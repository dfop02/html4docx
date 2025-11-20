import argparse
import inspect
import logging
import os
import re
from cgitb import handler
from io import BytesIO
from html.parser import HTMLParser

import docx
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor, Pt

from functools import lru_cache

from html4docx import constants
from html4docx import utils
from html4docx.constants import FONT_SIZES_NAMED
from html4docx.metadata import Metadata

# Paragraph-level styles (ParagraphFormat)
PARAGRAPH_FORMAT_STYLES = {
    'text-align': '_apply_alignment_paragraph',
    'line-height': '_apply_line_height_paragraph',
    'margin-left': '_apply_margins_paragraph',
    'margin-right': '_apply_margins_paragraph',
}

# Run-level styles (affect text formatting within runs)
PARAGRAPH_RUN_STYLES = {
    'font-weight': '_apply_font_weight_paragraph',
    'font-style': '_apply_font_style_paragraph',
    'text-decoration': '_apply_text_decoration_paragraph',
    'text-transform': '_apply_text_transform_paragraph',
    'font-size': '_apply_font_size_paragraph',
    'font-family': '_apply_font_family_paragraph',
    'color': '_apply_color_paragraph',
    'background-color': '_apply_background_color_paragraph'
}

RUN_STYLES = {
'font-weight': '_apply_font_weight_to_run',
    'font-style': '_apply_font_style_to_run',
    'text-decoration': '_apply_text_decoration_to_run',
    'text-transform': '_apply_text_transform_to_run',
    'font-size': '_apply_font_size_to_run',
    'font-family': '_apply_font_family_to_run',
    'color': '_apply_color_to_run',
    'background-color': '_apply_background_color_to_run'
}

class HtmlToDocx(HTMLParser):
    """
        Class to convert HTML to Docx
        source: https://docs.python.org/3/library/html.parser.html
    """
    def __init__(self):
        super().__init__()
        self.options = dict(constants.DEFAULT_OPTIONS)
        self.table_row_selectors = constants.DEFAULT_TABLE_ROW_SELECTORS
        self.table_style = constants.DEFAULT_TABLE_STYLE

    def set_initial_attrs(self, document = None):
        self.tags = {
            'span': [],
            'list': [],
        }
        self.doc = document if document else Document()
        self.bs = self.options['fix-html'] # whether or not to clean with BeautifulSoup
        self.paragraph = None
        self.skip = False
        self.skip_tag = None
        self.instances_to_skip = 0
        self.bookmark_id = 0
        self.in_li = False
        self.list_restart_counter = 0
        self.current_ol_num_id = None
        self._list_num_ids = {}

    @property
    def metadata(self) -> dict[str, any]:
        if not hasattr(self, '_metadata'):
            self._metadata = Metadata(self.doc)
        return self._metadata

    @property
    def include_tables(self) -> bool:
        return self.options.get('tables', True)

    @property
    def include_images(self) -> bool:
        return self.options.get('images', True)

    @property
    def include_styles(self) -> bool:
        return self.options.get('styles', True)

    def save(self, destination) -> None:
        """Save the document to a file path or BytesIO object."""
        if isinstance(destination, str):
            destination, _ = os.path.splitext(destination)
            self.doc.save(f'{destination}.docx')
        elif isinstance(destination, BytesIO):
            self.doc.save(destination)
        else:
            raise TypeError('destination must be a str path or BytesIO object')

    def copy_settings_from(self, other):
        """Copy settings from another instance of HtmlToDocx"""
        self.table_style = other.table_style

    def get_cell_html(self, soup):
        """
        Returns string of td element with opening and closing <td> tags removed
        Cannot use find_all as it only finds element tags and does not find text which
        is not inside an element
        """
        return ' '.join([str(i) for i in soup.contents])

    def set_cell_background(self, cell, color):
        """Set the background color of a table cell."""
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), color.lstrip('#'))
        tcPr.append(shd)

    def set_cell_borders(self, cell, styles):
        """
        Apply custom borders to a table cell using CSS-like styles.

        Args:
            cell: The table cell to style.
            styles (dict): A dictionary of CSS-like border styles.
                          Example: {"border": "1px 5px 10px 20px", "border-style": "solid", "border-color": "#FF0000"}
        sources:
            Border attributes: http://officeopenxml.com/WPtableBorders.php
        """
        tc = cell._tc  # Get the table cell element
        tcPr = tc.get_or_add_tcPr()  # Get or add the table cell properties

        # Default border properties
        default_size = constants.DEFAULT_BORDER_SIZE
        default_color = constants.DEFAULT_BORDER_COLOR
        default_style = constants.DEFAULT_BORDER_STYLE
        borders = constants.default_borders()
        border_styles = constants.BORDER_STYLES
        keywords = constants.BORDER_KEYWORDS
        border_sides = ("top", "right", "bottom", "left")
        border_width_pattern = re.compile(r'^[0-9]*\.?[0-9]+(px|pt|cm|in|rem|em|%)$')

        def parse_border_style(value: str) -> str:
            """Parses border styles to match word standart"""
            return constants.BORDER_STYLES[value] if value in constants.BORDER_STYLES.keys() else 'none'

        def check_unit_keywords(value: str) -> str:
            """Convert medium, thin, thick keywords to numeric values (px)"""
            lower_val = value.lower()
            return keywords.get(lower_val, value)

        @lru_cache(maxsize=None)
        def border_unit_converter(unit_value: str):
            """Convert multiple units to pt that is used on Word table cell border"""
            unit_value = utils.remove_important_from_style(unit_value)
            unit_value = check_unit_keywords(unit_value)

            # Return default if no value or empty
            if not unit_value or unit_value == '':
                return default_size

            unit = re.sub(r'[0-9\.]+', '', unit_value)
            value = float(re.sub(r'[a-zA-Z\!\%]+', '', unit_value))  # Allow float values

            if unit == 'px':
                result = value * 0.75  # 1 px = 0.75 pt
            elif unit == 'cm':
                result = value * 28.35  # 1 cm = 28.35 pt
            elif unit == 'in':
                result = value * 72  # 1 inch = 72 pt
            elif unit == 'pt':
                result = value # default is pt
            elif unit == 'rem' or unit == 'em':
                result = value * 12  # Assuming 1rem/em = 16px, converted to pt
            elif unit == '%':
                result = constants.MAX_INDENT * (value / 100)
            else:
                return None  # Unsupported units return None

            return result

        def parse_border_value(value: str):
            """
            Parses a border value like:
            '1px solid #000000', 'solid 1px red', or '#000000 medium dashed' in any order.
            """
            parts = value.split()

            # Return all default if there is only 'none' or empty
            if (len(parts) == 1 and parts[0] == 'none') or (not value or value.strip() == ''):
                return default_size, default_style, default_color

            size = None
            style = default_style
            color = default_color

            for part in parts:
                clean_part = utils.remove_important_from_style(part).lower()

                # Detect size (units or keywords)
                if border_width_pattern.match(clean_part) or clean_part in keywords:
                    size = border_unit_converter(clean_part) or default_size
                    continue

                # Detect style
                if clean_part in border_styles:
                    style = parse_border_style(clean_part)
                    continue

                # Detect color
                if utils.is_color(clean_part):
                    color = utils.parse_color(clean_part, return_hex=True)
                    continue

            # If only style or color was given without size, we need to apply 1pt size to render it
            if len(parts) >= 1 and size is None:
                size = 1.0

            return size, style, color

        for css_prop, css_value in styles.items():
            if not css_prop.startswith("border"):
                continue

            # Case 1: 'border' shorthand applies to all sides
            if css_prop == "border":
                values = css_value.split()
                if len(values) in (1, 3):
                    # Single value or full triple — parse all parts in any order
                    size, style, color = parse_border_value(css_value)
                    for side in border_sides:
                        borders[side].update({"size": size, "style": style, "color": color})
                elif len(values) == 2:
                    # Two widths (top/bottom, left/right)
                    tb_size = border_unit_converter(values[0]) or default_size
                    lr_size = border_unit_converter(values[1]) or default_size
                    for side in ("top", "bottom"):
                        borders[side]["size"] = tb_size
                    for side in ("left", "right"):
                        borders[side]["size"] = lr_size
                elif len(values) == 4:
                    # Four widths (top, right, bottom, left)
                    for side, val in zip(border_sides, values):
                        borders[side]["size"] = border_unit_converter(val) or default_size

            # Case 2: 'border-width', 'border-color', 'border-style' — apply to all sides
            elif css_prop in ("border-width", "border-color", "border-style"):
                prop_type = css_prop.split("-")[1]
                if prop_type == "width":
                    size = border_unit_converter(css_value) or default_size
                    for side in border_sides:
                        borders[side]["size"] = size
                elif prop_type == "color":
                    color = utils.parse_color(css_value, return_hex=True)
                    for side in border_sides:
                        borders[side]["color"] = color
                elif prop_type == "style":
                    style = parse_border_style(css_value.lower())
                    for side in border_sides:
                        borders[side]["style"] = style

            # Case 3: 'border-top', 'border-right-width', etc.
            else:
                parts = css_prop.split("-")
                side = parts[1]
                prop_type = parts[2] if len(parts) > 2 else None
                if prop_type == "width":
                    borders[side]["size"] = border_unit_converter(css_value) or default_size
                elif prop_type == "color":
                    borders[side]["color"] = utils.parse_color(css_value, return_hex=True)
                elif prop_type == "style":
                    borders[side]["style"] = parse_border_style(css_value.lower())
                else:
                    # Full side shorthand in any order
                    size, style, color = parse_border_value(css_value)
                    borders[side].update({"size": size, "style": style, "color": color})

        # Check if w:tcBorders exists, otherwise create it
        tcBorders = tcPr.first_child_found_in('w:tcBorders')
        if tcBorders is None:
            tcBorders = OxmlElement('w:tcBorders')
            tcPr.append(tcBorders)

        # Apply borders to the cell
        for side, border_info in borders.items():
            if border_info["size"] > 0:
                border = OxmlElement(f"w:{side}")
                border.set(qn("w:val"), border_info["style"])  # Set border style
                border.set(qn("w:sz"), str(border_info["size"] * 8))  # Word uses eighths of a point
                border.set(qn("w:color"), border_info["color"].replace('#', ''))  # Set border color
                tcBorders.append(border)

    def add_bookmark(self, bookmark_name):
        """Adds a word bookmark to an existing paragraph"""
        bookmark_start = OxmlElement('w:bookmarkStart')
        bookmark_start.set(qn('w:id'), str(self.bookmark_id))
        bookmark_start.set(qn('w:name'), bookmark_name)

        bookmark_end = OxmlElement('w:bookmarkEnd')
        bookmark_end.set(qn('w:id'), str(self.bookmark_id))

        if not self.paragraph:
            self.paragraph = self.doc.add_paragraph()

        self.paragraph._element.insert(0, bookmark_start)
        self.paragraph._element.append(bookmark_end)
        self.bookmark_id += 1

    def apply_styles_to_run(self, run, style):
        if not style or not hasattr(run, 'font'):
            return

        for style_name, style_value in style.items():
            # Skip paragraph-level styles for runs
            if style_name in PARAGRAPH_FORMAT_STYLES:
                continue

            elif style_name in RUN_STYLES:
                handler = getattr(self, RUN_STYLES[style_name])

                param_name = style_name.replace('-', '_')
                handler(
                    run=run,
                    **{param_name: style_value}
                )

            else:
                logging.warning(f"Warning: Unrecognized style '{style_name}', will be skipped.")

    def apply_styles_to_paragraph(self, paragraph, style, init=False):
        if not style or not hasattr(paragraph, 'paragraph_format'):
            return

        for style_name, style_value in style.items():
            if init and style_name in PARAGRAPH_FORMAT_STYLES:
                handler = getattr(self, PARAGRAPH_FORMAT_STYLES[style_name])
            elif not init and style_name in PARAGRAPH_RUN_STYLES:
                handler = getattr(self, PARAGRAPH_RUN_STYLES[style_name])
            else:
                logging.warning(f"Warning: Unrecognized paragraph style '{style_name}', will be skipped.")
                continue

            handler(
                paragraph=paragraph,
                style_name=style_name,
                value=style_value,
                all_styles=style
            )

    def _apply_alignment_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        value = kwargs['value']

        align = utils.remove_important_from_style(value)

        if 'center' in align:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif 'left' in align:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif 'right' in align:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif 'justify' in align:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def _apply_line_height_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        value = kwargs['value']

        line_height = utils.remove_important_from_style(value)

        if line_height in ('normal', 'inherit'):
            paragraph.paragraph_format.line_spacing = None
        else:
            try:
                if line_height.replace('.', '').replace('%', '').isdigit():
                    multiplier = float(line_height[:-1]) / 100.0 if line_height.endswith('%') else float(line_height)
                    paragraph.paragraph_format.line_spacing = multiplier
                else:
                    converted = utils.unit_converter(line_height, target_unit="pt")
                    if converted is not None:
                        paragraph.paragraph_format.line_spacing = converted
            except (ValueError, TypeError):
                paragraph.paragraph_format.line_spacing = None

    def _apply_margins_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        style_name = kwargs['style_name']
        all_styles = kwargs['all_styles']

        margin_left = all_styles.get('margin-left')
        margin_right = all_styles.get('margin-right')

        if margin_left and margin_right:
            if 'auto' in margin_left and 'auto' in margin_right:
                paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                return

        if style_name == 'margin-left' and margin_left and 'auto' not in margin_left:
            paragraph.paragraph_format.left_indent = utils.unit_converter(margin_left)

        if style_name == 'margin-right' and margin_right and 'auto' not in margin_right:
            paragraph.paragraph_format.right_indent = utils.unit_converter(margin_right)

    def _apply_font_weight_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        value = kwargs['value']

        font_weight = utils.remove_important_from_style(value).lower()

        for run in paragraph.runs:
            self._apply_font_weight_to_run(
                run=run,
                font_weight=font_weight,
            )

    def _apply_font_weight_to_run(self, **kwargs):
        font_weight = kwargs['font_weight']
        run = kwargs['run']
        if font_weight in ('bold', 'bolder', '700', '800', '900'):
            run.font.bold = True
        elif font_weight in ('normal', 'lighter', '400', '300', '100'):
            run.font.bold = False
        # Note: Decide what to do for values between 400-700
        elif font_weight.isdigit():
            weight = int(font_weight)
            run.font.bold = weight >= 700

    def _apply_font_style_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        value = kwargs['value']

        font_style = utils.remove_important_from_style(value).lower()

        for run in paragraph.runs:
            self._apply_font_style_to_run(
                run=run,
                font_style=font_style,
            )

    def _apply_font_style_to_run(self, **kwargs):
        font_style = kwargs['font_style']
        run = kwargs['run']

        if font_style in ('italic', 'oblique'):
            run.font.italic = True
        elif font_style == 'normal':
            run.font.italic = False

    def _apply_font_size_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        value = kwargs['value']

        font_size = utils.remove_important_from_style(value).lower()

        if font_size in FONT_SIZES_NAMED:
            font_size = FONT_SIZES_NAMED[font_size]

        for run in paragraph.runs:
            self._apply_font_size_to_run(
                run=run,
                font_size=font_size,
            )

    def _apply_font_size_to_run(self, **kwargs):
        run = kwargs['run']
        font_size = kwargs['font_size']

        font_size = utils.remove_important_from_style(font_size).lower()
        font_size = utils.adapt_font_size(font_size)

        try:
            if font_size in ('normal', 'initial', 'inherit'):
                run.font.size = None
            else:
                converted_size = utils.unit_converter(font_size, target_unit="pt")
                if converted_size is not None:
                    run.font.size = converted_size

        except (ValueError, TypeError) as e:
            logging.warning(f"Warning: Could not parse font-size '{font_size}': {e}")

    def _apply_font_family_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        value = kwargs['value']

        font_family = utils.remove_important_from_style(value).strip()

        for run in paragraph.runs:
            self._apply_font_family_to_run(
                run=run,
                font_family=font_family,
            )

    def _apply_font_family_to_run(self, **kwargs):
        run = kwargs['run']
        font_family = kwargs['font_family']

        if not font_family or font_family in ('inherit', 'initial', 'unset'):
            return

        try:
            font_families = [f.strip().strip('"\'') for f in font_family.split(',')]

            for font_name in font_families:
                if font_name and font_name not in ('inherit', 'initial', 'unset', 'serif', 'sans-serif', 'monospace',
                                                   'cursive', 'fantasy', 'system-ui'):
                    run.font.name = font_name
                    break
                elif font_name in ('serif', 'sans-serif', 'monospace'):
                    generic_font_map = {
                        'serif': 'Times New Roman',
                        'sans-serif': 'Arial',
                        'monospace': 'Courier New'
                    }
                    run.font.name = generic_font_map[font_name]
                    break

        except (AttributeError, Exception) as e:
            logging.warning(f"Warning: Could not apply font-family '{font_family}': {e}")

    def _apply_color_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        all_styles = kwargs['all_styles']

        color_value = utils.remove_important_from_style(all_styles.get('color', '')).lower().strip()
        if color_value in ('inherit', 'initial', 'transparent', 'currentcolor'):
            return

        for run in paragraph.runs:
            self._apply_color_to_run(
                run=run,
                color=color_value,
            )

    def _apply_color_to_run(self, **kwargs):
        run = kwargs['run']
        color_value = kwargs['color']

        try:
            colors = utils.parse_color(color_value)
            run.font.color.rgb = RGBColor(*colors)
        except (ValueError, AttributeError) as e:
            logging.warning(f"Could not apply color '{color_value}': {e}")

    def _apply_text_transform_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        value = kwargs['value']

        text_transform = utils.remove_important_from_style(value).lower()

        for run in paragraph.runs:
            self._apply_text_transform_to_run(
                run=run,
                text_transform=text_transform,
            )

    def _apply_text_transform_to_run(self, **kwargs):
        run = kwargs['run']
        text_transform = kwargs['text_transform']

        if not run.text:
            return

        try:
            if text_transform == 'uppercase':
                run.text = run.text.upper()
            elif text_transform == 'lowercase':
                run.text = run.text.lower()
            elif text_transform == 'capitalize':
                run.text = run.text.title()
            elif text_transform in ('none', 'initial', 'inherit'):
                # No transformation needed
                pass
            elif text_transform in ('full-width', 'math-auto', 'full-size-kana'):
                logging.warning(f"Warning: Unsupported text transform '{text_transform}'")

        except (AttributeError, Exception) as e:
            logging.warning(f"Warning: Could not apply text-transform '{text_transform}': {e}")

    def _apply_text_decoration_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        value = kwargs['value']

        text_decoration = utils.remove_important_from_style(value).lower()

        for run in paragraph.runs:
            self._apply_text_decoration_to_run(
                run=run,
                text_decoration=text_decoration,
            )

    def _apply_text_decoration_to_run(self, **kwargs):
        run = kwargs['run']
        text_decoration = kwargs['text_decoration']

        decorations = text_decoration.split()

        for decoration in decorations:
            if decoration in ('underline', 'overline', 'line-through', 'blink'):
                if decoration == 'underline':
                    run.font.underline = True
                elif decoration == 'line-through':
                    run.font.strike = True
                else:
                    logging.warning(f"Warning: Unsupported text decoration '{decoration}'")

            elif decoration == 'none':
                run.font.underline = False

            elif decoration in ('solid', 'double', 'dotted', 'dashed', 'wavy'):
                if run.font.underline:
                    if decoration == 'double':
                        run.font.underline = WD_UNDERLINE.DOUBLE
                    elif decoration == 'dotted':
                        run.font.underline = WD_UNDERLINE.DOTTED
                    elif decoration == 'dashed':
                        run.font.underline = WD_UNDERLINE.DASH
                    elif decoration == 'wavy':
                        run.font.underline = WD_UNDERLINE.WAVY

                    # 'solid' is the default, no change needed

        # Note: Check if adding color support is possible.

    def _apply_background_color_paragraph(self, **kwargs):
        paragraph = kwargs['paragraph']
        value = kwargs['value']

        background_color = utils.remove_important_from_style(value).lower().strip()

        if background_color in ('inherit', 'initial', 'transparent', 'none'):
            return

        try:
            color_hex = utils.parse_color(background_color, return_hex=True)
            if not color_hex:
                return

            # Apply to PARAGRAPH
            from docx.oxml.shared import qn
            from docx.oxml import OxmlElement

            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), color_hex.lstrip('#'))

            # Apply to paragraph properties
            p_pr = paragraph._element.get_or_add_pPr()

            existing_shd = p_pr.find(qn('w:shd'))
            if existing_shd is not None:
                p_pr.remove(existing_shd)

            p_pr.append(shd)

        except Exception as e:
            logging.warning(f"Could not apply background-color to paragraph: {e}")

    def _apply_background_color_to_run(self, run, background_color):
        try:
            color_hex = utils.parse_color(background_color, return_hex=True)
            if not color_hex:
                return

            from docx.oxml.shared import qn
            from docx.oxml import OxmlElement

            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), color_hex.lstrip('#'))

            r_pr = run._element.get_or_add_rPr()

            # Remove existing shading
            existing_shd = r_pr.find(qn('w:shd'))
            if existing_shd is not None:
                r_pr.remove(existing_shd)

            r_pr.append(shd)

        except Exception as e:
            logging.warning(f"Could not apply background-color to run: {e}")

    def add_text_align_or_margin_to(self, obj, style):
        """Styles that can be applied on multiple objects"""
        if 'text-align' in style:
            align = utils.remove_important_from_style(style['text-align'])

            if 'center' in align:
                obj.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif 'left' in align:
                obj.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif 'right' in align:
                obj.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif 'justify' in align:
                obj.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        if 'margin-left' in style and 'margin-right' in style:
            if 'auto' in style['margin-left'] and 'auto' in style['margin-right']:
                obj.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif 'margin-left' in style:
            obj.left_indent = utils.unit_converter(style['margin-left'])

    def add_styles_to_table_cell(self, styles, doc_cell, cell_row):
        """Styles that must be applied specifically in a _Cell object"""
        # Set background color
        if 'background-color' in styles:
            self.set_cell_background(doc_cell, styles['background-color'])

        # Set width (approximate, since DOCX uses different units)
        if 'width' in styles:
            doc_cell.width = utils.unit_converter(styles['width'])

        # Set height (due word limitations, cannot set individually cell height, only whole row)
        if 'height' in styles:
            cell_row.height = utils.unit_converter(styles['height'])

        # Set text color
        if 'color' in styles:
            color = utils.parse_color(styles['color'])
            if color:
                for paragraph in doc_cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.color.rgb = color

        # Set vertical align (for individual cells)
        if 'vertical-align' in styles:
            align = utils.remove_important_from_style(styles['vertical-align'])

            if 'top' in align:
                doc_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            elif 'middle' in align:
                doc_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            elif 'bottom' in align:
                doc_cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

        # Set borders
        if any('border' in style for style in styles.keys()):
            self.set_cell_borders(doc_cell, styles)

        self.add_text_align_or_margin_to(doc_cell.paragraphs[0], styles)

    def add_styles_to_run(self, style):
        if 'font-size' in style:
            font_size = utils.remove_important_from_style(style['font-size'])
            # Adapt font_size when text, ex.: small, medium, etc.
            font_size = utils.adapt_font_size(font_size)

            if font_size:
                for run in self.paragraph.runs:
                    run.font.size = utils.unit_converter(font_size)

        if 'color' in style:
            colors = utils.parse_color(style['color'])
            self.run.font.color.rgb = RGBColor(*colors)

        if 'background-color' in style:
            # color = utils.parse_color(style['background-color'], return_hex=True)
            self._apply_background_color_to_run(self.run, style['background-color'])

            # Little trick to apply background-color to paragraph
            # because `self.run.font.highlight_color`
            # has a very limited amount of colors
            # shd = OxmlElement('w:shd')
            # shd.set(qn('w:val'), 'clear')
            # shd.set(qn('w:color'), 'auto')
            # shd.set(qn('w:fill'), color.lstrip('#'))
            #
            # # Make sure the paragraph styling element exists
            # self.paragraph.paragraph_format.element.get_or_add_pPr()
            #
            # # Append the shading element
            # self.paragraph.paragraph_format.element.pPr.append(shd)

    def handle_li(self):
        '''
            Handle li tags
            source: https://stackoverflow.com/a/78685353/17274446
        '''
        list_depth = len(self.tags['list']) or 1
        list_type = self.tags['list'][-1] if self.tags['list'] else 'ul'
        level = min(list_depth, 3)
        style_key = list_type if level <= 1 else f"{list_type}{level}"
        list_style = constants.STYLES.get(style_key, 'List Number' if list_type == 'ol' else 'List Bullet')

        self.paragraph = self.doc.add_paragraph(style=list_style)
        self.in_li = True

        if list_type == "ol":
            # Use your current_ol_num_id (generated on <ol> open) as key
            ol_id = self.current_ol_num_id or -1

            if ol_id not in self._list_num_ids:
                # First time using this <ol> → create a new numId

                style_obj = self.paragraph.style
                num_id_style = None

                if hasattr(style_obj._element.pPr, 'numPr'):
                    num_id_style = style_obj._element.pPr.numPr.numId.val

                if num_id_style is not None:
                    ct_numbering = self.doc.part.numbering_part.numbering_definitions._numbering
                    ct_num = ct_numbering.num_having_numId(num_id_style)
                    abstractNumId = ct_num.abstractNumId.val

                    # Add new numId linked to same abstractNumId
                    ct_num_new = ct_numbering.add_num(abstractNumId)
                    new_num_id = ct_num_new.numId

                    # Apply startOverride for level 0
                    lvl_override = ct_num_new.add_lvlOverride(0)
                    start_override = lvl_override._add_startOverride()
                    start_override.val = 1

                    # Cache this new numId
                    self._list_num_ids[ol_id] = new_num_id
            else:
                new_num_id = self._list_num_ids[ol_id]

            # Assign this numId to the paragraph
            pPr = self.paragraph._p.get_or_add_pPr()
            numPr = OxmlElement('w:numPr')

            numId_elem = OxmlElement('w:numId')
            numId_elem.set(qn('w:val'), str(new_num_id))

            ilvl = OxmlElement('w:ilvl')
            ilvl.set(qn('w:val'), str(level - 1))

            numPr.append(ilvl)
            numPr.append(numId_elem)
            pPr.append(numPr)

    def add_image_to_cell(self, cell, image, width=None, height=None):
        # python-docx doesn't have method yet for adding images to table cells. For now we use this
        paragraph = cell.add_paragraph()
        run = paragraph.add_run()
        run.add_picture(image, width, height)

    def handle_img(self, current_attrs):
        if not self.include_images:
            self.skip = True
            self.skip_tag = 'img'
            return

        if 'src' not in current_attrs:
            self.doc.add_paragraph("<image: no_src>")
            return

        src = current_attrs['src']

        # added image dimension, interpreting values as pixel only
        height = utils.unit_converter(current_attrs['height']) if 'height' in current_attrs else None
        width = utils.unit_converter(current_attrs['width']) if 'width' in current_attrs else None

        # fetch image
        image = utils.fetch_image_data(src)

        if not self.paragraph:
            self.paragraph = self.doc.add_paragraph()

        self.run = self.paragraph.add_run()

        # add image to doc
        if image:
            try:
                if isinstance(self.doc, docx.document.Document):
                    self.run.add_picture(image, width, height)
                else:
                    self.add_image_to_cell(self.doc, image, width, height)
            except FileNotFoundError:
                image = None

        if not image:
            if utils.is_url(src):
                self.doc.add_paragraph("<image: %s>" % src)
            else:
                # avoid exposing filepaths in document
                self.doc.add_paragraph("<image: %s>" % utils.get_filename_from_url(src))

        '''
        #adding style
        For right-alignment: `'float: right;'`
        For center-alignment: `'display: block; margin-left: auto; margin-right: auto;'`
        Everything else would be Left aligned
        '''
        if 'style' in current_attrs:
            style = current_attrs['style']
            image_alignment = utils.get_image_alignment(style)
            last_paragraph = self.doc.paragraphs[-1]
            if image_alignment == utils.ImageAlignment.RIGHT:
                last_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if image_alignment == utils.ImageAlignment.CENTER:
                last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def handle_table(self, current_attrs):
        """
        To handle nested tables, we will parse tables manually as follows:
        Get table soup
        Create docx table
        Iterate over soup and fill docx table with new instances of this parser
        Tell HTMLParser to ignore any tags until the corresponding closing table tag
        """
        table_soup = self.tables[self.table_no]
        rows, cols = self.get_table_dimensions(table_soup)
        self.table = self.doc.add_table(rows, cols)

        if self.table_style:
            # Available Table Styles
            # https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html#table-styles-in-default-template
            try:
                # Fixed 'style lookup by style_id is deprecated.'
                # https://stackoverflow.com/a/29567907/17274446
                self.table_style = ' '.join(re.findall(r'[A-Z][a-z]*|[0-9]', self.table_style))
                # Available Table Styles
                # https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html#table-styles-in-default-template
                self.table.style = self.table_style
            except KeyError as e:
                raise ValueError(f"Unable to apply style {self.table_style}.") from e

        # Reference:
        # https://python-docx.readthedocs.io/en/latest/api/table.html#cell-objects
        used_cells = [[False] * cols for _ in range(rows)]
        for cell_row, row in enumerate(self.get_table_rows(table_soup)):
            col_offset = 0  # Shift index if some columns are occupied
            for col in self.get_table_columns(row):
                while used_cells[cell_row][col_offset]:
                    col_offset += 1

                current_row = cell_row
                current_col = col_offset

                cell_html = self.get_cell_html(col)
                if col.name == 'th':
                    cell_html = "<b>%s</b>" % cell_html

                # Get _Cell object from table based on cell_row and cell_col
                docx_cell = self.table.cell(current_row, current_col)

                # Reference:
                # https://python-docx.readthedocs.io/en/latest/dev/analysis/features/table/cell-merge.html
                rowspan = int(col.get('rowspan', 1))
                colspan = int(col.get('colspan', 1))

                if rowspan > 1 or colspan > 1:
                    docx_cell = docx_cell.merge(
                        self.table.cell(
                            current_row + (rowspan - 1),
                            current_col + (colspan - 1)
                        )
                    )

                    # Mark all merged cells as used
                    for r in range(current_row, current_row + rowspan):
                        for c in range(current_col, current_col + colspan):
                            used_cells[r][c] = True
                else:
                    used_cells[current_row][current_col] = True

                # Parse cell styles
                cell_styles = utils.parse_dict_string(col.get('style', ''))

                if 'width' in cell_styles or 'height' in cell_styles:
                    self.table.autofit = False
                    self.table.allow_autofit = False

                child_parser = HtmlToDocx()
                child_parser.copy_settings_from(self)
                child_parser.add_html_to_cell(cell_html, docx_cell)
                child_parser.add_styles_to_table_cell(cell_styles, docx_cell, self.table.rows[cell_row])

                col_offset += colspan  # Move to the next real column

        if 'style' in current_attrs and self.table:
            style = utils.parse_dict_string(current_attrs['style'])
            self.add_text_align_or_margin_to(self.table, style)

        # skip all tags until corresponding closing tag
        self.instances_to_skip = len(table_soup.find_all('table'))
        self.skip_tag = 'table'
        self.skip = True
        self.table = None

    def handle_div(self, current_attrs):
        self.paragraph = self.doc.add_paragraph()

        # handle page break
        if 'style' in current_attrs and 'page-break-after: always' in current_attrs['style']:
            self.doc.add_page_break()

    def handle_hr(self):
        # This implementation was taken from:
        # https://github.com/python-openxml/python-docx/issues/105#issuecomment-62806373
        self.paragraph = self.doc.add_paragraph()
        pPr = self.paragraph._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        pPr.insert_element_before(
            pBdr,
            'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
            'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
            'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
            'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
            'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
            'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'
        )
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), 'auto')
        pBdr.append(bottom)

    def handle_link(self, href, text, tooltip=None):
        """
        A function that places a hyperlink within a paragraph object.

        Args:
            href: A string containing the required url.
            text: The text displayed for the url.
            tooltip: The text displayed when holder link.
        """
        is_external = href.startswith('http') if href else False
        hyperlink = OxmlElement('w:hyperlink')

        if is_external:
            # Create external hyperlink
            rel_id = self.paragraph.part.relate_to(
                href,
                docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK,
                is_external=True
            )

            # Create the w:hyperlink tag and add needed values
            hyperlink.set(qn('r:id'), rel_id)
        else:
            # Create internal hyperlink (anchor)
            hyperlink.set(qn('w:anchor'), href.replace('#', ''))

        if tooltip is not None:
            # set tooltip to hyperlink
            hyperlink.set(qn('w:tooltip'), tooltip)

        # Create sub-run
        subrun = self.paragraph.add_run()
        rPr = OxmlElement('w:rPr')

        # add default color
        c = OxmlElement('w:color')
        c.set(qn('w:val'), "0000EE")
        rPr.append(c)

        # add underline
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

        subrun._r.append(rPr)
        subrun._r.text = text

        # Add subrun to hyperlink
        hyperlink.append(subrun._r)

        # Add hyperlink to run
        self.paragraph._p.append(hyperlink)

    def handle_starttag(self, tag, attrs):
        if self.skip:
            return
        if tag == 'head':
            self.skip = True
            self.skip_tag = tag
            self.instances_to_skip = 0
            return
        elif tag == 'body':
            return

        current_attrs = dict(attrs)

        if tag == 'span':
            self.tags['span'].append(current_attrs)
            return
        elif tag in ['ol', 'ul']:
            if tag == 'ol':
                # Assign new ID if it's a fresh top-level list
                self.list_restart_counter += 1
                self.current_ol_num_id = self.list_restart_counter
            else:
                self.current_ol_num_id = None  # unordered list

            self.tags['list'].append(tag)
            return # don't apply styles for now
        elif tag == 'br':
            try:
                self.run.add_break()
            except AttributeError:
                self.paragraph = self.doc.add_paragraph()
                self.run = self.paragraph.add_run()
                self.run.add_break()
            return

        self.tags[tag] = current_attrs
        if tag in ['p', 'pre']:
            if not self.in_li:
                self.paragraph = self.doc.add_paragraph()

        elif tag == 'li':
            self.handle_li()

        elif tag == 'hr':
            self.handle_hr()

        elif re.match('h[1-9]', tag):
            if isinstance(self.doc, docx.document.Document):
                h_size = int(tag[1])
                self.paragraph = self.doc.add_heading(level=min(h_size, 9))
            else:
                self.paragraph = self.doc.add_paragraph()

        elif tag == 'img':
            self.handle_img(current_attrs)
            self.paragraph = self.doc.paragraphs[-1]

        elif tag == 'table':
            if self.include_tables:
                self.handle_table(current_attrs)
                return

        elif tag == 'div':
            self.handle_div(current_attrs)

        # set new run reference point in case of leading line breaks
        if tag in ['p', 'li', 'pre']:
            self.run = self.paragraph.add_run()

        if 'id' in current_attrs:
            self.add_bookmark(current_attrs['id'])

        # add style
        if not self.include_styles:
            return

        if 'style' in current_attrs and self.paragraph:
            style = utils.parse_dict_string(current_attrs['style'])
            self.apply_styles_to_paragraph(self.paragraph, style, True)

    def handle_endtag(self, tag):
        if self.skip:
            if not tag == self.skip_tag:
                return

            if self.instances_to_skip > 0:
                self.instances_to_skip -= 1
                return

            self.skip = False
            self.skip_tag = None
            self.paragraph = None

        if tag == 'span':
            if self.tags['span']:
                self.tags['span'].pop()
                return
        elif tag in ['ol', 'ul']:
            utils.remove_last_occurence(self.tags['list'], tag)
            if tag == 'ol':
                self._list_num_ids.pop(self.current_ol_num_id, None)
                self.current_ol_num_id = None
            return
        elif tag == 'table':
            if self.include_tables:
                self.table_no += 1
                self.table = None
                self.paragraph = None
        elif tag == 'li':
            self.in_li = False

        if tag in self.tags:
            self.tags.pop(tag)
        # maybe set relevant reference to None?

    def handle_data(self, data):
        if self.skip:
            return

        # Only remove white space if we're not in a pre block.
        if 'pre' not in self.tags:
            # remove leading and trailing whitespace in all instances
            data = utils.remove_whitespace(data, True, True)

        if not self.paragraph:
            self.paragraph = self.doc.add_paragraph()

        # There can only be one nested link in a valid html document
        # You cannot have interactive content in an A tag, this includes links
        # https://html.spec.whatwg.org/#interactive-content
        link = self.tags.get('a', {})
        href = link.get('href', None)
        title = link.get('title', None)

        if link and href:
            self.handle_link(href, data, title)
            return

        # If there's a link, dont put the data directly in the run
        self.run = self.paragraph.add_run(data)

        for span in self.tags['span']:
            if 'style' in span:
                style = utils.parse_dict_string(span['style'])
                self.apply_styles_to_run(self.run, style)

        # add font style and name
        for tag, attrs in self.tags.items():
            if tag in constants.FONT_STYLES:
                font_style = constants.FONT_STYLES[tag]
                setattr(self.run.font, font_style, True)

            if tag in constants.FONT_NAMES:
                font_name = constants.FONT_NAMES[tag]
                self.run.font.name = font_name

            if 'style' in attrs and (tag in ['p']):
                style = utils.parse_dict_string(attrs['style'])
                self.apply_styles_to_paragraph(self.paragraph, style)

            if 'style' in attrs and (tag in ['div', 'li', 'pre'] or re.match(r'h[1-9]', tag)):
                style = utils.parse_dict_string(attrs['style'])
                self.add_styles_to_run(style)

    def ignore_nested_tables(self, tables_soup):
        """
        Returns array containing only the highest level tables
        Operates on the assumption that bs4 returns child elements immediately after
        the parent element in `find_all`. If this changes in the future, this method will need to be updated

        :return:
        """
        new_tables = []
        nest = 0
        for table in tables_soup:
            if nest:
                nest -= 1
                continue
            new_tables.append(table)
            nest = len(table.find_all('table'))
        return new_tables

    def get_table_rows(self, table_soup):
        # If there's a header, body, footer or direct child tr tags, add row dimensions from there
        return table_soup.select(', '.join(self.table_row_selectors), recursive=False)

    def get_table_columns(self, row):
        # Get all columns for the specified row tag.
        return row.find_all(['th', 'td'], recursive=False) if row else []

    def get_table_dimensions(self, table_soup):
        # Get rows for the table
        rows = self.get_table_rows(table_soup)
        # Table is either empty or has non-direct children between table and tr tags
        # Thus the row dimensions and column dimensions are assumed to be 0
        # A table can have a varying number of columns per row,
        # so it is important to find the maximum number of columns in any row
        if not rows:
            return 0, 0

        default_span = 1
        max_cols = 0
        max_rows = len(rows)

        for row_idx, row in enumerate(rows):
            cols = self.get_table_columns(row)
            # Handle colspan
            row_col_count = sum(int(col.get('colspan', default_span)) for col in cols)
            max_cols = max(max_cols, row_col_count)

            # Handle rowspan
            for col in cols:
                rowspan = int(col.get('rowspan', default_span))
                if rowspan > default_span:
                    max_rows = max(max_rows, row_idx + rowspan)

        return max_rows, max_cols

    def get_tables(self) -> None:
        if not hasattr(self, 'soup'):
            self.options['tables'] = False
            return

        self.tables = self.ignore_nested_tables(self.soup.find_all('table'))
        self.table_no = 0

    def run_process(self, html: str) -> None:
        if self.bs and BeautifulSoup:
            self.soup = BeautifulSoup(html, 'html.parser')

            html = str(self.soup)
        if self.include_tables:
            self.get_tables()
        self.feed(html)

    def add_html_to_document(self, html: str, document) -> None:
        if not isinstance(html, str):
            raise ValueError(f'First argument needs to be a {str}')
        elif not isinstance(document, docx.document.Document) and not isinstance(document, docx.table._Cell):
            raise ValueError(f'Second argument needs to be a {docx.document.Document}')

        self.set_initial_attrs(document)
        self.run_process(html)

    def add_html_to_cell(self, html: str, cell: docx.table._Cell) -> None:
        if not isinstance(cell, docx.table._Cell):
            raise ValueError(f'Second argument needs to be a {docx.table._Cell}')

        unwanted_paragraph = cell.paragraphs[0]
        utils.delete_paragraph(unwanted_paragraph)
        self.set_initial_attrs(cell)
        self.run_process(html)
        # cells must end with a paragraph or will get message about corrupt file
        # https://stackoverflow.com/a/29287121
        if not self.doc.paragraphs:
            self.doc.add_paragraph('')

    def parse_html_file(self, filename_html: str, filename_docx, encoding: str = 'utf-8') -> None:
        with open(filename_html, 'r', encoding=encoding) as infile:
            html = infile.read()

        self.set_initial_attrs()
        self.run_process(html)

        if not filename_docx:
            path, filename = os.path.split(filename_html)
            filename_docx = f'{path}/new_docx_file_{filename}'

        self.save(filename_docx)

    def parse_html_string(self, html: str) -> docx.document.Document:
        self.set_initial_attrs()
        self.run_process(html)
        return self.doc

if __name__ == '__main__':
    arg_parser = argparse.ArgumentParser(description='Convert .html file into .docx file with formatting')
    arg_parser.add_argument('filename_html', help='The .html file to be parsed')
    arg_parser.add_argument(
        'filename_docx',
        nargs='?',
        help='The name of the .docx file to be saved. Default new_docx_file_[filename_html]',
        default=None
    )
    arg_parser.add_argument('--bs', action='store_true',
                            help='Attempt to fix html before parsing. Requires bs4. Default True')

    args = vars(arg_parser.parse_args())
    file_html = args.pop('filename_html')
    html_parser = HtmlToDocx()
    html_parser.parse_html_file(file_html, **args)
