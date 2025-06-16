"""
Make 'span' in tags dict a stack
maybe do the same for all tags in case of unclosed tags?
optionally use bs4 to clean up invalid html?

the idea is that there is a method that converts html files into docx
but also have api methods that let user have more control e.g. so they
can nest calls to something like 'convert_chunk' in loops

user can pass existing document object as arg
(if they want to manage rest of document themselves)

How to deal with block level style applied over table elements? e.g. text align
"""
import argparse
import os
import re
from html.parser import HTMLParser

import docx
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor

from html4docx import utils

# values in inches
INDENT = 0.25
LIST_INDENT = 0.5
MAX_INDENT = 5.5 # To stop indents going off the page

# Style to use with tables. By default no style is used.
DEFAULT_TABLE_STYLE = None

class HtmlToDocx(HTMLParser):
    def __init__(self):
        super().__init__()
        self.options = {
            'fix-html': True,
            'images': True,
            'tables': True,
            'styles': True,
        }
        self.table_row_selectors = [
            'table > tr',
            'table > thead > tr',
            'table > tbody > tr',
            'table > tfoot > tr'
        ]
        self.table_style = DEFAULT_TABLE_STYLE

    def set_initial_attrs(self, document=None):
        self.tags = {
            'span': [],
            'list': [],
        }
        self.doc = document if document else Document()
        self.document = self.doc
        self.bs = self.options['fix-html'] # whether or not to clean with BeautifulSoup
        self.include_tables = True # TODO add this option back in?
        self.include_images = self.options['images']
        self.include_styles = self.options['styles']
        self.paragraph = None
        self.skip = False
        self.skip_tag = None
        self.instances_to_skip = 0
        self.bookmark_id = 0
        # This counter simulates unique numbering IDs for <ol> elements.
        # Each new top-level ordered list increments this to trick python-docx into restarting list numbering.
        # Required because python-docx doesn't expose fine-grained list numbering control.
        self.in_li = False
        self.list_restart_counter = 0
        self.current_ol_num_id = None
        self._list_num_ids = {}

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
        default_size = 0
        default_color = "#000000"
        default_style = "single"
        borders = {
            "top": {"size": default_size, "color": default_color, "style": default_style},
            "right": {"size": default_size, "color": default_color, "style": default_style},
            "bottom": {"size": default_size, "color": default_color, "style": default_style},
            "left": {"size": default_size, "color": default_color, "style": default_style},
        }
        border_styles = {
            "none": "none",
            "hidden": "none",
            "initial": "none",
            "solid": "single",
            "dotted": "dotted",
            "dashed": "dashed",
            "double": "double",
            "inset": "inset",
            "outset": "outset"
        }

        def parse_border_style(value: str):
            """Parses border styles to match word standart"""
            return border_styles[value] if value in border_styles.keys() else 'none'

        def border_unit_converter(unit_value: str):
            """Convert multiple units to pt that is used on Word table cell border"""
            unit_value = utils.remove_important_from_style(unit_value)
            unit = re.sub(r'[0-9\.]+', '', unit_value)
            value = float(re.sub(r'[a-zA-Z\!\%]+', '', unit_value))  # Allow float values

            if unit == 'px':
                result = int(value * 0.75)  # 1 px = 0.75 pt
            elif unit == 'cm':
                result = int(value * 28.35)  # 1 cm = 28.35 pt
            elif unit == 'in':
                result = int(value * 72)  # 1 inch = 72 pt
            elif unit == 'pt':
                result = int(value) # default is pt
            elif unit == 'rem' or unit == 'em':
                result = int(value * 12)  # Assuming 1rem/em = 16px, converted to pt
            elif unit == '%':
                result = int(MAX_INDENT * (value / 100))
            else:
                return None  # Unsupported units return None

            return result

        def parse_border_value(value: str):
            """Parses a border value like '1px solid #000000' or '5px'"""
            parts = value.split()
            size = border_unit_converter(parts[0]) if parts else default_size
            style = parse_border_style(parts[1]) if len(parts) > 1 else default_style
            color = utils.parse_color(parts[2], return_hex=True) if len(parts) > 2 else default_color
            return size, style, color

        for style, value in styles.items():
            if "border" not in style:
                continue

            if style == "border":  # Handle shorthand border
                border_values = value.split()
                num_values = len(border_values)

                if num_values == 1:  # "5px"
                    size, style, color = parse_border_value(value)
                    for side in borders:
                        borders[side].update({"size": size, "style": style, "color": color})
                elif num_values == 2:  # "10px 20px"
                    borders["top"].update({"size": border_unit_converter(border_values[0])})
                    borders["bottom"].update({"size": border_unit_converter(border_values[0])})
                    borders["left"].update({"size": border_unit_converter(border_values[1])})
                    borders["right"].update({"size": border_unit_converter(border_values[1])})
                elif num_values == 3:  # "5px solid #000000"
                    size, style, color = parse_border_value(value)
                    for side in borders:
                        borders[side].update({"size": size, "style": style, "color": color})
                elif num_values == 4:  # "1px 5px 10px 20px"
                    borders["top"].update({"size": border_unit_converter(border_values[0])})
                    borders["right"].update({"size": border_unit_converter(border_values[1])})
                    borders["bottom"].update({"size": border_unit_converter(border_values[2])})
                    borders["left"].update({"size": border_unit_converter(border_values[3])})

            elif style in ("border-width", "border-color", "border-style"):
                for side in borders:
                    prop = style.split("-")[-1]
                    if prop == "width":
                        borders[side]["size"] = border_unit_converter(value)
                    elif prop == "color":
                        borders[side]["color"] = utils.parse_color(value, return_hex=True)
                    elif prop == "style":
                        borders[side]["style"] = parse_border_style(value)

            elif re.match(r"^border-(top|right|bottom|left)(-(width|color|style))?$", style):
                parts = style.split("-")
                side = parts[1]
                prop = parts[2] if len(parts) > 2 else None

                if prop == "width":
                    borders[side]["size"] = border_unit_converter(value)
                elif prop == "color":
                    borders[side]["color"] = utils.parse_color(value, return_hex=True)
                elif prop == "style":
                    borders[side]["style"] = parse_border_style(value)
                else:
                    size, style, color = parse_border_value(value)
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
        self.paragraph._element.insert(0, bookmark_start)

        bookmark_end = OxmlElement('w:bookmarkEnd')
        bookmark_end.set(qn('w:id'), str(self.bookmark_id))
        self.paragraph._element.append(bookmark_end)

        self.bookmark_id += 1

    def add_text_align_or_margin_to(self, obj, style):
        if 'text-align' in style:
            align = utils.remove_important_from_style(style['text-align'])

            if 'center' in align:
                obj.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif 'right' in align:
                obj.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif 'justify' in align:
                obj.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        if 'margin-left' in style and 'margin-right' in style:
            margin_left = style['margin-left']
            margin_right = style['margin-right']
            if 'auto' in margin_left and 'auto' in margin_right:
                obj.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif 'margin-left' in style:
            obj.left_indent = utils.unit_converter(style['margin-left'])

    def add_styles_to_table_cell(self, styles, doc_cell, cell_row):
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

        # Set borders
        if any('border' in style for style in styles.keys()):
            self.set_cell_borders(doc_cell, styles)

        self.add_text_align_or_margin_to(doc_cell, styles)

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
            color = utils.parse_color(style['background-color'], return_hex=True)

            # Little trick to apply background-color to paragraph
            # because `self.run.font.highlight_color`
            # has a very limited amount of colors
            #
            # Create XML element
            shd = OxmlElement('w:shd')
            # Add attributes to the element
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), color.lstrip('#'))

            # Make sure the paragraph styling element exists
            self.paragraph.paragraph_format.element.get_or_add_pPr()

            # Append the shading element
            self.paragraph.paragraph_format.element.pPr.append(shd)

    def handle_li(self):
        '''
            Handle li tags
            source: https://stackoverflow.com/a/78685353/17274446
        '''
        list_depth = len(self.tags['list']) or 1
        list_type = self.tags['list'][-1] if self.tags['list'] else 'ul'
        level = min(list_depth, 3)
        style_key = list_type if level <= 1 else f"{list_type}{level}"
        list_style = utils.styles.get(style_key, 'List Number' if list_type == 'ol' else 'List Bullet')

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
        # Available Table Styles
        # https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html#table-styles-in-default-template
        self.table = self.doc.add_table(rows, cols)

        if self.table_style:
            try:
                # Fixed 'style lookup by style_id is deprecated.'
                # https://stackoverflow.com/a/29567907/17274446
                self.table_style = ' '.join(re.findall(r'[A-Z][a-z]*|[0-9]', self.table_style))
                self.table.style = self.table_style
            except KeyError as e:
                raise ValueError(f"Unable to apply style {self.table_style}.") from e

        # Reference:
        # https://python-docx.readthedocs.io/en/latest/api/table.html#cell-objects
        for cell_row, row in enumerate(self.get_table_rows(table_soup)):
            for cell_col, col in enumerate(self.get_table_columns(row)):
                cell_html = self.get_cell_html(col)
                if col.name == 'th':
                    cell_html = "<b>%s</b>" % cell_html

                # Get _Cell object from table based on cell_row and cell_col
                docx_cell = self.table.cell(cell_row, cell_col)

                # Parse cell styles
                cell_styles = utils.parse_dict_string(col.get('style', ''))

                if 'width' in cell_styles or 'height' in cell_styles:
                    self.table.autofit = False

                child_parser = HtmlToDocx()
                child_parser.copy_settings_from(self)
                child_parser.add_html_to_cell(cell_html, docx_cell)
                child_parser.add_styles_to_table_cell(cell_styles, docx_cell, self.table.rows[cell_row])

        if 'style' in current_attrs and self.table:
            style = utils.parse_dict_string(current_attrs['style'])
            self.add_text_align_or_margin_to(self.table, style)

        # skip all tags until corresponding closing tag
        self.instances_to_skip = len(table_soup.find_all('table'))
        self.skip_tag = 'table'
        self.skip = True
        self.table = None

    def handle_div(self, current_attrs):
        # handle page break
        if 'style' in current_attrs and 'page-break-after: always' in current_attrs['style']:
            self.doc.add_page_break()

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
                'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
                'w:pPrChange'
            )
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '6')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), 'auto')
            pBdr.append(bottom)

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
            self.add_text_align_or_margin_to(self.paragraph.paragraph_format, style)

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
            self.table_no += 1
            self.table = None
            self.doc = self.document
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
        link = self.tags.get('a')
        href = link.get('href', None) if link else None
        if link and href:
            self.handle_link(href, data, link.get('title', None))
        else:
            # If there's a link, dont put the data directly in the run
            self.run = self.paragraph.add_run(data)
            spans = self.tags['span']
            for span in spans:
                if 'style' in span:
                    style = utils.parse_dict_string(span['style'])
                    self.add_styles_to_run(style)

            # add font style and name
            for tag in self.tags:
                if tag in utils.font_styles:
                    font_style = utils.font_styles[tag]
                    setattr(self.run.font, font_style, True)

                if tag in utils.font_names:
                    font_name = utils.font_names[tag]
                    self.run.font.name = font_name

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
        #     so it is important to find the maximum number of columns in any row
        if rows:
            cols = max(len(self.get_table_columns(row)) for row in rows)
        else:
            cols = 0
        return len(rows), cols

    def get_tables(self):
        if not hasattr(self, 'soup'):
            self.include_tables = False
            return
            # find other way to do it, or require this dependency?
        self.tables = self.ignore_nested_tables(self.soup.find_all('table'))
        self.table_no = 0

    def run_process(self, html):
        if self.bs and BeautifulSoup:
            self.soup = BeautifulSoup(html, 'html.parser')

            html = str(self.soup)
        if self.include_tables:
            self.get_tables()
        self.feed(html)

    def add_html_to_document(self, html, document):
        if not isinstance(html, str):
            raise ValueError(f'First argument needs to be a {str}')
        elif not isinstance(document, docx.document.Document) and not isinstance(document, docx.table._Cell):
            raise ValueError(f'Second argument needs to be a {docx.document.Document}')
        self.set_initial_attrs(document)
        self.run_process(html)

    def add_html_to_cell(self, html, cell):
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

    def parse_html_file(self, filename_html, filename_docx=None, encoding='utf-8'):
        with open(filename_html, 'r', encoding=encoding) as infile:
            html = infile.read()
        self.set_initial_attrs()
        self.run_process(html)
        if not filename_docx:
            path, filename = os.path.split(filename_html)
            filename_docx = f'{path}/new_docx_file_{filename}'
        self.doc.save(f'{filename_docx}.docx')

    def parse_html_string(self, html):
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
