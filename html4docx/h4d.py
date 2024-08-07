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
import base64
import urllib

from io import BytesIO
from html.parser import HTMLParser

from bs4 import BeautifulSoup

import docx
from docx import Document
from docx.shared import RGBColor, Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from html4docx import utils
from html4docx.colors import Color

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

    def unit_converter(self, unit: str, value: int):
        result = None
        if unit == 'px':
            result = Inches(min(value // 10 * INDENT, MAX_INDENT))
        elif unit == 'cm':
            result = Cm(min(value // 10 * INDENT, MAX_INDENT) * 2.54)
        elif unit == 'pt':
            result = Pt(min(value // 10 * INDENT, MAX_INDENT) * 72)
        elif unit == '%':
            result = int(MAX_INDENT * (value / 100))

        # When unit is not supported returns None
        return result

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

    def add_styles_to_paragraph(self, style):
        if 'text-align' in style:
            align = utils.remove_important_from_style(style['text-align'])

            if 'center' in align:
                self.paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif 'right' in align:
                self.paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif 'justify' in align:
                self.paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        if 'margin-left' in style and 'margin-right' in style:
            margin_left = style['margin-left']
            margin_right = style['margin-right']
            if 'auto' in margin_left and 'auto' in margin_right:
                self.paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif 'margin-left' in style:
            margin = utils.remove_important_from_style(style['margin-left'])
            units = re.sub(r'[0-9]+', '', margin)
            margin = int(float(re.sub(r'[a-zA-Z\!\%]+', '', margin)))

            self.paragraph.paragraph_format.left_indent = self.unit_converter(units, margin)

    def add_styles_to_table(self, style):
        if 'text-align' in style:
            align = utils.remove_important_from_style(style['text-align'])

            if 'center' in align:
                self.table.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif 'right' in align:
                self.table.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif 'justify' in align:
                self.table.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        if 'margin-left' in style and 'margin-right' in style:
            margin_left = style['margin-left']
            margin_right = style['margin-right']
            if 'auto' in margin_left and 'auto' in margin_right:
                self.table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif 'margin-left' in style:
            margin = utils.remove_important_from_style(style['margin-left'])
            units = re.sub(r'[0-9]+', '', margin)
            margin = int(float(re.sub(r'[a-zA-Z\!\%]+', '', margin)))

            self.table.left_indent = self.unit_converter(units, margin)

    def add_styles_to_run(self, style):
        if 'font-size' in style:
            font_size = utils.remove_important_from_style(style['font-size'])
            # Adapt font_size when text, ex.: small, medium, etc.
            font_size = utils.adapt_font_size(font_size)

            units = re.sub(r'[0-9]+', '', font_size)
            font_size = int(float(re.sub(r'[a-zA-Z\!\%]+', '', font_size)))

            if units == 'px':
                font_size_unit = Inches(utils.px_to_inches(font_size))
            elif units == 'cm':
                font_size_unit = Cm(font_size)
            elif units == 'pt':
                font_size_unit = Pt(font_size)
            else:
                # When unit is not supported
                font_size_unit = None

            if font_size_unit:
                for run in self.paragraph.runs:
                    run.font.size = font_size_unit

        if 'color' in style:
            font_color = utils.remove_important_from_style(style['color'].lower())

            if 'rgb' in font_color:
                color = re.sub(r'[a-z()]+', '', font_color)
                colors = [int(x) for x in color.split(',')]
            elif '#' in font_color:
                color = font_color.lstrip('#')
                colors = RGBColor.from_string(color)
            elif font_color in Color._member_names_:
                colors = Color[font_color].value
            else:
                colors = [0, 0, 0]
                # Set color to black to prevent crashing
                # with inexpected colors

            self.run.font.color.rgb = RGBColor(*colors)

        if 'background-color' in style:
            background_color = utils.remove_important_from_style(style['background-color'].lower())

            if 'rgb' in background_color:
                color = re.sub(r'[a-z()]+', '', background_color)
                colors = [int(x) for x in color.split(',')]
            elif '#' in background_color:
                color = background_color.lstrip('#')
                colors = RGBColor.from_string(color)
            elif background_color in Color._member_names_:
                colors = Color[background_color].value
            else:
                colors = [0, 0, 0]
                # Set color to black to prevent crashing
                # with inexpected colors

            # Little trick to apply background-color to paragraph
            # because `self.run.font.highlight_color`
            # has a very limited amount of colors
            #
            # Create XML element
            shd = OxmlElement('w:shd')
            # Add attributes to the element
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), utils.rgb_to_hex(colors))

            # Make sure the paragraph styling element exists
            self.paragraph.paragraph_format.element.get_or_add_pPr()

            # Append the shading element
            self.paragraph.paragraph_format.element.pPr.append(shd)

    def parse_dict_string(self, string, separator=';'):
        new_string = string.replace(" ", '').split(separator)
        string_dict = dict([x.split(':') for x in new_string if ':' in x])
        return string_dict

    def handle_li(self):
        # check list stack to determine style and depth
        list_depth = len(self.tags['list'])
        if list_depth:
            list_type = self.tags['list'][-1]
        else:
            list_type = 'ul' # assign unordered if no tag

        if list_type == 'ol':
            list_style = utils.styles['LIST_NUMBER']
        else:
            list_style = utils.styles['LIST_BULLET']

        self.paragraph = self.doc.add_paragraph(style=list_style)
        self.paragraph.paragraph_format.left_indent = Inches(min(list_depth * LIST_INDENT, MAX_INDENT))
        self.paragraph.paragraph_format.line_spacing = 1

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
        height = Pt(int(re.sub(r'[^0-9]+$', '', current_attrs['height']))) if 'height' in current_attrs else None
        width = Pt(int(re.sub(r'[^0-9]+$', '', current_attrs['width']))) if 'width' in current_attrs else None

        # fetch image
        src_is_url = utils.is_url(src)
        if src_is_url:
            try:
                image = utils.fetch_image(src)
            except urllib.error.URLError:
                image = None
        else:
            image = src

        # check if image starts with data:.*base64 and
        # convert to bytes ready to insert to docx
        if image and isinstance(image, str) and image.startswith('data:image/'):
            image = image.split(',')[1]
            image = base64.b64decode(image)
            image = BytesIO(image)

        # add image to doc
        if image:
            try:
                if isinstance(self.doc, docx.document.Document):
                    self.doc.add_picture(image, width, height)
                else:
                    self.add_image_to_cell(self.doc, image, width, height)
            except FileNotFoundError:
                image = None

        if not image:
            if src_is_url:
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
            try:
                # Fixed 'style lookup by style_id is deprecated.'
                # https://stackoverflow.com/a/29567907/17274446
                self.table_style = ' '.join(re.findall(r'[A-Z][^A-Z]*', self.table_style))
            except KeyError as e:
                raise ValueError(f"Unable to apply style {self.table_style}.") from e

        rows = self.get_table_rows(table_soup)
        cell_row = 0
        docx_cells = self.table._cells
        for row in rows:
            cols = self.get_table_columns(row)
            cell_col = 0
            for col in cols:
                cell_html = self.get_cell_html(col)
                if col.name == 'th':
                    cell_html = "<b>%s</b>" % cell_html
                docx_cell = docx_cells[cell_col + (cell_row * self.table._column_count)]
                child_parser = HtmlToDocx()
                child_parser.copy_settings_from(self)
                child_parser.add_html_to_cell(cell_html, docx_cell)
                cell_col += 1
            cell_row += 1

        if 'style' in current_attrs and self.table:
            style = self.parse_dict_string(current_attrs['style'])
            self.add_styles_to_table(style)

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
        is_external = href.startswith('http')
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
        elif tag == 'ol' or tag == 'ul':
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
            style = self.parse_dict_string(current_attrs['style'])
            self.add_styles_to_paragraph(style)

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
        elif tag == 'ol' or tag == 'ul':
            utils.remove_last_occurence(self.tags['list'], tag)
            return
        elif tag == 'table':
            self.table_no += 1
            self.table = None
            self.doc = self.document
            self.paragraph = None

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
        if link:
            self.handle_link(link.get('href', None), data, link.get('title', None))
        else:
            # If there's a link, dont put the data directly in the run
            self.run = self.paragraph.add_run(data)
            spans = self.tags['span']
            for span in spans:
                if 'style' in span:
                    style = self.parse_dict_string(span['style'])
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
