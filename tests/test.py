import os
from pathlib import Path
import unittest
from docx import Document
from html4docx import HtmlToDocx
from .context import test_dir

class OutputTest(unittest.TestCase):
    @staticmethod
    def clean_up_docx():
        for filename in Path(test_dir).glob("*.docx"):
            filename.unlink()

    @staticmethod
    def get_html_from_file(filename: str):
        file_path = Path(f'{test_dir}/assets/htmls') / Path(filename)
        with open(file_path, 'r') as f:
            html = f.read()
        return html

    @classmethod
    def setUpClass(cls):
        cls.clean_up_docx()
        cls.document = Document()
        cls.text1 = cls.get_html_from_file('text1.html')
        cls.table_html = cls.get_html_from_file('tables1.html')
        cls.table2_html = cls.get_html_from_file('tables2.html')

    @classmethod
    def tearDownClass(cls):
        outputpath = os.path.join(test_dir, 'test.docx')
        cls.document.save(outputpath)

    def setUp(self):
        self.parser = HtmlToDocx()

    def test_html_with_images_links_style(self):
        self.document.add_heading(
            'Test: add regular html with images, links and some formatting to document',
            level=1
        )
        self.parser.add_html_to_document(self.text1, self.document)

    def test_html_with_default_paragraph_style(self):
        self.document.add_heading(
            'Test: add regular html with a default paragraph style defined',
            level=1
        )
        self.parser.paragraph_style = 'Quote'
        self.parser.add_html_to_document(self.text1, self.document)

    def test_add_html_to_table_cell_with_default_paragraph_style(self):
        self.document.add_heading(
            'Test: regular html to table cell with a default paragraph style defined',
            level=1
        )
        self.parser.paragraph_style = 'Quote'
        table = self.document.add_table(2, 2, style='Table Grid')
        cell = table.cell(1, 1)
        self.parser.add_html_to_document(self.text1, cell)

    def test_add_html_to_table_cell(self):
        self.document.add_heading(
            'Test: regular html with images, links, some formatting to table cell',
            level=1
        )
        table = self.document.add_table(2, 2, style='Table Grid')
        cell = table.cell(1, 1)
        self.parser.add_html_to_document(self.text1, cell)

    def test_add_html_skip_images(self):
        self.document.add_heading(
            'Test: regular html with images, but skip adding images',
            level=1
        )
        self.parser.options['images'] = False
        self.parser.add_html_to_document(self.text1, self.document)

        document = self.parser.parse_html_string(self.text1)
        assert any(['Graphic' in paragraph._p.xml for paragraph in document.paragraphs]) is False

    def test_add_html_with_tables(self):
        self.document.add_heading(
            'Test: add html with tables',
            level=1
        )
        self.parser.add_html_to_document(self.table_html, self.document)

    def test_add_html_with_tables_accent_style(self):
        self.document.add_heading(
            'Test: add html with tables with accent',
        )
        self.parser.table_style = 'Light Grid Accent 6'
        self.parser.add_html_to_document(self.table_html, self.document)

    def test_add_html_with_tables_basic_style(self):
        self.document.add_heading(
            'Test: add html with tables with basic style',
        )
        self.parser.table_style = 'TableGrid'
        self.parser.add_html_to_document(self.table_html, self.document)

    def test_add_nested_tables(self):
        self.document.add_heading(
            'Test: add nested tables',
        )
        self.parser.add_html_to_document(self.table2_html, self.document)

    def test_add_nested_tables_basic_style(self):
        self.document.add_heading(
            'Test: add nested tables with basic style',
        )
        self.parser.table_style = 'TableGrid'
        self.parser.add_html_to_document(self.table2_html, self.document)

    def test_add_nested_tables_accent_style(self):
        self.document.add_heading(
            'Test: add nested tables with accent style',
        )
        self.parser.table_style = 'Light Grid Accent 6'
        self.parser.add_html_to_document(self.table2_html, self.document)

    def test_add_html_skip_tables(self):
        # broken until feature readded
        self.document.add_heading(
            'Test: add html with tables, but skip adding tables',
            level=1
        )
        self.parser.options['tables'] = False
        self.parser.add_html_to_document(self.table_html, self.document)

    def test_wrong_argument_type_raises_error(self):
        try:
            self.parser.add_html_to_document(self.document, self.text1)
        except Exception as e:
            assert isinstance(e, ValueError)
            assert "First argument needs to be a <class 'str'>" in str(e)
        else:
            assert False, "Error not raised as expected"

        try:
            self.parser.add_html_to_document(self.text1, self.text1)
        except Exception as e:
            assert isinstance(e, ValueError)
            assert "Second argument" in str(e)
            assert "<class 'docx.document.Document'>" in str(e)
        else:
            assert False, "Error not raised as expected"

    def test_add_html_to_cells_method(self):
        self.document.add_heading(
            'Test: add_html_to_cells method',
            level=1
        )
        table = self.document.add_table(2, 3, style='Table Grid')
        cell = table.cell(0, 0)
        html = '''Line 0 without p tags<p>Line 1 with P tags</p>'''
        self.parser.add_html_to_cell(html, cell)

        cell = table.cell(0, 1)
        html = '''<p>Line 0 with p tags</p>Line 1 without p tags'''
        self.parser.add_html_to_cell(html, cell)

        cell = table.cell(0, 2)
        cell.text = "Pre-defined text that shouldn't be removed."
        html = '''<p>Add HTML to non-empty cell.</p>'''
        self.parser.add_html_to_cell(html, cell)

    def test_inline_code(self):
        self.document.add_heading(
            'Test: inline code block',
            level=1
        )

        html = "<p>This is a sentence that contains <code>some code elements</code> that " \
               "should appear as code.</p>"
        self.parser.add_html_to_document(html, self.document)

    def test_code_block(self):
        self.document.add_heading(
            'Test: code block',
            level=1
        )

        html = """<p><code>
This is a code block.
  That should be NOT be pre-formatted.
It should NOT retain carriage returns,

or blank lines.
</code></p>"""
        self.parser.add_html_to_document(html, self.document)

    def test_pre_block(self):
        self.document.add_heading(
            'Test: pre block',
            level=1
        )

        html = """<pre>
This is a pre-formatted block.
  That should be pre-formatted.
Retaining any carriage returns,

and blank lines.
</pre>
"""
        self.parser.add_html_to_document(html, self.document)

    def test_handling_hr(self):
        hr_html_example = '<p>paragraph</p><hr><p>paragraph</p>'

        self.document.add_heading(
            'Test: Handling of hr',
            level=1
        )
        # Add on document for human validation
        self.parser.add_html_to_document(hr_html_example, self.document)

        document = self.parser.parse_html_string(hr_html_example)
        assert '<w:pBdr>' in document._body._body.xml

    def test_external_hyperlink(self):
        hyperlink_html_example = "<a href=\"https://www.google.com\">Google External Link</a>"

        self.document.add_heading(
            'Test: Handling external hyperlink',
            level=1
        )
        # Add on document for human validation
        self.parser.add_html_to_document(hyperlink_html_example, self.document)

        document = self.parser.parse_html_string(hyperlink_html_example)
        # Extract external hyperlinks
        external_hyperlinks = []

        for rel in document.part.rels.values():
            if "hyperlink" in rel.reltype:
                external_hyperlinks.append(rel.target_ref)

        assert 'https://www.google.com' in external_hyperlinks
        assert '<w:hyperlink' in document._body._body.xml

    def test_internal_hyperlink(self):
        hyperlink_html_example = (
            "<p><h1 id=\"intro\">Introduction Header</h1></p>"
            "<p>Click here: <a href=\"#intro\" title=\"Link to intro\">Link to intro</a></p>"
        )

        self.document.add_heading(
            'Test: Handling internal hyperlink',
            level=1
        )
        # Add on document for human validation
        self.parser.add_html_to_document(hyperlink_html_example, self.document)

        document = self.parser.parse_html_string(hyperlink_html_example)
        document_body = document._body._body.xml
        assert '<w:bookmarkStart w:id="0" w:name="intro"/>' in document_body
        assert '<w:bookmarkEnd w:id="0"/>' in document_body
        assert '<w:hyperlink w:anchor="intro" w:tooltip="Link to intro">' in document_body

    def test_image_no_src(self):
        self.document.add_heading(
            'Test: Handling img without src',
            level=1
        )
        # Add on document for human validation
        self.parser.add_html_to_document('<img />', self.document)

        document = self.parser.parse_html_string('<img />')
        assert '<image: no_src>' in document.paragraphs[0].text

    def test_font_size(self):
        font_size_html_example = (
            "<p><span style=\"font-size:8px\">paragraph 8px</span></p>"
            "<p><span style=\"font-size: 1cm\">paragraph 1cm</span></p>"
            "<p><span style=\"font-size: 12em !important\">paragraph 12em not supported</span></p>"
            "<p><span style=\"font-size:14pt!important\">paragraph 14pt</span></p>"
            "<p><span style=\"font-size: 16pt!IMPORTANT\">paragraph 16pt</span></p>"
        )

        self.document.add_heading(
            'Test: Font-Size',
            level=1
        )
        # Add on document for human validation
        self.parser.add_html_to_document(font_size_html_example, self.document)

        document = self.parser.parse_html_string(font_size_html_example)
        font_sizes = [str(p.runs[1].font.size) for p in document.paragraphs]
        assert ['76200', '355600', 'None', '177800', '203200'] == font_sizes

    def test_color_by_name(self):
        color_html_example = (
            "<p><span style=\"color:red\">paragraph red</span></p>"
            "<p><span style=\"color: yellow\">paragraph yellow</span></p>"
            "<p><span style=\"color: blue !important\">paragraph blue</span></p>"
            "<p><span style=\"color: green!important\">paragraph green</span></p>"
            "<p><span style=\"color: darkgray!IMPORTANT\">paragraph darkgray</span></p>"
            "<p><span style=\"color: invalidcolor\">paragraph has default black because of invalid color name</span></p>"
        )

        self.document.add_heading(
            'Test: Color by name',
            level=1
        )
        # Add on document for human validation
        self.parser.add_html_to_document(color_html_example, self.document)

        document = self.parser.parse_html_string(color_html_example)
        colors = [str(p.runs[1].font.color.rgb) for p in document.paragraphs]

        assert 'FF0000' in colors # Red
        assert 'FFFF00' in colors # Yellow
        assert '0000FF' in colors # Blue
        assert '008000' in colors # Green
        assert 'A9A9A9' in colors # Darkgray
        assert '000000' in colors # Black

    def test_unbalanced_table(self):
        # A table with more td elements in latter rows than in the first
        self.document.add_heading(
            'Test: Handling unbalanced tables',
            level=1
        )
        self.parser.add_html_to_document(
            "<table>"
            "<tr><td>Hello</td></tr>"
            "<tr><td>One</td><td>Two</td></tr>"
            "</table>",
            self.document
        )

if __name__ == '__main__':
    unittest.main()
