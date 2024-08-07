# HTML FOR DOCX
Convert html to docx, this project is a fork from descontinued [pqzx/html2docx](https://github.com/pqzx/html2docx).

### How install

`pip install html-for-docx`

### Usage

The basic usage: Add HTML formatted to an existing Docx

```python
from html4docx import HtmlToDocx

parser = HtmlToDocx()
html_string = '<h1>Hello world</h1>'
parser.add_html_to_document(html_string, filename_docx)
```

You can use `python-docx` to manipulate the file as well, here an example

```python
from docx import Document
from html4docx import HtmlToDocx

document = Document()
new_parser = HtmlToDocx()

html_string = '<h1>Hello world</h1>'
new_parser.add_html_to_document(html_string, document)

document.save('your_file_name')
```

Convert files directly

```python
from html4docx import HtmlToDocx

new_parser = HtmlToDocx()
new_parser.parse_html_file(input_html_file_path, output_docx_file_path)
```

Convert files from a string

```python
from html4docx import HtmlToDocx

new_parser = HtmlToDocx()
docx = new_parser.parse_html_string(input_html_file_string)
```

Change table styles

Tables are not styled by default. Use the `table_style` attribute on the parser to set a table style. The style is used for all tables.

```python
from html4docx import HtmlToDocx

new_parser = HtmlToDocx()
new_parser.table_style = 'Light Shading Accent 4'
```

To add borders to tables, use the `TableGrid` style:

```python
new_parser.table_style = 'TableGrid'
```

Default table styles can be found
here: https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html#table-styles-in-default-template

### Why

My goal to fork and fix/update this package was to complete my current task at work that envolves manipulating a html to docs which the original couldn't complete because was lacking of few features and bugs, so instead creating a package from zero, I prefer update this one.

### Differences (fixes and new features)

**Fixes**
- Handle missing run for leading br tag | [dashingdove](https://github.com/dashingdove) from [PR](https://github.com/pqzx/html2docx/pull/53)
- Fix base64 images | [djplaner](https://github.com/djplaner) from [Issue](https://github.com/pqzx/html2docx/issues/28#issuecomment-1052736896)
- Handle img tag without src attribute | [johnjor](https://github.com/johnjor) from [PR](https://github.com/pqzx/html2docx/pull/63)
- Fix bug when any style has `!important` | [Dfop02](https://github.com/dfop02)
- Fix 'style lookup by style_id is deprecated.' | [Dfop02](https://github.com/dfop02)

**New Features**
- Add Witdh/Height style to images | [maifeeulasad](https://github.com/maifeeulasad) from [PR](https://github.com/pqzx/html2docx/pull/29)
- Support px, cm, pt and % for style margin-left to paragraphs | [Dfop02](https://github.com/dfop02)
- Improve performance on large tables | [dashingdove](https://github.com/dashingdove) from [PR](https://github.com/pqzx/html2docx/pull/58)
- Support for HTML Pagination | [Evilran](https://github.com/Evilran) from [PR](https://github.com/pqzx/html2docx/pull/39)
- Support Table style | [Evilran](https://github.com/Evilran) from [PR](https://github.com/pqzx/html2docx/pull/39)
- Support alternative encoding | [HebaElwazzan](https://github.com/HebaElwazzan) from [PR](https://github.com/pqzx/html2docx/pull/59)
- Support colors by name | [Dfop02](https://github.com/dfop02)
- Support font_size when text, ex.: small, medium, etc. | [Dfop02](https://github.com/dfop02)
- Support to internal links (Anchor) | [Dfop02](https://github.com/dfop02)
- Refactory Tests to be more consistent and less 'human validation' | [Dfop02](https://github.com/dfop02)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
