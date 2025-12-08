# HTML FOR DOCX
![Tests](https://github.com/dfop02/html4docx/actions/workflows/tests.yml/badge.svg)
[![PyPI Downloads](https://static.pepy.tech/personalized-badge/html-for-docx?period=total&units=INTERNATIONAL_SYSTEM&left_color=BLACK&right_color=GREEN&left_text=downloads)](https://pepy.tech/projects/html-for-docx)
![Version](https://img.shields.io/pypi/v/html-for-docx.svg)
![Supported Versions](https://img.shields.io/pypi/pyversions/html-for-docx.svg)

Convert html to docx, this project is a fork from descontinued [pqzx/html2docx](https://github.com/pqzx/html2docx).

## How install

`pip install html-for-docx`

## Usage

#### The basic usage

Add HTML-formatted content to an existing `.docx` document

```python
from html4docx import HtmlToDocx

parser = HtmlToDocx()
html_string = '<h1>Hello world</h1>'
parser.add_html_to_document(html_string, filename_docx)
```

You can use `python-docx` to manipulate directly the file, here an example

```python
from docx import Document
from html4docx import HtmlToDocx

document = Document()
parser = HtmlToDocx()

html_string = '<h1>Hello world</h1>'
parser.add_html_to_document(html_string, document)

document.save('your_file_name.docx')
```

or incrementally add new html to document and save it when finished, new content will always be added at the end

```python
from docx import Document
from html4docx import HtmlToDocx

document = Document()
parser = HtmlToDocx()

for part in ['First', 'Second', 'Third']:
    parser.add_html_to_document(f'<h1>{part} Part</h1>', document)

parser.save('your_file_name.docx')
```

When you pass a `Document` object, you can either use `document.save()` from python-docx or `parser.save()` from html4docx, both works well.

Both supports saving it in-memory, using `BytesIO`.

```python
from io import BytesIO
from docx import Document
from html4docx import HtmlToDocx

buffer = BytesIO()
document = Document()
parser = HtmlToDocx()

html_string = '<h1>Hello world</h1>'
parser.add_html_to_document(html_string, document)

# Save the document to the in-memory buffer
parser.save(buffer)

# If you need to read from the buffer again after saving,
# you might need to reset its position to the beginning
buffer.seek(0)
```

#### Convert files directly

```python
from html4docx import HtmlToDocx

parser = HtmlToDocx()
parser.parse_html_file(input_html_file_path, output_docx_file_path)
# You can also define a encoding, by default is utf-8
parser.parse_html_file(input_html_file_path, output_docx_file_path, 'utf-8')
```

#### Convert files from a string

```python
from html4docx import HtmlToDocx

parser = HtmlToDocx()
docx = parser.parse_html_string(input_html_file_string)
```

#### Change table styles

Tables are not styled by default. Use the `table_style` attribute on the parser to set a table style before convert html. The style is used for all tables.

```python
from html4docx import HtmlToDocx

parser = HtmlToDocx()
parser.table_style = 'Light Shading Accent 4'
docx = parser.parse_html_string(input_html_file_string)
```

To add borders to tables, use the `Table Grid` style:

```python
parser.table_style = 'Table Grid'
```

All table styles we support can be found [here](https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html#table-styles-in-default-template).

#### Options

There is 5 options that you can use to personalize your execution:
- Disable Images: Ignore all images.
- Disable Tables: If you do it, it will render just the raw tables content
- Disable Styles: Ignore all CSS styles. Also disables Style-Map.
- Disable Fix-HTML: Use BeautifulSoap to Fix possible HTML missing tags.
- Disable Style-Map: Ignore CSS classes to word styles mapping
- Disable Tag-Override: Ignore html tag overrides.
- Disable HTML-Comments: Ignore all "<!-- ... -->" comments from HTML.

This is how you could disable them if you want:

```python
from html4docx import HtmlToDocx

parser = HtmlToDocx()
parser.options['images'] = False # Default True
parser.options['tables'] = False # Default True
parser.options['styles'] = False # Default True
parser.options['fix-html'] = False # Default True
parser.options['html-comments'] = False # Default False
parser.options['style-map'] = False # Default True
parser.options['tag-override'] = False # Default True
docx = parser.parse_html_string(input_html_file_string)
```

## Extended Styling Features

### CSS Class to Word Style Mapping

Map HTML CSS classes to Word document styles:

```python
from html4docx import HtmlToDocx

style_map = {
    'code-block': 'Code Block',
    'numbered-heading-1': 'Heading 1 Numbered',
    'finding-critical': 'Finding Critical'
}

parser = HtmlToDocx(style_map=style_map)
parser.add_html_to_document(html, document)
```

### Tag Style Overrides

Override default tag-to-style mappings:

```python
tag_overrides = {
    'h1': 'Custom Heading 1',  # All <h1> use this style
    'pre': 'Code Block'        # All <pre> use this style
}

parser = HtmlToDocx(tag_style_overrides=tag_overrides)
```

### Default Paragraph Style

Set custom default paragraph style:

```python
# Use 'Body' as default (default behavior)
parser = HtmlToDocx(default_paragraph_style='Body')

# Use Word's default 'Normal' style
parser = HtmlToDocx(default_paragraph_style=None)
```

### Inline CSS Styles

Full support for inline CSS styles on any element:

```html
<p style="color: red; font-size: 14pt">Red 14pt paragraph</p>
<span style="font-weight: bold; color: blue">Bold blue text</span>
```

Supported CSS properties:

- color
- font-size
- font-weight (bold)
- font-style (italic)
- text-decoration (underline, line-through)
- font-family
- text-align
- background-color
- Border (for tables)
- Verticial Align (for tables)

### !important Flag Support

Proper CSS precedence with !important:

```html
<span style="color: gray">
  Gray text with <span style="color: red !important">red important</span>.
</span>
```

The !important flag ensures highest priority.

### Style Precedence Order

Styles are applied in this order (lowest to highest priority):

1. Base HTML tag styles (`<b>`, `<em>`, `<code>`)
2. Parent span styles
3. CSS class-based styles (from `style_map`)
4. Inline CSS styles (from `style` attribute)
5. !important inline CSS styles (highest priority)

#### Metadata

You're able to read or set docx metadata:

```python
from docx import Document
from html4docx import HtmlToDocx

document = Document()
parser = HtmlToDocx()
parser.set_initial_attrs(document)
metadata = parser.metadata

# You can get metadata as dict
metadata_json = metadata.get_metadata()
print(metadata_json['author']) # Jane
# or just print all metadata if if you want
metadata.get_metadata(print_result=True)

# Set new metadata
metadata.set_metadata(author="Jane", created="2025-07-18T09:30:00")
document.save('your_file_name.docx')
```

You can find all available metadata attributes [here](https://python-docx.readthedocs.io/en/latest/dev/analysis/features/coreprops.html).

### Why

My goal in forking and fixing/updating this package was to complete my current task at work, which involves converting HTML to DOCX. The original package lacked a few features and had some bugs, preventing me from completing the task. Instead of creating a new package from scratch, I preferred to update this one.

### Differences (fixes and new features)

**Fixes**
- Fix `table_style` not working | [Dfop02](https://github.com/dfop02) from [Issue](https://github.com/dfop02/html4docx/issues/11)
- Handle missing run for leading br tag | [dashingdove](https://github.com/dashingdove) from [PR](https://github.com/pqzx/html2docx/pull/53)
- Fix base64 images | [djplaner](https://github.com/djplaner) from [Issue](https://github.com/pqzx/html2docx/issues/28#issuecomment-1052736896)
- Handle img tag without src attribute | [johnjor](https://github.com/johnjor) from [PR](https://github.com/pqzx/html2docx/pull/63)
- Fix bug when any style has `!important` | [Dfop02](https://github.com/dfop02)
- Fix 'style lookup by style_id is deprecated.' | [Dfop02](https://github.com/dfop02)
- Fix `background-color` not working | [Dfop02](https://github.com/dfop02)
- Fix crashes when img or bookmark is created without paragraph | [Dfop02](https://github.com/dfop02)
- Fix Ordered and Unordered Lists | [TaylorN15](https://github.com/TaylorN15) from [PR](https://github.com/dfop02/html4docx/pull/16)
- Fixed styles was only being applied to span tag. | [Dfop02](https://github.com/dfop02) from [PR](https://github.com/dfop02/html4docx/issues/40)
- Fixed bug on styles parsing when style contains multiple colon. | [Dfop02](https://github.com/dfop02)
- Fixed highlighting a single word | [Lynuxen](https://github.com/Lynuxen)

**New Features**
- Add Witdh/Height style to images | [maifeeulasad](https://github.com/maifeeulasad) from [PR](https://github.com/pqzx/html2docx/pull/29)
- Support px, cm, pt, in, rem, em, mm, pc and % units for styles | [Dfop02](https://github.com/dfop02)
- Improve performance on large tables | [dashingdove](https://github.com/dashingdove) from [PR](https://github.com/pqzx/html2docx/pull/58)
- Support for HTML Pagination | [Evilran](https://github.com/Evilran) from [PR](https://github.com/pqzx/html2docx/pull/39)
- Support Table style | [Evilran](https://github.com/Evilran) from [PR](https://github.com/pqzx/html2docx/pull/39)
- Support alternative encoding | [HebaElwazzan](https://github.com/HebaElwazzan) from [PR](https://github.com/pqzx/html2docx/pull/59)
- Support colors by name | [Dfop02](https://github.com/dfop02)
- Support font_size when text, ex.: small, medium, etc. | [Dfop02](https://github.com/dfop02)
- Support to internal links (Anchor) | [Dfop02](https://github.com/dfop02)
- Support to rowspan and colspan in tables. | [Dfop02](https://github.com/dfop02) from [Issue](https://github.com/dfop02/html4docx/issues/25)
- Support to 'vertical-align' in table cells. | [Dfop02](https://github.com/dfop02)
- Support to metadata | [Dfop02](https://github.com/dfop02)
- Add support to table cells style (border, background-color, width, height, margin) | [Dfop02](https://github.com/dfop02)
- Being able to use inline images on same paragraph. | [Dfop02](https://github.com/dfop02)
- Refactory Tests to be more consistent and less 'human validation' | [Dfop02](https://github.com/dfop02)
- Support for common CSS properties for text | [Lynuxen](https://github.com/Lynuxen)
- Support for CSS classes to Word Styles | [raithedavion](https://github.com/raithedavion)
- Support for HTML tag style overrides | [raithedavion](https://github.com/raithedavion)

## To-Do

These are the ideas I'm planning to work on in the future to make this project even better:

- Add support for the `<style>` tag, including all classes, and ensure they are correctly applied throughout the file.
- Add support for the `<link>` tag to load external CSS and apply it properly across the file.

## Known Issues

- **Maximum Nesting Depth:** Ordered lists support up to 3 nested levels. Any additional depth beyond level 3 will be treated as level 3.
- **Counter Reset Behavior:**
  - At level 1, starting a new ordered list will reset the counter.
  - At levels 2 and 3, the counter will continue from the previous item unless explicitly reset.

## Project Guidelines

This project is primarily designed for compatibility with Microsoft Word, but it currently works well with LibreOffice and Google Docs, based on our testing. The goal is to maintain this cross-platform harmony while continuing to implement fixes and updates.

> ⚠️ However, please note that Microsoft Word is the priority. Bugs or issues specific to other editors (e.g., LibreOffice or Google Docs) may be considered, but fixing them is secondary to maintaining full compatibility with Word.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
