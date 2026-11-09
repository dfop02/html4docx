# HTML FOR DOCX
![Tests](https://github.com/dfop02/html4docx/actions/workflows/tests.yml/badge.svg)
[![PyPI Downloads](https://static.pepy.tech/personalized-badge/html-for-docx?period=total&units=INTERNATIONAL_SYSTEM&left_color=BLACK&right_color=GREEN&left_text=downloads)](https://pepy.tech/projects/html-for-docx)
![Version](https://img.shields.io/pypi/v/html-for-docx.svg)
![Supported Versions](https://img.shields.io/pypi/pyversions/html-for-docx.svg)

Convert html to docx, this project is a fork from the discontinued [pqzx/html2docx](https://github.com/pqzx/html2docx).

## How to install

`pip install html-for-docx`

## Usage

#### The basic usage

Add HTML-formatted content to an existing `.docx` document

```python
from docx import Document
from html4docx import HtmlToDocx

document = Document('existing_file.docx')
parser = HtmlToDocx()
html_string = '<h1>Hello world</h1>'
parser.add_html_to_document(html_string, document)
document.save('existing_file.docx')
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

When you pass a `Document` object, you can either use `document.save()` from python-docx or `parser.save()` from html4docx, both work well.

Both support saving it in-memory, using `BytesIO`.

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

If the HTML file references local images or stylesheets, allow the folder that holds them (see [External resources security](#external-resources-security)):

```python
import os

parser.options['local-files-root'] = os.path.dirname(os.path.abspath(input_html_file_path))
parser.parse_html_file(input_html_file_path, output_docx_file_path)
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

## Options

All options are set through `parser.options` before parsing. Copy only the ones you need to change:

```python
from html4docx import HtmlToDocx

parser = HtmlToDocx()

# --- Rendering ---
parser.options['images'] = False         # Default True.  False ignores all images.
parser.options['tables'] = False         # Default True.  False renders only the raw table content.
parser.options['styles'] = False         # Default True.  False ignores all CSS (also disables 'style-map').
parser.options['fix-html'] = False       # Default True.  False skips BeautifulSoup fixing of missing HTML tags.
parser.options['style-map'] = False      # Default True.  False ignores CSS class to Word style mapping.
parser.options['tag-override'] = False   # Default True.  False ignores HTML tag overrides.
parser.options['html-comments'] = True   # Default False. True renders "<!-- ... -->" comments.

# --- External resources security (see the section below) ---
parser.options['allow-remote-resources'] = False                    # Default True.  False skips all http(s) images and stylesheets.
parser.options['private-network-allowlist'] = ["10.0.5.20", "192.168.10.0/24"]  # Default []. Private IPs/CIDRs that may be fetched.
parser.options['local-files-root'] = '/srv/app/assets'              # Default None.  Directory local files may be read from (None = disabled).
parser.options['max-resource-bytes'] = 5 * 1024 * 1024              # Default 10 MB. Maximum size of each fetched resource.
parser.options['allow-private-network'] = False                     # Default False. True allows EVERY private address and ignores the allowlist: trusted HTML only!

docx = parser.parse_html_string(input_html_file_string)
```

### External resources security

HTML is often untrusted input, so images (`<img src>`) and stylesheets (`<link rel="stylesheet">`) are fetched under a restrictive policy by default:

- **Remote resources:** only `http`/`https` are fetched. `file://`, `ftp://` and any other scheme are always refused. Inline images (`data:image/...`) keep working.
- **Private network:** requests to non-public addresses (loopback, private ranges, link-local such as the cloud metadata endpoint `169.254.169.254`, reserved, multicast) are refused. The check is applied to the IP actually connected to, on every redirect hop, so DNS rebinding and open redirects can't bypass it.
- **Local files:** disabled by default. When `local-files-root` is set, only files inside that directory (after resolving `..` and symlinks) can be read, and relative paths are resolved against it.
- **Size:** each resource, remote or local, is limited to `max-resource-bytes`.

**Reaching an internal server.** Add its addresses to `private-network-allowlist`. Only those IPs/CIDRs are opened, and everything else stays blocked. Public addresses (CDNs, websites) never need to be listed. Keep entries as narrow as possible: a broad range such as `169.254.0.0/16` also opens cloud metadata endpoints. Invalid entries raise `ValueError` instead of being silently ignored. `allow-private-network` opens every private address at once and should only be used when the HTML is fully trusted.

> **Converting HTML from users?** Keep `allow-private-network` disabled, keep `private-network-allowlist` as small as possible and never point `local-files-root` at a directory containing sensitive files. If you don't need external images or stylesheets at all, set `allow-remote-resources` to `False`.

Blocked or unavailable resources are skipped and a warning is logged (logger `html4docx.resources`); the conversion itself does not fail.

#### Migrating from <=1.2.x

Public `http`/`https` URLs and `data:image/...` images work as before. What changed:

- **Local paths** in `<img src>` and `<link href>` are disabled by default. Set `local-files-root` to the folder that holds your assets. Relative paths are now resolved against it, not the current working directory:

```python
import os

parser.options['local-files-root'] = os.path.abspath('./assets')
```

- **Private network addresses** (e.g. an intranet server) are blocked. Add them to `private-network-allowlist`.
- **Proxies:** `HTTP_PROXY` / `HTTPS_PROXY` environment variables are ignored, because routing through a proxy would bypass the private network protection. If your server can only reach the internet through a proxy, remote resources will not load. Embed images as `data:image/...` URIs, put CSS in an inline `<style>` tag, or serve the files from `local-files-root`.

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

**Custom styles from a Word template:** Use a document created from a `.docx` that already defines the styles (e.g. "Code Block", "Custom Markdown"). Pass **that same document** to the parser and save it so the custom styles are preserved:

```python
from docx import Document
from html4docx import HtmlToDocx

doc = Document("path/to/template.docx")  # template has Code Block, Custom Markdown, etc.
parser = HtmlToDocx(tag_style_overrides={"code": "Custom Markdown", "pre": "Code Block"})
parser.add_html_to_document(html, doc)
doc.save("output.docx")  # save the template-based doc so custom styles are preserved
```

If you save a different document (for example, by creating a new `Document()` instead of loading your template), the output file will **not** contain the template’s custom styles.

If a referenced custom style does not exist in the document at generation time, a warning will be logged to help you detect the missing style.

### Default Paragraph Style

Set custom default paragraph style:

```python
# Word's 'Normal' style is used by default
parser = HtmlToDocx()

# Use a custom style, e.g. 'Body', as default
parser = HtmlToDocx(default_paragraph_style='Body')
```

## CSS

#### Inline CSS Styles

Full support for inline CSS styles on any element:

```html
<p style="color: red; font-size: 14pt">Red 14pt paragraph</p>
<span style="font-weight: bold; color: blue">Bold blue text</span>
```

#### External CSS Styles

Limited support for external CSS styles (local or url). Remote stylesheets follow the [External resources security](#external-resources-security) rules, and local ones require `local-files-root`:

**HTML**

```html
<head>
    <title>My page</title>
    <!-- Must contain rel="stylesheet" otherwise will be ignored -->
    <link rel="stylesheet" href="https://my-external-css.domain.com/style.css">
    <!-- OR, a file inside local-files-root -->
    <link rel="stylesheet" href="css/style.css">
</head>
<body>
    <div class="classFromExternalCss">
        <p>This should be Red and <span class="blueSpan">Blue</span>.<p>
        <p id="specialParagraph" style="font-weight: normal">A very special paragraph.</p>
    </div>
</body>
```

**CSS**

```css
#specialParagraph {
  text-decoration: underline solid red;
  color: gray
}

.blueSpan {
  color: blue;
}

p {
  color: red;
  font-weight: bold;
}
```

### Supported CSS properties:

- color
- font-size
- font-weight (bold)
- font-style (italic)
- text-decoration (underline, line-through)
- font-family
- text-align
- background-color
- Border (for tables)
- Vertical Align (for tables)

### !important Flag Support

Proper CSS precedence with !important:

```html
<span style="color: gray">
  Gray text with <span style="color: red !important">red important</span>.
</span>
```

The `!important` flag ensures highest priority.

### Style Precedence Order

Styles are applied in this order (lowest to highest priority):

1. Base HTML tag styles (`<b>`, `<em>`, `<code>`)
2. Parent span styles
3. CSS class-based styles (from `style_map`)
4. External CSS Styles (from `style` tag or external files)
5. Inline CSS styles (from `style` attribute)
6. `!important` inline CSS styles (highest priority)

## Metadata

You're able to read or set docx metadata:

```python
from docx import Document
from html4docx import HtmlToDocx

document = Document()
parser = HtmlToDocx()
parser.set_initial_attrs(document)
metadata = parser.metadata

# Set new metadata
metadata.set_metadata(author="Jane", created="2025-07-18T09:30:00")

# You can get metadata as dict
metadata_json = metadata.get_metadata()
print(metadata_json['author']) # Jane
# or just print all metadata if you want
metadata.get_metadata(print_result=True)

document.save('your_file_name.docx')
```

You can find all available metadata attributes [here](https://python-docx.readthedocs.io/en/latest/dev/analysis/features/coreprops.html).

## Logging

html4docx uses Python's standard `logging` module with named, hierarchical loggers (for example, `html4docx.h4d` and `html4docx.resources`). The library never logs directly to the root logger and installs a `NullHandler` by default, so it remains silent unless your application configures logging.

**Silence all html4docx logs:**

```python
import logging

logging.getLogger("html4docx").setLevel(logging.ERROR) # suppresses WARNING and below errors
```

**Enable debug logging:**

```python
import logging

logging.basicConfig(level=logging.DEBUG)
logging.getLogger("html4docx").setLevel(logging.DEBUG)
```

**Django / framework `LOGGING` dict** — add an entry for the `html4docx` parent and it applies to all sub-loggers:

```python
LOGGING = {
    "version": 1,
    "loggers": {
        "html4docx": {
            "level": "ERROR",   # suppresses WARNING and below
            "propagate": False,
        },
    },
}
```

> **Note:** Unrecognised CSS properties (e.g. `letter-spacing`, `margin`, `padding`) are intentionally logged at `DEBUG` level because they are expected skips for any real-world HTML, not problems. Only genuinely unexpected situations (missing styles, unsupported colour formats, blocked external resources, etc.) are logged at `WARNING`.

## Why

My goal in forking and fixing/updating this package was to complete my current task at work, which involves converting HTML to DOCX. The original package lacked a few features and had some bugs, preventing me from completing the task. Instead of creating a new package from scratch, I preferred to update this one.

## Differences (fixes and new features)

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
- Fixed styles was only being applied to span tag. | [Dfop02](https://github.com/dfop02) from [Issue](https://github.com/dfop02/html4docx/issues/40)
- Fixed bug on styles parsing when style contains multiple colon. | [Dfop02](https://github.com/dfop02)
- Fixed highlighting a single word | [Lynuxen](https://github.com/Lynuxen)
- Fix color parsing failing due to invalid colors, falling back to black. | [dfop02](https://github.com/dfop02) from [Issue](https://github.com/dfop02/html4docx/issues/53)
- Fix logging noise: replace root-logger calls with named module loggers so consumers can silence or configure html4docx output independently. | [dfop02](https://github.com/dfop02) from [Issue](https://github.com/dfop02/html4docx/issues/80)
- Fixed non-breaking spaces (`&nbsp;`) being collapsed into ordinary spaces, which let Word break lines inside amounts like `1 000,00 €` | [Lynuxen](https://github.com/Lynuxen)
- Fix SSRF and arbitrary local file read through `<img src>` and `<link rel="stylesheet">` when converting untrusted HTML | [dfop02](https://github.com/dfop02)

**New Features**
- Add Width/Height style to images | [maifeeulasad](https://github.com/maifeeulasad) from [PR](https://github.com/pqzx/html2docx/pull/29)
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
- Refactor Tests to be more consistent and less 'human validation' | [Dfop02](https://github.com/dfop02)
- Support for common CSS properties for text | [Lynuxen](https://github.com/Lynuxen)
- Support for CSS classes to Word Styles | [raithedavion](https://github.com/raithedavion)
- Support for HTML tag style overrides | [raithedavion](https://github.com/raithedavion)
- Support style tag and external CSS styles | [Dfop02](https://github.com/dfop02)
- Security policy for external resources (`allow-remote-resources`, `private-network-allowlist`, `allow-private-network`, `local-files-root`, `max-resource-bytes`) | [dfop02](https://github.com/dfop02)

## To-Do

These are the ideas I'm planning to work on in the future to make this project even better:

- No ideas for now.

## Known Limitations

### Lists
- **Maximum Nesting Depth:** Ordered lists support up to 3 nested levels. Any additional depth beyond level 3 will be treated as level 3.
- **Counter Reset Behavior:**
  - At level 1, starting a new ordered list will reset the counter.
  - At levels 2 and 3, the counter will continue from the previous item unless explicitly reset.

### CSS
1. **Complex Selectors:**
   - Do not support complex selector like `div > p`, `h1 + h2`.
   - Complex selectors are simplified to root tag.

2. **Pseudo-classes:**
   - Do not support `:hover`, `:first-child`, etc.
   - They're all removed during CSS Parsing

3. **Media Queries:**
   - Do not support `@media` queries.
   - Would be necessary a responsive CSS, that's not applied to Docx.

4. **@import:**
   - Do not support `@import` in CSS.
   - We can try implement it in future if necessary.

5. **CSS Variables:**
   - Do not support `var(--name)` values. Frameworks that define colors through variables (e.g. Bootstrap 5 `text-danger`) fall back to black.

### External resources
- Proxies configured through `HTTP_PROXY` / `HTTPS_PROXY` are not used (see [Migrating from <=1.2.x](#migrating-from-12x)).
- `data:` URIs are only supported for images. Stylesheets must be remote, local (inside `local-files-root`) or inline in a `<style>` tag.

## Project Guidelines

This project is primarily designed for compatibility with Microsoft Word, but it currently works well with LibreOffice and Google Docs, based on our testing. The goal is to maintain this cross-platform harmony while continuing to implement fixes and updates.

> ⚠️ However, please note that Microsoft Word is the priority. Bugs or issues specific to other editors (e.g., LibreOffice or Google Docs) may be considered, but fixing them is secondary to maintaining full compatibility with Word.

## Security

If you find a security vulnerability, please **do not open a public issue**. Report it privately through [GitHub Security Advisories](https://github.com/dfop02/html4docx/security/advisories/new).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
