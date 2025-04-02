import os
from html4docx import HtmlToDocx
from .context import test_dir

# Manual test (requires inspection of result) for converting html with nested tables

filename = os.path.join(f'{test_dir}/assets/htmls', 'tables1.html')
d = HtmlToDocx()

d.parse_html_file(filename, f'{test_dir}/tables1')
