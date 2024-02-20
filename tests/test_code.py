import os
from html4docx import HtmlToDocx
from .context import test_dir

# Manual test (requires inspection of result) for converting code and pre blocks.

filename = os.path.join(f'{test_dir}/assets/htmls', 'code.html')
d = HtmlToDocx()

d.parse_html_file(filename, f'{test_dir}/code')
