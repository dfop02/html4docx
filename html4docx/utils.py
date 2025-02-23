import os
import re
import urllib.request

from io import BytesIO
from enum import Enum
from urllib.parse import urlparse
from docx.shared import RGBColor, Pt, Cm, Mm, Inches

from html4docx.colors import Color

font_styles = {
    'b': 'bold',
    'strong': 'bold',
    'em': 'italic',
    'i': 'italic',
    'u': 'underline',
    's': 'strike',
    'sup': 'superscript',
    'sub': 'subscript',
    'th': 'bold',
}

font_names = {
    'code': 'Courier',
    'pre': 'Courier',
}

font_sizes_named = {
    'xx-small': '9px',
    'x-small': '10px',
    'small': '13px',
    'medium': '16px',
    'large': '18px',
    'x-large': '24px',
    'xx-large': '32px'
}

styles = {
    'LIST_BULLET': 'List Bullet',
    'LIST_NUMBER': 'List Number',
}

# values in inches
INDENT = 0.25
MAX_INDENT = 5.5 # To stop indents going off the page

class ImageAlignment(Enum):
    LEFT = 1
    CENTER = 2
    RIGHT = 3

def get_filename_from_url(url):
    return os.path.basename(urlparse(url).path)

def is_url(url):
    """
    Not to be used for actually validating a url, but in our use case we only
    care if it's a url or a file path, and they're pretty distinguishable
    """
    parts = urlparse(url)
    return all([parts.scheme, parts.netloc, parts.path])

def rgb_to_hex(rgb):
    return '#' + ''.join(f'{i:02X}' for i in rgb)

def adapt_font_size(size):
    if (size in font_sizes_named.keys()):
        return font_sizes_named[size]

    return size

def remove_important_from_style(text):
    return re.sub('!important', '', text, flags=re.IGNORECASE).strip()

def fetch_image(url):
    """
    Attempts to fetch an image from a url.
    If successful returns a bytes object, else returns None

    :return:
    """
    try:
        with urllib.request.urlopen(url) as response:
            # security flaw?
            return BytesIO(response.read())
    except urllib.error.URLError:
        return None

def parse_dict_string(string: str, separator: str = ';'):
    new_string = re.sub(r'\s+', ' ', string.replace("\n", '')).split(separator)
    string_dict = dict((k.strip(), v.strip()) for x in new_string if ':' in x for k, v in [x.split(':')])
    return string_dict

def unit_converter(unit_value: str, target_unit: str = "pt"):
    """
    Converts a CSS unit value to a target unit (default is 'pt').
    Supported input units: px, pt, in, pc, cm, mm, em, rem, %.
    Supported target units: pt, px, in, cm, mm.

    Args:
        unit_value (str): The value with unit (e.g., "12px", "1.5in").
        target_unit (str): The target unit to convert to (default is "pt").

    Returns:
        Union[float, Pt, Cm, Inches]: The converted value in the target unit,
        clamped to MAX_INDENT. Returns a python-docx class (Pt, Cm, Inches) when possible.

    Sources:
        https://www.w3schools.com/cssref/css_units.php
    """
    # Remove whitespace and convert to lowercase
    unit_value = unit_value.strip().lower()

    # Extract numeric value and unit
    value = float(re.sub(r'[^0-9.]', '', unit_value))  # Extract numeric part
    unit = re.sub(r'[0-9.]', '', unit_value)           # Extract unit part

    # Conversion factors to points (pt)
    conversion_to_pt = {
        "px": value * 0.75,    # 1px = 0.75pt (assuming 96dpi)
        "pt": value * 1.0,     # 1pt = 1pt
        "in": value * 72.0,    # 1in = 72pt
        "pc": value * 12.0,    # 1pc = 12pt
        "cm": value * 28.3465, # 1cm = 28.3465pt
        "mm": value * 2.83465, # 1mm = 2.83465pt
        "em": value * 12.0,    # 1em = 12pt (assuming 1em = 16px)
        "rem": value * 12.0,   # 1rem = 12pt (assuming 1rem = 16px)
        "%": value,            # Percentage is context-dependent; return as-is
    }

    # Convert input value to points (pt)
    if unit in conversion_to_pt:
        value_in_pt = conversion_to_pt[unit]
    else:
        print(f'Warning: unsupported unit {unit}, return None instead.')
        return None

    # Clamp the value to MAX_INDENT (in points)
    value_in_pt = min(value_in_pt, MAX_INDENT * 72.0)  # MAX_INDENT is in inches

    # Convert from points (pt) to the target unit
    conversion_from_pt = {
        "pt": Pt(value_in_pt),              # 1pt = 1pt
        "px": round(value_in_pt / 0.75, 2), # 1pt = 1.33px
        "in": Inches(value_in_pt / 72.0),   # 1pt = 1/72in
        "cm": Cm(value_in_pt / 28.3465),    # 1pt = 0.0353cm
        "mm": Mm(value_in_pt / 2.83465),    # 1pt = 0.3527mm
    }

    if target_unit in conversion_from_pt:
        return conversion_from_pt[target_unit]
    else:
        raise ValueError(f"Unsupported target unit: {target_unit}")

# def unit_converter(unit_value: str):
#     unit_value = remove_important_from_style(unit_value.strip().lower())
#     unit = re.sub(r'[0-9\.]+', '', unit_value)
#     value = float(re.sub(r'[a-zA-Z\!\%]+', '', unit_value))

#     if unit == 'px':
#         result = Inches(min(value // 10 * INDENT, MAX_INDENT))
#     elif unit == 'in':
#         result = Inches(min(value // 10 * INDENT, MAX_INDENT) * 1)
#     elif unit == 'cm':
#         result = Cm(min(value * INDENT, MAX_INDENT) * 2.54)
#     elif unit == 'pt':
#         result = Pt(min(value // 10 * INDENT, MAX_INDENT) * 72)
#     elif unit == 'rem' or unit == 'em':
#         result = Inches(min(value * 16 // 10 * INDENT, MAX_INDENT))  # Assuming 1rem/em = 16px
#     elif unit == '%':
#         result = int(MAX_INDENT * (value / 100))
#     else:
#         print(f'Warning: unsupported unit {unit}, return None instead.')
#         return None

#     return result

def parse_color(color: str, return_hex: bool = False):
    color = remove_important_from_style(color.strip().lower())

    if 'rgb' in color:
        color = re.sub(r'[^0-9,]', '', color)
        colors = [int(x) for x in color.split(',')]
    elif color.startswith('#'):
        color = color.lstrip('#')
        colors = RGBColor.from_string(color)
    elif color in Color.__members__.keys():
        colors = Color[color].value
    else:
        colors = [0, 0, 0]  # Default to black for unexpected colors

    return rgb_to_hex(colors) if return_hex else colors

def remove_last_occurence(ls, x):
    ls.pop(len(ls) - ls[::-1].index(x) - 1)

def remove_whitespace(string, leading=False, trailing=False):
    """Remove white space from a string.

    Args:
        string(str): The string to remove white space from.
        leading(bool, optional): Remove leading new lines when True.
        trailing(bool, optional): Remove trailing new lines when False.

    Returns:
        str: The input string with new line characters removed and white space squashed.

    Examples:

        Single or multiple new line characters are replaced with space.

            >>> remove_whitespace("abc\\ndef")
            'abc def'
            >>> remove_whitespace("abc\\n\\n\\ndef")
            'abc def'

        New line characters surrounded by white space are replaced with a single space.

            >>> remove_whitespace("abc \\n \\n \\n def")
            'abc def'
            >>> remove_whitespace("abc  \\n  \\n  \\n  def")
            'abc def'

        Leading and trailing new lines are replaced with a single space.

            >>> remove_whitespace("\\nabc")
            ' abc'
            >>> remove_whitespace("  \\n  abc")
            ' abc'
            >>> remove_whitespace("abc\\n")
            'abc '
            >>> remove_whitespace("abc  \\n  ")
            'abc '

        Use ``leading=True`` to remove leading new line characters, including any surrounding
        white space:

            >>> remove_whitespace("\\nabc", leading=True)
            'abc'
            >>> remove_whitespace("  \\n  abc", leading=True)
            'abc'

        Use ``trailing=True`` to remove trailing new line characters, including any surrounding
        white space:

            >>> remove_whitespace("abc  \\n  ", trailing=True)
            'abc'
    """
    # Remove any leading new line characters along with any surrounding white space
    if leading:
        string = re.sub(r'^\s*\n+\s*', '', string)

    # Remove any trailing new line characters along with any surrounding white space
    if trailing:
        string = re.sub(r'\s*\n+\s*$', '', string)

    # Replace new line characters and absorb any surrounding space.
    string = re.sub(r'\s*\n\s*', ' ', string)
    # TODO need some way to get rid of extra spaces in e.g. text <span>   </span>  text
    return re.sub(r'\s+', ' ', string)

def delete_paragraph(paragraph):
    # https://github.com/python-openxml/python-docx/issues/33#issuecomment-77661907
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def get_image_alignment(image_style):
    if image_style == 'float: right;':
        return ImageAlignment.RIGHT
    if image_style == 'display: block; margin-left: auto; margin-right: auto;':
        return ImageAlignment.CENTER
    return ImageAlignment.LEFT
