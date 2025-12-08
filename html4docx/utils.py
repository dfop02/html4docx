import base64
import logging
import os
import re
import urllib
import urllib.request
from enum import Enum
from io import BytesIO
from urllib.parse import urlparse

from docx.shared import Cm, Inches, Mm, Pt, RGBColor

from html4docx import constants
from html4docx.colors import Color


class ImageAlignment(Enum):
    LEFT = 1
    CENTER = 2
    RIGHT = 3


def get_filename_from_url(url: str):
    return os.path.basename(urlparse(url).path)

def dict_to_style_string(style_dict):
    """Convert style dictionary back to CSS string"""
    return '; '.join([f'{k}: {v}' for k, v in style_dict.items()])

def is_url(url: str):
    """
    Not to be used for actually validating a url, but in our use case we only
    care if it's a url or a file path, and they're pretty distinguishable
    """
    parts = urlparse(url)
    return all([parts.scheme, parts.netloc, parts.path])


def rgb_to_hex(rgb: str):
    return "#" + "".join(f"{i:02X}" for i in rgb)


def adapt_font_size(size: str):
    if size in constants.FONT_SIZES_NAMED.keys():
        return constants.FONT_SIZES_NAMED[size]

    return size


def remove_important_from_style(text: str):
    return re.sub("!important", "", text, flags=re.IGNORECASE).strip()


def fetch_image(url: str):
    """
    Attempts to fetch an image from a url, with 5s timeout.
    If successful returns a bytes object, else returns None

    :return:
    """
    try:
        with urllib.request.urlopen(url, timeout=5) as response:
            return BytesIO(response.read())
    except urllib.error.URLError:
        return None


def fetch_image_data(src: str):
    """Fetches image data from a URL or local file."""
    if src.startswith("data:image/"):  # Handle Base64
        _, encoded = src.split(",", 1)
        return BytesIO(base64.b64decode(encoded))

    elif is_url(src):  # Handle URLs
        return fetch_image(src)

    else:  # Handle Local Files
        try:
            return open(src, "rb")
        except FileNotFoundError:
            return None


def parse_dict_string(string: str, separator: str = ";"):
    """Parse style string into dict, return empty dict if no style"""
    if not string:
        return dict()

    new_string = re.sub(r"\s+", " ", string.replace("\n", "")).split(separator)
    string_dict = dict(
        (k.strip(), v.strip())
        for x in new_string
        if ":" in x
        for k, v in [x.split(":", 1)]
    )
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
    value = float(re.sub(r"[^0-9.]", "", unit_value))  # Extract numeric part
    unit = re.sub(r"[0-9.]", "", unit_value)  # Extract unit part

    # Conversion factors to points (pt)
    conversion_to_pt = {
        "px": value * 0.75,  # 1px = 0.75pt (assuming 96dpi)
        "pt": value * 1.0,  # 1pt = 1pt
        "in": value * 72.0,  # 1in = 72pt
        "pc": value * 12.0,  # 1pc = 12pt
        "cm": value * 28.3465,  # 1cm = 28.3465pt
        "mm": value * 2.83465,  # 1mm = 2.83465pt
        "em": value * 12.0,  # 1em = 12pt (assuming 1em = 16px)
        "rem": value * 12.0,  # 1rem = 12pt (assuming 1rem = 16px)
        "%": value,  # Percentage is context-dependent; return as-is
    }

    # Convert input value to points (pt)
    if unit in conversion_to_pt:
        value_in_pt = conversion_to_pt[unit]
    else:
        print(f"Warning: unsupported unit {unit}, return None instead.")
        return None

    # Clamp the value to MAX_INDENT (in points)
    value_in_pt = min(
        value_in_pt, constants.MAX_INDENT * 72.0
    )  # MAX_INDENT is in inches

    # Convert from points (pt) to the target unit
    conversion_from_pt = {
        "pt": Pt(value_in_pt),  # 1pt = 1pt
        "px": round(value_in_pt / 0.75, 2),  # 1pt = 1.33px
        "in": Inches(value_in_pt / 72.0),  # 1pt = 1/72in
        "cm": Cm(value_in_pt / 28.3465),  # 1pt = 0.0353cm
        "mm": Mm(value_in_pt / 2.83465),  # 1pt = 0.3527mm
    }

    if target_unit in conversion_from_pt:
        return conversion_from_pt[target_unit]
    else:
        raise ValueError(f"Unsupported target unit: {target_unit}")


def is_color(color: str) -> bool:
    """
    Checks if a color string is a valid color.
    Supports RGB, hex, and color name strings.

    Args:
        color(str): The color string to check.

    Returns:
        bool: True if the color is valid, False otherwise.

    Examples:
        >>> is_color("red")
        True
        >>> is_color("#000000")
        True
        >>> is_color("rgb(0, 0, 0)")
        True
        >>> is_color("000000")
        False
    """
    is_rgb = 'rgb' in color
    is_hex = color.startswith('#')
    is_keyword = color == 'currentcolor'
    is_color_name = color in Color.__members__
    return is_rgb or is_hex or is_keyword or is_color_name


def parse_color(color: str, return_hex: bool = False):
    """
    Parses a color string into a tuple of RGB values.
    Supports RGB, hex, and color name strings.
    Returns a tuple of RGB values by default, or a hex string if return_hex is True.
    """
    color = remove_important_from_style(color.strip().lower())

    if "rgb" in color:
        color = re.sub(r"[^0-9,]", "", color)
        colors = [int(x) for x in color.split(",")]
    elif color.startswith("#"):
        color = color.lstrip("#")
        color = (
            "".join([x + x for x in color]) if len(color) == 3 else color
        )  # convert short hex to full hex
        colors = RGBColor.from_string(color)
    elif color in Color.__members__:
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
        string = re.sub(r"^\s*\n+\s*", "", string)

    # Remove any trailing new line characters along with any surrounding white space
    if trailing:
        string = re.sub(r"\s*\n+\s*$", "", string)

    # Replace new line characters and absorb any surrounding space.
    string = re.sub(r"\s*\n\s*", " ", string)
    # TODO need some way to get rid of extra spaces in e.g. text <span>   </span>  text
    return re.sub(r"\s+", " ", string)


def delete_paragraph(paragraph):
    # https://github.com/python-openxml/python-docx/issues/33#issuecomment-77661907
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None


def get_image_alignment(image_style):
    if image_style == "float: right;":
        return ImageAlignment.RIGHT
    if image_style == "display: block; margin-left: auto; margin-right: auto;":
        return ImageAlignment.CENTER
    return ImageAlignment.LEFT


def check_style_exists(document, style_name):
    try:
        return style_name in document.styles
    except Exception:
        return False

# Moved from h4d.py to here.... was _parse_text_decoration
def parse_text_decoration(text_decoration):
    """Parse text-decoration using regex to preserve color values."""
    # Pattern to match color values (rgb, hex, named colors) or other tokens
    pattern = r"rgb\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*\)|#[\da-fA-F]+|[\w-]+"

    tokens = re.findall(pattern, text_decoration)

    result = {"line_type": None, "line_style": "solid", "color": None}

    for token in tokens:
        if token in constants.FONT_UNDERLINE:
            result["line_type"] = token
        elif token == "none":
            result["line_type"] = "none"
        elif token in constants.FONT_UNDERLINE_STYLES:
            result["line_style"] = token
        elif is_color(token):
            result["color"] = token
        elif token in ("blink", "overline"):
            result["line_style"] = None
            logging.warning("Blink or overline not supported.")

    if result["line_type"] == "line-through" and result["color"] is not None:
        logging.warning(
            f"Word does not support colored strike-through. Color '{result['color']}' will be ignored for line-through."
        )
    return result


def parse_inline_styles(style_string):
    """
    Parse inline CSS styles and separate normal styles from !important ones.

    Args:
        style_string (str): CSS style string (e.g., "color: red; font-size: 12px !important")

    Returns:
        tuple: (normal_styles dict, important_styles dict)
    """
    normal_styles = {}
    important_styles = {}

    if not style_string:
        return normal_styles, important_styles

    # Parse style string into individual declarations
    style_dict = parse_dict_string(style_string)

    for prop, value in style_dict.items():
        # Check if value has !important flag
        if "!important" in value.lower():
            # Remove !important flag and store in important_styles
            clean_value = remove_important_from_style(value)
            important_styles[prop] = clean_value
        else:
            normal_styles[prop] = value

    return normal_styles, important_styles
