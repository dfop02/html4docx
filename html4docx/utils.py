import os
import re
import urllib.request

from io import BytesIO
from enum import Enum
from urllib.parse import urlparse

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

def px_to_inches(px):
    return px * 0.0104166667

def rgb_to_hex(rgb):
    return '#' + ''.join(f'{i:02X}' for i in rgb)

def adapt_font_size(size):
    if (size in font_sizes_named.keys()):
        return font_sizes_named[size]

    return size

def remove_important_from_style(text):
    return re.sub('!important', '', text, flags=re.IGNORECASE)

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
