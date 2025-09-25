# values in inches
INDENT = 0.25
LIST_INDENT = 0.5
MAX_INDENT = 5.5 # To stop indents going off the page

# Style to use with tables. By default no style is used.
DEFAULT_TABLE_STYLE = None

DEFAULT_OPTIONS = {
    'fix-html': True,
    'images': True,
    'tables': True,
    'styles': True
}

DEFAULT_TABLE_ROW_SELECTORS = [
    'table > tr',
    'table > thead > tr',
    'table > tbody > tr',
    'table > tfoot > tr'
]

FONT_STYLES = {
    'b': 'bold',
    'strong': 'bold',
    'em': 'italic',
    'i': 'italic',
    'u': 'underline',
    's': 'strike',
    'sup': 'superscript',
    'sub': 'subscript',
    'th': 'bold'
}

FONT_NAMES = {
    'code': 'Courier',
    'pre': 'Courier'
}

FONT_SIZES_NAMED = {
    'xx-small': '9px',
    'x-small': '10px',
    'small': '13px',
    'medium': '16px',
    'large': '18px',
    'x-large': '24px',
    'xx-large': '32px'
}

STYLES = {
    'ul': 'List Bullet',
    'ul2': 'List Bullet 2',
    'ul3': 'List Bullet 3',
    'ol': 'List Number',
    'ol2': 'List Number 2',
    'ol3': 'List Number 3'
}

# Default border properties #
DEFAULT_BORDER_SIZE = 0
DEFAULT_BORDER_COLOR = "#000000"
DEFAULT_BORDER_STYLE = "single"
BORDER_STYLES = {
    "none": "none",
    "hidden": "none",
    "initial": "none",
    "solid": "single",
    "dotted": "dotted",
    "dashed": "dashed",
    "double": "double",
    "inset": "inset",
    "outset": "outset"
}
BORDER_KEYWORDS = {
    'thin': '1px',
    'medium': '3px',
    'thick': '5px',
    '0': '0px',
}

def default_borders():
    return {
        "top": {"size": DEFAULT_BORDER_SIZE, "color": DEFAULT_BORDER_COLOR, "style": DEFAULT_BORDER_STYLE},
        "right": {"size": DEFAULT_BORDER_SIZE, "color": DEFAULT_BORDER_COLOR, "style": DEFAULT_BORDER_STYLE},
        "bottom": {"size": DEFAULT_BORDER_SIZE, "color": DEFAULT_BORDER_COLOR, "style": DEFAULT_BORDER_STYLE},
        "left": {"size": DEFAULT_BORDER_SIZE, "color": DEFAULT_BORDER_COLOR, "style": DEFAULT_BORDER_STYLE}
    }
