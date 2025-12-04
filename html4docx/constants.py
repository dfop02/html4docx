# values in inches
from docx.enum.text import WD_UNDERLINE

INDENT = 0.25
LIST_INDENT = 0.5
MAX_INDENT = 5.5  # To stop indents going off the page

# Style to use with tables. By default no style is used.
DEFAULT_TABLE_STYLE = None

DEFAULT_OPTIONS = {
    'fix-html': True,
    'images': True,
    'tables': True,
    'styles': True,
    'html-comments': False,
    "style-map": True,
    "tag-override": True,
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

FONT_UNDERLINE = {
    'underline',
    'line-through',
}

FONT_UNDERLINE_STYLES = {
    'solid': WD_UNDERLINE.SINGLE,
    'dashed': WD_UNDERLINE.DASH,
    'dotted': WD_UNDERLINE.DOTTED,
    'wavy': WD_UNDERLINE.WAVY,
    'double': WD_UNDERLINE.DOUBLE,
}

# NEW DEFAULTS
DEFAULT_STYLE_MAP = dict()
DEFAULT_TAG_OVERRIDES = dict()
DEFAULT_PARAGRAPH_STYLE = "Normal"

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

GENERIC_FONT_STYLES = {
    'serif': 'Times New Roman',
    'sans-serif': 'Arial',
    'monospace': 'Courier New'
}

# Paragraph-level styles (ParagraphFormat)
PARAGRAPH_FORMAT_STYLES = {
    'text-align': '_apply_alignment_paragraph',
    'line-height': '_apply_line_height_paragraph',
    'margin-left': '_apply_margins_paragraph',
    'margin-right': '_apply_margins_paragraph',
    'text-indent': '_apply_text_indent_paragraph',
}

# Run-level styles (affect text formatting within runs)
PARAGRAPH_RUN_STYLES = {
    'font-weight': '_apply_font_weight_paragraph',
    'font-style': '_apply_font_style_paragraph',
    'text-decoration': '_apply_text_decoration_paragraph',
    'text-decoration-line': '_apply_text_decoration_paragraph',
    'text-decoration-style': '_apply_text_decoration_paragraph',
    'text-decoration-color': '_apply_text_decoration_paragraph',
    'text-transform': '_apply_text_transform_paragraph',
    'font-size': '_apply_font_size_paragraph',
    'font-family': '_apply_font_family_paragraph',
    'color': '_apply_color_paragraph',
    'background-color': '_apply_background_color_paragraph'
}

RUN_STYLES = {
    'font-weight': '_apply_font_weight_to_run',
    'font-style': '_apply_font_style_to_run',
    'text-decoration': '_apply_text_decoration_to_run',
    'text-decoration-line': '_apply_text_decoration_line_to_run',
    'text-decoration-style': '_apply_text_decoration_style_to_run',
    'text-decoration-color': '_apply_text_decoration_color_to_run',
    'text-transform': '_apply_text_transform_to_run',
    'font-size': '_apply_font_size_to_run',
    'font-family': '_apply_font_family_to_run',
    'color': '_apply_color_to_run',
    'background-color': '_apply_background_color_to_run'
}

def default_borders():
    return {
        "top": {"size": DEFAULT_BORDER_SIZE, "color": DEFAULT_BORDER_COLOR, "style": DEFAULT_BORDER_STYLE},
        "right": {"size": DEFAULT_BORDER_SIZE, "color": DEFAULT_BORDER_COLOR, "style": DEFAULT_BORDER_STYLE},
        "bottom": {"size": DEFAULT_BORDER_SIZE, "color": DEFAULT_BORDER_COLOR, "style": DEFAULT_BORDER_STYLE},
        "left": {"size": DEFAULT_BORDER_SIZE, "color": DEFAULT_BORDER_COLOR, "style": DEFAULT_BORDER_STYLE}
    }
