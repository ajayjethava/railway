# ctr_pdf_generator.py
import os
import math
import json
import platform
import subprocess
import tempfile
import logging
import traceback
from datetime import datetime
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Rectangle
import numpy as np
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A3
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, Spacer, Table, TableStyle, PageBreak
from reportlab.platypus.flowables import Flowable
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A2, landscape

# ================= LOGGING SETUP =================
def setup_logging(log_level=logging.INFO, log_file=None):
    """Setup comprehensive logging configuration"""
    # Create logger
    logger = logging.getLogger('ctr_pdf_generator')
    logger.setLevel(log_level)
    
    # Clear any existing handlers
    if logger.handlers:
        logger.handlers.clear()
    
    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # File handler if log_file specified
    if log_file:
        # Create log directory if it doesn't exist
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
            
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(log_level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    # Prevent propagation to root logger
    logger.propagate = False
    
    return logger

# Initialize logger
logger = setup_logging(
    log_level=logging.INFO,
    log_file=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs', 'ctr_pdf_generator.log')
)

# ================= ERROR HANDLER DECORATOR =================
def log_exceptions(func):
    """Decorator to log exceptions with full traceback"""
    def wrapper(*args, **kwargs):
        try:
            logger.debug(f"Entering {func.__name__} with args: {args}, kwargs: {kwargs}")
            result = func(*args, **kwargs)
            logger.debug(f"Exiting {func.__name__} successfully")
            return result
        except Exception as e:
            logger.error(f"Exception in {func.__name__}: {str(e)}")
            logger.error(traceback.format_exc())
            raise
    return wrapper

# Import from your existing models if needed
try:
    from app import db
    from app.models import CTRUpload, CTRSummary, CTRDiagram, CTRRowDetail
    logger.info("Successfully imported database models")
except ImportError as e:
    # For standalone testing
    logger.warning(f"Could not import database models: {e}. Running in standalone mode.")

# ================= COLOR PARSING FUNCTIONS =================
@log_exceptions
def parse_color_string(color_str):
    """
    Parse color string from RowDetail sheet.
    Formats:
    - "red" -> (red background, black text)
    - "white red bg white text" -> (red background, white text)
    - "red bg white text" -> (red background, white text)
    - "green" -> (green background, black text)
    - etc.
    
    Returns: (background_color, text_color) tuple, or None if no valid color
    """
    logger.debug(f"Parsing color string: '{color_str}'")
    
    if not color_str or pd.isna(color_str) or str(color_str).strip() == "":
        logger.debug("Empty color string, returning None")
        return None, None  # Return None to indicate no color override
    
     # -------- special rules --------

    
    
    color_str = str(color_str).lower().strip()
    logger.debug(f"Normalized color string: '{color_str}'")
    
    if color_str == "red text":
        return None, colors.red

    if color_str in ["no color", "nocolor"]:
        return colors.white, colors.black
        
    
    # Try to parse complex format like "white red bg white text" or "red bg white text"
    if "bg" in color_str and "text" in color_str:
        parts = color_str.split()
        bg_color = None
        text_color = None
        
        for i, part in enumerate(parts):
            if part == "bg" and i > 0:
                # Previous word is background color
                bg_color_name = parts[i-1]
                bg_color = color_name_to_rgb(bg_color_name)
                logger.debug(f"Found background color '{bg_color_name}' -> {bg_color}")
            
            if part == "text" and i > 0:
                # Previous word is text color
                text_color_name = parts[i-1]
                text_color = color_name_to_rgb(text_color_name)
                logger.debug(f"Found text color '{text_color_name}' -> {text_color}")
        
        if bg_color and text_color:
            logger.debug(f"Parsed complex color: bg={bg_color}, text={text_color}")
            return bg_color, text_color
        elif bg_color:
            logger.debug(f"Parsed color with only background: bg={bg_color}, text=black")
            return bg_color, colors.black
    
    # Simple color name
    color_obj = color_name_to_rgb(color_str)
    if color_obj:
        logger.debug(f"Parsed simple color: {color_obj}")
        

        if color_str in ["navy blue","navyblue","red"]:
            return color_obj, colors.white

        return color_obj, colors.black
        
    
    logger.warning(f"Invalid color string: '{color_str}'")
    return None, None  # Invalid color string

@log_exceptions
def color_name_to_rgb(color_name):
    """Convert color name to reportlab color object"""
    if not color_name:
        return None, None
    
    color_name = str(color_name).lower().strip()
    logger.debug(f"Converting color name: '{color_name}'")
    
    color_map = {
        'red': colors.red,
        'green': colors.green,
        'blue': colors.blue,
        'yellow': colors.yellow,
        'white': colors.white,
        'black': colors.black,
        'gray': colors.gray,
        'grey': colors.grey,
        'lightred': colors.Color(1.0, 0.5, 0.5),  # Light Red
        'lightblue': colors.lightblue,
        'lightgray': colors.lightgrey,
        'light grey': colors.lightgrey,
        'darkgray': colors.darkgrey,
        'darkgrey': colors.darkgrey,
        'cyan': colors.cyan,
        'magenta': colors.magenta,
        'orange': colors.orange,
        'purple': colors.purple,
        'brown': colors.brown,
        'pink': colors.pink,
        'violet': colors.Color(0.58, 0, 0.83),  # RGB for violet
        'indigo': colors.Color(0.29, 0, 0.51),  # RGB for indigo
        'maroon': colors.Color(0.5, 0, 0),      # RGB for maroon
        'navy': colors.Color(0, 0, 0.5),        # RGB for navy
        'olive': colors.Color(0.5, 0.5, 0),     # RGB for olive
        'teal': colors.Color(0, 0.5, 0.5),      # RGB for teal
        'aqua': colors.Color(0, 1, 1),          # RGB for aqua
        'lime': colors.Color(0, 1, 0),          # RGB for lime
        'fuchsia': colors.Color(1, 0, 1),       # RGB for fuchsia
        'silver': colors.Color(0.75, 0.75, 0.75),  # RGB for silver
    }
    
    # Check for specific railway colors
    if color_name in ['skyblue', 'sky blue', 'skyblue']:
        color_obj = colors.skyblue  # Dark Navy Blue
        logger.debug(f"Railway color '{color_name}' -> {color_obj}")
        return color_obj
    elif color_name in ['lightgray', 'light gray']:
        color_obj = colors.lightgrey  # Light Blue
        logger.debug(f"Railway color '{color_name}' -> {color_obj}")
        return color_obj
    elif color_name in ['navy blue', 'navyblue']:
        color_obj = colors.navy  # Light Blue
        logger.debug(f"Railway color '{color_name}' -> {color_obj}")
        return color_obj
    elif color_name in ['lightorange', 'light orange']:
        color_obj = colors.Color(1.0, 0.85, 0.7)  # Light Orange
        logger.debug(f"Railway color '{color_name}' -> {color_obj}")
        return color_obj
    
    color_obj = color_map.get(color_name, None)
    if color_obj:
        logger.debug(f"Found color '{color_name}' in map -> {color_obj}")
    else:
        logger.warning(f"Color name '{color_name}' not found in color map")
    
    return color_obj

# ================= SAFE COLOR PARSING =================
@log_exceptions
def safe_parse_color(color_str):
    """Safely parse color string with error handling"""
    try:
        if not color_str or pd.isna(color_str) or str(color_str).strip() == "":
            return None,None
        
        result = parse_color_string(str(color_str))
        
        # Ensure result is a tuple of colors
        if result is not None and not isinstance(result, tuple):
            logger.warning(f"Invalid color format returned: {result}")
            return None,None
        
        return result
    except Exception as e:
        logger.error(f"Error parsing color '{color_str}': {e}")
        return None,None

# ================= CTR DIAGRAM CONFIG & FUNCTIONS =================
CTR_SYMBOL_HEIGHT = 0.6
CTR_SYMBOL_WIDTH = 0.35
CTR_SYMBOL_RADIUS = 0.15
CTR_TOTAL_COLUMNS = 15
CTR_PAGE_W, CTR_PAGE_H = 11.69, 8.27  # A4 landscape (inches)

@log_exceptions
def vtext(ax, x, y, text, fontsize=7, rotation=0, max_dim=None):
    """
    Add text with optional auto-sizing.
    If max_dim is provided (in inches), the font size will be reduced so that
    the text fits within that dimension (width for rotation=0, height for rotation=90).
    """
    if pd.isna(text) or str(text).strip() == "":
        text = "SP"
    if max_dim is not None:
        # Approximate character width/height factor (points per character)
        char_factor = 0.65
        text_len = len(text)
        # Required font size in points
        if rotation == 0:
            # max_dim is maximum width in inches
            required_fs = (max_dim * 72) / (text_len * char_factor)
        else:
            # max_dim is maximum height in inches (for rotated text)
            required_fs = (max_dim * 72) / (text_len * char_factor)
        # Use the smaller of the requested fontsize and the required one, but not below 4
        fontsize = max(4, min(fontsize, required_fs))
        logger.debug(f"Auto-sized '{text}' to {fontsize:.1f}pt (max_dim={max_dim:.2f}\")")
    ax.text(x, y, text, fontsize=fontsize, rotation=rotation,
            ha='center', va='center')

@log_exceptions
def generate_ctr_diagram_from_df(df, output_image="ctr_diagram.png"):
    """Generate CTR diagram from DataFrame - more tolerant column detection"""
    try:
        logger.info(f"Starting CTR diagram generation. DataFrame shape: {df.shape}")
        
        # Normalize column names
        df.columns = [str(c).lower().strip().replace(' ', '').replace('_', '').replace('#', 'no') 
                      for c in df.columns]
        logger.debug(f"Normalized columns: {list(df.columns)}")
        
        # Flexible column name detection for terminal
        possible_term_cols = [
            'terminalno', 'terminal_no', 'terminal', 'termno', 'term_no',
            'terminalnumber', 'termnumber', 'terminallo', 'term', 'no'
        ]
        
        term_col = None
        for col in df.columns:
            if col in possible_term_cols or 'terminal' in col or 'term' in col:
                term_col = col
                logger.info(f"Using terminal column: '{col}'")
                break
        
        if term_col is None:
            logger.warning("No terminal number column detected! Looking for columns containing 'term' or 'no'")
            logger.debug(f"Available columns: {list(df.columns)}")
            # Fallback: use first column that looks numeric
            for col in df.columns:
                if df[col].dtype in ['int64', 'float64'] or (hasattr(df[col], 'str') and df[col].str.isnumeric().any()):
                    term_col = col
                    logger.info(f"Fallback terminal column: '{col}'")
                    break
        
        if term_col is None:
            logger.error("Could not find any suitable terminal column")
            return None
        
        logger.info(f"Creating figure with size: {CTR_PAGE_W}x{CTR_PAGE_H}")
        fig = plt.figure(figsize=(CTR_PAGE_W, CTR_PAGE_H))
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_xlim(0, CTR_PAGE_W)
        ax.set_ylim(0, CTR_PAGE_H)
        ax.axis('off')
       
        #xs = np.linspace(1, CTR_PAGE_W-1, CTR_TOTAL_COLUMNS)
        xs = np.linspace(1, CTR_PAGE_W-1, len(df))  # <-- only as many columns as rows
        col_width = xs[1] - xs[0]  # approximate column width in inches
        
        y_pos, y_fuse, y_func, y_cap, y_neg = (
            CTR_PAGE_H-1.2, CTR_PAGE_H-2.2, CTR_PAGE_H-3.4, CTR_PAGE_H-4.6, CTR_PAGE_H-5.8
        )
       
        fuse_pts, cap_pts, fflags, cflags = [], [], [], []
       
        for i, x in enumerate(xs):
            row = df.iloc[i] if i < len(df) else pd.Series()
            
            if i >= len(df):
                logger.debug(f"Row {i+1}: Beyond DataFrame length, using empty row")
           
            # Get terminal number with better fallback
            terminal = "SP"
            if term_col and term_col in row:
                val = row[term_col]
                if pd.notna(val) and str(val).strip() not in ['', 'nan', 'NaN']:
                    terminal = str(val).strip().replace('.0', '')
                    logger.debug(f"Row {i+1}: Terminal value from column: {terminal}")
                elif i < len(df):
                    terminal = str(i+1)
                    logger.debug(f"Row {i+1}: Using index as terminal: {terminal}")
           
            # Draw positive text (horizontal)
            vtext(ax, x, y_pos, row.get('positive', 'SP'), 7, max_dim=col_width*0.9)
           
            # Draw S-FUSE with terminal
            ft, fb, fic, foc = draw_s_fuse(
                ax, x, y_fuse, terminal,
                row.get('fuse_input_left', ''),
                row.get('fuse_input_right', ''),
                row.get('fuse_output_left', ''),
                row.get('fuse_output_right', ''),
                row.get('fuse_input_connected', 'N'),
                row.get('fuse_output_connected', 'N')
            )
           
            draw_vertical_line(ax, x, y_pos-0.35, ft[1])
            draw_vertical_line(ax, x, fb[1], y_func+0.35)
           
            # Draw function text (vertical) – auto‑size to fit vertical space of 0.7 inches
            vtext(ax, x, y_func, row.get('function', 'SP'), 7, rotation=90, max_dim=0.7)
           
            ct, cb, cic, coc = draw_capsule(
                ax, x, y_cap, terminal,
                row.get('capsule_input_left', ''),
                row.get('capsule_input_right', ''),
                row.get('capsule_output_left', ''),
                row.get('capsule_output_right', ''),
                row.get('capsule_input_connected', 'N'),
                row.get('capsule_output_connected', 'N')
            )
           
            draw_vertical_line(ax, x, y_func-0.35, ct[1])
            draw_vertical_line(ax, x, cb[1], y_neg+0.35)
           
            # Draw negative text (horizontal)
            vtext(ax, x, y_neg, row.get('negative', 'SP'), 7, max_dim=col_width*0.9)
           
            fuse_pts.append(fb)
            cap_pts.append(ct)
            fflags.append(foc)
            cflags.append(cic)
       
        for i in range(len(fuse_pts)):
            if i < len(fflags) and i < len(cflags) and fflags[i] == 'Y' and cflags[i] == 'Y':
                draw_vertical_line(ax, fuse_pts[i][0], fuse_pts[i][1], cap_pts[i][1], lw=1.2)
       
        ax.add_patch(Rectangle(
            (0.5, y_neg-0.5), CTR_PAGE_W-1, (y_pos+0.5)-(y_neg-0.5),
            fill=False, lw=1.5
        ))
       
        logger.info(f"Saving CTR diagram to: {output_image}")
        plt.savefig(output_image, dpi=300, bbox_inches='tight', pad_inches=0.1)
        plt.close(fig)
        
        # Check if file was created
        if os.path.exists(output_image):
            file_size = os.path.getsize(output_image)
            logger.info(f"CTR diagram generated successfully: {output_image} ({file_size} bytes)")
        else:
            logger.error(f"Failed to save CTR diagram: {output_image}")
            return None
            
        return output_image
    
    except Exception as e:
        logger.error(f"Error generating CTR diagram: {e}")
        logger.error(traceback.format_exc())
        return None

@log_exceptions
def draw_s_fuse(ax, x, y_center, terminal_no,
                input_left=None, input_right=None,
                output_left=None, output_right=None,
                input_connected='N', output_connected='N'):
    """Draw S-FUSE symbol"""
    logger.debug(f"Drawing S-FUSE at ({x}, {y_center}) with terminal {terminal_no}")
    
    fuse_top = y_center + CTR_SYMBOL_HEIGHT / 2
    fuse_bottom = y_center - CTR_SYMBOL_HEIGHT / 2
    r = CTR_SYMBOL_RADIUS * 0.8
    
    ax.add_patch(Circle((x, fuse_top), r, ec='black', fc='white', lw=1))
    ax.add_patch(Circle((x, fuse_bottom), r, ec='black', fc='white', lw=1))
    
    start = (x, fuse_top - r)
    end = (x, fuse_bottom + r)
    ctrl1 = (x + CTR_SYMBOL_RADIUS * 2.2, y_center + CTR_SYMBOL_HEIGHT * 0.15)
    ctrl2 = (x - CTR_SYMBOL_RADIUS * 2.2, y_center - CTR_SYMBOL_HEIGHT * 0.15)
    
    t = np.linspace(0, 1, 100)
    xs = (1 - t)**3 * start[0] + 3*(1 - t)**2*t*ctrl1[0] + 3*(1 - t)*t**2*ctrl2[0] + t**3*end[0]
    ys = (1 - t)**3 * start[1] + 3*(1 - t)**2*t*ctrl1[1] + 3*(1 - t)*t**2*ctrl2[1] + t**3*end[1]
    ax.plot(xs, ys, color='black', lw=1)
    
    if pd.notna(terminal_no) and str(terminal_no).strip():
        ax.text(x - 0.1, y_center, str(terminal_no).replace('.0', ''),
                fontsize=7, ha='center', va='center')
    
    def fmt(t):
        t = str(t)
        return t[:7] + "\n" + t[7:] if len(t) > 7 else t
    
    if pd.notna(input_left) and str(input_left).strip():
        logger.debug(f"Adding input_left label: {input_left}")
        ax.text(x - 0.005, fuse_top + 0.18, fmt(input_left),
                fontsize=6, rotation=90, ha='right', va='bottom')
    if pd.notna(input_right) and str(input_right).strip():
        logger.debug(f"Adding input_right label: {input_right}")
        ax.text(x + 0.05, fuse_top + 0.18, fmt(input_right),
                fontsize=6, rotation=90, ha='left', va='bottom')
    if pd.notna(output_left) and str(output_left).strip():
        logger.debug(f"Adding output_left label: {output_left}")
        ax.text(x - 0.005, fuse_bottom - 0.15, fmt(output_left),
                fontsize=6, rotation=90, ha='right', va='top')
    if pd.notna(output_right) and str(output_right).strip():
        logger.debug(f"Adding output_right label: {output_right}")
        ax.text(x + 0.05, fuse_bottom - 0.18, fmt(output_right),
                fontsize=6, rotation=90, ha='left', va='top')
    
    return (x, fuse_top + r), (x, fuse_bottom - r), input_connected, output_connected

@log_exceptions
def draw_capsule(ax, x, y, terminal_no, il, ir, ol, or_, ic, oc):
    """Draw capsule symbol"""
    logger.debug(f"Drawing capsule at ({x}, {y}) with terminal {terminal_no}")
    
    top = y + CTR_SYMBOL_HEIGHT / 2
    bottom = y - CTR_SYMBOL_HEIGHT / 2
    r = CTR_SYMBOL_RADIUS * 0.8
    
    ax.add_patch(Circle((x, top), r, ec='black', fc='white', lw=1))
    ax.add_patch(Circle((x, bottom), r, ec='black', fc='white', lw=1))
    
    off = CTR_SYMBOL_WIDTH / 2 - 0.055
    for dx in (-off, off):
        y1 = bottom + np.sqrt(max(0, r*r - dx*dx))
        y2 = top - np.sqrt(max(0, r*r - dx*dx))
        ax.plot([x+dx, x+dx], [y1, y2], color='black', lw=1)
    
    if pd.notna(terminal_no) and str(terminal_no).strip():
        ax.text(x, y, str(terminal_no).replace('.0', ''),
                fontsize=7, ha='center', va='center')
    
    def lbl(t, xp, yp, ha, va):
        if pd.notna(t) and str(t).strip():
            logger.debug(f"Adding label: {t} at ({xp}, {yp})")
            ax.text(xp, yp, t, fontsize=6, rotation=90, ha=ha, va=va)
    
    lbl(il, x-0.005, top+0.18, 'right', 'bottom')
    lbl(ir, x+0.05, top+0.18, 'left', 'bottom')
    lbl(ol, x-0.005, bottom-0.15, 'right', 'top')
    lbl(or_, x+0.05, bottom-0.18, 'left', 'top')
    
    return (x, top+r), (x, bottom-r), ic, oc

@log_exceptions
def draw_vertical_line(ax, x, y1, y2, lw=1):
    """Draw vertical line"""
    logger.debug(f"Drawing vertical line from ({x}, {y1}) to ({x}, {y2})")
    ax.plot([x, x], [y1, y2], lw=lw, color='black')

# ================= FONT REGISTRATION =================
COMMON_FONT_PATHS = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
    "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf",
    "C:\\Windows\\Fonts\\DejaVuSans.ttf",
    "C:\\Windows\\Fonts\\Arial.ttf",
    "./DejaVuSans.ttf",
    "./DejaVuSans-Bold.ttf",
]

@log_exceptions
def register_unicode_font():
    """Register Unicode fonts for PDF generation"""
    logger.info("Registering Unicode fonts...")
    
    regular_name = "Helvetica"
    bold_name = "Helvetica-Bold"
    supports_box = False

    found_regular = found_bold = None
    for p in COMMON_FONT_PATHS:
        if os.path.exists(p) and p.lower().endswith(".ttf"):
            lower = os.path.basename(p).lower()
            if "dejavusans" in lower and "bold" not in lower:
                found_regular = p
            elif "dejavusans" in lower and "bold" in lower:
                found_bold = p
            elif "notosans" in lower and "regular" in lower:
                found_regular = p
            elif "notosans" in lower and "bold" in lower:
                found_bold = p

    if not found_regular:
        for p in COMMON_FONT_PATHS:
            if os.path.exists(p) and p.lower().endswith(".ttf"):
                found_regular = p
                break

    try:
        if found_regular:
            logger.info(f"Found regular font: {found_regular}")
            pdfmetrics.registerFont(TTFont("UserSans", found_regular))
            regular_name = "UserSans"
            if found_bold:
                logger.info(f"Found bold font: {found_bold}")
                pdfmetrics.registerFont(TTFont("UserSans-Bold", found_bold))
                bold_name = "UserSans-Bold"
            else:
                bold_name = regular_name
                logger.warning("No bold font found, using regular font for bold")

            try:
                w_box = pdfmetrics.stringWidth("\u2500", regular_name, 6)
                w_hy = pdfmetrics.stringWidth("-", regular_name, 6)
                supports_box = w_box > 0 and w_box != w_hy
                logger.info(f"Box drawing character support: {supports_box}")
            except Exception as e:
                logger.warning(f"Error checking box character support: {e}")
                supports_box = True
    except Exception as e:
        logger.error(f"Error registering fonts: {e}")

    logger.info(f"Font registration complete: Regular='{regular_name}', Bold='{bold_name}'")
    return regular_name, bold_name, supports_box

FONT_REGULAR, FONT_BOLD, SUPPORTS_BOX = register_unicode_font()

# ================= COLOR MAPPING =================
@log_exceptions
def get_description_color(desc_text: str, color_override=None):
    """
    Returns (background_color, text_color) tuple based on description text.
    Enhanced with railway signaling color coding rules.
    If color_override is provided and valid, use it instead of automatic detection.
    """
    logger.debug(f"Getting color for description: '{desc_text}'")
    
    # If color override is provided and is not None, use it
    if color_override is not None:
        logger.debug(f"Using color override: {color_override}")
        return color_override
    
    text = (desc_text or "").lower().strip()
    
    # ===== NEW ADDITIONS FROM TABLE =====
    
    # 1. Defective Core - White background with Red text
    if any(x in text for x in ["defective", "def core", "defective core", "faulty"]):
        logger.debug("Color: Defective Core -> White background, Red text")
        return colors.white, colors.red
    
    # 2. Spare - Grey background with Black text
    if any(x in text for x in ["spare", "sp.", "unused", "reserved"]):
        logger.debug("Color: Spare -> Grey background, Black text")
        return colors.Color(0.8, 0.8, 0.8), colors.black
    
    # 3. Cable details (location from to) - Yellow background with Black text
    # Example: 2403 JB-3 to PT-103
    if any(x in text for x in ["jb-", "to pt-", "to jb-", "jb to", "cable details", "location"]):
        logger.debug("Color: Cable details -> Yellow background, Black text")
        return colors.yellow, colors.black
    
    # 4. Jumper - Yellow background with Black text
    # Example: From Cable no. XXXX (A,B) to XXXX(C,D)
    if any(x in text for x in ["jumper", "jmp", "jpr", "from cable", "to cable", "jump"]):
        logger.debug("Color: Jumper -> Yellow background, Black text")
        return colors.yellow, colors.black
    
    # 5. Track/Axle counter (TPR & Charger) - Light Orange background with Black text
    # Examples: 101 TPR, 110BX Track CH, 12B AXTPR, Tx, Rx
    if any(x in text for x in ["tpr", "track ch", "axtpr", "axle", "track/axle", 
                               "tpw", "tpr ", "tpwx", "tx", "rx", "track counter",
                               "track circuit", "tc"]):
        logger.debug("Color: Track/Axle counter -> Light Orange background, Black text")
        # Light Orange color (RGB: 1.0, 0.85, 0.7)
        return colors.Color(1.0, 0.85, 0.7), colors.black
    
    # ===== EXISTING SIGNALING COLOR CODING =====
    
    # 1. RG & ICC - Red background with white text
    if any(x in text for x in ["rg", "red green", "r/g", "red/green", "recpr"]):
        logger.debug("Color: RG & ICC -> Red background, White text")
        return colors.red, colors.white
    
    # 2. HG/HHG & ICC - Yellow background with black text
    if any(x in text for x in ["hg", "hhg", "hpr", "hhecr", "hecr"]):
        logger.debug("Color: HG/HHG & ICC -> Yellow background, Black text")
        return colors.yellow, colors.black
    
    # 3. DG & ICC - Green background with black text
    if any(x in text for x in ["dg", "dpr", "decr"]):
        logger.debug("Color: DG & ICC -> Green background, Black text")
        return colors.green, colors.black
    
    # 4. Route, Co, Sign & its ICC, LC, CH, slot & other circuit - White background with black text
    # Example: Co-46 or S46 a "UG" or SH-44 PG
    if any(x in text for x in ["route", "co-", "co ", "sign", "ug", "pg", "lc", "ch", "slot",
                               "s-", "sh-", "co/", "co.", "rte", "signal", "sig"]):
        logger.debug("Color: Route/Co/Sign -> White background, Black text")
        return colors.white, colors.black
    
    # 5. Point Operation - Dark Navy Blue background with white text
    # Examples: PT 101/102 NW/RW/CW (Q series), W1/W2/W3/W4 (Siemens)
    if any(x in text for x in ["point", "pt "]):
        # Check if it's operation (NW, RW, CW, W1, W2, W3, W4)
        if any(x in text for x in ["nw", "rw", "cw", "w1", "w2", "w3", "w4", "op", "operation"]):
            logger.debug("Color: Point Operation -> Dark Navy Blue background, White text")
            # Dark Navy Blue
            return colors.Color(0.0, 0.0, 0.4), colors.white
        # Check if it's detection (NKR, RKR, RWKR, NWKR)
        elif any(x in text for x in ["nkr", "rkr", "rwkr", "nwkr", "detect", "detection"]):
            logger.debug("Color: Point Detection -> Light Blue background, Black text")
            # Light Blue
            return colors.Color(0.85, 0.90, 1.00), colors.black
        # Default for point (light blue for detection)
        else:
            logger.debug("Color: Point (default) -> Light Blue background, Black text")
            return colors.Color(0.85, 0.90, 1.00), colors.black
    
    # Additional existing mappings
    if any(x in text for x in ["bg", "blue green", "b/g", "blue/green"]):
        logger.debug("Color: BG -> Blue background, White text")
        return colors.blue, colors.white
    if any(x in text for x in ["yg", "yellow green", "y/g"]):
        logger.debug("Color: YG -> Yellow background, Black text")
        return colors.yellow, colors.black
    if any(x in text for x in ["w", "white", "wh"]):
        logger.debug("Color: White -> Light Yellow background, Black text")
        return colors.Color(0.96, 0.96, 0.7), colors.black
    if any(x in text for x in ["hr", "hpr", "head red"]):
        logger.debug("Color: HR -> Light Red background, Black text")
        return colors.Color(1.0, 0.75, 0.75), colors.black
    if any(x in text for x in ["ucr", "u", "up"]):
        logger.debug("Color: UCR -> Light Blue background, Black text")
        return colors.Color(0.80, 0.80, 1.00), colors.black
    if any(x in text for x in ["tc", "track", "t/c"]):
        logger.debug("Color: TC -> Light Gray background, Black text")
        return colors.Color(0.94, 0.94, 0.94), colors.black
    
    # Default fallback
    logger.debug("Color: Default -> White background, Black text")
    return colors.white, colors.black

# ================= TERMINAL DIAGRAM FLOWABLE =================
class TerminalDiagram(Flowable):
    def __init__(self, groups, descriptions, cable_names, total_terminals,
                 desc_block_sizes=None, cable_core_numbers=None,
                 desc_colors=None,
                 width=400*mm, row_marker="A", table_shift_right=10*mm, compact_mode=False,
                 is_overflow=False, overflow_index=0, start_terminal=1,
                 super_compact_mode=False):
        Flowable.__init__(self)
        self.groups = groups
        self.descriptions = descriptions
        self.cable_names = cable_names
        self.total_terminals = total_terminals
        self.width = width
        
        # Adjust heights based on mode
        if super_compact_mode:
            # SUPER-COMPACT: For 7 diagrams per page
            self.row_heights = [10*mm, 5*mm, 5*mm, 5*mm]  # Total: 25mm
        elif compact_mode:
            # COMPACT: For 4 diagrams per page
            self.row_heights = [12*mm, 6*mm, 6*mm, 6*mm]  # Total: 30mm
        else:
            # NORMAL: For fewer diagrams per page
            self.row_heights = [14*mm, 8*mm, 8*mm, 8*mm]  # Total: 38mm
        
        self.height = sum(self.row_heights)
        self.marker_width = 10*mm
        self.label_width = 30*mm
        self.desc_block_sizes = desc_block_sizes or [1] * len(groups)
        self.cable_core_numbers = cable_core_numbers or [list(range(1, g+1)) for g in groups]
        self.desc_colors = desc_colors or []
        self.row_marker = row_marker
        self.table_shift_right = table_shift_right
        self.compact_mode = compact_mode
        self.super_compact_mode = super_compact_mode
        self.is_overflow = is_overflow
        self.overflow_index = overflow_index
        self.start_terminal = start_terminal
        
        # Store the cable ranges for drawing cable names
        self.cable_ranges = []
        self._calculate_cable_ranges()
        
        # Store which groups belong to which cable range
        self.group_cable_range_indices = []
        self._map_groups_to_cable_ranges()
        
        if is_overflow:
            logger.info(f"Created Overflow TerminalDiagram: row_marker={row_marker}, total_terminals={total_terminals}, "
                       f"overflow_index={overflow_index}, start_terminal={start_terminal}, groups={len(groups)}")
        else:
            logger.info(f"Created TerminalDiagram: row_marker={row_marker}, total_terminals={total_terminals}, "
                       f"start_terminal={start_terminal}, compact_mode={compact_mode}, super_compact_mode={super_compact_mode}, groups={len(groups)}")

    def _calculate_cable_ranges(self):
        """Calculate the start and end terminal indices for each unique cable"""
        logger.debug("Calculating cable ranges...")
        current_terminal = 0
        current_cable = None
        cable_start = 0
        
        for i, (cable_name, group_size) in enumerate(zip(self.cable_names, self.groups)):
            if cable_name != current_cable:
                # If we have a previous cable, save it
                if current_cable is not None:
                    self.cable_ranges.append((cable_start, current_terminal - 1, current_cable))
                    logger.debug(f"Cable range: {cable_start} to {current_terminal - 1} = '{current_cable}'")
                # Start new cable
                current_cable = cable_name
                cable_start = current_terminal
            
            current_terminal += group_size
            
            # Check if this is the last group
            if i == len(self.groups) - 1:
                self.cable_ranges.append((cable_start, current_terminal - 1, current_cable))
                logger.debug(f"Final cable range: {cable_start} to {current_terminal - 1} = '{current_cable}'")
        
        logger.debug(f"Total cable ranges calculated: {len(self.cable_ranges)}")

    def _map_groups_to_cable_ranges(self):
        """Map each group to its cable range index"""
        logger.debug("Mapping groups to cable ranges...")
        current_terminal = 0
        range_idx = 0
        current_range_start, current_range_end, current_cable = self.cable_ranges[range_idx]
        
        for i, group_size in enumerate(self.groups):
            group_start = current_terminal
            group_end = current_terminal + group_size - 1
            
            # Find which cable range this group belongs to
            while group_end > current_range_end and range_idx < len(self.cable_ranges) - 1:
                range_idx += 1
                current_range_start, current_range_end, current_cable = self.cable_ranges[range_idx]
            
            self.group_cable_range_indices.append(range_idx)
            logger.debug(f"Group {i} (size={group_size}) -> cable range index {range_idx}")
            current_terminal += group_size

    def _wrap_text_to_fit_box(self, text, font_name, font_size, max_width, block_size=None):
        """Wrap text to fit within a given width, using line breaks if needed"""
        if not text:
            return text
            
        canv = self.canv
        canv.setFont(font_name, font_size)
        
        # For compact mode with single terminal blocks, be more aggressive with wrapping
        if self.compact_mode and block_size == 1:
            words = str(text).split()
            if not words:
                return text
                
            lines = []
            current_line = []
            
            for word in words:
                current_line.append(word)
                test_line = ' '.join(current_line)
                line_width = canv.stringWidth(test_line, font_name, font_size)
                
                if line_width > max_width * 0.9:
                    if len(current_line) == 1:
                        # Single word is too long, break it into chunks
                        if len(test_line) > 6:
                            chunks = [test_line[i:i+4] for i in range(0, len(test_line), 4)]
                            for chunk in chunks[:-1]:
                                lines.append(chunk)
                            current_line = [chunks[-1]]
                    else:
                        current_line.pop()
                        lines.append(' '.join(current_line))
                        current_line = [word]
            
            if current_line:
                lines.append(' '.join(current_line))
            
            result = '\n'.join(lines)
            return result
        else:
            # Original wrapping logic for larger blocks
            words = str(text).split()
            if not words:
                return text
                
            lines = []
            current_line = []
            
            for word in words:
                current_line.append(word)
                test_line = ' '.join(current_line)
                line_width = canv.stringWidth(test_line, font_name, font_size)
                
                if line_width > max_width:
                    if len(current_line) == 1:
                        if len(test_line) > 12:
                            chunks = [test_line[i:i+10] for i in range(0, len(test_line), 10)]
                            for chunk in chunks[:-1]:
                                lines.append(chunk)
                            current_line = [chunks[-1]]
                        else:
                            current_line.pop()
                            lines.append(' '.join(current_line))
                            current_line = [word]
                    else:
                        current_line.pop()
                        lines.append(' '.join(current_line))
                        current_line = [word]
            
            if current_line:
                lines.append(' '.join(current_line))
            
            result = '\n'.join(lines)
            return result

    def _get_optimal_font_size(self, text, font_name, max_width, max_height, min_font_size=4.5, block_size=None):
        """Determine the optimal font size to fit text in a box"""
        if not text:
            if self.compact_mode:
                return 3.5 if block_size == 1 else 4.0
            return 4.0 if block_size == 1 else 5.5
        
        # Adjust base font size based on block size and compact mode
        if self.compact_mode:
            if block_size == 1:
                base_font_size = 4.0
                min_font_size = 3.0
            else:
                base_font_size = 4.5
                min_font_size = 3.5
        else:
            if block_size == 1:
                base_font_size = 4.5
                min_font_size = 3.5
            else:
                base_font_size = 5.5
                min_font_size = 4.5
        
        # Count lines
        lines = str(text).count('\n') + 1
        
        # Start with appropriate base size
        font_size = base_font_size
        
        # If text has line breaks, adjust for line count
        if lines > 1:
            font_size = min(font_size, 4.0 if block_size == 1 else 4.5)
        
        # Check if text fits at current size
        canv = self.canv
        canv.setFont(font_name, font_size)
        
        # For multi-line text, check each line
        text_lines = str(text).split('\n')
        max_line_width = 0
        for line in text_lines:
            line_width = canv.stringWidth(line, font_name, font_size)
            max_line_width = max(max_line_width, line_width)
        
        # Reduce font size if text is too wide
        original_size = font_size
        while max_line_width > max_width * 0.9 and font_size > min_font_size:
            font_size -= 0.2
            canv.setFont(font_name, font_size)
            max_line_width = 0
            for line in text_lines:
                line_width = canv.stringWidth(line, font_name, font_size)
                max_line_width = max(max_line_width, line_width)
        
        return font_size

    def draw(self):
        """Draw the terminal diagram with proper text alignment"""
        try:
            if self.is_overflow:
                logger.debug(f"Drawing Overflow TerminalDiagram for row_marker={self.row_marker}, overflow_index={self.overflow_index}, start_terminal={self.start_terminal}")
            else:
                logger.debug(f"Drawing TerminalDiagram for row_marker={self.row_marker}, start_terminal={self.start_terminal}")
            
            canv = self.canv
            avail_width = self.width - 2*self.marker_width - self.label_width
            if self.total_terminals <= 0:
                logger.warning(f"Total terminals is {self.total_terminals}, skipping draw")
                return
            cell_width = avail_width / self.total_terminals

            current_terminal = self.start_terminal
            x_offset = self.marker_width + self.label_width

            # Adjust row positions based on mode
            if self.super_compact_mode:
                # For 7 rows per page
                y_term_bottom = 0  # Terminal numbers row (bottom)
                y_cable_bottom = y_term_bottom + 5*mm  # Cable names row
                y_core_bottom = y_cable_bottom + 5*mm  # Core numbers row
                y_desc_bottom = y_core_bottom + 5*mm   # Description row (top)
                desc_row_height = 10*mm
            elif self.compact_mode:
                # For 4 rows per page
                y_term_bottom = 0
                y_cable_bottom = y_term_bottom + 6*mm
                y_core_bottom = y_cable_bottom + 6*mm
                y_desc_bottom = y_core_bottom + 6*mm
                desc_row_height = 12*mm
            else:
                y_term_bottom = 0
                y_cable_bottom = y_term_bottom + 8*mm
                y_core_bottom = y_cable_bottom + 8*mm
                y_desc_bottom = y_core_bottom + 8*mm
                desc_row_height = 14*mm

            canv.setStrokeColor(colors.black)
            canv.setLineWidth(1)

            # Apply shift to ALL elements
            shift = self.table_shift_right

            # Left marker circle
            canv.rect(0 + shift, 0, self.marker_width, self.height, fill=0, stroke=1)
            cx = shift + self.marker_width / 2
            cy = self.height / 2
            r = min(self.marker_width, self.height) * 0.35
            
            canv.circle(cx, cy, r, stroke=1, fill=0)
            
            # Row marker in circle
            if self.super_compact_mode:
                marker_font_size = 7
            elif self.compact_mode:
                marker_font_size = 8
            else:
                marker_font_size = 10
            
            canv.setFont(FONT_BOLD, marker_font_size)
            # Calculate proper vertical centering for marker
            text_height = marker_font_size * 0.75  # Approximate text height factor
            text_y = cy - (text_height / 2)
            canv.drawCentredString(cx, text_y, str(self.row_marker))

            # Right marker circle
            rx = self.width - self.marker_width
            canv.rect(rx + shift, 0, self.marker_width, self.height, fill=0, stroke=1)
            cx = rx + shift + self.marker_width / 2
            
            canv.circle(cx, cy, r, stroke=1, fill=0)
            canv.drawCentredString(cx, text_y, str(self.row_marker))

            # Labels - with shift
            lx = self.marker_width + self.label_width - 2*mm + shift
            
            # Adjust label font size based on mode
            if self.super_compact_mode:
                label_font_size = 5
            elif self.compact_mode:
                label_font_size = 6
            else:
                label_font_size = 8
            
            canv.setFont(FONT_BOLD, label_font_size)
            
            # Calculate label positions with proper centering
            if self.super_compact_mode:
                # Terminal numbers label
                term_label_y = y_term_bottom + (5*mm - label_font_size*0.75) / 2
                canv.drawRightString(lx, term_label_y, "TERMINAL NO")
                
                # Cable names label
                cable_label_y = y_cable_bottom + (5*mm - label_font_size*0.75) / 2
                canv.drawRightString(lx, cable_label_y, "CABLE NO")
                
                # Core numbers label
                core_label_y = y_core_bottom + (5*mm - label_font_size*0.75) / 2
                canv.drawRightString(lx, core_label_y, "CABLE CORE NO")
                
                # Description label
                desc_label_y = y_desc_bottom + (desc_row_height - label_font_size*0.75) / 2
                canv.drawRightString(lx, desc_label_y, "DESCRIPTION")
            elif self.compact_mode:
                term_label_y = y_term_bottom + (6*mm - label_font_size*0.75) / 2
                canv.drawRightString(lx, term_label_y, "TERMINAL NO")
                
                cable_label_y = y_cable_bottom + (6*mm - label_font_size*0.75) / 2
                canv.drawRightString(lx, cable_label_y, "CABLE NO")
                
                core_label_y = y_core_bottom + (6*mm - label_font_size*0.75) / 2
                canv.drawRightString(lx, core_label_y, "CABLE CORE NO")
                
                desc_label_y = y_desc_bottom + (desc_row_height - label_font_size*0.75) / 2
                canv.drawRightString(lx, desc_label_y, "DESCRIPTION")
            else:
                term_label_y = y_term_bottom + (8*mm - label_font_size*0.75) / 2
                canv.drawRightString(lx, term_label_y, "TERMINAL NO")
                
                cable_label_y = y_cable_bottom + (8*mm - label_font_size*0.75) / 2
                canv.drawRightString(lx, cable_label_y, "CABLE NO")
                
                core_label_y = y_core_bottom + (8*mm - label_font_size*0.75) / 2
                canv.drawRightString(lx, core_label_y, "CABLE CORE NO")
                
                desc_label_y = y_desc_bottom + (desc_row_height - label_font_size*0.75) / 2
                canv.drawRightString(lx, desc_label_y, "DESCRIPTION")

            canv.setLineWidth(0.5)
            # Vertical lines with shift
            canv.line(self.marker_width + self.label_width + shift, 0, 
                     self.marker_width + self.label_width + shift, self.height)
            canv.line(rx + shift, 0, rx + shift, self.height)

            line_char = "\u2500" if SUPPORTS_BOX else "-"

            # Draw description blocks
            desc_block_start = 0
            for g_idx, g_size in enumerate(self.groups):
                g_width = g_size * cell_width
                block_size = self.desc_block_sizes[g_idx]
                n_blocks = math.ceil(g_size / block_size)

                # Draw vertical line ONLY if this group starts a new cable or is the first group
                if g_idx > 0:
                    if g_idx >= len(self.group_cable_range_indices):
                        canv.line(x_offset + shift, 0, x_offset + shift, self.height)
                    else:
                        current_range_idx = self.group_cable_range_indices[g_idx]
                        prev_range_idx = self.group_cable_range_indices[g_idx-1]
                        if current_range_idx != prev_range_idx:
                            canv.line(x_offset + shift, 0, x_offset + shift, self.height)
                        else:
                            canv.line(x_offset + shift, 0, x_offset + shift, y_cable_bottom)
                            if self.super_compact_mode:
                                canv.line(x_offset + shift, y_cable_bottom + 5*mm, x_offset + shift, self.height)
                            elif self.compact_mode:
                                canv.line(x_offset + shift, y_cable_bottom + 6*mm, x_offset + shift, self.height)
                            else:
                                canv.line(x_offset + shift, y_cable_bottom + 8*mm, x_offset + shift, self.height)

                for b in range(n_blocks):
                    start = b * block_size
                    end = min(start + block_size, g_size)
                    bx = x_offset + shift + start * cell_width
                    bw = (end - start) * cell_width

                    desc = self.descriptions[g_idx][b] if b < len(self.descriptions[g_idx]) else ""
                    
                    # Get colors - Fixed condition
                    if (self.desc_colors and 
                        g_idx < len(self.desc_colors) and 
                        len(self.desc_colors[g_idx]) > b and 
                        self.desc_colors[g_idx][b] is not None):
                        fill_c, text_c = self.desc_colors[g_idx][b]
                    else:
                        fill_c, text_c = get_description_color(desc)
                    if fill_c :
                       canv.setFillColor(fill_c)
                    else :
                       canv.setFillColor(colors.white)
                           
                    canv.rect(bx, y_desc_bottom, bw, desc_row_height, fill=1, stroke=1)
                    if text_c :
                       canv.setFillColor(text_c)
                    else :
                       canv.setFillColor(colors.black)
                       
                    # AUTO-TEXT WRAPPING AND SIZING
                    if desc:
                        available_width = bw - 4*mm
                        group_block_size = self.desc_block_sizes[g_idx] if g_idx < len(self.desc_block_sizes) else 1
                        
                        # Determine appropriate font size for wrapping based on mode and block size
                        if self.super_compact_mode:
                            wrap_font_size = 3.0 if group_block_size == 1 else 3.5
                        elif self.compact_mode:
                            wrap_font_size = 3.5 if group_block_size == 1 else 4.0
                        else:
                            wrap_font_size = 4.5 if group_block_size == 1 else 5.5
                        
                        wrapped_text = self._wrap_text_to_fit_box(
                            desc, FONT_BOLD, wrap_font_size, available_width, group_block_size
                        )
                        
                        # Compute optimal font size for the wrapped text
                        optimal_font_size = self._get_optimal_font_size(
                            wrapped_text, FONT_BOLD, available_width, desc_row_height,
                            min_font_size=2.5 if self.super_compact_mode else (3.0 if self.compact_mode else 3.5),
                            block_size=group_block_size
                        )
                        
                        canv.setFont(FONT_BOLD, optimal_font_size)
                        
                        # Split wrapped text into lines
                        text_lines = wrapped_text.split('\n')
                        line_height = optimal_font_size * 1.1  # Reduced line spacing
                        total_text_height = len(text_lines) * line_height
                        
                        # Calculate vertical starting position to center text
                        if len(text_lines) == 1:
                            # Single line: center vertically
                            text_y = y_desc_bottom + (desc_row_height - optimal_font_size*0.75) / 2
                            canv.drawCentredString(bx + bw/2, text_y, text_lines[0])
                        else:
                            # Multiple lines: distribute evenly
                            start_y = y_desc_bottom + (desc_row_height - total_text_height) / 2 + line_height
                            for i, line in enumerate(text_lines):
                                line_y = start_y - (i * line_height)
                                canv.drawCentredString(bx + bw/2, line_y, line)

                # Draw core numbers
                canv.setFillColor(colors.black)
                if self.super_compact_mode:
                    core_font_size = 4.0
                elif self.compact_mode:
                    core_font_size = 5.0
                else:
                    core_font_size = 6.0
                
                canv.setFont(FONT_BOLD, core_font_size)
                for t in range(g_size):
                    tx = x_offset + shift + t * cell_width
                    if self.super_compact_mode:
                        row_height = 5*mm
                    elif self.compact_mode:
                        row_height = 6*mm
                    else:
                        row_height = 8*mm
                    
                    canv.rect(tx, y_core_bottom, cell_width, row_height, fill=0, stroke=1)
                    if g_idx < len(self.cable_core_numbers) and t < len(self.cable_core_numbers[g_idx]):
                        core = self.cable_core_numbers[g_idx][t]
                    else:
                        core = desc_block_start + t + 1
                    
                    # Draw centered core number
                    core_str = str(core)
                    text_y = y_core_bottom + (row_height - core_font_size*0.75) / 2
                    canv.drawCentredString(tx + cell_width/2, text_y, core_str)

                # Draw terminal numbers - using the actual terminal number
                if self.super_compact_mode:
                    term_font_size = 5.0
                elif self.compact_mode:
                    term_font_size = 6.0
                else:
                    term_font_size = 7.5
                
                canv.setFont(FONT_BOLD, term_font_size)
                for t in range(g_size):
                    tx = x_offset + shift + t * cell_width
                    if self.super_compact_mode:
                        row_height = 5*mm
                    elif self.compact_mode:
                        row_height = 6*mm
                    else:
                        row_height = 8*mm
                    
                    canv.rect(tx, y_term_bottom, cell_width, row_height, fill=0, stroke=1)
                    num_str = str(current_terminal)
                    
                    # Draw centered terminal number
                    text_y = y_term_bottom + (row_height - term_font_size*0.75) / 2
                    canv.drawCentredString(tx + cell_width/2, text_y, num_str)
                    current_terminal += 1

                desc_block_start += g_size
                x_offset += g_width

            # Draw cable names (merged for same cable)
            canv.setFillColor(colors.black)
            if self.super_compact_mode:
                cable_font_size = 5
            elif self.compact_mode:
                cable_font_size = 6
            else:
                cable_font_size = 8
            
            canv.setFont(FONT_BOLD, cable_font_size)
            for cable_start, cable_end, cable_name in self.cable_ranges:
                start_x = self.marker_width + self.label_width + shift + cable_start * cell_width
                cable_width = (cable_end - cable_start + 1) * cell_width
                
                # Draw the rectangle for the cable name area
                if self.super_compact_mode:
                    row_height = 5*mm
                elif self.compact_mode:
                    row_height = 6*mm
                else:
                    row_height = 8*mm
                
                canv.rect(start_x, y_cable_bottom, cable_width, row_height, fill=0, stroke=1)
                
                # Draw cable name with formatting
                padding = 6*mm
                avail = max(0, cable_width - padding)
                tw = canv.stringWidth(cable_name, FONT_BOLD, cable_font_size)
                cw = canv.stringWidth(line_char, FONT_BOLD, cable_font_size)
                sw = canv.stringWidth("  ", FONT_BOLD, cable_font_size)
                rem = avail - tw - sw*2

                if rem <= 0 or cw <= 0:
                    txt = f"<  {cable_name}  >"
                else:
                    n = max(2, int(rem // cw))
                    left = n // 2
                    right = n - left
                    txt = "<" + line_char*left + "  " + cable_name + "  " + line_char*right + ">"
                
                # Draw centered cable name
                text_y = y_cable_bottom + (row_height - cable_font_size*0.75) / 2
                canv.drawCentredString(start_x + cable_width/2, text_y, txt)

            canv.setStrokeColor(colors.black)
            canv.setLineWidth(1)
            # Main rectangle with shift
            canv.rect(0 + shift, 0, self.width, self.height)
            canv.setLineWidth(0.5)
            sx = self.marker_width + shift
            ex = self.width - self.marker_width + shift
            canv.line(sx, y_desc_bottom, ex, y_desc_bottom)
            canv.line(sx, y_core_bottom, ex, y_core_bottom)
            canv.line(sx, y_cable_bottom, ex, y_cable_bottom)
            
            if self.is_overflow:
                logger.debug(f"Successfully drew Overflow TerminalDiagram for row_marker={self.row_marker}, overflow_index={self.overflow_index}, start_terminal={self.start_terminal}")
            else:
                logger.debug(f"Successfully drew TerminalDiagram for row_marker={self.row_marker}, start_terminal={self.start_terminal}")
            
        except Exception as e:
            logger.error(f"Error drawing TerminalDiagram: {e}")
            logger.error(traceback.format_exc())
            raise

# ================= FOOTER FUNCTION FOR PDF =================
@log_exceptions
def add_pdf_footer(canvas, doc, footer_data, total_pages):
    """Add footer to PDF pages, now including version number if available"""
    try:
        logger.debug(f"Adding footer to page {canvas.getPageNumber()}")
        
        canvas.saveState()
        
        page_width, page_height = doc.pagesize
        footer_y = 40
        footer_height = 90
        
        # Keep the footer box on the right side
        footer_x_start = page_width - 370
        footer_width = 340
        
        # Left margin for the long horizontal lines
        page_left_margin = 50 

        page_num = canvas.getPageNumber()
        
        # Get values with defaults
        des1 = footer_data.get('designation1', 'Sr DSTE/ADI')
        des2 = footer_data.get('designation2', 'ADSTE/ADI')
        des3 = footer_data.get('designation3', 'SSE/SIG')
        station = footer_data.get('station_name', 'VASADVA STN').replace('STATION', 'STN').upper().strip()
        track = footer_data.get('junction_name', 'C.TRACK -1')
        drawing_no = footer_data.get('station_code', 'SC/PL 411/14')
        zone = footer_data.get('zone', 'WESTERN')
        division = footer_data.get('division', 'AHMEDABAD')
        version = footer_data.get('version', '')  # New version field
        
        division_line1 = f"{zone} RLY."
        division_line2 = f"{division} DIVISION"
        
        logger.debug(f"Footer data: station={station}, track={track}, drawing_no={drawing_no}, version={version}")
        
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(1)

        # Horizontal lines (full width from left margin to right edge)
        canvas.line(page_left_margin, footer_y + footer_height, footer_x_start + footer_width, footer_y + footer_height)
        canvas.line(page_left_margin, footer_y, footer_x_start + footer_width, footer_y)

        # Vertical lines inside the footer box (original short ones)
        canvas.line(footer_x_start, footer_y, footer_x_start, footer_y + footer_height)
        canvas.line(footer_x_start + footer_width, footer_y, footer_x_start + footer_width, footer_y + footer_height)
        
        # Extended vertical lines with independent top & bottom spacing
        top_gap = 25
        bottom_gap = 40
        
        start_y = bottom_gap
        horizontal_y = page_height - top_gap
        end_y = horizontal_y
        
        # Left long vertical line
        canvas.line(page_left_margin, start_y, page_left_margin, end_y)
        
        # Right long vertical line
        canvas.line(footer_x_start + footer_width, start_y, footer_x_start + footer_width, end_y)

        # Connecting horizontal line at the top of the two verticals
        canvas.line(page_left_margin, horizontal_y, footer_x_start + footer_width, horizontal_y)

        # Signature & other column vertical lines (only inside footer height)
        sig_box_width = 80 
        sig_col_x = footer_x_start + sig_box_width
        col1_x = footer_x_start + 160 
        col2_x = footer_x_start + 260
        
        canvas.line(sig_col_x, footer_y, sig_col_x, footer_y + footer_height)
        canvas.line(col1_x, footer_y, col1_x, footer_y + footer_height)
        canvas.line(col2_x, footer_y, col2_x, footer_y + footer_height)
        
        # Row dividers (inside footer)
        row_height = footer_height / 3
        top_div_y = footer_y + 2 * row_height
        mid_div_y = footer_y + row_height
        
        canvas.line(footer_x_start, top_div_y, footer_x_start + footer_width, top_div_y)
        canvas.line(footer_x_start, mid_div_y, col2_x, mid_div_y)
        
        top_cell_top = footer_y + footer_height
        mid_cell_top = top_cell_top - row_height
        bot_cell_top = mid_cell_top - row_height
        
        # Font setup
        try:
            canvas.setFont(FONT_BOLD, 9)
        except:
            canvas.setFont("Helvetica-Bold", 9)
        
        # ========== TOP ROW (right cell: page & version; middle cell: station) ==========
        right_edge_x = footer_x_start + footer_width - 10  # 10pt from right border

        # Page number (top row, right cell)
        page_num_text = f"PAGE NO-{page_num}"
        canvas.drawRightString(right_edge_x, top_cell_top - 20, page_num_text)   # moved down from -17 to -20
        
        # Version (top row, right cell) – only if present
        if version:
            version_text = f"VER {version}"
            canvas.drawRightString(right_edge_x, top_cell_top - 28, version_text) # moved down from -27 to -28
        
        # Station (top row, middle cell) – raised slightly to avoid overlap
        middle_center_x = col1_x + (col2_x - col1_x) / 2
        canvas.drawCentredString(middle_center_x, top_cell_top - 12, station)    # was -18, now -12
        
        # ========== MIDDLE ROW (track) ==========
        canvas.drawCentredString(middle_center_x, mid_cell_top - 18, track)      # unchanged
        
        # ========== BOTTOM ROW (drawing number & division text) ==========
        # Drawing number (bottom row, right cell) – raised to avoid overlap with division text
        canvas.drawRightString(right_edge_x, bot_cell_top - 12, drawing_no)      # was -18, now -12
        
        # "SIG PLAN NO-" label (middle row, right cell) – unchanged
        canvas.drawString(col2_x + 5, mid_cell_top - 18, "SIG PLAN NO-")
        
        # ========== DESIGNATIONS (signature column) ==========
        des_y_offset = 18
        canvas.drawString(sig_col_x + 5, top_cell_top - des_y_offset, des1)
        canvas.drawString(sig_col_x + 5, mid_cell_top - des_y_offset, des2)
        canvas.drawString(sig_col_x + 5, bot_cell_top - des_y_offset, des3)
        
        # ========== DIVISION TEXT (bottom row, middle column) ==========
        try:
            canvas.setFont(FONT_BOLD, 8)
        except:
            canvas.setFont("Helvetica-Bold", 8)
        
        division_y = footer_y + 12
        division_center_x = footer_x_start + (footer_width / 2) + 40
        
        canvas.drawCentredString(division_center_x, division_y + 10, division_line1)
        canvas.drawCentredString(division_center_x, division_y, division_line2)
        
        # ========== COMPLETION LABEL (above footer) ==========
        try:
            canvas.setFont(FONT_BOLD, 9)
        except:
            canvas.setFont("Helvetica-Bold", 9)
        completion_x = footer_x_start + footer_width - 65
        completion_y = footer_y + footer_height + 5
        canvas.drawString(completion_x, completion_y, "COMPLETION")
        
        canvas.restoreState()
        logger.debug(f"Footer added successfully to page {page_num}")
        
    except Exception as e:
        logger.error(f"Error drawing footer: {e}")
        logger.error(traceback.format_exc())
        canvas.restoreState()
        
# ================= LOAD FOOTER DATA FROM EXCEL =================
@log_exceptions
def load_footer_data_from_excel(excel_path):
    """Load footer data from Summary sheet in Excel, now including version"""
    try:
        logger.info(f"Loading footer data from: {excel_path}")
        
        summary_df = pd.read_excel(excel_path, sheet_name='Summary')
        
        footer_data = {}
        
        # Updated column mapping to match your template
        column_mapping = {
            'designation1': ['desg1', 'designation1', 'designation 1', 'sr dste'],
            'designation2': ['desg2', 'designation2', 'designation 2', 'adste'],
            'designation3': ['desg3', 'designation3', 'designation 3', 'sse'],
            'station_name': ['station_name', 'Station', 'station name', 'station'],
            'junction_name': ['name', 'Project', 'junction_name', 'junction name', 'junction', 'track'],
            'station_code': ['station_code', 'station code', 'drawing_name', 'drg_no', 'drawing no', 'big plan'],
            'zone': ['zone', 'zone name', 'zone_no'],
            'division': ['division', 'division name', 'div'],
            'version': ['version', 'version_no', 'revision', 'ver']   # New version column mapping
        }
        
        for key, possible_names in column_mapping.items():
            for col in summary_df.columns:
                col_lower = str(col).lower().strip()
                if any(name in col_lower for name in possible_names):
                    if pd.notna(summary_df.iloc[0][col]):
                        footer_data[key] = str(summary_df.iloc[0][col]).strip()
                        logger.debug(f"Found {key}: '{footer_data[key]}' in column '{col}'")
                    break
        
        logger.info(f"Loaded footer data: {footer_data}")
        return footer_data
        
    except Exception as e:
        logger.error(f"Error loading footer data: {e}")
        logger.error(traceback.format_exc())
        return {}

# ================= IMPROVED DATA LOADING FUNCTIONS =================
@log_exceptions
def load_data_from_excel(excel_path):
    """
    Load all data from Excel file with enhanced logging and validation.
    Returns: (station_name, project_name, ctr_df, rows, no_of_rows, no_of_terminal_per_row)
    """
    if not os.path.exists(excel_path):
        logger.error(f"File not found: {excel_path}")
        return None, None, None, [], None, None

    logger.info(f"Loading data from Excel file: {excel_path}")
    excel_file = pd.ExcelFile(excel_path)
    sheet_names = excel_file.sheet_names
    logger.info(f"Available sheets: {sheet_names}")

    # Defaults
    station_name = "VASADVA STATION"
    project_name = "CTR-1"
    ctr_df = None
    rows = []
    no_of_rows = 0
    no_of_terminal_per_row = 0

    # ---- Summary sheet ----
    if 'Summary' in sheet_names:
        try:
            summary_df = pd.read_excel(excel_path, sheet_name='Summary')
            logger.debug(f"Summary columns: {list(summary_df.columns)}")
            if not summary_df.empty:
                row0 = summary_df.iloc[0].to_dict()
                logger.debug(f"Summary first row: {row0}")

                # station_name
                for col in ['station_name', 'Station']:
                    if col in summary_df.columns and pd.notna(row0.get(col)):
                        station_name = str(row0[col]).strip()
                        logger.info(f"Station name from '{col}': {station_name}")
                        break

                # project_name (track name)
                for col in ['name', 'Project']:
                    if col in summary_df.columns and pd.notna(row0.get(col)):
                        val = str(row0[col]).strip()
                        if not val.startswith('C.TRACK'):
                            val = f"C.TRACK -{val}" if val.isdigit() else val
                        project_name = val
                        logger.info(f"Project name from '{col}': {project_name}")
                        break

                # rows and terminals per row (optional)
                if 'no_of_rows' in summary_df.columns:
                    try:
                        no_of_rows = int(float(str(row0['no_of_rows'])))
                        logger.info(f"no_of_rows from Summary: {no_of_rows}")
                    except:
                        pass
                if 'no_of_terminal_per_row' in summary_df.columns:
                    try:
                        no_of_terminal_per_row = int(float(str(row0['no_of_terminal_per_row'])))
                        logger.info(f"no_of_terminal_per_row from Summary: {no_of_terminal_per_row}")
                    except:
                        pass
        except Exception as e:
            logger.error(f"Error reading Summary sheet: {e}")
    else:
        logger.warning("No 'Summary' sheet found – using default station/project names.")

    # ---- Diagram sheet ----
    if 'Diagram' in sheet_names:
        try:
            ctr_df = pd.read_excel(excel_path, sheet_name='Diagram')
            logger.info(f"Loaded Diagram sheet with {len(ctr_df)} rows")
        except Exception as e:
            logger.error(f"Error reading Diagram sheet: {e}")
    else:
        logger.warning("No 'Diagram' sheet found – CTR diagram will be skipped.")

    # ---- RowDetail sheet (critical) ----
    if 'RowDetail' not in sheet_names:
        logger.error("RowDetail sheet not found in Excel file! Cannot generate terminal diagrams.")
        return station_name, project_name, ctr_df, [], no_of_rows, no_of_terminal_per_row

    try:
        # First, compute max terminals per row if not already known
        if no_of_terminal_per_row == 0:
            # Quick scan to calculate max terminals per main row (before splitting)
            df_raw = pd.read_excel(excel_path, sheet_name='RowDetail', dtype=str).fillna("")
            marker_col = next((col for col in df_raw.columns if 'row' in col.lower() and 'marker' in col.lower()), None)
            term_col = next((col for col in df_raw.columns if 'terminal' in col.lower() and 'no' in col.lower()), None)

            if marker_col and term_col:
                marker_groups = {}
                for _, row in df_raw.iterrows():
                    marker = str(row[marker_col]).strip()[0] if str(row[marker_col]).strip() else 'A'
                    term_val = 1
                    if str(row[term_col]).strip().isdigit():
                        term_val = int(float(str(row[term_col]).strip()))
                    marker_groups[marker] = marker_groups.get(marker, 0) + term_val
                if marker_groups:
                    no_of_terminal_per_row = max(marker_groups.values())
                    logger.info(f"Calculated no_of_terminal_per_row from raw data: {no_of_terminal_per_row}")
                else:
                    no_of_terminal_per_row = 60  # fallback
            else:
                logger.warning("Could not locate row marker or terminal columns for auto-calculation; using 60.")
                no_of_terminal_per_row = 60

        # Now load and split rows with the determined limit
        rows = load_rows_from_rowdetail_sheet(excel_path, max_terminals_per_row=no_of_terminal_per_row)
        logger.info(f"Loaded RowDetail sheet and produced {len(rows)} terminal diagram rows (split by {no_of_terminal_per_row})")

        # Calculate no_of_rows if not set
        if no_of_rows == 0 and rows:
            unique_markers = set()
            for r in rows:
                if not r.get('is_overflow', False) or r.get('overflow_index', 0) == 0:
                    unique_markers.add(r.get('row_marker', 'A'))
            no_of_rows = len(unique_markers)
            logger.info(f"Calculated no_of_rows from processed rows: {no_of_rows}")

    except Exception as e:
        logger.error(f"Error processing RowDetail sheet: {e}")
        logger.error(traceback.format_exc())

    return station_name, project_name, ctr_df, rows, no_of_rows, no_of_terminal_per_row
    
@log_exceptions
def split_row_by_terminal_limit(row_data, max_terminals_per_row=60):
    """
    Split a row if it has more than max_terminals_per_row terminals.
    Returns a list of row_data dictionaries.
    """
    total_terminals = row_data["total_terminals"]
    
    if total_terminals <= max_terminals_per_row:
        # No splitting needed, main row starts from terminal 1
        row_data["start_terminal"] = 1
        row_data["is_overflow"] = False
        row_data["overflow_index"] = 0
        return [row_data]
    
    logger.info(f"Splitting row with marker {row_data['row_marker']} and {total_terminals} terminals")
    
    # We need to split the groups to create multiple rows
    groups = row_data["groups"]
    descriptions = row_data["descriptions"]
    cable_names = row_data["cable_names"]
    cable_core_numbers = row_data.get("cable_core_numbers", [])
    desc_block_sizes = row_data.get("desc_block_sizes", [])
    desc_colors = row_data.get("desc_colors", [])
    
    # Calculate how many rows we need
    rows_needed = math.ceil(total_terminals / max_terminals_per_row)
    logger.info(f"Need {rows_needed} rows for {total_terminals} terminals (max {max_terminals_per_row} per row)")
    
    # Initialize lists for each row
    split_rows = []
    
    current_terminal_count = 0
    current_groups = []
    current_descriptions = []
    current_cable_names = []
    current_cable_core_numbers = []
    current_desc_block_sizes = []
    current_desc_colors = []
    
    # We'll process groups and assign them to rows
    terminal_counter = 1  # Start from terminal 1 for the entire original row
    
    for i, group_size in enumerate(groups):
        # Check if adding this group would exceed the limit in current row
        if current_terminal_count + group_size > max_terminals_per_row and current_terminal_count > 0:
            # Finish current row and start a new one
            # Calculate start terminal for this row
            start_terminal_for_row = terminal_counter - current_terminal_count
            
            split_rows.append({
                "total_terminals": current_terminal_count,
                "row_marker": row_data["row_marker"],
                "groups": current_groups.copy(),
                "descriptions": current_descriptions.copy(),
                "cable_names": current_cable_names.copy(),
                "cable_core_numbers": current_cable_core_numbers.copy() if current_cable_core_numbers else None,
                "desc_block_sizes": current_desc_block_sizes.copy(),
                "desc_colors": current_desc_colors.copy() if current_desc_colors else None,
                "is_overflow": len(split_rows) > 0,  # First split row is main, rest are overflow
                "overflow_index": len(split_rows),  # 0 for main, 1 for first overflow, etc.
                "start_terminal": start_terminal_for_row
            })
            
            # Reset for new row
            current_terminal_count = 0
            current_groups = []
            current_descriptions = []
            current_cable_names = []
            current_cable_core_numbers = []
            current_desc_block_sizes = []
            current_desc_colors = []
        
        # Add group to current row
        current_groups.append(group_size)
        current_descriptions.append(descriptions[i] if i < len(descriptions) else [])
        current_cable_names.append(cable_names[i] if i < len(cable_names) else "")
        
        if cable_core_numbers and i < len(cable_core_numbers):
            current_cable_core_numbers.append(cable_core_numbers[i])
        
        if desc_block_sizes and i < len(desc_block_sizes):
            current_desc_block_sizes.append(desc_block_sizes[i])
        
        if desc_colors and i < len(desc_colors):
            current_desc_colors.append(desc_colors[i])
        
        current_terminal_count += group_size
        terminal_counter += group_size
    
    # Don't forget the last row
    if current_groups:
        start_terminal_for_row = terminal_counter - current_terminal_count
        
        split_rows.append({
            "total_terminals": current_terminal_count,
            "row_marker": row_data["row_marker"],
            "groups": current_groups,
            "descriptions": current_descriptions,
            "cable_names": current_cable_names,
            "cable_core_numbers": current_cable_core_numbers if current_cable_core_numbers else None,
            "desc_block_sizes": current_desc_block_sizes,
            "desc_colors": current_desc_colors if current_desc_colors else None,
            "is_overflow": len(split_rows) > 0,  # If we have previous rows, this is overflow
            "overflow_index": len(split_rows),
            "start_terminal": start_terminal_for_row
        })
    
    # Sort by start_terminal in DESCENDING order (higher terminal numbers first)
    # This puts overflow rows (61-120) BEFORE main rows (1-60)
    split_rows.sort(key=lambda x: x["start_terminal"], reverse=True)
    
    logger.info(f"Split row into {len(split_rows)} parts:")
    for i, row in enumerate(split_rows):
        logger.info(f"  Row {i}: terminals {row['start_terminal']} to {row['start_terminal'] + row['total_terminals'] - 1} "
                   f"({row['total_terminals']} terminals), is_overflow={row['is_overflow']}, "
                   f"overflow_index={row['overflow_index']}, marker={row['row_marker']}")
    
    return split_rows


def clean_cable_name(name):
    if not name:
        return ""
    
    parts = name.split()
    seen = []
    for p in parts:
        if p not in seen:
            seen.append(p)
    return " ".join(seen)
@log_exceptions
def load_rows_from_rowdetail_sheet(excel_path, max_terminals_per_row=None):
    """
    Load rows from RowDetail sheet, respecting a terminal limit per physical row.
    Includes detailed column mapping and validation.
    """
    logger.info(f"Loading RowDetail sheet from: {excel_path} with max_terminals_per_row={max_terminals_per_row}")
    df = pd.read_excel(excel_path, sheet_name='RowDetail', dtype=str).fillna("")

    logger.info(f"RowDetail columns: {list(df.columns)}")
    logger.info(f"RowDetail shape: {df.shape}")
    if not df.empty:
        logger.debug(f"First 3 rows:\n{df.head(3)}")

    # Map expected columns with flexible matching
    col_map = {
        'TerminalNo': None,
        'RowMarker': None,
        'Description': None,
        'CableName': None,
        'CableCoreStart': None,
        'CableCoreEnd': None,
        'BlockSize': None,
        'Color': None
    }

    for col in df.columns:
        col_lower = col.lower().strip()
        if 'terminal' in col_lower and 'no' in col_lower:
            col_map['TerminalNo'] = col
        elif 'row' in col_lower and 'marker' in col_lower:
            col_map['RowMarker'] = col
        elif 'description' in col_lower:
            col_map['Description'] = col
        elif 'cable' in col_lower and 'name' in col_lower:
            col_map['CableName'] = col
        elif 'cable' in col_lower and 'core' in col_lower and ('start' in col_lower or 'begin' in col_lower):
            col_map['CableCoreStart'] = col
        elif 'cable' in col_lower and 'core' in col_lower and ('end' in col_lower or 'finish' in col_lower):
            col_map['CableCoreEnd'] = col
        elif 'block' in col_lower and 'size' in col_lower:
            col_map['BlockSize'] = col
        elif 'color' in col_lower:
            col_map['Color'] = col

    # Log what we found
    logger.info(f"Column mapping: {col_map}")
    missing = [k for k, v in col_map.items() if v is None]
    if missing:
        logger.warning(f"Missing columns: {missing}. Some data may be empty.")

    grouped_data = {}

    for idx, row in df.iterrows():
        try:
            # Row marker
            if col_map['RowMarker']:
                marker_val = str(row[col_map['RowMarker']]).strip()
                marker = marker_val[0].upper() if marker_val else 'A'
            else:
                marker = 'A'

            if marker not in grouped_data:
                grouped_data[marker] = {
                    'groups': [],
                    'descriptions': [],
                    'cable_names': [],
                    'cable_core_starts': [],
                    'cable_core_ends': [],
                    'block_sizes': [],
                    'colors': []
                }

            # Terminal count (number of terminals in this group)
            if col_map['TerminalNo']:
                term_str = str(row[col_map['TerminalNo']]).strip()
                terminal_no = int(float(term_str)) if term_str and term_str.replace('.','',1).isdigit() else 1
            else:
                terminal_no = 1  # fallback

            # Description
            description = str(row[col_map['Description']]).strip() if col_map['Description'] else ""

            # Cable name
            cable_name = str(row[col_map['CableName']]).strip() if col_map['CableName'] else ""
            cable_name = clean_cable_name(cable_name)
            
            #cable_name = str(row[col_map['CableName']]).strip() if col_map['CableName'] else ""

            # Core start/end
            core_start = None
            if col_map['CableCoreStart']:
                cs = str(row[col_map['CableCoreStart']]).strip()
                if cs and cs.isdigit():
                    core_start = int(cs)
            core_end = None
            if col_map['CableCoreEnd']:
                ce = str(row[col_map['CableCoreEnd']]).strip()
                if ce and ce.isdigit():
                    core_end = int(ce)

            # Block size
            block_size = 1
            if col_map['BlockSize']:
                bs = str(row[col_map['BlockSize']]).strip()
                if bs and bs.isdigit():
                    block_size = int(bs)
            block_size = min(block_size, terminal_no)  # cannot be larger than group size

            # Color override
            color_override = None
            if col_map['Color']:
                color_str = str(row[col_map['Color']]).strip()
                if color_str and color_str.lower() not in ['', 'nan', 'none', 'null']:
                    color_override = safe_parse_color(color_str)
                    if color_override:
                        logger.debug(f"Row {idx+2}: color '{color_str}' -> {color_override}")

            grouped_data[marker]['groups'].append(terminal_no)
            grouped_data[marker]['descriptions'].append([description])
            grouped_data[marker]['cable_names'].append(cable_name)
            grouped_data[marker]['cable_core_starts'].append(core_start)
            grouped_data[marker]['cable_core_ends'].append(core_end)
            grouped_data[marker]['block_sizes'].append(block_size)
            grouped_data[marker]['colors'].append(color_override)

        except Exception as e:
            logger.error(f"Error processing row {idx+2}: {e}")
            continue

    # Build rows from grouped data
    all_rows = []
    for marker, data in grouped_data.items():
        if not data['groups']:
            continue

        total_terminals = sum(data['groups'])
        logger.debug(f"Marker {marker}: total terminals {total_terminals} in {len(data['groups'])} groups")

        # Construct core numbers per group
        cable_core_numbers = []
        for i, group_size in enumerate(data['groups']):
            start = data['cable_core_starts'][i]
            end = data['cable_core_ends'][i]
            if start is not None and end is not None:
                cores = list(range(start, end + 1))
                # If the range is shorter than group size, pad with empty strings
                if len(cores) < group_size:
                    cores.extend([''] * (group_size - len(cores)))
                else:
                    cores = cores[:group_size]
            else:
                cores = [''] * group_size
            cable_core_numbers.append(cores)

        # Build descriptions and colors per block
        descriptions = []
        desc_colors = []
        for i, group_size in enumerate(data['groups']):
            block_size = data['block_sizes'][i]
            n_blocks = math.ceil(group_size / block_size)
            group_desc = data['descriptions'][i]  # list with one element usually
            group_color = data['colors'][i]

            if len(group_desc) == 1 and n_blocks > 1:
                block_descs = [group_desc[0]] * n_blocks
                block_colors = [group_color] * n_blocks if group_color else [None] * n_blocks
            else:
                block_descs = group_desc[:n_blocks]  # should match
                block_colors = [group_color] * len(block_descs) if group_color else [None] * len(block_descs)

            descriptions.append(block_descs)
            desc_colors.append(block_colors)

        row_data = {
            "total_terminals": total_terminals,
            "row_marker": marker,
            "groups": data['groups'],
            "desc_block_sizes": data['block_sizes'],
            "descriptions": descriptions,
            "cable_names": data['cable_names'],
            "cable_core_numbers": cable_core_numbers,
            "desc_colors": desc_colors,
        }

        limit = max_terminals_per_row if max_terminals_per_row and max_terminals_per_row > 0 else 60
        split_rows = split_row_by_terminal_limit(row_data, max_terminals_per_row=limit)
        all_rows.extend(split_rows)

    logger.info(f"Finished processing RowDetail: total rows after splitting = {len(all_rows)}")
    return all_rows
        
# ================= PDF CREATION =================
@log_exceptions
def create_terminal_diagram_pdf(filename, rows, station_name, project_name, ctr_image_path=None, footer_data=None, diagram_start_y_offset=80):
    """
    Create the main PDF document with adjustable vertical positioning
    
    Args:
        filename: Output filename
        rows: Terminal diagram rows (already split by 60 terminals)
        station_name: Station name
        project_name: Project/track name
        ctr_image_path: Path to CTR diagram image
        footer_data: Footer data dictionary
        diagram_start_y_offset: Vertical offset in points (positive = move down, negative = move up)
    """
    pdf_path = f"{filename}.pdf"
    try:
        logger.info(f"Creating PDF: {pdf_path}")
        logger.info(f"Parameters: station={station_name}, project={project_name}, rows={len(rows)}, y_offset={diagram_start_y_offset}")
        
        # Create document with adjusted bottom margin to accommodate footer
        doc = SimpleDocTemplate(pdf_path, pagesize=landscape(A2),
                               leftMargin=10*mm, rightMargin=10*mm,
                               topMargin=10*mm, bottomMargin=60*mm)
        
        def add_footer(canvas, doc):
            add_pdf_footer(canvas, doc, footer_data or {}, 0)
        
        elements = []
        styles = getSampleStyleSheet()

        station_style = ParagraphStyle('Station', parent=styles['Heading1'], 
                                       fontSize=30, alignment=2,
                                       spaceAfter=0, fontName=FONT_BOLD,
                                       textColor=colors.black,
                                       leftIndent=40*mm,
                                       spaceBefore=0)
        
        project_style = ParagraphStyle('Project', parent=styles['Heading2'], 
                                       fontSize=20, alignment=1,
                                       spaceAfter=0, fontName=FONT_BOLD,
                                       textColor=colors.black,
                                       leftIndent=65*mm,
                                       spaceBefore=12*mm)
        
        station_text = Paragraph(station_name, station_style)
        project_text = Paragraph(project_name, project_style)
        
        # Hardcoded logo path
        logo_path = r"C:\Railway\git\Circuitbuilding\app\static\images\railway_logo.jpg"
        
        # ===== FIRST PAGE HEADER ONLY =====
        logger.info("Adding header for first page only")
        if ctr_image_path and os.path.exists(ctr_image_path):
            logger.info(f"Including CTR image: {ctr_image_path}")
            ctr_img = Image(ctr_image_path, width=380, height=228)
            
            # Check if logo exists
            logo_img = None
            if os.path.exists(logo_path):
                try:
                    # Bigger logo size - 140x140
                    logo_img = Image(logo_path, width=140, height=140)
                    logger.info(f"Loaded logo: {logo_path}")
                except Exception as e:
                    logger.warning(f"Could not load logo {logo_path}: {e}")
                    logo_img = None
            
            if logo_img:
                # Three columns: Logo, Text, CTR Image
                logo_col_width = 180
                text_col_width = doc.width * 0.45
                ctr_col_width = doc.width * 0.37
                
                text_table_data = [
                    [station_text],
                    [project_text]
                ]
                
                text_table = Table(text_table_data, colWidths=[text_col_width])
                text_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('BOTTOMPADDING', (0, 0), (0, 0), 0),
                    ('TOPPADDING', (0, 1), (0, 1), 8*mm),
                    ('LEFTPADDING', (0, 0), (-1, -1), 15*mm),
                    ('ALIGN', (0, 1), (0, 1), 'CENTER'),
                ]))
                
                # Create a separate table for logo with perfect border
                logo_table = Table([[logo_img]], 
                                  colWidths=[140],
                                  rowHeights=[140])
                
                logo_table.setStyle(TableStyle([
                    ('BOX', (0, 0), (0, 0), 1, colors.black),
                    ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                    ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                    ('LEFTPADDING', (0, 0), (0, 0), 0),
                    ('RIGHTPADDING', (0, 0), (0, 0), 0),
                    ('TOPPADDING', (0, 0), (0, 0), 0),
                    ('BOTTOMPADDING', (0, 0), (0, 0), 0),
                ]))
                
                header_table_data = [
                    [logo_table, text_table, ctr_img]
                ]
                
                header_table = Table(header_table_data, colWidths=[logo_col_width, text_col_width, ctr_col_width])
                header_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                    ('VALIGN', (0, 0), (0, 0), 'TOP'),
                    ('LEFTPADDING', (0, 0), (0, 0), 30),
                    ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                    ('VALIGN', (1, 0), (1, 0), 'TOP'),
                    ('ALIGN', (2, 0), (2, 0), 'RIGHT'),
                    ('VALIGN', (2, 0), (2, 0), 'MIDDLE'),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                    ('TOPPADDING', (0, 0), (-1, -1), 0),
                ]))
                
                elements.append(header_table)
                logger.info("Created header with logo and CTR image")
            else:
                # Original two columns (no logo)
                col1_width = doc.width * 0.60
                col2_width = doc.width * 0.40
                
                text_table_data = [
                    [station_text],
                    [project_text]
                ]
                
                text_table = Table(text_table_data, colWidths=[col1_width])
                text_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('BOTTOMPADDING', (0, 0), (0, 0), 0),
                    ('TOPPADDING', (0, 1), (0, 1), 8*mm),
                    ('LEFTPADDING', (0, 0), (-1, -1), 30*mm),
                    ('ALIGN', (0, 1), (0, 1), 'CENTER'),
                ]))
                
                header_table_data = [
                    [text_table, ctr_img]
                ]
                
                header_table = Table(header_table_data, colWidths=[col1_width, col2_width])
                header_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (0, 0), 'RIGHT'),
                    ('VALIGN', (0, 0), (0, 0), 'TOP'),
                    ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                    ('VALIGN', (1, 0), (1, 0), 'MIDDLE'),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                    ('TOPPADDING', (0, 0), (-1, -1), 0),
                ]))
                
                elements.append(header_table)
                logger.info("Created header with CTR image (no logo)")
        else:
            # No CTR image - just station and project text
            if os.path.exists(logo_path):
                try:
                    logo_img = Image(logo_path, width=140, height=140)
                    logo_col_width = 180
                    text_col_width = doc.width - logo_col_width
                    
                    text_table_data = [
                        [station_text],
                        [project_text]
                    ]
                    
                    text_table = Table(text_table_data, colWidths=[text_col_width])
                    text_table.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ('BOTTOMPADDING', (0, 0), (0, 0), 0),
                        ('TOPPADDING', (0, 1), (0, 1), 8*mm),
                        ('LEFTPADDING', (0, 0), (-1, -1), 15*mm),
                        ('ALIGN', (0, 1), (0, 1), 'CENTER'),
                    ]))
                    
                    logo_table = Table([[logo_img]], 
                                      colWidths=[140],
                                      rowHeights=[140])
                    
                    logo_table.setStyle(TableStyle([
                        ('BOX', (0, 0), (0, 0), 1, colors.black),
                        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
                        ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                        ('LEFTPADDING', (0, 0), (0, 0), 0),
                        ('RIGHTPADDING', (0, 0), (0, 0), 0),
                        ('TOPPADDING', (0, 0), (0, 0), 0),
                        ('BOTTOMPADDING', (0, 0), (0, 0), 0),
                    ]))
                    
                    header_table_data = [
                        [logo_table, text_table]
                    ]
                    
                    header_table = Table(header_table_data, colWidths=[logo_col_width, text_col_width])
                    header_table.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                        ('VALIGN', (0, 0), (0, 0), 'TOP'),
                        ('LEFTPADDING', (0, 0), (0, 0), 30),
                        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
                        ('VALIGN', (1, 0), (1, 0), 'TOP'),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                        ('TOPPADDING', (0, 0), (-1, -1), 0),
                    ]))
                    
                    elements.append(header_table)
                    logger.info("Created header with logo (no CTR image)")
                except Exception as e:
                    logger.warning(f"Could not add logo: {e}")
                    elements.append(station_text)
                    elements.append(Spacer(1, 12*mm))
                    elements.append(project_text)
                    logger.info("Created simple header (no logo, no CTR image)")
            else:
                elements.append(station_text)
                elements.append(Spacer(1, 12*mm))
                elements.append(project_text)
                logger.info("Created simple header")
        
        # Add adjustable vertical spacing after header (ONLY on first page)
        #vertical_spacing = diagram_start_y_offset
        vertical_spacing=0
        elements.append(Spacer(1, vertical_spacing))
        logger.info(f"Added vertical spacing of {vertical_spacing} points after header on first page")
        
        diagram_width = doc.width - 20*mm
        
        # ===== CORRECT DESCENDING ORDER =====
        # We need to sort rows in DESCENDING order:
        # 1. First by marker (F, E, D, C, B, A) - descending
        # 2. Then by start_terminal (61-120 first, then 1-60) - descending
        
        logger.info(f"Total rows before sorting: {len(rows)}")
        
        # Create a custom sort function
        def sort_key(row):
            marker = row.get('row_marker', 'A')
            start_terminal = row.get('start_terminal', 1)
            # For marker: Z > A, so F > E > D > C > B > A
            # For start_terminal: higher first (61-120 before 1-60)
            return (-ord(marker), -start_terminal)
        
        # Sort rows
        ordered_rows = sorted(rows, key=sort_key)
        
        # Log the order for debugging
        logger.info(f"Sorted rows in descending order:")
        for i, row in enumerate(ordered_rows):
            logger.info(f"  Row {i}: marker={row.get('row_marker', 'A')}, "
                       f"start={row.get('start_terminal', 1)}, "
                       f"terminals={row.get('total_terminals', 0)}")
        
        # ===== CALCULATE ROWS PER PAGE =====
        # We want 7 rows per page
        rows_per_page = 7
        
        # For 7 rows per page, use super compact mode
        diagram_height_super_compact = 23*mm
        
        # Calculate available space
        page_height = landscape(A2)[1]  # Height in points
        header_height = 200  # Estimated header height
        footer_height = 60*mm
        
        # Available height on FIRST page (with header)
        #first_page_available = page_height - header_height - footer_height - vertical_spacing
        first_page_available = page_height - header_height - footer_height  
        
        # Available height on SUBSEQUENT pages (no header)
        subsequent_page_available = page_height - footer_height - 10*mm  # 10mm top margin
        
        # Calculate how many rows can fit on each page
        #rows_per_page_first = min(rows_per_page, int(first_page_available // diagram_height_super_compact))
        #rows_per_page_subsequent = min(rows_per_page, int(subsequent_page_available // diagram_height_super_compact))
        rows_per_page_first=len(rows)
        rows_per_page_subsequent=0
        
        # Ensure at least 1 row per page
        rows_per_page_first = max(1, rows_per_page_first)
        rows_per_page_subsequent = max(1, rows_per_page_subsequent)
        
        logger.info(f"First page can fit: {rows_per_page_first} rows (available={first_page_available:.1f})")
        logger.info(f"Subsequent pages can fit: {rows_per_page_subsequent} rows (available={subsequent_page_available:.1f})")
        
        # ===== ADD DIAGRAMS TO PAGES =====
        current_index = 0
        page_number = 0
        '''
        while current_index < len(ordered_rows):
            page_number += 1
            
            if page_number == 1:
                # First page with header
                rows_this_page = rows_per_page_first
                use_super_compact = True
            else:
                # Subsequent pages
                rows_this_page = rows_per_page_subsequent
                use_super_compact = True
                # Add page break for pages after first
                elements.append(PageBreak())
                logger.info(f"Added PageBreak for page {page_number}")
            
            # Get rows for this page
            chunk = ordered_rows[current_index:current_index + rows_this_page]
            logger.info(f"Page {page_number}: {len(chunk)} rows (rows {current_index+1} to {current_index+len(chunk)})")
            
            # Add each diagram
            for j, row in enumerate(chunk):
                is_overflow = row.get('is_overflow', False)
                overflow_index = row.get('overflow_index', 0)
                start_terminal = row.get('start_terminal', 1)
                
                logger.debug(f"  Adding diagram {j+1}: marker={row.get('row_marker', 'A')}, "
                           f"start={start_terminal}, overflow={is_overflow}")
                
                diagram = TerminalDiagram(
                    row['groups'],
                    row['descriptions'],
                    row['cable_names'],
                    row['total_terminals'],
                    desc_block_sizes=row.get('desc_block_sizes'),
                    cable_core_numbers=row.get('cable_core_numbers'),
                    desc_colors=row.get('desc_colors'),
                    width=diagram_width,
                    row_marker=row.get('row_marker', 'A'),
                    table_shift_right=10*mm,
                    compact_mode=True,
                    super_compact_mode=use_super_compact,
                    is_overflow=is_overflow,
                    overflow_index=overflow_index,
                    start_terminal=start_terminal
                )
                elements.append(diagram)
            
            current_index += rows_this_page
        '''
        use_super_compact = True 
        for j, row in enumerate(ordered_rows):
                is_overflow = row.get('is_overflow', False)
                overflow_index = row.get('overflow_index', 0)
                start_terminal = row.get('start_terminal', 1)
                
                logger.debug(f"  Adding diagram {j+1}: marker={row.get('row_marker', 'A')}, "
                           f"start={start_terminal}, overflow={is_overflow}")
                
                diagram = TerminalDiagram(
                    row['groups'],
                    row['descriptions'],
                    row['cable_names'],
                    row['total_terminals'],
                    desc_block_sizes=row.get('desc_block_sizes'),
                    cable_core_numbers=row.get('cable_core_numbers'),
                    desc_colors=row.get('desc_colors'),
                    width=diagram_width,
                    row_marker=row.get('row_marker', 'A'),
                    table_shift_right=10*mm,
                    compact_mode=True,
                    super_compact_mode=use_super_compact,
                    is_overflow=is_overflow,
                    overflow_index=overflow_index,
                    start_terminal=start_terminal
                )
                elements.append(diagram)
                print(diagram) 
        
       
        logger.info(f"Building PDF with {len(elements)} elements...")
        doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
        
        # Check if PDF was created
        if os.path.exists(pdf_path):
            file_size = os.path.getsize(pdf_path)
            logger.info(f"PDF generated successfully: {pdf_path} ({file_size} bytes)")
            logger.info(f"Location: {os.path.abspath(pdf_path)}")
            
            # Calculate total pages
            if len(ordered_rows) <= rows_per_page_first:
                total_pages = 1
            else:
                remaining = len(ordered_rows) - rows_per_page_first
                total_pages = 1 + (remaining + rows_per_page_subsequent - 1) // rows_per_page_subsequent
            
            logger.info(f"Total pages: {total_pages}")
            logger.info(f"First page rows: {rows_per_page_first}")
            logger.info(f"Subsequent page rows: {rows_per_page_subsequent}")
            logger.info(f"Total rows: {len(ordered_rows)}")
            
            return pdf_path
        else:
            logger.error(f"PDF file was not created: {pdf_path}")
            return None
        
    except Exception as e:
        logger.error(f"Error generating PDF: {e}")
        logger.error(traceback.format_exc())
        return None


# ================= IMPROVED PDF GENERATION FUNCTION =================
@log_exceptions
def generate_ctr_pdf_from_excel(excel_path, output_dir=None, logo_path=None, diagram_start_y_offset=80, version=None):
    """
    Main function to generate CTR PDF from Excel file.

    Args:
        excel_path: Path to the uploaded Excel file
        output_dir: Directory to save the PDF (optional)
        logo_path: Path to logo image (optional)
        diagram_start_y_offset: Vertical offset in points (positive = move down)
        version: Optional version string. If provided, it overrides any version in the Excel file.
                 This ensures that the version number in the PDF footer is correct even after
                 multiple generations.
    Returns:
        Dictionary with keys: pdf_path, no_of_rows, no_of_terminal_per_row
    """
    logger.info(f"Starting PDF generation from: {excel_path}")
    logger.info(f"Output dir: {output_dir}, Logo: {logo_path}, Y-offset: {diagram_start_y_offset}, Version: {version}")

    # Load all data
    station_name, project_name, ctr_df, rows, no_of_rows, no_of_terminal_per_row = load_data_from_excel(excel_path)

    if not rows:
        logger.error("No terminal diagram data loaded from 'RowDetail' sheet!")
        return {'pdf_path': None, 'no_of_rows': 0, 'no_of_terminal_per_row': 0}

    # Load footer data (including version from Excel if present)
    footer_data = load_footer_data_from_excel(excel_path)

    # Fill missing footer fields from station/project if needed
    if 'station_name' not in footer_data and station_name:
        footer_data['station_name'] = station_name
    if 'junction_name' not in footer_data and project_name:
        footer_data['junction_name'] = project_name

    # IMPORTANT: Version precedence – passed version overrides any version from Excel
    if version is not None:
        footer_data['version'] = str(version)
        logger.info(f"Using passed version: {version}")
    elif 'version' not in footer_data:
        footer_data['version'] = '0'   # default if missing
        logger.info("No version found; using default '0'")
    else:
        logger.info(f"Using version from Excel: {footer_data['version']}")

    # Generate CTR diagram image if possible
    ctr_image_path = None
    if ctr_df is not None and not ctr_df.empty:
        logger.info("Generating CTR diagram from Diagram sheet...")
        ctr_image_path = generate_ctr_diagram_from_df(ctr_df, r"C:\Railway\git\static\ctr_diagram.png")
    else:
        logger.warning("No Diagram sheet data; CTR diagram will be omitted.")

    # Determine output filename: include version and timestamp to avoid collisions
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    version_str = footer_data.get('version', '0')
    output_filename = f"{base_name}_v{version_str}_{timestamp}"
    app_root = os.path.dirname(__file__)
    if output_dir:
        output_dir = os.path.join(app_root, output_dir)
        output_dir = r"C:\Railway\git\Circuitbuilding\uploads_ctr"
        os.makedirs(output_dir, exist_ok=True)
        output_filename = os.path.join(output_dir, output_filename)

    # Generate PDF
    pdf_path = create_terminal_diagram_pdf(
        output_filename,
        rows,
        station_name,
        project_name,
        ctr_image_path,
        footer_data,
        diagram_start_y_offset=diagram_start_y_offset
    )

    # Cleanup temporary image
    if ctr_image_path and os.path.exists(ctr_image_path):
        try:
            os.remove(ctr_image_path)
            logger.info(f"Cleaned up temporary CTR image: {ctr_image_path}")
        except Exception as e:
            logger.warning(f"Could not remove temporary CTR image: {e}")

    result = {
        'pdf_path': pdf_path,
        'no_of_rows': no_of_rows,
        'no_of_terminal_per_row': no_of_terminal_per_row
    }

    if pdf_path:
        logger.info(f"PDF generation completed successfully: {pdf_path}")
    else:
        logger.error("PDF generation failed!")

    return result
        
# ================= MAIN EXECUTION BLOCK =================
if __name__ == "__main__":
    """Main execution block for standalone testing"""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python ctr_pdf_generator.py <excel_file_path> [output_directory] [y_offset] [version]")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    y_offset = int(sys.argv[3]) if len(sys.argv) > 3 else 80
    version = sys.argv[4] if len(sys.argv) > 4 else None
    
    logger.info(f"Starting standalone execution with Excel file: {excel_path}")
    
    try:
        result = generate_ctr_pdf_from_excel(excel_path, output_dir, diagram_start_y_offset=y_offset, version=version)
        if result['pdf_path']:
            logger.info(f"Successfully generated PDF: {result['pdf_path']}")
            logger.info(f"Calculated Summary: Rows={result['no_of_rows']}, Terminals per row={result['no_of_terminal_per_row']}")
            print(f"PDF generated: {result['pdf_path']}")
            print(f"Rows: {result['no_of_rows']}")
            print(f"Terminals per row: {result['no_of_terminal_per_row']}")
        else:
            logger.error("Failed to generate PDF")
            print("Failed to generate PDF")
    except Exception as e:
        logger.error(f"Standalone execution failed: {e}")
        print(f"Error: {e}")