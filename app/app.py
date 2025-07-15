from flask import Flask, render_template, request, send_file, redirect, url_for, jsonify, session
import pandas as pd
from fpdf import FPDF
import os
import zipfile
from werkzeug.utils import secure_filename
from PIL import Image  # Added for image dimension detection
from datetime import datetime  # Added for date formatting
from report_generator import ReportGenerator  # Import our new reporting module
import uuid
import threading
import time
import webbrowser

app = Flask(__name__)
app.secret_key = 'your-secret-key-here-change-this-in-production'  # Add secret key for sessions

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


UPLOAD_FOLDER = os.path.join(BASE_DIR, "..", "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "..", "outputs")
LOGO_DB_FOLDER = os.path.join(BASE_DIR, "..", "logo_database")
LOGO_IMAGES_FOLDER = os.path.join(BASE_DIR, "..", "logo_images")
ZIP_NAME = f"art_instructions_pdfs_{datetime.now().strftime('%m_%d_%Y_%H_%M_%S')}.zip"
LOGO_DB_FILE = "ArtDBSample.xlsx"  # Logo database file
STATIC_IMAGE_PATH = os.path.join(BASE_DIR, "static", "jauniforms.png")

# Create all necessary folders
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(LOGO_DB_FOLDER, exist_ok=True)
os.makedirs(LOGO_IMAGES_FOLDER, exist_ok=True)

# Global variable to store logo database
logo_database = None

# Progress tracking
progress_status = {}  # Dictionary to store progress for each session

def update_progress(session_id, status, progress=0, message="", current_step="", total_steps=0):
    """Update progress status for a session"""
    if session_id in progress_status:
        progress_status[session_id].update({
            'status': status,  # 'processing', 'completed', 'error'
            'progress': progress,  # 0-100
            'message': message,
            'current_step': current_step,
            'total_steps': total_steps,
            'timestamp': time.time()
        })

def load_logo_database():
    """Load the logo database into memory"""
    global logo_database
    logo_db_path = os.path.join(LOGO_DB_FOLDER, LOGO_DB_FILE)
    
    if os.path.exists(logo_db_path):
        try:
            logo_database = pd.read_excel(logo_db_path)
            logo_database.columns = [col.strip() for col in logo_database.columns]
            print(f"Logo database loaded successfully with {len(logo_database)} entries")
        except Exception as e:
            print(f"Error loading logo database: {e}")
            logo_database = None
    else:
        print(f"Logo database file not found at: {logo_db_path}")
        logo_database = None

def get_logo_info(logo_sku):
    """Get logo information from the database based on SKU (preserving leading zeros)"""
    if logo_database is None or pd.isna(logo_sku) or logo_sku == "" or str(logo_sku) == "0000":
        return None
    
    try:
        # Preserve the original format including leading zeros
        logo_sku_str = str(logo_sku).strip()
        
        # Search for the logo SKU in the database (try both original and numeric formats)
        logo_row = logo_database[logo_database['Logo SKU'].astype(str).str.strip() == logo_sku_str]
        
        # If not found with original format, try numeric conversion for backward compatibility
        if logo_row.empty and logo_sku_str.isdigit():
            numeric_sku = str(int(logo_sku_str))
            logo_row = logo_database[logo_database['Logo SKU'].astype(str).str.strip() == numeric_sku]
        
        if not logo_row.empty:
            row = logo_row.iloc[0]
            return {
                'logo_sku': logo_sku_str,  # Use original format
                'client': safe_get(row['CLIENT']),
                'logo_position': safe_get(row['Logo Position']),
                'operation_type': safe_get(row['Operation Type']),
                'stitch_count': safe_get(row['Stitch Count']),
                'file_name': safe_get(row['File Name']),
                'notes': safe_get(row['Notes']),
                'size': safe_get(row['Size']),
                'logo_colors': get_logo_colors(row)
            }
    except Exception as e:
        print(f"Error looking up logo SKU {logo_sku}: {e}")
    
    return None

def get_logo_colors(row):
    """Extract logo colors from the database row"""
    colors = []
    for i in range(1, 16):  # Logo Color 1 through Logo Color 15
        color_col = f'Logo Color {i}'
        if color_col in row and pd.notna(row[color_col]) and str(row[color_col]).strip():
            colors.append(str(row[color_col]).strip())
    return colors

def find_logo_images_by_sku(logo_sku):
    """Find all logo image files based on SKU number with suffix letters (preserving leading zeros)"""
    if not logo_sku or pd.isna(logo_sku) or logo_sku == "":
        return []
    
    # Preserve original format including leading zeros
    sku_str = str(logo_sku).strip()
    
    # Skip if it's the default "0000" or equivalent
    if sku_str == "0000" or sku_str == "0":
        return []
    
    # Common image extensions
    extensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff']
    found_images = []
    
    # Check for images with suffix letters (a, b, c, d, ...)
    for suffix in 'abcdefghijklmnopqrstuvwxyz':
        for ext in extensions:
            image_filename = f"{sku_str}{suffix}{ext}"
            image_path = os.path.join(LOGO_IMAGES_FOLDER, image_filename)
            if os.path.exists(image_path):
                found_images.append({
                    'path': image_path,
                    'filename': image_filename,
                    'suffix': suffix
                })
    
    # Sort by suffix to maintain order (a, b, c, ...)
    found_images.sort(key=lambda x: x['suffix'])
    
    return found_images

def format_date_consistently(date_value):
    """Convert various date formats to MM/dd/yy format"""
    if pd.isna(date_value) or date_value == "":
        return ""
    
    try:
        # Handle different input types
        if isinstance(date_value, (int, float)):
            # Excel serial date number
            if date_value > 40000:  # Reasonable range for Excel dates (2009+)
                # Convert Excel serial date to Python datetime
                excel_epoch = datetime(1899, 12, 30)
                date_obj = excel_epoch + pd.Timedelta(days=date_value)
            else:
                return str(int(date_value))
        elif isinstance(date_value, str):
            # String date - try to parse various formats
            date_str = str(date_value).strip()
            
            # If it's already in MM/dd/yy format, return as-is
            if len(date_str) == 8 and date_str.count('/') == 2:
                parts = date_str.split('/')
                if len(parts[2]) == 2:  # Already in MM/dd/yy format
                    return date_str
            
            # Try to parse the string as a date
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y', '%Y/%m/%d', '%d/%m/%Y']:
                try:
                    date_obj = datetime.strptime(date_str, fmt)
                    break
                except ValueError:
                    continue
            else:
                return date_str
        else:
            date_obj = pd.to_datetime(date_value)
        
        # Format as MM/dd/yy
        return date_obj.strftime('%m/%d/%y')
        
    except Exception as e:
        print(f"Error formatting date '{date_value}': {e}")
        return str(date_value)

def read_file_with_format_detection(file_path):
    """Read Excel or CSV file"""
    file_extension = os.path.splitext(file_path)[1].lower()
    
    try:
        if file_extension == '.csv':
            df = pd.read_csv(file_path, dtype={'LOGO': str})
            print(f"Successfully read CSV file: {file_path}")
        else:
            df = pd.read_excel(file_path, dtype={'LOGO': str})
            print(f"Successfully read Excel file: {file_path}")
        return df
    except:
        # Fallback
        try:
            if file_extension == '.csv':
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            if 'LOGO' in df.columns:
                df['LOGO'] = df['LOGO'].astype(str)
            return df
        except Exception as e:
            raise e

def safe_get(value):
    return "" if pd.isna(value) else str(value)

def truncate_text(text, pdf, max_width):
    ellipsis = '...'
    original = text
    while pdf.get_string_width(text) > max_width:
        if len(text) <= len(ellipsis):
            return ellipsis
        text = text[:-1]
    return text + ellipsis if text != original else text

def render_items_section(pdf, vendor_styles, total_width):
    styles = vendor_styles.split(", ")
    label_width = 30
    value_width = total_width - label_width
    max_width = value_width - 5

    pdf.set_font("Arial", "", 8.5)
    line = ""
    for style in styles:
        appended = style + ", "
        if pdf.get_string_width(line + appended) < max_width:
            line += appended
        else:
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(label_width, 5, "ITEMS:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            pdf.cell(value_width, 5, line.strip(", "), border=1)
            pdf.ln()
            line = appended

    if line:
        pdf.set_font("Arial", "B", 8.5)
        pdf.cell(label_width, 5, "ITEMS:", border=1, align="C")
        pdf.set_font("Arial", "", 8.5)
        pdf.cell(value_width, 5, line.strip(", "), border=1)
        pdf.ln()

def calculate_text_height(pdf, text, available_width, line_height=5):
    """
    Calculate the required height for text that may need to wrap
    """
    if not text or text.strip() == "":
        return line_height
    
    text_str = str(text).strip()
    available_width = available_width - 4  # Account for padding
    
    # If text fits in one line
    if pdf.get_string_width(text_str) <= available_width:
        return line_height
    
    # Calculate number of lines needed
    words = text_str.split()
    lines = []
    current_line = ""
    
    for word in words:
        test_line = current_line + (" " if current_line else "") + word
        if pdf.get_string_width(test_line) <= available_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
                current_line = word
            else:
                # Single word is too long - break it by characters
                if pdf.get_string_width(word) > available_width:
                    # Calculate how many lines this long word will need
                    chars_so_far = ""
                    for char in word:
                        test_chars = chars_so_far + char
                        if pdf.get_string_width(test_chars) > available_width:
                            if chars_so_far:
                                lines.append(chars_so_far)
                                chars_so_far = char
                            else:
                                lines.append(char)  # Single character that's too wide
                                chars_so_far = ""
                        else:
                            chars_so_far = test_chars
                    if chars_so_far:
                        current_line = chars_so_far
                else:
                    current_line = word
    
    if current_line:
        lines.append(current_line)
    
    # Return height needed with minimum of original line_height
    return max(len(lines) * line_height, line_height)

def add_multiline_text_to_cell(pdf, text, x, y, width, height, border=1, align="L", fill=False):
    """
    Add text to a cell with proper line wrapping and boundary control
    """
    # Draw background fill first if needed
    if fill:
        pdf.set_fill_color(255, 255, 0)  # Yellow color
        pdf.rect(x, y, width, height, style='F')
        pdf.set_fill_color(255, 255, 255)  # Reset to white background
    
    # Draw the cell border
    if border:
        pdf.rect(x, y, width, height)
    
    if not text or text.strip() == "":
        return
    
    text_str = str(text).strip()
    # More conservative padding to ensure text stays within borders
    padding = 1
    available_width = width - (2 * padding)
    line_height = 4  # Slightly smaller line height for better fit
    
    # Calculate maximum lines that can fit
    max_lines = max(1, int((height - 2) / line_height))
    
    # If text fits in one line
    if pdf.get_string_width(text_str) <= available_width:
        if align == "C":
            text_x = x + (width - pdf.get_string_width(text_str)) / 2 
        elif align == "L":
            text_x = x + padding
        else:  # Right align
            text_x = x + width - pdf.get_string_width(text_str) - padding
        
        text_y = y + (height - line_height) / 2
        pdf.set_xy(text_x, text_y)
        pdf.cell(pdf.get_string_width(text_str), line_height, text_str, 0, 0, 'L')
        return
    
    
    # For long text, break it properly
    words = text_str.split()
    lines = []
    current_line = ""
    
    for word in words:
        test_line = current_line + (" " if current_line else "") + word
        if pdf.get_string_width(test_line) <= available_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
                current_line = word
            else:
                # Single word is too long - break it by characters
                remaining_word = word
                while remaining_word and len(lines) < max_lines:
                    char_line = ""
                    for char in remaining_word:
                        test_char_line = char_line + char
                        if pdf.get_string_width(test_char_line) <= available_width:
                            char_line = test_char_line
                        else:
                            break
                    
                    if char_line:
                        lines.append(char_line)
                        remaining_word = remaining_word[len(char_line):]
                    else:
                        # Even single character doesn't fit - just add it
                        lines.append(remaining_word[0])
                        remaining_word = remaining_word[1:]
                
                current_line = remaining_word if len(lines) < max_lines else ""
    
    if current_line and len(lines) < max_lines:
        lines.append(current_line)
    
    # Limit to max lines that fit in cell height
    lines = lines[:max_lines]
    
    # Calculate starting Y position to center the text block vertically
    total_text_height = len(lines) * line_height
    start_y = y + max(1, (height - total_text_height) / 2)
    
    # Add lines to PDF with strict boundary control
    for i, line in enumerate(lines):
        line_y = start_y + (i * line_height)
        
        # Make sure we don't draw outside the cell
        if line_y + line_height > y + height:
            break
        
        if align == "C":
            line_x = x + (width - pdf.get_string_width(line)) / 2
        elif align == "L":
            line_x = x + padding
        else:  # Right align
            line_x = x + width - pdf.get_string_width(line) - padding
        
        # Ensure text doesn't go outside cell boundaries
        line_x = max(x + padding, min(line_x, x + width - pdf.get_string_width(line) - padding))
        
        pdf.set_xy(line_x, line_y)
        pdf.cell(pdf.get_string_width(line), line_height, line, 0, 0, 'L')

def add_logo_color_table(pdf, logo_colors=None):
    """Enhanced logo color table with actual colors from database and truncation (using consistent width)"""
    #pdf.ln(5)
    # Use the same usable_width as other tables for consistent right margin
    usable_width = 190 - (2 * 0.8)  # Same calculation as in main function
    logo_color_width = usable_width * 0.20
    number_width = usable_width * 0.05
    value_width = usable_width * 0.35

    # First row: LOGO COLOR header
    pdf.set_font("Arial", "B", 8.5)
    pdf.cell(logo_color_width, 5, "LOGO COLOR:", border=1, align="C")
    pdf.set_font("Arial", "", 8.5)
    
    # Add first color if available (truncated to 95% of cell width)
    color1 = logo_colors[0] if logo_colors and len(logo_colors) > 0 else ""
    color1_display = truncate_text(color1, pdf, value_width * 0.95)
    pdf.cell(number_width, 5, "1", border=1, align="C")
    pdf.cell(value_width, 5, color1_display, border=1)
    
    # Add ninth color if available (truncated to 95% of cell width)
    color9 = logo_colors[8] if logo_colors and len(logo_colors) > 8 else ""
    color9_display = truncate_text(color9, pdf, value_width * 0.95)
    pdf.cell(number_width, 5, "9", border=1, align="C")
    pdf.cell(value_width, 5, color9_display, border=1)
    pdf.ln()

    # Second row: PRODUCTION DAY directly under LOGO COLOR
    pdf.set_font("Arial", "B", 8.5)
    pdf.cell(logo_color_width, 5, "PRODUCTION DAY:", border=1, align="C")
    pdf.set_font("Arial", "", 8.5)
    
    # Add second and tenth colors if available (truncated to 95% of cell width)
    color2 = logo_colors[1] if logo_colors and len(logo_colors) > 1 else ""
    color2_display = truncate_text(color2, pdf, value_width * 0.95)
    color10 = logo_colors[9] if logo_colors and len(logo_colors) > 9 else ""
    color10_display = truncate_text(color10, pdf, value_width * 0.95)
    pdf.cell(number_width, 5, "2", border=1, align="C")
    pdf.cell(value_width, 5, color2_display, border=1)
    pdf.cell(number_width, 5, "10", border=1, align="C")
    pdf.cell(value_width, 5, color10_display, border=1)
    pdf.ln()

    # Calculate the height of the merged cell (6 rows * 5 units = 30 units)
    merged_cell_height = 6 * 5  # 6 rows of height 5 each
    
    # Store current position to draw the merged cell
    current_x = pdf.get_x()
    current_y = pdf.get_y()
    
    # Draw the large merged cell for logo color column
    pdf.cell(logo_color_width, merged_cell_height, "", border=1)
    
    # Move to the position right after the merged cell to continue with other columns
    pdf.set_xy(current_x + logo_color_width, current_y)
    
    # Draw rows 3-8 (numbers 3-8 and 11-16)
    for i in range(3, 8):
        color_left = logo_colors[i-1] if logo_colors and len(logo_colors) > i-1 else ""
        color_right = logo_colors[i+7] if logo_colors and len(logo_colors) > i+7 else ""
        
        # Truncate colors to 95% of cell width
        color_left_display = truncate_text(color_left, pdf, value_width * 0.95)
        color_right_display = truncate_text(color_right, pdf, value_width * 0.95)
        
        pdf.cell(number_width, 5, str(i), border=1, align="C")
        pdf.cell(value_width, 5, color_left_display, border=1)
        pdf.cell(number_width, 5, str(i + 8), border=1, align="C")
        pdf.cell(value_width, 5, color_right_display, border=1)
        # Move to next line, but stay at the same x position (after the merged cell)
        pdf.set_xy(current_x + logo_color_width, pdf.get_y() + 5)

    # Last row with only left half filled (number 8), right half blank
    color8 = logo_colors[7] if logo_colors and len(logo_colors) > 7 else ""
    color8_display = truncate_text(color8, pdf, value_width * 0.95)
    pdf.cell(number_width, 5, "8", border=1, align="C")
    pdf.cell(value_width, 5, color8_display, border=1)
    pdf.cell(number_width + value_width, 5, "", border=1)
    pdf.ln()

def apply_max_size_constraint(width_mm, height_mm, max_width=91.9, max_height=58.1):
    """Apply maximum size constraint while preserving aspect ratio"""
    if width_mm <= max_width and height_mm <= max_height:
        # Image is within limits, return as-is
        return width_mm, height_mm
    
    # Calculate scaling factors for both dimensions
    width_scale = max_width / width_mm
    height_scale = max_height / height_mm
    
    # Use the smaller scaling factor to ensure both dimensions fit
    scale_factor = min(width_scale, height_scale)
    
    # Apply scaling
    new_width = width_mm * scale_factor
    new_height = height_mm * scale_factor
    
    return new_width, new_height

def get_image_dimensions_mm(image_path, dpi=300):
    """Get image dimensions in millimeters with max size constraint"""
    try:
        with Image.open(image_path) as img:
            width_px, height_px = img.size
            # Convert pixels to millimeters (assuming 300 DPI)
            width_mm = (width_px / dpi) * 25.4
            height_mm = (height_px / dpi) * 25.4
            
            # Apply maximum size constraint
            constrained_width, constrained_height = apply_max_size_constraint(width_mm, height_mm)
            
            return constrained_width, constrained_height
    except Exception as e:
        print(f"Error getting dimensions for {image_path}: {e}")
        return 25, 20  # Default fallback size

def calculate_optimal_layout(images, available_width, available_height, margin=5, max_width=91.9, max_height=58.1):
    """Calculate optimal layout for images with max size constraint"""
    if not images:
        return []
    
    # Get actual dimensions for all images (already constrained by max size)
    image_info = []
    for img in images:
        width, height = get_image_dimensions_mm(img['path'])
        image_info.append({
            'path': img['path'],
            'filename': img['filename'],
            'suffix': img['suffix'],
            'original_width': width,  # Already constrained
            'original_height': height  # Already constrained
        })
    
    # Try to fit images at their constrained sizes first
    layout = []
    current_row = []
    current_row_width = 0
    current_row_height = 0
    total_height_used = 0
    
    for img_info in image_info:
        img_width = img_info['original_width']
        img_height = img_info['original_height']
        
        # Check if image fits in current row
        needed_width = current_row_width + (margin if current_row else 0) + img_width
        
        if needed_width <= available_width and total_height_used + img_height <= available_height:
            # Fits in current row at constrained size
            current_row.append({
                **img_info,
                'display_width': img_width,
                'display_height': img_height,
                'use_actual_size': True
            })
            current_row_width = needed_width
            current_row_height = max(current_row_height, img_height)
        else:
            # Start new row if current row has images
            if current_row:
                layout.append(current_row)
                total_height_used += current_row_height + margin
                current_row = []
                current_row_width = 0
                current_row_height = 0
            
            # Check if single image fits at constrained size in new row
            if img_width <= available_width and total_height_used + img_height <= available_height:
                current_row.append({
                    **img_info,
                    'display_width': img_width,
                    'display_height': img_height,
                    'use_actual_size': True
                })
                current_row_width = img_width
                current_row_height = img_height
            else:
                # Need to resize further - will handle this in optimization phase
                current_row.append({
                    **img_info,
                    'display_width': img_width,
                    'display_height': img_height,
                    'use_actual_size': False
                })
                current_row_width = img_width
                current_row_height = img_height
    
    # Add last row
    if current_row:
        layout.append(current_row)
        total_height_used += current_row_height
    
    # Check if we need to optimize sizes (if any images don't fit at constrained size)
    needs_optimization = any(
        not img['use_actual_size'] 
        for row in layout 
        for img in row
    ) or total_height_used > available_height
    
    if needs_optimization:
        # Optimize layout to fit all images (with max size constraint)
        layout = optimize_image_layout(image_info, available_width, available_height, margin, max_width, max_height)
    
    return layout

def optimize_image_layout(image_info, available_width, available_height, margin=5, max_width=91.9, max_height=58.1):
    """Optimize image layout to fit all images in available space"""
    num_images = len(image_info)
    if num_images == 0:
        return []
    
    # Calculate how many images per row and rows needed
    best_layout = None
    best_waste = float('inf')
    
    # Try different arrangements (1 to num_images per row)
    for images_per_row in range(1, num_images + 1):
        rows_needed = (num_images + images_per_row - 1) // images_per_row
        
        # Calculate available space per image
        width_per_image = (available_width - (images_per_row - 1) * margin) / images_per_row
        height_per_row = (available_height - (rows_needed - 1) * margin) / rows_needed
        
        # Check if this arrangement is feasible
        max_aspect_ratio = max(img['original_width'] / img['original_height'] for img in image_info)
        min_aspect_ratio = min(img['original_width'] / img['original_height'] for img in image_info)
        
        # Calculate what size images would be with this constraint
        if width_per_image / height_per_row >= max_aspect_ratio:
            # Height is the limiting factor
            actual_height = height_per_row
            actual_width = min(width_per_image, height_per_row * max_aspect_ratio)
        else:
            # Width is the limiting factor
            actual_width = width_per_image
            actual_height = min(height_per_row, width_per_image / min_aspect_ratio)
        
        # Calculate wasted space
        used_width = images_per_row * actual_width + (images_per_row - 1) * margin
        used_height = rows_needed * actual_height + (rows_needed - 1) * margin
        wasted_space = (available_width * available_height) - (used_width * used_height)
        
        if wasted_space < best_waste:
            best_waste = wasted_space
            best_layout = {
                'images_per_row': images_per_row,
                'rows_needed': rows_needed,
                'width_per_image': actual_width,
                'height_per_image': actual_height
            }
    
    # Create the optimized layout
    if not best_layout:
        return []
    
    layout = []
    current_row = []
    
    for i, img_info in enumerate(image_info):
        # Maintain aspect ratio while fitting in allocated space
        aspect_ratio = img_info['original_width'] / img_info['original_height']
        
        if best_layout['width_per_image'] / best_layout['height_per_image'] > aspect_ratio:
            # Height is limiting
            display_height = best_layout['height_per_image']
            display_width = display_height * aspect_ratio
        else:
            # Width is limiting
            display_width = best_layout['width_per_image']
            display_height = display_width / aspect_ratio
        
        current_row.append({
            **img_info,
            'display_width': display_width,
            'display_height': display_height,
            'use_actual_size': False
        })
        
        # Start new row if needed
        if len(current_row) == best_layout['images_per_row']:
            layout.append(current_row)
            current_row = []
    
    # Add last row if it has images
    if current_row:
        layout.append(current_row)
    
    return layout

def add_logo_images_to_pdf(pdf, logo_sku, logo_info=None):
    """Add logo images to PDF with intelligent sizing and layout"""
    if not logo_sku or pd.isna(logo_sku) or logo_sku == "":
        return
    
    # Find all images for this SKU
    logo_images = find_logo_images_by_sku(logo_sku)
    
    # Note: Due to pre-validation, this should not happen, but keeping minimal check
    if not logo_images:
        print(f"Warning: No logo images found for SKU: {logo_sku} during PDF generation")
        return
    
    print(f"Found {len(logo_images)} logo image(s) for SKU {logo_sku}: {[img['filename'] for img in logo_images]}")
    
    # Calculate available space on current page
    current_y = pdf.get_y()
    page_height = pdf.h - pdf.b_margin  # Page height minus bottom margin
    available_height = page_height - current_y - 10  # Leave 10mm buffer
    available_width = pdf.w - pdf.l_margin - pdf.r_margin  # Available width
    margin = 5  # Margin between images
    
    # Get optimal layout
    layout = calculate_optimal_layout(logo_images, available_width, available_height, margin)
    
    if not layout:
        print(f"Could not fit images for SKU {logo_sku}")
        
        # Add "could not fit images" message to PDF
        current_y = pdf.get_y() + 5
        pdf.set_xy(pdf.l_margin, current_y)
        pdf.set_font("Arial", "I", 10)
        pdf.set_text_color(128, 128, 128)  # Gray text
        pdf.cell(0, 8, f"Logo images too large to fit for SKU: {logo_sku}", align="C")
        pdf.ln(12)
        pdf.set_text_color(0, 0, 0)  # Reset to black text
        return
    
    try:
        start_y = current_y + 5  # Small buffer from previous content
        
        for row_index, row in enumerate(layout):
            if not row:
                continue
                
            # Calculate starting X position to center the row
            row_width = sum(img['display_width'] for img in row) + margin * (len(row) - 1)
            start_x = pdf.l_margin + (available_width - row_width) / 2
            
            # Calculate Y position for this row
            if row_index == 0:
                row_y = start_y
            else:
                # Position based on previous row's height
                prev_row_height = max(img['display_height'] for img in layout[row_index - 1])
                row_y = start_y + sum(
                    max(img['display_height'] for img in layout[i]) + margin 
                    for i in range(row_index)
                )
            
            # Place images in this row
            current_x = start_x
            row_height = max(img['display_height'] for img in row)
            
            for img_info in row:
                # Center image vertically in the row
                img_y = row_y + (row_height - img_info['display_height']) / 2
                
                # Add image to PDF
                pdf.image(
                    img_info['path'], 
                    x=current_x, 
                    y=img_y, 
                    w=img_info['display_width'], 
                    h=img_info['display_height']
                )
                
                # Add suffix label below image
                label_y = img_y + img_info['display_height'] + 1
                pdf.set_xy(current_x, label_y)
                pdf.set_font("Arial", "", 8)
                pdf.cell(img_info['display_width'], 3, f"({img_info['suffix']})", align="C")
                
                # Move to next image position
                current_x += img_info['display_width'] + margin
                
                # Debug info
                size_info = "actual size" if img_info.get('use_actual_size', False) else "optimized size"
                print(f"  Added {img_info['filename']} at {size_info}: {img_info['display_width']:.1f}x{img_info['display_height']:.1f}mm")
        
        # Update PDF cursor position
        total_layout_height = sum(
            max(img['display_height'] for img in row) + margin 
            for row in layout
        ) + 10  # Extra buffer
        pdf.set_xy(pdf.l_margin, start_y + total_layout_height)
        
    except Exception as e:
        print(f"Error adding logo images for SKU {logo_sku}: {e}")
        
        # Add error message to PDF
        current_y = pdf.get_y() + 5
        pdf.set_xy(pdf.l_margin, current_y)
        pdf.set_font("Arial", "I", 10)
        pdf.set_text_color(255, 0, 0)  # Red text
        pdf.cell(0, 8, f"Error loading logo images for SKU: {logo_sku}", align="C")
        pdf.ln(12)
        pdf.set_text_color(0, 0, 0)  # Reset to black text
        
def filter_by_sales_order(df, sales_order_filter):
    """Filter dataframe by sales order number if provided (exact match only)"""
    if not sales_order_filter or sales_order_filter.strip() == "":
        return df
    
    sales_order_filter = sales_order_filter.strip()
    print(f"Filtering by Sales Order (exact match): '{sales_order_filter}'")
    
    # Check if Document Number column exists
    if 'Document Number' not in df.columns:
        print("Warning: 'Document Number' column not found in data")
        return pd.DataFrame()  # Return empty dataframe
    
    # Filter by exact match only
    original_count = len(df)
    filtered_df = df[df['Document Number'].astype(str).str.strip() == sales_order_filter]
    
    print(f"Sales Order filter result: {len(filtered_df)} rows found out of {original_count} total rows")
    
    if filtered_df.empty:
        print(f"No exact match found for Sales Order: '{sales_order_filter}'")
    else:
        found_orders = filtered_df['Document Number'].unique()
        print(f"Found Sales Orders: {list(found_orders)}")
    
    return filtered_df

def validate_row_for_processing(row, report_data, approval_filter="approved_only"):  # MODIFIED LINE
    """
    Validate a single row for processing and return validation result
    Returns: (is_valid, error_message)
    """
    doc_num = safe_get(row.get("Document Number", ""))
    logo_sku = safe_get(row.get("LOGO", ""))
    
    # Record entry for reporting
    row_data = {
        'Document Number': doc_num,
        'LOGO': logo_sku,
        'VENDOR STYLE': safe_get(row.get("VENDOR STYLE", "")),
        'COLOR': safe_get(row.get("COLOR", "")),
        'SIZE': safe_get(row.get("SIZE", "")),
        'SUBCATEGORY': safe_get(row.get("SUBCATEGORY", "")),
        'Quantity': safe_get(row.get("Quantity", "")),
        'Customer/Vendor Name': safe_get(row.get("Customer/Vendor Name", "")),
        'Due Date': safe_get(row.get("Due Date", "")),
        'DueDateStatus': safe_get(row.get("DueDateStatus", "")),
        'OPERATIONAL CODE': safe_get(row.get("OPERATIONAL CODE", "")),
        'List of Operation Codes': safe_get(row.get("List of Operation Codes", "")),
        'LOGO POSITION': safe_get(row.get("LOGO POSITION", "")),
        'STITCH COUNT': safe_get(row.get("STITCH COUNT", "")),
        'NOTES': safe_get(row.get("NOTES", "")),
        'FILE NAME': safe_get(row.get("FILE NAME", ""))
    }
    
    # Validation 1: Check DueDateStatus based on approval filter - MODIFIED SECTION
    due_date_status = safe_get(row.get("DueDateStatus", "")).strip().upper()
    
    if approval_filter == "approved_only":
        # Only process approved orders (skip "Not Approved")
        if due_date_status == "NOT APPROVED":
            row_data['Execution Status'] = 'FAILED'
            row_data['Error Message'] = 'Status: Not Approved (filtered out)'
            report_data.append(row_data)
            return False, "Status: Not Approved (filtered out)"
    elif approval_filter == "not_approved_only":
        # Only process not approved orders (skip anything that's not "Not Approved")
        if due_date_status != "NOT APPROVED":
            row_data['Execution Status'] = 'FAILED'
            row_data['Error Message'] = 'Status: Approved (filtered out)'
            report_data.append(row_data)
            return False, "Status: Approved (filtered out)"
    # For approval_filter == "both", we don't filter based on approval status
    
    # Rest of the validation logic remains the same...
    
    # Validation 2: Check Logo SKU validity
    logo_sku_str = str(logo_sku).strip()
    if pd.isna(logo_sku) or logo_sku_str in ["", "0", "0000", "nan", "NaN"]:
        row_data['Execution Status'] = 'FAILED'
        row_data['Error Message'] = f'Invalid Logo SKU: "{logo_sku_str}"'
        report_data.append(row_data)
        return False, f'Invalid Logo SKU: "{logo_sku_str}"'
    
    # Validation 3: Check if logo info exists in database
    logo_info = get_logo_info(logo_sku_str)
    if logo_info is None:
        row_data['Execution Status'] = 'FAILED'
        row_data['Error Message'] = f'Logo info not found in database for SKU: {logo_sku_str}'
        report_data.append(row_data)
        return False, f'Logo info not found in database for SKU: {logo_sku_str}'
    
    # Validation 4: Check if logo images exist
    logo_images = find_logo_images_by_sku(logo_sku_str)
    if not logo_images:
        row_data['Execution Status'] = 'FAILED'
        row_data['Error Message'] = f'Logo images not found for SKU: {logo_sku_str}'
        report_data.append(row_data)
        return False, f'Logo images not found for SKU: {logo_sku_str}'
    
    # Validation 5: Check Operational Code validity
    operational_code = None
    op_code_raw = row.get("OPERATIONAL CODE")
    
    if pd.notna(op_code_raw) and str(op_code_raw).strip():
        op_code_str = str(op_code_raw).strip()
        
        if op_code_str not in ["00", "0", ""]:
            try:
                if '.' in op_code_str:
                    operational_code = int(float(op_code_str))
                else:
                    operational_code = int(op_code_str)
            except (ValueError, TypeError):
                row_data['Execution Status'] = 'FAILED'
                row_data['Error Message'] = f'Invalid Operational Code format: "{op_code_str}"'
                report_data.append(row_data)
                return False, f'Invalid Operational Code format: "{op_code_str}"'
        else:
            row_data['Execution Status'] = 'FAILED'
            row_data['Error Message'] = f'Invalid Operational Code: "{op_code_str}" (00, 0, or empty)'
            report_data.append(row_data)
            return False, f'Invalid Operational Code: "{op_code_str}" (00, 0, or empty)'
    else:
        row_data['Execution Status'] = 'FAILED'
        row_data['Error Message'] = 'Missing or empty Operational Code'
        report_data.append(row_data)
        return False, "Missing or empty Operational Code"
    
    # Validation 6: Check Operational Code conditions
    if operational_code == 11:
        # Valid - Operational Code is 11
        row_data['Execution Status'] = 'SUCCESS'
        row_data['Error Message'] = ''
        report_data.append(row_data)
        return True, ""
    elif operational_code > 89:
        # Check List of Operation Codes
        list_operation_codes = []
        list_codes_raw = row.get("List of Operation Codes")
        
        if pd.notna(list_codes_raw) and str(list_codes_raw).strip():
            list_codes_str = str(list_codes_raw).strip()
            
            # Parse comma-separated codes
            if ',' in list_codes_str:
                individual_codes = list_codes_str.split(',')
                for individual_code in individual_codes:
                    clean_code = individual_code.strip()
                    if clean_code and clean_code.replace('.', '').isdigit():
                        try:
                            if '.' in clean_code:
                                list_operation_codes.append(int(float(clean_code)))
                            else:
                                list_operation_codes.append(int(clean_code))
                        except (ValueError, TypeError):
                            pass
            else:
                # Single code
                if list_codes_str.replace('.', '').isdigit():
                    try:
                        if '.' in list_codes_str:
                            list_operation_codes.append(int(float(list_codes_str)))
                        else:
                            list_operation_codes.append(int(list_codes_str))
                    except (ValueError, TypeError):
                        pass
        
        # Validate List of Operation Codes
        if not list_operation_codes:
            row_data['Execution Status'] = 'FAILED'
            row_data['Error Message'] = 'No valid List of Operation Codes found for Operational Code > 89'
            report_data.append(row_data)
            return False, "No valid List of Operation Codes found for Operational Code > 89"
        
        # Must contain exactly one 11
        count_of_11 = list_operation_codes.count(11)
        if count_of_11 != 1:
            row_data['Execution Status'] = 'FAILED'
            row_data['Error Message'] = f'List must contain exactly one 11 (found {count_of_11})'
            report_data.append(row_data)
            return False, f'List must contain exactly one 11 (found {count_of_11})'
        
        # No operation code should be less than 60 (except 11)
        codes_less_than_60 = [code for code in list_operation_codes if code < 60 and code != 11]
        if codes_less_than_60:
            row_data['Execution Status'] = 'FAILED'
            row_data['Error Message'] = f'List contains codes < 60 (excluding 11): {codes_less_than_60}'
            report_data.append(row_data)
            return False, f'List contains codes < 60 (excluding 11): {codes_less_than_60}'
        
        # Valid
        row_data['Execution Status'] = 'SUCCESS'
        row_data['Error Message'] = ''
        report_data.append(row_data)
        return True, ""
    else:
        row_data['Execution Status'] = 'FAILED'
        row_data['Error Message'] = f'Operational Code {operational_code} is not 11 and not > 89'
        report_data.append(row_data)
        return False, f'Operational Code {operational_code} is not 11 and not > 89'

def process_file_with_progress(file_path, sales_order_filter, session_id, approval_filter="approved_only"):
    """
    Process the file with progress updates - this replaces your main processing logic
    """
    try:
        # Initialize progress
        progress_status[session_id] = {
            'status': 'processing',
            'progress': 0,
            'message': 'Starting file processing...',
            'current_step': 'Initializing',
            'total_steps': 8,
            'timestamp': time.time()
        }
        
        # Step 1: Load logo database
        update_progress(session_id, 'processing', 5, 'Loading logo database...', 'Database Loading', 8)
        load_logo_database()
        time.sleep(0.5)  # Small delay for user feedback
        
        # Step 2: Read and process file
        update_progress(session_id, 'processing', 15, 'Reading and processing uploaded file...', 'File Processing', 8)
        df = read_file_with_format_detection(file_path)
        df.columns = [col.strip() for col in df.columns]
        time.sleep(0.5)
        
        # Step 3: Apply filters
        update_progress(session_id, 'processing', 25, 'Applying filters and validation...', 'Data Filtering', 8)
        if sales_order_filter:
            df = filter_by_sales_order(df, sales_order_filter)
            if df.empty:
                update_progress(session_id, 'error', 0, f'No exact match found for Sales Order: {sales_order_filter}', 'Error', 8)
                return {'success': False, 'error': f'No exact match found for Sales Order: {sales_order_filter}'}
        
        # Clean LOGO column
        if 'LOGO' in df.columns:
            def clean_logo_value(logo_val):
                if pd.isna(logo_val) or logo_val in ['nan', 'NaN', '']:
                    return ""
                logo_str = str(logo_val).strip()
                if logo_str.endswith('.0'):
                    logo_str = logo_str[:-2]
                if logo_str in ['', 'nan', 'NaN', '0', '0000']:
                    return ""
                if logo_str.isdigit() and len(logo_str) < 4:
                    logo_str = logo_str.zfill(4)
                return logo_str
            
            df['LOGO'] = df['LOGO'].apply(clean_logo_value)
        
        # Step 4: Validate data
        update_progress(session_id, 'processing', 40, 'Validating data and checking requirements...', 'Data Validation', 8)
        report_data = []
        
        # Process each row for validation
        for index, row in df.iterrows():
            is_valid, error_msg = validate_row_for_processing(row, report_data, approval_filter)
            if not is_valid:
                print(f"Row {index + 1}: {error_msg}")
        
        time.sleep(0.5)
        
        # Step 5: Clear output folder and prepare for PDF generation
        update_progress(session_id, 'processing', 50, 'Preparing output folder...', 'Setup', 8)
        for f in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, f))
        
        # Step 6: Generate PDFs
        update_progress(session_id, 'processing', 60, 'Generating PDF documents...', 'PDF Generation', 8)
        
        # Filter valid rows and group by Document Number and Logo SKU
        valid_df = df[df['LOGO'] != ""]
        final_valid_rows = []
        for index, row in valid_df.iterrows():
            temp_report = []
            is_valid, _ = validate_row_for_processing(row, temp_report, approval_filter)
            if is_valid:
                final_valid_rows.append(row)
        
        pdf_count = 0
        if final_valid_rows:
            valid_df = pd.DataFrame(final_valid_rows)
            grouped = valid_df.groupby(["Document Number", "LOGO"])
            total_groups = len(grouped)
            
            for group_index, ((doc_num, logo_sku), group) in enumerate(grouped):
                # Update progress for each PDF
                pdf_progress = 60 + (group_index / total_groups) * 20  # PDF generation takes 20% (60-80%)
                update_progress(session_id, 'processing', pdf_progress, 
                              f'Generating PDF {group_index + 1} of {total_groups} (SO: {doc_num}, Logo: {logo_sku})', 
                              'PDF Generation', 8)
                
                try:
                    # Generate PDF (your existing PDF generation code)
                    pdf = FPDF(orientation="P", unit="mm", format=(190.5, 254.0))
                    pdf.set_margins(0.8, 0.8, 0.8)
                    pdf.add_page()
                    pdf.set_auto_page_break(auto=True, margin=0.8)
                    pdf.set_font("Arial", "", 8.5)

                    due_date = datetime.now().strftime("%m/%d/%Y")

                    full_width = 190
                    usable_width = full_width - (2 * 0.8)
                    left_width = full_width * 0.75
                    right_width = full_width - left_width

                    # Calculate client_name after left_width is defined
                    client_name = truncate_text(safe_get(group["Customer/Vendor Name"].iloc[0]), pdf, (left_width - 20) * 0.95)

                    pdf.set_font("Arial", "B", 10)
                    pdf.cell(left_width, 8, "ART INSTRUCTIONS", border=1, align="C")
                    pdf.cell(right_width, 8, "", border=0)
                    pdf.image(STATIC_IMAGE_PATH, x=pdf.get_x() - right_width + 3, y=pdf.get_y() + 1, w=right_width - 6)
                    pdf.ln()

                    pdf.set_font("Arial", "B", 8.5)
                    pdf.cell(20, 6, "CLIENT:", border=1, align="C")
                    pdf.set_font("Arial", "", 8.5)
                    pdf.cell(left_width - 20, 6, client_name, border=1)
                    pdf.cell(right_width, 6, "", border=0)
                    pdf.ln()

                    so_section_width = left_width * 0.70
                    date_section_width = left_width * 0.30

                    pdf.set_font("Arial", "B", 8.5)
                    pdf.cell(20, 6, "SO#:", border=1, align="C")
                    pdf.set_font("Arial", "", 8.5)
                    so_display = truncate_text(str(doc_num), pdf, (so_section_width - 20) * 0.95)
                    pdf.cell(so_section_width - 20, 6, so_display, border=1)

                    pdf.set_font("Arial", "B", 8.5)
                    pdf.cell(15, 6, "DATE:", border=1, align="C")
                    pdf.set_font("Arial", "", 8.5)
                    pdf.cell(date_section_width - 15, 6, due_date, border=1, align="C")
                    pdf.cell(right_width, 6, "", border=0)
                    pdf.ln(8)

                    vendor_styles = ", ".join(group["VENDOR STYLE"].dropna().astype(str).unique())
                    render_items_section(pdf, vendor_styles, usable_width)

                    pdf.ln(2)
                    COLOR_WIDTH = usable_width * 0.55
                    DESC_WIDTH = usable_width * 0.30
                    QTY_WIDTH = usable_width * 0.15

                    pdf.set_font("Arial", "B", 8.5)
                    pdf.cell(COLOR_WIDTH, 5, "COLOR", 1, align="C")
                    pdf.cell(DESC_WIDTH, 5, "DESCRIPTION", 1, align="C")
                    pdf.cell(QTY_WIDTH, 5, "QTY", 1, align="C")
                    pdf.ln()

                    # Group by COLOR and DESCRIPTION, then sum quantities
                    color_desc_groups = {}
                    total_qty = 0
                    
                    for _, row in group.iterrows():
                        color = safe_get(row.get("COLOR")).strip().upper()
                        desc = safe_get(row.get("SUBCATEGORY")).strip().upper()
                        
                        try:
                            qty = float(row.get("Quantity", 0))
                        except:
                            qty = 0
                        
                        # Create a key from color and description for grouping
                        group_key = f"{color}|{desc}"
                        
                        if group_key in color_desc_groups:
                            color_desc_groups[group_key]['quantity'] += qty
                        else:
                            color_desc_groups[group_key] = {
                                'color': color,
                                'description': desc,
                                'quantity': qty
                            }
                        
                        total_qty += qty

                    # Display grouped results with enhanced formatting
                    pdf.set_font("Arial", "", 8.5)
                    for group_key, group_data in color_desc_groups.items():
                        color_display = truncate_text(group_data['color'], pdf, COLOR_WIDTH * 0.90)
                        desc_display = truncate_text(group_data['description'], pdf, DESC_WIDTH * 0.90)
                        qty_display = str(int(group_data['quantity']))
                        
                        # Calculate if quantity needs multiple lines
                        qty_width = pdf.get_string_width(qty_display)
                        qty_cell_width = QTY_WIDTH * 0.95  # Use 95% of quantity cell width
                        
                        if qty_width <= qty_cell_width:
                            # Single line - normal height
                            cell_height = 5
                            pdf.cell(COLOR_WIDTH, cell_height, color_display, 1, align="C")
                            pdf.cell(DESC_WIDTH, cell_height, desc_display, 1, align="C")
                            pdf.cell(QTY_WIDTH, cell_height, qty_display, 1, align="C")
                            pdf.ln()
                        else:
                            # Multi-line quantity - calculate needed height
                            lines_needed = int(qty_width / qty_cell_width) + 1
                            cell_height = 5 * lines_needed
                            
                            # Store current position
                            current_x = pdf.get_x()
                            current_y = pdf.get_y()
                            
                            # Draw color and description cells with increased height
                            pdf.cell(COLOR_WIDTH, cell_height, color_display, 1, align="C")
                            pdf.cell(DESC_WIDTH, cell_height, desc_display, 1, align="C")
                            
                            # Draw quantity cell border first
                            pdf.cell(QTY_WIDTH, cell_height, "", 1)
                            
                            # Now add multi-line quantity text
                            pdf.set_xy(current_x + COLOR_WIDTH + DESC_WIDTH + 1, current_y + 1)
                            
                            # Split quantity into chunks that fit
                            qty_chars = list(qty_display)
                            chars_per_line = int(len(qty_chars) / lines_needed)
                            
                            for line_num in range(lines_needed):
                                start_idx = line_num * chars_per_line
                                if line_num == lines_needed - 1:  # Last line gets remaining chars
                                    line_text = ''.join(qty_chars[start_idx:])
                                else:
                                    line_text = ''.join(qty_chars[start_idx:start_idx + chars_per_line])
                                
                                pdf.set_x(current_x + COLOR_WIDTH + DESC_WIDTH + 1)
                                pdf.cell(QTY_WIDTH - 2, 5, line_text, 0, align="C")
                                if line_num < lines_needed - 1:  # Don't move down after last line
                                    pdf.ln(5)
                            
                            # Move to next row
                            pdf.set_xy(current_x, current_y + cell_height)

                    # Enhanced total row with multi-line support
                    pdf.set_font("Arial", "B", 8.5)
                    total_display = str(int(total_qty))
                    
                    # Calculate if total needs multiple lines
                    total_width = pdf.get_string_width(total_display)
                    total_cell_width = QTY_WIDTH * 0.95  # Use 95% of quantity cell width
                    
                    if total_width <= total_cell_width:
                        # Single line - normal height
                        cell_height = 5
                        pdf.cell(COLOR_WIDTH, cell_height, "", 1)
                        pdf.cell(DESC_WIDTH, cell_height, "TOTAL:", 1, align="C")
                        pdf.cell(QTY_WIDTH, cell_height, total_display, 1, align="C")
                        pdf.ln()
                    else:
                        # Multi-line total - calculate needed height
                        lines_needed = int(total_width / total_cell_width) + 1
                        cell_height = 5 * lines_needed
                        
                        # Store current position
                        current_x = pdf.get_x()
                        current_y = pdf.get_y()
                        
                        # Draw empty color cell and description cell with increased height
                        pdf.cell(COLOR_WIDTH, cell_height, "", 1)
                        pdf.cell(DESC_WIDTH, cell_height, "TOTAL:", 1, align="C")
                        
                        # Draw total cell border first
                        pdf.cell(QTY_WIDTH, cell_height, "", 1)
                        
                        # Now add multi-line total text
                        pdf.set_xy(current_x + COLOR_WIDTH + DESC_WIDTH + 1, current_y + 1)
                        
                        # Split total into chunks that fit
                        total_chars = list(total_display)
                        chars_per_line = int(len(total_chars) / lines_needed)
                        
                        for line_num in range(lines_needed):
                            start_idx = line_num * chars_per_line
                            if line_num == lines_needed - 1:  # Last line gets remaining chars
                                line_text = ''.join(total_chars[start_idx:])
                            else:
                                line_text = ''.join(total_chars[start_idx:start_idx + chars_per_line])
                            
                            pdf.set_x(current_x + COLOR_WIDTH + DESC_WIDTH + 1)
                            pdf.cell(QTY_WIDTH - 2, 5, line_text, 0, align="C")
                            if line_num < lines_needed - 1:  # Don't move down after last line
                                pdf.ln(5)
                        
                        # Move to next section
                        pdf.set_xy(current_x, current_y + cell_height)
                        pdf.ln()
                    
                    pdf.ln(5)

                    # Enhanced logo section with database lookup and multi-line support
                    logo_info = get_logo_info(str(logo_sku).strip())
                    
                    # Check if logo image exists
                    logo_images = find_logo_images_by_sku(logo_sku)
                    if not logo_images:
                        print(f"Error: Logo image not found for SKU {logo_sku}. Skipping PDF generation.")
    
                        # Update report status
                        for idx, row_data in enumerate(report_data):
                            if row_data['Document Number'] == str(doc_num) and row_data['LOGO'] == str(logo_sku):
                                report_data[idx]['Execution Status'] = 'FAILED'
                                report_data[idx]['Error Message'] = f'Logo image not found for SKU: {logo_sku}'
    
                        continue  # Skip this group

                    # Calculate proportional widths that add up to usable_width
                    logo_sku_label_width = usable_width * 0.12   # Increased from 0.10 to 0.14
                    logo_sku_value_width = usable_width * 0.08   
                    logo_pos_label_width = usable_width * 0.17   
                    logo_pos_value_width = usable_width * 0.39   # Decreased from 0.44 to 0.40
                    stitch_label_width = usable_width * 0.14     # Increased from 0.13 to 0.18
                    stitch_value_width = usable_width * 0.10    # Decreased from 0.10 to 0.05
                    # Prepare values for multi-line processing
                    logo_display = str(logo_sku).strip()
                    logo_pos = ""
                    if logo_info and logo_info['logo_position']:
                        logo_pos = logo_info['logo_position']
                    elif "LOGO POSITION" in group.columns:
                        logo_pos = safe_get(group["LOGO POSITION"].iloc[0])
                    
                    stitch_count = ""
                    if logo_info and logo_info['stitch_count']:
                        stitch_count = str(logo_info['stitch_count']).replace('.0', '')
                    elif "STITCH COUNT" in group.columns:
                        stitch_count = safe_get(group["STITCH COUNT"].iloc[0]).replace('.0', '')
                    
                    # Enhanced logo section (simplified for space)
                    # Calculate heights needed for each field
                    # Set standard row height
                    standard_height = 5

                    # Check if all text fits in standard height using same padding as add_multiline_text_to_cell
                    pdf.set_font("Arial", "", 8.5)
                    padding = 2  # Same as used in add_multiline_text_to_cell (2 * 1)
                    logo_sku_fits = pdf.get_string_width(logo_display) <= (logo_sku_value_width - padding)
                    logo_pos_fits = pdf.get_string_width(logo_pos) <= (logo_pos_value_width - padding)
                    stitch_fits = pdf.get_string_width(stitch_count) <= (stitch_value_width - padding)

                    # If everything fits, use standard height; otherwise calculate needed height
                    if logo_sku_fits and logo_pos_fits and stitch_fits:
                        row_height = standard_height
                    else:
                        # Calculate heights only when needed
                        logo_sku_height = calculate_text_height(pdf, logo_display, logo_sku_value_width - 2)
                        logo_pos_height = calculate_text_height(pdf, logo_pos, logo_pos_value_width - 2)
                        stitch_height = calculate_text_height(pdf, stitch_count, stitch_value_width - 2)
                        row_height = max(logo_sku_height, logo_pos_height, stitch_height, standard_height)

                    # Store current position
                    current_x = pdf.get_x()
                    current_y = pdf.get_y()

                    # Draw LOGO SKU section
                    pdf.set_font("Arial", "B", 8.5)
                    add_multiline_text_to_cell(pdf, "LOGO SKU:", current_x, current_y, logo_sku_label_width, row_height, border=1, align="C")

                    pdf.set_font("Arial", "", 8.5)
                    add_multiline_text_to_cell(pdf, logo_display, current_x + logo_sku_label_width, current_y, logo_sku_value_width, row_height, border=1, align="C")

                    # Draw LOGO POSITION section
                    pdf.set_font("Arial", "B", 8.5)
                    add_multiline_text_to_cell(pdf, "LOGO POSITION:", current_x + logo_sku_label_width + logo_sku_value_width, current_y, logo_pos_label_width, row_height, border=1, align="C")

                    pdf.set_font("Arial", "", 8.5)
                    # Check if logo position needs yellow highlighting
                    if logo_pos.strip().upper() != "LEFT CHEST":
                        add_multiline_text_to_cell(pdf, logo_pos, current_x + logo_sku_label_width + logo_sku_value_width + logo_pos_label_width, current_y, logo_pos_value_width, row_height, border=1, align="L", fill=True)
                    else:
                        add_multiline_text_to_cell(pdf, logo_pos, current_x + logo_sku_label_width + logo_sku_value_width + logo_pos_label_width, current_y, logo_pos_value_width, row_height, border=1, align="L")

                    # Draw STITCH COUNT section
                    pdf.set_font("Arial", "B", 8.5)
                    add_multiline_text_to_cell(pdf, "STITCH COUNT:", current_x + logo_sku_label_width + logo_sku_value_width + logo_pos_label_width + logo_pos_value_width, current_y, stitch_label_width, row_height, border=1, align="C")

                    pdf.set_font("Arial", "", 8.5)
                    add_multiline_text_to_cell(pdf, stitch_count, current_x + logo_sku_label_width + logo_sku_value_width + logo_pos_label_width + logo_pos_value_width + stitch_label_width, current_y, stitch_value_width, row_height, border=1, align="C")

                    # Move to next section
                    pdf.set_xy(current_x, current_y + row_height + 2)

                    # Enhanced notes section with multi-line support
                    notes = ""
                    if logo_info and logo_info['notes']:
                        notes = logo_info['notes']
                    elif "NOTES" in group.columns:
                        notes = safe_get(group["NOTES"].iloc[0])

                    # Calculate height needed for notes
                   # Check if notes fit in standard height first
                    standard_notes_height = 5
                    pdf.set_font("Arial", "", 8.5)
                    notes_fits = pdf.get_string_width(notes) <= ((usable_width * 0.90) - 4)

                    if notes_fits:
                        notes_height = standard_notes_height
                    else:
                        # Calculate height only when needed
                        notes_height = calculate_text_height(pdf, notes, (usable_width * 0.90) - 2)

                    # Store current position for notes
                    notes_x = pdf.get_x()
                    notes_y = pdf.get_y()

                    # Draw NOTES section
                    pdf.set_font("Arial", "B", 8.5)
                    add_multiline_text_to_cell(pdf, "NOTES:", notes_x, notes_y, usable_width * 0.10, notes_height, border=1, align="C")

                    pdf.set_font("Arial", "", 8.5)
                    add_multiline_text_to_cell(pdf, notes, notes_x + (usable_width * 0.10), notes_y, usable_width * 0.90, notes_height, border=1, align="L")

                    # Move to next section
                    pdf.set_xy(notes_x, notes_y + notes_height + 5)                    

                    # Enhanced logo color table with actual colors
                    logo_colors = logo_info['logo_colors'] if logo_info else None
                    add_logo_color_table(pdf, logo_colors)

                    pdf.ln(2)
                    pdf.set_font("Arial", "B", 8.5)
                    pdf.cell(25, 5, "FILE NAME:", border=1, align="C")
                    pdf.set_font("Arial", "", 8.5)
                    
                    file_name = ""
                    if logo_info and logo_info['file_name']:
                        file_name = logo_info['file_name']
                    elif "FILE NAME" in group.columns:
                        file_name = safe_get(group["FILE NAME"].iloc[0])
                    
                    file_name_display = truncate_text(file_name, pdf, (usable_width - 25) * 0.95)
                    pdf.cell(usable_width - 25, 5, file_name_display, border=1)
                    pdf.ln(7)

                    # Add logo images
                    add_logo_images_to_pdf(pdf, str(logo_sku).strip(), logo_info)

                    # Generate filename
                    safe_doc_num = str(doc_num).replace("/", "_").replace("\\", "_")
                    safe_logo_sku = str(logo_sku).strip().replace("/", "_").replace("\\", "_")
                    filename = f"SO_{safe_doc_num}_AI_{safe_logo_sku}.pdf"
                    
                    pdf.output(os.path.join(OUTPUT_FOLDER, filename))
                    print(f"Generated PDF: {filename}")
                    pdf_count += 1
                    
                except Exception as e:
                    print(f"Error generating PDF for {doc_num}-{logo_sku}: {e}")
                    # Update report data for this group to show error
                    for idx, row_data in enumerate(report_data):
                        if (row_data['Document Number'] == str(doc_num) and 
                            row_data['LOGO'] == str(logo_sku)):
                            report_data[idx]['Execution Status'] = 'FAILED'
                            report_data[idx]['Error Message'] = f'PDF generation error: {str(e)}'

        # Step 7: Generate reports
        update_progress(session_id, 'processing', 85, 'Generating comprehensive reports...', 'Report Generation', 8)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        try:
            report_gen = ReportGenerator()
            
            # Create filter info string for reports
            filter_info = ""
            if sales_order_filter:
                filter_info += f"_SO_{sales_order_filter}"
            if approval_filter != "approved_only":
                approval_suffix = {
                    "not_approved_only": "_NotApproved",
                    "both": "_AllStatus"
                }.get(approval_filter, "")
                filter_info += approval_suffix
            
            report_gen.generate_all_reports(
                report_data=report_data,
                output_folder=OUTPUT_FOLDER,
                timestamp=timestamp,
                sales_order_filter=sales_order_filter,
                approval_filter=approval_filter,
                filter_info=filter_info
            )
        except Exception as e:
            print(f"Error generating reports: {e}")
        
        time.sleep(0.5)
        
        # Step 8: Create ZIP file
        zip_path = os.path.join(OUTPUT_FOLDER, ZIP_NAME)
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for fname in os.listdir(OUTPUT_FOLDER):
                if fname.endswith((".pdf", ".xlsx", ".txt", ".json")) and fname != ZIP_NAME:
                    # Check if it's a PDF with SO in the name
                    if fname.startswith("SO_") and "_AI_" in fname and fname.endswith(".pdf"):
                        # Extract SO number
                        try:
                            so_part = fname.split("_AI_")[0]  # e.g., SO_12345
                            so_number = so_part.replace("SO_", "")
                            arcname = os.path.join(so_number, fname)  # e.g., 12345/SO_12345_AI_0012.pdf
                        except:
                            arcname = fname
                    else:
                        arcname = fname  # non-PDF files go to root
            
                    zipf.write(os.path.join(OUTPUT_FOLDER, fname), arcname)
        
        # Completion
        success_msg = f"Successfully generated {pdf_count} art instruction PDF(s) with execution report"
        if sales_order_filter:
            success_msg += f" for Sales Order '{sales_order_filter}'"
        
        # Add approval filter info to success message
        if approval_filter == "approved_only":
            success_msg += " (Approved orders only)"
        elif approval_filter == "not_approved_only":
            success_msg += " (Not approved orders only)"
        elif approval_filter == "both":
            success_msg += " (Both approved and not approved orders)"
        
        update_progress(session_id, 'completed', 100, success_msg, 'Complete', 8)
        
        return {'success': True, 'message': success_msg, 'pdf_count': pdf_count}
        
    except Exception as e:
        update_progress(session_id, 'error', 0, f'Error during processing: {str(e)}', 'Error', 8)
        return {'success': False, 'error': str(e)}

@app.route("/", methods=["GET", "POST"])
def upload_file():
    # Load logo database on each request
    load_logo_database()
    
    if request.method == "POST":
        file = request.files["excel"]
        sales_order_filter = request.form.get("sales_order", "").strip()
        approval_filter = request.form.get("approval_filter", "approved_only").strip()  # NEW LINE
        
        if file.filename == "":
            return redirect(request.url)
            
        # Generate unique session ID for this processing task
        session_id = str(uuid.uuid4())
        session['processing_id'] = session_id
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)
        
        # Start processing in background thread
        def background_process():
            process_file_with_progress(file_path, sales_order_filter, session_id, approval_filter)  # MODIFIED LINE
        
        thread = threading.Thread(target=background_process)
        thread.daemon = True
        thread.start()
        
        # Redirect to progress page
        return redirect(url_for('progress_page', session_id=session_id))
    
    return render_template("upload.html")

@app.route("/progress/<session_id>")
def progress_page(session_id):
    """Display progress page"""
    return render_template("progress.html", session_id=session_id)

@app.route("/api/progress/<session_id>")
def get_progress(session_id):
    """API endpoint to get current progress"""
    if session_id in progress_status:
        return jsonify(progress_status[session_id])
    else:
        return jsonify({
            'status': 'not_found',
            'progress': 0,
            'message': 'Processing session not found',
            'current_step': '',
            'total_steps': 0
        })

@app.route("/download/<session_id>")
def download_file_with_session(session_id):
    """Download file after processing complete"""
    if session_id in progress_status and progress_status[session_id]['status'] == 'completed':
        # Clean up progress status
        del progress_status[session_id]
        return send_file(os.path.join(OUTPUT_FOLDER, ZIP_NAME), as_attachment=True)
    else:
        return redirect(url_for('upload_file'))

@app.route("/download")
def download_file():
    success_msg = request.args.get('success', '')
    return send_file(os.path.join(OUTPUT_FOLDER, ZIP_NAME), as_attachment=True)

if __name__ == "__main__":
    # Load logo database at startup
    load_logo_database()
    webbrowser.open("http://127.0.0.1:5000/")
    app.run(debug=True)