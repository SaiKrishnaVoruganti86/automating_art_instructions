from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
from fpdf import FPDF
import os
import zipfile
from werkzeug.utils import secure_filename
from PIL import Image  # Added for image dimension detection
from datetime import datetime  # Added for date formatting

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
LOGO_DB_FOLDER = "logo_database"  # New folder for logo database
LOGO_IMAGES_FOLDER = "logo_images"  # New folder for logo images
ZIP_NAME = "art_instructions_pdfs.zip"
LOGO_DB_FILE = "ArtDBSample.xlsx"  # Logo database file

# Create all necessary folders
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(LOGO_DB_FOLDER, exist_ok=True)
os.makedirs(LOGO_IMAGES_FOLDER, exist_ok=True)

# Global variable to store logo database
logo_database = None

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

def add_logo_color_table(pdf, logo_colors=None):
    """Enhanced logo color table with actual colors from database and truncation (using consistent width)"""
    pdf.ln(5)
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
    """Add logo images to PDF with intelligent sizing and layout, or display message if logo info not found"""
    if not logo_sku or pd.isna(logo_sku) or logo_sku == "":
        return
    
    # Check if logo info was found in database
    if logo_info is None:
        print(f"Logo info not found in database for SKU: {logo_sku}")
        
        # Add "logo info not found" message to PDF
        current_y = pdf.get_y() + 5
        pdf.set_xy(pdf.l_margin, current_y)
        pdf.set_font("Arial", "B", 12)
        pdf.set_text_color(255, 0, 0)  # Red text
        pdf.cell(0, 10, f"Logo info not found in database for SKU: {logo_sku}", align="C")
        pdf.ln(15)
        pdf.set_text_color(0, 0, 0)  # Reset to black text
        return
    
    # Find all images for this SKU
    logo_images = find_logo_images_by_sku(logo_sku)
    
    if not logo_images:
        print(f"No logo images found for SKU: {logo_sku}")
        
        # Add "no images found" message to PDF
        current_y = pdf.get_y() + 5
        pdf.set_xy(pdf.l_margin, current_y)
        pdf.set_font("Arial", "I", 10)
        pdf.set_text_color(128, 128, 128)  # Gray text
        pdf.cell(0, 8, f"No logo images found for SKU: {logo_sku}", align="C")
        pdf.ln(12)
        pdf.set_text_color(0, 0, 0)  # Reset to black text
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

@app.route("/", methods=["GET", "POST"])
def upload_file():
    # Load logo database on each request (or you could load it once at startup)
    load_logo_database()
    
    if request.method == "POST":
        file = request.files["excel"]
        sales_order_filter = request.form.get("sales_order", "").strip()
        
        if file.filename == "":
            return redirect(request.url)
        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

        # Read Excel or CSV file with LOGO column as string to preserve leading zeros
        df = read_file_with_format_detection(file_path)
        df.columns = [col.strip() for col in df.columns]
        
        # Apply sales order filter if provided
        if sales_order_filter:
            df = filter_by_sales_order(df, sales_order_filter)
            if df.empty:
                # Return to upload page with error message
                return render_template("upload.html", 
                                     error_message=f"No exact match found for Sales Order: '{sales_order_filter}'. Please enter the complete and exact sales order number.")
        
        # Clean and preserve LOGO format
        if 'LOGO' in df.columns:
            def clean_logo_value(logo_val):
                if pd.isna(logo_val) or logo_val in ['nan', 'NaN', '']:
                    return ""
                
                # Convert to string and clean
                logo_str = str(logo_val).strip()
                
                # Handle float-like strings (e.g., "9.0" -> "9")
                if logo_str.endswith('.0'):
                    logo_str = logo_str[:-2]
                
                # Skip empty or invalid values
                if logo_str in ['', 'nan', 'NaN', '0', '0000']:
                    return ""
                
                # Auto-pad numeric values to 4 digits for consistency
                if logo_str.isdigit() and len(logo_str) < 4:
                    logo_str = logo_str.zfill(4)  # Pad to 4 digits: "9" -> "0009"
                    
                return logo_str
            
            df['LOGO'] = df['LOGO'].apply(clean_logo_value)
            
            # Filter out rows with empty/invalid logos
            df = df[df['LOGO'] != ""]
            
            print("LOGO column processed to preserve original format")
            
            # Show sample of LOGO values for debugging
            sample_logos = df['LOGO'].dropna().unique()[:10]
            print(f"Sample LOGO values detected: {list(sample_logos)}")

        # DEBUG: Check OPERATIONAL CODE column
        print("=== DEBUGGING OPERATIONAL CODE COLUMN ===")
        print(f"Available columns: {list(df.columns)}")

        if 'OPERATIONAL CODE' in df.columns:
            print("\nOPERATIONAL CODE column found!")
            
            # Show unique values and their types
            unique_op_codes = df['OPERATIONAL CODE'].unique()
            print(f"\nUnique OPERATIONAL CODE values ({len(unique_op_codes)} total):")
            for i, code in enumerate(unique_op_codes[:20]):  # Show first 20
                print(f"  {i+1}. '{code}' (type: {type(code)}, pandas null: {pd.isna(code)})")
            
            # Check rows where OPERATIONAL CODE might be 90
            print(f"\nRows where OPERATIONAL CODE appears to be 90:")
            # Try different ways to find 90
            op_90_candidates = []
            
            # Method 1: Direct string comparison
            op_90_str = df[df['OPERATIONAL CODE'].astype(str).str.strip() == '90']
            op_90_candidates.extend(op_90_str.head(3).to_dict('records'))
            
            # Method 2: Numeric comparison (if possible)
            try:
                op_90_num = df[pd.to_numeric(df['OPERATIONAL CODE'], errors='coerce') == 90]
                op_90_candidates.extend(op_90_num.head(3).to_dict('records'))
            except:
                pass
            
            print(f"Found {len(op_90_candidates)} potential rows with OPERATIONAL CODE = 90")
            for i, row in enumerate(op_90_candidates[:5]):
                doc_num = row['Document Number']
                op_code = row['OPERATIONAL CODE']
                list_codes = row.get('List of Operation Codes', 'N/A')
                logo = row['LOGO']
                print(f"    {i+1}. Doc: {doc_num}, OPERATIONAL CODE: '{op_code}' (type: {type(op_code)}), List: '{list_codes}', Logo: '{logo}'")

        else:
            print("OPERATIONAL CODE column NOT FOUND!")
            print("Available columns:")
            for col in df.columns:
                print(f"  - '{col}'")

        print("=== END DEBUG ===\n")
        
        # Group by both Document Number AND Logo SKU to handle multiple logos per SO
        grouped = df.groupby(["Document Number", "LOGO"])

        # Clear output folder
        for f in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, f))

        pdf_count = 0  # Track number of PDFs generated

        # ENHANCED FILTERING LOGIC WITH CORRECT ORDER
        for (doc_num, logo_sku), group in grouped:
            # Filter 1: Skip entries with no valid logo SKU
            if pd.isna(logo_sku) or str(logo_sku).strip() in ["", "0", "0000"]:
                print(f"Skipping Document {doc_num} - No valid logo SKU")
                continue
            
            # Filter 2: Skip "Not Approved" orders using DueDateStatus column
            due_date_status = safe_get(group["DueDateStatus"].iloc[0]) if "DueDateStatus" in group.columns else ""
            if due_date_status.strip().upper() == "NOT APPROVED":
                print(f"Skipping Document {doc_num} - Status: Not Approved")
                continue
            
            # Filter 3: Check LOGO validity first
            logo_sku_str = str(logo_sku).strip()
            
            # LOGO is not valid if it's "0000", "0", or empty
            if logo_sku_str in ["0000", "0", ""]:
                print(f"Skipping Document {doc_num} - Invalid Logo SKU: '{logo_sku_str}'")
                continue
            
            # Filter 4: Check OPERATIONAL CODE validity
            operational_code = None
            if "OPERATIONAL CODE" in group.columns:
                # Get the operational code for this group (should be same for all rows in group)
                op_code_raw = group["OPERATIONAL CODE"].iloc[0]
                print(f"Document {doc_num} - Raw OPERATIONAL CODE: '{op_code_raw}' (type: {type(op_code_raw)})")
                
                if pd.notna(op_code_raw) and str(op_code_raw).strip():
                    op_code_str = str(op_code_raw).strip()
                    print(f"Document {doc_num} - OPERATIONAL CODE string: '{op_code_str}'")
                    
                    # OPERATIONAL CODE is not valid if it's "00", "0", or empty
                    if op_code_str not in ["00", "0", ""]:
                        try:
                            # Handle both integer and float formats
                            if '.' in op_code_str:
                                operational_code = int(float(op_code_str))
                            else:
                                operational_code = int(op_code_str)
                            print(f"Document {doc_num} - Parsed OPERATIONAL CODE: {operational_code}")
                        except (ValueError, TypeError) as e:
                            print(f"Document {doc_num} - Error parsing OPERATIONAL CODE '{op_code_str}': {e}")
                    else:
                        print(f"Document {doc_num} - OPERATIONAL CODE '{op_code_str}' is invalid (00, 0, or empty)")
                else:
                    print(f"Document {doc_num} - OPERATIONAL CODE is null or empty")
            else:
                print(f"Document {doc_num} - OPERATIONAL CODE column not found")
            
            # If OPERATIONAL CODE is not valid, skip
            if operational_code is None:
                print(f"Skipping Document {doc_num} - Invalid or missing Operational Code")
                continue
            
            print(f"Document {doc_num} - Valid Logo: {logo_sku_str}, Valid Operational Code: {operational_code}")
            
            # Filter 5: Check OPERATIONAL CODE conditions
            
            # Sub-case 1: If OPERATIONAL CODE is 11, generate regardless of List of Operation Codes
            if operational_code == 11:
                print(f"✓ Document {doc_num} - Operational Code is 11, generating art instruction")
            
            # Sub-case 2: If OPERATIONAL CODE > 89, check List of Operation Codes
            elif operational_code > 89:
                print(f"Document {doc_num} - Operational Code {operational_code} > 89, checking List of Operation Codes")
                
                # Get List of Operation Codes
                list_operation_codes = []
                if "List of Operation Codes" in group.columns:
                    list_codes_raw = group["List of Operation Codes"].iloc[0]
                    print(f"Document {doc_num} - Raw List of Operation Codes: '{list_codes_raw}' (type: {type(list_codes_raw)})")
                    
                    if pd.notna(list_codes_raw) and str(list_codes_raw).strip():
                        list_codes_str = str(list_codes_raw).strip()
                        print(f"Document {doc_num} - List of Operation Codes string: '{list_codes_str}'")
                        
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
                    else:
                        print(f"Document {doc_num} - List of Operation Codes is empty or null")
                else:
                    print(f"Document {doc_num} - List of Operation Codes column not found")
                
                print(f"Document {doc_num} - Parsed List of Operation Codes: {list_operation_codes}")
                
                # Check conditions for List of Operation Codes
                if not list_operation_codes:
                    print(f"Skipping Document {doc_num} - No valid List of Operation Codes found")
                    continue
                
                # Must contain exactly one 11 (mandatory)
                count_of_11 = list_operation_codes.count(11)
                if count_of_11 != 1:
                    print(f"Skipping Document {doc_num} - List must contain exactly one 11 (found {count_of_11})")
                    continue
                
                # No operation code should be less than 60 (except 11, which is required)
                codes_less_than_60 = [code for code in list_operation_codes if code < 60 and code != 11]
                if codes_less_than_60:
                    print(f"Skipping Document {doc_num} - List contains codes < 60 (excluding 11): {codes_less_than_60}")
                    continue
                
                print(f"✓ Document {doc_num} - All List of Operation Codes conditions satisfied, generating art instruction")
            
            # Sub-case 3: OPERATIONAL CODE is anything other than 11 and not > 89
            else:
                print(f"Skipping Document {doc_num} - Operational Code {operational_code} is not 11 and not > 89")
                continue
            
            # If we reach here, all conditions are satisfied
            print(f"✓ PROCESSING Document {doc_num} with Logo SKU: {logo_sku_str}, Operational Code: {operational_code}")
            
            pdf = FPDF(orientation="P", unit="mm", format=(190.5, 254.0))
            pdf.set_margins(0.8, 0.8, 0.8)
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=0.8)
            pdf.set_font("Arial", "", 8.5)

            due_date = format_date_consistently(group["Due Date"].iloc[0])

            full_width = 190
            usable_width = full_width - (2 * 0.8)
            left_width = full_width * 0.75
            right_width = full_width - left_width

            # Now calculate client_name after left_width is defined
            client_name = truncate_text(safe_get(group["Customer/Vendor Name"].iloc[0]), pdf, (left_width - 20) * 0.95)

            pdf.set_font("Arial", "B", 10)
            pdf.cell(left_width, 8, "ART INSTRUCTIONS", border=1, align="C")
            pdf.cell(right_width, 8, "", border=0)
            pdf.image("static/jauniforms.png", x=pdf.get_x() - right_width + 3, y=pdf.get_y() + 1, w=right_width - 6)
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
                
                # Debug output to check truncation
                print(f"Original description: '{group_data['description']}'")
                print(f"Truncated description: '{desc_display}'")
                print(f"Available width: {DESC_WIDTH * 0.90}")
                print(f"Text width: {pdf.get_string_width(desc_display)}")
                
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

            # Enhanced total row with multi-line support (immediately after the loop, no extra line break)
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
            
            pdf.ln(7)

            # Enhanced logo section with database lookup and multi-line support (using full usable width)
            logo_info = get_logo_info(logo_sku_str)
            
            # Calculate proportional widths that add up to usable_width (your specified sizing)
            logo_sku_label_width = usable_width * 0.10   # Logo SKU label - 10%
            logo_sku_value_width = usable_width * 0.08   # Logo SKU value - 8%
            logo_pos_label_width = usable_width * 0.15   # Logo Position label - 15%
            logo_pos_value_width = usable_width * 0.44   # Logo Position value - 44%
            stitch_label_width = usable_width * 0.13     # Stitch Count label - 13%
            stitch_value_width = usable_width * 0.10     # Stitch Count value - 10%
            
            # Debug: verify total width doesn't exceed usable_width
            total_logo_width = (logo_sku_label_width + logo_sku_value_width + 
                              logo_pos_label_width + logo_pos_value_width + 
                              stitch_label_width + stitch_value_width)
            print(f"Logo section total width: {total_logo_width:.2f}mm, usable width: {usable_width:.2f}mm")
            
            # Ensure we don't exceed usable width (safety check)
            if total_logo_width > usable_width:
                scale_factor = usable_width / total_logo_width
                logo_sku_label_width *= scale_factor
                logo_sku_value_width *= scale_factor
                logo_pos_label_width *= scale_factor
                logo_pos_value_width *= scale_factor
                stitch_label_width *= scale_factor
                stitch_value_width *= scale_factor
                print(f"Applied scaling factor: {scale_factor:.3f}")
            
            # Prepare values for multi-line processing
            logo_display = logo_sku_str
            logo_pos = ""
            if logo_info and logo_info['logo_position']:
                logo_pos = logo_info['logo_position']
            elif "LOGO POSITION" in group.columns:
                logo_pos = safe_get(group["LOGO POSITION"].iloc[0])
            
            stitch_count = ""
            if logo_info and logo_info['stitch_count']:
                stitch_count = str(logo_info['stitch_count'])
            elif "STITCH COUNT" in group.columns:
                stitch_count = safe_get(group["STITCH COUNT"].iloc[0])
            
            # Calculate available widths for each field (95% of cell width)
            logo_sku_width = logo_sku_value_width * 0.95
            logo_pos_width = logo_pos_value_width * 0.95
            stitch_count_width = stitch_value_width * 0.95
            
            # Check which fields need multiple lines
            pdf.set_font("Arial", "", 8.5)
            logo_text_width = pdf.get_string_width(logo_display)
            pos_text_width = pdf.get_string_width(logo_pos)
            stitch_text_width = pdf.get_string_width(stitch_count)
            
            # Calculate lines needed for each field
            logo_lines = max(1, int(logo_text_width / logo_sku_width) + 1) if logo_text_width > logo_sku_width else 1
            pos_lines = max(1, int(pos_text_width / logo_pos_width) + 1) if pos_text_width > logo_pos_width else 1
            stitch_lines = max(1, int(stitch_text_width / stitch_count_width) + 1) if stitch_text_width > stitch_count_width else 1
            
            # Use the maximum lines needed for consistent row height
            max_lines = max(logo_lines, pos_lines, stitch_lines)
            cell_height = 5 * max_lines
            
            # Store current position
            current_x = pdf.get_x()
            current_y = pdf.get_y()
            
            # Draw cell borders first (using proportional widths)
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(logo_sku_label_width, cell_height, "", border=1)  # Logo SKU label cell
            pdf.set_font("Arial", "", 8.5)
            pdf.cell(logo_sku_value_width, cell_height, "", border=1)  # Logo SKU value cell
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(logo_pos_label_width, cell_height, "", border=1)  # Logo Position label cell
            pdf.set_font("Arial", "", 8.5)
            pdf.cell(logo_pos_value_width, cell_height, "", border=1)  # Logo Position value cell
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(stitch_label_width, cell_height, "", border=1)  # Stitch Count label cell
            pdf.set_font("Arial", "", 8.5)
            pdf.cell(stitch_value_width, cell_height, "", border=1)  # Stitch Count value cell
            
            # Now add the labels (centered vertically)
            label_y_offset = (cell_height - 5) / 2
            
            # Logo SKU label
            pdf.set_xy(current_x, current_y + label_y_offset)
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(logo_sku_label_width, 5, "LOGO SKU:", align="C")
            
            # Logo Position label
            pdf.set_xy(current_x + logo_sku_label_width + logo_sku_value_width, current_y + label_y_offset)
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(logo_pos_label_width, 5, "LOGO POSITION:", align="C")
            
            # Stitch Count label
            pdf.set_xy(current_x + logo_sku_label_width + logo_sku_value_width + logo_pos_label_width + logo_pos_value_width, current_y + label_y_offset)
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(stitch_label_width, 5, "STITCH COUNT:", align="C")
            
            # Add multi-line values
            pdf.set_font("Arial", "", 8.5)
            
            # Logo SKU value (multi-line if needed)
            if logo_lines > 1:
                logo_chars = list(logo_display)
                chars_per_line = max(1, len(logo_chars) // logo_lines)
                for line_num in range(logo_lines):
                    start_idx = line_num * chars_per_line
                    if line_num == logo_lines - 1:
                        line_text = ''.join(logo_chars[start_idx:])
                    else:
                        line_text = ''.join(logo_chars[start_idx:start_idx + chars_per_line])
                    
                    line_y = current_y + (line_num * 5) + ((cell_height - (logo_lines * 5)) / 2)
                    pdf.set_xy(current_x + logo_sku_label_width + 1, line_y)
                    pdf.cell(logo_sku_value_width - 2, 5, line_text, align="C")
            else:
                pdf.set_xy(current_x + logo_sku_label_width + 1, current_y + label_y_offset)
                pdf.cell(logo_sku_value_width - 2, 5, logo_display, align="C")
            
            # Logo Position value (multi-line if needed)
            if pos_lines > 1:
                pos_chars = list(logo_pos)
                chars_per_line = max(1, len(pos_chars) // pos_lines)
                for line_num in range(pos_lines):
                    start_idx = line_num * chars_per_line
                    if line_num == pos_lines - 1:
                        line_text = ''.join(pos_chars[start_idx:])
                    else:
                        line_text = ''.join(pos_chars[start_idx:start_idx + chars_per_line])
                    
                    line_y = current_y + (line_num * 5) + ((cell_height - (pos_lines * 5)) / 2)
                    pdf.set_xy(current_x + logo_sku_label_width + logo_sku_value_width + logo_pos_label_width + 1, line_y)
                    pdf.cell(logo_pos_value_width - 2, 5, line_text, align="L")
            else:
                pdf.set_xy(current_x + logo_sku_label_width + logo_sku_value_width + logo_pos_label_width + 1, current_y + label_y_offset)
                pdf.cell(logo_pos_value_width - 2, 5, logo_pos, align="L")
            
            # Stitch Count value (multi-line if needed)
            if stitch_lines > 1:
                stitch_chars = list(stitch_count)
                chars_per_line = max(1, len(stitch_chars) // stitch_lines)
                for line_num in range(stitch_lines):
                    start_idx = line_num * chars_per_line
                    if line_num == stitch_lines - 1:
                        line_text = ''.join(stitch_chars[start_idx:])
                    else:
                        line_text = ''.join(stitch_chars[start_idx:start_idx + chars_per_line])
                    
                    line_y = current_y + (line_num * 5) + ((cell_height - (stitch_lines * 5)) / 2)
                    pdf.set_xy(current_x + logo_sku_label_width + logo_sku_value_width + logo_pos_label_width + logo_pos_value_width + stitch_label_width + 1, line_y)
                    pdf.cell(stitch_value_width - 2, 5, line_text, align="C")
            else:
                pdf.set_xy(current_x + logo_sku_label_width + logo_sku_value_width + logo_pos_label_width + logo_pos_value_width + stitch_label_width + 1, current_y + label_y_offset)
                pdf.cell(stitch_value_width - 2, 5, stitch_count, align="C")
            
            # Move to next section
            pdf.set_xy(current_x, current_y + cell_height)
            
            # Minimal consistent spacing after logo section
            pdf.ln(2)

            # Enhanced notes section with multi-line support
            pdf.set_font("Arial", "B", 8.5)
            
            # Get notes value from database or original data
            notes = ""
            if logo_info and logo_info['notes']:
                notes = logo_info['notes']
            elif "NOTES" in group.columns:
                notes = safe_get(group["NOTES"].iloc[0])
            
            # Calculate available width for notes (95% of notes value cell)
            notes_value_width = (usable_width * 0.90) * 0.95  # 95% of the 90% notes value cell
            
            # Check if notes need multiple lines
            pdf.set_font("Arial", "", 8.5)
            notes_text_width = pdf.get_string_width(notes)
            
            if notes_text_width <= notes_value_width:
                # Single line - normal height
                cell_height = 5
                pdf.set_font("Arial", "B", 8.5)
                pdf.cell(usable_width * 0.10, cell_height, "NOTES:", border=1, align="C")
                pdf.set_font("Arial", "", 8.5)
                pdf.cell(usable_width * 0.90, cell_height, notes, border=1)
            else:
                # Multi-line notes - calculate needed height using word-based wrapping
                # Split text into words for better line breaks
                words = notes.split()
                lines = []
                current_line = ""
                
                # Build lines by adding words until width limit is reached
                for word in words:
                    test_line = current_line + (" " if current_line else "") + word
                    test_width = pdf.get_string_width(test_line)
                    
                    if test_width <= notes_value_width:
                        current_line = test_line
                    else:
                        if current_line:  # If current line has content, save it
                            lines.append(current_line)
                            current_line = word
                        else:  # Single word is too long, need to break it
                            # For very long single words, break by characters
                            while word:
                                char_line = ""
                                for char in word:
                                    if pdf.get_string_width(char_line + char) <= notes_value_width:
                                        char_line += char
                                    else:
                                        break
                                if char_line:
                                    lines.append(char_line)
                                    word = word[len(char_line):]
                                else:
                                    # Single character is too wide (shouldn't happen)
                                    lines.append(word[0])
                                    word = word[1:]
                            current_line = ""
                
                # Add the last line if it has content
                if current_line:
                    lines.append(current_line)
                
                lines_needed = len(lines)
                cell_height = 5 * lines_needed
                
                print(f"Notes text: '{notes}'")
                print(f"Split into {lines_needed} lines:")
                for i, line in enumerate(lines):
                    print(f"  Line {i+1}: '{line}'")
                
                # Store current position
                current_x = pdf.get_x()
                current_y = pdf.get_y()
                
                # Draw cell borders first
                pdf.set_font("Arial", "B", 8.5)
                pdf.cell(usable_width * 0.10, cell_height, "", border=1)  # Notes label cell
                pdf.set_font("Arial", "", 8.5)
                pdf.cell(usable_width * 0.90, cell_height, "", border=1)  # Notes value cell
                
                # Add NOTES label (centered vertically)
                label_y_offset = (cell_height - 5) / 2
                pdf.set_xy(current_x, current_y + label_y_offset)
                pdf.set_font("Arial", "B", 8.5)
                pdf.cell(usable_width * 0.10, 5, "NOTES:", align="C")
                
                # Add multi-line notes value using the word-wrapped lines
                pdf.set_font("Arial", "", 8.5)
                for line_num, line_text in enumerate(lines):
                    # Calculate Y position for this line (centered within the multi-line area)
                    line_y = current_y + (line_num * 5) + ((cell_height - (lines_needed * 5)) / 2)
                    pdf.set_xy(current_x + (usable_width * 0.10) + 1, line_y)
                    pdf.cell((usable_width * 0.90) - 2, 5, line_text, align="L")
                
                # Move to next section
                pdf.set_xy(current_x, current_y + cell_height)
                
                # Minimal consistent spacing after notes section
                pdf.ln(2)

            # Enhanced logo color table with actual colors
            logo_colors = logo_info['logo_colors'] if logo_info else None
            add_logo_color_table(pdf, logo_colors)

            # Minimal consistent spacing after logo color table
            pdf.ln(2)
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(25, 5, "FILE NAME:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            # Use file name from database if available, otherwise from original data
            file_name = ""
            if logo_info and logo_info['file_name']:
                file_name = logo_info['file_name']
            elif "FILE NAME" in group.columns:
                file_name = safe_get(group["FILE NAME"].iloc[0])
            
            # Truncate file name to use 95% of available space
            file_name_display = truncate_text(file_name, pdf, (usable_width - 25) * 0.95)
            pdf.cell(usable_width - 25, 5, file_name_display, border=1)
            
            # Minimal consistent spacing after file name table
            pdf.ln(2)

            # Add logo images based on preserved SKU number and pass logo_info
            pdf.ln(5)  # Add some space before images
            add_logo_images_to_pdf(pdf, logo_sku_str, logo_info)

            # Generate filename: SO_SO#_AI_LOGO SKU.pdf
            safe_doc_num = str(doc_num).replace("/", "_").replace("\\", "_")
            safe_logo_sku = logo_sku_str.replace("/", "_").replace("\\", "_")
            filename = f"SO_{safe_doc_num}_AI_{safe_logo_sku}.pdf"
            
            pdf.output(os.path.join(OUTPUT_FOLDER, filename))
            print(f"Generated PDF: {filename}")
            pdf_count += 1

        # Check if any PDFs were generated
        if pdf_count == 0:
            if sales_order_filter:
                error_msg = f"No art instructions generated for Sales Order '{sales_order_filter}'. Please check that the sales order exists and meets the processing criteria."
            else:
                error_msg = "No art instructions generated. Please check that your data meets the processing criteria (valid logos, operational codes, etc.)."
            
            return render_template("upload.html", error_message=error_msg)

        # Create ZIP file
        zip_path = os.path.join(OUTPUT_FOLDER, ZIP_NAME)
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for fname in os.listdir(OUTPUT_FOLDER):
                if fname.endswith(".pdf"):
                    zipf.write(os.path.join(OUTPUT_FOLDER, fname), fname)

        # Success message
        success_msg = f"Successfully generated {pdf_count} art instruction PDF(s)"
        if sales_order_filter:
            success_msg += f" for Sales Order '{sales_order_filter}'"
        
        return redirect(url_for("download_file", success=success_msg))

    return render_template("upload.html")

@app.route("/download")
def download_file():
    success_msg = request.args.get('success', '')
    return send_file(os.path.join(OUTPUT_FOLDER, ZIP_NAME), as_attachment=True)

if __name__ == "__main__":
    # Load logo database at startup
    load_logo_database()
    app.run(debug=True)