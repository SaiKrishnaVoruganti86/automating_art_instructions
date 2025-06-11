from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
from fpdf import FPDF
import os
import zipfile
from werkzeug.utils import secure_filename
from PIL import Image  # Added for image dimension detection

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
    """Enhanced logo color table with actual colors from database"""
    pdf.ln(5)
    total_width = 190.5 - (2 * 0.8)
    logo_color_width = total_width * 0.20
    number_width = total_width * 0.05
    value_width = total_width * 0.35

    # First row: LOGO COLOR header
    pdf.set_font("Arial", "B", 8.5)
    pdf.cell(logo_color_width, 5, "LOGO COLOR:", border=1, align="C")
    pdf.set_font("Arial", "", 8.5)
    
    # Add first color if available
    color1 = logo_colors[0] if logo_colors and len(logo_colors) > 0 else ""
    pdf.cell(number_width, 5, "1", border=1, align="C")
    pdf.cell(value_width, 5, color1, border=1)
    
    # Add ninth color if available
    color9 = logo_colors[8] if logo_colors and len(logo_colors) > 8 else ""
    pdf.cell(number_width, 5, "9", border=1, align="C")
    pdf.cell(value_width, 5, color9, border=1)
    pdf.ln()

    # Second row: PRODUCTION DAY directly under LOGO COLOR
    pdf.set_font("Arial", "B", 8.5)
    pdf.cell(logo_color_width, 5, "PRODUCTION DAY:", border=1, align="C")
    pdf.set_font("Arial", "", 8.5)
    
    # Add second and tenth colors if available
    color2 = logo_colors[1] if logo_colors and len(logo_colors) > 1 else ""
    color10 = logo_colors[9] if logo_colors and len(logo_colors) > 9 else ""
    pdf.cell(number_width, 5, "2", border=1, align="C")
    pdf.cell(value_width, 5, color2, border=1)
    pdf.cell(number_width, 5, "10", border=1, align="C")
    pdf.cell(value_width, 5, color10, border=1)
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
        
        pdf.cell(number_width, 5, str(i), border=1, align="C")
        pdf.cell(value_width, 5, color_left, border=1)
        pdf.cell(number_width, 5, str(i + 8), border=1, align="C")
        pdf.cell(value_width, 5, color_right, border=1)
        # Move to next line, but stay at the same x position (after the merged cell)
        pdf.set_xy(current_x + logo_color_width, pdf.get_y() + 5)

    # Last row with only left half filled (number 8), right half blank
    color8 = logo_colors[7] if logo_colors and len(logo_colors) > 7 else ""
    pdf.cell(number_width, 5, "8", border=1, align="C")
    pdf.cell(value_width, 5, color8, border=1)
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

def optimize_image_layout(image_info, available_width, available_height, margin=5):
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

@app.route("/", methods=["GET", "POST"])
def upload_file():
    # Load logo database on each request (or you could load it once at startup)
    load_logo_database()
    
    if request.method == "POST":
        file = request.files["excel"]
        if file.filename == "":
            return redirect(request.url)
        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

        # Read Excel file with LOGO column as string to preserve leading zeros
        try:
            # First, try reading with LOGO as string type
            df = pd.read_excel(file_path, dtype={'LOGO': str})
        except:
            # Fallback: read normally and convert LOGO to string
            df = pd.read_excel(file_path)
            if 'LOGO' in df.columns:
                df['LOGO'] = df['LOGO'].astype(str)
        
        df.columns = [col.strip() for col in df.columns]
        
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
                
                # IMPORTANT: If you want automatic padding for short numeric values,
                # uncomment the lines below. This will convert "9" to "0009"
                # if logo_str.isdigit() and len(logo_str) < 4:
                #     logo_str = logo_str.zfill(4)  # Pad to 4 digits
                    
                return logo_str
            
            df['LOGO'] = df['LOGO'].apply(clean_logo_value)
            
            # Filter out rows with empty/invalid logos
            df = df[df['LOGO'] != ""]
            
            print("LOGO column processed to preserve original format")
            
            # Show sample of LOGO values for debugging
            sample_logos = df['LOGO'].dropna().unique()[:10]
            print(f"Sample LOGO values detected: {list(sample_logos)}")
            
            # Warning if we detect short numeric values that might need leading zeros
            short_numeric = [logo for logo in sample_logos if logo.isdigit() and len(logo) < 4]
            if short_numeric:
                print(f"⚠️  WARNING: Found short numeric LOGO values that might need leading zeros: {short_numeric}")
                print("   If these should have leading zeros (e.g., '9' should be '0009'):")
                print("   1. Format the LOGO column as Text in Excel before entering data")
                print("   2. Or uncomment the auto-padding code in clean_logo_value function")
        
        # Group by both Document Number AND Logo SKU to handle multiple logos per SO
        grouped = df.groupby(["Document Number", "LOGO"])

        # Clear output folder
        for f in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, f))

        for (doc_num, logo_sku), group in grouped:
            # Skip entries with no logo or default logo
            if pd.isna(logo_sku) or str(logo_sku).strip() in ["", "0", "0000"]:
                print(f"Skipping Document {doc_num} - No valid logo SKU")
                continue
                
            # Preserve original logo SKU format
            logo_sku_str = str(logo_sku).strip()
            
            print(f"Processing Document {doc_num} with Logo SKU: {logo_sku_str}")
            
            pdf = FPDF(orientation="P", unit="mm", format=(190.5, 254.0))
            pdf.set_margins(0.8, 0.8, 0.8)
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=0.8)
            pdf.set_font("Arial", "", 8.5)

            client_name = truncate_text(safe_get(group["Customer/Vendor Name"].iloc[0]), pdf, 72)
            due_date = str(group["Due Date"].iloc[0]).split(" ")[0]

            full_width = 190
            usable_width = full_width - (2 * 0.8)
            left_width = full_width * 0.75
            right_width = full_width - left_width

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
            pdf.cell(so_section_width - 20, 6, str(doc_num), border=1)

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

            # Display grouped results
            pdf.set_font("Arial", "", 8.5)
            for group_key, group_data in color_desc_groups.items():
                color_display = truncate_text(group_data['color'], pdf, COLOR_WIDTH * 0.90)
                desc_display = group_data['description']
                qty_display = str(int(group_data['quantity']))
                
                pdf.cell(COLOR_WIDTH, 5, color_display, 1, align="C")
                pdf.cell(DESC_WIDTH, 5, desc_display, 1, align="C")
                pdf.cell(QTY_WIDTH, 5, qty_display, 1, align="C")
                pdf.ln()

            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(COLOR_WIDTH, 5, "", 1)
            pdf.cell(DESC_WIDTH, 5, "TOTAL:", 1, align="C")
            pdf.cell(QTY_WIDTH, 5, str(int(total_qty)), 1, align="C")
            pdf.ln(7)

            # Enhanced logo section with database lookup (using preserved SKU format)
            logo_info = get_logo_info(logo_sku_str)
            
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(18.89, 5, "LOGO SKU:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            logo_display = truncate_text(logo_sku_str, pdf, 15.11 * 0.90)
            pdf.cell(15.11, 5, logo_display, border=1, align="C")

            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(28.34, 5, "LOGO POSITION:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            # Use logo position from database if available, otherwise from original data
            logo_pos = ""
            if logo_info and logo_info['logo_position']:
                logo_pos = logo_info['logo_position']
            elif "LOGO POSITION" in group.columns:
                logo_pos = safe_get(group["LOGO POSITION"].iloc[0])
            pdf.cell(83.12, 5, logo_pos, border=1)

            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(24.56, 5, "STITCH COUNT:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            # Use stitch count from database if available, otherwise from original data
            stitch_count = ""
            if logo_info and logo_info['stitch_count']:
                stitch_count = logo_info['stitch_count']
            elif "STITCH COUNT" in group.columns:
                stitch_count = safe_get(group["STITCH COUNT"].iloc[0])
            pdf.cell(18.89, 5, str(stitch_count), border=1)
            pdf.ln(7)

            # Enhanced notes section
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(usable_width * 0.10, 5, "NOTES:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            # Use notes from database if available, otherwise from original data
            notes = ""
            if logo_info and logo_info['notes']:
                notes = logo_info['notes']
            elif "NOTES" in group.columns:
                notes = safe_get(group["NOTES"].iloc[0])
            pdf.cell(usable_width * 0.90, 5, notes, border=1)
            pdf.ln(2)

            # Enhanced logo color table with actual colors
            logo_colors = logo_info['logo_colors'] if logo_info else None
            add_logo_color_table(pdf, logo_colors)

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
            pdf.cell(usable_width - 25, 5, file_name, border=1)
            pdf.ln(8)

            # Add logo images based on preserved SKU number and pass logo_info
            pdf.ln(5)  # Add some space before images
            add_logo_images_to_pdf(pdf, logo_sku_str, logo_info)

            # Generate filename: SO_SO#_AI_LOGO SKU.pdf
            safe_doc_num = str(doc_num).replace("/", "_").replace("\\", "_")
            safe_logo_sku = logo_sku_str.replace("/", "_").replace("\\", "_")
            filename = f"SO_{safe_doc_num}_AI_{safe_logo_sku}.pdf"
            
            pdf.output(os.path.join(OUTPUT_FOLDER, filename))
            print(f"Generated PDF: {filename}")

        # Create ZIP file
        zip_path = os.path.join(OUTPUT_FOLDER, ZIP_NAME)
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for fname in os.listdir(OUTPUT_FOLDER):
                if fname.endswith(".pdf"):
                    zipf.write(os.path.join(OUTPUT_FOLDER, fname), fname)

        return redirect(url_for("download_file"))

    return render_template("upload.html")

@app.route("/download")
def download_file():
    return send_file(os.path.join(OUTPUT_FOLDER, ZIP_NAME), as_attachment=True)

if __name__ == "__main__":
    # Load logo database at startup
    load_logo_database()
    app.run(debug=True)