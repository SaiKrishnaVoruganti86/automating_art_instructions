from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
from fpdf import FPDF
import os
import zipfile
from werkzeug.utils import secure_filename

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
    """Get logo information from the database based on SKU"""
    if logo_database is None or pd.isna(logo_sku) or logo_sku == "" or logo_sku == "0000":
        return None
    
    try:
        # Convert logo_sku to string for comparison
        logo_sku_str = str(int(float(logo_sku))) if str(logo_sku).replace('.', '').isdigit() else str(logo_sku)
        
        # Search for the logo SKU in the database
        logo_row = logo_database[logo_database['Logo SKU'].astype(str) == logo_sku_str]
        
        if not logo_row.empty:
            row = logo_row.iloc[0]
            return {
                'logo_sku': safe_get(row['Logo SKU']),
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

def find_logo_image(file_name):
    """Find logo image file in the logo images folder"""
    if not file_name or pd.isna(file_name):
        return None
    
    # Common image extensions
    extensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff']
    
    for ext in extensions:
        # Try with exact filename
        image_path = os.path.join(LOGO_IMAGES_FOLDER, f"{file_name}{ext}")
        if os.path.exists(image_path):
            return image_path
        
        # Try without extension if filename already has one
        base_name = os.path.splitext(file_name)[0]
        image_path = os.path.join(LOGO_IMAGES_FOLDER, f"{base_name}{ext}")
        if os.path.exists(image_path):
            return image_path
    
    return None

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

def add_logo_image_to_pdf(pdf, logo_info):
    """Add logo image to PDF if available"""
    if not logo_info or not logo_info.get('file_name'):
        return
    
    image_path = find_logo_image(logo_info['file_name'])
    if image_path:
        try:
            # Add logo image in a designated area
            current_y = pdf.get_y()
            # You can adjust these coordinates based on your layout needs
            pdf.image(image_path, x=150, y=current_y, w=30, h=20)
            print(f"Added logo image: {image_path}")
        except Exception as e:
            print(f"Error adding logo image {image_path}: {e}")

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

        df = pd.read_excel(file_path)
        df.columns = [col.strip() for col in df.columns]
        grouped = df.groupby("Document Number")

        # Clear output folder
        for f in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, f))

        for doc_num, group in grouped:
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

            total_qty = 0
            pdf.set_font("Arial", "", 8.5)
            for _, row in group.iterrows():
                color = truncate_text(safe_get(row.get("COLOR")), pdf, COLOR_WIDTH * 0.90)
                desc = safe_get(row.get("SUBCATEGORY"))
                try:
                    qty = float(row.get("Quantity"))
                except:
                    qty = 0
                total_qty += qty
                pdf.cell(COLOR_WIDTH, 5, color, 1, align="C")
                pdf.cell(DESC_WIDTH, 5, desc, 1, align="C")
                pdf.cell(QTY_WIDTH, 5, str(int(qty)), 1, align="C")
                pdf.ln()

            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(COLOR_WIDTH, 5, "", 1)
            pdf.cell(DESC_WIDTH, 5, "TOTAL:", 1, align="C")
            pdf.cell(QTY_WIDTH, 5, str(int(total_qty)), 1, align="C")
            pdf.ln(7)

            # Enhanced logo section with database lookup
            raw_logo = safe_get(group["LOGO"].iloc[0]) if "LOGO" in group.columns else ""
            try:
                logo_sku = str(int(float(raw_logo)))
            except:
                logo_sku = raw_logo
            
            # Get logo information from database
            logo_info = get_logo_info(logo_sku)
            
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(18.89, 5, "LOGO SKU:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            logo_display = truncate_text(logo_sku, pdf, 15.11 * 0.90)
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

            # Add logo image if available
            add_logo_image_to_pdf(pdf, logo_info)

            pdf.output(os.path.join(OUTPUT_FOLDER, f"ART_INSTRUCTIONS_SO_{doc_num}.pdf"))
            print(f"Generated PDF for Document {doc_num} with logo info: {logo_info is not None}")

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