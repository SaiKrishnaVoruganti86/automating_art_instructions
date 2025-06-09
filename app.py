from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
from fpdf import FPDF
import os
import zipfile
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
ZIP_NAME = "art_instructions_pdfs.zip"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

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


def add_logo_color_table(pdf):
    pdf.ln(5)
    total_width = 190.5 - (2 * 0.8)
    logo_color_width = total_width * 0.20
    number_width = total_width * 0.05
    value_width = total_width * 0.35

    pdf.set_font("Arial", "B", 8.5)
    pdf.cell(logo_color_width, 5, "LOGO COLOR:", border=1, align="C")
    pdf.set_font("Arial", "", 8.5)
    pdf.cell(number_width, 5, "1", border=1, align="C")
    pdf.cell(value_width, 5, "", border=1)
    pdf.cell(number_width, 5, "9", border=1, align="C")
    pdf.cell(value_width, 5, "", border=1)
    pdf.ln()

    pdf.cell(logo_color_width, 5, "", border=1)
    pdf.cell(number_width, 5, "2", border=1, align="C")
    pdf.cell(value_width, 5, "", border=1)
    pdf.cell(number_width, 5, "10", border=1, align="C")
    pdf.cell(value_width, 5, "", border=1)
    pdf.ln()

    pdf.set_font("Arial", "B", 8.5)
    pdf.cell(logo_color_width, 5, "PRODUCTION DAY:", border=1, align="C")
    pdf.set_font("Arial", "", 8.5)
    pdf.cell(number_width, 5, "3", border=1, align="C")
    pdf.cell(value_width, 5, "", border=1)
    pdf.cell(number_width, 5, "11", border=1, align="C")
    pdf.cell(value_width, 5, "", border=1)
    pdf.ln()

    for i in range(4, 9):
        pdf.cell(logo_color_width, 5, "", border=1)
        pdf.cell(number_width, 5, str(i), border=1, align="C")
        pdf.cell(value_width, 5, "", border=1)
        pdf.cell(number_width, 5, str(i + 8), border=1, align="C")
        pdf.cell(value_width, 5, "", border=1)
        pdf.ln()

@app.route("/", methods=["GET", "POST"])
def upload_file():
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

        for f in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, f))

        for doc_num, group in grouped:
            pdf = FPDF(orientation="P", unit="mm", format=(190.5, 254.0))  # 7.5x10 in
            pdf.set_margins(0.8, 0.8, 0.8)
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=0.8)

            pdf.set_font("Arial", "", 8.5)

            client_name = truncate_text(safe_get(group["Customer/Vendor Name"].iloc[0]), pdf, 72)
            due_date = str(group["Due Date"].iloc[0]).split(" ")[0]

            
            # Width setup
            full_width = 190
            usable_width = full_width - (2 * 0.8)  # 190.5 - 1.6 = 188.9mm
            left_width = full_width * 0.75
            right_width = full_width - left_width

            # Save starting position
            x_left = pdf.get_x()
            y_top = pdf.get_y()

            # Draw ART INSTRUCTIONS title across the top-left box
            pdf.set_font("Arial", "B", 10)
            pdf.cell(left_width, 8, "ART INSTRUCTIONS", border=1, align="C")

            # Draw right-side logo container
            pdf.cell(right_width, 8, "", border=0)
            pdf.image("static/jauniforms.png", x=pdf.get_x() - right_width + 3, y=pdf.get_y() + 1, w=right_width - 6)

            pdf.ln()

            # CLIENT | VALUE (spans full left_width), no DATE
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(20, 6, "CLIENT:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            pdf.cell(left_width - 20, 6, client_name, border=1)
            pdf.cell(right_width, 6, "", border=0)  # logo spacer
            pdf.ln()

            # Row: SO#: value (65%) | DATE: value (35%) of left_width
            so_section_width = left_width * 0.70
            date_section_width = left_width * 0.30

            # SO# section
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(20, 6, "SO#:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            pdf.cell(so_section_width - 20, 6, str(doc_num), border=1)

            # DATE section
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(15, 6, "DATE:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            pdf.cell(date_section_width - 15, 6, due_date, border=1, align="C")

            # Spacer for logo column
            pdf.cell(right_width, 6, "", border=0)
            pdf.ln(8)




            vendor_styles = ", ".join(group["VENDOR STYLE"].dropna().astype(str).unique())
            render_items_section(pdf, vendor_styles, usable_width)


            pdf.ln(2)

            usable_width = 190.5 - (2 * 0.8)  # total page width - left & right margins
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

            pdf.cell(30, 5, "LOGO POSITION:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            logo_pos = safe_get(group["LOGO POSITION"].iloc[0]) if "LOGO POSITION" in group.columns else ""
            pdf.cell(60, 5, logo_pos, border=1)
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(20, 5, "NOTES:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            notes = safe_get(group["NOTES"].iloc[0]) if "NOTES" in group.columns else ""
            pdf.cell(60, 5, notes, border=1)
            pdf.ln(2)

            add_logo_color_table(pdf)

            pdf.ln(2)
            pdf.set_font("Arial", "B", 8.5)
            pdf.cell(30, 5, "LOGO SKU:", border=1, align="C")
            pdf.set_font("Arial", "", 8.5)
            raw_logo = safe_get(group["LOGO"].iloc[0]) if "LOGO" in group.columns else ""
            try:
                logo = str(int(float(raw_logo)))
            except:
                logo = raw_logo
            logo = truncate_text(logo, pdf, (190 - 30) * 0.98)
            pdf.cell(150, 5, logo, border=1)
            pdf.ln(8)

            pdf.output(os.path.join(OUTPUT_FOLDER, f"ART_INSTRUCTIONS_SO_{doc_num}.pdf"))

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
    app.run(debug=True)