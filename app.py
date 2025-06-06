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

def render_items_section(pdf, vendor_styles):
    styles = vendor_styles.split(", ")
    label_width = 40
    value_width = 150
    max_width = value_width - 5

    pdf.set_font("Arial", "", 10)
    line = ""
    for style in styles:
        appended = style + ", "
        if pdf.get_string_width(line + appended) < max_width:
            line += appended
        else:
            pdf.set_font("Arial", "B", 10)
            pdf.cell(label_width, 8, "ITEMS:", border=1, align="C")
            pdf.set_font("Arial", "", 10)
            pdf.cell(value_width, 8, line.strip(", "), border=1)
            pdf.ln()
            line = appended

    if line:
        pdf.set_font("Arial", "B", 10)
        pdf.cell(label_width, 8, "ITEMS:", border=1, align="C")
        pdf.set_font("Arial", "", 10)
        pdf.cell(value_width, 8, line.strip(", "), border=1)
        pdf.ln()

def add_logo_color_table(pdf):
    pdf.ln(10)
    total_width = 190
    logo_color_width = total_width * 0.20
    number_width = total_width * 0.05
    value_width = total_width * 0.35

    pdf.set_font("Arial", "B", 10)
    pdf.cell(logo_color_width, 8, "LOGO COLOR:", border=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(number_width, 8, "1", border=1, align="C")
    pdf.cell(value_width, 8, "", border=1)
    pdf.cell(number_width, 8, "9", border=1, align="C")
    pdf.cell(value_width, 8, "", border=1)
    pdf.ln()

    pdf.cell(logo_color_width, 8, "", border=1)
    pdf.cell(number_width, 8, "2", border=1, align="C")
    pdf.cell(value_width, 8, "", border=1)
    pdf.cell(number_width, 8, "10", border=1, align="C")
    pdf.cell(value_width, 8, "", border=1)
    pdf.ln()

    pdf.set_font("Arial", "B", 10)
    pdf.cell(logo_color_width, 8, "PRODUCTION DAY:", border=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(number_width, 8, "3", border=1, align="C")
    pdf.cell(value_width, 8, "", border=1)
    pdf.cell(number_width, 8, "11", border=1, align="C")
    pdf.cell(value_width, 8, "", border=1)
    pdf.ln()

    for i in range(4, 9):
        pdf.cell(logo_color_width, 8, "", border=1)
        pdf.cell(number_width, 8, str(i), border=1, align="C")
        pdf.cell(value_width, 8, "", border=1)
        pdf.cell(number_width, 8, str(i + 8), border=1, align="C")
        pdf.cell(value_width, 8, "", border=1)
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
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)

            pdf.set_font("Arial", "", 10)
            client_name = truncate_text(safe_get(group["Customer/Vendor Name"].iloc[0]), pdf, 72)
            due_date = str(group["Due Date"].iloc[0]).split(" ")[0]

            pdf.image("static/jauniforms.png", x=158, y=12, w=35)

            x_left = 10
            w_total = 190 * 0.75
            y_start = 10

            pdf.set_font("Arial", "B", 12)
            pdf.set_xy(x_left, y_start)
            pdf.cell(w_total, 10, "ART INSTRUCTIONS", border=1, align="C")

            pdf.set_xy(x_left, y_start + 10)
            pdf.set_font("Arial", "B", 10)
            pdf.cell(25, 8, "CLIENT:", border=1, align="C")
            pdf.set_font("Arial", "", 10)
            pdf.cell(80, 8, client_name, border=1)
            pdf.set_font("Arial", "B", 10)
            pdf.cell(w_total - 25 - 80, 8, "DATE:", border=1, align="C")

            pdf.set_xy(x_left, y_start + 18)
            pdf.set_font("Arial", "B", 10)
            pdf.cell(25, 8, "SO#:", border=1, align="C")
            pdf.set_font("Arial", "", 10)
            pdf.cell(80, 8, str(doc_num), border=1)
            pdf.set_font("Arial", "", 10)
            pdf.cell(w_total - 25 - 80, 8, due_date, border=1, align="C")

            pdf.ln(10)

            vendor_styles = ", ".join(group["VENDOR STYLE"].dropna().astype(str).unique())
            render_items_section(pdf, vendor_styles)

            pdf.ln(2.5)

            COLOR_WIDTH = 104.5
            DESC_WIDTH = 57.0
            QTY_WIDTH = 28.5

            pdf.set_font("Arial", "B", 10)
            pdf.cell(COLOR_WIDTH, 8, "COLOR", 1, align="C")
            pdf.cell(DESC_WIDTH, 8, "DESCRIPTION", 1, align="C")
            pdf.cell(QTY_WIDTH, 8, "QTY", 1, align="C")
            pdf.ln()

            pdf.set_font("Arial", "", 10)
            total_qty = 0
            for _, row in group.iterrows():
                color_text = truncate_text(safe_get(row.get("COLOR")), pdf, COLOR_WIDTH * 0.90)
                description_text = safe_get(row.get("SUBCATEGORY"))
                try:
                    qty = float(row.get("Quantity"))
                except (ValueError, TypeError):
                    qty = 0
                total_qty += qty
                qty_text = str(int(qty)) if pd.notna(qty) else ""

                pdf.cell(COLOR_WIDTH, 8, color_text, 1, align="C")
                pdf.cell(DESC_WIDTH, 8, description_text, 1, align="C")
                pdf.cell(QTY_WIDTH, 8, qty_text, 1, align="C")
                pdf.ln()

            pdf.set_font("Arial", "B", 10)
            pdf.cell(COLOR_WIDTH, 8, "", 1)
            pdf.cell(DESC_WIDTH, 8, "TOTAL:", 1, align="C")
            pdf.cell(QTY_WIDTH, 8, str(int(total_qty)), 1, align="C")
            pdf.ln(11)

            label1_width = 41.8
            value1_width = 72.2
            label2_width = 19.0
            value2_width = 57.0

            pdf.set_font("Arial", "B", 10)
            pdf.cell(label1_width, 8, "LOGO POSITION:", border=1, align="C")
            pdf.set_font("Arial", "", 10)
            logo_position_value = safe_get(group["LOGO POSITION"].iloc[0]) if "LOGO POSITION" in group else ""
            pdf.cell(value1_width, 8, logo_position_value, border=1)
            pdf.set_font("Arial", "B", 10)
            pdf.cell(label2_width, 8, "NOTES:", border=1, align="C")
            pdf.set_font("Arial", "", 10)
            notes_value = safe_get(group["NOTES"].iloc[0]) if "NOTES" in group else ""
            pdf.cell(value2_width, 8, notes_value, border=1)
            pdf.ln(0.25)

            add_logo_color_table(pdf)

            pdf.ln(2.5)
            pdf.set_font("Arial", "B", 10)
            pdf.cell(30, 8, "LOGO SKU:", border=1, align="C")
            pdf.set_font("Arial", "", 10)
            raw_logo = safe_get(group["LOGO"].iloc[0])
            try:
                logo_value = str(int(float(raw_logo)))
            except (ValueError, TypeError):
                logo_value = raw_logo
            logo_value = truncate_text(logo_value, pdf, (190 - 30) * 0.98)
            pdf.cell(160, 8, logo_value, border=1)
            pdf.ln(10)

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
