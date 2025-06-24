# Art Instructions Generator v3.0

## Quick Start Guide

### 1. First Time Setup
1. **Double-click** `Art_Instructions_Generator.exe` to start the application
2. The application will automatically open in your web browser
3. If required files are missing, you'll see instructions on screen

### 2. Required Files Setup

**Logo Database:**
- Place your `ArtDBSample.xlsx` file in the `logo_database/` folder
- This file contains logo information (SKUs, colors, positions, etc.)

**Company Logo:**
- Place your `jauniforms.png` file in the `static/` folder
- This appears on the generated PDF documents

**Logo Images:**
- Place your logo image files in the `logo_images/` folder
- Supported formats: .png, .jpg, .jpeg, .gif, .bmp, .tiff
- Naming convention: `[SKU][suffix].[extension]` (e.g., `0001a.png`, `0001b.png`)

### 3. How to Use
1. **Start Application**: Double-click the .exe file
2. **Upload Data**: Choose your Excel/CSV file with sales order data
3. **Optional Filter**: Enter specific sales order number if needed
4. **Generate PDFs**: Click "Generate Art Instructions PDFs"
5. **Track Progress**: Watch the real-time progress indicator
6. **Download Results**: Get your ZIP file with PDFs and reports

### 4. File Structure
```
Art_Instructions_Package/
├── Art_Instructions_Generator.exe    # Main application
├── logo_database/                    # Place ArtDBSample.xlsx here
│   └── ArtDBSample.xlsx              # Logo database file
├── logo_images/                      # Place logo images here
│   ├── 0001a.png                     # Logo images with SKU naming
│   ├── 0001b.png
│   └── ...
├── static/                           # Static files
│   └── jauniforms.png                # Company logo
└── README.txt                        # This file
```

### 5. Features
- **Real-time Progress Tracking**: See exactly what's happening during processing
- **Comprehensive Reporting**: Get detailed Excel and PDF reports
- **Error Handling**: Clear error messages and validation
- **Sales Order Filtering**: Process specific orders or all orders
- **Intelligent Logo Placement**: Automatic image sizing and layout

### 6. System Requirements
- Windows 10 or later
- At least 4GB RAM
- 500MB free disk space
- Modern web browser (Chrome, Firefox, Edge)

### 7. Troubleshooting

**Application won't start:**
- Make sure you have administrative permissions
- Check Windows Defender/antivirus isn't blocking the file
- Try running as administrator (right-click → "Run as administrator")

**Logo database not loading:**
- Ensure `ArtDBSample.xlsx` is in the `logo_database/` folder
- Check the file isn't corrupted or password-protected
- File must be named exactly `ArtDBSample.xlsx`

**Logo images not showing:**
- Check logo images are in the `logo_images/` folder
- Verify naming convention: `[SKU][suffix].[extension]`
- Supported formats: .png, .jpg, .jpeg, .gif, .bmp, .tiff

**Browser doesn't open automatically:**
- Manually open your browser and go to: http://127.0.0.1:5000
- Make sure no other application is using port 5000

### 8. Data Requirements

**Required Excel/CSV Columns:**
- Document Number (Sales Order Number)
- LOGO (Logo SKU)
- OPERATIONAL CODE

**Optional but Recommended Columns:**
- VENDOR STYLE, COLOR, SIZE, SUBCATEGORY
- Customer/Vendor Name, Due Date, DueDateStatus
- Quantity, List of Operation Codes
- LOGO POSITION, STITCH COUNT, NOTES, FILE NAME

### 9. Processing Rules
- Logo SKU must not be empty, 0000, or 0
- Operational Code must be 11, OR > 89 with valid List of Operation Codes
- Status must not be "Not Approved"
- Logo must exist in database and have corresponding image files

### 10. Version Information
- Version: 3.0
- Features: Progress tracking, enhanced reporting, PDF formatting
- Build Date: 1750710931.5325966

## Support
For technical support, contact your system administrator or IT department.
