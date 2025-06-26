#!/usr/bin/env python
"""
Build script for creating the .exe file
Run this script to build the executable
"""

import os
import subprocess
import sys
import shutil
import time
from pathlib import Path

def install_requirements():
    """Install required packages for building"""
    print("📦 Installing build requirements...")
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'])
        print("✅ Requirements installed!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install requirements: {e}")
        return False
    return True

def safe_remove(path):
    """Safely remove a file or directory with retry logic"""
    if not os.path.exists(path):
        return True
    
    max_retries = 5
    for attempt in range(max_retries):
        try:
            if os.path.isfile(path):
                # Make file writable before deletion
                os.chmod(path, 0o777)
                os.unlink(path)
            else:
                # Make directory and all contents writable
                for root, dirs, files in os.walk(path):
                    for dir in dirs:
                        os.chmod(os.path.join(root, dir), 0o777)
                    for file in files:
                        os.chmod(os.path.join(root, file), 0o777)
                shutil.rmtree(path)
            return True
        except PermissionError as e:
            if attempt < max_retries - 1:
                print(f"  ⚠️  Retry {attempt + 1}/{max_retries}: {e}")
                time.sleep(1)  # Wait a second before retry
            else:
                print(f"  ❌ Failed to remove {path}: {e}")
                return False
        except Exception as e:
            print(f"  ❌ Error removing {path}: {e}")
            return False
    return False

def clean_build():
    """Clean previous build artifacts"""
    folders_to_clean = ['build', 'dist', '__pycache__', 'Art_Instructions_Package']
    
    print("🧹 Cleaning previous build artifacts...")
    for folder in folders_to_clean:
        if os.path.exists(folder):
            if safe_remove(folder):
                print(f"  Removed {folder}/")
            else:
                print(f"  ⚠️  Could not remove {folder}/, continuing anyway...")
    
    # Remove .pyc files
    for pyc_file in Path('.').rglob('*.pyc'):
        try:
            pyc_file.unlink()
        except:
            pass
        
    # Remove __pycache__ directories
    for pycache_dir in Path('.').rglob('__pycache__'):
        if pycache_dir.is_dir():
            safe_remove(str(pycache_dir))

def build_executable():
    """Build the executable using PyInstaller"""
    print("🔨 Building executable...")
    
    try:
        cmd = [
            sys.executable, '-m', 'PyInstaller',
            '--clean',
            'art_instructions.spec'
        ]
        
        subprocess.check_call(cmd)
        print("✅ Build completed successfully!")
        
        # Check if exe was created (now it should be in a folder)
        exe_path = Path('dist/Art_Instructions_Generator/Art_Instructions_Generator.exe')
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"📦 Executable created: {exe_path}")
            print(f"📏 Size: {size_mb:.1f} MB")
            return True
        else:
            print("❌ Executable not found after build")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        return False

def create_distribution_package():
    """Create a distribution package with all necessary files"""
    print("📦 Creating distribution package...")
    
    package_dir = Path('Art_Instructions_Package')
    
    # Clean and create package directory
    if package_dir.exists():
        if not safe_remove(str(package_dir)):
            print("⚠️  Could not remove existing package directory, continuing...")
            return False
    
    package_dir.mkdir()
    
    # Copy the entire dist folder contents (this is now a directory, not a single file)
    dist_source = Path('dist/Art_Instructions_Generator')
    if dist_source.exists():
        # Copy all contents from the dist folder
        for item in dist_source.iterdir():
            if item.is_file():
                shutil.copy2(item, package_dir / item.name)
                print(f"  ✓ Copied {item.name}")
            elif item.is_dir():
                shutil.copytree(item, package_dir / item.name)
                print(f"  ✓ Copied {item.name}/ folder")
    else:
        print("❌ Dist folder not found")
        return False
    
    # Create folder structure for user data
    (package_dir / 'logo_database').mkdir(exist_ok=True)
    (package_dir / 'logo_images').mkdir(exist_ok=True)
    print("  ✓ Created data folders")
    
    # Copy sample files if they exist
    sample_files = [
        ('logo_database/ArtDBSample.xlsx', 'logo_database/'),
        ('static/jauniforms.png', 'static/')
    ]
    
    # Create static folder if it doesn't exist
    (package_dir / 'static').mkdir(exist_ok=True)
    
    for source, dest in sample_files:
        source_path = Path(source)
        if source_path.exists():
            dest_path = package_dir / dest
            dest_path.mkdir(exist_ok=True)
            shutil.copy2(source_path, dest_path / source_path.name)
            print(f"  ✓ Copied {source}")
        else:
            print(f"  ⚠️  {source} not found (optional)")
    
    # Copy some sample logo images if they exist
    logo_images_source = Path('logo_images')
    if logo_images_source.exists():
        sample_count = 0
        for img_file in logo_images_source.glob('*'):
            if img_file.is_file() and sample_count < 5:  # Copy first 5 images as samples
                shutil.copy2(img_file, package_dir / 'logo_images/')
                sample_count += 1
        if sample_count > 0:
            print(f"  ✓ Copied {sample_count} sample logo images")
    
    # Create comprehensive README
    readme_content = """# Art Instructions Generator v3.0

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
├── _internal/                        # Application files (DO NOT MODIFY)
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
- Ensure `ArtDBSample.xlsx` is in the `logo_database/` folder next to the .exe
- Check the file isn't corrupted or password-protected
- File must be named exactly `ArtDBSample.xlsx`

**Logo images not showing:**
- Check logo images are in the `logo_images/` folder next to the .exe
- Verify naming convention: `[SKU][suffix].[extension]`
- Supported formats: .png, .jpg, .jpeg, .gif, .bmp, .tiff

**Browser doesn't open automatically:**
- Manually open your browser and go to: http://127.0.0.1:5000
- Make sure no other application is using port 5000

**Download button not working:**
- Make sure the application has write permissions to its folder
- Try running as administrator if needed
- Check that antivirus isn't blocking file creation

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
- Architecture: Directory-based distribution (not single-file)

## Support
For technical support, contact your system administrator or IT department.

## Important Notes
- Do NOT delete or modify files in the `_internal/` folder
- Keep all folders together - the application needs them to work properly
- The application creates `outputs/` and `uploads/` folders automatically
"""
    
    with open(package_dir / 'README.txt', 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print("  ✓ Created comprehensive README")
    
    print(f"✅ Distribution package created: {package_dir}/")
    print("📋 Package contents:")
    for item in package_dir.rglob('*'):
        if item.is_file():
            size_kb = item.stat().st_size / 1024
            relative_path = item.relative_to(package_dir)
            print(f"   📄 {relative_path} ({size_kb:.1f} KB)")
    
    return True

def main():
    """Main build process"""
    print("=" * 70)
    print("    ART INSTRUCTIONS GENERATOR - BUILD SCRIPT v3.0")
    print("=" * 70)
    
    # Check if we're in the right directory
    required_files = ['app.py', 'report_generator.py', 'requirements.txt']
    missing_files = [f for f in required_files if not Path(f).exists()]
    
    if missing_files:
        print(f"❌ Missing required files: {missing_files}")
        print("Please run this script from the project directory containing your Flask app.")
        return
    
    # Install requirements
    if not install_requirements():
        return
    
    # Clean previous builds
    clean_build()
    
    # Build executable
    if build_executable():
        # Create distribution package
        if create_distribution_package():
            print("\n" + "=" * 70)
            print("✅ BUILD COMPLETED SUCCESSFULLY!")
            print("📦 Your executable is ready in: Art_Instructions_Package/")
            print("📋 See README.txt in the package for setup instructions")
            print("🚀 To test: Go to Art_Instructions_Package/ and run Art_Instructions_Generator.exe")
            print("💡 Note: This is now a DIRECTORY-based distribution (not single file)")
            print("📁 Keep all files and folders together for the app to work properly")
            print("=" * 70)
        else:
            print("\n❌ Failed to create distribution package.")
    else:
        print("\n❌ Build failed. Please check the error messages above.")

if __name__ == "__main__":
    main()