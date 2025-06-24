#!/usr/bin/env python
"""
Art Instructions Generator - Standalone Executable
Main entry point for the .exe version
"""

import os
import sys
import webbrowser
import threading
import time
from pathlib import Path

# Add the current directory to Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

# Import your Flask app
from app import app, load_logo_database

def setup_directories():
    """Create necessary directories for the application"""
    directories = [
        'uploads',
        'outputs', 
        'logo_database',
        'logo_images',
        'templates',
        'static'
    ]
    
    for directory in directories:
        dir_path = os.path.join(current_dir, directory)
        os.makedirs(dir_path, exist_ok=True)
        print(f"✓ Directory verified: {directory}/")

def check_required_files():
    """Check if required files exist"""
    required_files = {
        'logo_database/ArtDBSample.xlsx': 'Logo database file',
        'static/jauniforms.png': 'Company logo image'
    }
    
    missing_files = []
    for file_path, description in required_files.items():
        full_path = os.path.join(current_dir, file_path)
        if not os.path.exists(full_path):
            missing_files.append(f"  ❌ {file_path} - {description}")
        else:
            print(f"  ✓ {file_path} - Found")
    
    return missing_files

def open_browser():
    """Open the web browser after a short delay"""
    time.sleep(3)  # Wait for Flask to start
    webbrowser.open('http://127.0.0.1:5000')

def main():
    """Main function to run the Flask app"""
    print("=" * 70)
    print("           ART INSTRUCTIONS GENERATOR v3.0")
    print("=" * 70)
    print("🚀 Starting application...")
    
    # Setup directories
    print("\n📁 Setting up directories...")
    setup_directories()
    
    # Check required files
    print("\n📋 Checking required files...")
    missing_files = check_required_files()
    
    if missing_files:
        print("\n❌ MISSING REQUIRED FILES:")
        for file_info in missing_files:
            print(file_info)
        print("\n📝 SETUP INSTRUCTIONS:")
        print("1. Create logo_database/ folder if it doesn't exist")
        print("2. Add your ArtDBSample.xlsx file to logo_database/")
        print("3. Create static/ folder if it doesn't exist") 
        print("4. Add jauniforms.png to static/ folder")
        print("5. Add your logo images to logo_images/ folder")
        print("\n⚠️  Press Enter to continue anyway or close this window to exit...")
        input()
    
    # Load logo database
    print("\n💾 Loading logo database...")
    try:
        load_logo_database()
        print("✓ Logo database loaded successfully")
    except Exception as e:
        print(f"⚠️  Warning: Could not load logo database: {e}")
    
    # Start browser in a separate thread
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()
    
    print("\n🌐 Starting web server...")
    print("📱 Web interface will open automatically")
    print("🔗 Manual access: http://127.0.0.1:5000")
    print("\n⚠️  To stop the application:")
    print("   • Close this window, or")
    print("   • Press Ctrl+C in this window")
    print("=" * 70)
    
    try:
        # Run Flask app
        app.run(
            host='127.0.0.1',
            port=5000,
            debug=False,  # Disable debug mode in production
            use_reloader=False,  # Disable reloader for exe
            threaded=True
        )
    except KeyboardInterrupt:
        print("\n\n👋 Application stopped by user")
    except Exception as e:
        print(f"\n❌ Error running application: {e}")
        print("Press Enter to exit...")
        input()

if __name__ == "__main__":
    main()