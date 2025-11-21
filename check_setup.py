"""
Pre-flight check script to verify everything is ready.

Run this before processing to make sure all files and dependencies are in place.
"""

import sys
import os
from config import PDF_INPUT_PATH, OLD_EXCEL_PATH, OUTPUT_EXCEL_PATH


def check_dependencies():
    """Check if required Python packages are installed."""
    print("🔍 Checking Python dependencies...")
    
    required = ['pdfplumber', 'pandas', 'openpyxl']
    missing = []
    
    for package in required:
        try:
            __import__(package)
            print(f"   ✓ {package}")
        except ImportError:
            print(f"   ✗ {package} - NOT INSTALLED")
            missing.append(package)
    
    if missing:
        print(f"\n❌ Missing packages: {', '.join(missing)}")
        print("\nRun this command to install:")
        print("   pip install -r requirements.txt")
        return False
    
    print("   ✓ All dependencies installed\n")
    return True


def check_files():
    """Check if required input files exist."""
    print("📁 Checking input files...")
    
    # Use paths from config
    pdf_file = PDF_INPUT_PATH
    excel_file = OLD_EXCEL_PATH
    
    pdf_exists = os.path.exists(pdf_file)
    excel_exists = os.path.exists(excel_file)
    
    if pdf_exists:
        size = os.path.getsize(pdf_file) / 1024 / 1024  # MB
        print(f"   ✓ {pdf_file} ({size:.2f} MB)")
    else:
        print(f"   ✗ {pdf_file} - NOT FOUND")
    
    if excel_exists:
        size = os.path.getsize(excel_file) / 1024  # KB
        print(f"   ✓ {excel_file} ({size:.2f} KB)")
    else:
        print(f"   ⚠️  {excel_file} - NOT FOUND")
        print("      (This is OK for first run - no tracking data will be merged)")
    
    if not pdf_exists:
        print("\n❌ PDF file is required!")
        print("\nPlease add your dental software PDF as:")
        print(f"   {pdf_file}")
        return False
    
    print()
    return True


def check_write_permissions():
    """Check if we can write output files."""
    print("✏️  Checking write permissions...")
    
    try:
        test_file = '.test_write_permissions.tmp'
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        print("   ✓ Can write files\n")
        return True
    except Exception as e:
        print(f"   ✗ Cannot write files: {e}\n")
        return False


def show_next_steps():
    """Show what to do next."""
    print("=" * 70)
    print("✅ READY TO RUN!")
    print("=" * 70)
    print("\nNext step:")
    print("   python process_statements.py")
    print()


def main():
    """Run all checks."""
    print("=" * 70)
    print("DENTAL AUTOMATION - PRE-FLIGHT CHECK")
    print("=" * 70)
    print()
    
    # Run all checks
    deps_ok = check_dependencies()
    files_ok = check_files()
    write_ok = check_write_permissions()
    
    # Summary
    all_ok = deps_ok and files_ok and write_ok
    
    if all_ok:
        show_next_steps()
    else:
        print("=" * 70)
        print("❌ NOT READY - Please fix the issues above")
        print("=" * 70)
        print()
        sys.exit(1)


if __name__ == '__main__':
    main()

