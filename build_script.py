# ============================================================================
# build.py
# ============================================================================
"""
PyInstaller build script for 5G Stability Report Generator.
Creates a standalone executable with embedded icon.

Usage:
    python build.py

Requirements:
    pip install pyinstaller pillow
"""

import os
import sys
import base64
import subprocess
from pathlib import Path

# ============================================================================
# APPLICATION METADATA
# ============================================================================
APP_NAME = "5G_Stability"
APP_VERSION = "1.0.0"
APP_AUTHOR = "debu69er"
APP_DESCRIPTION = "5G Performance Stability Report Generator"

# ============================================================================
# ICON GENERATION (BASE64)
# ============================================================================
# This is a sample base64-encoded ICO file (16x16 simple icon)
# Replace this with your actual icon base64 data
ICON_BASE64 = """

"""

# Alternative: Path to existing ICO file
ICON_PATH_OPTION = "assets/app_icon.ico"  # Set to None to use base64


# ============================================================================
# ICON FILE CREATION
# ============================================================================
def create_icon_from_base64(base64_data: str, output_path: str = "app_icon.ico") -> str:
    """
    Create ICO file from base64 encoded data.
    
    Args:
        base64_data: Base64 encoded ICO file data
        output_path: Path where to save the ICO file
        
    Returns:
        Path to created ICO file
    """
    print(f"📦 Creating icon file from base64 data...")
    
    # Decode base64 to binary
    icon_data = base64.b64decode(base64_data.strip())
    
    # Write to file
    with open(output_path, 'wb') as f:
        f.write(icon_data)
    
    print(f"✓ Icon created: {output_path}")
    return output_path


def create_default_icon_with_pillow(output_path: str = "app_icon.ico") -> str:
    """
    Create a simple default icon using Pillow if no icon is provided.
    
    Args:
        output_path: Path where to save the ICO file
        
    Returns:
        Path to created ICO file
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
        
        print(f"🎨 Creating default icon with Pillow...")
        
        # Create a 256x256 image with gradient background
        size = 256
        img = Image.new('RGBA', (size, size), (33, 150, 243, 255))  # Blue background
        draw = ImageDraw.Draw(img)
        
        # Draw a simple 5G symbol
        # Draw circle
        circle_bbox = [size//4, size//4, 3*size//4, 3*size//4]
        draw.ellipse(circle_bbox, fill=(255, 255, 255, 255), outline=(255, 255, 255, 255))
        
        # Draw "5G" text
        try:
            # Try to use a system font
            font = ImageFont.truetype("arial.ttf", 80)
        except:
            font = ImageFont.load_default()
        
        text = "5G"
        # Get text bounding box
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        # Center text
        x = (size - text_width) // 2
        y = (size - text_height) // 2 - 10
        
        draw.text((x, y), text, fill=(33, 150, 243, 255), font=font)
        
        # Save as ICO with multiple sizes
        img.save(output_path, format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
        
        print(f"✓ Default icon created: {output_path}")
        return output_path
        
    except ImportError:
        print("⚠ Pillow not installed. Using base64 icon instead.")
        return create_icon_from_base64(ICON_BASE64, output_path)


def get_icon_path() -> str:
    """
    Determine which icon to use: existing file, base64, or generate default.
    
    Returns:
        Path to icon file
    """
    # Option 1: Use existing ICO file if path is specified and exists
    if ICON_PATH_OPTION and os.path.exists(ICON_PATH_OPTION):
        print(f"✓ Using existing icon: {ICON_PATH_OPTION}")
        return ICON_PATH_OPTION
    
    # Option 2: Create from base64
    if ICON_BASE64:
        return create_icon_from_base64(ICON_BASE64)
    
    # Option 3: Create default icon with Pillow
    return create_default_icon_with_pillow()


# ============================================================================
# PYINSTALLER BUILD CONFIGURATION
# ============================================================================
def build_executable():
    """
    Build standalone executable using PyInstaller.
    """
    print("="*70)
    print(f"🏗️  Building {APP_NAME} v{APP_VERSION}")
    print("="*70)
    
    # Get icon path
    icon_path = get_icon_path()
    
    # Ensure main.py exists
    if not os.path.exists("main.py"):
        print("❌ Error: main.py not found!")
        sys.exit(1)
    
    # PyInstaller command arguments
    pyinstaller_args = [
        'pyinstaller',
        '--name', APP_NAME,                          # Application name
        '--onefile',                                 # Create single executable
        '--windowed',                                # No console window (GUI app)
        '--icon', icon_path,                         # Application icon
        '--clean',                                   # Clean cache before build
        '--noconfirm',                               # Overwrite without asking
        
        # Add metadata (Windows only)
        '--version-file', 'version_info.txt' if os.path.exists('version_info.txt') else '',
        
        # Hidden imports (add modules that PyInstaller might miss)
        '--hidden-import', 'pandas',
        '--hidden-import', 'openpyxl',
        '--hidden-import', 'PyQt6',
        '--hidden-import', 'PyQt6.QtCore',
        '--hidden-import', 'PyQt6.QtGui',
        '--hidden-import', 'PyQt6.QtWidgets',
        
        # Add data files (if any)
        # '--add-data', 'config.ini:.',              # Example: include config file
        
        # Optimization
        '--optimize', '2',                           # Optimization level
        
        # Entry point
        'main.py'
    ]
    
    # Remove empty arguments
    pyinstaller_args = [arg for arg in pyinstaller_args if arg]
    
    print("\n📋 PyInstaller Configuration:")
    print(f"   Name: {APP_NAME}")
    print(f"   Version: {APP_VERSION}")
    print(f"   Icon: {icon_path}")
    print(f"   Mode: Single File (--onefile)")
    print(f"   GUI: Yes (--windowed)")
    
    # Run PyInstaller
    print("\n🔨 Running PyInstaller...")
    print("   This may take a few minutes...\n")
    
    try:
        result = subprocess.run(pyinstaller_args, check=True)
        
        print("\n" + "="*70)
        print("✅ BUILD SUCCESSFUL!")
        print("="*70)
        print(f"\n📦 Executable created:")
        
        # Determine executable path
        if sys.platform == "win32":
            exe_path = Path("dist") / f"{APP_NAME}.exe"
        else:
            exe_path = Path("dist") / APP_NAME
        
        if exe_path.exists():
            print(f"   {exe_path.absolute()}")
            print(f"\n💾 File size: {exe_path.stat().st_size / (1024*1024):.2f} MB")
        
        print("\n📝 Build artifacts:")
        print(f"   - Executable: dist/{APP_NAME}.exe")
        print(f"   - Build files: build/")
        print(f"   - Spec file: {APP_NAME}.spec")
        
        print("\n🚀 You can now distribute the executable to users!")
        print("   No Python installation required on target machines.\n")
        
    except subprocess.CalledProcessError as e:
        print("\n❌ BUILD FAILED!")
        print(f"   Error: {e}")
        sys.exit(1)
    except FileNotFoundError:
        print("\n❌ PyInstaller not found!")
        print("   Please install it with: pip install pyinstaller")
        sys.exit(1)


# ============================================================================
# VERSION INFO FILE GENERATOR (WINDOWS)
# ============================================================================
def create_version_info():
    """
    Create version_info.txt for Windows executable metadata.
    """
    version_info_content = f"""# UTF-8
#
# Version Info for {APP_NAME}
#

VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'{APP_AUTHOR}'),
        StringStruct(u'FileDescription', u'{APP_DESCRIPTION}'),
        StringStruct(u'FileVersion', u'{APP_VERSION}'),
        StringStruct(u'InternalName', u'{APP_NAME}'),
        StringStruct(u'LegalCopyright', u'Copyright (c) 2026'),
        StringStruct(u'OriginalFilename', u'{APP_NAME}.exe'),
        StringStruct(u'ProductName', u'{APP_NAME}'),
        StringStruct(u'ProductVersion', u'{APP_VERSION}')])
      ]
    ),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
"""
    
    with open('version_info.txt', 'w', encoding='utf-8') as f:
        f.write(version_info_content)
    
    print("✓ Created version_info.txt")


# ============================================================================
# CLEANUP FUNCTION
# ============================================================================
def cleanup_build_files():
    """
    Clean up temporary build files (optional).
    """
    import shutil
    
    print("\n🧹 Cleaning up build artifacts...")
    
    # Remove build directories
    dirs_to_remove = ['build', '__pycache__']
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"   Removed: {dir_name}/")
    
    # Remove spec file (optional - comment out if you want to keep it)
    # spec_file = f"{APP_NAME}.spec"
    # if os.path.exists(spec_file):
    #     os.remove(spec_file)
    #     print(f"   Removed: {spec_file}")
    
    # Remove generated icon if it was temporary
    if os.path.exists("app_icon.ico") and not ICON_PATH_OPTION:
        os.remove("app_icon.ico")
        print(f"   Removed: app_icon.ico")
    
    print("✓ Cleanup complete")


# ============================================================================
# MAIN EXECUTION
# ============================================================================
def main():
    """
    Main build script execution.
    """
    print("""
    ╔═══════════════════════════════════════════════════════════════╗
    ║                                                               ║
    ║        5G STABILITY REPORTER - BUILD SCRIPT                   ║
    ║                                                               ║
    ║        Creating standalone executable with PyInstaller        ║
    ║                                                               ║
    ╚═══════════════════════════════════════════════════════════════╝
    """)
    
    # Step 1: Create version info file (Windows)
    if sys.platform == "win32":
        create_version_info()
    
    # Step 2: Build executable
    build_executable()
    
    # Step 3: Optional cleanup
    # Uncomment the line below if you want to auto-cleanup
    # cleanup_build_files()
    
    print("\n✨ Build process complete!\n")


if __name__ == "__main__":
    main()


# ============================================================================
# ADVANCED CONFIGURATION EXAMPLES
# ============================================================================
"""
ADVANCED USAGE EXAMPLES:

1. Custom Icon from File:
   ICON_PATH_OPTION = "path/to/your/icon.ico"

2. Custom Icon from Base64:
   - Convert your ICO file to base64:
     
     import base64
     with open("your_icon.ico", "rb") as f:
         icon_base64 = base64.b64encode(f.read()).decode()
         print(icon_base64)
   
   - Copy the output and paste into ICON_BASE64 variable

3. Include Additional Files:
   Add to pyinstaller_args:
   '--add-data', 'templates:templates',
   '--add-data', 'config.ini:.',

4. Multi-file Build (not --onefile):
   Change '--onefile' to '--onedir'
   Results in a folder with multiple files instead of single exe

5. Console Mode (for debugging):
   Remove '--windowed' to see console output

6. Splash Screen:
   '--splash', 'splash.png',

7. Additional Hidden Imports:
   '--hidden-import', 'module_name',

8. UPX Compression (smaller file size):
   '--upx-dir', 'path/to/upx',

BUILD COMMAND ALTERNATIVES:

# Quick build (from command line):
pyinstaller --onefile --windowed --icon=app_icon.ico --name=5G_Stability_Reporter main.py

# With all optimizations:
pyinstaller --onefile --windowed --icon=app_icon.ico --name=5G_Stability_Reporter --optimize=2 --clean main.py

DISTRIBUTION NOTES:

1. The executable is self-contained - no Python needed on target machine
2. First run may be slow (extracting files), subsequent runs are faster
3. Antivirus may flag it (false positive) - you may need code signing certificate
4. Test on clean Windows machine before distributing
5. Consider creating installer with tools like Inno Setup or NSIS

"""
