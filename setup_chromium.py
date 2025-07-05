"""
Script để copy Chromium vào thư mục dự án để đóng gói
"""
import os
import shutil
import sys
from pathlib import Path

# Thiết lập encoding UTF-8 cho console
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def copy_chromium_to_project():
    """
    Copy Chromium từ thư mục cache của Playwright vào thư mục dự án
    """
    # Đường dẫn thư mục dự án
    project_dir = Path(__file__).parent
    chromium_project_dir = project_dir / "chromium-browser"
    
    # Tìm Chromium trong cache của Playwright
    playwright_cache = Path.home() / "AppData" / "Local" / "ms-playwright"
    
    print("🔍 Searching for Chromium in Playwright cache...")
    
    chromium_dirs = list(playwright_cache.glob("chromium-*"))
    if not chromium_dirs:
        print("❌ No Chromium found in Playwright cache!")
        return False
    
    # Lấy version mới nhất
    latest_chromium = sorted(chromium_dirs, key=lambda x: x.name)[-1]
    print(f"📦 Found Chromium: {latest_chromium}")
    
    # Copy vào thư mục dự án
    if chromium_project_dir.exists():
        print("🗑️  Removing old Chromium...")
        shutil.rmtree(chromium_project_dir)
    
    print(f"📋 Copying Chromium to: {chromium_project_dir}")
    shutil.copytree(latest_chromium, chromium_project_dir)
    
    # Tìm file chrome.exe
    chrome_exe = None
    for chrome_path in chromium_project_dir.rglob("chrome.exe"):
        chrome_exe = chrome_path
        break
    
    if chrome_exe:
        print(f"✅ Chromium copied successfully!")
        print(f"🌐 Chrome executable: {chrome_exe}")
        return True
    else:
        print("❌ chrome.exe not found in copied directory!")
        return False

def main():
    print("=" * 60)
    print("🚀 SETUP CHROMIUM FOR PROJECT PACKAGING")
    print("=" * 60)
    
    success = copy_chromium_to_project()
    
    if success:
        print("\n✅ Setup completed! Chromium is now in your project directory.")
        print("📦 You can now package your project with PyInstaller.")
    else:
        print("\n❌ Setup failed! Please check the errors above.")

if __name__ == "__main__":
    main()
