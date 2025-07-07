"""
Script để đóng gói dự án Order Checker v2.0 thành thư mục
Tất cả build đều tạo thư mục để chạy nhanh, không cần giải nén
"""
import os
import subprocess
import sys
from pathlib import Path

def build_menu_exe():
    """
    Đóng gói menu.py thành thư mục (recommended - chạy nhanh, không cần giải nén)
    """
    project_dir = Path(__file__).parent
    
    print("🎛️ Building Order Checker Menu v2.0 directory...")
    
    # Command PyInstaller cho menu - build thành thư mục
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",                     # Tạo thư mục thay vì 1 file
        "--name=OrderChecker-Menu",     # Tên thư mục
        "--icon=icon.ico",              # Icon (nếu có)
        "--add-data=chromium-browser;chromium-browser",  # Include chromium
        "--add-data=input;input",       # Include input folder
        "--add-data=output;output",     # Include output folder
        "--add-data=check_oder.py;.",   # Include check_oder.py
        "--add-data=process_excel.py;.", # Include process_excel.py
        "--add-data=test_system.py;.",  # Include test_system.py
        "--add-data=HUONG_DAN.md;.",    # Include hướng dẫn
        "--hidden-import=process_excel", # Import process_excel
        "--hidden-import=check_oder",   # Import check_oder
        "--clean",                      # Clean cache
        "-y",                          # Overwrite without confirmation
        "menu.py"                      # File menu chính
    ]
    
    # Nếu không có icon thì bỏ qua
    if not (project_dir / "icon.ico").exists():
        cmd = [item for item in cmd if not item.startswith("--icon")]
    
    print(f"📦 Running: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, cwd=project_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ Menu build successful!")
            print(f"📁 Directory created at: {project_dir}/dist/OrderChecker-Menu/")
            print(f"🚀 Run: {project_dir}/dist/OrderChecker-Menu/OrderChecker-Menu.exe")
            
            # Hiển thị thông tin thư mục
            exe_path = project_dir / "dist" / "OrderChecker-Menu" / "OrderChecker-Menu.exe"
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"📊 Main executable size: {size_mb:.1f} MB")
        else:
            print("❌ Build failed!")
            print("Error:", result.stderr)
            
    except Exception as e:
        print(f"❌ Error: {e}")

def build_exe():
    """
    Đóng gói check_oder.py thành thư mục (chạy nhanh, không cần giải nén)
    """
    project_dir = Path(__file__).parent
    
    print("🚀 Building Order Checker v2.0 directory...")
    
    # Command PyInstaller - build thành thư mục
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",                     # Tạo thư mục thay vì 1 file
        "--name=OrderChecker",          # Tên thư mục
        "--icon=icon.ico",              # Icon (nếu có)
        "--add-data=chromium-browser;chromium-browser",  # Include chromium
        "--add-data=input;input",       # Include input folder
        "--add-data=output;output",     # Include output folder
        "--add-data=process_excel.py;.", # Include process_excel.py
        "--add-data=menu.py;.",         # Include menu.py
        "--add-data=test_system.py;.",  # Include test_system.py
        "--add-data=HUONG_DAN.md;.",    # Include hướng dẫn
        "--hidden-import=process_excel", # Import process_excel
        "--clean",                      # Clean cache
        "-y",                          # Overwrite without confirmation
        "check_oder.py"                 # File chính
    ]
    
    # Nếu không có icon thì bỏ qua
    if not (project_dir / "icon.ico").exists():
        cmd = [item for item in cmd if not item.startswith("--icon")]
    
    print(f"📦 Running: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, cwd=project_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ Build successful!")
            print(f"📁 Directory created at: {project_dir}/dist/OrderChecker/")
            print(f"🚀 Run: {project_dir}/dist/OrderChecker/OrderChecker.exe")
            
            # Hiển thị thông tin thư mục
            exe_path = project_dir / "dist" / "OrderChecker" / "OrderChecker.exe"
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"📊 Main executable size: {size_mb:.1f} MB")
        else:
            print("❌ Build failed!")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            
    except Exception as e:
        print(f"❌ Error: {e}")

def build_dir():
    """
    Đóng gói dự án thành thư mục với tối ưu tốc độ và kích thước
    """
    project_dir = Path(__file__).parent
    
    print("� Building Order Checker optimized directory...")
    
    # Command PyInstaller với tối ưu
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",                     # Tạo thư mục
        "--name=OrderChecker-Optimized", # Tên thư mục
        "--add-data=chromium-browser;chromium-browser",  # Include chromium
        "--add-data=input;input",       # Include input folder  
        "--add-data=output;output",     # Include output folder
        "--add-data=process_excel.py;.", # Include process_excel.py
        "--add-data=menu.py;.",         # Include menu.py
        "--add-data=test_system.py;.",  # Include test_system.py
        "--add-data=HUONG_DAN.md;.",    # Include hướng dẫn
        "--hidden-import=process_excel", # Import process_excel
        "--exclude-module=tkinter",     # Loại bỏ tkinter không cần
        "--exclude-module=matplotlib",  # Loại bỏ matplotlib không cần
        "--clean",                      # Clean cache
        "-y",                          # Overwrite without confirmation
        "check_oder.py"                 # File chính
    ]
    
    print(f"📦 Running: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, cwd=project_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ Build successful!")
            print(f"📁 Optimized directory created at: {project_dir}/dist/OrderChecker-Optimized/")
            print(f"🚀 Run: {project_dir}/dist/OrderChecker-Optimized/OrderChecker-Optimized.exe")
            
            # Hiển thị thông tin thư mục
            exe_path = project_dir / "dist" / "OrderChecker-Optimized" / "OrderChecker-Optimized.exe"
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"📊 Main executable size: {size_mb:.1f} MB")
        else:
            print("❌ Build failed!")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            
    except Exception as e:
        print(f"❌ Error: {e}")

def main():
    print("=" * 60)
    print("📦 ORDER CHECKER v2.0 - BUILD DIRECTORIES")
    print("=" * 60)
    
    print("Choose what to build (all create directories for fast startup):")
    print("1. 🎛️ Menu directory (RECOMMENDED - OrderChecker-Menu/)")
    print("2. 🚀 Direct directory (OrderChecker/)")
    print("3. 📁 Optimized directory (OrderChecker-Optimized/)")
    print("0. ❌ Exit")
    
    choice = input("\nEnter choice (1-3): ").strip()
    
    if choice == "1":
        build_menu_exe()
    elif choice == "2":
        build_exe()
    elif choice == "3":
        build_dir()
    elif choice == "0":
        print("👋 Goodbye!")
    else:
        print("❌ Invalid choice!")

if __name__ == "__main__":
    main()
