"""
Script để đóng gói dự án Order Checker thành file .exe
"""
import os
import subprocess
import sys
from pathlib import Path

def build_exe():
    """
    Đóng gói dự án thành file .exe
    """
    project_dir = Path(__file__).parent
    
    print("🚀 Building Order Checker executable...")
    
    # Command PyInstaller
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # Tạo 1 file exe duy nhất
        # "--windowed",                   # Không hiện cửa sổ console - tắt để debug
        "--name=OrderChecker",          # Tên file exe
        "--icon=icon.ico",              # Icon (nếu có)
        "--add-data=chromium-browser;chromium-browser",  # Include chromium
        "--add-data=input;input",       # Include input folder
        "--add-data=output;output",     # Include output folder
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
            print(f"📁 Executable created at: {project_dir}/dist/OrderChecker.exe")
            
            # Hiển thị thông tin file
            exe_path = project_dir / "dist" / "OrderChecker.exe"
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"📊 File size: {size_mb:.1f} MB")
        else:
            print("❌ Build failed!")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            
    except Exception as e:
        print(f"❌ Error: {e}")

def build_dir():
    """
    Đóng gói dự án thành thư mục (không nén thành 1 file)
    """
    project_dir = Path(__file__).parent
    
    print("🚀 Building Order Checker directory...")
    
    # Command PyInstaller
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onedir",                     # Tạo thư mục
        "--name=OrderChecker",          # Tên thư mục
        "--add-data=chromium-browser;chromium-browser",  # Include chromium
        "--add-data=input;input",       # Include input folder  
        "--add-data=output;output",     # Include output folder
        "--clean",                      # Clean cache
        "-y",                          # Overwrite without confirmation
        "check_oder.py"                 # File chính
    ]
    
    print(f"📦 Running: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, cwd=project_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ Build successful!")
            print(f"📁 Application created at: {project_dir}/dist/OrderChecker/")
            
        else:
            print("❌ Build failed!")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            
    except Exception as e:
        print(f"❌ Error: {e}")

def main():
    print("=" * 60)
    print("📦 ORDER CHECKER - BUILD EXECUTABLE")
    print("=" * 60)
    
    print("Choose build type:")
    print("1. Single executable file (--onefile)")
    print("2. Directory with files (--onedir)")
    
    choice = input("Enter choice (1 or 2): ").strip()
    
    if choice == "1":
        build_exe()
    elif choice == "2":
        build_dir()
    else:
        print("❌ Invalid choice!")

if __name__ == "__main__":
    main()
