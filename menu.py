"""
Script tiện ích quản lý hệ thống Check Orders
Cung cấp menu lựa chọn các chức năng chính
"""

import os
import sys
from pathlib import Path
from datetime import datetime

def show_menu():
    """Hiển thị menu chính"""
    print("\n" + "="*60)
    print("🎯 HỆ THỐNG CHECK ORDERS - MENU CHÍNH")
    print("="*60)
    print("1. 🚀 Chạy hệ thống hoàn chỉnh (check orders + xử lý Excel)")
    print("2. 📊 Chỉ xử lý file Excel")
    print("3. 🔍 Kiểm tra hệ thống")
    print("4. 📋 Bật/tắt tạo file tổng hợp")
    print("5. 📁 Mở thư mục kết quả")
    print("6. 📖 Xem hướng dẫn")
    print("7. 🔧 Cài đặt/kiểm tra môi trường")
    print("0. ❌ Thoát")
    print("="*60)

def get_script_path(script_name):
    """Lấy đường dẫn đúng của script khi chạy từ package hoặc dev"""
    import sys
    from pathlib import Path
    
    if getattr(sys, 'frozen', False):
        # Chạy từ package (exe)
        if hasattr(sys, '_MEIPASS'):
            # PyInstaller bundle
            base_path = Path(sys._MEIPASS)
        else:
            # Fallback
            base_path = Path(sys.executable).parent
    else:
        # Dev mode
        base_path = Path(__file__).parent
    
    script_path = base_path / script_name
    if script_path.exists():
        return str(script_path)
    else:
        # Fallback - chỉ dùng tên file
        return script_name

def run_full_system():
    """Chạy hệ thống hoàn chỉnh"""
    print("\n🚀 Đang khởi chạy hệ thống hoàn chỉnh...")
    
    try:
        # Kiểm tra xem có đang chạy trong package không
        import sys
        if getattr(sys, 'frozen', False):
            # Đang chạy trong package - import và chạy trực tiếp
            from check_oder import main as check_order_main
            check_order_main()
        else:
            # Đang chạy development mode
            script_path = get_script_path("check_oder.py")
            os.system(f'python "{script_path}"')
    except Exception as e:
        print(f"❌ Lỗi chạy hệ thống: {e}")

def run_excel_only():
    """Chỉ xử lý Excel"""
    print("\n📊 Đang xử lý file Excel...")
    
    try:
        # Kiểm tra xem có đang chạy trong package không
        import sys
        if getattr(sys, 'frozen', False):
            # Đang chạy trong package - import và chạy trực tiếp
            from process_excel import main as process_excel_main
            process_excel_main()
        else:
            # Đang chạy development mode
            script_path = get_script_path("process_excel.py")
            os.system(f'python "{script_path}"')
    except Exception as e:
        print(f"❌ Lỗi xử lý Excel: {e}")

def check_system():
    """Kiểm tra hệ thống"""
    print("\n🔍 Đang kiểm tra hệ thống...")
    
    try:
        # Kiểm tra xem có đang chạy trong package không
        import sys
        if getattr(sys, 'frozen', False):
            # Đang chạy trong package - import và chạy trực tiếp
            from test_system import main as test_system_main
            test_system_main()
        else:
            # Đang chạy development mode
            script_path = get_script_path("test_system.py")
            os.system(f'python "{script_path}"')
    except Exception as e:
        print(f"❌ Lỗi kiểm tra hệ thống: {e}")

def toggle_summary():
    """Bật/tắt tạo file tổng hợp"""
    print("\n📋 QUẢN LÝ FILE TỔNG HỢP")
    print("-"*40)
    print("1. Bật tạo file tổng hợp")
    print("2. Tắt tạo file tổng hợp")
    print("3. Kiểm tra trạng thái hiện tại")
    print("0. Quay lại")
    
    choice = input("\nChọn: ").strip()
    
    if choice == "1":
        print("\n🔧 Tạo script bật file tổng hợp...")
        script_content = """
from process_excel import ExcelProcessor

processor = ExcelProcessor()
processor.enable_summary_creation()
print("✅ Đã BẬT tạo file tổng hợp!")
success = processor.process_excel_files()
if success:
    print("✅ Xử lý hoàn thành!")
else:
    print("❌ Có lỗi xảy ra!")
"""
        with open("run_with_summary.py", "w", encoding="utf-8") as f:
            f.write(script_content)
        print("📝 Đã tạo file run_with_summary.py")
        print("🚀 Chạy: python run_with_summary.py")
        
    elif choice == "2":
        print("✅ File tổng hợp đã TẮT mặc định!")
        print("📊 Chạy bình thường: python process_excel.py")
        
    elif choice == "3":
        try:
            from process_excel import ExcelProcessor
            processor = ExcelProcessor()
            status = "BẬT" if processor.create_summary else "TẮT"
            print(f"📋 Trạng thái hiện tại: {status}")
        except Exception as e:
            print(f"❌ Lỗi kiểm tra: {e}")

def open_output_folder():
    """Mở thư mục kết quả"""
    import sys
    
    # Xác định base path đúng khi chạy trong package
    if getattr(sys, 'frozen', False):
        # Chạy từ package
        base_path = Path(sys.executable).parent
    else:
        # Dev mode
        base_path = Path(__file__).parent
    
    output_dir = base_path / "output"
    today = datetime.now().strftime("%d%m%Y")
    daily_dir = output_dir / today
    
    if daily_dir.exists():
        print(f"📁 Mở thư mục: {daily_dir}")
        os.startfile(str(daily_dir))
    elif output_dir.exists():
        print(f"📁 Mở thư mục: {output_dir}")
        os.startfile(str(output_dir))
    else:
        print("❌ Không tìm thấy thư mục output!")

def show_guide():
    """Hiển thị hướng dẫn"""
    import sys
    
    # Xác định base path đúng khi chạy trong package  
    if getattr(sys, 'frozen', False):
        # Chạy từ package
        base_path = Path(sys.executable).parent
    else:
        # Dev mode
        base_path = Path(__file__).parent
        
    guide_file = base_path / "HUONG_DAN.md"
    if guide_file.exists():
        print(f"📖 Mở file hướng dẫn: {guide_file}")
        os.startfile(str(guide_file))
    else:
        print("❌ Không tìm thấy file hướng dẫn!")

def setup_environment():
    """Cài đặt môi trường"""
    print("\n🔧 KIỂM TRA/CÀI ĐẶT MÔI TRƯỜNG")
    print("-"*40)
    
    # Kiểm tra Python packages
    try:
        import openpyxl
        print("✅ openpyxl: OK")
    except ImportError:
        print("❌ openpyxl: CHƯA CÀI")
        print("🔧 Cài đặt: pip install openpyxl")
    
    try:
        import playwright
        print("✅ playwright: OK")
    except ImportError:
        print("❌ playwright: CHƯA CÀI")
        print("🔧 Cài đặt: pip install playwright")
    
    # Kiểm tra browser
    print("\n🌐 Kiểm tra browser...")
    import sys
    
    # Xác định base path đúng khi chạy trong package
    if getattr(sys, 'frozen', False):
        # Chạy từ package
        base_path = Path(sys.executable).parent
    else:
        # Dev mode
        base_path = Path(__file__).parent
    
    browser_dirs = [
        base_path / "chromium-browser",
        base_path / "chromium"
    ]
    
    browser_found = False
    for browser_dir in browser_dirs:
        if browser_dir.exists():
            chrome_files = list(browser_dir.rglob("chrome.exe"))
            if chrome_files:
                print(f"✅ Chromium: {chrome_files[0]}")
                browser_found = True
                break
    
    if not browser_found:
        print("❌ Chromium: CHƯA CÀI")
        print("🔧 Cài đặt: python -m playwright install chromium")
    
    # Kiểm tra thư mục
    print("\n📁 Kiểm tra thư mục...")
    dirs = ["input", "output"]
    for dir_name in dirs:
        dir_path = base_path / dir_name
        if dir_path.exists():
            print(f"✅ {dir_name}/: OK")
        else:
            print(f"❌ {dir_name}/: KHÔNG TỒN TẠI")
            print(f"🔧 Tạo thư mục: mkdir {dir_name}")

def main():
    """Hàm chính"""
    while True:
        show_menu()
        choice = input("\nChọn chức năng: ").strip()
        
        if choice == "1":
            run_full_system()
        elif choice == "2":
            run_excel_only()
        elif choice == "3":
            check_system()
        elif choice == "4":
            toggle_summary()
        elif choice == "5":
            open_output_folder()
        elif choice == "6":
            show_guide()
        elif choice == "7":
            setup_environment()
        elif choice == "0":
            print("\n👋 Tạm biệt!")
            break
        else:
            print("\n❌ Lựa chọn không hợp lệ!")
        
        input("\n🔄 Nhấn Enter để tiếp tục...")

if __name__ == "__main__":
    main()
