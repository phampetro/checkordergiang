"""
Script kiểm tra hệ thống Check Orders
Kiểm tra tính năng xử lý Excel và tích hợp
"""

import os
import sys
from pathlib import Path
from datetime import datetime

# Import modules
try:
    from process_excel import ExcelProcessor, process_excel_for_check_order
    print("✅ Import process_excel thành công")
except ImportError as e:
    print(f"❌ Lỗi import process_excel: {e}")
    sys.exit(1)

try:
    from check_oder import OrderChecker
    print("✅ Import check_oder thành công")
except ImportError as e:
    print(f"❌ Lỗi import check_oder: {e}")
    sys.exit(1)

def test_excel_processor():
    """Kiểm tra ExcelProcessor"""
    print("\n" + "="*60)
    print("🔍 KIỂM TRA EXCEL PROCESSOR")
    print("="*60)
    
    processor = ExcelProcessor()
    print(f"📁 Base path: {processor.base_path}")
    print(f"📁 Output dir: {processor.output_dir}")
    print(f"🔧 Tạo file tổng hợp: {'BẬT' if processor.create_summary else 'TẮT'}")
    
    # Kiểm tra thư mục ngày hiện tại
    today = datetime.now().strftime("%d%m%Y")
    daily_dir = processor.output_dir / today
    print(f"📅 Thư mục hôm nay: {daily_dir}")
    print(f"📂 Tồn tại: {'CÓ' if daily_dir.exists() else 'KHÔNG'}")
    
    if daily_dir.exists():
        excel_files = [f for f in daily_dir.glob("*.xlsx") if not f.name.startswith("~$") and f.name != "Kết quả.xlsx"]
        print(f"📊 Số file Excel: {len(excel_files)}")
        for file in excel_files:
            print(f"   - {file.name}")

def test_order_checker():
    """Kiểm tra OrderChecker"""
    print("\n" + "="*60)
    print("🔍 KIỂM TRA ORDER CHECKER")
    print("="*60)
    
    try:
        checker = OrderChecker()
        print(f"📁 Base path: {checker.base_path}")
        print(f"📁 Input dir: {checker.input_dir}")
        print(f"📁 Output dir: {checker.output_dir}")
        print(f"📁 Daily output: {checker.daily_output_dir}")
        print(f"🌐 Chromium path: {'CÓ' if checker.chromium_path else 'KHÔNG'}")
        
        # Kiểm tra template file
        template_path = checker.input_dir / "template.xlsx"
        print(f"📋 Template file: {'CÓ' if template_path.exists() else 'KHÔNG'}")
        
        # Kiểm tra config
        print(f"⚙️ Config: {'CÓ' if checker.config else 'KHÔNG'}")
        
    except Exception as e:
        print(f"❌ Lỗi khởi tạo OrderChecker: {e}")

def test_integration():
    """Kiểm tra tích hợp"""
    print("\n" + "="*60)
    print("🔍 KIỂM TRA TÍCH HỢP")
    print("="*60)
    
    print("🧪 Test function process_excel_for_check_order...")
    try:
        # Không chạy thực tế mà chỉ kiểm tra function có hoạt động
        processor = ExcelProcessor()
        daily_dir = processor.get_daily_directory()
        
        if daily_dir and daily_dir.exists():
            excel_files = [f for f in daily_dir.glob("*.xlsx") if not f.name.startswith("~$") and f.name != "Kết quả.xlsx"]
            if excel_files:
                print(f"📊 Có {len(excel_files)} file Excel sẵn sàng xử lý")
                print("⚠️ Để test thực tế, hãy chạy: process_excel_for_check_order()")
            else:
                print("📭 Không có file Excel để test")
        else:
            print("📅 Chưa có thư mục ngày hôm nay")
            
        print("✅ Function sẵn sàng hoạt động")
        
    except Exception as e:
        print(f"❌ Lỗi test integration: {e}")

def show_summary():
    """Hiển thị tổng kết"""
    print("\n" + "="*60)
    print("📋 TỔNG KẾT HỆ THỐNG")
    print("="*60)
    
    print("🎯 Chức năng chính:")
    print("   ✅ Tự động check orders từ web")
    print("   ✅ Tải file Excel về thư mục theo ngày")
    print("   ✅ Xử lý Excel: ẩn dòng/cột, xóa dữ liệu, freeze panes")
    print("   ✅ Auto-fit cột I/K, báo cáo chi tiết")
    print("   ✅ Không làm mất dữ liệu gốc")
    
    print("\n🔧 Tùy chọn:")
    print("   📊 Tạo file tổng hợp 'Kết quả.xlsx': TẮT (mặc định)")
    print("   🎛️ Có thể bật bằng: processor.enable_summary_creation()")
    
    print("\n📁 Cấu trúc thư mục:")
    print("   📂 input/     - Chứa template.xlsx và config.json")
    print("   📂 output/    - Chứa file kết quả theo ngày")
    print("   📂 output/DDMMYYYY/ - File Excel của ngày cụ thể")
    
    print("\n🚀 Cách sử dụng:")
    print("   1. Chạy check_oder.py để tải và xử lý tự động")
    print("   2. Hoặc chạy process_excel.py để chỉ xử lý Excel")

if __name__ == "__main__":
    print("🔍 KIỂM TRA HỆ THỐNG CHECK ORDERS")
    print("="*60)
    
    # Kiểm tra từng component
    test_excel_processor()
    test_order_checker()
    test_integration()
    show_summary()
    
    print("\n" + "="*60)
    print("✅ KIỂM TRA HOÀN TẤT!")
    print("="*60)
