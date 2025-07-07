"""
HƯỚNG DẪN SỬ DỤNG HỆ THỐNG CHECK ORDERS
=====================================

🎯 MÔ TẢ HỆ THỐNG:
- Tự động check orders từ web và tải file Excel về
- Xử lý file Excel theo 10 bước nghiệp vụ (ẩn dòng/cột, xóa dữ liệu, freeze panes, auto-fit)
- Không làm mất dữ liệu gốc, dễ mở rộng

📁 CẤU TRÚC THỦ MỤC:
- input/template.xlsx - Danh sách báo cáo cần check
- input/config.json - Cấu hình hệ thống
- output/DDMMYYYY/ - File Excel kết quả theo ngày
- myenv/ - Python virtual environment

🚀 CÁCH SỬ DỤNG:

1. CHẠY HỆ THỐNG HOÀN CHỈNH:
   python check_oder.py
   
   → Tự động check orders và xử lý Excel

2. CHỈ XỬ LÝ FILE EXCEL:
   python process_excel.py
   
   → Chỉ xử lý file Excel đã có trong thư mục ngày hiện tại

3. KIỂM TRA HỆ THỐNG:
   python test_system.py
   
   → Kiểm tra các component và tình trạng hệ thống

⚙️ CẤU HÌNH:

1. TEMPLATE (input/template.xlsx):
   - Cột A: Tên viết tắt (VD: DHTC)
   - Cột B: Tên báo cáo đầy đủ (VD: DHTC - Đơn hàng thành công)

2. CONFIG (input/config.json):
   - URL website
   - Selector các element
   - Thời gian chờ

📊 XỬ LÝ EXCEL (10 BƯỚC):

B1: Ẩn từ dòng 1 đến dòng 3
B2: Ẩn dòng có cột A rỗng
B3: Ẩn dòng có cột B rỗng  
B4: Ẩn dòng có cột D rỗng AND cột C <> ""
B4: Xóa dữ liệu của các dòng có cột C rỗng, xóa từ K trở đi
B5: Ẩn các dòng K có chứa nội dung "NPP bán"
B6: Ẩn dòng có cột Q > 0 (giữ lại dòng rỗng và 0)
B7: Kiểm tra cột Q nếu có 2 dòng rỗng liên tiếp thì ẩn dòng thứ 2
B8: Ẩn cột S trở đi, cột A đến F, cột M và N
B9: Cố định xem được tiêu đề (freeze panes)
B10: Tối ưu cột I, K (bỏ xuống dòng + tự động điều chỉnh độ rộng)

🎛️ TÙY CHỌN NÂNG CAO:

1. BẬT TẠO FILE TỔNG HỢP:
   ```python
   from process_excel import ExcelProcessor
   processor = ExcelProcessor()
   processor.enable_summary_creation()
   processor.process_excel_files()
   ```

2. TẮT TẠO FILE TỔNG HỢP:
   ```python
   processor.disable_summary_creation()
   ```

3. XỬ LÝ FILE CỤ THỂ:
   ```python
   processor.process_single_excel(Path("file.xlsx"))
   ```

⚠️ LƯU Ý:

1. File tổng hợp "Kết quả.xlsx" mặc định TẮT do hạn chế format
2. Mỗi file lẻ sau xử lý có format hoàn hảo
3. Hệ thống không làm mất dữ liệu gốc
4. Chromium browser cần được cài đặt (tự động)

🔧 KHẮC PHỤC LỖI:

1. Lỗi import module:
   - Kiểm tra Python environment
   - Cài đặt: pip install openpyxl playwright

2. Lỗi browser:
   - Chạy: python -m playwright install chromium

3. Lỗi file Excel:
   - Kiểm tra file không bị mở trong Excel
   - Kiểm tra quyền ghi file

4. Lỗi thư mục:
   - Kiểm tra thư mục output/DDMMYYYY tồn tại
   - Kiểm tra quyền ghi thư mục

📞 HỖ TRỢ:

- Chạy test_system.py để kiểm tra chi tiết
- Xem log chi tiết khi chạy
- Kiểm tra file config.json và template.xlsx

🎉 HOÀN TẤT!
Hệ thống đã sẵn sàng sử dụng!
"""
