# 🎯 ORDER CHECKER - EXCEL AUTOMATION TOOL v2.0

**Hệ thống tự động check orders và xử lý Excel chuyên nghiệp với 10 bước nghiệp vụ**

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://python.org)
[![Playwright](https://img.shields.io/badge/Playwright-Latest-green.svg)](https://playwright.dev)
[![Version](https://img.shields.io/badge/Version-2.0.0-brightgreen.svg)](https://github.com/your-repo/releases)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Contributions](https://img.shields.io/badge/Contributions-Welcome-brightgreen.svg)](CONTRIBUTING.md)

## ✨ Tính năng chính v2.0

🚀 **Hệ thống tự động hoàn chỉnh**: Check orders + xử lý Excel tích hợp  
📊 **Xử lý Excel chuyên nghiệp**: 10 bước nghiệp vụ (ẩn dòng/cột, freeze panes, auto-fit)  
🎛️ **Menu quản lý thân thiện**: Giao diện dễ sử dụng với menu.py  
🔍 **Công cụ kiểm tra**: test_system.py để debug và monitoring  
📋 **File tổng hợp tùy chọn**: Tạo "Kết quả.xlsx" (mặc định tắt)  
📖 **Tài liệu chi tiết**: Hướng dẫn đầy đủ và tổng kết hệ thống  
🔧 **Không mất dữ liệu gốc**: Bảo toàn file nguồn, dễ mở rộng  
⚡ **Hiệu suất cao**: Xử lý nhanh, báo cáo chi tiết  

## 📋 Luồng hoạt động

1. **Đọc template** từ `input/template.xlsx` để lấy danh sách báo cáo cần tải
2. **Đăng nhập** tự động vào hệ thống với thông tin từ config
3. **Điều hướng** đến phần báo cáo KPI
4. **Lặp qua từng báo cáo**:
   - Chọn KPI theo tên báo cáo
   - Chọn tháng/năm (ngày hiện tại lùi 1 ngày)
   - Tải file Excel và đặt tên theo "Tên viết tắt"
5. **Xử lý Excel files**:
   - Ẩn dòng theo điều kiện (cột C có dữ liệu AND cột D trống)
   - Ẩn cột A-F, M-N và từ S trở đi
   - Bỏ wrap text, canh giữa dọc, auto width
6. **Tạo file kết quả** gộp tất cả sheet vào "Kết quả.xlsx"
7. **Phân tích dữ liệu** và tạo báo cáo chi tiết  

## 🚀 Cài đặt và chạy

### 🎯 Cài đặt nhanh (Recommended)

```bash
# Clone repository
git clone <repository-url>
cd "Check Oders"

# Chạy setup tự động
python setup.py
```

Setup script sẽ tự động:
- ✅ Tạo virtual environment  
- ✅ Cài đặt dependencies từ requirements.txt
- ✅ Tải và setup Chromium browser
- ✅ Tạo thư mục input/output
- ✅ Kiểm tra config files

### 🔧 Cài đặt thủ công

```bash
# Clone repository
git clone <repository-url>
cd "Check Oders"

# Tạo virtual environment
python -m venv myenv

# Kích hoạt virtual environment
# Windows:
myenv\Scripts\activate
# Linux/Mac:
source myenv/bin/activate

# Cài đặt dependencies
pip install -r requirements.txt

# Cài đặt Chromium
python -m playwright install chromium

# Copy Chromium vào dự án (để đóng gói)
python setup_chromium.py
```

### 2. Cấu hình template báo cáo

File `input/template.xlsx` chứa danh sách báo cáo cần tải:

| Tên viết tắt | Tên báo cáo |
|--------------|-------------|
| NMCD 6 loại | NMCD 6 loại |
| Mắm HV (500&750ML) | Mắm HV (500&750ML) |

- **Tên viết tắt**: Tên file Excel sau khi tải về
- **Tên báo cáo**: Tên hiển thị trong dropdown KPI trên web

### 3. Cấu hình đăng nhập

Chạy lần đầu tiên để tạo file config:

```bash
python check_oder.py
```

Sau đó cập nhật `input/config.json`:

```json
{
  "website": {
    "url": "https://cholimexfood.dmsone.vn/login",
    "name": "Order Management System"
  },
  "credentials": {
    "username": "YOUR_USERNAME",
    "password": "YOUR_PASSWORD"
  },
  "selectors": {
    "username_field": "#username",
    "password_field": "#password",
    "login_button": "//*[@id=\"fm1\"]/fieldset/div[3]/input[3]",
    "dashboard_indicator": "//*[@id=\"638\"]",
    "search_button": "//*[@id=\"btnSearch\"]/span",
    "menu_638": "//*[@id=\"638\"]/a",
    "dms_report_kpi": "//*[@id=\"DMS_REPORT_KPI\"]/ins",
    "rpt_kpi_staff": "//*[@id=\"RPT_KPI_STAFF\"]/a",
    "kpi_dropdown": "//*[@id=\"lstKPI\"]",
    "dhtc_option_text": "DHTC - Đơn hàng thành công",
    "from_month_field": "//*[@id=\"fromMonth\"]",
    "month_year_picker_year": "select.mtz-monthpicker.mtz-monthpicker-year",
    "month_year_picker_month": "td.mtz-monthpicker-month[data-month=\"{month}\"]"
  },
  "settings": {
    "wait_timeout": 30000,
    "auto_wait": true,
    "fast_mode": false
  }
}
```

### 4. Chạy automation

```bash
# Chạy từ source code
python check_oder.py
```

### 5. Đóng gói thành executable

```bash
# Đóng gói
python build.py

# Chọn option:
# 1 = Single file .exe (--onefile)
# 2 = Directory with files (--onedir)

# File .exe sẽ được tạo tại:
# dist/OrderChecker.exe (option 1)
# dist/OrderChecker/OrderChecker.exe (option 2)
```
##  Cấu trúc dự án

```
Check Oders/
├── 📄 check_oder.py          # File chính - OrderChecker class
├── 🔧 setup_chromium.py      # Setup Chromium cho đóng gói
├── 📦 build.py               # Script đóng gói executable
├── 📋 requirements.txt       # Python dependencies
├── 📖 README.md              # Tài liệu này
├── 🌐 myenv/                 # Virtual environment
├── 🖥️ chromium-browser/      # Chromium browser (sau setup)
│
├── 📥 input/                 # Thư mục input (config & template)
│   ├── 📊 template.xlsx      # Template danh sách báo cáo
│   ├── ⚙️ config.json        # Cấu hình đăng nhập & selectors
│   ├── 📝 config.template.json # Template config mẫu
│   └── 📘 CONFIG_GUIDE.md    # Hướng dẫn chi tiết
│
├── 📤 output/                # Thư mục output (kết quả theo ngày)
│   └── 📅 DDMMYYYY/          # Thư mục theo ngày (vd: 05072025)
│       ├── 📈 NMCD 6 loại.xlsx
│       ├── 📈 Mắm HV (500&750ML).xlsx
│       ├── 🎯 Kết quả.xlsx   # File gộp tất cả sheet
│       └── 📊 analysis_report.txt
│
└── 🚀 dist/                  # Executable files
    └── OrderChecker/
        ├── OrderChecker.exe  # File thực thi chính
        └── _internal/        # Dependencies & browser
```

## 💡 Cách sử dụng

### Bước 1: Chuẩn bị template
Chỉnh sửa `input/template.xlsx` với danh sách báo cáo cần tải:
- **Cột A**: Tên viết tắt (tên file sau khi tải)  
- **Cột B**: Tên báo cáo (hiển thị trong dropdown web)

### Bước 2: Chạy automation
```bash
# Từ source code
python check_oder.py

# Hoặc từ executable
./dist/OrderChecker/OrderChecker.exe
```

### Bước 3: Kiểm tra kết quả
- **Raw files**: `output/DDMMYYYY/*.xlsx` (files gốc đã tải)
- **Processed file**: `output/DDMMYYYY/Kết quả.xlsx` (file đã xử lý)
- **Analysis**: `output/DDMMYYYY/analysis_report.txt` (phân tích chi tiết)

## 🔧 Excel Processing Logic

### Filtering Rules
- **Hide Rows**: Cột C có dữ liệu AND cột D trống
- **Hide Columns**: A-F, M-N, và từ S trở đi
- **Visible Columns**: G-L, O-R (dữ liệu chính)

### Formatting Applied
- ✅ Bỏ wrap text từ dòng 6 trở đi
- ✅ Canh giữa theo chiều dọc (vertical center)
- ✅ Auto width cho các cột hiển thị
- ✅ Giữ nguyên 5 dòng tiêu đề (1-5)
- ✅ Copy đầy đủ formatting (font, border, color)

## 🎯 Output Console

```
────────────────────────────────────────────────────────────
📥 Đang tải báo cáo...
   📅 Chọn tháng/năm: 7/2025
   ✅ Đã chọn năm: 2025
   ✅ Đã chọn tháng: T7
   ✅ Tải file 1: NMCD 6 loại thành công
   📅 Chọn tháng/năm: 7/2025
   ✅ Đã chọn năm: 2025
   ✅ Đã chọn tháng: T7
   ✅ Tải file 2: Mắm HV (500&750ML) thành công
✅ Đã tải thành công 2 báo cáo
────────────────────────────────────────────────────────────
────────────────────────────────────────────────────────────
📊 Đang xử lý và tạo file kết quả...
   ✅ Thêm Sheet 1: Mắm HV (500&750ML) thành công
   ✅ Thêm Sheet 2: NMCD 6 loại thành công
🎉 Hoàn thành tạo file: Kết quả.xlsx
────────────────────────────────────────────────────────────
────────────────────────────────────────────────────────────
🎉 HOÀN THÀNH TẤT CẢ!
────────────────────────────────────────────────────────────
```

## 🛠️ Technical Details

### Dependencies
- **playwright**: Web automation
- **openpyxl**: Excel file processing  
- **pyinstaller**: Executable packaging

### Browser Mode
- **Headless**: True (chạy ẩn, không hiển thị giao diện)
- **Auto-close**: Safe browser cleanup
- **Chromium**: Embedded, portable

### Date Logic
- **Target Date**: Ngày hiện tại lùi 1 ngày
- **Auto Selection**: Tự động chọn tháng/năm tương ứng

## � Troubleshooting

### ❌ Config không hợp lệ
**Nguyên nhân**: Thiếu username/password hoặc URL không đúng  
**Giải pháp**: Kiểm tra và cập nhật `input/config.json`

### ❌ Đăng nhập thất bại  
**Nguyên nhân**: Sai thông tin đăng nhập hoặc selector không đúng  
**Giải pháp**: 
- Kiểm tra username/password
- Cập nhật CSS selectors trong config
- Xem hướng dẫn chi tiết: `input/CONFIG_GUIDE.md`

### ❌ Điều hướng thất bại
**Nguyên nhân**: Selector menu không đúng  
**Giải pháp**: Cập nhật các selector trong phần "selectors" của config

### ❌ Tải file thất bại
**Nguyên nhân**: 
- Tên báo cáo trong template không khớp với dropdown
- Selector tháng/năm không đúng
**Giải pháp**:
- Kiểm tra tên báo cáo trong `input/template.xlsx`
- Cập nhật selector cho month/year picker

### ❌ Browser not found
```bash
python setup_chromium.py
python -m playwright install chromium
```

### ❌ Lỗi đóng gói
```bash
pip uninstall pyinstaller
pip install pyinstaller
python build.py
```

### ❌ Permission Error
**Windows**: Chạy terminal với quyền Administrator  
**Linux/Mac**: Kiểm tra quyền thực thi cho file executable

## 🔍 Debug Mode

Để debug, thay đổi trong code:
```python
# Trong run_browser_test()
browser = p.chromium.launch(headless=False)  # Hiển thị browser
```

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details on:

- 🐛 **Bug Reports**: How to report issues effectively
- ✨ **Feature Requests**: Suggesting new functionality  
- 🔧 **Code Contributions**: Development setup and guidelines
- 📚 **Documentation**: Improving guides and examples

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 📋 Changelog

See [CHANGELOG.md](CHANGELOG.md) for detailed version history and updates.

## 📞 Support

📧 **Issues**: Tạo issue trên GitHub repository  
📚 **Documentation**: Xem `input/CONFIG_GUIDE.md` để biết chi tiết  
🛠️ **Custom Development**: Contact cho tùy chỉnh selector theo website khác

---
**Made with ❤️ using Playwright & Python**

## 🗂️ Cấu trúc Project v2.0

```
📦 Check Oders/
├── 🚀 check_oder.py           # Hệ thống chính - tự động check orders
├── 📊 process_excel.py        # Xử lý Excel chuyên nghiệp (10 bước)
├── 🎛️ menu.py                 # Menu quản lý hệ thống
├── 🔍 test_system.py          # Kiểm tra và debug hệ thống
├── 📖 HUONG_DAN.md            # Hướng dẫn chi tiết v2.0
├── 📋 FINAL_SUMMARY.md        # Tổng kết hệ thống hoàn thiện
├── 📂 input/                  # Cấu hình và template
│   ├── 📋 template.xlsx       # Danh sách báo cáo
│   └── ⚙️ config.json         # Cấu hình hệ thống
├── 📂 output/                 # Kết quả theo ngày
│   └── 📅 DDMMYYYY/           # File Excel đã xử lý
└── 🐍 myenv/                  # Python virtual environment
```

## 🎯 Cách sử dụng v2.0

### 🚀 Cách nhanh nhất - Menu quản lý:

```bash
python menu.py
```

Menu cung cấp:
- 🚀 Chạy hệ thống hoàn chỉnh
- 📊 Chỉ xử lý Excel
- 🔍 Kiểm tra hệ thống
- 📋 Bật/tắt file tổng hợp
- 📁 Mở thư mục kết quả
- 📖 Xem hướng dẫn

### 📊 Xử lý Excel chuyên nghiệp (10 bước):

```bash
python process_excel.py
```

**Nghiệp vụ xử lý:**
1. Ẩn dòng 1-3 (header)
2. Ẩn dòng có cột A rỗng
3. Ẩn dòng có cột B rỗng
4. Ẩn dòng có cột D rỗng AND cột C ≠ ""
5. Xóa dữ liệu dòng có cột C rỗng (từ K trở đi)
6. Ẩn dòng K chứa "NPP bán"
7. Ẩn dòng có cột Q > 0
8. Ẩn dòng rỗng liên tiếp cột Q
9. Ẩn cột S+, A-F, M-N
10. Freeze panes + Auto-fit cột I/K

### 🔍 Kiểm tra hệ thống:

```bash
python test_system.py
```

Kiểm tra:
- ✅ Import modules
- ✅ Thư mục và file
- ✅ Cấu hình
- ✅ Tích hợp
