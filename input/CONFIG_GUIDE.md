# 📋 CONFIG FILE GUIDE

## Hướng dẫn cấu hình file config.json

File này chứa tất cả thông tin cần thiết để automation hoạt động với website của bạn.

### 1. Website Configuration
```json
"website": {
  "url": "https://cholimexfood.dmsone.vn/login",
  "name": "Order Management System"
}
```
- **url**: URL trang đăng nhập của hệ thống
- **name**: Tên hiển thị của hệ thống (chỉ để ghi nhớ)

### 2. Credentials
```json
"credentials": {
  "username": "GADMIN",
  "password": "Ch0oL!iM3ex#2024"
}
```
- **username**: Tên đăng nhập của bạn
- **password**: Mật khẩu của bạn

### 3. CSS/XPath Selectors - QUAN TRỌNG NHẤT
```json
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
}
```

#### Giải thích từng selector:

**Đăng nhập:**
- `username_field`: Ô nhập tên đăng nhập
- `password_field`: Ô nhập mật khẩu  
- `login_button`: Nút đăng nhập
- `dashboard_indicator`: Element hiển thị sau khi đăng nhập thành công

**Navigation Menu:**
- `menu_638`: Menu chính (ID 638)
- `dms_report_kpi`: Submenu DMS Report KPI
- `rpt_kpi_staff`: Link đến báo cáo KPI Staff

**Báo cáo:**
- `kpi_dropdown`: Dropdown chọn loại KPI/báo cáo
- `dhtc_option_text`: Text hiển thị của option trong dropdown
- `from_month_field`: Field chọn tháng/năm
- `month_year_picker_year`: Dropdown chọn năm
- `month_year_picker_month`: Cell chọn tháng (có placeholder {month})
- `search_button`: Nút Search để tải báo cáo
  "dashboard_indicator": "//*[@id=\"638\"]",
  "search_button": "//*[@id=\"btnSearch\"]/span",
  "menu_638": "//*[@id=\"638\"]/a",
  "dms_report_kpi": "//*[@id=\"DMS_REPORT_KPI\"]/ins",
  "rpt_kpi_staff": "//*[@id=\"RPT_KPI_STAFF\"]/a",
  "kpi_dropdown": "//*[@id=\"lstKPI\"]",
  "dhtc_option_value": "3811"
}
```

#### 📝 Mô tả từng selector:
- **username_field**: Ô nhập tên đăng nhập
- **password_field**: Ô nhập mật khẩu  
- **login_button**: Nút đăng nhập
- **dashboard_indicator**: Element xác nhận đăng nhập thành công
- **search_button**: Nút Search để tải báo cáo
- **menu_638**: Menu item chính (ID=638)
- **dms_report_kpi**: Nút expand menu DMS_REPORT_KPI
- **rpt_kpi_staff**: Menu con RPT_KPI_STAFF
- **kpi_dropdown**: Dropdown chọn loại KPI
- **dhtc_option_value**: Giá trị option "DHTC - Đơn hàng thành công"

#### 🔄 Template Backup:
- File `config.template.json` chứa backup đầy đủ
- Nếu `config.json` bị xóa/hỏng → copy từ template
- Template luôn được cập nhật với selectors mới nhất

#### Cách tìm CSS Selectors:
1. Mở trang đăng nhập trong browser
2. Nhấn F12 để mở Developer Tools
3. Click vào nút "Select Element" (mũi tên)
4. Click vào element cần tìm
5. Trong HTML, right-click element → Copy → Copy selector

#### Ví dụ selectors phổ biến:
- **ID**: `#username`, `#email`, `#login`
- **Class**: `.form-control`, `.btn-primary`
- **Attribute**: `[name='username']`, `[type='submit']`
- **Combined**: `input[name='username']`, `button.login-btn`

### 4. Settings
```json
"settings": {
  "wait_timeout": 10000,
  "screenshot_on_success": true,
  "screenshot_on_error": true
}
```
- **wait_timeout**: Thời gian chờ tối đa (milliseconds)
- **screenshot_on_success**: Chụp ảnh khi thành công
- **screenshot_on_error**: Chụp ảnh khi có lỗi

## 🔧 Troubleshooting

### Lỗi "Username field not found"
- Kiểm tra lại selector của ô username
- Thử dùng các selector khác như `[name='username']`, `.username-input`

### Lỗi "Login button not found"  
- Kiểm tra selector của nút đăng nhập
- Thử `button[type='submit']`, `.btn-login`, `input[type='submit']`

### Lỗi "Dashboard not loaded"
- Kiểm tra selector của element xuất hiện sau khi đăng nhập
- Có thể là `.header`, `.navbar`, `.user-menu`

### Tips để tìm selectors:
1. Tìm element có **id** trước (ưu tiên #id)
2. Nếu không có id, tìm **class** duy nhất
3. Cuối cùng dùng **attribute** như name, type
4. Test selector trong Console: `document.querySelector("YOUR_SELECTOR")`

## ⚠️ Security Notes
- Không share file config.json chứa mật khẩu
- Cân nhắc sử dụng biến môi trường cho production
- Backup file config template trước khi chỉnh sửa
