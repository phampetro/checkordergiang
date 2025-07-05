import os
import sys
import json
import time
import shutil
import warnings
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

import openpyxl

# Tắt warning openpyxl về default style
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

class OrderChecker:
    def ensure_template_excel(self):
        """
        Kiểm tra file input/template.xlsx, nếu chưa có thì tạo file mẫu với 2 cột: Tên viết tắt, Tên báo cáo
        """
        excel_path = self.input_dir / "template.xlsx"
        if not excel_path.exists():
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.append(["Tên viết tắt", "Tên báo cáo"])
            ws.append(["DHTC", "DHTC - Đơn hàng thành công"])
            wb.save(excel_path)
    def load_report_list_from_excel(self):
        """
        Đọc danh sách báo cáo từ input/template.xlsx
        Trả về list các dict: { 'short_name': ..., 'report_name': ... }
        """
        excel_path = self.input_dir / "template.xlsx"
        if not excel_path.exists():
            return None
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        if not rows:
            return None
        report_list = []
        for row in rows:
            if row[0] and row[1]:
                report_list.append({
                    'short_name': str(row[0]).strip(),
                    'report_name': str(row[1]).strip()
                })
        if not report_list:
            return None
        return report_list
    def __init__(self):
        self.base_path = self.get_base_path()
        self.input_dir = self.setup_directory("input")
        self.output_dir = self.setup_directory("output")
        self.daily_output_dir = self.setup_daily_output_directory()
        self.ensure_template_excel()
        self.chromium_path = self.get_chromium_path()
        self.config_file = self.input_dir / "config.json"
        self.config = self.load_or_create_config()
        
    def get_base_path(self):
        """
        Kiểm tra chế độ chạy và trả về đường dẫn base
        - Dev mode: thư mục chứa file .py
        - Packaged mode: thư mục chứa file .exe hoặc _internal
        """
        if getattr(sys, 'frozen', False):
            # Chế độ đóng gói (exe)
            # Kiểm tra xem có phải --onefile hay --onedir
            if hasattr(sys, '_MEIPASS'):
                # --onefile: sys._MEIPASS là thư mục tạm
                # Base path là thư mục chứa .exe
                base_path = Path(sys.executable).parent
            else:
                # --onedir: sys.executable là trong thư mục dist
                base_path = Path(sys.executable).parent
        else:
            # Chế độ dev
            base_path = Path(__file__).parent
            
        return base_path
    
    def setup_directory(self, dir_name):
        """
        Tạo thư mục nếu chưa tồn tại
        Trong chế độ --onedir, file được đặt trong _internal
        """
        if getattr(sys, 'frozen', False):
            # Packaged mode - tìm trong _internal trước
            internal_path = self.base_path / "_internal" / dir_name
            if internal_path.exists():
                return internal_path
            
        # Dev mode hoặc fallback
        dir_path = self.base_path / dir_name
        if not dir_path.exists():
            dir_path.mkdir(parents=True, exist_ok=True)
        return dir_path
    
    def get_chromium_path(self):
        """
        Tìm đường dẫn Chromium trong thư mục dự án
        """
        # Danh sách thư mục để tìm Chromium theo chế độ
        if getattr(sys, 'frozen', False):
            # Packaged mode - tìm trong _internal hoặc cùng cấp
            chromium_base_dirs = [
                self.base_path / "_internal" / "chromium-browser",  # --onedir
                self.base_path / "chromium-browser",               # Backup location
                self.base_path / "chromium",
            ]
        else:
            # Dev mode - tìm trong thư mục project
            chromium_base_dirs = [
                self.base_path / "chromium-browser",
                self.base_path / "chromium", 
                self.base_path / "browser",
            ]
        
        for base_dir in chromium_base_dirs:
            if base_dir.exists():
                # Tìm file chrome.exe trong tất cả thư mục con
                for chrome_exe in base_dir.rglob("chrome.exe"):
                    return str(chrome_exe)
        
        return None
    
    def install_browser(self):
        """
        Cài đặt Chromium vào thư mục dự án
        """
        browser_dir = self.base_path / "chromium-browser"
        
        # Tạo thư mục browser nếu chưa có
        browser_dir.mkdir(exist_ok=True)
        
        # Set environment variable để playwright cài vào thư mục này
        os.environ['PLAYWRIGHT_BROWSERS_PATH'] = str(browser_dir)
        
        # Cài đặt chromium
        os.system(f'"{sys.executable}" -m playwright install chromium')
        
        # Tìm lại đường dẫn sau khi cài
        self.chromium_path = self.get_chromium_path()
    
    def run_browser_test(self, url=None):
        """
        Chạy automation: đăng nhập, duyệt từng báo cáo trong template.xlsx, tải về và đặt tên file theo Tên viết tắt
        """
        if not self.config:
            return False

        report_list = self.load_report_list_from_excel()
        if not report_list:
            return False

        if not url:
            url = self.config['website']['url']

        try:
            with sync_playwright() as p:
                if self.chromium_path and os.path.exists(self.chromium_path):
                    browser = p.chromium.launch(
                        executable_path=self.chromium_path,
                        headless=True
                    )
                else:
                    browser = p.chromium.launch(headless=True)

                context = browser.new_context()
                page = context.new_page()

                login_success = self.login_to_website(page)
                if not login_success:
                    print("❌ Đăng nhập thất bại!")
                    return False

                navigation_success = self.navigate_to_reports(page)
                if not navigation_success:
                    print("❌ Điều hướng thất bại!")
                    return False

                # Lặp qua từng báo cáo trong Excel
                all_success = True
                downloaded_files = []
                
                print("─" * 60)
                print("📥 Đang tải báo cáo...")
                for idx, report in enumerate(report_list, 1):
                    success = self.select_kpi_and_download(page, report['report_name'], report['short_name'])
                    if success:
                        downloaded_files.append(report['short_name'])
                        print(f"   ✅ Tải file {idx}: {report['short_name']} thành công")
                    else:
                        print(f"   ❌ Tải file {idx}: {report['short_name']} thất bại")
                        all_success = False
                
                if downloaded_files:
                    print(f"✅ Đã tải thành công {len(downloaded_files)} báo cáo")
                print("─" * 60)
                
                # Đóng browser an toàn
                try:
                    context.close()
                    browser.close()
                except:
                    pass
                
                # Xử lý các file Excel đã tải về
                if downloaded_files:
                    print("\n" + "─" * 60)
                    print("📊 Đang xử lý và tạo file kết quả...")
                    process_success = self.process_downloaded_excel_files()
                    if process_success:
                        self.analyze_excel_data()
                    print("─" * 60)
                
                return all_success
        except Exception as e:
            print(f"❌ Lỗi browser: {str(e)}")
            return False

    def select_kpi_and_download(self, page, kpi_text, short_name):
        """
        Chọn KPI theo tên (kpi_text), tải file và đặt tên theo short_name
        """
        try:
            kpi_dropdown = self.get_locator(page, self.config['selectors']['kpi_dropdown'])
            kpi_dropdown.click()
            try:
                kpi_dropdown.select_option(label=kpi_text)
            except Exception as e:
                try:
                    dhtc_option = self.get_locator(page, f'//option[contains(text(), "{kpi_text}")]')
                    dhtc_option.click()
                except Exception as e2:
                    try:
                        dhtc_option = self.get_locator(page, f'//option[contains(text(), "{kpi_text.split()[0]}")]')
                        dhtc_option.click()
                    except Exception as e3:
                        return False

            selected_text = kpi_dropdown.locator('option:checked').inner_text()
            if kpi_text in selected_text or kpi_text.split()[0] in selected_text:
                return self.click_search_and_download(page, custom_filename=short_name)
            else:
                return False
        except PlaywrightTimeoutError:
            return False
        except Exception as e:
            return False

    def load_or_create_config(self):
        """
        Tải hoặc tạo file config.json
        """
        if not self.config_file.exists():
            print("📝 Tạo file config.json từ template...")
            self.create_default_config()
            # Thử load lại sau khi tạo
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                return config
            except:
                return None
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # Kiểm tra thông tin bắt buộc
            if not config.get('credentials', {}).get('username') or not config.get('credentials', {}).get('password'):
                print("❌ Config thiếu username/password!")
                return None
            
            if not config.get('website', {}).get('url'):
                print("❌ Config thiếu URL website!")
                return None
            
            return config
            
        except json.JSONDecodeError as e:
            print(f"❌ Config JSON không hợp lệ: {e}")
            return None
        except Exception as e:
            print(f"❌ Lỗi đọc config: {e}")
            return None
    
    def create_default_config(self):
        """
        Tạo file config từ template
        """
        template_file = self.input_dir / "config.template.json"
        
        # Kiểm tra file template có tồn tại không
        if template_file.exists():
            try:
                # Copy từ template
                with open(template_file, 'r', encoding='utf-8') as f:
                    template_config = json.load(f)
                
                # Lưu thành config.json
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    json.dump(template_config, f, indent=2, ensure_ascii=False)
                
                print(f"✅ Đã tạo config.json từ template")
                return
                
            except Exception as e:
                print(f"❌ Lỗi đọc template: {e}")
        
        # Fallback: tạo config mặc định nếu không có template
        default_config = {
            "website": {
                "url": "https://example.com/login",
                "name": "Order Management System"
            },
            "credentials": {
                "username": "",
                "password": ""
            },
            "selectors": {
                "username_field": "#username",
                "password_field": "#password", 
                "login_button": "#login-button",
                "dashboard_indicator": ".dashboard",
                "search_button": "#search-btn",
                "menu_638": "#menu-638",
                "dms_report_kpi": "#dms-report-kpi",
                "rpt_kpi_staff": "#rpt-kpi-staff",
                "kpi_dropdown": "#kpi-dropdown",
                "dhtc_option_text": "DHTC - Đơn hàng thành công",
                "from_month_field": "#from-month-field",
                "month_year_picker_year": "#month-year-picker-year",
                "month_year_picker_month": "//select[@id='month-year-picker-month']/option[text()='T{month}']"
            },
            "settings": {
                "wait_timeout": 30000,
                "auto_wait": True,
                "fast_mode": False
            }
        }
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Đã tạo config.json mặc định")
    
    def login_to_website(self, page):
        """
        Đăng nhập vào website sử dụng thông tin từ config
        """
        if not self.config:
            return False
            
        try:
            url = self.config['website']['url']
            username = self.config['credentials']['username']
            password = self.config['credentials']['password']
            selectors = self.config['selectors']
            timeout = self.config['settings']['wait_timeout']
            
            # Mở trang đăng nhập
            page.goto(url)
            page.wait_for_load_state('networkidle')
            
            # Tìm và điền username
            try:
                username_field = self.get_locator(page, selectors['username_field'])
                self.smart_wait_for_element(username_field, 'visible')
                username_field.fill(username)
            except PlaywrightTimeoutError:
                return False
            
            # Tìm và điền password
            try:
                password_field = self.get_locator(page, selectors['password_field'])
                self.smart_wait_for_element(password_field, 'visible')
                password_field.fill(password)
            except PlaywrightTimeoutError:
                return False
            
            # Click nút đăng nhập
            try:
                login_button = self.get_locator(page, selectors['login_button'])
                self.smart_wait_for_element(login_button, 'visible')
                login_button.click()
            except PlaywrightTimeoutError:
                return False
            
            # Đợi trang dashboard load
            try:
                dashboard_indicator = self.get_locator(page, selectors['dashboard_indicator'])
                self.smart_wait_for_element(dashboard_indicator, 'visible')
                return True
                
            except PlaywrightTimeoutError:
                return False
                
        except Exception as e:
            return False

    def navigate_to_reports(self, page):
        """
        Điều hướng đến trang báo cáo sau khi đăng nhập thành công
        """
        try:
            menu_item_638 = self.get_locator(page, self.config['selectors']['menu_638'])
            menu_item_638.click()
            page.wait_for_load_state('networkidle')

            dms_report_kpi = self.get_locator(page, self.config['selectors']['dms_report_kpi'])
            dms_report_kpi.click()

            rpt_kpi_staff = self.get_locator(page, self.config['selectors']['rpt_kpi_staff'])
            try:
                rpt_kpi_staff.scroll_into_view_if_needed()
            except:
                pass
            rpt_kpi_staff.click()
            page.wait_for_load_state('networkidle')
            return True
        except PlaywrightTimeoutError:
            return False
        except Exception as e:
            return False

    def select_kpi_dropdown(self, page):
        """
        Chọn KPI từ dropdown lstKPI
        """
        try:
            print("   Looking for dropdown #lstKPI...")
            kpi_dropdown = self.get_locator(page, self.config['selectors']['kpi_dropdown'])
            kpi_dropdown.click()
            print("✅ Step 3a: Clicked KPI dropdown")

            print(f"   Selecting '{self.config['selectors']['dhtc_option_text']}'...")
            try:
                kpi_dropdown.select_option(label=self.config['selectors']['dhtc_option_text'])
                print("✅ Step 3b: Selected DHTC option using label")
            except Exception as e:
                print(f"   Method 1 failed: {e}")
                try:
                    option_text = self.config['selectors']['dhtc_option_text']
                    dhtc_option = self.get_locator(page, f'//option[contains(text(), "{option_text}")]')
                    dhtc_option.click()
                    print("✅ Step 3b: Selected DHTC option using text search")
                except Exception as e2:
                    print(f"   Method 2 failed: {e2}")
                    try:
                        dhtc_option = self.get_locator(page, '//option[contains(text(), "DHTC")]')
                        dhtc_option.click()
                        print("✅ Step 3b: Selected DHTC option using partial text")
                    except Exception as e3:
                        print(f"   All methods failed: {e3}")
                        return False

            selected_text = kpi_dropdown.locator('option:checked').inner_text()
            expected_text = self.config['selectors']['dhtc_option_text']
            if expected_text in selected_text or "DHTC" in selected_text:
                print(f"✅ Step 3c: Confirmed DHTC option is selected: {selected_text}")
                print("🔍 Step 4: Looking for Search button...")
                search_success = self.click_search_and_download(page)
                if search_success:
                    print("✅ Step 4: Search and download completed successfully!")
                    return True
                else:
                    print("❌ Step 4: Search and download failed!")
                    return False
            else:
                print(f"❌ Step 3c: Wrong option selected. Current text: {selected_text}")
                return False
        except PlaywrightTimeoutError:
            print("❌ Step 3 failed: KPI dropdown not found")
            return False
        except Exception as e:
            print(f"❌ KPI selection error: {e}")
            return False

    def select_month_year_before_search(self, page):
        """
        Chọn tháng/năm trước khi click search
        Chọn ngày hiện tại lùi 1 ngày để xác định tháng/năm
        """
        try:
            from datetime import datetime, timedelta
            
            # Tính ngày hiện tại lùi 1 ngày
            yesterday = datetime.now() - timedelta(days=1)
            target_year = yesterday.year
            target_month = yesterday.month
            
            print(f"   📅 Chọn tháng/năm: {target_month}/{target_year}")
            
            # Bước 1: Click vào field fromMonth để mở month/year picker
            from_month_field = self.get_locator(page, self.config['selectors']['from_month_field'])
            from_month_field.click()
            page.wait_for_timeout(1000)  # Đợi picker hiện lên
            
            # Bước 2: Chọn năm
            year_selector = self.get_locator(page, self.config['selectors']['month_year_picker_year'])
            year_selector.select_option(value=str(target_year))
            print(f"   ✅ Đã chọn năm: {target_year}")
            page.wait_for_timeout(500)
            
            # Bước 3: Chọn tháng
            month_selector_template = self.config['selectors']['month_year_picker_month']
            month_selector = month_selector_template.format(month=target_month)
            month_element = self.get_locator(page, month_selector)
            month_element.click()
            print(f"   ✅ Đã chọn tháng: T{target_month}")
            page.wait_for_timeout(1000)
            
            return True
            
        except Exception as e:
            print(f"   ❌ Lỗi chọn tháng/năm: {str(e)}")
            return False

    def click_search_and_download(self, page, custom_filename=None):
        """
        Click nút Search và xử lý download file với retry logic
        Nếu custom_filename được truyền vào thì đặt tên file tải về theo tên này
        """
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
                # Bước mới: Chọn tháng/năm trước khi search
                month_year_success = self.select_month_year_before_search(page)
                if not month_year_success:
                    print(f"   ❌ Không thể chọn tháng/năm, thử lại lần {attempt}")
                    if attempt < max_retries:
                        if self.retry_navigation_from_menu(page):
                            continue
                    return False
                
                search_button = self.get_locator(page, self.config['selectors']['search_button'])
                if custom_filename:
                    download_path = self.daily_output_dir / f"{custom_filename}.xlsx"
                else:
                    download_path = self.daily_output_dir / f"report_{int(time.time())}.xlsx"
                with page.expect_download(timeout=180000) as download_info:
                    search_button.click()
                download = download_info.value
                download.save_as(str(download_path))
                if download_path.exists() and download_path.stat().st_size > 0:
                    return True
                else:
                    if attempt < max_retries:
                        if self.retry_navigation_from_menu(page):
                            continue
                    print(f"❌ Tải thất bại: {custom_filename or download_path.name}")
                    return False
            except Exception as e:
                if attempt < max_retries:
                    if self.retry_navigation_from_menu(page):
                        continue
                else:
                    print(f"❌ Tải thất bại: {custom_filename or 'Unknown'} - Lỗi: {str(e)}")
                    return False
        return False
    
    def retry_navigation_from_menu(self, page):
        """
        Thực hiện lại navigation từ menu #638 khi retry
        """
        try:
            timeout = self.config['settings']['wait_timeout']
            
            # Refresh page và đợi load
            page.reload()
            page.wait_for_load_state('networkidle')
            page.wait_for_timeout(3000)
            
            # Click menu #638
            menu_item_638 = self.get_locator(page, self.config['selectors']['menu_638'])
            menu_item_638.wait_for(state='visible', timeout=timeout)
            menu_item_638.click()
            
            page.wait_for_load_state('networkidle')
            page.wait_for_timeout(2000)
            
            # Click DMS_REPORT_KPI
            dms_report_kpi = self.get_locator(page, self.config['selectors']['dms_report_kpi'])
            dms_report_kpi.wait_for(state='visible', timeout=timeout)
            dms_report_kpi.click()
            
            page.wait_for_timeout(1000)
            
            # Click RPT_KPI_STAFF
            rpt_kpi_staff = self.get_locator(page, self.config['selectors']['rpt_kpi_staff'])
            rpt_kpi_staff.wait_for(state='visible', timeout=timeout)
            rpt_kpi_staff.click()
            
            page.wait_for_load_state('networkidle')
            page.wait_for_timeout(2000)
            
            # Bỏ việc chọn KPI dropdown ở đây vì sẽ được xử lý trong select_kpi_and_download
            
            return True
            
        except Exception as e:
            return False

    def smart_wait_for_element(self, locator, state='visible'):
        """
        Smart waiting strategy - sử dụng auto wait hoặc custom timeout
        """
        settings = self.config.get('settings', {})
        auto_wait = settings.get('auto_wait', True)
        fast_mode = settings.get('fast_mode', False)
        
        if auto_wait:
            # Sử dụng Playwright default timeout (30s) - nhanh và thông minh
            if fast_mode:
                # Fast mode: timeout ngắn hơn
                locator.wait_for(state=state, timeout=5000)  # 5 giây
            else:
                # Normal mode: dùng default của Playwright
                locator.wait_for(state=state)  # 30 giây default
        else:
            # Custom timeout từ config
            timeout = settings.get('wait_timeout', 30000)
            locator.wait_for(state=state, timeout=timeout)

    def get_locator(self, page, selector):
        """
        Tạo locator từ selector (hỗ trợ CSS và XPath)
        """
        if selector.startswith('/'):
            # XPath selector
            return page.locator(f"xpath={selector}")
        else:
            # CSS selector
            return page.locator(selector)

    def setup_daily_output_directory(self):
        """
        Tạo thư mục output theo ngày hiện tại (format: DDMMYYYY)
        Nếu đã tồn tại thì xóa và tạo lại để đảm bảo dữ liệu mới nhất
        """
        today = datetime.now().strftime("%d%m%Y")
        daily_dir = self.output_dir / today
        
        if daily_dir.exists():
            shutil.rmtree(daily_dir)
        
        daily_dir.mkdir(parents=True, exist_ok=True)
        
        return daily_dir

    def process_downloaded_excel_files(self):
        """
        Xử lý tất cả các file Excel đã tải về trong thư mục ngày hiện tại
        """
        # Chỉ xử lý file trong thư mục ngày hiện tại
        excel_files = list(self.daily_output_dir.glob("*.xlsx"))
        
        if not excel_files:
            print("❌ Không tìm thấy file Excel nào để xử lý!")
            return False
        
        processed_count = 0
        processed_files = []  # Lưu danh sách file đã xử lý để gộp
        
        for excel_file in excel_files:
            try:
                # Xử lý nâng cao với filtering
                processed_ws = self.process_excel_with_advanced_filtering_return_sheet(excel_file)
                if processed_ws:
                    processed_files.append((excel_file, processed_ws))
                    processed_count += 1
                else:
                    print(f"❌ Xử lý thất bại: {excel_file.name}")
            except Exception as e:
                print(f"❌ Xử lý thất bại: {excel_file.name} - Lỗi: {str(e)}")
        
        # Tạo file kết quả gộp
        if processed_files:
            result_file = self.create_consolidated_result_file(processed_files)
            if result_file:
                print("🎉 Hoàn thành tạo file: Kết quả.xlsx")
                return True
        else:
            print("❌ Không có file nào được xử lý thành công!")
        
        return processed_count > 0
    
    def process_single_excel_file(self, excel_file):
        """
        Xử lý một file Excel cụ thể theo quy trình:
        1. Giữ nguyên 5 dòng tiêu đề
        2. Lọc cột C (bỏ blanks), cột D (chỉ blanks)
        3. Ẩn cột A-F và M
        """
        try:
            # Mở file Excel
            wb = openpyxl.load_workbook(excel_file)
            
            # Lấy thông tin cơ bản
            sheet_names = wb.sheetnames
            print(f"   📄 Sheets: {sheet_names}")
            
            # Xử lý sheet đầu tiên
            ws = wb.active
            
            # Đếm số dòng có dữ liệu
            row_count = ws.max_row
            col_count = ws.max_column
            print(f"   📊 Kích thước: {row_count} dòng × {col_count} cột")
            
            # Lấy header (dòng đầu tiên)
            if row_count > 0:
                header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
                headers = [str(cell) if cell is not None else "" for cell in header_row]
                print(f"   📋 Headers: {headers[:5]}...")  # Hiển thị 5 cột đầu
            
            # Bước 1: Áp dụng auto filter cho toàn bộ dữ liệu
            if row_count > 5:  # Chỉ áp dụng nếu có dữ liệu ngoài 5 dòng tiêu đề
                data_range = f"A6:{openpyxl.utils.get_column_letter(col_count)}{row_count}"
                ws.auto_filter.ref = data_range
                print(f"   � Áp dụng auto filter cho range: {data_range}")
                
                # Bước 2: Tạo filter cho cột C (bỏ blanks)
                # Filter cột C: chỉ hiển thị các ô có dữ liệu
                col_c_filter = openpyxl.worksheet.filters.FilterColumn(colId=2)  # Cột C (index 2)
                col_c_filter.filters = openpyxl.worksheet.filters.Filters()
                # Thêm filter để loại bỏ blank values
                blank_filter = openpyxl.worksheet.filters.Filter(val="")
                col_c_filter.filters.filter.append(blank_filter)
                ws.auto_filter.filterColumn.append(col_c_filter)
                
                # Bước 3: Tạo filter cho cột D (chỉ blanks)
                col_d_filter = openpyxl.worksheet.filters.FilterColumn(colId=3)  # Cột D (index 3)
                col_d_filter.filters = openpyxl.worksheet.filters.Filters()
                # Chỉ hiển thị blank values
                blank_only_filter = openpyxl.worksheet.filters.Filter(val="", blank=True)
                col_d_filter.filters.filter.append(blank_only_filter)
                ws.auto_filter.filterColumn.append(col_d_filter)
                
                print(f"   ✅ Đã áp dụng filter: Cột C (bỏ blanks), Cột D (chỉ blanks)")
            
            # Bước 4: Ẩn các cột A-F, M-N và từ S trở đi
            columns_to_hide = ['A', 'B', 'C', 'D', 'E', 'F', 'M', 'N']
            
            # Thêm các cột từ S trở đi vào danh sách ẩn
            for col_num in range(19, col_count + 1):  # S = 19, T = 20, ...
                col_letter = openpyxl.utils.get_column_letter(col_num)
                columns_to_hide.append(col_letter)
            
            # Ẩn các cột
            for col_letter in columns_to_hide:
                ws.column_dimensions[col_letter].hidden = True
            
            print(f"   🙈 Đã ẩn {len(columns_to_hide)} cột: A-F, M-N, S trở đi")
            
            # Bước 5: Hiển thị và format các cột còn lại (G-L, O-R)
            visible_columns = []
            for col_num in range(1, col_count + 1):
                col_letter = openpyxl.utils.get_column_letter(col_num)
                if col_letter not in columns_to_hide:
                    ws.column_dimensions[col_letter].hidden = False
                    visible_columns.append(col_letter)
            
            print(f"   👁️ Các cột hiển thị: {', '.join(visible_columns)}")
            
            # Bước 6: Bỏ thuộc tính wrap text từ dòng 6 trở đi
            print(f"   📝 Bỏ wrap text từ dòng 6-{row_count}...")
            for row_num in range(6, row_count + 1):
                for col_num in range(1, col_count + 1):
                    cell = ws.cell(row_num, col_num)
                    if cell.alignment and cell.alignment.wrap_text:
                        from openpyxl.styles import Alignment
                        cell.alignment = Alignment(wrap_text=False)
            
            # Bước 7: Tự động điều chỉnh độ rộng cột
            print(f"   📏 Tự động điều chỉnh độ rộng cột...")
            for col_letter in visible_columns:
                col_num = openpyxl.utils.column_index_from_string(col_letter)
                max_length = 0
                
                # Tìm độ dài tối đa của nội dung trong cột
                for row_num in range(1, row_count + 1):
                    cell_value = ws.cell(row_num, col_num).value
                    if cell_value:
                        cell_length = len(str(cell_value))
                        if cell_length > max_length:
                            max_length = cell_length
                
                # Đặt độ rộng cột (tối thiểu 8, tối đa 50)
                adjusted_width = min(max(max_length + 2, 8), 50)
                ws.column_dimensions[col_letter].width = adjusted_width
            
            print(f"   ✅ Đã điều chỉnh độ rộng cho {len(visible_columns)} cột hiển thị")
            
            # Tạo file processed
            processed_file = excel_file.parent / f"processed_{excel_file.name}"
            wb.save(processed_file)
            print(f"   💾 Lưu file đã xử lý: {processed_file.name}")
            
            # Tạo file tóm tắt
            summary_file = excel_file.parent / f"summary_{excel_file.stem}.txt"
            with open(summary_file, 'w', encoding='utf-8') as f:
                f.write(f"📊 PROCESSING SUMMARY FOR {excel_file.name}\n")
                f.write(f"{'='*50}\n")
                f.write(f"Original File: {excel_file.name}\n")
                f.write(f"Processed File: processed_{excel_file.name}\n")
                f.write(f"File Size: {excel_file.stat().st_size:,} bytes\n")
                f.write(f"Sheets: {len(sheet_names)}\n")
                f.write(f"Sheet Names: {', '.join(sheet_names)}\n")
                f.write(f"Dimensions: {row_count} rows × {col_count} columns\n")
                f.write(f"Title Rows: 1-5 (preserved)\n")
                f.write(f"Data Rows: 6-{row_count}\n")
                f.write(f"Filter Applied: Column C (non-blanks), Column D (blanks only)\n")
                f.write(f"Hidden Columns: {', '.join(columns_to_hide)}\n")
                f.write(f"Visible Columns: {', '.join(visible_columns)}\n")
                f.write(f"Processed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            
            print(f"   📝 Tạo file tóm tắt: {summary_file.name}")
            
            wb.close()
            return True
            
        except Exception as e:
            print(f"   ❌ Lỗi khi xử lý file: {e}")
            return False

    def analyze_excel_data(self):
        """
        Phân tích chi tiết dữ liệu trong các file Excel đã tải về
        """
        excel_files = list(self.daily_output_dir.glob("*.xlsx"))
        
        if not excel_files:
            return False
        
        analysis_results = {}
        
        for excel_file in excel_files:
            try:
                result = self.analyze_single_excel(excel_file)
                analysis_results[excel_file.name] = result
                
            except Exception as e:
                analysis_results[excel_file.name] = {"error": str(e)}
        
        # Tạo báo cáo tổng hợp
        self.create_analysis_report(analysis_results)
        
        return True
    
    def analyze_single_excel(self, excel_file):
        """
        Phân tích chi tiết một file Excel
        """
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
        
        analysis = {
            "file_name": excel_file.name,
            "file_size": excel_file.stat().st_size,
            "sheets": wb.sheetnames,
            "total_rows": 0,
            "total_cols": 0,
            "has_data": False,
            "headers": [],
            "sample_data": []
        }
        
        # Đếm số dòng và cột có dữ liệu
        max_row = ws.max_row
        max_col = ws.max_column
        
        # Tìm dòng cuối cùng có dữ liệu thực sự
        actual_max_row = 0
        for row_num in range(1, max_row + 1):
            row_data = [ws.cell(row_num, col).value for col in range(1, max_col + 1)]
            if any(cell for cell in row_data):
                actual_max_row = row_num
        
        analysis["total_rows"] = actual_max_row
        analysis["total_cols"] = max_col
        analysis["has_data"] = actual_max_row > 0
        
        if actual_max_row > 0:
            # Lấy headers (dòng đầu tiên)
            headers = [str(ws.cell(1, col).value) if ws.cell(1, col).value else f"Column_{col}" 
                      for col in range(1, max_col + 1)]
            analysis["headers"] = headers
            
            # Lấy 3 dòng dữ liệu mẫu (bỏ qua header)
            sample_rows = min(3, actual_max_row - 1)
            for row_num in range(2, 2 + sample_rows):
                row_data = [str(ws.cell(row_num, col).value) if ws.cell(row_num, col).value else "" 
                           for col in range(1, max_col + 1)]
                analysis["sample_data"].append(row_data)
        
        wb.close()
        return analysis
    
    def create_analysis_report(self, analysis_results):
        """
        Tạo báo cáo tổng hợp phân tích
        """
        report_file = self.daily_output_dir / "analysis_report.txt"
        
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("📊 EXCEL FILES ANALYSIS REPORT\n")
            f.write(f"{'='*50}\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Directory: {self.daily_output_dir}\n")
            f.write(f"Total Files: {len(analysis_results)}\n\n")
            
            for file_name, analysis in analysis_results.items():
                f.write(f"📄 FILE: {file_name}\n")
                f.write(f"{'-'*30}\n")
                
                if "error" in analysis:
                    f.write(f"❌ Error: {analysis['error']}\n")
                else:
                    f.write(f"Size: {analysis['file_size']:,} bytes\n")
                    f.write(f"Sheets: {', '.join(analysis['sheets'])}\n")
                    f.write(f"Dimensions: {analysis['total_rows']} rows × {analysis['total_cols']} columns\n")
                    f.write(f"Has Data: {'Yes' if analysis['has_data'] else 'No'}\n")
                    
                    if analysis['headers']:
                        f.write(f"Headers: {', '.join(analysis['headers'][:5])}...\n")
                    
                    if analysis['sample_data']:
                        f.write("Sample Data:\n")
                        for i, row in enumerate(analysis['sample_data'][:2]):
                            f.write(f"  Row {i+2}: {', '.join(row[:3])}...\n")
                
                f.write("\n")
        
        return report_file

    def process_excel_with_advanced_filtering(self, excel_file):
        """
        Xử lý Excel với filtering nâng cao:
        1. Giữ nguyên 5 dòng tiêu đề
        2. Ẩn các dòng thỏa mãn điều kiện (C có dữ liệu và D trống)
        3. Ẩn cột A-F, M-N và từ S trở đi
        4. Bỏ thuộc tính wrap text từ dòng 6 trở đi
        5. Tự động điều chỉnh độ rộng cột
        """
        try:
            wb = openpyxl.load_workbook(excel_file)
            ws = wb.active
            
            row_count = ws.max_row
            col_count = ws.max_column
            
            print(f"   🔄 Xử lý nâng cao: {row_count} dòng × {col_count} cột")
            
            # Bước 1: Phân tích dữ liệu cột C và D
            rows_to_hide = []
            visible_rows = []
            
            for row_num in range(6, row_count + 1):  # Bắt đầu từ dòng 6 (sau tiêu đề)
                cell_c = ws.cell(row_num, 3)  # Cột C
                cell_d = ws.cell(row_num, 4)  # Cột D
                
                c_has_data = cell_c.value is not None and str(cell_c.value).strip() != ""
                d_is_blank = cell_d.value is None or str(cell_d.value).strip() == ""
                
                # Điều kiện: Ẩn các dòng có C có dữ liệu VÀ D trống
                if c_has_data and d_is_blank:
                    rows_to_hide.append(row_num)
                else:
                    visible_rows.append(row_num)
            
            print(f"   📊 Phân tích: {len(visible_rows)} dòng hiển thị, {len(rows_to_hide)} dòng ẩn (C có dữ liệu AND D trống)")
            
            # Bước 2: Ẩn các dòng thỏa mãn điều kiện
            for row_num in rows_to_hide:
                ws.row_dimensions[row_num].hidden = True
            
            # Bước 3: Ẩn các cột A-F, M-N và từ S trở đi
            columns_to_hide = ['A', 'B', 'C', 'D', 'E', 'F', 'M', 'N']
            
            # Thêm các cột từ S trở đi (S=19, T=20, ...)
            for col_num in range(19, col_count + 1):  # S=19 trở đi
                col_letter = openpyxl.utils.get_column_letter(col_num)
                columns_to_hide.append(col_letter)
            
            for col_letter in columns_to_hide:
                if col_letter <= openpyxl.utils.get_column_letter(col_count):
                    ws.column_dimensions[col_letter].hidden = True
            
            # Bước 4: Bỏ wrap text từ dòng 6 trở đi
            print(f"   📝 Bỏ wrap text từ dòng 6 đến {row_count}...")
            for row_num in range(6, row_count + 1):
                for col_num in range(1, col_count + 1):
                    cell = ws.cell(row_num, col_num)
                    if cell.alignment and cell.alignment.wrap_text:
                        cell.alignment = openpyxl.styles.Alignment(wrap_text=False)
            
            # Bước 5: Tự động điều chỉnh độ rộng cột cho các cột hiển thị
            visible_columns = []
            for col_num in range(1, col_count + 1):
                col_letter = openpyxl.utils.get_column_letter(col_num)
                if col_letter not in columns_to_hide:
                    visible_columns.append(col_letter)
            
            print(f"   📏 Điều chỉnh độ rộng cho {len(visible_columns)} cột hiển thị...")
            for col_letter in visible_columns:
                col_num = openpyxl.utils.column_index_from_string(col_letter)
                max_length = 0
                
                # Tìm độ dài tối đa của nội dung trong cột
                for row_num in range(1, row_count + 1):
                    cell_value = ws.cell(row_num, col_num).value
                    if cell_value:
                        cell_length = len(str(cell_value))
                        if cell_length > max_length:
                            max_length = cell_length
                
                # Đặt độ rộng cột (tối thiểu 8, tối đa 50)
                adjusted_width = min(max(max_length + 2, 8), 50)
                ws.column_dimensions[col_letter].width = adjusted_width
            
            print(f"   ✅ Ẩn {len(columns_to_hide)} cột, hiển thị {len(visible_columns)} cột")
            print(f"   ✅ Bỏ wrap text và điều chỉnh độ rộng cột hoàn thành")
            
            return True
            
        except Exception as e:
            print(f"   ❌ Lỗi xử lý nâng cao: {e}")
            return False

    def process_excel_with_advanced_filtering_return_sheet(self, excel_file):
        """
        Xử lý Excel với filtering nâng cao và trả về worksheet đã xử lý
        1. Giữ nguyên 5 dòng tiêu đề
        2. Ẩn các dòng thỏa mãn điều kiện (C có dữ liệu và D trống)
        3. Ẩn cột A-F, M-N và từ S trở đi
        4. Bỏ wrap text từ dòng 6 trở đi
        5. Tự động điều chỉnh độ rộng cột
        """
        try:
            wb = openpyxl.load_workbook(excel_file)
            ws = wb.active
            
            row_count = ws.max_row
            col_count = ws.max_column
            
            # Bước 1: Phân tích dữ liệu cột C và D
            rows_to_hide = []
            visible_rows = []
            
            for row_num in range(6, row_count + 1):  # Bắt đầu từ dòng 6 (sau tiêu đề)
                cell_c = ws.cell(row_num, 3)  # Cột C
                cell_d = ws.cell(row_num, 4)  # Cột D
                
                c_has_data = cell_c.value is not None and str(cell_c.value).strip() != ""
                d_is_blank = cell_d.value is None or str(cell_d.value).strip() == ""
                
                # Điều kiện: Ẩn các dòng có C có dữ liệu VÀ D trống
                if c_has_data and d_is_blank:
                    rows_to_hide.append(row_num)
                else:
                    visible_rows.append(row_num)
            
            # Bước 2: Ẩn các dòng thỏa mãn điều kiện
            for row_num in rows_to_hide:
                ws.row_dimensions[row_num].hidden = True
            
            # Bước 3: Ẩn các cột A-F, M-N và từ S trở đi
            columns_to_hide = ['A', 'B', 'C', 'D', 'E', 'F', 'M', 'N']
            
            # Thêm các cột từ S trở đi (S=19, T=20, ...)
            for col_num in range(19, col_count + 1):  # S=19 trở đi
                col_letter = openpyxl.utils.get_column_letter(col_num)
                columns_to_hide.append(col_letter)
            
            for col_letter in columns_to_hide:
                if col_letter <= openpyxl.utils.get_column_letter(col_count):
                    ws.column_dimensions[col_letter].hidden = True
            
            # Bước 4: Bỏ wrap text từ dòng 6 trở đi
            for row_num in range(6, row_count + 1):
                for col_num in range(1, col_count + 1):
                    cell = ws.cell(row_num, col_num)
                    if cell.alignment and cell.alignment.wrap_text:
                        cell.alignment = openpyxl.styles.Alignment(wrap_text=False)
            
            # Bước 5: Tự động điều chỉnh độ rộng cột cho các cột hiển thị
            visible_columns = []
            for col_num in range(1, col_count + 1):
                col_letter = openpyxl.utils.get_column_letter(col_num)
                if col_letter not in columns_to_hide:
                    visible_columns.append(col_letter)
            
            for col_letter in visible_columns:
                col_num = openpyxl.utils.column_index_from_string(col_letter)
                max_length = 0
                
                # Tìm độ dài tối đa của nội dung trong cột
                for row_num in range(1, row_count + 1):
                    cell_value = ws.cell(row_num, col_num).value
                    if cell_value:
                        cell_length = len(str(cell_value))
                        if cell_length > max_length:
                            max_length = cell_length
                
                # Đặt độ rộng cột (tối thiểu 8, tối đa 50)
                adjusted_width = min(max(max_length + 2, 8), 50)
                ws.column_dimensions[col_letter].width = adjusted_width
            
            return ws  # Trả về worksheet đã xử lý
            
        except Exception as e:
            return None

    def create_consolidated_result_file(self, processed_files):
        """
        Tạo file Kết quả.xlsx gộp tất cả các sheet đã xử lý
        """
        if not processed_files:
            print("❌ Không có file nào để gộp!")
            return False
        
        # Tạo workbook mới cho kết quả
        result_wb = openpyxl.Workbook()
        result_wb.remove(result_wb.active)  # Xóa sheet mặc định
        
        for i, (original_file, sheet_data) in enumerate(processed_files, 1):
            try:
                # Tạo tên sheet từ tên file gốc (bỏ .xlsx)
                sheet_name = original_file.stem
                # Đảm bảo tên sheet không quá dài (Excel giới hạn 31 ký tự)
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                
                # Tạo sheet mới
                ws_result = result_wb.create_sheet(title=sheet_name)
                
                # Copy dữ liệu từ sheet_data vào sheet mới
                self.copy_worksheet_data(sheet_data, ws_result)
                
                print(f"   ✅ Thêm Sheet {i}: {sheet_name} thành công")
                
            except Exception as e:
                print(f"   ❌ Thêm Sheet {i}: {original_file.stem} thất bại")
        
        # Lưu file kết quả
        result_path = self.daily_output_dir / "Kết quả.xlsx"
        result_wb.save(result_path)
        result_wb.close()
        
        return result_path
    
    def copy_worksheet_data(self, source_ws, target_ws):
        """
        Copy dữ liệu từ worksheet nguồn sang worksheet đích kèm theo formatting
        """
        # Copy tất cả dữ liệu và formatting
        for row in source_ws.iter_rows():
            for cell in row:
                target_cell = target_ws.cell(row=cell.row, column=cell.column)
                
                # Copy value
                target_cell.value = cell.value
                
                # Copy formatting nếu có
                if cell.has_style:
                    # Copy font
                    if cell.font:
                        target_cell.font = openpyxl.styles.Font(
                            name=cell.font.name,
                            size=cell.font.size,
                            bold=cell.font.bold,
                            italic=cell.font.italic,
                            vertAlign=cell.font.vertAlign,
                            underline=cell.font.underline,
                            strike=cell.font.strike,
                            color=cell.font.color
                        )
                    
                    # Copy fill (background color)
                    if cell.fill:
                        target_cell.fill = openpyxl.styles.PatternFill(
                            fill_type=cell.fill.fill_type,
                            start_color=cell.fill.start_color,
                            end_color=cell.fill.end_color
                        )
                    
                    # Copy border
                    if cell.border:
                        target_cell.border = openpyxl.styles.Border(
                            left=cell.border.left,
                            right=cell.border.right,
                            top=cell.border.top,
                            bottom=cell.border.bottom,
                            diagonal=cell.border.diagonal,
                            diagonal_direction=cell.border.diagonal_direction,
                            outline=cell.border.outline,
                            vertical=cell.border.vertical,
                            horizontal=cell.border.horizontal
                        )
                    
                    # Copy alignment (nhưng sẽ override wrap_text sau)
                    if cell.alignment:
                        target_cell.alignment = openpyxl.styles.Alignment(
                            horizontal=cell.alignment.horizontal,
                            vertical=cell.alignment.vertical,
                            text_rotation=cell.alignment.text_rotation,
                            wrap_text=cell.alignment.wrap_text,
                            shrink_to_fit=cell.alignment.shrink_to_fit,
                            indent=cell.alignment.indent
                        )
                    
                    # Copy number format
                    if cell.number_format:
                        target_cell.number_format = cell.number_format
        
        # Copy row heights trước khi áp dụng hidden rows
        for row_num in range(1, source_ws.max_row + 1):
            if source_ws.row_dimensions[row_num].height:
                target_ws.row_dimensions[row_num].height = source_ws.row_dimensions[row_num].height
        
        # Copy hidden rows từ source (đã được xử lý)
        for row_num in range(1, source_ws.max_row + 1):
            if source_ws.row_dimensions[row_num].hidden:
                target_ws.row_dimensions[row_num].hidden = True
        
        # Bước quan trọng: Áp dụng lại logic ẩn cột và điều chỉnh độ rộng
        col_count = source_ws.max_column
        row_count = source_ws.max_row
        
        # Ẩn các cột A-F, M-N và từ S trở đi (áp dụng lại logic)
        columns_to_hide = ['A', 'B', 'C', 'D', 'E', 'F', 'M', 'N']
        
        # Thêm các cột từ S trở đi (S=19, T=20, U=21, ..., Z=26)
        for col_num in range(19, col_count + 1):  # S=19 trở đi
            col_letter = openpyxl.utils.get_column_letter(col_num)
            columns_to_hide.append(col_letter)
        
        # Ẩn các cột - FIX: Phải ẩn tất cả cột có trong danh sách
        for col_letter in columns_to_hide:
            target_ws.column_dimensions[col_letter].hidden = True
        
        # Bỏ wrap text từ dòng 6 trở đi và THÊM canh giữa theo chiều dọc
        for row_num in range(6, row_count + 1):
            for col_num in range(1, col_count + 1):
                cell = target_ws.cell(row_num, col_num)
                if cell.alignment:
                    # Giữ nguyên các thuộc tính alignment khác, bỏ wrap_text và THÊM vertical center
                    target_ws.cell(row_num, col_num).alignment = openpyxl.styles.Alignment(
                        horizontal=cell.alignment.horizontal,
                        vertical='center',  # CANH GIỮA THEO CHIỀU DỌC
                        text_rotation=cell.alignment.text_rotation,
                        wrap_text=False,  # Bỏ wrap text
                        shrink_to_fit=cell.alignment.shrink_to_fit,
                        indent=cell.alignment.indent
                    )
                else:
                    # Nếu chưa có alignment, tạo mới với vertical center
                    target_ws.cell(row_num, col_num).alignment = openpyxl.styles.Alignment(
                        vertical='center',
                        wrap_text=False
                    )
        
        # THÊM: Canh giữa cho tất cả các merged cells (bao gồm cả tiêu đề)
        for merged_range in target_ws.merged_cells.ranges:
            # Lấy cell đầu tiên của merged range
            start_cell = target_ws.cell(merged_range.min_row, merged_range.min_col)
            start_cell.alignment = openpyxl.styles.Alignment(
                horizontal='center',  # Canh giữa ngang
                vertical='center',    # Canh giữa dọc
                wrap_text=False
            )
        
        # Tự động điều chỉnh độ rộng cột cho các cột hiển thị
        visible_columns = []
        for col_num in range(1, col_count + 1):
            col_letter = openpyxl.utils.get_column_letter(col_num)
            if col_letter not in columns_to_hide:
                visible_columns.append(col_letter)
        
        # Điều chỉnh độ rộng cột dựa trên nội dung thực tế
        for col_letter in visible_columns:
            col_num = openpyxl.utils.column_index_from_string(col_letter)
            max_length = 0
            
            # Tìm độ dài tối đa của nội dung trong cột (chỉ tính các dòng hiển thị)
            for row_num in range(1, row_count + 1):
                # Bỏ qua các dòng bị ẩn
                if target_ws.row_dimensions[row_num].hidden:
                    continue
                    
                cell_value = target_ws.cell(row_num, col_num).value
                if cell_value:
                    # Tính độ dài hiển thị thực tế (có thể có font size khác nhau)
                    display_length = len(str(cell_value))
                    
                    # Điều chỉnh theo font size nếu có
                    cell = target_ws.cell(row_num, col_num)
                    if cell.font and cell.font.size:
                        # Font size lớn hơn thì cần width lớn hơn
                        size_factor = cell.font.size / 11  # 11 là font size chuẩn
                        display_length = int(display_length * size_factor)
                    
                    if display_length > max_length:
                        max_length = display_length
            
            # Đặt độ rộng cột (tối thiểu 8, tối đa 40, và thêm padding vừa phải)
            if max_length == 0:
                adjusted_width = 10  # Default width cho cột trống
            else:
                # Tính width tối ưu: nội dung + padding nhỏ
                adjusted_width = min(max(max_length + 1, 8), 40)  # Giảm padding từ +2 xuống +1
            
            target_ws.column_dimensions[col_letter].width = adjusted_width
        
        # Copy merged cells
        for merged_range in source_ws.merged_cells.ranges:
            target_ws.merge_cells(str(merged_range))

def main():
    """
    Hàm main chạy chương trình
    """
    # Khởi tạo OrderChecker
    checker = OrderChecker()
    
    # Kiểm tra config trước khi chạy
    if not checker.config:
        print("❌ Config không hợp lệ!")
        return
    
    # Chạy automation với thông tin từ config
    success = checker.run_browser_test()
    
    print("\n" + "─" * 60)
    if success:
        print("🎉 HOÀN THÀNH TẤT CẢ!")
    else:
        print("❌ THẤT BẠI!")
    print("─" * 60)

if __name__ == "__main__":
    main()