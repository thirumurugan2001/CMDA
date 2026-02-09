import os
import sys
import requests
from io import BytesIO
from pathlib import Path
from datetime import datetime
from urllib.parse import urljoin
from Integration import lead_import
from pdf_report import generate_pdf_report
from playwright.sync_api import sync_playwright
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QMovie, QColor, QPalette, QPixmap
from approved_letter import extract_registered_architect_from_bytes
from extractor import extract_text_from_pdf_bytesio, extract_fields, export_to_xlsx
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QVBoxLayout, QComboBox,QMessageBox, QLabel, QGroupBox, QHBoxLayout, QSizePolicy, QSpacerItem, QProgressBar, QFileDialog)

def setup_playwright_path():
    if getattr(sys, 'frozen', False):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))        
        browser_path = os.path.join(base_path, 'ms-playwright')
        os.environ['PLAYWRIGHT_BROWSERS_PATH'] = browser_path        
        print(f"Playwright browser path set to: {browser_path}")
        if os.path.exists(browser_path):
            print("✅ Browser path exists!")
            for root, dirs, files in os.walk(browser_path):
                for dir in dirs:
                    print(f"  - Found: {dir}")
        else:
            print("❌ Warning: Browser path does not exist!")            
        return browser_path
    else:
        return None

setup_playwright_path()

class ScrapeWorker(QThread):
    progress = pyqtSignal(int, int)
    finished = pyqtSignal(list, tuple)
    error = pyqtSignal(str)    
    def __init__(self, year, entries):
        super().__init__()
        self.year = year
        self.entries = entries
        self.total_attempted = 0
        self.successful_scraped = 0
        self.failed_scraped = 0
        self.failed_file_numbers = []
    def run(self):
        url = f"https://cmdachennai.gov.in/OnlinePPAApprovalDetails/{self.year}.html"
        entry_value = {'10': '10', '25': '25', '50': '50', 'All': '-1'}.get(self.entries, '10')
        pdf_streams = []
        urls = []
        approved_plan_links = []
        architect_details = []        
        try:
            browser = None
            playwright_instance = None            
            try:
                playwright_instance = sync_playwright().start()                
                if getattr(sys, 'frozen', False):
                    base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
                    chromium_path = os.path.join(base_path, 'ms-playwright', 'chromium-1187', 'chrome-win', 'chrome.exe')
                    if os.path.exists(chromium_path):
                        print(f"🚀 Launching browser with explicit path: {chromium_path}")
                        browser = playwright_instance.chromium.launch(
                            executable_path=chromium_path,
                            headless=True
                        )
                    else:
                        print("⚠️ Browser not found at explicit path, trying default...")
                        browser = playwright_instance.chromium.launch(headless=True)
                else:
                    browser = playwright_instance.chromium.launch(headless=True)
            except Exception as browser_error:
                print(f"❌ Browser launch failed: {browser_error}")
                if not playwright_instance:
                    playwright_instance = sync_playwright().start()
                browser = playwright_instance.chromium.launch(headless=True)            
            if not browser:
                self.error.emit("Failed to launch browser")
                return       
            page = browser.new_page()
            page.goto(url, timeout=60000)
            page.select_option('select[name="DataTables_Table_0_length"]', value=entry_value)
            page.wait_for_timeout(3000)            
            page.wait_for_selector('table tbody tr', timeout=30000)            
            links = page.locator('table tbody tr td:nth-child(9) a')
            approved_links = page.locator('table tbody tr td:nth-child(7) a')
            approved_letter = page.locator('table tbody tr td:nth-child(6) a')
            file_no_cells = page.locator('table tbody tr td:nth-child(2)')            
            total = links.count()
            self.total_attempted = total
            approved_letter_links = []            
            for i in range(total):
                try:
                    href = links.nth(i).get_attribute("href")
                    approved_href = approved_links.nth(i).get_attribute("href")
                    letter_href = approved_letter.nth(i).get_attribute("href")
                    file_no_text = file_no_cells.nth(i).inner_text() if file_no_cells.count() > i else f"Unknown_{i+1}"
                    if href and href.lower().endswith(".pdf"):
                        full_url = urljoin("https://cmdachennai.gov.in/", href)
                        approved_url = urljoin("https://cmdachennai.gov.in/", approved_href) if approved_href else ""
                        letter_url = urljoin("https://cmdachennai.gov.in/", letter_href) if letter_href else ""
                        try:
                            r = requests.get(full_url, timeout=30)
                            r.raise_for_status()
                            pdf_io = BytesIO(r.content)
                            text = extract_text_from_pdf_bytesio(pdf_io)
                            fields = extract_fields(text)                            
                            extracted_file_no = fields.get("File No.", file_no_text)                            
                            if fields.get("File No.") in ["Not Found", "Error", ""] or not fields.get("File No."):
                                self.failed_scraped += 1
                                self.failed_file_numbers.append(extracted_file_no)
                            else:
                                self.successful_scraped += 1
                            pdf_streams.append(fields)
                            urls.append(full_url)
                            approved_plan_links.append(approved_url)
                            approved_letter_links.append(letter_url)                            
                            if letter_url:
                                try:
                                    r_letter = requests.get(letter_url, timeout=30)
                                    r_letter.raise_for_status()
                                    letter_io = BytesIO(r_letter.content)
                                    architect_info = extract_registered_architect_from_bytes(letter_io)
                                    architect_details.append(architect_info)
                                except Exception as e:
                                    architect_details.append({"error": str(e), "status": "failure"})
                        except requests.exceptions.RequestException as e:
                            self.failed_scraped += 1
                            self.failed_file_numbers.append(file_no_text)
                            pdf_streams.append({
                                "File No.": file_no_text,
                                "Planning Permission No.": "Failed",
                                "Permit No.": "Failed",
                                "Date of permit": "Failed",
                                "Date of Application": "Failed",
                                "Mobile No.": "Failed",
                                "Email ID": "Failed",
                                "Applicant Name": "Failed",
                                "Applicant Address": "Failed",
                                "Nature of Development": "Failed",
                                "Dwelling Unit Info": "",
                                "Site Address": "Failed",
                                "Area Name": "Failed"
                            })
                            urls.append(full_url)
                            approved_plan_links.append(approved_url)
                            approved_letter_links.append(letter_url)
                            architect_details.append({"error": str(e), "status": "failure"})
                        except Exception as e:
                            print(f"⚠️ Processing failed: {e}")
                            self.failed_scraped += 1
                            self.failed_file_numbers.append(file_no_text)                
                except Exception as e:
                    print(f"⚠️ Error processing row {i+1}: {e}")
                    self.failed_scraped += 1
                    self.failed_file_numbers.append(f"Row_{i+1}")                
                self.progress.emit(i + 1, total)            
            browser.close()
            if playwright_instance:
                playwright_instance.stop()
            scraping_stats = {
                'total_attempted': self.total_attempted,
                'successful_scraped': self.successful_scraped,
                'failed_scraped': self.failed_scraped,
                'failed_file_numbers': self.failed_file_numbers
            }
            if self.failed_file_numbers:
                print(f"  Failed File Numbers: {', '.join(self.failed_file_numbers[:5])}")
                if len(self.failed_file_numbers) > 5:
                    print(f"    ... and {len(self.failed_file_numbers) - 5} more")
            self.finished.emit(pdf_streams, (urls, approved_plan_links, approved_letter_links, architect_details, scraping_stats))
        except Exception as e:
            print(f"❌ Scraping error: {e}")
            self.error.emit(str(e))
            if playwright_instance:
                playwright_instance.stop()

class ScraperApp(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("CMDA Plan Permit Scraper - Ajantha Bathroom Products")
        self.setMinimumSize(1200, 800)
        self.selected_year = None
        self.selected_entries = '10'
        self.temp_file_path = None 
        self.local_file_path = None
        self.scraping_stats = None
        self.crm_import_result = None
        self.setup_ui()
    
    def setup_ui(self):
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(40, 20, 40, 20)
        self.layout.setSpacing(20)        
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor("#ffffff"))
        palette.setColor(QPalette.WindowText, QColor("#000000"))
        palette.setColor(QPalette.Base, QColor("#ffffff"))
        palette.setColor(QPalette.AlternateBase, QColor("#f8f9fa"))
        palette.setColor(QPalette.Text, QColor("#000000"))
        self.setPalette(palette)
        self.setAutoFillBackground(True)        
        self.setStyleSheet("""
            QWidget {
                font-family: 'Segoe UI', 'Inter', sans-serif;
                font-size: 14px;
                color: #000000;
                background-color: #ffffff;
            }            
            QLabel#TitleLabel {
                color: #dc2626;
                font-size: 36px;
                font-weight: 800;
                margin-bottom: 8px;
                letter-spacing: -0.5px;
                text-align: center;
            }            
            QLabel#SubTitleLabel {
                color: #404040;
                font-size: 16px;
                font-weight: 400;
                margin-bottom: 24px;
                text-align: center;
            }            
            QLabel#ClientLabel {
                color: #dc2626;
                font-size: 18px;
                font-weight: 700;
                margin-bottom: 8px;
                text-align: center;
            }            
            QLabel#AutomationLabel {
                color: #666666;
                font-size: 14px;
                font-weight: 400;
                margin-bottom: 20px;
                text-align: center;
            }           
            QGroupBox {
                border: 2px solid #dc2626;
                border-radius: 10px;
                padding: 15px;
                background-color: #ffffff;
                margin-top: 8px;
                font-weight: 700;
                color: #000000;
            }            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px 0 8px;
                color: #dc2626;
                font-weight: 700;
                font-size: 14px;
                background-color: #ffffff;
            }            
            QComboBox {
                background-color: #ffffff;
                border: 2px solid #dc2626;
                border-radius: 6px;
                padding: 8px 12px;
                font-weight: 600;
                min-width: 100px;
                color: #000000;
                font-size: 13px;
                min-height: 20px;
            }            
            QComboBox:focus {
                border-color: #b91c1c;
                outline: none;
            }            
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #dc2626;
                border-radius: 0px;
            }            
            QComboBox::down-arrow {
                image: url(down_arrow.png);
                width: 12px;
                height: 12px;
            }            
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                border: 2px solid #dc2626;
                color: #000000;
                selection-background-color: #dc2626;
                selection-color: #ffffff;
                outline: none;
                padding: 4px;
            }            
            QComboBox QAbstractItemView::item {
                padding: 6px 8px;
                border-radius: 4px;
            }            
            QComboBox QAbstractItemView::item:selected {
                background-color: #dc2626;
                color: #ffffff;
            }            
            QPushButton {
                background-color: #dc2626;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 24px;
                font-weight: 700;
                font-size: 14px;
                letter-spacing: 0.3px;
                min-height: 20px;
            }            
            QPushButton:hover {
                background-color: #b91c1c;
            }            
            QPushButton:pressed {
                background-color: #991b1b;
            }            
            QPushButton:disabled {
                background-color: #d1d5db;
                color: #6b7280;
            }            
            QPushButton#YearButton {
                background-color: #ffffff;
                color: #000000;
                border: 2px solid #dc2626;
                font-weight: 700;
                padding: 8px 16px;
            }            
            QPushButton#YearButton:hover {
                background-color: #fef2f2;
                border-color: #b91c1c;
            }            
            QPushButton#YearButton:checked {
                background-color: #dc2626;
                color: white;
                border-color: #dc2626;
            }            
            QProgressBar {
                border: 2px solid #dc2626;
                border-radius: 8px;
                background-color: #ffffff;
                text-align: center;
                color: #000000;
                font-weight: 600;
                height: 24px;
                font-size: 12px;
            }            
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #dc2626, stop:1 #b91c1c);
                border-radius: 6px;
            }""")
        
        header_layout = QVBoxLayout()
        header_layout.setSpacing(12)
        header_layout.setAlignment(Qt.AlignCenter)        
        self.logo_label = QLabel()
        logo_path = str(Path(__file__).resolve().parent / "client_logo.png")
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            scaled_pixmap = pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.logo_label.setPixmap(scaled_pixmap)
        else:
            self.logo_label.setText("🏭")
            self.logo_label.setStyleSheet("font-size: 50px; color: #dc2626;")
        self.logo_label.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(self.logo_label)        
        client_label = QLabel("Ajantha Bathroom Products and Pipes Pvt.Ltd")
        client_label.setObjectName("ClientLabel")
        client_label.setAlignment(Qt.AlignCenter)
        client_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        header_layout.addWidget(client_label)
        separator = QLabel()
        separator.setFixedHeight(2)
        separator.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #dc2626, stop:0.5 #000000, stop:1 #dc2626); margin: 15px 0px;")
        header_layout.addWidget(separator)        
        title = QLabel("CMDA PLAN PERMIT SCRAPER")
        title.setObjectName("TitleLabel")
        title.setAlignment(Qt.AlignCenter)
        title.setFont(QFont("Segoe UI", 28, QFont.Black))
        header_layout.addWidget(title)        
        subtitle = QLabel("Extract digitally generated Planning Permits from CMDA - Zoho Automation and CRM integration")
        subtitle.setObjectName("SubTitleLabel")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setFont(QFont("Segoe UI", 13))
        header_layout.addWidget(subtitle)        
        self.layout.addLayout(header_layout)        
        self.selected_year_label = QLabel("Year: Not Selected")
        self.selected_year_label.setAlignment(Qt.AlignCenter)
        self.selected_year_label.setFont(QFont("Segoe UI", 18, QFont.Black))
        self.selected_year_label.setStyleSheet("color: #dc2626; margin: 15px 0; background: transparent; font-weight: 800;")
        self.layout.addWidget(self.selected_year_label)        
        year_group = QGroupBox("Choose Year")
        year_layout = QHBoxLayout()
        year_layout.setSpacing(12)
        year_layout.setAlignment(Qt.AlignCenter)        
        for year in ['2026','2025', '2024', '2023', '2022']:
            btn = QPushButton(year)
            btn.setObjectName("YearButton")
            btn.setFixedSize(90, 40)
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked, y=year: self.on_year_selected(y) if checked else None)
            year_layout.addWidget(btn)        
        year_group.setLayout(year_layout)
        self.layout.addWidget(year_group)        
        self.entry_group = QGroupBox("Select Entry Count")
        self.entry_group.setVisible(False)
        self.entry_group.setMaximumHeight(100)
        entry_layout = QVBoxLayout()
        entry_layout.setContentsMargins(10, 5, 10, 5)
        entry_layout.setSpacing(8)
        entry_layout.setAlignment(Qt.AlignCenter)        
        self.entry_dropdown = QComboBox()
        self.entry_dropdown.addItems(['10', '25', '50', 'All'])
        self.entry_dropdown.setFixedSize(120, 35)
        self.entry_dropdown.currentIndexChanged.connect(self.update_dropdown_selection)
        entry_layout.addWidget(self.entry_dropdown)        
        self.entry_group.setLayout(entry_layout)
        self.layout.addWidget(self.entry_group)        
        self.row_count_label = QLabel("")
        self.row_count_label.setVisible(False)
        self.row_count_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.row_count_label.setAlignment(Qt.AlignCenter)
        self.row_count_label.setStyleSheet("color: #000000; font-weight: 700; background: transparent; padding: 8px;")
        self.layout.addWidget(self.row_count_label)        
        self.loader = QLabel()
        self.loader.setAlignment(Qt.AlignCenter)
        self.loader.setMinimumHeight(80)
        gif_path = str(Path(__file__).resolve().parent / "loader.gif")
        if os.path.exists(gif_path):
            self.loader_movie = QMovie(gif_path)
            self.loader.setMovie(self.loader_movie)
        self.loader.setVisible(False)
        self.layout.addWidget(self.loader)        
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setFormat("Scraping: %p% (%v of %m)")
        self.layout.addWidget(self.progress)        
        self.layout.addItem(QSpacerItem(20, 30, QSizePolicy.Minimum, QSizePolicy.Expanding))        
        self.scrape_btn = QPushButton("Start Scraping")
        self.scrape_btn.setVisible(False)
        self.scrape_btn.setFixedSize(200, 50)
        self.scrape_btn.clicked.connect(self.scrape)
        self.layout.addWidget(self.scrape_btn, alignment=Qt.AlignCenter)        
        self.report_btn = QPushButton("Generate PDF Report")
        self.report_btn.setVisible(False)
        self.report_btn.setFixedSize(200, 50)
        self.report_btn.clicked.connect(self.generate_pdf_report)
        self.layout.addWidget(self.report_btn, alignment=Qt.AlignCenter)        
        footer_label = QLabel("© 2024 Ajantha Bathroom Products - AI Powered Zoho CRM Automation")
        footer_label.setAlignment(Qt.AlignCenter)
        footer_label.setStyleSheet("color: #666666; font-size: 11px; margin-top: 15px; background: transparent; font-weight: 600;")
        self.layout.addWidget(footer_label)        
        self.setLayout(self.layout)    
    
    def on_year_selected(self, selected_year):
        for i in range(self.layout.count()):
            item = self.layout.itemAt(i)
            if isinstance(item.widget(), QGroupBox) and item.widget().title() == "Choose Year":
                year_group = item.widget()
                for j in range(year_group.layout().count()):
                    widget = year_group.layout().itemAt(j).widget()
                    if isinstance(widget, QPushButton) and widget.text() != selected_year:
                        widget.setChecked(False)
        self.load_year(selected_year)
    
    def load_year(self, year):
        self.selected_year = year
        self.selected_year_label.setText(f"Selected Year: {year}")
        self.selected_entries = self.entry_dropdown.currentText()
        self.row_count_label.setVisible(False)
        success = self.update_row_count()
        self.entry_group.setVisible(success)
        self.scrape_btn.setVisible(success)
        self.report_btn.setVisible(False)
    
    def update_dropdown_selection(self):
        self.selected_entries = self.entry_dropdown.currentText()
        self.update_row_count()
    
    def update_row_count(self):
        if not self.selected_year:
            return False
        url = f"https://cmdachennai.gov.in/OnlinePPAApprovalDetails/{self.selected_year}.html"
        entry_value = {'10': '10', '25': '25', '50': '50', 'All': '-1'}.get(self.selected_entries, '10')
        self.loader.setVisible(True)
        if hasattr(self, 'loader_movie'):
            self.loader_movie.start()
        QApplication.processEvents()        
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page()
                page.goto(url, timeout=60000)
                page.select_option('select[name="DataTables_Table_0_length"]', value=entry_value)
                page.wait_for_timeout(2000)
                rows = page.locator('table tbody tr').all_inner_texts()
                self.row_count_label.setText(f"<span style='color: #000000;'> Total entries found: <b>{len(rows)}</b></span>")
                self.row_count_label.setVisible(True)
                browser.close()
                return True
        except Exception as e:
            self.row_count_label.setText("<span style='color: #000000;'> Error fetching entry count.</span>")
            print(f"Error: {e}")
            return False
        finally:
            if hasattr(self, 'loader_movie'):
                self.loader_movie.stop()
            self.loader.setVisible(False)
    
    def scrape(self):
        self.scrape_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.loader.setVisible(True)
        if hasattr(self, 'loader_movie'):
            self.loader_movie.start()        
        self.worker = ScrapeWorker(self.selected_year, self.selected_entries)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_scrape_finished)
        self.worker.error.connect(self.on_scrape_error)
        self.worker.start()
    
    def update_progress(self, current, total):
        self.progress.setMaximum(total)
        self.progress.setValue(current)
    
    def on_scrape_finished(self, pdfs, urls):
        urls, approved_links, approved_letter, architect_details, scraping_stats = urls
        if hasattr(self, 'loader_movie'):
            self.loader_movie.stop()
        self.loader.setVisible(False)
        self.scrape_btn.setEnabled(True)
        self.progress.setVisible(False)        
        self.scraping_stats = scraping_stats        
        temp_path, download_path = export_to_xlsx(
            pdfs, self.selected_year, urls, approved_links, approved_letter, architect_details
        )        
        self.temp_file_path = temp_path
        self.local_file_path = download_path        
        import_result = lead_import(file_path=self.temp_file_path)
        self.crm_import_result = import_result        
        self.show_completion_message(len(pdfs), import_result, download_path)        
        self.report_btn.setVisible(True)
    
    def show_completion_message(self, pdf_count, import_result, download_path):
        success = import_result.get("status", False) if import_result else False        
        message = f"""
        <div style='font-family: Segoe UI; font-size: 14px; color: #000000; background: #ffffff; padding: 25px; border-radius: 12px; border: 3px solid #dc2626; max-width: 600px;'>
            <div style='text-align: center; margin-bottom: 20px;'>
                <h3 style='color: #dc2626; margin: 0; font-size: 24px; font-weight: 800;'>✅ SCRAPING COMPLETED</h3>
                <div style='color: #666666; font-size: 14px; margin-top: 5px;'>Data successfully extracted and processed</div>
            </div>
            
            <div style='background: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px;'>
                <div style='display: flex; justify-content: space-between; margin-bottom: 15px;'>
                    <div style='text-align: center; flex: 1;'>
                        <div style='font-size: 24px; font-weight: 800; color: #dc2626;'>{pdf_count}</div>
                        <div style='font-size: 12px; color: #666666;'>Documents Processed</div>
                    </div>
                    <div style='text-align: center; flex: 1;'>
                        <div style='font-size: 24px; font-weight: 800; color: #dc2626;'>{self.selected_year}</div>
                        <div style='font-size: 12px; color: #666666;'>Year Processed</div>
                    </div>
                    <div style='text-align: center; flex: 1;'>
                        <div style='font-size: 24px; font-weight: 800; color: {'#16a34a' if success else '#dc2626'};'>{'✓' if success else '✗'}</div>
                        <div style='font-size: 12px; color: #666666;'>CRM Import</div>
                    </div>
                </div>
            </div>
            
            <div style='background: #fef2f2; padding: 15px; border-radius: 8px; border-left: 4px solid #dc2626;'>
                <div style='font-weight: 700; color: #000000; margin-bottom: 8px;'>📁 File Saved Location:</div>
                <div style='color: #666666; font-family: Consolas, monospace; font-size: 12px; background: #ffffff; padding: 10px; border-radius: 4px; border: 1px solid #e5e7eb; word-break: break-all;'>
                    {download_path}
                </div>
            </div>
            
            <div style='margin-top: 20px; padding: 12px; background: #dc2626; border-radius: 6px; text-align: center;'>
                <span style='color: #ffffff; font-weight: 700; font-size: 13px;'>🚀 Ready for analysis in Zoho CRM</span>
            </div>
            
            <div style='margin-top: 20px; text-align: center;'>
                <span style='color: #666666; font-size: 12px;'>You can now generate a PDF report with detailed statistics.</span>
            </div>
        </div>
        """
        
        msg_box = QMessageBox()
        msg_box.setWindowTitle("Scraping Completed - Ajantha Automation")
        msg_box.setTextFormat(Qt.RichText)
        msg_box.setText(message)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()
    
    def on_scrape_error(self, message):
        if hasattr(self, 'loader_movie'):
            self.loader_movie.stop()
        self.loader.setVisible(False)
        self.scrape_btn.setEnabled(True)
        self.progress.setVisible(False)        
        error_message = f"""
        <div style='font-family: Segoe UI; font-size: 14px; color: #000000; background: #ffffff; padding: 25px; border-radius: 12px; border: 3px solid #dc2626; max-width: 500px;'>
            <div style='text-align: center; margin-bottom: 20px;'>
                <h3 style='color: #dc2626; margin: 0; font-size: 24px; font-weight: 800;'>❌ SCRAPING ERROR</h3>
            </div>
            
            <div style='background: #fef2f2; padding: 15px; border-radius: 8px; border-left: 4px solid #dc2626; margin-bottom: 15px;'>
                <div style='font-weight: 700; color: #000000; margin-bottom: 8px;'>Error Details:</div>
                <div style='color: #dc2626; font-weight: 600; font-size: 13px;'>{message}</div>
            </div>
            
            <div style='color: #666666; font-size: 12px; text-align: center; background: #f8f9fa; padding: 12px; border-radius: 6px;'>
                Please check your internet connection and try again.<br>
                Contact <strong>Ajantha IT Support</strong> if the problem persists.
            </div>
        </div>
        """        
        msg_box = QMessageBox()
        msg_box.setWindowTitle("Scraping Error - Ajantha Automation")
        msg_box.setTextFormat(Qt.RichText)
        msg_box.setText(error_message)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()
    
    def generate_pdf_report(self):
        if not self.scraping_stats or not self.crm_import_result:
            QMessageBox.warning(self, "No Data", "Please run scraping first to generate a report.")
            return        
        default_filename = f"CMDA_Report_{self.selected_year}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save PDF Report",
            str(Path.home() / "Downloads" / default_filename),
            "PDF Files (*.pdf)"
        )        
        if not file_path:
            return        
        try:
            success = generate_pdf_report(
                file_path=file_path,
                scraping_stats=self.scraping_stats,
                crm_result=self.crm_import_result,
                year=self.selected_year,
                local_file_path=self.local_file_path
            )            
            if success:
                QMessageBox.information(self,"Report Generated",f"PDF report has been successfully generated and saved to:\n{file_path}")
            else:
                QMessageBox.warning(self,"Report Generation Failed","Failed to generate PDF report. Please check the logs for details.")                
        except Exception as e:
            QMessageBox.critical(self,"Error",f"An error occurred while generating the PDF report:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScraperApp()
    window.showMaximized()
    sys.exit(app.exec_())