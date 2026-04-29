import sys
import os
import time
import threading
import re
from datetime import datetime
import requests
from urllib.parse import urlparse, parse_qs

import openpyxl
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QTextEdit, 
    QVBoxLayout, QGridLayout, QCheckBox, QComboBox, QFileDialog
)
from PyQt6.QtCore import QSettings, pyqtSignal, QUrl, Qt
from PyQt6.QtGui import QDesktopServices

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoAlertPresentException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager






JOB_COUNTER = 0

class MenuPolicyApp(QWidget):
    log_signal = pyqtSignal(str)
    update_sheet_label_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.settings = QSettings("MyCompany", "MenuPolicyApp")  # Lưu dữ liệu dưới tên ứng dụng
        # Initialize all sheet values
        self.value_A1 = ""
        self.value_name = ""
        self.value_address = ""
        self.value_phone = ""
        self.value_link_booking = ""
        self.value_email = ""
        self.value_link_facebook = ""
        self.value_link_ins = ""
        self.value_link_x = ""
        self.value_link_yelp = ""
        self.value_link_map = ""
        self.value_hours = ""
        self.value_link_rv_yelp = ""
        self.value_link_rv_map = ""
        self.value_keyword_B27 = ""
        self.value_AM_B23 = ""
        self.value_Assignee_C23 = ""
        self.initUI()
        self.load_settings()  # Tải dữ liệu khi khởi động
        self.driver = None

    def initUI(self):
        self.log_signal.connect(self.append_log)
        self.update_sheet_label_signal.connect(self.set_sheet_label)
        self.setWindowTitle("Menu Policy App")
        self.setGeometry(100, 100, 600, 400)
        self.is_logged_in = False
        # ✅ Tạo layout
        self.layout = QGridLayout()  
        self.setLayout(self.layout)  # ✅ Chỉ dùng self.layout, không dùng layout

        # ✅ Thêm QLabel để hiển thị dữ liệu từ Google Sheets
        self.sheet_data_label = QLabel("📌 No Data")  
        self.sheet_data_label.setOpenExternalLinks(True)
        self.layout.addWidget(self.sheet_data_label, 0, 0)  

        # ✅ Gọi hàm để lấy dữ liệu Google Sheets
        self.get_google_sheets_data()

        print(f"Type of self.layout: {type(self.layout)}")

    # Website Password New
        self.website_pass_new = QLineEdit(self)
        self.website_pass_new.setEchoMode(QLineEdit.EchoMode.Password)
        self.layout.addWidget(self.website_pass_new, 0, 2)


    # Configuration for sheets
        SHEET_CONFIG = {
            "Tony": "445860585",
            "Lucas": "1922397353",
            "Alex": "1127382190",
            "Allen": "208823018",
            "Hayden": "1606998811",
            "Leo": "730036757",
            "Davis": "443113123",
            "Backup": "1373613829"
        }

    # Google Sheet ID cố định
        self.spreadsheet_id = "1a_wPoit-UqeJGTLu7EyMTnIljllOD-LVJedtMnXmqn4"
        
    # Tạo combo box để chọn sheet/tab
        self.sheet_selector = QComboBox()
        self.sheet_selector.addItems(list(SHEET_CONFIG.keys()))
        self.sheet_selector.currentTextChanged.connect(self.get_google_sheets_data)
        self.layout.addWidget(self.sheet_selector, 0, 1)
        
    # Mapping giữa tab và gid
        self.sheet_gids = SHEET_CONFIG

        self.get_google_sheets_data()


    # License Menu User
        self.license_user_label = QLabel("License Menu User/Pass:")
        self.layout.addWidget(self.license_user_label, 1, 0)
        self.license_user_input = QLineEdit(self)
        self.layout.addWidget(self.license_user_input, 1, 1)
        
    # License Menu Password
        self.license_pass_input = QLineEdit(self)
        self.license_pass_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.layout.addWidget(self.license_pass_input, 1, 2)   

    # Website User
        self.website_user_label = QLabel("Website User/Pass:")
        self.layout.addWidget(self.website_user_label, 2, 0)
        self.website_user_input = QLineEdit(self)
        self.layout.addWidget(self.website_user_input, 2, 1)
        
    # Website Password

        self.website_pass_input = QLineEdit(self)
        self.website_pass_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.layout.addWidget(self.website_pass_input, 2, 2)

    # Email Input
        self.label_email = QLabel("Your Email:")
        self.layout.addWidget(self.label_email, 3, 0)
        self.email_input = QLineEdit(self)
        self.layout.addWidget(self.email_input,  3, 1)

    # Domain Input
        self.label = QLabel("Domain:")
        self.layout.addWidget(self.label, 4, 0)
        self.domain_input = QLineEdit(self)
        self.layout.addWidget(self.domain_input, 4, 1, 1, 2)

    # Button Show/Hide Password
        self.show_pass_button = QPushButton("👁 Show Password", self)
        self.show_pass_button.setCheckable(True)  # Cho phép bật/tắt
        self.show_pass_button.clicked.connect(self.toggle_password_visibility)
        self.layout.addWidget(self.show_pass_button, 3, 2)

    # Checkbox chọn chức năng
        self.website_text = QLabel("Select Function:==================")
        self.layout.addWidget(self.website_text, 10, 0, 1, 2)

        self.chk_license = QCheckBox("Active Menu")
        self.chk_license.setChecked(True)  # Mặc định chọn
        self.layout.addWidget(self.chk_license, 11, 0)

        self.chk_change_admin_pass = QCheckBox("Reset Pass Admin")
        self.layout.addWidget(self.chk_change_admin_pass, 11, 1)

        self.chk_change_user_pass = QCheckBox("Reset Pass Tech")
        self.layout.addWidget(self.chk_change_user_pass, 11, 2)

        self.chk_limit_login_settings = QCheckBox("Limit login")
        self.layout.addWidget(self.chk_limit_login_settings, 12, 0)

        self.chk_settings_webinfo = QCheckBox("Web Info")
        self.layout.addWidget(self.chk_settings_webinfo, 12, 1)

        self.chk_settings_name = QCheckBox("Change Name")
        self.layout.addWidget(self.chk_settings_name, 12, 2)

        self.chk_contact_form_settings = QCheckBox("Set CF7")
        self.layout.addWidget(self.chk_contact_form_settings, 13, 0)
        
        self.chk_add_pass_wpm = QCheckBox("Add WPM Pass")
        self.layout.addWidget(self.chk_add_pass_wpm, 13, 1)

    # Lưu thông tin
        self.save_button = QPushButton("💾 Save Info", self)
        self.save_button.clicked.connect(self.save_settings)
        self.layout.addWidget(self.save_button, 23, 0)

    # Run Button
        self.run_button = QPushButton("Run", self)
        self.run_button.clicked.connect(self.run_script)
        self.layout.addWidget(self.run_button, 23, 1)
        
    # Log Output
        
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.layout.addWidget(self.log_output, 24, 0, 1, 3)

    def log(self, message):
        self.log_signal.emit(message)

    def append_log(self, message):
        self.log_output.append(message)

    def set_sheet_label(self, message):
        self.sheet_data_label.setText(message)
        
    def toggle_password_visibility(self):
        if self.show_pass_button.isChecked():
            self.license_pass_input.setEchoMode(QLineEdit.EchoMode.Normal)
            self.website_pass_input.setEchoMode(QLineEdit.EchoMode.Normal)
            self.website_pass_new.setEchoMode(QLineEdit.EchoMode.Normal)  # Hiện pass
            self.show_pass_button.setText("🙈 Hide Password")
        else:
            self.license_pass_input.setEchoMode(QLineEdit.EchoMode.Password)
            self.website_pass_input.setEchoMode(QLineEdit.EchoMode.Password)
            self.website_pass_new.setEchoMode(QLineEdit.EchoMode.Password)  # Ẩn pass
            self.show_pass_button.setText("👁 Show Password")

#lấy thông tin sheet           


#lấy thông tin web iofo từ sheet
    def get_google_sheets_data(self):
        try:
            selected_sheet = self.sheet_selector.currentText()

            credentials_dict = {
                    "type": "service_account",
                    "project_id": "tool-web-472209",
                    "private_key_id": "9a246c7b94dd306858e6f6dbeea29827bdbeb3bb",
                    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvwIBADANBgkqhkiG9w0BAQEFAASCBKkwggSlAgEAAoIBAQDcQSXc7hITA7Wu\ntKvUkzDpb63PhfQJLjx/amx+te0dlygu8Y5GjiciH9WzQnGKe1MDJi4i76jd74vX\nP+EHWUZu6nEWf+6IyS3ookzQLPPyT2lVh3KCldlES1GU6CRou8uMWPtKb4uPBXeD\nAXCwzeqFrXB6+DUmTFp2kLtNCyeA6T96CzE/ELQw3ZbEcIjs9akdpqPTUFLZKght\nKjw/mNcByQUdlAAcQ4Eqx5uBrAmZWGpRa6CN1AIN2o8jy1hxwrVdNLtKzrBd1mZy\nM3BIO3uSIHIlN2SMGF3g3BsjsAP7Mnn7AbJrROiHAB4WvjHCxzTEx/BE5cA0Q3Rq\n/3vcb4ALAgMBAAECggEANQeOeTHssh1LigO/9091EE8uNu/WgLtcH4VdF+jvbRdu\nula4Xa5uJ54okp9lvOTSaMixhQHU05KQRcQAhwdsmNnjXnpw0Qg3AHLLMbgRB/8s\nqkYRQPWZOQM1Yje+RGUqreDip9pRbJ7xVl5rZnY9h+/pTAIksdLiWPeh0Pvlh/iu\nKo+9CAcDcBem8wfLzKGhKLIPTv/xpVRRQ/gr1KH3jI2qa171ZgO/0zdvDnbJlOVH\ndfAEUNF8XVz36O+hUJBQz2CD7HpTvFO55h64hALcaEGweD/cS0LE+eK/HT8RuU+b\nmL8Lr8WpQD8f9i5CBj61ZOEpnRpZ+liYoVSzwhMmGQKBgQD6PtsaVyqSUFAEKNMH\nwR1a63j/S9X+lomNKqGXtotv6EcxSX9CvnQMFmho9n1YewLlEhhFOxOSD/P7VTFK\nTcxFoo3lb3pwAdzv1mBmlSi/M8UyGtAQ1epJ30WyTigtEgz0L63qNL9B7eEhD1Mb\nLl8FHhSHd17jBcU7wBbuliRszQKBgQDhUb2q476vGy945dg46x681sz2lve/RRwZ\ndNZsnTS6HCHIAB+Qnu521RPVA3AbT3JYTZhWZM/2AVncgj4PgApWXA9Ym1p6XbP0\no9Q6aJk15sA3lsC+0jbV10/iJEuuWCV93wNNtTG0+WOD/gVgQurs/jiFJFvWHOFu\nnSlEsQygNwKBgQCx+4qyTVTGA8EldDPDzCIozFmemj11eXQTp0KPORIrYbVg5LlS\nq0q2XimcndPA3pzMd/YzJzVgKWCKXalVA8hJrrfle0hF6c1N99dQnr4AX73dSRy7\nHKoqKFbV3qjMhY4ZDuBPN3zgU2RPsyqUpoKGjUJkpw4hwbTqLlEhGECH8QKBgQDX\nTCHtzpx/+XwNC6LmEFRYoO9MmMi2bTUCZhAVzMl7JDJrRyLiL9swlT3UBuryTaG3\nGr37n2zPZk8VUyY17WTzTBgl1JxJ3It9saWzAguT45+7/kLCk19uScS9E211dCiu\n84/WitKqWLpsfydn6clNF0WugyV1nDcUWPv79SlZVQKBgQDpvBrzfgtsocwxBRui\nBnmw8DY/gn/j8luvRlYsNYkyDuHUWfcDV5J4QTKHElDTQ5ISQORUJNQTZZJlJSaj\nPmPWsKEMPcV/2RxtA8JBVywX/CkzktDLl5pp23rNQnQVWoMu867UqtS22cwI/+8Y\nUeLiKQBG64My4pb0sbcrh/Z3Yg==\n-----END PRIVATE KEY-----\n",
                    "client_email": "auto-189@tool-web-472209.iam.gserviceaccount.com",
                    "client_id": "110372871813895552037",
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/auto-189%40tool-web-472209.iam.gserviceaccount.com",
                    "universe_domain": "googleapis.com"
                }


            # Constants for Google Sheets configuration
            SHEET_SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            SPREADSHEET_ID = "1a_wPoit-UqeJGTLu7EyMTnIljllOD-LVJedtMnXmqn4"
            SHEET_GIDS = {
                "Tony": "445860585",
                "Lucas": "1922397353",
                "Alex": "1127382190",
                "Allen": "208823018",
                "Hayden": "1606998811",
                "Leo": "730036757",
                "Davis": "443113123",
                "Backup": "1373613829",
            }
            DEFAULT_VALUES = {
                "name": "Mac Marketing",
                "address": "#address",
                "phone": "#phone",
                "link_booking": "#link_booking",
                "email": "#email"
            }
            
            try:
                # Initialize Google Sheets client with caching
                if not hasattr(self, '_sheets_client') or self._sheets_client is None:
                    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, SHEET_SCOPE)
                    self._sheets_client = gspread.authorize(creds)

                # Validate selected sheet
                gid = SHEET_GIDS.get(selected_sheet)
                if gid is None:
                    self.update_sheet_label_signal.emit(f"⚠️ Không tìm thấy tab {selected_sheet}")
                    return

                # Get worksheet and batch read cells for better performance
                spreadsheet = self._sheets_client.open_by_key(SPREADSHEET_ID)
                sheet = spreadsheet.worksheet(selected_sheet)
                
                try:
                    # Batch read all required cells at once
                    ranges = [
                        'A1',          # value_A1 0
                        'B2',          # value_name 1
                        'B3',          # value_address 2
                        'B4',          # value_phone 3
                        'B6',          # value_email 4
                        'B7',          # value_link_booking 5
                        'B8',          # value_hours 6
                        'B10',         # value_link_facebook  7
                        'B11',         # value_link_ins 8
                        'B12',         # value_link_yelp 9
                        'B13',         # value_link_map 10
                        'B14',         # value_link_x 11
                        'C12',         # value_link_rv_yelp 12
                        'C13',         # value_link_rv_map 13
                        'B27',         # value_keyword_B27 14
                        'B23',         # value_AM_B23 15
                        'C23',         # value_Assignee_C23 16
                    ]
                    
                    cell_list = sheet.batch_get(ranges)
                    
                    # Helper function to safely get cell value
                    def get_cell_value(cell_data, default=""):
                        try:
                            return cell_data[0][0] if cell_data and cell_data[0] and cell_data[0][0] else default
                        except (IndexError, TypeError):
                            return default
                    
                    # Assign values with defaults
                    self.value_A1 = get_cell_value(cell_list[0])
                    self.value_name = get_cell_value(cell_list[1], "Mac Marketing")
                    self.value_address = get_cell_value(cell_list[2], "#address")
                    self.value_phone = get_cell_value(cell_list[3], "#phone")
                    self.value_email = get_cell_value(cell_list[4], "#email")
                    self.value_link_booking = get_cell_value(cell_list[5], "#link_booking")
                    self.value_hours = get_cell_value(cell_list[6], "#hours")
                    self.value_link_facebook = get_cell_value(cell_list[7], "#link_facebook")
                    self.value_link_ins = get_cell_value(cell_list[8], "#link_instagram")
                    self.value_link_yelp = get_cell_value(cell_list[9], "#link_yelp")
                    self.value_link_map = get_cell_value(cell_list[10], "#link_map")
                    self.value_link_x = get_cell_value(cell_list[11], "#link_twitter")
                    self.value_link_rv_yelp = get_cell_value(cell_list[12], "#link_review_yelp")
                    self.value_link_rv_map = get_cell_value(cell_list[13], "#link_review_map")
                    self.value_keyword_B27 = get_cell_value(cell_list[14], "")
                    self.value_AM_B23 = get_cell_value(cell_list[15], "")
                    self.value_Assignee_C23 = get_cell_value(cell_list[16], "")
                    
                    print(f"✅ Successfully fetched data from sheet: {selected_sheet}")
                    
                except Exception as e:
                    print(f"⚠️ Error reading sheet values: {str(e)}")
                    # Set default values in case of error
                    self.value_name = "Mac Marketing"
                    self.value_address = "#address"
                    self.value_phone = "#phone"
                    self.value_link_booking = "#link_booking"
                    self.value_email = "#email"
                    self.value_hours = "#hours"
                    self.value_link_facebook = "#link_facebook"
                    self.value_link_ins = "#link_instagram"
                    self.value_link_x = "#link_twitter"
                    self.value_link_map = "#link_map"
                    self.value_link_yelp = "#link_yelp"
                    self.value_link_rv_map = "#link_review_map"
                    self.value_link_rv_yelp = "#link_review_yelp"
                    self.value_keyword_B27 = ""
                    self.value_AM_B23 = ""
                    self.value_Assignee_C23 = ""
                
            except gspread.exceptions.APIError as e:
                print(f"❌ Google Sheets API Error: {str(e)}")
                self.update_sheet_label_signal.emit("⚠️ Lỗi kết nối Google Sheets API")
            except Exception as e:
                print(f"❌ Unexpected error: {str(e)}")
                self.update_sheet_label_signal.emit("⚠️ Lỗi không xác định khi đọc dữ liệu")
                try:
                    # Batch get all remaining data at once to minimize API calls
                    all_remaining_data = sheet.batch_get(['B10:B12', 'B13:B14', 'B8', 'C13:C14', 'B27', 'B23', 'C23'])
                    
                    # Process social media links (B10:B12)
                    try:
                        social_links = all_remaining_data[0] if all_remaining_data else [[""] * 3]
                        social_values = (social_links[0] + [""] * 3)[:3] if social_links else [""] * 3
                        
                        self.value_link_facebook = social_values[0] if social_values[0] else "#link_facebook"
                        self.value_link_ins = social_values[1] if social_values[1] else "#link_instagram"
                        self.value_link_x = social_values[2] if social_values[2] else "#link_twitter"
                    except (IndexError, TypeError) as e:
                        print(f"⚠️ Warning: Error processing social links: {str(e)}")
                        self.value_link_facebook = "#link_facebook"
                        self.value_link_ins = "#link_instagram"
                        self.value_link_x = "#link_twitter"
                    
                    # Process Yelp and Map links (B13:B14)
                    try:
                        yelp_map_data = all_remaining_data[1] if len(all_remaining_data) > 1 else [[""] * 2]
                        yelp_map_values = (yelp_map_data[0] + [""] * 2)[:2] if yelp_map_data else [""] * 2
                        
                        self.value_link_yelp = yelp_map_values[0] if yelp_map_values[0] else "#link_yelp"
                        self.value_link_map = yelp_map_values[1] if yelp_map_values[1] else "#link_google_map"
                    except (IndexError, TypeError) as e:
                        print(f"⚠️ Warning: Error processing yelp/map links: {str(e)}")
                        self.value_link_yelp = "#link_yelp"
                        self.value_link_map = "#link_google_map"
                    
                    # Process Business Hours (B8)
                    try:
                        hours_data = all_remaining_data[2] if len(all_remaining_data) > 2 else [[""]]
                        self.value_hours = hours_data[0][0] if hours_data and hours_data[0] else "Business hours"
                    except (IndexError, TypeError) as e:
                        print(f"⚠️ Warning: Error processing business hours: {str(e)}")
                        self.value_hours = "Business hours"
                    
                    # Process Review links (C13:C14)
                    try:
                        review_data = all_remaining_data[3] if len(all_remaining_data) > 3 else [[""] * 2]
                        review_values = (review_data[0] + [""] * 2)[:2] if review_data else [""] * 2
                        
                        self.value_link_rv_yelp = review_values[0] if review_values[0] else "#link_review_yelp"
                        self.value_link_rv_map = review_values[1] if review_values[1] else "#link_review_map"
                    except (IndexError, TypeError) as e:
                        print(f"⚠️ Warning: Error processing review links: {str(e)}")
                        self.value_link_rv_yelp = "#link_review_yelp"
                        self.value_link_rv_map = "#link_review_map"

                    # Process WPM fields (B27, B23, C23)
                    try:
                        wpm_b27_data = all_remaining_data[4] if len(all_remaining_data) > 4 else [[""]]
                        wpm_b23_data = all_remaining_data[5] if len(all_remaining_data) > 5 else [[""]]
                        wpm_c23_data = all_remaining_data[6] if len(all_remaining_data) > 6 else [[""]]
                        
                        self.value_keyword_B27 = wpm_b27_data[0][0] if wpm_b27_data and wpm_b27_data[0] else ""
                        self.value_AM_B23 = wpm_b23_data[0][0] if wpm_b23_data and wpm_b23_data[0] else ""
                        self.value_Assignee_C23 = wpm_c23_data[0][0] if wpm_c23_data and wpm_c23_data[0] else ""
                    except (IndexError, TypeError) as e:
                        print(f"⚠️ Warning: Error processing wpm config: {str(e)}")
                        self.value_keyword_B27 = ""
                        self.value_AM_B23 = ""
                        self.value_Assignee_C23 = ""
                    
                    # Update UI with success message
                    self.update_sheet_label_signal.emit(f"✅ Đã tải dữ liệu từ sheet {selected_sheet} thành công!")
                    
                except Exception as e:
                    print(f"⚠️ Error processing remaining data: {str(e)}")
                    # Set all default values in case of error
                    self.value_link_facebook = "#link_facebook"
                    self.value_link_ins = "#link_instagram"
                    self.value_link_x = "#link_twitter"
                    self.value_link_yelp = "#link_yelp"
                    self.value_link_map = "#link_google_map"
                    self.value_hours = "Business hours"
                    self.value_link_rv_yelp = "#link_review_yelp"
                    self.value_link_rv_map = "#link_review_map"
                    self.value_keyword_B27 = ""
                    self.value_AM_B23 = ""
                    self.value_Assignee_C23 = ""

            if self.value_A1 and self.value_A1.strip() != "":
                link = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/view?gid={gid}"
                self.update_sheet_label_signal.emit(f"link Sheet: <a href='{link}'>Click</a>")
            else:
                self.update_sheet_label_signal.emit("✅ Không có dữ liệu trong ô A1")

        except Exception as e:
            self.update_sheet_label_signal.emit(f"⚠️ Fail: {e}")

#Lưu thông tin người dùng vào bộ nhớ cục bộ
    def save_settings(self):

        self.settings.setValue("license_user", self.license_user_input.text())
        self.settings.setValue("license_pass", self.license_pass_input.text())
        self.settings.setValue("website_pass_new", self.website_pass_new.text())
        self.settings.setValue("website_user", self.website_user_input.text())
        self.settings.setValue("website_pass", self.website_pass_input.text())
        self.settings.setValue("domain", self.domain_input.text())
        self.settings.setValue("email", self.email_input.text())
        self.settings.setValue("sheet_name", self.sheet_selector.currentText())  # ✅ Lưu tên sheet/tab
        self.settings.sync()
        self.log("✅ Done save your info!")

#Tải thông tin đã lưu khi mở ứng dụng
    def load_settings(self):

        self.license_user_input.setText(self.settings.value("license_user", ""))
        self.license_pass_input.setText(self.settings.value("license_pass", ""))
        self.website_pass_new.setText(self.settings.value("website_pass_new", ""))
        self.website_user_input.setText(self.settings.value("website_user", ""))
        self.website_pass_input.setText(self.settings.value("website_pass", ""))
        self.domain_input.setText(self.settings.value("domain", ""))
        self.email_input.setText(self.settings.value("email", ""))
        saved_sheet = self.settings.value("sheet_name", "Tony")
        
    # ✅ Tải lại tab đã chọn trước đó
        index = self.sheet_selector.findText(saved_sheet)
        if index != -1:
            self.sheet_selector.setCurrentIndex(index)

    def run_script(self):
        license_user = self.license_user_input.text().strip()
        license_pass = self.license_pass_input.text().strip()
        domain = self.domain_input.text().strip()
        email = self.email_input.text().strip()
        website_user = self.website_user_input.text().strip()
        website_pass = self.website_pass_input.text().strip()
        website_pass_new = self.website_pass_new.text().strip()
        sheet_name = self.sheet_selector.currentText().strip()  # ✅ Lấy sheet đang chọn

        if not domain or not license_user or not license_pass or not email or not website_user or not website_pass or not website_pass_new:
            self.log("❌ Incomplete information!")
            return

        self.log(f"🔄 Load... (Tab: {sheet_name})")  # ✅ Gợi ý: log tab hiện tại
        threading.Thread(
            target=self.automate,
            args=(domain, license_user, license_pass, email, website_user, website_pass, website_pass_new, sheet_name),
            daemon=True
        ).start()

#===========> Hàm Logic chạy chương trình <============
    def automate(self, domain, license_user, license_pass, email, website_user, website_pass, website_pass_new, sheet_name):
        # ✅ Chuẩn hóa domain: loại bỏ "/" ở cuối nếu có
        domain = domain.rstrip("/")
        self.log(f"📌 Domain đã được chuẩn hóa: {domain}")
        
        options = Options()
        options.add_argument("--incognito")  # Luôn chạy ở chế độ ẩn danh
        options.add_argument("--start-maximized")

        # Đóng trình duyệt cũ (nếu có)
        if self.driver is not None:
            self.driver.quit()
            self.driver = None
            self.is_logged_in = False  # Đảm bảo trạng thái đăng nhập được reset

        # Mở trình duyệt mới
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 15)

        try:
            # Bắt buộc đăng nhập lại mỗi lần chạy
            self.is_logged_in = False
            if not self.is_logged_in:
                if self.login(self.driver, self.wait, domain, website_user, website_pass):
                    self.is_logged_in = True  # Đánh dấu đã đăng nhập
                else:
                    self.log("⚠️ Fail, skipping next steps!")
                    return  # Dừng luôn nếu không đăng nhập được
                
            if self.chk_change_admin_pass.isChecked():
                try:
                    self.change_admin_password(self.driver, self.wait, domain, website_pass_new)
                except Exception as e:
                    self.log(f"⚠️ Fail pass Admin: {e}")

            if self.chk_change_user_pass.isChecked():
                try:
                    self.change_user_password(self.driver, self.wait, domain)
                except Exception as e:
                    self.log(f"⚠️ Fail pass Tech: {e}")

            if self.chk_limit_login_settings.isChecked():
                try:
                    self.configure_limit_login(self.driver, self.wait, domain, email)
                except Exception as e:
                    self.log(f"⚠️ Fail limit Login: {e}")

            if self.chk_settings_webinfo.isChecked():
                try:
                    self.get_google_sheets_data()
                    self.settings_webinfo(domain, self.driver)
                except Exception as e:
                    self.log(f"⚠️ Fail setting webinfo: {e}")

            if self.chk_contact_form_settings.isChecked():
                try:
                    self.get_google_sheets_data()
                    self.settings_contact_form(domain, self.driver)
                except Exception as e:
                    self.log(f"⚠️ Fail setting contact form: {e}")

            if self.chk_settings_name.isChecked():
                try:
                    self.get_google_sheets_data()
                    self.settings_name_salon(domain, self.driver)
                except Exception as e:
                    self.log(f"⚠️ Fail change name: {e}")

            if self.chk_license.isChecked():
                try:
                    self.automate_active_license(self.driver, self.wait, domain, license_user, license_pass)
                except Exception as e:
                    self.log(f"⚠️ Fail Active License: {e}")
                    
            if self.chk_add_pass_wpm.isChecked():
                try:
                    self.get_google_sheets_data()
                    self.automate_add_pass_wpm(self.driver, self.wait, domain, license_user, license_pass)
                except Exception as e:
                    self.log(f"⚠️ Fail Add pass WPM: {e}")

            self.log("✅ Process complete!")

            # gửi telegram khi thành công
            msg = self.build_report_message(domain)
            self.send_telegram(msg)

        except Exception as e:
            self.log(f"❌ An error occurred: {e}")

            msg = f"""
❌ JOB FAILED

🌐 Domain: {domain}
👤 Người làm: {getattr(self, 'value_Assignee_C23', 'Unknown')}
💥 Error: {str(e)}
"""
            self.send_telegram(msg)

        finally:
            self.log("🔔 The browser is still open for further testing")

    def fill_input(self, element_id, value):
        """Hàm hỗ trợ nhập liệu vào ô input theo ID"""
        try:
            field = self.driver.find_element(By.ID, element_id)
            field.clear()
            field.send_keys(value)
            self.log(f"✔️ Nhập thành công: {element_id} → {value}")
        except Exception as e:
            self.log(f"⚠️ Lỗi khi nhập {element_id}: {e}")

    def settings_webinfo(self, domain, driver, *_):
        try:
            self.log("")  # Xuống 1 dòng
            self.log(f"🚀 Đang cập nhật Web Info từ Google Sheets...")

            # Điều hướng đến trang web info trong WordPress
            url = f"{domain}/wp-admin/options-general.php?page=web-info"
            self.driver.get(url)

            # Chờ trang tải hoàn     
            self.wait.until(EC.presence_of_element_located((By.ID, "booking-link")))

            # Nhập dữ liệu vào các input trên WordPress Admin
            self.fill_input("booking-link", self.value_link_booking)
            self.fill_input("business_phone", self.value_phone)
            self.fill_input("address", self.value_address)
            self.fill_input("map-link", self.value_link_map)
            self.fill_input("business_hours", self.value_hours)
            self.fill_input("facebook-link", self.value_link_facebook)
            self.fill_input("instagram-link", self.value_link_ins)
            self.fill_input("yelp-link", self.value_link_yelp)
            self.fill_input("twitter_link", self.value_link_x)
            self.fill_input("email", self.value_email)
            self.fill_input("google-review-link", self.value_link_rv_map)
            self.fill_input("yelp-review-link", self.value_link_rv_yelp)

            # Lưu cài đặt
            save_button = driver.find_element(By.XPATH, "//button[@type='submit' and contains(@class, 'cx-button-primary-style')]")
            save_button.click()

            time.sleep(5) 

            try:
                self.log("🔄 Đang cập nhật Site Title từ Google Sheets...")

                # Điều hướng đến trang cài đặt chung trong WordPress
                url = f"{domain}/wp-admin/options-general.php"
                driver.get(url)

                # Tìm và cập nhật trường "Site Title" (blogname)
                title_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "blogname"))
                )

                # Kiểm tra nếu trường có dữ liệu, thì xóa trước khi nhập mới
                if title_input.get_attribute("value").strip():
                    title_input.clear()

                # Nhập giá trị mới từ Google Sheets
                title_input.send_keys(self.value_name)

                # Click vào nút "Save Changes"
                save_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.NAME, "submit"))
                )
                save_button.click()

                self.log(f"✅ Đã cập nhật Site Title: {self.value_name}")

            except Exception as e:
                self.log(f"⚠️ Lỗi khi cập nhật Site Title: {e}")

            self.log("✅ Cập nhật Web Info thành công!")
        except Exception as e:
            self.log(f"⚠️ Lỗi khi cập nhật Web Info: {e}")

# Hàm settings Name Salon 
    def settings_name_salon(self, domain, driver, *_):
        try:
            self.log("")  # Xuống 1 dòng, self.log("\n") 2 dòng
            self.log(f"🚀 Đang cập nhật Tên tiệm từ Google Sheets...")

            driver.get(f"{domain}/wp-admin/admin.php?page=wp-mail-smtp")  # Điều hướng đến trang Web Info

            # Tìm và cập nhật trường "From Name"
            name_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "wp-mail-smtp[mail][from_name]"))
            )

            # Kiểm tra nếu trường có dữ liệu, thì xóa trước khi nhập mới
            if name_input.get_attribute("value").strip():
                name_input.clear()

            # Nhập giá trị mới từ Google Sheets
            name_input.send_keys(self.value_name)
            save_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Save Settings')]"))
            )
            save_button.click()   

            self.log(f"✅ Đã cập nhật Name Salon: {self.value_name}")
            time.sleep(5)


        except Exception as e:
            self.log(f"⚠️ Lỗi khi cập nhật Web Info: {e}")

# Hàm settings contact form 7 
    def settings_contact_form(self, domain, driver, *_):
        try:
            # 🚀 Cập nhật HTML Form Policy-------------------------------------------
            
            self.log("")  # Xuống 1 dòng, self.log("\n") 2 dòng
            self.log(f"🚀 Đang cập nhật email cho policy......")
            wait = WebDriverWait(driver, 15)
            driver.get(f"{domain}/wp-admin/admin.php?page=wpcf7&post=62&action=edit#mail-panel")# Tìm và click vào tab "Mail"
            mail_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '#mail-panel')]")))
            mail_tab.click()

            email_input = wait.until(EC.presence_of_element_located((By.ID, "wpcf7-mail-recipient")))

            # Xóa nội dung cũ
            email_input.clear()

            # Nhập email mới từ Google Sheets
            email_input.send_keys(self.value_email)
            time.sleep(5) 
            save_button = wait.until(EC.presence_of_element_located((By.NAME, "wpcf7-save")))
            driver.execute_script("arguments[0].click();", save_button)

            self.log(f"✅ Đã cập nhật email cho Form policy: {self.value_email}")

            time.sleep(5) 


            # 🚀 Cập nhật HTML Form QR Code----------------------------------------
            self.log(f"🚀 Đang cập nhật HTML Form QR...")
            driver.get(f"{domain}/wp-admin/admin.php?page=wpcf7&post=7&action=edit#form-panel")

            # 🔹 Chờ đến khi ô Title hiển thị
            title_input = wait.until(EC.presence_of_element_located((By.ID, "title")))

            # 🔹 Xóa nội dung cũ
            title_input.clear()

            # 🔹 Điền tiêu đề mới
            new_title = "QR Code Form"
            title_input.send_keys(new_title)

            self.log(f"✅ Đã cập nhật tiêu đề Form thành: {new_title}")

            time.sleep(5)

            # 🔹 Chờ textarea xuất hiện
            form_textarea = wait.until(EC.presence_of_element_located((By.ID, "wpcf7-form")))

            # 🔹 Xóa nội dung cũ
            form_textarea.clear()

            # 🔹 Nội dung HTML mới
            new_form_content = """<div class="cf-container">
                <div class="cf-col-6">[text* your-name placeholder "Your name*"]</div>
                <div class="cf-col-6">[tel* your-phone placeholder "Phone*"]</div>
                <div class="cf-col-12">[email your-email placeholder "Email"]</div>
                <div class="cf-col-12">[textarea your-message placeholder "Your message (optional)"]</div>
                <div class="cf-col-12">[submit "Submit"]</div>
            </div>
            <style>
            .cf-container {
                display: -ms-flexbox;
                display: flex;
                -ms-flex-wrap: wrap;
                flex-wrap: wrap;
                margin-right: -5px;
                margin-left: -5px;
            }
            .cf-col-1, .cf-col-2, .cf-col-3, .cf-col-4, .cf-col-5, .cf-col-6, .cf-col-7, .cf-col-8, .cf-col-9, .cf-col-10, .cf-col-11, .cf-col-12 {
                position: relative;
                width: 100%;
                min-height: 1px;
                padding-right: 5px;
                padding-left: 5px;
            }
            @media ( min-width: 576px ) {
                .cf-col-6 {
                    -ms-flex: 0 0 50%;
                    flex: 0 0 50%;
                    max-width: 50%;
                }
                .cf-col-12 {
                    -ms-flex: 0 0 100%;
                    flex: 0 0 100%;
                    max-width: 100%;
                }
            }
            </style>"""

            # 🔹 Điền nội dung mới
            form_textarea.send_keys(new_form_content)

            # 🔹 Lưu form
            save_button = wait.until(EC.presence_of_element_located((By.NAME, "wpcf7-save")))
            driver.execute_script("arguments[0].click();", save_button)
            self.log("✅ Đã cập nhật title + nội dung HTML form!")
            time.sleep(5)

            # 🚀 Chuyển sang tab "Mail" để cập nhật Email & Subject
            self.log(f"🚀 Đang cập nhật email gửi & tiêu đề Subject...")
            driver.get(f"{domain}/wp-admin/admin.php?page=wpcf7&post=7&action=edit#mail-panel")

            wait = WebDriverWait(driver, 15)

            # 🔹 Chờ tab "Mail" hiển thị
            mail_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '#mail-panel')]")))
            mail_tab.click()

            # 🔹 Cập nhật email gửi
            email_input = wait.until(EC.presence_of_element_located((By.ID, "wpcf7-mail-recipient")))
            email_input.clear()
            email_input.send_keys(self.value_email)

            # 🔹 Xóa nội dung cũ của Subject
            subject_input = wait.until(EC.presence_of_element_located((By.ID, "wpcf7-mail-subject")))
            subject_input.clear()

            # 🔹 Điền Subject mới
            new_subject = "[_site_title] Unhappy Customer's Feedback"
            subject_input.send_keys(new_subject)

            # 🔹 Chờ đến khi phần "Additional Headers" hiển thị
            additional_headers = wait.until(EC.presence_of_element_located((By.ID, "wpcf7-mail-additional-headers")))

            # 🔹 Xóa nội dung cũ
            additional_headers.clear()
            # 🔹 Chờ đến khi ô "Message Body" hiển thị
            message_body = wait.until(EC.presence_of_element_located((By.ID, "wpcf7-mail-body")))

            # 🔹 Xóa nội dung cũ
            message_body.clear()

            # 🔹 Nội dung mới cho "Message Body"
            new_message_body = """From: [your-name]
            Email: [your-email]
            Phone: [your-phone]
            Message Body: [your-message]

            Send on: [_date], [_time]
            -- 
            This is a notification that a contact form was submitted on your website ([_site_title] [_site_url])."""

            # 🔹 Điền nội dung mới vào ô
            message_body.send_keys(new_message_body)

            # 🔹 Lưu thay đổi
            save_button = wait.until(EC.presence_of_element_located((By.NAME, "wpcf7-save")))
            driver.execute_script("arguments[0].click();", save_button)

            self.log(f"✅ Đã cập nhật QR Code Form.")

            time.sleep(5)
                        
        
        except Exception as e:
            self.log(f"❌ Lỗi khi cập nhật email: {e}")

# Hàm Lấy key license 
    def automate_active_license(self, driver, wait, domain, license_user, license_pass):

        self.log("")  # Xuống 1 dòng, self.log("\n") 2 dòng
        self.log(f"🚀 Login License...")
        driver.get("https://wpm.macusaone.com/login")
        if not wait:
            wait = WebDriverWait(driver, 20)

        try:
            # Đăng nhập (theo HTML mới: email + password)
            wait.until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(license_user)
            wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(license_pass)
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))).click()
            self.log("2 Đăng nhập thành công!")

            # Truy cập trang License
            driver.get("https://wpm.macusaone.com/websites/license")
            self.log("3 Đang lấy License Key...")

            try:
                # Tìm và click nút Generate Key
                get_key_button = wait.until(
                    EC.element_to_be_clickable((By.ID, "generateKeyBtn"))
                )
                driver.execute_script("arguments[0].scrollIntoView();", get_key_button)
                get_key_button.click()
                self.log("4 Đã click Generate Key!")
            except Exception as e:
                self.log(f"Lỗi khi click nút Generate Key: {e}")
                

            time.sleep(5)

            # Lấy key trực tiếp từ span id="keyGenerated"
            key_span = wait.until(
                EC.presence_of_element_located((By.ID, "keyGenerated"))
            )
            key_value = key_span.text.strip()

            if not key_value:
                self.log("Không tìm thấy Key trong nội dung!")
                

            self.log(f"5 Key lấy được: {key_value}")


            self.log("6 Đang kích hoạt Key trên website...")
            driver.get(f"{domain}/wp-admin/admin.php?page=mac-core")
            time.sleep(5)


        #    try:
        #        toggle_btn = wait.until(EC.element_to_be_clickable((By.ID, "toggle-license-form")))
        #       toggle_btn.click()
        #        self.log("Đã click nút 'Add/Change License Key'.")
        #   except Exception as e:
        #       # fallback nếu click thường không được
        #        try:
        #            toggle_btn = driver.find_element(By.ID, "toggle-license-form")
        #            driver.execute_script("arguments[0].click();", toggle_btn)
        #            self.log("Đã click nút bằng JS (fallback).")
        #        except Exception as e2:
        #            self.log(f"Không thể click nút 'Add/Change License Key': {e2}")

            try:
                key_input = wait.until(EC.presence_of_element_located((By.ID, "kvp-key-input")))
                key_input.clear()
                key_input.send_keys(key_value)
                time.sleep(2)

                validate_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Validate']")))
                validate_button.click()
                time.sleep(5)
                self.log("7 Key đã được kích hoạt!")
                driver.get(f"{domain}/wp-admin/admin.php?page=mac-core")
            except Exception as e:
                self.log(f"Lỗi kích hoạt Key: {e}")

            self.log("8 Install MAC Menu ")

            try:
                install_btn = wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "button.button.button-primary.mac-core-install-plugin[data-plugin-slug='mac-menu']")
                    )
                )
                driver.execute_script("arguments[0].scrollIntoView();", install_btn)
                install_btn.click()
                self.log("9 Activate MAC Menu ")
                time.sleep(5)
                

            except Exception as e:
                self.log(f"❌ Lỗi khi click Install Now: {e}")
                time.sleep(5)
            

# 🚀 Handle alert nếu xuất hiện
            try:
                WebDriverWait(driver, 5).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                self.log(f"⚠️ Alert xuất hiện: {alert.text}")
                alert.accept()  # Hoặc alert.dismiss()
                self.log("✅ Đã đóng alert thành công.")
            except NoAlertPresentException:
                self.log("ℹ️ Không có alert nào bật ra.")
            except Exception as e:
                self.log(f"❌ Lỗi khi xử lý alert: {e}")

           

            self.log("🚀 Đang Activate Menu...")

            driver.get(f"{domain}/wp-admin/plugins.php")
            time.sleep(5)

            try:
                activate_btn = wait.until(
                    EC.element_to_be_clickable((By.ID, "activate-mac-menu"))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", activate_btn)
                activate_btn.click()
                self.log("Đã click nút Activate cho plugin MAC Menu.")
            except Exception as e:
                try:
                    activate_btn = driver.find_element(
                        By.CSS_SELECTOR,
                        "tr[data-slug='mac-menu'] .row-actions .activate a"
                    )
                    driver.execute_script("arguments[0].click();", activate_btn)
                    self.log("Đã click Activate cho MAC Menu bằng fallback selector.")
                except Exception as e2:
                    self.log(f"Không thể click nút Activate MAC Menu: {e2}")

            # Sau khi xử lý xong, code vẫn tiếp tục chạy xuống
            self.log("🚀 Đang cập nhật Privacy Policy...")
            driver.get(f"{domain}/wp-admin/options-general.php?page=mac-privacy-policy-settings")
            time.sleep(5)

            try:
                wait = WebDriverWait(driver, 15)

                # 🔹 Click lần 1
                update_privacy_button = wait.until(
                    EC.element_to_be_clickable((By.NAME, "submit-add-shortcode"))
                )
                update_privacy_button.click()
                time.sleep(3)

                # 🔹 Click lần 2
                wait.until(EC.presence_of_element_located((By.NAME, "submit-add-shortcode")))
                update_privacy_button = wait.until(
                    EC.element_to_be_clickable((By.NAME, "submit-add-shortcode"))
                )
                update_privacy_button.click()
                time.sleep(3)

                self.log("✅ Đã cập nhật Privacy Policy!")
            except Exception as e:
                self.log(f"❌ Lỗi khi cập nhật Privacy Policy: {e}")



            # self.log("9 Kiểm tra cập nhật plugin MAC Menu...")
            #
            # try:
            #     driver.get(f"{domain}/wp-admin/plugins.php")
            #     time.sleep(5)
            #
            #     # Tìm phần tử thông báo cập nhật plugin MAC Menu
            #     update_element = driver.find_elements(By.CSS_SELECTOR, "p.update-message a[href*='update_mac=mac-menu']")
            #     
            #     if update_element:
            #         update_url = update_element[0].get_attribute("href")
            #         full_update_url = domain + update_url if update_url.startswith("/") else update_url
            #
            #         self.log(f"🔄 Có bản cập nhật cho MAC Menu. Đang tiến hành cập nhật...")
            #         driver.get(full_update_url)
            #         time.sleep(5)  # Có thể thay bằng WebDriverWait nếu muốn chắc chắn
            #
            #         self.log("✅ Đã cập nhật plugin MAC Menu thành công.")
            #     else:
            #         self.log("ℹ️ Không có bản cập nhật cho plugin MAC Menu.")
            #
            # except Exception as e:
            #     self.log(f"❌ Lỗi khi kiểm tra/cập nhật plugin: {e}")



        except Exception as e:
            self.log(f"Lỗi trong quá trình lấy và kích hoạt License: {e}")
     
# Hàm đăng nhập vào website
    def login(self, driver, wait, domain, website_user, website_pass):
        
        self.log("")  # Xuống 1 dòng, self.log("\n") 2 dòng
        self.log(f"🚀 Login website...")
        driver.get(f"{domain}/mac-login")
        time.sleep(5)

        try:
            username_input = wait.until(EC.presence_of_element_located((By.NAME, "log")))
            password_input = wait.until(EC.presence_of_element_located((By.NAME, "pwd")))
            login_button = wait.until(EC.element_to_be_clickable((By.ID, "wp-submit")))

            username_input.send_keys(website_user)
            password_input.send_keys(website_pass)
            login_button.click()
            time.sleep(5)
            driver.refresh()  
            time.sleep(5)

            # Kiểm tra đăng nhập thành công bằng cách tìm một phần tử chỉ có sau khi đăng nhập
            if "wp-admin" in driver.current_url:  # Kiểm tra URL có chứa 'wp-admin' không
                self.log("✅ Log in successfully!")
                return True
            else:
                self.log("⚠️ Fail, check again!")
                return False

        except Exception as e:
            self.log(f"❌ Fail Login: {e}")
            return False

# Cài đặt giới hạn đăng nhập
    def configure_limit_login(self, driver, wait, domain, email):

        self.log("")  # Xuống 1 dòng, self.log("\n") 2 dòng
        self.log(f"🚀 Load Limit Login...")
        driver.get(f"{domain}/wp-admin/admin.php?page=limit-login-attempts&tab=settings")

            # Chờ trang load hoàn toàn trước khi nhập liệu
        time.sleep(5)  # Chờ 5 giây
        try:
            # Tick vào checkbox "Email to" nếu chưa chọn
            checkbox = driver.find_element(By.NAME, "lockout_notify_email")
            if not checkbox.is_selected():
                checkbox.click()

            # Đổi email admin
            email_field = driver.find_element(By.NAME, "admin_notify_email")
            email_field.clear()
            email_field.send_keys(email)

            # Tìm nút Save Settings và click
            try:
                save_button = driver.find_element(By.NAME, "llar_update_settings")
                driver.execute_script("arguments[0].click();", save_button)

                self.log(f"✅ Changed mail to :{email}")

            except Exception as e:
                self.log(f"❌ No Check Save Settings! : {e}")


        except Exception as e:
            self.log(f"❌ Fail seting Limit Login: {e}")

# Hàm lưu hoặc cập nhật mật khẩu
    def save_or_update_passwords(self, domain, new_admin_password="", new_user_password=""):
        self.log(f"💾 Đường dẫn file lưu mật khẩu: {os.path.abspath('passwords.txt')}")
        file_path = "passwords.txt"
        updated_lines = []
        found = False

        try:
            with open(file_path, "r", encoding="utf-8") as file:
                lines = file.readlines()

            for line in lines:
                if domain in line:
                    found = True
                    parts = line.strip().split(" | ")
                    admin_pass = new_admin_password if new_admin_password else next((p.split(": ")[-1] for p in parts if "Admin" in p), "")
                    user_pass = new_user_password if new_user_password else next((p.split(": ")[-1] for p in parts if "User" in p), "")
                    updated_lines.append(f"{domain} | Admin: {admin_pass} | User: {user_pass}\n")
                else:
                    updated_lines.append(line)

            if not found:
                updated_lines.append(f"{domain} | Admin: {new_admin_password} | User: {new_user_password}\n")

            with open(file_path, "w", encoding="utf-8") as file:
                file.writelines(updated_lines)

        except FileNotFoundError:
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(f"{domain} | Admin: {new_admin_password} | User: {new_user_password}\n")

# Đổi mật khẩu Admin
    def change_admin_password(self, driver, wait, domain, website_pass_new):
        self.log("")  # Xuống 1 dòng
        self.log(f"🚀 Reset pass Admin...")
        driver.get(f"{domain}/wp-admin/profile.php")
        time.sleep(5)

        try:
            prefix = self.website_pass_new.text().strip()

            set_new_password_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "wp-generate-pw")))
            set_new_password_btn.click()

            password_field = wait.until(EC.element_to_be_clickable((By.ID, "pass1")))
            driver.execute_script("arguments[0].removeAttribute('disabled')", password_field)
            random_password = password_field.get_attribute("value")

            final_password = prefix + random_password
            password_field.clear()
            password_field.send_keys(final_password)

            save_button = wait.until(EC.element_to_be_clickable((By.ID, "submit")))
            save_button.click()
            time.sleep(5)

            self.save_or_update_passwords(domain, new_admin_password=final_password, new_user_password="")
            self.log(f"✅ Reset pass Admin successfully!: {final_password}")

        except Exception as e:
            self.log(f"❌ Fail Reset pass Admin: {e}")

# Đổi mật khẩu User
    def change_user_password(self, driver, wait, domain):
        self.log("")  # Xuống 1 dòng
        self.log("🚀 Reset pass Tech...")
        driver.get(f"{domain}/wp-admin/users.php")
        time.sleep(5)

        try:
            mac_tech_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'mac-tech')]")))
            mac_tech_link.click()

            set_new_password_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "wp-generate-pw")))
            set_new_password_btn.click()

            password_field = wait.until(EC.element_to_be_clickable((By.ID, "pass1")))
            driver.execute_script("arguments[0].removeAttribute('disabled')", password_field)

            # Lấy password tự sinh
            new_password = password_field.get_attribute("value")

            # Không cần clear hay send lại nữa
            save_button = wait.until(EC.element_to_be_clickable((By.ID, "submit")))
            save_button.click()
            time.sleep(5)

            self.save_or_update_passwords(domain, new_admin_password="", new_user_password=new_password)
            self.log(f"✅ Reset pass Tech successfully!: {new_password}")

        except Exception as e:
            self.log(f"❌ Fail Reset pass Tech: {e}")


    
# Add  Assignee pas wpm

    def automate_add_pass_wpm(self, driver, wait, domain, license_user, license_pass):
        self.log("") 
        self.log("🚀 Bắt đầu Add pass WPM...")
        
        try:
            # 1. Tới trang đăng nhập
            driver.get("https://wpm.macusaone.com/login")
            time.sleep(3)
            
            # Kiểm tra xem đã login sẵn hay chưa bằng cách check URL
            # Nếu đã login, sẽ redirect tới /dashboard
            current_url = driver.current_url
            
            if "/dashboard" in current_url:
                # Đã login sẵn, không cần đăng nhập lại
                self.log("✅ Đã đăng nhập sẵn trên WPM, bỏ qua bước login.")
            elif "/login" in current_url:
                # Chưa login, tiến hành đăng nhập
                try:
                    email_input = wait.until(EC.presence_of_element_located((By.NAME, "email")))
                    password_input = wait.until(EC.presence_of_element_located((By.NAME, "password")))
                    submit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))
                    
                    email_input.send_keys(license_user)
                    password_input.send_keys(license_pass)
                    submit_button.click()
                    
                    self.log("🔄 Đang đăng nhập WPM...")
                    # Chờ redirect sang /dashboard sau khi login thành công
                    wait.until(lambda d: "/dashboard" in d.current_url)
                    time.sleep(2)
                    self.log("✅ Đã đăng nhập WPM thành công!")
                except Exception as e:
                    self.log(f"⚠️ Lỗi khi đăng nhập WPM: {e}")
                    raise
            else:
                self.log(f"⚠️ URL không mong đợi: {current_url}")

            # 2. Vào trang website list và tìm kiếm
            driver.get("https://wpm.macusaone.com/projects")
            time.sleep(3)
            
            self.log(f"🔍 Đang tìm kiếm từ khóa: {self.value_keyword_B27}")
            keyword_input = wait.until(EC.presence_of_element_located((By.ID, "keyword")))
            keyword_input.clear()
            keyword_input.send_keys(self.value_keyword_B27)
            
            # Lưu url đợi reload khi bấm Submit
            current_url = driver.current_url
            keyword_input.send_keys(Keys.RETURN)
            
            try:
                wait.until(lambda d: d.current_url != current_url)
                self.log("Đã nhảy trang kết quả tìm kiếm.")
            except Exception:
                self.log("...Quá thời gian chờ URL đổi nhảy.")
            
            time.sleep(3)  # Chờ thêm cho file JS render grid view
            
            # 3. Click nút edit
            edit_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.btn.btn-actions.btn-edit")))
            driver.execute_script("arguments[0].click();", edit_btn)
            self.log("✅ Đã mở bảng điều khiển Edit cho bản ghi.")
            
            time.sleep(5)  # Chờ trang edit tải xong (các select2 cũng load xong)

            # Hàm nhỏ để execute custom JS update select2
            def update_select2(select_id, target_text, label_name):
                script = f"""
                var select = document.getElementById('{select_id}');
                var targetText = arguments[0];
                var found = false;
                for (var i = 0; i < select.options.length; i++) {{
                    if (select.options[i].text.trim() === targetText.trim()) {{
                        select.value = select.options[i].value;
                        found = true;
                        if (typeof jQuery !== 'undefined') {{
                            jQuery(select).trigger('change');
                        }} else {{
                            select.dispatchEvent(new Event('change', {{bubbles: true}}));
                        }}
                        break;
                    }}
                }}
                return found;
                """
                
                if target_text.strip():
                    self.log(f"🔄 Đang gán {label_name} = {target_text}")
                    found = driver.execute_script(script, target_text)
                    if not found:
                        self.log(f"⚠️ Khống tìm thấy tùy chọn {target_text} cho {label_name}!")
                else:
                    self.log(f"ℹ️ {label_name} trống, bỏ qua gán.")

            # 4. Cập nhật Account Manager và Assignee
            update_select2('account-manager-id', self.value_AM_B23, 'Account Manager')
            update_select2('assignee-id', self.value_Assignee_C23, 'Assignee')
            time.sleep(2)
            
            # 5. Lưu lại thông tin
            submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "submitBtn")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", submit_btn)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", submit_btn)
            
            self.log("✅ Đã Submit form gán tài khoản WPM thành công!")
            time.sleep(3)
            
            # Click OK button on the success popup
            try:
                ok_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "swal2-confirm")))
                ok_button.click()
                self.log("✅ Đã nhấn OK trên popup thành công!")
                time.sleep(2)
            except Exception as e:
                self.log(f"⚠️ Lỗi khi nhấn OK popup: {e}")

            # 6. Mở tab "Website Information"
            self.log("▶️ Đang mở tab Website Information...")
            try:
                web_info_tab = wait.until(EC.element_to_be_clickable((By.ID, "tab-web-info-tab")))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", web_info_tab)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", web_info_tab)
                time.sleep(2)
                self.log("✅ Đã mở tab Website Information thành công!")
            except Exception as e:
                # Fallback: thử dùng CSS selector khác
                try:
                    web_info_tab = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#tab-web-info']")))
                    driver.execute_script("arguments[0].click();", web_info_tab)
                    time.sleep(2)
                    self.log("✅ Đã mở tab Website Information (fallback) thành công!")
                except Exception as e2:
                    self.log(f"⚠️ Lỗi mở tab Website Information: {e2}")

            # 6.5. Mở thẻ Web Content (webContent_0)
            self.log("▶️ Đang kiểm tra trạng thái thẻ Web Content...")
            try:
                # Kiểm tra nội dung bên trong có class 'show' không
                web_content = wait.until(EC.presence_of_element_located((By.ID, "webContent_0")))
                if "show" not in web_content.get_attribute("class"):
                    # Click trực tiếp vào icon chevron
                    chevron_icon = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.card-header[data-target='#webContent_0'] i.fa-chevron-down")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", chevron_icon)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", chevron_icon)
                    time.sleep(2)
                    self.log("✅ Đã mở thẻ Web Content thành công!")
                else:
                    self.log("ℹ️ Thẻ Web Content đã mở sẵn.")
            except Exception as e:
                self.log(f"⚠️ Lỗi mở thẻ Web Content: {e}")

            # 7. Mở tab "Endpoint 1"
            self.log("▶️ Đang kiểm tra trạng thái tab Endpoint 1...")
            try:
                # Kiểm tra nội dung bên trong có class 'show' không
                endpoint_content = wait.until(EC.presence_of_element_located((By.ID, "endpoint_content_0_0")))
                if "show" not in endpoint_content.get_attribute("class"):
                    # Click trực tiếp vào icon chevron
                    chevron_icon = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'endpoint-header') and .//h5[contains(text(), 'Endpoint 1')]]//i[contains(@class, 'fa-chevron-down')]")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", chevron_icon)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", chevron_icon)
                    time.sleep(2)
                    self.log("✅ Đã mở tab Endpoint 1 thành công!")
                else:
                    self.log("ℹ️ Tab Endpoint 1 đã mở sẵn.")
            except Exception as e:
                self.log(f"⚠️ Lỗi mở tab Endpoint 1: {e}")

            # 8. Chọn "Add Account" 2 lần
            self.log("▶️ Đang bấm Add Account 2 lần...")
            try:
                # Selector cụ thể hơn cho button Add Account với các class Bootstrap
                btn_selector = "button.btn.btn-success.btn-sm.add-account[data-web-index='0'][data-endpoint-index='0']"
                
                for click_num in range(1, 3):
                    try:
                        add_account_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, btn_selector)))
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", add_account_btn)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", add_account_btn)
                        self.log(f"✅ Lần {click_num}: Đã bấm Add Account thành công!")
                        time.sleep(1)
                    except Exception as e:
                        self.log(f"⚠️ Lần {click_num}: Lỗi bấm button - {e}")
                        raise
                
                time.sleep(2)
                self.log("✅ Đã bấm Add Account 2 lần hoàn tất!")

                # 9. Điền thông tin Account 1 và Account 2
                self.log("▶️ Đang đọc mật khẩu từ file và điền vào form...")
                try:
                    admin_pwd = ""
                    tech_pwd = ""
                    if os.path.exists("passwords.txt"):
                        with open("passwords.txt", "r", encoding="utf-8") as f_pass:
                            for line in f_pass:
                                if domain in line:
                                    parts = line.strip().split(" | ")
                                    admin_pwd = next((p.split(": ")[-1] for p in parts if "Admin" in p), "")
                                    tech_pwd = next((p.split(": ")[-1] for p in parts if "User" in p), "")
                                    break
                    
                    if not admin_pwd or not tech_pwd:
                        self.log("⚠️ Cảnh báo: Không tìm thấy mật khẩu Admin hoặc Tech đầy đủ trong passwords.txt cho domain này!")

                    # Điền thông tin Account 1
                    acc0_user = wait.until(EC.visibility_of_element_located((By.NAME, "webs[0][login_endpoints][0][accounts][0][username]")))
                    acc0_pass = wait.until(EC.visibility_of_element_located((By.NAME, "webs[0][login_endpoints][0][accounts][0][password]")))
                    
                    # Scroll vào view và chờ until clickable
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", acc0_user)
                    time.sleep(0.5)
                    wait.until(EC.element_to_be_clickable((By.NAME, "webs[0][login_endpoints][0][accounts][0][username]")))
                    
                    # Dùng Javascript để fill value an toàn hơn
                    driver.execute_script("arguments[0].value = ''; arguments[0].focus();", acc0_user)
                    acc0_user.send_keys("mac-admin")
                    time.sleep(0.3)
                    
                    driver.execute_script("arguments[0].focus();", acc0_pass)
                    if admin_pwd:
                        acc0_pass.send_keys(admin_pwd)
                    time.sleep(0.3)

                    # Điền thông tin Account 2
                    acc1_user = wait.until(EC.visibility_of_element_located((By.NAME, "webs[0][login_endpoints][0][accounts][1][username]")))
                    acc1_pass = wait.until(EC.visibility_of_element_located((By.NAME, "webs[0][login_endpoints][0][accounts][1][password]")))
                    
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", acc1_user)
                    time.sleep(0.5)
                    wait.until(EC.element_to_be_clickable((By.NAME, "webs[0][login_endpoints][0][accounts][1][username]")))
                    
                    driver.execute_script("arguments[0].value = ''; arguments[0].focus();", acc1_user)
                    acc1_user.send_keys("mac-tech")
                    time.sleep(0.3)
                    
                    driver.execute_script("arguments[0].focus();", acc1_pass)
                    if tech_pwd:
                        acc1_pass.send_keys(tech_pwd)
                    time.sleep(0.3)

                    # Chuyển Select2 của account 2 Role thành "tech"
                    time.sleep(1)  # Chờ select2 initialize
                    try:
                        driver.execute_script("""
                            var select = document.querySelector("select[name='webs[0][login_endpoints][0][accounts][1][role]']");
                            if(select) {
                                var el = select;
                                // Đảm bảo element visible trước khi thay đổi
                                el.scrollIntoView({block: 'center'});
                                el.value = 'tech';
                                
                                // Trigger change event
                                if (typeof jQuery !== 'undefined') {
                                    jQuery(el).trigger('change');
                                } else {
                                    el.dispatchEvent(new Event('change', {bubbles: true}));
                                    el.dispatchEvent(new Event('input', {bubbles: true}));
                                }
                            }
                        """)
                        time.sleep(0.5)
                    except Exception as ex_role:
                        self.log(f"⚠️ Lỗi khi đổi Role sang Tech: {ex_role}")

                    self.log("✅ Đã điền xong mac-admin và mac-tech vào form!")
                    
                    # 10. Ấn submit lưu lại
                    try:
                        # Dùng CSS Selector rõ ràng để chọn đúng nút submit nằm bên trong tab Web Info, tránh bị nhầm với nút submit của tab Project Info
                        final_submit_btn = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#tab-web-info #submitBtn")))
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", final_submit_btn)
                        time.sleep(1)
                        
                        # Đợi cho đến khi button clickable
                        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#tab-web-info #submitBtn")))
                        time.sleep(0.5)
                        
                        # Dùng Javascript kích hoạt click để đảm bảo ăn lấu
                        driver.execute_script("arguments[0].click();", final_submit_btn)
                        self.log("✅ Đã nhấn Nút Submit Lưu Form cuối cùng!")
                        time.sleep(3)
                    except Exception as err_submit:
                        self.log(f"⚠️ Lỗi khi nhấn nút khóa Submit cuối: {err_submit}")

                except Exception as fill_e:
                    self.log(f"⚠️ Lỗi khi điền thông tin tài khoản: {fill_e}")

            except Exception as e:
                self.log(f"⚠️ Lỗi bấm Add Account: {e}")

        except Exception as e:
            self.log(f"❌ Lỗi tự động Add WPM Pass: {e}")

    def send_telegram(self, message):
        TOKEN = "8641265689:AAHktzXJ8kTlHDvs6IgdEd-eZsj6BHSoFqg"
        CHAT_ID = "1767769751"

        try:

            response = requests.post(
                f"https://api.telegram.org/bot{TOKEN}/sendMessage",
                data={
                    "chat_id": CHAT_ID,
                    "text": message
                }
            )
            # Nếu request thất bại ở mức API (như sai token/chat_id), in trực tiếp lỗi ra
            if response.status_code != 200:
                self.log(f"⚠️ Telegram API lỗi {response.status_code}: {response.text}")
            else:
                self.log("✅ Đã gửi log báo cáo lên Telegram!")
        except Exception as e:
            self.log(f"⚠️ Telegram connection error: {e}")

    def build_report_message(self, domain):
        global JOB_COUNTER

        JOB_COUNTER += 1
        now = datetime.now().strftime("%d/%m/%Y : %H:%M:%S")
        assignee = getattr(self, "value_Assignee_C23", "Unknown")

        message = f"""
📊 JOB REPORT

📅 Ngày: {now}
🔢 STT: {JOB_COUNTER}
🌐 Domain: {domain}
👤 Người làm: {assignee}
"""
        return message

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MenuPolicyApp()
    window.show()
    sys.exit(app.exec())
