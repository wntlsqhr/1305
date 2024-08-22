from PyQt5.QtGui import QFont, QIcon, QStandardItemModel, QStandardItem, QTextBlock, QTextCursor
from PyQt5.QtCore import Qt, QThread, QObject, pyqtSignal, QCoreApplication
from PyQt5.QtWidgets import (QApplication, QWidget, QGroupBox, QRadioButton
, QCheckBox, QPushButton, QMenu, QGridLayout, QVBoxLayout)
from selenium.common.exceptions import SessionNotCreatedException
from openpyxl.utils.exceptions import InvalidFileException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoAlertPresentException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from gspread.utils import rowcol_to_a1
from openpyxl import Workbook, load_workbook
from selenium import webdriver
from datetime import datetime, date, timedelta
from gspread_formatting import *
from PyQt5.QtWidgets import *
import pandas as pd
import chromedriver_autoinstaller
import datetime
import threading
import openpyxl
import gspread
import json
import time
import glob
import csv
import sys
import os
import re


class Rawdata_extractor(QWidget):

    def __init__(self):
        super().__init__()
        self.UI초기화()

    def UI초기화(self):

        self.setWindowTitle("Raw data 자동 추출기")
        self.setFixedSize(1000, 800)


# 매출 group box
        self.sales_group_box = QGroupBox("매출",self)
        self.sales_group_box.setFont(QFont('Helvetia', 20, QFont.Bold))
        self.sales_group_box.move(40,10)
        self.sales_group_box.setFixedSize(400, 300)

    # 카페24
        self.salesCafe24 = QLabel("카페24",self)
        self.salesCafe24.move(50,50)
        self.salesCafe24.setFont(QFont('Helvetia', 14, QFont.Bold))

        # 하엔
        self.haen_salesCafe24 = QCheckBox("하엔",self)
        self.haen_salesCafe24.move(50,75)
        self.haen_salesCafe24.setFont(QFont('Helvetia', 11))

        # 러블로
        self.love_salesCafe24 = QCheckBox("러블로",self)
        self.love_salesCafe24.move(130,75)
        self.love_salesCafe24.setFont(QFont('Helvetia', 11))

        # 노마셀
        self.know_salesCafe24 = QCheckBox("노마셀",self)
        self.know_salesCafe24.move(210,75)
        self.know_salesCafe24.setFont(QFont('Helvetia', 11))

        # 제니크
        self.zq_salesCafe24 = QCheckBox("제니크",self)
        self.zq_salesCafe24.move(290,75)
        self.zq_salesCafe24.setFont(QFont('Helvetia', 11))

    # 쿠팡
        self.salesCoup = QLabel("쿠팡",self)
        self.salesCoup.move(50,110)
        self.salesCoup.setFont(QFont('Helvetia', 15, QFont.Bold))

        # 하엔
        self.haen_salesCoup = QCheckBox("하엔",self)
        self.haen_salesCoup.move(50,135)
        self.haen_salesCoup.setFont(QFont('Helvetia', 11))

        # 러블로
        self.love_salesCoup = QCheckBox("러블로",self)
        self.love_salesCoup.move(130,135)
        self.love_salesCoup.setFont(QFont('Helvetia', 11))

        # 노마셀
        self.know_salesCoup = QCheckBox("노마셀",self)
        self.know_salesCoup.move(210,135)
        self.know_salesCoup.setFont(QFont('Helvetia', 11))

    # 네이버
        self.salesNaver = QLabel("네이버",self)
        self.salesNaver.move(50,170)
        self.salesNaver.setFont(QFont('Helvetia', 15, QFont.Bold))

        # 하엔
        self.haen_salesNaver = QCheckBox("하엔",self)
        self.haen_salesNaver.move(50,195)
        self.haen_salesNaver.setFont(QFont('Helvetia', 11))

        # 러블로
        self.love_salesNaver = QCheckBox("러블로",self)
        self.love_salesNaver.move(130,195)
        self.love_salesNaver.setFont(QFont('Helvetia', 11))

        # 노마셀
        self.know_salesNaver = QCheckBox("노마셀",self)
        self.know_salesNaver.move(210,195)
        self.know_salesNaver.setFont(QFont('Helvetia', 11))


# 광고 group box
        self.advt_group_box = QGroupBox("광고",self)
        self.advt_group_box.setFont(QFont('Helvetia', 20, QFont.Bold))
        self.advt_group_box.move(490,10)
        self.advt_group_box.setFixedSize(400, 300)

    # 쿠팡
        self.advtCoup = QLabel("쿠팡",self)
        self.advtCoup.move(500,50)
        self.advtCoup.setFont(QFont('Helvetia', 15, QFont.Bold))

        # 하엔
        self.haen_advtCoup = QCheckBox("하엔",self)
        self.haen_advtCoup.move(500,75)
        self.haen_advtCoup.setFont(QFont('Helvetia', 11))

        # 러블로
        self.love_advtCoup = QCheckBox("러블로",self)
        self.love_advtCoup.move(580,75)
        self.love_advtCoup.setFont(QFont('Helvetia', 11))

        # 노마셀
        self.know_advtCoup = QCheckBox("노마셀",self)
        self.know_advtCoup.move(660,75)
        self.know_advtCoup.setFont(QFont('Helvetia', 11))

    # 네이버
        self.advtNaver = QLabel("네이버",self)
        self.advtNaver.move(500,110)
        self.advtNaver.setFont(QFont('Helvetia', 15, QFont.Bold))

        # 하엔
        self.haen_advtNaver = QCheckBox("하엔",self)
        self.haen_advtNaver.move(500,135)
        self.haen_advtNaver.setFont(QFont('Helvetia', 11))

        # 러블로
        self.love_advtNaver = QCheckBox("러블로",self)
        self.love_advtNaver.move(580,135)
        self.love_advtNaver.setFont(QFont('Helvetia', 11))

        # 노마셀
        self.know_advtNaver = QCheckBox("노마셀",self)
        self.know_advtNaver.move(660,135)
        self.know_advtNaver.setFont(QFont('Helvetia', 11))

    # GFA
        self.advtGFA = QLabel("GFA",self)
        self.advtGFA.move(500,170)
        self.advtGFA.setFont(QFont('Helvetia', 15, QFont.Bold))

        # 하엔
        self.haen_advtGFA = QCheckBox("하엔",self)
        self.haen_advtGFA.move(500,195)
        self.haen_advtGFA.setFont(QFont('Helvetia', 11))

        # 러블로
        self.love_advtGFA = QCheckBox("러블로",self)
        self.love_advtGFA.move(580,195)
        self.love_advtGFA.setFont(QFont('Helvetia', 11))

        # 노마셀
        self.know_advtGFA = QCheckBox("노마셀",self)
        self.know_advtGFA.move(660,195)
        self.know_advtGFA.setFont(QFont('Helvetia', 11))

    # 파워컨텐츠
        self.advtPC = QLabel("파워컨텐츠",self)
        self.advtPC.move(500,230)
        self.advtPC.setFont(QFont('Helvetia', 15, QFont.Bold))

        # 하엔
        self.haen_advtPC = QCheckBox("하엔",self)
        self.haen_advtPC.move(500,255)
        self.haen_advtPC.setFont(QFont('Helvetia', 11))

        # 러블로
        self.love_advtPC = QCheckBox("러블로",self)
        self.love_advtPC.move(580,255)
        self.love_advtPC.setFont(QFont('Helvetia', 11))

        # 노마셀
        self.know_advtPC = QCheckBox("노마셀",self)
        self.know_advtPC.move(660,255)
        self.know_advtPC.setFont(QFont('Helvetia', 11))


# 기타 group box
        self.etc_group_box = QGroupBox("기타",self)
        self.etc_group_box.setFont(QFont('Helvetia', 20, QFont.Bold))
        self.etc_group_box.move(40,340)
        self.etc_group_box.setFixedSize(400, 300)

    # 카페24 방문자수
        self.visitors = QLabel("방문자수",self)
        self.visitors.move(50,380)
        self.visitors.setFont(QFont('Helvetia', 15, QFont.Bold))

        # 하엔
        self.haen_visitors = QCheckBox("하엔",self)
        self.haen_visitors.move(50,405)
        self.haen_visitors.setFont(QFont('Helvetia', 11))

        # 러블로
        self.love_visitors = QCheckBox("러블로",self)
        self.love_visitors.move(130,405)
        self.love_visitors.setFont(QFont('Helvetia', 11))

        # 노마셀
        self.know_visitors = QCheckBox("노마셀",self)
        self.know_visitors.move(210,405)
        self.know_visitors.setFont(QFont('Helvetia', 11))

    # 카페24 신규가입자수
        self.newMemb = QLabel("신규가입자수",self)
        self.newMemb.move(50,440)
        self.newMemb.setFont(QFont('Helvetia', 15, QFont.Bold))

        # 하엔
        self.haen_newMemb = QCheckBox("하엔",self)
        self.haen_newMemb.move(50,465)
        self.haen_newMemb.setFont(QFont('Helvetia', 11))

        # 러블로
        self.love_newMemb = QCheckBox("러블로",self)
        self.love_newMemb.move(130,465)
        self.love_newMemb.setFont(QFont('Helvetia', 11))

        # 노마셀
        self.know_newMemb = QCheckBox("노마셀",self)
        self.know_newMemb.move(210,465)
        self.know_newMemb.setFont(QFont('Helvetia', 11))

        #불러오기 체크박스설정
        self.loadCheckboxState()


# 버튼

    # 다운로드
        # 다운로드폴더 버튼
        self.slt_folder = QPushButton('다운로드폴더',self)
        self.slt_folder.setGeometry(330,511,100,29)
        self.slt_folder.clicked.connect(self.folderopen)
        self.slt_folder.setStyleSheet(
            """
            QPushButton {
                background-color: white;
                border-radius: 1.5px;
                border-width: 1px;
                border-color: black;
                border-style: solid;
            }
            QPushButton:hover {
                background-color: rgb(120,120,120);
            }
            QPushButton:pressed {
                background-color: rgb(50, 50, 50);
            }
            """
        )

        # 다운로드폴더 설정저장 버튼
        self.saveButton = QPushButton('설정저장', self)
        self.saveButton.setGeometry(440,511,100,29)
        self.saveButton.clicked.connect(self.saveText)
        self.saveButton.setStyleSheet(
            """
            QPushButton {
                background-color: white;
                border-radius: 1.5px;
                border-width: 1px;
                border-color: black;
                border-style: solid;
            }
            QPushButton:hover {
                background-color: rgb(120,120,120);
            }
            QPushButton:pressed {
                background-color: rgb(50, 50, 50);
            }
            """
        )

         # 다운로드폴더 경로
        self.path_folder = QLineEdit(self)
        self.path_folder.setGeometry(80,511,240,27)
        self.path_folder.setStyleSheet(
                        "background-color: white;"
                        "border-radius: 1.5px;"
                        "border-width: 1px;"
                        "border-color: black;"
                        "border-style: solid;")  # 테두리 스타일 추가
        self.path_folder.setReadOnly(True)

    # 크롬폴더
        # 크롬폴더 버튼
        self.chrome_slt_folder = QPushButton('크롬 폴더',self)
        self.chrome_slt_folder.setGeometry(330,560,100,29)
        self.chrome_slt_folder.clicked.connect(self.chromefolderopen)
        self.chrome_slt_folder.setStyleSheet(
            """
            QPushButton {
                background-color: white;
                border-radius: 1.5px;
                border-width: 1px;
                border-color: black;
                border-style: solid;
            }
            QPushButton:hover {
                background-color: rgb(120,120,120);
            }
            QPushButton:pressed {
                background-color: rgb(50, 50, 50);
            }
            """
        )

        # 크롬폴더 경로
        self.chrome_path_folder = QLineEdit(self)
        self.chrome_path_folder.setGeometry(80,560,240,27)
        self.chrome_path_folder.setStyleSheet(
                        "background-color: white;"
                        "border-radius: 1.5px;"
                        "border-width: 1px;"
                        "border-color: black;"
                        "border-style: solid;")  # 테두리 스타일 추가
        self.chrome_path_folder.setReadOnly(True)
        self.loadText()

    # 날짜
        # 날짜 선택
        self.combo = QComboBox(self)
        self.combo.setGeometry(75, 613, 50, 39)
        self.combo.addItems(["1", "2", "3", "4", "5", "6", "7"])
        self.combo.setFont(QFont('Helvetia', 12, QFont.Bold))

        # 날짜 레이블
        self.daybefore = QLabel("일 전까지", self)
        self.daybefore.move(75, 655)
        self.daybefore.setFont(QFont('Helvetia', 12, QFont.Bold))

    # 추출
        # 추출하기
        self.extr_button = QPushButton('추출하기',self)
        self.extr_button.setGeometry(130,612,410,40)
        self.extr_button.clicked.connect(self.extract)
        self.extr_button.setStyleSheet(
            """
            QPushButton {
                background-color: white;
                border-radius: 1.5px;
                border-width: 1px;
                border-color: black;
                border-style: solid;
            }
            QPushButton:hover {
                background-color: rgb(120,120,120);
            }
            QPushButton:pressed {
                background-color: rgb(50, 50, 50);
            }
            """
        )

    def extract(self):

    # 타겟날짜 변수 저장
        target_days_input = int(self.combo.currentText())

        # 크롬 On
        chromedriver_autoinstaller.install()
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_argument("--start-maximized") #최대 크기로 시작
        # chrome_options.add_argument('--incognito')
        # chrome_options.add_argument('--window-size=1920,1080')  
        # chrome_options.add_argument('--headless')
        chrome_options.add_experimental_option('detach', True)

        user_data = self.chrome_path_folder.text()
        user_data = 'C:\\Users\\A\\AppData\\Local\\Google\\Chrome\\User Data1'
        chrome_options.add_argument(f"user-data-dir={user_data}")
        chrome_options.add_argument("--profile-directory=Profile 1")

        
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
        headers = {'user-agent' : user_agent}

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )

        download_folder = self.path_folder.text()


        def count_files(folder):
            """ 폴더 내 파일의 개수를 반환합니다. """
            return len([name for name in os.listdir(folder) if os.path.isfile(os.path.join(folder, name))])

        def get_latest_file(folder):
            """ 폴더 내에서 가장 최신의 파일을 반환합니다. """
            files = [os.path.join(folder, f) for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
            latest_file = max(files, key=os.path.getctime)
            return latest_file
        
        def get_previous_latest_file(folder):
            """폴더 내에서 가장 최신 파일을 제외한 이전 파일을 반환합니다."""
            # 폴더 내의 파일들의 전체 경로와 함께 리스트를 생성합니다.
            files = [os.path.join(folder, f) for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
            
            # 파일이 없다면 None을 반환
            if not files:
                return None
            
            # 파일들을 생성 시간 기준으로 정렬합니다.
            files.sort(key=os.path.getctime)
            
            # 가장 최신 파일을 제외한 가장 최신 파일을 찾습니다.
            # 파일이 하나만 있는 경우에는 그 파일이 최신 파일이므로, None을 반환합니다.
            if len(files) > 1:
                previous_latest_file = files[-2]  # 뒤에서 두 번째 항목 선택
                print(previous_latest_file)
                return previous_latest_file
            else:
                return None
            
        def get_nth_latest_file(folder, n):
            """폴더 내에서 n번째로 최신 파일을 반환합니다. n이 1이면 가장 최신, 2면 두 번째로 최신 파일을 반환합니다."""
            # 폴더 내의 파일들의 전체 경로와 함께 리스트를 생성합니다.
            files = [os.path.join(folder, f) for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
            
            # 파일이 없다면 None을 반환
            if not files:
                return None
            
            # 파일들을 생성 시간 기준으로 정렬합니다.
            files.sort(key=os.path.getctime, reverse=True)
            
            # 요청한 순위의 파일을 반환합니다. n이 파일 수보다 많거나 0 이하인 경우 None을 반환합니다.
            if 1 <= n <= len(files):
                nth_latest_file = files[n-1]  # n번째 파일 선택
                print(nth_latest_file)
                return nth_latest_file
            else:
                return None

        def check_download():
            # 다운로드 전의 파일 개수 확인
                initial_file_count = count_files(download_folder)

                # 다운로드 시작 ...

                # 새 파일이 다운로드될 때까지 기다림
                global check
                check = 0
                i = 0

                while i < 20:
                    current_file_count = count_files(download_folder)
                    if current_file_count > initial_file_count:
                        print("A new file has been downloaded.")
                        latest_file = get_latest_file(download_folder)
                        print(f"Downloaded file: {latest_file}")
                        # 여기서 필요한 작업을 수행하세요, 예를 들면 파일 열기 등
                        check = 1
                        break
                    else:
                        print("Still no new file")
                    time.sleep(0.3)  # 폴더 상태를 0.3초마다 체크
                    i += 1
                return check

        def check_data_in_second_row(file_path):
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active
            second_row = list(sheet.iter_rows(min_row=2, max_row=2, values_only=True))
            if second_row and any(cell is not None for cell in second_row[0]):
                return True
            return False
        
        def convert_data(data):
            result = []
            for item in data:
                if isinstance(item, str) and '%' in item:
                    result.append(float(item.strip('%')) / 100)
                elif isinstance(item, str) and ',' in item:
                    result.append(int(item.replace(',', '')))
                elif item.isdigit():
                    result.append(int(item))
                else:
                    result.append(item)
            return result
    
        # 날짜 변수
        dayx = datetime.timedelta(days=target_days_input)
        day1 = datetime.timedelta(days=1)
        today = date.today()

        today_date = today.strftime("%d")
        today_month = str(int(today.strftime("%m")))

        weekday_korean = {
            0: '월',
            1: '화',
            2: '수',
            3: '목',
            4: '금',
            5: '토',
            6: '일'
        }

        # 오늘 날짜 구하기
        today_yday = today-day1
        today_tday = today-dayx
        today_Tday년월 = (today-dayx).strftime("%Y년 %m월")
        today_Yday년월 = (today-day1).strftime("%Y년 %m월")
        Tday_month월 = str(int(today_tday.strftime("%m"))) + "월"
        Yday_month월 = str(int(today_yday.strftime("%m"))) + "월"
        today_Tday일 = str(int((today-dayx).strftime("%d")))
        today_Yday일 = str(int((today-day1).strftime("%d")))


        weekday_num = today.weekday()  # 요일 번호 (월요일이 0, 일요일이 6)
        weekday_numy = today_yday.weekday()  # 요일 번호 (월요일이 0, 일요일이 6)
        weekday_numt = today_tday.weekday()  # 요일 번호 (월요일이 0, 일요일이 6)
        # 요일을 한국어로 변환
        weekday_kr = weekday_korean[weekday_num]
        weekday_kry = weekday_korean[weekday_numy]
        weekday_krt = weekday_korean[weekday_numt]

        weekday = f"{today}({weekday_kr})"
        weekday_y = f"{today_yday}({weekday_kry})"
        weekday_t = f"{today_tday}({weekday_krt})"

        #카페24
        def cafe24(url_cafe24, url_cafe24_req, cafe24_id, cafe24_pw, sheet_urlR, sheet_nameR, sheet_nameD):
            
            driver.get(url_cafe24)

            ##################################### 로그인
            ##################################### 로그인
            ##################################### 로그인
            ##################################### 로그인

            # ID
            input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#mall_id")))
            input_field.click()
            time.sleep(1)
            input_field.send_keys(Keys.CONTROL + "a")
            input_field.send_keys(Keys.BACKSPACE)
            driver.find_element(By.CSS_SELECTOR, "#mall_id").send_keys(cafe24_id)

            # PW
            input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#userpasswd")))
            input_field.click()
            input_field.send_keys(Keys.CONTROL + "a")
            input_field.send_keys(Keys.BACKSPACE)
            driver.find_element(By.CSS_SELECTOR, "#userpasswd").send_keys(cafe24_pw)

            # 로그인클릭
            driver.find_element(By.CSS_SELECTOR,'#frm_user > div > div.mButton > button').click()

            #비밀번호변경안내
            try: WebDriverWait(driver, 5).until(EC.element_to_be_clickable(((By.CSS_SELECTOR,"#iptBtnEm")))).click() 
            except: pass

            try:
                time.sleep(3)
                popup = driver.find_element(By.XPATH, '//*[contains(text(), "오늘 하루 보지 않기")]')
                popup.click()

            except: pass

            #화면로딩대기
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[contains(text(), "오늘의 할 일")]')))

            # 데이터 접근
            driver.get(url_cafe24_req)

            # 자세히보기클릭
            element = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#QA_day3 > div.mBoard.gScroll > table")))
            driver.execute_script("arguments[0].scrollIntoView(true);", element) # 스크롤다운
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#sReportGabView"))).click() 
            
            element = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#QA_day3 > div.mBoard.gScroll > table")))
            driver.execute_script("arguments[0].scrollIntoView(true);", element) # 스크롤다운
            rows = driver.find_elements(By.CSS_SELECTOR, 'tbody.right tr')

            cover = []
            cover0 = []
            for element in rows:
                new_data_list = []
                rawdata = element.text
                # 문자열을 공백을 기준으로 분리하여 리스트로 변환
                data_list = rawdata.split()
                data_list = [x.replace(',', '') for x in data_list]

                for items in data_list:

                # 숫자인 경우 숫자로 변환
                    try:
                        numeric_value = int(items)
                        new_data_list.append(numeric_value)
                    except:
                        # 숫자가 아닌 경우 원래 값 유지
                        new_data_list.append(items)
                cover.append(new_data_list[1:])
                cover0.append(new_data_list[0])


            today_tdayTemp = today_tday
            today_tdayTempDay = today_tday.strftime(f"%Y-%m-%d({weekday_krt})")
            cover.reverse()
            cover0.reverse()
            print(cover)

            print(cover0)
            print(today_tdayTemp)
            for i in range(target_days_input):
                print(today_tdayTemp)

                    # 서비스 계정 키 파일 경로
                credential_file = 'triple-nectar-412808-da4dac0cc16e.json'

                # gspread 클라이언트 초기화
                client = gspread.service_account(filename=credential_file)

                # Google 시트 열기
                spreadsheet = client.open_by_url(sheet_urlR)

                # 첫 번째 시트 선택
                sheet = spreadsheet.worksheet(sheet_nameR)

                column_values = sheet.col_values(1)
                for idx, cell_value in enumerate(column_values, start=1):  # start=1로 설정하여 행 번호를 1부터 시작
                    if cell_value == str(today_tdayTemp):
                        print(cell_value)
                        print(gspread.utils.rowcol_to_a1(idx, 1))
                        cell_addr = gspread.utils.rowcol_to_a1(idx, 1)
                        # return f"{gspread.utils.rowcol_to_a1(idx, 1)}"  # 셀 주소 반환
                    
                (start_row, start_col) = gspread.utils.a1_to_rowcol(cell_addr)
                print(start_row)
                print(today_tdayTempDay)

                # last_row = len(sheet.col_values(3))
                # next_row = last_row + 1
                # print(last_row)
                # print(next_row)

                if today_tdayTempDay in cover0:
                    print("성립")
                    keynum = cover0.index(today_tdayTempDay)
                    data_to_paste = cover[keynum]    

                    data1 = data_to_paste[:9]
                    data2 = data_to_paste[9]
                    data3 = data_to_paste[10:]
                # 카페24 R, 데이터 없으면 0 입력 되도록 코드 수정 -2
                else:
                    data1 = [0, 0, 0, 0, 0, 0, 0, 0, 0]
                    data2 = 0
                    data3 = [0, 0, 0]


                print(data1)
                print(data2)
                print(data3)

                range1 = f'C{start_row}:K{start_row}'
                range2 = f'M{start_row}'
                range3 = f'O{start_row}:Q{start_row}'
                
                sheet.update([data1], range1)
                sheet.update([[data2]], range2)
                sheet.update([data3], range3)

                # 날짜 넘김 처리
                today_tdayTemp = today_tdayTemp + timedelta(days=1)
                weekday_numtTemp = today_tdayTemp.weekday()  # 요일 번호 (월요일이 0, 일요일이 6)
                weekday_krtTemp = weekday_korean[weekday_numt]
                weekday_krtTemp = weekday_korean[weekday_numtTemp]
                today_tdayTempDay = f"{today_tdayTemp}({weekday_krtTemp})"


        #카페24 하엔
        if self.haen_salesCafe24.isChecked() == True:

            url_cafe24 = "https://eclogin.cafe24.com/Shop/"
            url_cafe24_req_haen = "https://woo8425.cafe24.com/disp/admin/shop1/report/DailyList"
            
            cafe24_id_haen = self.login_info("CAFE_HAEN_ID")
            cafe24_pw_haen = self.login_info("CAFE_HAEN_PW")

            sheet_haenR_url = 'https://docs.google.com/spreadsheets/d/145lVmBVqp87AwsRK9KCclE-Dgkh0B7jbwsfaHKmwOz0/edit#gid=1894651086'
            sheet_haenR = '하엔R'
            sheet_haenD = "하엔D"
        
            cafe24(url_cafe24, url_cafe24_req_haen, cafe24_id_haen, cafe24_pw_haen, sheet_haenR_url, sheet_haenR, sheet_haenD)

        #카페24 러블로
        if self.love_salesCafe24.isChecked() == True:

            url_cafe24 = "https://eclogin.cafe24.com/Shop/" 
            url_cafe24_req_lovelo = "https://wooo8425.cafe24.com/disp/admin/shop1/report/DailyList"

            cafe24_id_lovelo = self.login_info("CAFE_LOVE_ID")
            cafe24_pw_lovelo = self.login_info("CAFE_LOVE_PW")

            sheet_loveR_url = 'https://docs.google.com/spreadsheets/d/145lVmBVqp87AwsRK9KCclE-Dgkh0B7jbwsfaHKmwOz0/edit#gid=872830966'
            sheet_loveR = '러블로R'
            sheet_loveD = "러블로D"

            cafe24(url_cafe24, url_cafe24_req_lovelo, cafe24_id_lovelo, cafe24_pw_lovelo, sheet_loveR_url, sheet_loveR, sheet_loveD)

        #카페24 노마셀
        if self.know_salesCafe24.isChecked() == True:

            url_cafe24 = "https://eclogin.cafe24.com/Shop/" 
            url_cafe24_req_knowmycell = "https://fkark12.cafe24.com/disp/admin/shop1/report/DailyList"

            cafe24_id_knowmycell = self.login_info("CAFE_KNOW_ID")
            cafe24_pw_knowmycell = self.login_info("CAFE_KNOW_PW")

            sheet_knowR_url = 'https://docs.google.com/spreadsheets/d/145lVmBVqp87AwsRK9KCclE-Dgkh0B7jbwsfaHKmwOz0/edit#gid=567505346'
            sheet_knowR = '노마셀R'
            sheet_knowD = "노마셀D"

            cafe24(url_cafe24, url_cafe24_req_knowmycell, cafe24_id_knowmycell, cafe24_pw_knowmycell, sheet_knowR_url, sheet_knowR, sheet_knowD)

        #카페24 제니크
        if self.zq_salesCafe24.isChecked() == True:

            url_cafe24 = "https://eclogin.cafe24.com/Shop/" 
            url_cafe24_req_ZQ = "https://fkark08.cafe24.com/disp/admin/shop1/report/DailyList"

            cafe24_id_ZQ = self.login_info("CAFE_ZQ_ID")
            cafe24_pw_ZQ = self.login_info("CAFE_ZQ_PW")

            sheet_ZQR_url = 'https://docs.google.com/spreadsheets/d/145lVmBVqp87AwsRK9KCclE-Dgkh0B7jbwsfaHKmwOz0/edit#gid=567505346'
            sheet_ZQR = '제니크R'
            sheet_ZQD = "제니크D"
            brand = "제니크"

            cafe24(url_cafe24, url_cafe24_req_ZQ, cafe24_id_ZQ, cafe24_pw_ZQ, sheet_ZQR_url, sheet_ZQR, sheet_ZQD)
    
<<<<<<< HEAD
        # 쿠팡
        def coupang(url_coupang_daily, coupang_id, coupang_pw, coupC_url):
            try:
                driver.get(url_coupang_daily)  # 로그인 시작
                if driver.find_element(By.CSS_SELECTOR, "body > pre"):
                    driver.get(url_coupang_daily)  # 요소가 존재하면 페이지를 다시 로드
            except NoSuchElementException:
                # 요소가 없을 때 처리할 로직
                pass

            try:
                driver.get(url_coupang_daily)  # 로그인 시작
                if driver.find_element(By.CSS_SELECTOR, "body > h1"):
                    driver.get(url_coupang_daily)  # 요소가 존재하면 페이지를 다시 로드
            except NoSuchElementException:
                pass
            
            try:
                loginElements = driver.find_elements(By.XPATH, '//*[contains(text(), "로그인하기")]')

                if len(loginElements) > 1:  # 요소가 두 개 이상 있는지 확인
                    
                    loginElements[0].click()

                # if driver.find_element(By.CSS_SELECTOR, "#main-container > div > div.sc-30ec2de1-0.tedRR > ul > li:nth-child(1) > a > span"):
                #     driver.find_element(By.CSS_SELECTOR, "#main-container > div > div.sc-30ec2de1-0.tedRR > ul > li:nth-child(1) > a > span").click()

            except NoSuchElementException:
                # 요소가 없을 때 처리할 로직
                pass

            print(coupang_id)
            print(coupang_pw)
            input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#username")))
            input_field.click()
            time.sleep(0.7)
            input_field.send_keys(Keys.CONTROL + "a")
            input_field.send_keys(Keys.BACKSPACE)
            driver.find_element(By.CSS_SELECTOR, "#username").send_keys(coupang_id)
            input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#password")))
            input_field.click()
            input_field.send_keys(Keys.CONTROL + "a")
            input_field.send_keys(Keys.BACKSPACE)
            driver.find_element(By.CSS_SELECTOR, "#password").send_keys(coupang_pw)
            driver.find_element(By.CSS_SELECTOR,'#kc-login').click()

            # 로그인 오류 발생하면 재시도
            ### 비밀번호 오류 예외문
            try:
                loginErrorMessage = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[contains(text(), "아이디 또는 비밀번호가 다릅니다.")]')))
                if loginErrorMessage:
                    input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#username")))
                    input_field.click()
                    time.sleep(0.7)
                    input_field.send_keys(Keys.CONTROL + "a")
                    input_field.send_keys(Keys.BACKSPACE)
                    driver.find_element(By.CSS_SELECTOR, "#username").send_keys(coupang_id)
                    input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#password")))
                    input_field.click()
                    input_field.send_keys(Keys.CONTROL + "a")
                    input_field.send_keys(Keys.BACKSPACE)
                    driver.find_element(By.CSS_SELECTOR, "#password").send_keys(coupang_pw)
                    driver.find_element(By.CSS_SELECTOR,'#kc-login').click()
                

            except: pass

            try:
                WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.CSS_SELECTOR,"#startDateId"))) #클릭 시작일
            except:
                driver.refresh()
                WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.CSS_SELECTOR,"#startDateId"))) #클릭 시작일
            driver.find_element(By.CSS_SELECTOR, "#startDateId").click()

            before_Ym = today_tday.strftime("%Y년 %m월")
            before_d = str(int(today_tday.strftime("%d")))
            yesterday_Ym = today_yday.strftime("%Y년 %m월")
            yesterday_d = str(int(today_yday.strftime("%d")))

            firstCal = driver.find_elements(By.XPATH, "//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[2]")
            secondCal = driver.find_elements(By.XPATH, "//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[3]")

            # 시작날짜
            try:
                for i in firstCal:
                    # 텍스트를 줄 단위로 나누기
                    lines = (i.text).strip().split('\n')
                    if lines[0] == today_Tday년월:
                        print("OK")
                        i.find_element(By.XPATH, f"""//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/
                        div/div[2]/div/table//*[text()='{today_Tday일}']""").click()

                for i in secondCal:
                    # 텍스트를 줄 단위로 나누기
                    lines = (i.text).strip().split('\n')
                    if lines[0] == today_Tday년월:
                        print("OK")
                        i.find_element(By.XPATH, f"//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[3]/div/table//*[text()='{today_Tday일}']").click()

            except: pass
                

                    
            time.sleep(0.1)

            # 종료날짜
            try:
                for i in firstCal:
                    # 텍스트를 줄 단위로 나누기
                    lines = (i.text).strip().split('\n')
                    if lines[0] == today_Yday년월:
                        print("OK")
                        i.find_element(By.XPATH, f"//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/table//*[text()='{today_Yday일}']").click()

                for i in secondCal:
                    # 텍스트를 줄 단위로 나누기
                    lines = (i.text).strip().split('\n')
                    if lines[0] == today_Yday년월:
                        print("OK")
                        i.find_element(By.XPATH, f"//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[3]/div/table//*[text()='{today_Yday일}']").click()

            except: pass
            element = driver.find_element(By.CSS_SELECTOR, '#ad-reporting-app > div.self-service-ad-reporting-ui > div > div.sc-11l2gxs-0.fcpsUc > div.sc-ipia07-0.iCqAxH > div.panel-options > div.sc-19odvm9-0.kgfJLF > div.select-date-group')#기간 구분
            element.click() 
            ActionChains(driver).move_to_element_with_offset(element,5,75).click().perform() #클릭 일별
            time.sleep(0.3)

            driver.find_element(By.CSS_SELECTOR,'#ad-reporting-app > div.self-service-ad-reporting-ui > div > div.sc-11l2gxs-0.fcpsUc > div.sc-ipia07-0.iCqAxH > div.panel-options > div.sc-1jpf51e-0.hSjByk > div > div.campaign-picker-container > div > button > span.text').click() #캠페인 선택
            time.sleep(0.5)
            checkbox = driver.find_element(By.CSS_SELECTOR, "body > div.sc-1jpf51e-1.jljGiJ.popper > div > div.select-all-campaigns > label > span.ant-checkbox > input[type='checkbox']")
            if not checkbox.is_selected():
                checkbox.click()  # 체크박스가 체크되어 있지 않다면 클릭하여 체크합니다.
            driver.find_element(By.CSS_SELECTOR, "body > div.sc-1jpf51e-1.jljGiJ.popper > div > div.button-container > button.ant-btn.ant-btn-primary > span").click()
            time.sleep(0.3)

### 보고서 생성 실패하면 페이지 다시 로딩 후 생성
            try:
                try:
                    driver.find_element(By.CSS_SELECTOR, "#generateReport > span").click() #보고서 생성

                except:
                    driver.find_element(By.CSS_SELECTOR,'#ad-reporting-app > div.self-service-ad-reporting-ui > div > div.sc-11l2gxs-0.fcpsUc > div.sc-ipia07-0.iCqAxH > div.panel-options > div.sc-1jpf51e-0.hSjByk > div > div.campaign-picker-container > div > button > span.text').click() #캠페인 선택
                    time.sleep(0.5)
                    checkbox = driver.find_element(By.CSS_SELECTOR, "body > div.sc-1jpf51e-1.jljGiJ.popper > div > div.select-all-campaigns > label > span.ant-checkbox > input[type='checkbox']")

                    if not checkbox.is_selected():
                        checkbox.click()  # 체크박스가 체크되어 있지 않다면 클릭하여 체크합니다.

                    driver.find_element(By.CSS_SELECTOR, "body > div.sc-1jpf51e-1.jljGiJ.popper > div > div.button-container > button.ant-btn.ant-btn-primary > span").click()
                    driver.find_element(By.CSS_SELECTOR, "#generateReport > span").click() #보고서 생성

                time.sleep(1.5)

                if driver.find_element(By.CSS_SELECTOR, "#rc-tabs-0-panel-requestedReport > div > div.react-grid-Container > div > div > div:nth-child(2) > div > div > div:nth-child(2) > div:nth-child(1) > div:nth-child(5) > div > div > span > div").text == "생성 실패":
                    driver.find_element(By.CSS_SELECTOR, "#generateReport > span").click() 
                    time.sleep(5)

                

                # 보고서 다운로드
                element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#rc-tabs-0-panel-requestedReport > div > div.react-grid-Container > div > div > div:nth-child(2) > div > div > div:nth-child(2) > div:nth-child(1) > div:nth-child(6) > div > div > span > div > div:nth-child(2) > button > span"))) 

                # 다운로드 확인
                cnt = 1
                while cnt < 10:
                    current_file_count1 = count_files(download_folder)
                    element.click()
                    time.sleep(3)
                    current_file_count2 = count_files(download_folder)
                    if current_file_count1 != current_file_count2:
                        break

                    cnt += 1

                # check_download()
                time.sleep(1)
                try:
                    WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div.MuiDialog-root.sc-852clq-0.efPzRF > div.MuiDialog-container.MuiDialog-scrollPaper > div > div:nth-child(3) > button"))).click()
                except: pass

            except:
                driver.get(url_coupang_daily)

                try:
                    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.CSS_SELECTOR,"#startDateId"))) #클릭 시작일
                except:
                    driver.refresh()
                    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.CSS_SELECTOR,"#startDateId"))) #클릭 시작일
                driver.find_element(By.CSS_SELECTOR, "#startDateId").click()

                before_Ym = today_tday.strftime("%Y년 %m월")
                before_d = str(int(today_tday.strftime("%d")))
                yesterday_Ym = today_yday.strftime("%Y년 %m월")
                yesterday_d = str(int(today_yday.strftime("%d")))

                firstCal = driver.find_elements(By.XPATH, "//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[2]")
                secondCal = driver.find_elements(By.XPATH, "//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[3]")

                # 시작날짜
                try:
                    for i in firstCal:
                        # 텍스트를 줄 단위로 나누기
                        lines = (i.text).strip().split('\n')
                        if lines[0] == today_Tday년월:
                            print("OK")
                            i.find_element(By.XPATH, f"""//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/
                            div/div[2]/div/table//*[text()='{today_Tday일}']""").click()

                    for i in secondCal:
                        # 텍스트를 줄 단위로 나누기
                        lines = (i.text).strip().split('\n')
                        if lines[0] == today_Tday년월:
                            print("OK")
                            i.find_element(By.XPATH, f"//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[3]/div/table//*[text()='{today_Tday일}']").click()

                except: pass
                    

                        
                time.sleep(0.1)

                # 종료날짜
                try:
                    for i in firstCal:
                        # 텍스트를 줄 단위로 나누기
                        lines = (i.text).strip().split('\n')
                        if lines[0] == today_Yday년월:
                            print("OK")
                            i.find_element(By.XPATH, f"//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/table//*[text()='{today_Yday일}']").click()

                    for i in secondCal:
                        # 텍스트를 줄 단위로 나누기
                        lines = (i.text).strip().split('\n')
                        if lines[0] == today_Yday년월:
                            print("OK")
                            i.find_element(By.XPATH, f"//*[@id='ad-reporting-app']/div[2]/div/div[2]/div[1]/div[2]/div[3]/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[3]/div/table//*[text()='{today_Yday일}']").click()

                except: pass
                element = driver.find_element(By.CSS_SELECTOR, '#ad-reporting-app > div.self-service-ad-reporting-ui > div > div.sc-11l2gxs-0.fcpsUc > div.sc-ipia07-0.iCqAxH > div.panel-options > div.sc-19odvm9-0.kgfJLF > div.select-date-group')#기간 구분
                element.click() 
                ActionChains(driver).move_to_element_with_offset(element,5,75).click().perform() #클릭 일별
                time.sleep(0.3)

                driver.find_element(By.CSS_SELECTOR,'#ad-reporting-app > div.self-service-ad-reporting-ui > div > div.sc-11l2gxs-0.fcpsUc > div.sc-ipia07-0.iCqAxH > div.panel-options > div.sc-1jpf51e-0.hSjByk > div > div.campaign-picker-container > div > button > span.text').click() #캠페인 선택
                time.sleep(0.5)
                checkbox = driver.find_element(By.CSS_SELECTOR, "body > div.sc-1jpf51e-1.jljGiJ.popper > div > div.select-all-campaigns > label > span.ant-checkbox > input[type='checkbox']")
                if not checkbox.is_selected():
                    checkbox.click()  # 체크박스가 체크되어 있지 않다면 클릭하여 체크합니다.
                driver.find_element(By.CSS_SELECTOR, "body > div.sc-1jpf51e-1.jljGiJ.popper > div > div.button-container > button.ant-btn.ant-btn-primary > span").click()
                time.sleep(0.3)

                driver.find_element(By.CSS_SELECTOR, "#generateReport > span").click() #보고서 생성

                # 보고서 다운로드
                element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#rc-tabs-0-panel-requestedReport > div > div.react-grid-Container > div > div > div:nth-child(2) > div > div > div:nth-child(2) > div:nth-child(1) > div:nth-child(6) > div > div > span > div > div:nth-child(2) > button > span"))) 

                # 다운로드 확인
                cnt = 1
                while cnt < 10:
                    current_file_count1 = count_files(download_folder)
                    element.click()
                    time.sleep(3)
                    current_file_count2 = count_files(download_folder)
                    if current_file_count1 != current_file_count2:
                        break

                    cnt += 1

                # check_download()
                time.sleep(1)
                # 알림창 제거
                try:
                    WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div.MuiDialog-root.sc-852clq-0.efPzRF > div.MuiDialog-container.MuiDialog-scrollPaper > div > div:nth-child(3) > button"))).click()
                except: pass
=======
>>>>>>> 1e354683fa81441d941cc7e0d069d7eff51441ae


        

    def saveText(self):
        text = self.path_folder.text()
        text1 = self.chrome_path_folder.text()
        with open('saved_text.txt', 'w') as file:
            file.write(text)
            file.write("\n")
            file.write(text1)
        QMessageBox.information(self,'알림','저장되었습니다.')

        with open('checkbox_state.txt', 'w') as file:
            file.write(f"{self.haen_salesCafe24.isChecked()}\n")
            file.write(f"{self.love_salesCafe24.isChecked()}\n")
            file.write(f"{self.know_salesCafe24.isChecked()}\n")
            file.write(f"{self.zq_salesCafe24.isChecked()}\n")

            file.write(f"{self.haen_salesCoup.isChecked()}\n")
            file.write(f"{self.love_salesCoup.isChecked()}\n")
            file.write(f"{self.know_salesCoup.isChecked()}\n")

            file.write(f"{self.haen_salesNaver.isChecked()}\n")
            file.write(f"{self.love_salesNaver.isChecked()}\n")
            file.write(f"{self.know_salesNaver.isChecked()}\n")

            file.write(f"{self.haen_advtCoup.isChecked()}\n")
            file.write(f"{self.love_advtCoup.isChecked()}\n")
            file.write(f"{self.know_advtCoup.isChecked()}\n")

            file.write(f"{self.haen_advtNaver.isChecked()}\n")
            file.write(f"{self.love_advtNaver.isChecked()}\n")
            file.write(f"{self.know_advtNaver.isChecked()}\n")

            file.write(f"{self.haen_advtGFA.isChecked()}\n")
            file.write(f"{self.love_advtGFA.isChecked()}\n")
            file.write(f"{self.know_advtGFA.isChecked()}\n")

            file.write(f"{self.haen_advtPC.isChecked()}\n")
            file.write(f"{self.love_advtPC.isChecked()}\n")
            file.write(f"{self.know_advtPC.isChecked()}\n")

            file.write(f"{self.haen_visitors.isChecked()}\n")
            file.write(f"{self.love_visitors.isChecked()}\n")
            file.write(f"{self.know_visitors.isChecked()}\n")

            file.write(f"{self.haen_newMemb.isChecked()}\n")
            file.write(f"{self.love_newMemb.isChecked()}\n")
            file.write(f"{self.know_newMemb.isChecked()}\n")

    def loadCheckboxState(self):
        try:
            with open('checkbox_state.txt', 'r') as file:
                states = file.readlines()
                self.haen_salesCafe24.setChecked(states[0].strip() == 'True')
                self.love_salesCafe24.setChecked(states[1].strip() == 'True')
                self.know_salesCafe24.setChecked(states[2].strip() == 'True')
                self.zq_salesCafe24.setChecked(states[3].strip() == 'True')

                self.haen_salesCoup.setChecked(states[4].strip() == 'True')
                self.love_salesCoup.setChecked(states[5].strip() == 'True')
                self.know_salesCoup.setChecked(states[6].strip() == 'True')

                self.haen_salesNaver.setChecked(states[7].strip() == 'True')
                self.love_salesNaver.setChecked(states[8].strip() == 'True')
                self.know_salesNaver.setChecked(states[9].strip() == 'True')

                self.haen_advtCoup.setChecked(states[10].strip() == 'True')
                self.love_advtCoup.setChecked(states[11].strip() == 'True')
                self.know_advtCoup.setChecked(states[12].strip() == 'True')

                self.haen_advtNaver.setChecked(states[13].strip() == 'True')
                self.love_advtNaver.setChecked(states[14].strip() == 'True')
                self.know_advtNaver.setChecked(states[15].strip() == 'True')

                self.haen_advtGFA.setChecked(states[16].strip() == 'True')
                self.love_advtGFA.setChecked(states[17].strip() == 'True')
                self.know_advtGFA.setChecked(states[18].strip() == 'True')

                self.haen_advtPC.setChecked(states[19].strip() == 'True')
                self.love_advtPC.setChecked(states[20].strip() == 'True')
                self.know_advtPC.setChecked(states[21].strip() == 'True')

                self.haen_visitors.setChecked(states[22].strip() == 'True')
                self.love_visitors.setChecked(states[23].strip() == 'True')
                self.know_visitors.setChecked(states[24].strip() == 'True')

                self.haen_newMemb.setChecked(states[25].strip() == 'True')
                self.love_newMemb.setChecked(states[26].strip() == 'True')
                self.know_newMemb.setChecked(states[27].strip() == 'True')
                # 나머지 체크박스도 동일하게 불러옵니다.
        except FileNotFoundError:
            pass

    def loadText(self):
            try:
                with open('saved_text.txt', 'r') as f:
                    saved_text = f.read()
                    texts = saved_text.split("\n")

                    self.path_folder.setText(texts[0])
                    self.chrome_path_folder.setText(texts[1])

                    
            except FileNotFoundError:
                pass

    def login_info(self, target_word):
        try:
            with open('login_info.txt', 'r', encoding='utf-8') as f:
                lines = f.readlines()  # 파일의 모든 줄을 읽어 리스트로 저장

            # 모든 줄을 순회하면서 target_word 찾기
            for i, line in enumerate(lines):
                if target_word in line:  # 현재 줄에 target_word가 포함되어 있는지 확인
                    if i + 1 < len(lines):  # 다음 줄이 존재하는지 확인
                        print(lines[i + 1].strip())  # 다음 줄의 내용을 프린트 (공백 제거)
                        return(lines[i + 1].strip())
        except FileNotFoundError: print("cannot find login information.")
        
    def folderopen(self):
        fname = QFileDialog.getExistingDirectory(self,'폴더선택','')
        self.path_folder.setText(fname)
    
    def chromefolderopen(self):
        fname = QFileDialog.getExistingDirectory(self,'폴더선택','')
        self.chrome_path_folder.setText(fname)

    def my_exception_hook(exctype, value, traceback):
        # Print the error and traceback
        print(exctype, value, traceback)
        # Call the normal Exception hook after
        sys._excepthook(exctype, value, traceback)
        # sys.exit(1)

    # Back up the reference to the exceptionhook
    sys._excepthook = sys.excepthook

    # Set the exception hook to our wrapping function
    sys.excepthook = my_exception_hook

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = Rawdata_extractor()
    win.show()
    sys.exit(app.exec_())