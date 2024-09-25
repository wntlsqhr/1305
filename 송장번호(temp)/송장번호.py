from PyQt5.QtGui import QFont, QIcon, QStandardItemModel, QStandardItem, QTextBlock, QTextCursor
from PyQt5.QtCore import Qt, QThread, QObject, pyqtSignal, QCoreApplication
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
from gspread_formatting import *
from gspread.exceptions import APIError
from openpyxl import Workbook, load_workbook
from selenium import webdriver
from datetime import datetime, date, timedelta
from PyQt5.QtWidgets import *
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
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
import shutil


# 엑셀 파일 불러오기 (파일 경로를 입력하세요)
file_path = '골드레인 코퍼레이션 배송여부 확인 요청 건_주문번호.xlsx'
df = pd.read_excel(file_path)

# E열 데이터를 모두 가져오기 (E열은 5번째 열, 인덱스로는 4)
e_column_data = df.iloc[:, 4].tolist()

# 결과 출력
print(e_column_data)
print(len(e_column_data))



# 크롬 On
### chromedriver_autoinstaller.install() 사용 추가
chromedriver_path = chromedriver_autoinstaller.install()
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_argument("--start-maximized") #최대 크기로 시작
# chrome_options.add_argument('--incognito')
# chrome_options.add_argument('--window-size=1920,1080')  
# chrome_options.add_argument('--headless')
chrome_options.add_experimental_option('detach', True)

user_data = 'C:\\Users\\A\\AppData\\Local\\Google\\Chrome\\User Data1'
chrome_options.add_argument(f"user-data-dir={user_data}")
chrome_options.add_argument("--profile-directory=Profile 1")

user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
headers = {'user-agent' : user_agent}

driver = webdriver.Chrome(
    service=Service(chromedriver_path),
    options=chrome_options
)



driver.get("https://eclogin.cafe24.com/Shop/")

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
driver.find_element(By.CSS_SELECTOR, "#mall_id").send_keys("woo8425")

# PW
input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#userpasswd")))
input_field.click()
input_field.send_keys(Keys.CONTROL + "a")
input_field.send_keys(Keys.BACKSPACE)
driver.find_element(By.CSS_SELECTOR, "#userpasswd").send_keys("1305gold^^")

# 로그인클릭
driver.find_element(By.CSS_SELECTOR,'#frm_user > div > div.mButton > button').click()

time.sleep(1)

driver.get("https://woo8425.cafe24.com/admin/php/shop1/s_new/order_list.php")





cnt = 1
while cnt < 864:
    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#sBaseSearchBox')))
    inputElement = driver.find_element(By.CSS_SELECTOR, "#sBaseSearchBox")
    inputElement.click()
    time.sleep(0.1)
    inputElement.send_keys(Keys.CONTROL + "a")
    inputElement.send_keys(Keys.BACKSPACE)
    inputElement.send_keys(e_column_data[cnt-1])
    inputElement.send_keys(Keys.ENTER)
    
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, 'copyarea_0')))
    driver.find_element(By.ID, "copyarea_0").click()

### 새 창 으로 스위치
    try:
        #새 창 대기
        current_window_handle = driver.current_window_handle

        new_window_handle = None
        start_time = time.time()
        timeout = 5  # 15초 타임아웃
        while not new_window_handle and (time.time() - start_time) < timeout:
            for handle in driver.window_handles:
                if handle != current_window_handle:
                    new_window_handle = handle
                    break
        if not new_window_handle:
            print("New window did not open within the timeout. Exiting.")
            break

    except:
        driver.find_element(By.ID, "copyarea_0").click()

        #새 창 대기
        current_window_handle = driver.current_window_handle

        new_window_handle = None
        start_time = time.time()
        timeout = 5  # 15초 타임아웃
        while not new_window_handle and (time.time() - start_time) < timeout:
            for handle in driver.window_handles:
                if handle != current_window_handle:
                    new_window_handle = handle
                    break
        if not new_window_handle:
            print("New window did not open within the timeout. Exiting.")
            break
    
    driver.switch_to.window(driver.window_handles[1])

    element = driver.find_element(By.CSS_SELECTOR, "#tabNumber > div.mBoard.typeOrder.gScroll.gCellSingle > table > tbody > tr:nth-child(2) > td > div > div.control > input.fText")
    text_value = element.get_attribute('value')  # 'value' 속성 가져오기
    print(text_value)

    excel_path = "골드레인 코퍼레이션 배송여부 확인 요청 건_주문번호.xlsx"  # 엑셀 파일 경로
    workbook = load_workbook(excel_path)
    sheet = workbook.active  # 활성 시트 선택

    cell = f'M{cnt+1}'
    sheet[cell] = text_value

    


    if driver.find_element(By.CSS_SELECTOR, "#tabNumber > div.mBoard.typeOrder.gScroll.gCellSingle > table > tbody > tr:nth-child(1) > td:nth-child(10) > a.txtLink.eLayerClick").text == "배송완료":

        cell = f'K{cnt+1}'
        sheet[cell] = 'Y'

    else:
        cell = f'K{cnt+1}'
        sheet[cell] = 'N'

    #tabNumber > div.mBoard.typeOrder.gScroll.gCellSingle > table > tbody > tr:nth-child(1) > td:nth-child(10) > a.txtLink.eLayerClick

    # 3. 엑셀 저장 및 닫기
    workbook.save(excel_path)
    workbook.close()

    driver.close()
    driver.switch_to.window(driver.window_handles[0])

    cnt += 1





















