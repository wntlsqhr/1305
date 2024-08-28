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


# while True:
try:
    driver.get("https://advertising.coupang.com/marketing-reporting/billboard/reports/pa")  # 로그인 시작
    if driver.find_element(By.CSS_SELECTOR, "body > pre"):
        driver.get("https://advertising.coupang.com/marketing-reporting/billboard/reports/pa")  # 요소가 존재하면 페이지를 다시 로드
except NoSuchElementException:
    # 요소가 없을 때 처리할 로직
    pass

try:
    driver.get("https://advertising.coupang.com/marketing-reporting/billboard/reports/pa")  # 로그인 시작
    if driver.find_element(By.CSS_SELECTOR, "body > h1"):
        driver.get("https://advertising.coupang.com/marketing-reporting/billboard/reports/pa")  # 요소가 존재하면 페이지를 다시 로드
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

print("rmsdn0417")
print("zmsdn44^^^")
input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#username")))
input_field.click()
time.sleep(0.7)
input_field.send_keys(Keys.CONTROL + "a")
input_field.send_keys(Keys.BACKSPACE)
driver.find_element(By.CSS_SELECTOR, "#username").send_keys("rmsdn0417")
input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#password")))
input_field.click()
input_field.send_keys(Keys.CONTROL + "a")
input_field.send_keys(Keys.BACKSPACE)
driver.find_element(By.CSS_SELECTOR, "#password").send_keys("zmsdn44^^^")
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
        driver.find_element(By.CSS_SELECTOR, "#username").send_keys("rmsdn0417")
        input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#password")))
        input_field.click()
        input_field.send_keys(Keys.CONTROL + "a")
        input_field.send_keys(Keys.BACKSPACE)
        driver.find_element(By.CSS_SELECTOR, "#password").send_keys("zmsdn44^^^")
        driver.find_element(By.CSS_SELECTOR,'#kc-login').click()
    

except: pass

try:
    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.CSS_SELECTOR,"#startDateId"))) #클릭 시작일
except:
    print("보고서페이지 로딩실패... retry")

    driver.close()
    driver = webdriver.Chrome(
    service=Service(chromedriver_path),
    options=chrome_options
    )
    driver.get("https://advertising.coupang.com/marketing-reporting/billboard/reports/pa")  # 로그인 시작

    # driver.find_element(By.CSS_SELECTOR, "#cap-sidebar > nav > ul > li.ant-menu-item.ant-menu-item-selected > span.ant-menu-title-content > span").click()

    # WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ad-reporting-app > div.self-service-ad-reporting-ui > div > div.root > ul > li:nth-child(2) > div > span"))).click()

    # WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[contains(text(), "쿠팡 상품광고 보고서")]'))).click()


    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.CSS_SELECTOR,"#startDateId"))) #클릭 시작일

driver.find_element(By.CSS_SELECTOR, "#startDateId").click()
driver.quit()

        

