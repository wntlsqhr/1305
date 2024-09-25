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
import functools
import threading
import openpyxl
import gspread
import json
import time


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

driver.get("https://partner.alps.llogis.com/main/pages/sec/authentication")
input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#principal > input:nth-child(2)")))
input_field.click()
time.sleep(1)
input_field.send_keys(Keys.CONTROL + "a")
input_field.send_keys(Keys.BACKSPACE)
input_field.send_keys("305895")

input_field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#credential > input:nth-child(2)")))
input_field.click()
input_field.send_keys(Keys.CONTROL + "a")
input_field.send_keys(Keys.BACKSPACE)
input_field.send_keys("thsrkfka2!")

driver.find_element(By.CSS_SELECTOR, "#btn-login").click()
WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.col:nth-child(3)"))).click()
driver.find_element(By.CSS_SELECTOR, "div.row:nth-child(2) > div:nth-child(3) > ul:nth-child(2) > li:nth-child(2) > a:nth-child(1)").click()
WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[contains(text(), "조회")]'))).click()
text = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#grdList01 > div:nth-child(1) > div:nth-child(1) > canvas:nth-child(1)"))).text
print(text)
input()



