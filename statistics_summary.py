
# https://medium.com/@great.investor/how-to-scrape-data-from-yahoo-finance-with-python-and-selenium-81e8577db69c
# 
# ============================================================
#  Yahoo Finance Statistics Scraper
#  Compatible with Windows / Render / Headless environments
# ============================================================

import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
# ============================================================
# 1️⃣ WebDriver configuration (Render-ready)
# ============================================================

from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

# Edge or Chrome both work — Render provides a Chromium runtime by default
options = Options()

# Run visible by default (local dev)
headless = os.getenv("HEADLESS", "false").lower() in ("true", "1", "yes")
if headless:
    options.add_argument("--headless=new")

options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-extensions")
options.add_argument("--remote-debugging-port=9222")
options.add_argument("--window-size=1920,1080")
options.add_argument("--log-level=3")

service = Service(r"C:\Windows\System32\msedgedriver.exe")
driver = webdriver.Edge(service=service, options=options)
driver.implicitly_wait(5)
# ============================================================
# 2️⃣ Initialize DataFrame and symbols
# ============================================================
taglist = ["MU", "MSFT"]
StatisticsSummary = pd.DataFrame()

# Optional: base URL
BASE_URL = "https://finance.yahoo.com/lookup"
driver.get(BASE_URL)
time.sleep(2)

# Search for an elements of taglist to disable the button that appears after first search
def SearchElements(tag):
    SearchBar = driver.find_element(By.XPATH, "//form[@class='_yb_1vht69d _yb_1ndj1ez ybar-enable-search-ui ybar-assist_billboard_v2 _yb_oow1q0 _yb_9vl95e _yb_1o792pg']//input[@class='_yb_d7nywr _yb_owcm9n _yb_ikgg1k finsrch-inpt ']")
    SearchBar.send_keys(tag)
    SearchBar.send_keys(Keys.ENTER)
    time.sleep(2)

def getStatistics(tag):
    global StatisticsSummary

    # statistic button을 찾음
    StatisticsBtn = driver.find_element(By.XPATH, "//li[@class='yf-1mwz2jq']/a[@category='key-statistics']")
    action = ActionChains(driver)
    action.click(on_element=StatisticsBtn)
    action.perform()
    time.sleep(2)

    # VAluation Measures 테이블 첫번째 행 레이블(thead) 추출
    labelElements = driver.find_elements(
    By.XPATH, "//div[@class='table-container yf-kbx2lo']//table[@class='table yf-kbx2lo']//thead//tr[@class='yf-kbx2lo']//th[@class='yf-kbx2lo']")

    # Extract visible text from each element
    labelTexts = [el.text for el in labelElements if el.text.strip() != '']
    # print(f"labelTexts: {labelTexts}")

    # VAluation Measures 테이블 2번째행과 이후에 있는 데이터(tbody) 추출
    bodyElements = driver.find_elements(
        By.XPATH, "//div[@class='table-container yf-kbx2lo']//table[@class='table yf-kbx2lo']//tbody//tr[@class='yf-kbx2lo alt' or @class='yf-kbx2lo']//td[@class='yf-kbx2lo']")

    dataList = [value.text for value in bodyElements]

    # Create dictionary to store time label
    time_values = labelTexts[:] 

    # Create dictionary to store statistics
    column_labels = dataList[0::7]

    # Extract the dataList rows (all elements except the labels, reshaped into columns)
    data_values = [dataList[i::7] for i in range(1, 7)]

    # Create the DataFrame for tag label
    df = pd.DataFrame(data=data_values, columns= column_labels)   

    # Add time and tag columns directly to df
    df.insert(0, 'tag', tag)            # add tag as second column
    df.insert(1, 'time', time_values)   # add as first column

    # Append the combined DataFrame to StatisticsSummary
    StatisticsSummary = pd.concat([StatisticsSummary, df], ignore_index=True)

for tag in taglist:
    SearchElements(tag)
    getStatistics(tag)
# Print the DataFrame to verify
print(StatisticsSummary)
print(f"✅ Finished {tag} ({len(StatisticsSummary)} rows total)")

# Use os.path.expanduser
output_path = os.path.join(os.path.dirname(__file__), "StatisticsSummary.xlsx")
StatisticsSummary.to_excel(output_path, sheet_name="statistics_summary", index=False)
# add a polite driver close
driver.quit()
