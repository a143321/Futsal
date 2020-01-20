# coding: utf_8

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import matplotlib.dates as mdates
import chrome_driver
import excel_access

STOCK_WEB_SITE_PATH = 'https://yoyaku.harp.lg.jp/resident/Login.aspx'

#chrome = chrome_driver.ChromeWebDriver()
#chrome.change_url(STOCK_WEB_SITE_PATH)

excel = excel_access.ExcelAccess("database.xlsx")


login_info = excel.get_excel_data()

print(login_info[0][0])
print(login_info[0][1])
print(login_info[0][2])
print(login_info[0][3])