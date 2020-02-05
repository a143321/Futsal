# coding: utf_8

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import matplotlib.dates as mdates
import chrome_driver
import excel_access

STOCK_WEB_SITE_PATH = 'https://yoyaku.harp.lg.jp/resident/Login.aspx'

# Excelファイルを文字列として読み込む
excel = excel_access.ExcelAccess("database.xlsx")

# セルの値を取得
print(excel.ReadCell(1,1))

# 行の値を取得
for item in excel.ReadRow(1):
    print(item)

# 列の値を取得
for item in excel.ReadColumn(1):
    print(item)

# 範囲の値を取得
for item in excel.ReadArea((2,2),(3,3)):
    print(item)


excel.WriteCell(1,1,"test")

print(excel.ReadAll())

excel.SaveExcel()
