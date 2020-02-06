# coding: utf_8

import chrome_driver
import excel_access

STOCK_WEB_SITE_PATH = 'https://yoyaku.harp.lg.jp/resident/Login.aspx'

# Excelファイルを文字列として読み込む
excel = excel_access.ExcelAccess("database.xlsx", 2)

# 全データを取得する 
print(excel.ReadAll())

# セルの値を取得
print(excel.ReadCell(1,1))

# 行の値を取得
for item in excel.ReadRow(0):
    print(item)

# 列の値を取得
for item in excel.ReadColumn(0):
    print(item)

# 範囲の値を取得
for item in excel.ReadArea((0,0),(2,2)):
    print(item)

# 全データを取得する 
print(excel.ReadAll())

# セルに値を設定する
excel.WriteCell(0, 3,"Password21")

# 全データを取得する 
print(excel.ReadAll())

# Excelに保存する
excel.SaveExcel()
