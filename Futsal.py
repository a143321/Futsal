# coding: utf_8

import chrome_driver
import excel_access
from user_data import User
from user_data import Facility
import pandas as pd
from bs4 import BeautifulSoup
import time



# Excelファイルを文字列として読み込む
excel = excel_access.ExcelAccess("database.xlsx", 0)

# ログイン情報を格納する
user_list = []

for item in excel.ReadAll():
    user_id = item[2]
    password = item[3]
    user_list.append(User(user_id, password))



WEB_SITE_PATH = 'https://yoyaku.harp.lg.jp/resident/Menu.aspx?lgCode=011002'

driver = chrome_driver.SeleniumChromeAccess()
driver.change_url(WEB_SITE_PATH)

# 施設予約ログインボタンクリック
driver.click_button_using_XPath("/html/body/form/div[5]/div/div[1]/div/table/tbody/tr[4]/td[1]/div/ul/li[1]/input")


input_user_id_xml = "/html/body/form/div[5]/div/div[1]/div/table[1]/tbody/tr[2]/td/input"
input_password_xml = "/html/body/form/div[5]/div/div[1]/div/table[1]/tbody/tr[3]/td/input"
push_button_login =  "/html/body/form/div[5]/div/div[1]/div/table[1]/tbody/tr[4]/th/input"

driver.input_textbox_using_XPath(input_user_id_xml, user_list[0].user_id)
driver.input_textbox_using_XPath(input_password_xml, user_list[0].password)
driver.click_button_using_XPath(push_button_login)

# 空き施設検索・抽選申込・利用申込クリック
BUTTON_APPLICATION = "/html/body/form/div[5]/div/div[1]/div/table[1]/tbody/tr[4]/td[1]/div/ul/li[1]/input"
driver.click_button_using_XPath(BUTTON_APPLICATION)

# 利用目的分類
# スポーツ（屋内
driver.select_dropdown_list_using_id("ctl00_ContentPlaceHolder1_ShinseiKumiawaseInp1_drpPurposeBunrui", "04")

# 利用目的
# サロンフットボール・フットサル
driver.select_dropdown_list_using_id("ctl00_ContentPlaceHolder1_ShinseiKumiawaseInp1_drpPurpose", "064")

# 利用開始日
driver.input_textbox_using_XPath("/html/body/form/div[5]/div/div[1]/div/table[2]/tbody/tr[4]/td/div/input[1]", "2020/3/1")

# 利用終了日
driver.input_textbox_using_XPath("/html/body/form/div[5]/div/div[1]/div/table[2]/tbody/tr[4]/td/div/input[3]", "2020/3/31")

# 施設検索ボタンを押す
driver.click_button_using_XPath("/html/body/form/div[5]/div/div[1]/div/table[2]/tbody/tr[6]/th/input")

html=driver.get_page_source()

soup=BeautifulSoup(html,'html.parser')
results=soup.find_all('span')

print(f"ページ数 : {results[0]}")

item_list = []

for i, item in enumerate(results):
    if i == 0 or i % 2  == 1:
        continue
    # 学校名だけ取得
    item_list.append(item.text)

school_list = []

for item in item_list:
    school_list.append(item)

# 検索対象文字列

search_list = []
search_list.append("伏見中学校")
search_list.append("白石中学校")
search_list.append("米里中学校")

target = search_list[0]

print(f"target = {target}")
target_index = -1

if target in school_list:
    target_index = school_list.index(target)
    print(f"target_index = {target_index}")
else:
    print(f"not found = {target in school_list}")

if target_index != -1:
    serch_xml = f"/html/body/form/div[5]/div/div[1]/div/table[3]/tbody/tr[{target_index + 2}]/td[1]/div/a"
    print(f"serch_xml = {serch_xml}")

    driver.click_button_using_XPath("/html/body/form/div[5]/div/div[1]/div/table[3]/tbody/tr[2]/td[1]/div/a")

    time.sleep(3)

    print("test1")

    html2=driver.get_page_source()
    soup2=BeautifulSoup(html2,'html.parser')
    results2=soup2.find_all('a')

    print("test2")

    select_list = []
    for item in results2:
        if item.get('style') == "color:Black":
            select_list.append(item)

    for item in select_list:
        print(f"result = {item.text}")

    # https://qiita.com/VA_nakatsu/items/0095755dc48ad7e86e2f
    driver.click_button_using_XPath("//a[@title='3月2日']")

