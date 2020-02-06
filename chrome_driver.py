from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os

MAX_WAIT_TIME_SEC = 60

class SeleniumChromeAccess():

    def __init__(self):
        self.driver = self.get_web_driver()

    def get_web_driver(self):
        CHROME_DRIVER_PATH = "chromedriver.exe"
        options = Options()
        return webdriver.Chrome(executable_path=CHROME_DRIVER_PATH, options=options)

    def change_url(self, url_path):
        self.driver.get(url_path)

    # XPathを用いて、ボタンをクリックする
    def click_button_using_XPath(self, element_XPath : str):
        
        element = WebDriverWait(self.driver, MAX_WAIT_TIME_SEC).until(
            EC.presence_of_element_located((By.XPATH, element_XPath))
        )
        
        element.click()
        
        return element

    # XPathを用いて、テキストボックスに文字列を入力する
    def input_textbox_using_XPath(self, element_XPath : str, send_code : str):
        
        element = WebDriverWait(self.driver, MAX_WAIT_TIME_SEC).until(
            EC.presence_of_element_located((By.XPATH, element_XPath))
        )
        
        # 検索テキストボックスをクリアする
        for item in range(0,100) :
            element.send_keys(Keys.BACK_SPACE)
                                    
        element.send_keys(send_code)
        
        return element

    # 要素を取得する関数
    def get_element(self, element_XPath : str):
        
        element = WebDriverWait(self.driver, MAX_WAIT_TIME_SEC).until(
            EC.presence_of_element_located((By.XPATH, element_XPath))
        )
        return element

    # 要素を取得する関数
    def get_element_by_class_name(self, path : str):
        
        element = WebDriverWait(self.driver, MAX_WAIT_TIME_SEC).until(
            EC.presence_of_element_located((By.CLASS_NAME, path))
        )
        return element

    # 要素のテキスト文を取得する関数
    def get_element_text(self, element_XPath : str) -> str:
        return self.get_element(element_XPath).text

    # キャプチャを取得する
    def get_screen_shot(self, folder_path, file_path):

        folder_path = folder_path.replace(os.path.sep, '/')
        
        os.makedirs(folder_path, exist_ok=True)

        full_path = os.path.join(folder_path, file_path)

        self.driver.save_screenshot(full_path)







