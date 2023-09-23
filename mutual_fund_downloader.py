import time
from selenium.webdriver.common.keys import Keys
import xlwings as xw
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from chrome_driver_connection import ChromeDriverConnection  # Importing the ChromeDriverConnection class
from selenium import webdriver

class MutualFundDownloader:
    def __init__(self, driver_path, download_directory, headless=False):
        self.connection = ChromeDriverConnection(driver_path, download_directory, headless)

    def download_mutual_fund_files(self, url,  sheet_name, start_cell, excel_file_path=None,):
        self.connection.get(url)
        if excel_file_path:
            xt = xw.Book(excel_file_path)
        else:
            xt = xw.Book.caller()

            
        mutual_fund_sheet = xt.sheets[sheet_name]
        mutual_fund_list = mutual_fund_sheet.range(start_cell).options(expand="down").value
        for mutual_fund_name in mutual_fund_list:
            input_name = self.connection.find_element((By.ID, "srch-term"))
            input_name.clear()
            time.sleep(2)
#             for char in mutual_fund_name[:-8]:
#                 input_name.send_keys(char)
#                 time.sleep(0.5) 
#             time.sleep(3)
#             drop_down = self.connection.find_element((By.XPATH,"/html/body/ul[2]/li[1]"))
            input_name.send_keys(mutual_fund_name)
            input_name.send_keys(Keys.ENTER)
            time.sleep(4)
            
#             drop_down.click()

            download_button = self.connection.find_element((By.ID, "download_smf"))
            time.sleep(2)
            download_button.click()
            time.sleep(2)

    def close(self):
        self.connection.close()
