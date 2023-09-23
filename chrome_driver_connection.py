from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class ChromeDriverConnection:
    def __init__(self, driver_path, download_directory, headless=False):
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument('--headless')
        chrome_prefs = {'download.default_directory': download_directory}
        options.add_experimental_option('prefs', chrome_prefs)  
        self.driver = webdriver.Chrome(options=options)  
        self.driver.maximize_window()
    def get(self, url):
        self.driver.get(url)

    def find_element(self, selector):
        return self.driver.find_element(*selector)

    def find_elements(self, selector):
        return self.driver.find_elements(*selector)

    def close(self):
        self.driver.quit()
        
    def refresh(self):
        self.driver.refresh()

    def get_url(self):
        current_url = self.driver.execute_script("return window.location.href")
        return current_url
    
    def wait_for_element(self, selector, timeout=10):
        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located(selector)
            )
            return element
        except:
            raise Exception(f"Element with selector {selector} not found within {timeout} seconds.")
    
