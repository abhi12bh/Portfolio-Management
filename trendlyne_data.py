from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from collections import OrderedDict
from chrome_driver_connection import ChromeDriverConnection
import time
import pandas as pd
import xlwings as xw
import numpy as np
import math

class TrendlyneData:
    def __init__(self, driver_path, download_directory, headless=False):
        self.driver_connection = ChromeDriverConnection(driver_path, download_directory, headless)
        self.trendlyne_url_list = []
        self.durability_list = []
        self.valuation_list = []
        self.momentum_list = []
        self.forecast_price_list = []
        self.forecast_percent_list = []
        self.graham_number_list = []
        self.graham_ratio_list = []
        self.graham_s_comment_list = []
        self.graham_r_comment_list = []
        self.swot_strengths_list = []
        self.swot_weakness_list = []
        self.swot_opportunity_list = []
        self.swot_threats_list = []
        self.s1_list = []
        self.s2_list = []
        self.s3_list = []
        self.r1_list = []
        self.r2_list = []
        self.r3_list = []
        self.consensus_list = []
        self.consensus_no_list = []
        self.consensus_comment_list = []
      
                    
    def trendlyne_login(self, trendlyne_url, user_id, password):
        self.driver_connection.get(trendlyne_url)
#         time.sleep(3)
#         login_button = self.driver_connection.find_element((By.ID, "login-signup-btn"))
        login_button = self.driver_connection.wait_for_element((By.ID, "login-signup-btn"))
        login_button.click()
#         time.sleep(4)
        login_mail_trendlyne = self.driver_connection.wait_for_element((By.ID, "id_login"))
        login_mail_trendlyne.send_keys(user_id + Keys.ENTER)

        password_trendlyne = self.driver_connection.find_element((By.ID, "id_password"))
        password_trendlyne.send_keys(password + Keys.ENTER)
        
    def read_credentials_from_text(self, text_file):
        with open(text_file, 'r') as file:
            lines = file.readlines()
            user_id_trendlyne = lines[0].strip()
            password_trendlyne = lines[1].strip()
        return user_id_trendlyne, password_trendlyne
    
    
    def find_element_safe(self, by, selector):
        try:
            element = self.driver_connection.find_element((by, selector))
            return element
        except:
            return np.nan
    
    def handle_error(self, element, default_value=np.nan):
        try:
            if element:
                return element.text
        except Exception:
            pass
        return default_value
    
    def handle_split_error(self,element,default_value = np.nan):
        try:
            return element
        except:
            pass
        return default_value

    def overview_data(self):
        time.sleep(4)
        over_view_data_list = self.driver_connection.find_elements((By.CLASS_NAME, "fs1p8rem")) 
        for index in range(4): 
            try:
                value = self.handle_error(over_view_data_list[index])
            except IndexError:
                value = np.nan
            if index == 0:
                durability_value = value
            elif index == 1:
                valuation_value = value
            elif index == 2:
                momentum_value = value
            elif index == 3:
                forecast_price_value = value
                
    
        try:
            forecast_percent = self.driver_connection.find_element((By.CSS_SELECTOR, "div.bottom-insight > span"))
            forecast_percent_value = self.handle_split_error(self.handle_error(forecast_percent).split("%")[0])
        except:
            forecast_percent = self.find_element_safe(By.CSS_SELECTOR, "div.forecaster-block > a > div.bottom-right > div > span")
            forecast_percent_value = self.handle_split_error(self.handle_error(forecast_percent))


        self.durability_list.append(durability_value)
        self.valuation_list.append(valuation_value)
        self.momentum_list.append(momentum_value)
        self.forecast_price_list.append(forecast_price_value)
        self.forecast_percent_list.append(forecast_percent_value)

    
    def graham_analysis(self):
        graham_number = self.find_element_safe(By.XPATH,'//*[@id="stock_performance_parameters"]/div/div[2]/table/tbody/tr[8]/td[2]/span/span[1]')
        graham_number_value = self.handle_error(graham_number)

        graham_ratio = self.find_element_safe(By.XPATH,'//*[@id="stock_performance_parameters"]/div/div[2]/table/tbody/tr[7]/td[2]/span/span[1]')
        graham_ratio_value = self.handle_error(graham_ratio)

        graham_s_comment = self.find_element_safe(By.XPATH,'//*[@id="stock_performance_parameters"]/div/div[2]/table/tbody/tr[8]/td[1]/a/div[2]/span')
        graham_s_comment_value = self.handle_error(graham_s_comment)

        graham_r_comment = self.find_element_safe(By.XPATH,'//*[@id="stock_performance_parameters"]/div/div[2]/table/tbody/tr[7]/td[1]/a/div[2]/span')
        graham_r_comment_value = self.handle_error(graham_r_comment)

        self.graham_number_list.append(graham_number_value)
        self.graham_ratio_list.append(graham_ratio_value)
        self.graham_s_comment_list.append(graham_s_comment_value)
        self.graham_r_comment_list.append(graham_r_comment_value)

    def swot_analysis(self):
        swot_strength = self.find_element_safe(By.CSS_SELECTOR,"tbody > tr:nth-child(9) > td.stcard-rating.tl__st_card_rating--3TNfu")
        swot_strength_value = self.handle_error(swot_strength)

        swot_weakness = self.find_element_safe(By.CSS_SELECTOR,"tbody > tr:nth-child(10) > td.stcard-rating.tl__st_card_rating--3TNfu")
        swot_weakness_value = self.handle_error(swot_weakness)

        swot_opportunity = self.find_element_safe(By.CSS_SELECTOR,"tbody > tr:nth-child(11) > td.stcard-rating.tl__st_card_rating--3TNfu")
        swot_opportunity_value = self.handle_error(swot_opportunity)

        swot_threats = self.find_element_safe(By.CSS_SELECTOR,"tbody > tr:nth-child(12) > td.stcard-rating.tl__st_card_rating--3TNfu")
        swot_threats_value = self.handle_error(swot_threats)

        self.swot_strengths_list.append(swot_strength_value)
        self.swot_weakness_list.append(swot_weakness_value)
        self.swot_opportunity_list.append(swot_opportunity_value)
        self.swot_threats_list.append(swot_threats_value)
        
    def pivot_analysis(self):
        technical_trendlyne = self.driver_connection.wait_for_element((By.CSS_SELECTOR, "ul > li.nav-item.navs-heading-pills > div > a"))
        technical_trendlyne.click()
        time.sleep(2)

        technical_trendlyne_again = self.driver_connection.wait_for_element((By.CSS_SELECTOR, "ul > li.nav-item.navs-heading-pills > div > div > a:nth-child(1)"))
        technical_trendlyne_again.click()

        try:
            s1 = self.driver_connection.wait_for_element((By.CSS_SELECTOR, "div.pivot-table-rows > div:nth-child(1) > span.react-tech-value.positive"))
        except:
            s1_value = np.nan
        else:
            s1_value = self.handle_error(s1)

        s2 = self.find_element_safe(By.CSS_SELECTOR, "div.pivot-table-rows > div:nth-child(2) > span.react-tech-value.positive")
        s2_value = self.handle_error(s2)

        s3 = self.find_element_safe(By.CSS_SELECTOR, "div.pivot-table-rows > div:nth-child(3) > span.react-tech-value.positive")
        s3_value = self.handle_error(s3)

        r1 = self.find_element_safe(By.CSS_SELECTOR, "div.pivot-table-rows > div:nth-child(1) > span.react-tech-value.negative")
        r1_value = self.handle_error(r1)

        r2 = self.find_element_safe(By.CSS_SELECTOR, "div.pivot-table-rows > div:nth-child(2) > span.react-tech-value.negative")
        r2_value = self.handle_error(r2)

        r3 = self.find_element_safe(By.CSS_SELECTOR, "div.pivot-table-rows > div:nth-child(3) > span.react-tech-value.negative")
        r3_value = self.handle_error(r3)

        self.s1_list.append(s1_value)
        self.s2_list.append(s2_value)
        self.s3_list.append(s3_value)
        self.r1_list.append(r1_value)
        self.r2_list.append(r2_value)
        self.r3_list.append(r3_value)
        
    def forecaster_analysis(self):

        try:
            forecaster = self.driver_connection.find_element((By.CSS_SELECTOR, "ul > li:nth-child(3) > a > img"))
            forecaster.click()
        except:
            consensus = np.nan
            consensus_no = np.nan
            consensus_comment = np.nan
        else:    
            
            time.sleep(3)
            consensus_element = self.find_element_safe(By.CLASS_NAME, "title1")
            consensus = self.handle_error(consensus_element)

            try:
                consensus_no_element = self.driver_connection.find_element((By.CSS_SELECTOR, ".highlight.fw500"))
                consensus_no = self.handle_split_error(self.handle_error(consensus_no_element).split(" ")[0])
            except:
                consensus_no = np.nan     

            consensus_comment_element = self.find_element_safe(By.CLASS_NAME, "insight-title")
            consensus_comment = self.handle_error(consensus_comment_element)
        
        finally:
            self.consensus_list.append(consensus)
            self.consensus_no_list.append(consensus_no)
            self.consensus_comment_list.append(consensus_comment)

    def analyze_stock(self, stock):
        search_button = self.driver_connection.find_element((By.CLASS_NAME, "tl-navbar-search-input"))

        search_button.send_keys(stock + Keys.ENTER)

        equity_button = self.driver_connection.wait_for_element((By.CSS_SELECTOR, '#ui-id-1 > li:nth-child(1)'))
        if equity_button.text =="Equity":
            time.sleep(5)
            next_button = self.driver_connection.wait_for_element((By.CSS_SELECTOR, '#ui-id-1 > li:nth-child(2)'))
            next_button.click()
            trendlyne_url = self.driver_connection.get_url()
            self.trendlyne_url_list.append(trendlyne_url)
            # Call other analysis functions
            self.overview_data()
            time.sleep(0.5)
            
            self.swot_analysis()
            time.sleep(0.5)
            self.graham_analysis()            
            self.pivot_analysis()
            self.forecaster_analysis()

        else:
            self.durability_list.append(np.nan)
            self.valuation_list.append(np.nan)
            self.momentum_list.append(np.nan)
            self.forecast_price_list.append(np.nan)
            self.forecast_percent_list.append(np.nan)
            self.graham_number_list.append(np.nan)
            self.graham_ratio_list.append(np.nan)
            self.graham_s_comment_list.append(np.nan)
            self.graham_r_comment_list.append(np.nan)
            self.swot_strengths_list.append(np.nan)
            self.swot_weakness_list.append(np.nan)
            self.swot_opportunity_list.append(np.nan)
            self.swot_threats_list.append(np.nan)
            self.s1_list.append(np.nan)
            self.s2_list.append(np.nan)
            self.s3_list.append(np.nan)
            self.r1_list.append(np.nan)
            self.r2_list.append(np.nan)
            self.r3_list.append(np.nan)
            self.consensus_list.append(np.nan)
            self.consensus_no_list.append(np.nan)
            self.consensus_comment_list.append(np.nan)
            self.trendlyne_url_list.append(np.nan)
            self.driver_connection.refresh()
            time.sleep(4)
            
        
    def create_dataframe(self):
        data_dict = {
            "Durability": self.durability_list,
            "Valuation": self.valuation_list,
            "Momentum": self.momentum_list,
            "Forecast Price": self.forecast_price_list,
            "Forecast Percent": self.forecast_percent_list,
            "Forecast Consensus": self.consensus_list,
            "Consensus No": self.consensus_no_list,
            "Consensus Comment": self.consensus_comment_list,            
            "R1": self.r1_list,
            "R2": self.r2_list,
            "R3": self.r3_list,
            "S1": self.s1_list,
            "S2": self.s2_list,
            "S3": self.s3_list,
            "S": self.swot_strengths_list,
            "W": self.swot_weakness_list,
            "O": self.swot_opportunity_list,
            "T": self.swot_threats_list,
            "Graham Number": self.graham_number_list,
            "Graham Ratio": self.graham_ratio_list,
            "GS Comment": self.graham_s_comment_list,
            "GR Comment": self.graham_r_comment_list,
        }

        df = pd.DataFrame(data_dict)
        return df

    def trendlyne_data_into_excel(self, stock_list_sheet, stock_sheet_list_cell,trendline_sheet,trendline_sheet_cell,stock_column = "A6",
                                  trendlyne_url_column= "C6", excel_file_path=None):
        if excel_file_path:
            xt = xw.Book(excel_file_path)
        else:
            xt = xw.Book.caller()
        stock_sheet_list_name = xt.sheets[stock_list_sheet]
        trendline_analsis_sheeet = xt.sheets[trendline_sheet]
        stock_name_list = stock_sheet_list_name.range(stock_sheet_list_cell).options(expand="down").value
        filtered_stock_name_list = [name for name in stock_name_list if name not in ('', ',')]
        filtered_stock_name_list = sorted(list(OrderedDict.fromkeys(filtered_stock_name_list)))

        for stock in filtered_stock_name_list:
            self.analyze_stock(stock)
        
        time.sleep(1)        
        data_frame = self.create_dataframe()
        data_frame.set_index(data_frame.columns[0], inplace=True)
        trendline_analsis_sheeet.range(trendline_sheet_cell).options(expand="down").value = data_frame
        trendline_analsis_sheeet.range(stock_column).options(transpose=True).value = filtered_stock_name_list
        trendline_analsis_sheeet.range(trendlyne_url_column).options(transpose=True).value = self.trendlyne_url_list
    
    
    def logout(self):
        try:
            logout_button = self.driver_connection.find_element((By.CSS_SELECTOR,"#topnav-right-col > div > div:nth-child(3) > ul > li > a > img"))
            logout_button.click()
            time.sleep(1)
            logout = self.driver_connection.find_element((By.CSS_SELECTOR,"#topnav-right-col > div > div:nth-child(3) > ul > li > div > a.dropdown-item.modal-trigger"))
            logout.click()
            time.sleep(1)
            logout_confirmation = self.driver_connection.find_element((By.CSS_SELECTOR,"#basemodal > div > div > div > div.logout-content > div.container-fluid.logout-modal-content > form > button"))
            logout_confirmation.click()
            time.sleep(2)
            self.driver_connection.close()
            print("Logged out successfully.")
        except Exception as e:
            print("Error during logout:", str(e))
        
    
    def close_connection(self):
        self.driver_connection.close()

