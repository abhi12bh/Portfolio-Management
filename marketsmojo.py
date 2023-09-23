from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from collections import OrderedDict
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from chrome_driver_connection import ChromeDriverConnection
import time
import pandas as pd
import xlwings as xw
import numpy as np

class MarketsMojoData:
    def __init__(self, driver_path, download_directory, headless=False):
        self.driver_connection = ChromeDriverConnection(driver_path, download_directory, headless)
        self.cmp_list = []
        self.mojo_score_list = []
        self.mojo_stock_yes_list = []
        self.mojo_stock_since_list = []
        self.dashboard_recommendation_list = []
        self.dashboard_comment_list = []
        self.recommendated_technical_list = []
        self.recommendated_technical_comment_list = []
        self.quality_score_list = []
        self.quality_comments_list = []
        self.valuation_score_list = []
        self.financial_trend_positive_list = []
        self.financial_trend_negative_list = []
        self.financial_trend_total_list = []
        self.financial_trend_status_list = []
        self.one_day_list = []
        self.one_week_list = []
        self.one_month_list = []
        self.three_month_list = []
        self.six_month_list = []
        self.ytd_list = []
        self.one_year_list = []
        self.two_year_list = []
        self.three_year_list = []
        self.four_year_list = []
        self.five_year_list = []
        self.ten_year_list = []
        self.majority_shareholders_list = []
        self.pledged_list = []
        self.mf_list = []
        self.mf_scheme_percent_list = []
        self.fii_list = []
        self.fii_percent_list = []
        self.promoter_list = []
        self.promoter_percent_list = []
        self.highest_public_holder_list = []
        self.highest_public_holder_percent_list = []
        self.individual_investor_list = []
        self.markets_mojo_url = []
           

    def markets_mojo_login(self,markets_mojo_url, user_id, password, temporary_url=False):
        self.driver_connection.get(markets_mojo_url)
        time.sleep(3)
        if temporary_url:
            username_mojo = self.driver_connection.find_element((By.ID,"emailID"))
            password_mojo = self.driver_connection.find_element((By.ID,"regpassword"))
            remember_me = self.driver_connection.find_element((By.ID,"rememberMe"))
            remember_me.click()
            time.sleep(1)
            username_mojo.send_keys(user_id + Keys.ENTER)
            password_mojo.send_keys(password + Keys.ENTER)
        else:   
            username_mojo = self.driver_connection.find_element((By.ID,"username"))
            password_mojo = self.driver_connection.find_element((By.ID,"password"))
            login = self.driver_connection.find_element((By.CSS_SELECTOR,'form > div:nth-child(2) > button'))
            username_mojo.send_keys(user_id + Keys.ENTER)
            password_mojo.send_keys(password + Keys.ENTER)
            close_add = self.driver_connection.find_element((By.CLASS_NAME, "bi-x-circle"))
            close_add.click()
            
     #   login.click()
        time.sleep(5)

    def find_element_safe(self, by, selector):
        try:
            element = self.driver_connection.find_element((by, selector))
            return element
        except:
            return np.nan

    
    def handle_split_error(self,element,default_value = np.nan):
        try:
            return element
        except:
            pass
        return default_value        
        
    
    def read_credentials_from_text(self, text_file):
        with open(text_file, 'r') as file:
            lines = file.readlines()
            user_id_markets_mojo = lines[0].strip()
            password_markets_mojo = lines[1].strip()
        return user_id_markets_mojo, password_markets_mojo

    def analyze_stock(self, stock):

        search_button = self.driver_connection.wait_for_element((By.CSS_SELECTOR,".menu a"))
        search_button.click()
        time.sleep(2)
        search_input_stocks =  self.driver_connection.wait_for_element((By.CSS_SELECTOR,"#mm-header > \
                                                                                 div > form > input"))
        #search_input_stocks.send_keys(stock + Keys.ENTER)
        for char in stock:
            search_input_stocks.send_keys(char)
            time.sleep(0.5) 
        
        try:
            stock_option = self.driver_connection.wait_for_element((By.CSS_SELECTOR,"#mm-header > div > div > div > \
                                                                         div.mm-search-wrapper > div > div.filter-top >\
                                                                         span:nth-child(2)"))
        except:
            self.cmp_list.append(np.nan)
            self.mojo_score_list.append(np.nan)
            self.mojo_stock_yes_list.append(np.nan)
            self.mojo_stock_since_list.append(np.nan)
            self.dashboard_recommendation_list.append(np.nan)
            self.dashboard_comment_list.append(np.nan)
            self.recommendated_technical_list.append(np.nan)
            self.recommendated_technical_comment_list.append(np.nan)
            self.quality_score_list.append(np.nan)
            self.quality_comments_list.append(np.nan)
            self.valuation_score_list.append(np.nan)
            self.financial_trend_positive_list.append(np.nan)
            self.financial_trend_negative_list.append(np.nan)
            self.financial_trend_total_list.append(np.nan)
            self.financial_trend_status_list.append(np.nan)
            self.one_day_list.append(np.nan)
            self.one_week_list.append(np.nan)
            self.one_month_list.append(np.nan)
            self.three_month_list.append(np.nan)
            self.six_month_list.append(np.nan)
            self.ytd_list.append(np.nan)
            self.one_year_list.append(np.nan)
            self.two_year_list.append(np.nan)
            self.three_year_list.append(np.nan)
            self.four_year_list.append(np.nan)
            self.five_year_list.append(np.nan)
            self.ten_year_list.append(np.nan)
            self.majority_shareholders_list.append(np.nan)
            self.pledged_list.append(np.nan)
            self.mf_list.append(np.nan)
            self.mf_scheme_percent_list.append(np.nan)
            self.fii_list.append(np.nan)
            self.fii_percent_list.append(np.nan)
            self.promoter_list.append(np.nan)
            self.promoter_percent_list.append(np.nan)
            self.highest_public_holder_list.append(np.nan)
            self.highest_public_holder_percent_list.append(np.nan)
            self.individual_investor_list.append(np.nan)
            self.markets_mojo_url.append(np.nan) 
            self.markets_mojo_url.append(np.nan)
            self.driver_connection.refresh()
            
        else:
            stock_option = self.driver_connection.wait_for_element((By.CSS_SELECTOR,"#mm-header > div > div > div > \
                                                                         div.mm-search-wrapper > div > div.filter-top >\
                                                                         span:nth-child(2)"))
                                                                                  
            stock_option.click()
            time.sleep(8)
            fund_name = self.driver_connection.find_element((By.CLASS_NAME,"fund-name"))
            fund_name.click()
            time.sleep(2)
            market_mojo_url = self.driver_connection.get_url()
            self.markets_mojo_url.append(market_mojo_url)
            self.dashboard_analysis()
            self.technical_analysis()
            self.quality_analysis()
            self.financial_analysis()
            self.returns_analysis()
            self.shareholding_analysis()
            
           

    def handle_element_text(self, element):
        try:
            if element:
                return element.text
        except Exception:
            pass
        return np.nan


    def dashboard_analysis(self):
        time.sleep(5)        
        cmp_element = self.find_element_safe(By.XPATH,"/html/body/app-root/stock-root/div/div[1]/div[1]/div[2]/div[1]/p[2]/span[1]")
        cmp = self.handle_element_text(cmp_element)
        self.cmp_list.append(cmp)

        mojo_score_element = self.find_element_safe(By.CLASS_NAME,'precentage')
        mojo_score = self.handle_element_text(mojo_score_element)
        self.mojo_score_list.append(mojo_score)
        
        try:
            element = self.find_element_safe(By.CLASS_NAME, "ms1") 
            
        except NoSuchElementException:
            mojo_stock_yes = "No"
            mojo_stock_since = np.nan
        else:
            mojo_stock = self.handle_element_text(element)
            if mojo_stock == "MOJO STOCK":
                mojo_stock_yes = "Yes"
            else:
                mojo_stock_yes = np.nan    
            mojo_stoc_sin = self.find_element_safe(By.CLASS_NAME, "ms2") 
            mojo_stock_since = self.handle_element_text(mojo_stoc_sin)
            
        finally:
            self.mojo_stock_yes_list.append(mojo_stock_yes)
            self.mojo_stock_since_list.append(mojo_stock_since)

        dashboard_recommendation_element = self.find_element_safe(By.CSS_SELECTOR,"#newstockscore > div.scorecontent1.no-mob.no-tab2.ng-tns-c93-0 > div > span")
        dashboard_recommendation = self.handle_element_text(dashboard_recommendation_element)
        self.dashboard_recommendation_list.append(dashboard_recommendation)

        dashboard_comment_element = self.find_element_safe(By.CLASS_NAME,"scorefooter")
        dashboard_comment = self.handle_element_text(dashboard_comment_element)
        self.dashboard_comment_list.append(dashboard_comment)

    def technical_analysis(self):
        technical_button = self.driver_connection.find_element((By.CSS_SELECTOR, "#myNavbar > ul > li:nth-child(2) > a"))
        technical_button.click()
        time.sleep(2)

        technical_elements = self.find_element_safe(By.XPATH, '//*[@id="sectIndigraph_graph"]/div/div[2]/div[2]/div[2]/div[1]/div[2]/p[1]')
        recommendated_technical = self.handle_element_text(technical_elements)
        technical_comment = self.find_element_safe(By.XPATH, '//*[@id="sectIndigraph_graph"]/div/div[2]/div[2]/div[2]/div[1]/div[2]/p[2]')
        recommendated_technical_comment = self.handle_element_text(technical_comment)
        self.recommendated_technical_list.append(recommendated_technical)
        self.recommendated_technical_comment_list.append(recommendated_technical_comment)


    def quality_analysis(self):
        quality_button = self.driver_connection.find_element((By.CSS_SELECTOR, "#myNavbar > ul > li:nth-child(3) > a"))
        quality_button.click()
        time.sleep(2)

        quality_score_element = self.find_element_safe(By.CLASS_NAME, "score-status2")
        quality_score = self.handle_element_text(quality_score_element)
        self.quality_score_list.append(quality_score)

        quality_comment_elements =self.driver_connection.find_elements((By.CLASS_NAME, "qualityularea_card"))
        if quality_comment_elements:
            quality_comments = " ".join([comment.text for comment in quality_comment_elements])
        else:
            quality_comments = np.nan
        self.quality_comments_list.append(quality_comments)

    def financial_analysis(self):
        valuation_button = self.driver_connection.find_element((By.CSS_SELECTOR, "#myNavbar > ul > li:nth-child(4) > a"))
        valuation_button.click()
        time.sleep(2)

        valuation_score_element = self.find_element_safe(By.CLASS_NAME, "dsh-txt6")
        valuation_score = self.handle_element_text(valuation_score_element)
        self.valuation_score_list.append(valuation_score)

        financial_trend_button = self.find_element_safe(By.CSS_SELECTOR, "#myNavbar > ul > li:nth-child(5) > a")

        if financial_trend_button:
            financial_trend_button.click()
            time.sleep(2)
            financial_trend_positive = np.nan
            financial_trend_negative = np.nan
            financial_trend_total = np.nan
            financial_trend_status = np.nan
            try:
                financial_trend_positive_element = self.find_element_safe(By.CLASS_NAME, "w-no")
                if financial_trend_positive_element:
                    financial_trend_positive = financial_trend_positive_element.text.split(" ")[1]
                else:
                    financial_trend_positive = np.nan
            except:
                financial_trend_positive = np.nan
                financial_trend_negative = np.nan
                financial_trend_total = np.nan
                financial_trend_status = np.nan                
            else:
                financial_trend_negative_element = self.driver_connection.find_elements((By.CLASS_NAME, "w-no"))[1]
                financial_trend_negative = self.handle_element_text(financial_trend_negative_element)
                
                financial_trend_length_element = self.driver_connection.find_elements((By.CSS_SELECTOR, "#valuation-main .score-status2"))
                if len(financial_trend_length_element) == 2:
                    financial_trend_total_element = financial_trend_length_element[1]
                else:
                    financial_trend_total_element = financial_trend_length_element[2]
                financial_trend_total = financial_trend_total_element.text.split(" ")[1]
                financial_trend_status_element = self.driver_connection.find_elements((By.CSS_SELECTOR, "#valuation-main .score-status2"))[1]
                financial_trend_status = self.handle_element_text(financial_trend_status_element)

                
            finally:
                self.financial_trend_positive_list.append(financial_trend_positive)
                self.financial_trend_negative_list.append(financial_trend_negative)
                self.financial_trend_total_list.append(financial_trend_total)
                self.financial_trend_status_list.append(financial_trend_status)
                
                

    def returns_analysis(self):
        try:
            returns_button = self.driver_connection.find_element((By.CSS_SELECTOR, "#myNavbar > ul > li:nth-child(7) > a"))
            returns_button.click()
            time.sleep(2)

            one_day = self.get_element_text(".ng-star-inserted > div:nth-child(1) > div:nth-child(1) > h6")
            one_week = self.get_element_text(".ng-star-inserted > div:nth-child(1) > div:nth-child(2) > h6")
            one_month = self.get_element_text(".ng-star-inserted > div:nth-child(1) > div:nth-child(3) > h6")
            three_month = self.get_element_text(".ng-star-inserted > div:nth-child(1) > div:nth-child(4) > h6")
            six_month = self.get_element_text(".no-bor.ng-star-inserted > div:nth-child(2) > div:nth-child(1) > h6")
            ytd = self.get_element_text(".no-bor.ng-star-inserted > div:nth-child(2) > div:nth-child(2) > h6")
            one_year = self.get_element_text(".no-bor.ng-star-inserted > div:nth-child(2) > div:nth-child(3) > h6")
            two_year = self.get_element_text(".no-bor.ng-star-inserted > div:nth-child(2) > div:nth-child(4) > h6")
            three_year = self.get_element_text(".no-bor.ng-star-inserted > div:nth-child(3) > div:nth-child(1) > h6")
            four_year = self.get_element_text(".no-bor.ng-star-inserted > div:nth-child(3) > div:nth-child(2) > h6")
            five_year = self.get_element_text(".no-bor.ng-star-inserted > div:nth-child(3) > div:nth-child(3) > h6")
            ten_year = self.get_element_text(".no-bor.ng-star-inserted > div:nth-child(3) > div:nth-child(4) > h6")

            self.one_day_list.append(one_day)
            self.one_week_list.append(one_week)
            self.one_month_list.append(one_month)
            self.three_month_list.append(three_month)
            self.six_month_list.append(six_month)
            self.ytd_list.append(ytd)
            self.one_year_list.append(one_year)
            self.two_year_list.append(two_year)
            self.three_year_list.append(three_year)
            self.four_year_list.append(four_year)
            self.five_year_list.append(five_year)
            self.ten_year_list.append(ten_year)
        except Exception:
            self.one_day_list.append(np.nan)
            self.one_week_list.append(np.nan)
            self.one_month_list.append(np.nan)
            self.three_month_list.append(np.nan)
            self.six_month_list.append(np.nan)
            self.ytd_list.append(np.nan)
            self.one_year_list.append(np.nan)
            self.two_year_list.append(np.nan)
            self.three_year_list.append(np.nan)
            self.four_year_list.append(np.nan)
            self.five_year_list.append(np.nan)
            self.ten_year_list.append(np.nan)

    def get_element_text(self, selector):
        try:
            element = self.driver_connection.find_element((By.CSS_SELECTOR, selector))
            return element.text.split("\n")[1].replace("%", "")
        except Exception:
            return np.nan


                
    def shareholding_analysis(self):
        shareholding = self.driver_connection.find_element((By.CSS_SELECTOR, "#myNavbar > ul > li:nth-child(10) > a"))
        shareholding.click()
        time.sleep(2)

        try:
            majority = self.driver_connection.find_element((By.CSS_SELECTOR,
                                                            "div.sec2.ng-star-inserted > ul > li:nth-child(1)")).text.split(": ")[1]
        except Exception:
            majority = np.nan
        self.majority_shareholders_list.append(majority)

        try:
            pledged = self.driver_connection.find_element((By.CSS_SELECTOR,
                                                           "div.sec2.ng-star-inserted > ul > li:nth-child(2)")).text.split(": ")[1]
        except Exception:
            pledged = np.nan
        self.pledged_list.append(pledged)

        try:
            mf_element = self.driver_connection.find_element((By.CSS_SELECTOR, "div.sec2.ng-star-inserted > ul > li:nth-child(3)"))
            mf = mf_element.text.split(" Schemes ")[0].split(" ")[-1]
            mf_scheme_percent = mf_element.text.split(" Schemes ")[1].replace("(", "").replace(")", "")
        except Exception:
            mf = np.nan
            mf_scheme_percent = np.nan
        self.mf_list.append(mf)
        self.mf_scheme_percent_list.append(mf_scheme_percent)

        try:
            fii_element = self.driver_connection.find_element((By.CSS_SELECTOR, "div.sec2.ng-star-inserted > ul > li:nth-child(4)"))
            fii = fii_element.text.split(" FIIs ")[0].split(" ")[-1]
            fii_percent = fii_element.text.split(" FIIs ")[1].replace("(", "").replace(")", "")
        except Exception:
            fii = np.nan
            fii_percent = np.nan
        self.fii_list.append(fii)
        self.fii_percent_list.append(fii_percent)

        try:
            promoter_element = self.driver_connection.find_element((By.CSS_SELECTOR,
                                                                    "div.sec2.ng-star-inserted > ul > li:nth-child(5)"))
            promoter = promoter_element.text.split(" :")[1].split(" (")[0]
            promoter_percent = promoter_element.text.split(" :")[1].split(" (")[1].replace(")", "")
        except Exception:
            promoter = np.nan
            promoter_percent = np.nan
        self.promoter_list.append(promoter)
        self.promoter_percent_list.append(promoter_percent)

        try:
            highest_public_holder_element = self.driver_connection.find_element((By.CSS_SELECTOR,
                                                                                 "div.sec2.ng-star-inserted > ul > li:nth-child(6)"))
            highest_public_holder = highest_public_holder_element.text.split(": ")[1].split(" (")[0]
            highest_public_holder_percent = highest_public_holder_element.text.split(": ")[1].split("(")[1].replace(")", "")
        except Exception:
            highest_public_holder = np.nan
            highest_public_holder_percent = np.nan
        self.highest_public_holder_list.append(highest_public_holder)
        self.highest_public_holder_percent_list.append(highest_public_holder_percent)

        try:
            individual_investor = self.driver_connection.find_element((By.CSS_SELECTOR,
                                                                       "div.sec2.ng-star-inserted > ul > li:nth-child(7)")).text.split(": ")[1]
        except Exception:
            individual_investor = np.nan
        self.individual_investor_list.append(individual_investor)

    


    def create_dataframe(self):
        data_dict = {
            "CMP": self.cmp_list,
            "Mojo Score": self.mojo_score_list,
            "Mojo Stock": self.mojo_stock_yes_list,
            "Mojo Stock Since": self.mojo_stock_since_list,
            "Recommendation": self.dashboard_recommendation_list,
            "Comment": self.dashboard_comment_list,
            "Technical Recommendated": self.recommendated_technical_list,
            "Technical Comment": self.recommendated_technical_comment_list,
            "Quality Score": self.quality_score_list,
            "Quality Comments": self.quality_comments_list,
            "Valuation Score": self.valuation_score_list,
            "FT Positive": self.financial_trend_positive_list,
            "FT Negative": self.financial_trend_negative_list,
            "FT Total": self.financial_trend_total_list,
            "Financial Trend": self.financial_trend_status_list,
            "1D": self.one_day_list,
            "1W": self.one_week_list,
            "1M": self.one_month_list,
            "3m": self.three_month_list,
            "6M": self.six_month_list,
            "YTD": self.ytd_list,
            "1Y": self.one_year_list,
            "2Y": self.two_year_list,
            "3Y": self.three_year_list,
            "4Y": self.four_year_list,
            "5Y": self.five_year_list,
            "10Y": self.ten_year_list,
            "Majority Shareholders": self.majority_shareholders_list,
            "Pledged": self.pledged_list,
            "MF": self.mf_list,
            "MF Scheme%": self.mf_scheme_percent_list,
            "FII": self.fii_list,
            "FII%": self.fii_percent_list,
            "Promoter": self.promoter_list,
            "Promoter%": self.promoter_percent_list,
            "Highest Pub-Holder": self.highest_public_holder_list,
            "Highest Pub-Holder%": self.highest_public_holder_percent_list,
            "Ind Investor": self.individual_investor_list,
        }

        df = pd.DataFrame(data_dict)
        return df        

    
    def market_mojo_data_into_excel(self, stock_list_sheet, stock_sheet_list_cell,mojo_sheet,mojo_sheet_cell,market_mojo_url_column="B6",
                                  excel_file_path=None):
        if excel_file_path:
            xt = xw.Book(excel_file_path)
        else:
            xt = xw.Book.caller()
        stock_sheet_list_name = xt.sheets[stock_list_sheet]
        mojo_analsis_sheeet = xt.sheets[mojo_sheet]
        stock_name_list = stock_sheet_list_name.range(stock_sheet_list_cell).options(expand="down").value 
        filtered_stock_name_list = [name for name in stock_name_list if name not in ('', ',')]
        filtered_stock_name_list = sorted(list(OrderedDict.fromkeys(filtered_stock_name_list)))

        print(stock_name_list)
        for stock in filtered_stock_name_list:
            self.analyze_stock(stock)
        
        time.sleep(1)        
        data_frame = self.create_dataframe()
        data_frame.set_index(data_frame.columns[0], inplace=True)
        mojo_analsis_sheeet.range(mojo_sheet_cell).options(expand="down").value = data_frame
        mojo_analsis_sheeet.range(market_mojo_url_column).options(transpose=True).value = self.markets_mojo_url 
    
    
    def logout(self):
        try:
            logout_button = self.driver_connection.find_element((By.CSS_SELECTOR,"li.user-login.user-login-web > \
                                                                  div > button > img"))
            logout_button.click()
            logout_button_again = self.driver_connection.find_element((By.XPATH,'//*[@id="mm-header"]\
                                                                        /div/nav[2]/ul/li[3]/div/div/a[11]'))
            logout_button_again.click()
            time.sleep(2)
            self.driver_connection.close()
            print("Logged out successfully.")
        except Exception as e:
            print("Error during logout:", str(e))
            
            
            
    def close_connection(self):
        self.driver_connection.close()