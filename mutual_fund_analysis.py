from mutual_fund_downloader import MutualFundDownloader
from mutualfundanalyzer import MutualFundAnalyzer
from trendlyne_data import TrendlyneData
from marketsmojo import MarketsMojoData
from data_grabber import DataGrabber
from stock_filter import StockFilter


import os
import time


current_dir = os.getcwd()

WEB_DRIVER_PATH = os.path.join(current_dir, "webdriver", "chromedriver.exe")

CURRENT_PATH = 'C:\\Users\\Abhishek\\Downloads\\Freelancing\\Stock Market Projects\\Mutual Funds\\'
DOWNLOAD_DIRECTORY = f"{CURRENT_PATH}mutual_funds_file_download\\"
TARGET_FOLDER  = f"{CURRENT_PATH}Mutual Funds\\database\\"
TRENDLNE_CREDENTIALS = f"{CURRENT_PATH}trendlyne_credentials.txt"
MARKETS_MOJO_CREDENTIALS = f"{CURRENT_PATH}marketsmojo_credentials.txt"
ALL_DATA_FILE = f"{CURRENT_PATH}database\\all_data.xlsx"
NSE_LIST_PATH =  f"{CURRENT_PATH}MARKET_LIST\\nse_list.xlsx"
EXCEL_FILE_PATH = f"{CURRENT_PATH}mutual_fund_analysis.xlsm"


# current_dir = os.getcwd()

# WEB_DRIVER_PATH = os.path.join(current_dir, "webdriver", "chromedriver.exe")
# DOWNLOAD_DIRECTORY = os.path.join(current_dir, "mutual_funds_file_download")
# TARGET_FOLDER = os.path.join(current_dir, "database")
# TRENDLNE_CREDENTIALS = os.path.join(current_dir,  "trendlyne_credentials.txt")
# MARKETS_MOJO_CREDENTIALS = os.path.join(current_dir,  "marketsmojo_credentials.txt")
# ALL_DATA_FILE = os.path.join(current_dir,"database", "all_data.xlsx")
# NSE_LIST_PATH = os.path.join(current_dir, "MARKET_LIST", "nse_list.xlsx")
# EXCEL_FILE_PATH = os.path.join(current_dir, "mutual_fund_analysis.xlsm")


rupeevest_url = "https://www.rupeevest.com/Mutual-Fund-Portfolio-Tracker"
trendlyne_url = "https://trendlyne.com/features/"
markets_mojo_url = "https://www.marketsmojo.com/mojofeed/login"
temporary_mojo_url = "https://www.marketsmojo.com/mojo/login?redirect=mojo/stock-research?"
#temporary_mojo_url = "https://www.marketsmojo.com/mojo/profileupdate"


mutual_fund_list_sheet = "mutual_fund_list"
mutual_fund_list_cell = "A4"
rupeevest_fetch_data_sheet = "rupeevest_fetch_data"
rupeevest_fetch_data_cell = "A4"

stock_list_sheet = "stock_list"
stock_list_name_cell = "A2"
portfolio_analysis_sheet="portfolio_analysis"
portfolio_analysis_trendline_cell = "AO5"
portfolio_analysis_mojo_cell = "C5"



mutual_fund_analysis_sheet = "mutual_fund_analysis"
mutual_fund_analysis_trendline_cell = "AO5"
mutual_fund_analysis_mojo_cell = "C5"
added_new_sheet_name = "added_new_stocks"
add_new_sheet_cell_data = "D7"
added_new_sheet_cell_nse = "F6"
stock_column = "A6"
trendlyne_url_column= "BL6"
market_mojo_url_column = "BK6"



large_cap_mf_sheet_name = "large_cap_filter_mf"
mid_cap_mf_sheet_name = "mid_cap_filter_mf"
small_cap_mf_sheet_name = "small_cap_filter_mf"
large_cap_portfolio_sheet_name = "large_cap_filter_portfolio"
mid_cap_portfolio_sheet_name = "mid_cap_filter_portfolio"
small_cap_portfolio_sheet_name = "small_cap_filter_portfolio"



def download_mutual_fund():
    downloader = MutualFundDownloader(WEB_DRIVER_PATH, DOWNLOAD_DIRECTORY)
    downloader.download_mutual_fund_files(rupeevest_url, mutual_fund_list_sheet, mutual_fund_list_cell)
    downloader.close()
        

def mutual_fund_data_analyze():
    analyzer = MutualFundAnalyzer(DOWNLOAD_DIRECTORY)
    analyzer.analyze_mutual_fund_data(rupeevest_fetch_data_sheet,rupeevest_fetch_data_cell)
    analyzer.move_files_to_database(TARGET_FOLDER,ALL_DATA_FILE)
    
def fetch_trend_line_portfolio_data():
    trendlyne_data = TrendlyneData(WEB_DRIVER_PATH, DOWNLOAD_DIRECTORY, headless=False)
    user_id_trendlyne, password_trendlyne = trendlyne_data.read_credentials_from_text(TRENDLNE_CREDENTIALS)
    trendlyne_data.trendlyne_login(trendlyne_url, user_id_trendlyne, password_trendlyne)
    trendlyne_data.trendlyne_data_into_excel(stock_list_sheet,stock_list_name_cell,portfolio_analysis_sheet,
                                            portfolio_analysis_trendline_cell,stock_column,trendlyne_url_column)
    time.sleep(1)    
    trendlyne_data.logout()
    
def fetch_market_mojo_portfolio_data():
    markets_mojo_data = MarketsMojoData(WEB_DRIVER_PATH, DOWNLOAD_DIRECTORY, headless=False)
    user_id_markets_mojo, password_markets_mojo = markets_mojo_data.read_credentials_from_text(MARKETS_MOJO_CREDENTIALS)
    markets_mojo_data.markets_mojo_login(temporary_mojo_url, user_id_markets_mojo, password_markets_mojo, temporary_url=True)
    markets_mojo_data.market_mojo_data_into_excel(stock_list_sheet,stock_list_name_cell,portfolio_analysis_sheet,
                                           portfolio_analysis_mojo_cell,market_mojo_url_column)
    time.sleep(1)    
    #markets_mojo_data.logout()

def fetch_trend_line_mf_data():
    trendlyne_data = TrendlyneData(WEB_DRIVER_PATH, DOWNLOAD_DIRECTORY, headless=False)
    user_id_trendlyne, password_trendlyne = trendlyne_data.read_credentials_from_text(TRENDLNE_CREDENTIALS)
    trendlyne_data.trendlyne_login(trendlyne_url, user_id_trendlyne, password_trendlyne)
    trendlyne_data.trendlyne_data_into_excel(added_new_sheet_name,add_new_sheet_cell_data,mutual_fund_analysis_sheet,
                                            mutual_fund_analysis_trendline_cell,stock_column,trendlyne_url_column)
    time.sleep(1)    
    trendlyne_data.logout()
    
def fetch_market_mojo_mf_data():
    markets_mojo_data = MarketsMojoData(WEB_DRIVER_PATH, DOWNLOAD_DIRECTORY, headless=False)
    user_id_markets_mojo, password_markets_mojo = markets_mojo_data.read_credentials_from_text(MARKETS_MOJO_CREDENTIALS)
    markets_mojo_data.markets_mojo_login(temporary_mojo_url, user_id_markets_mojo, password_markets_mojo, temporary_url=True)
    markets_mojo_data.market_mojo_data_into_excel(added_new_sheet_name,add_new_sheet_cell_data,mutual_fund_analysis_sheet,
                                            mutual_fund_analysis_mojo_cell,market_mojo_url_column)
    time.sleep(1)    
    #markets_mojo_data.logout()    
    
    
def get_data():
    data_grabber = DataGrabber(added_new_sheet_name)
    data_grabber.read_nse_sheet(NSE_LIST_PATH,added_new_sheet_cell_nse)

    
def large_cap_filter_mf():
    use_cols = "A:J"
    stock_filter = StockFilter(mutual_fund_analysis_sheet,large_cap_mf_sheet_name,use_cols,EXCEL_FILE_PATH)

    numeric_column_list = ['Durability', 'Valuation', 'Momentum', 'Forecast Price'] 
    no_rows = 12
    stock_filter.load_data(numeric_column_list,no_rows)
    stock_filter.filter_market_data()
    stock_filter.data_to_excel(filter_sheet_cell = "A12" )
    
    
def mid_cap_filter_mf():
    use_cols = "A:J"
    stock_filter = StockFilter(mutual_fund_analysis_sheet,mid_cap_mf_sheet_name,use_cols,EXCEL_FILE_PATH)

    numeric_column_list = ['Durability', 'Valuation', 'Momentum', 'Forecast Price'] 
    no_rows = 12
    stock_filter.load_data(numeric_column_list,no_rows)
    stock_filter.filter_market_data()
    stock_filter.data_to_excel(filter_sheet_cell = "A12" )
    
    
def small_cap_filter_mf():
    use_cols = "A:J"
    stock_filter = StockFilter(mutual_fund_analysis_sheet,small_cap_mf_sheet_name,use_cols,EXCEL_FILE_PATH)

    numeric_column_list = ['Durability', 'Valuation', 'Momentum', 'Forecast Price'] 
    no_rows = 12
    stock_filter.load_data(numeric_column_list,no_rows)
    stock_filter.filter_market_data()
    stock_filter.data_to_excel(filter_sheet_cell = "A12" )
    
    
def large_cap_filter_portfolio():
    use_cols = "A:J"
    stock_filter = StockFilter(portfolio_analysis_sheet,large_cap_portfolio_sheet_name,use_cols,EXCEL_FILE_PATH)

    numeric_column_list = ['Durability', 'Valuation', 'Momentum', 'Forecast Price'] 
    no_rows = 12
    stock_filter.load_data(numeric_column_list,no_rows)
    stock_filter.filter_market_data()
    stock_filter.data_to_excel(filter_sheet_cell = "A12" )
    
    
def mid_cap_filter_portfolio():
    use_cols = "A:J"
    stock_filter = StockFilter(portfolio_analysis_sheet,mid_cap_portfolio_sheet_name,use_cols,EXCEL_FILE_PATH)

    numeric_column_list = ['Durability', 'Valuation', 'Momentum', 'Forecast Price'] 
    no_rows = 12
    stock_filter.load_data(numeric_column_list,no_rows)
    stock_filter.filter_market_data()
    stock_filter.data_to_excel(filter_sheet_cell = "A12" )
    
    
def small_cap_filter_portfolio():
    use_cols = "A:J"
    stock_filter = StockFilter(portfolio_analysis_sheet,small_cap_portfolio_sheet_name,use_cols,EXCEL_FILE_PATH)

    numeric_column_list = ['Durability', 'Valuation', 'Momentum', 'Forecast Price'] 
    no_rows =12
    stock_filter.load_data(numeric_column_list,no_rows)
    stock_filter.filter_market_data()
    stock_filter.data_to_excel(filter_sheet_cell = "A12" )
    
