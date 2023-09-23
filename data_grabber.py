import pandas as pd
import xlwings as xw

class DataGrabber:
    """
    A class for fetching NSE and BSE Data using Excel and Pandas.
    """

    def __init__(self, sheet_name, excel_file_path=None):
        """
        Initialize the DataGrabber class.
        
        Parameters:
            sheet_name (str): The name of the sheet to work with.
            excel_file_path (str, optional): The path to the Excel file. If None, the caller's active workbook is used.
        """
        self.sheet_name = sheet_name
        
        if excel_file_path:
            xt = xw.Book(excel_file_path)
        else:
            xt = xw.Book.caller()
        self.added_new_stocks = xt.sheets[self.sheet_name]
            
    def read_nse_sheet(self, nse_list_path, nse_cell):
        """
        Read NSE stock data from an Excel sheet and update the target Excel sheet with the data.
        
        Parameters:
            nse_list_path (str): The path to the Excel file containing NSE stock data.
            nse_cell (str): The target cell in the target sheet where the data should be written.
        """
        df = pd.read_excel(nse_list_path, index_col=[0])
        df.dropna(inplace=True)
        columns_list = ['NSE Symbol', 'Company Name', 'Market Cap']
        df.columns = columns_list

        cap_list = []
        for _, row in df.iterrows():
            market_cap = row['Market Cap']
            if isinstance(market_cap, (int, float)):
                if market_cap > 2000000:
                    cap_list.append('Large')
                elif 500000 < market_cap <= 2000000:
                    cap_list.append('Mid')
                elif 50000 < market_cap <= 500000:
                    cap_list.append('Small')
                elif 10000 < market_cap <= 50000:
                    cap_list.append('Macro')
                else:
                    cap_list.append('No Criteria')
            else:
                cap_list.append('No Data')
        df['Cap'] = cap_list

        new_df = df.loc[:, ['NSE Symbol', 'Company Name', 'Cap']]
        new_df.set_index("NSE Symbol", inplace=True)
        self.added_new_stocks.range(nse_cell).options(expand="down").value = new_df
