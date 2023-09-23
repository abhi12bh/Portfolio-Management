import xlwings as xw
import pandas as pd
import numpy as np

class StockFilter:
    def __init__(self, fund_sheet_name, filter_sheet_name, usecols = "A:J",excel_file_path=None):
        self.excel_file_path = excel_file_path
        self.xt = xw.Book.caller()
        self.filter_sheet_name = filter_sheet_name
        self.fund_sheet_name = fund_sheet_name
        self.filter_sheet = self.xt.sheets[self.filter_sheet_name]
        self.fund_analysis = self.xt.sheets[self.fund_sheet_name]
        self.usecols = usecols

    def load_data(self, numeric_column_list = ['Durability', 'Valuation', 'Momentum', 'Forecast Price'], no_rows =11):
        df = pd.read_excel(self.excel_file_path, sheet_name=self.filter_sheet_name, usecols=self.usecols,skiprows=1)
        market_data = pd.read_excel(self.excel_file_path, sheet_name=self.fund_sheet_name, skiprows=4)

        market_data.replace("-", np.nan, inplace=True)
        numeric_columns = numeric_column_list
        market_data[numeric_columns] = market_data[numeric_columns].astype(float)
        df = df.iloc[:no_rows]
        columns_list = df.columns
        self.columns_list = columns_list
        self.df = df
        self.market_data = market_data

    def custom_isin(self, condition_list, value):
        results = []
        for condition in condition_list:
            if isinstance(condition, str) and condition.startswith(('>', '<', '>=', '<=')):
                operator = condition[0]
                number_str = condition[1:]
                try:
                    number = float(number_str)
                except ValueError:
                    results.append(False)  # Invalid number format
                else:
                    if operator == '>':
                        result = value > number
                    elif operator == '<':
                        result = value < number
                    elif operator == '>=':
                        result = value >= number
                    elif operator == '<=':
                        result = value <= number
                    else:
                        result = False  # Invalid operator
                    results.append(result)
            else:
                results.append(value == condition)
        return any(results)

    def filter_market_data(self):
        filtered_rows = []
        for index, row in self.market_data.iterrows():
            conditions = []
            for column in self.columns_list:
                conditions.append(self.custom_isin(list(self.df[column].dropna()), row[column]))
            if all(conditions):
                filtered_rows.append(row)

        self.filtered_market_data = pd.DataFrame(filtered_rows)
        
    def data_to_excel(self,filter_sheet_cell = "A12" ):
        if not self.filtered_market_data.empty:
            filtered_market_data_copy = self.filtered_market_data[['NSE Symbol','Recommendation', 'Technical Recommendated', 'Quality Score','Valuation Score', 'Durability', 'Valuation', 'Momentum','Financial Trend', 'Forecast Consensus', 'CAP','Comment']]
            filtered_market_data_copy.set_index(filtered_market_data_copy.columns[0], inplace=True)
            self.filter_sheet.range(filter_sheet_cell).options(expand="down").value = filtered_market_data_copy

