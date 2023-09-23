import os
import pandas as pd
import shutil
import xlwings as xw
from openpyxl import Workbook


class MutualFundAnalyzer:
    """
    A class for analyzing mutual fund data from CSV files.
    """

    def __init__(self, folder_path):
        """
        Initialize the MutualFundAnalyzer instance.

        Parameters:
        - folder_path (str): Path to the folder containing CSV files.
        """
        self.folder_path = folder_path
        self.results_final = pd.DataFrame()

    def process_files(self):
        """
        Process all CSV files in the specified folder.
        """
        file_list = os.listdir(self.folder_path)

        for file_name in file_list:
            if file_name.endswith('.csv'):
                self.process_csv_file(file_name)

    def process_csv_file(self, file_name):
        """
        Process a CSV file.

        Parameters:
        - file_name (str): Name of the CSV file.
        """
        mutual_fund_name = file_name.split(".csv")[0]
        file_path = os.path.join(self.folder_path, file_name)

        df = pd.read_csv(file_path, skiprows=2)
        df.set_index("Company", inplace=True)
        df.rename(columns={'Unnamed: 2': 'No. of Shares 2', 'Unnamed: 4': 'No. of Shares 4',
                           'Unnamed: 6': 'No. of Shares 6', 'Unnamed: 8': 'No. of Shares 8'}, inplace=True)
        df = df.iloc[1:]
        df.dropna(inplace=True)

        new_df = df.iloc[:, :4]
        new_df = new_df.replace('-', 0)
        #new_df = new_df.apply(pd.to_numeric, errors='coerce')
        column_list = list(new_df.columns)
        columns_to_convert = column_list
        new_df[columns_to_convert] = new_df[columns_to_convert].apply(pd.to_numeric, errors='coerce')
        new_df['Change'] = new_df['No. of Shares 2'] - new_df['No. of Shares 4']
        new_df.sort_values(column_list[0], ascending=False,inplace=True)
        new_df['Preferred Rank'] = new_df[column_list[0]].rank(ascending=False, method='first').astype(int)

        status = []
        for _, row in new_df.iterrows():
            if pd.isnull(row[1]):
                status.append("EXIT")
            elif row[1]==0 and row[3]==0:
                status.append("REMOVED")
            elif row[4]>0 and row[3]==0:
                status.append("NEW")
            elif row[4]>0:
                status.append('ADDED')
            elif row[4]<0:
                status.append('PARTIAl EXIT')
            elif row[4]==0:
                status.append('NOCHANGE')
        month = new_df.columns[0].split(" AUM")[0]
        new_df.insert(0, 'Fund', mutual_fund_name)
        column_names = new_df.columns
        new_column_names = {column_names[1]: column_names[1].split(" ")[0],
                            column_names[3]: column_names[3].split(" ")[0]}
        new_df.rename(columns=new_column_names, inplace=True)
        new_df['Status'] = status
        self.results_final = pd.concat([self.results_final, new_df], axis=0)
        
        new_file_name = f"{mutual_fund_name}_{month}.csv"
        old_path = os.path.join(self.folder_path, file_name)
        new_path = os.path.join(self.folder_path, new_file_name)
        os.rename(old_path, new_path)
 
        
        
    def move_files_to_database(self, target_folder,excel_file):
        """
        Move processed CSV files to the target database folder.

        Parameters:
        - target_folder (str): Path to the target folder.
        """
        writer = pd.ExcelWriter(excel_file, engine='openpyxl')
        sheet_name = self.results_final.columns[1]
        self.results_final.to_excel(writer, sheet_name=sheet_name)

        writer.close()
        for file_name in os.listdir(self.folder_path):
            if file_name.endswith('.csv'):
                source_path = os.path.join(self.folder_path, file_name)
                target_path = os.path.join(target_folder, file_name)
                shutil.move(source_path, target_path)
                
                

    def analyze_mutual_fund_data(self, sheet_name, start_cell, excel_file_path=None,):
        if excel_file_path:
            xt = xw.Book(excel_file_path)
        else:
            xt = xw.Book.caller()
     
        mutual_fund_data_sheet = xt.sheets[sheet_name]
        self.process_files()
        mutual_fund_data_sheet.range(start_cell).options(expand="down").value = self.results_final
 