import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import matplotlib.dates as mdates

class ExcelAccess():

    def __init__(self, file_path, header_row = None):
        """
        引数:
            file_path : string : 開くExcelファイルのフルパスを入力
            header_row : int : ヘッダー行を指定 (デフォルトはNone)
        """
        self.file_path = file_path
        self.df = self.ReadExcel(file_path, header_row)


    def ReadExcel(self, file_path, header_row = None):
        """Excelファイル(xls, xlsx)を開く

        引数:
            file_path : string : 開くExcelファイルのフルパスを入力
            header_row : int : ヘッダー行を指定 (デフォルトはNone)
        戻り値 : <class 'pandas.core.frame.DataFrame'>
        """
        df = pd.read_excel(file_path, dtype=str, encoding='utf8', header=header_row)
        return df


    def ReadCell(self, row_index, column_index):
        """行のデータを取得する
        引数 : 
            row_index : int : 行番号
            column_index : int : 列番号

        戻り値 : str
        """
        print(type(self.df.iloc[row_index - 1, column_index - 1]))
        return self.df.iloc[row_index - 1, column_index - 1]

    def ReadRow(self, row_index):
        """行のデータを取得する
        引数 : 
            column_index : int : 列番号

        戻り値 : str型の一次元配列 : <class 'numpy.ndarray'>
        """
        for item in self.df.iloc[row_index - 1, 0:]:
            print(item)

        
        print(type(self.df.iloc[row_index - 1, 0:].values))

        return self.df.iloc[row_index - 1, 0:].values


    def ReadColumn(self, column_index):
        """列のデータを取得する
        引数 : 
            column_index : int : 列番号

        戻り値 : 一次元配列 : <class 'numpy.ndarray'>
        """
        for item in self.df.iloc[0:, column_index - 1]:
            print(item)


        print(type(self.df.iloc[0:, column_index - 1].values))
        return self.df.iloc[0:, column_index - 1].values


    def ReadArea(self, start_cell, end_cell):
        """範囲を取得する
        引数 : 
            start_cell : tuple(int,int) : 開始地点のセルのタプル(行番号 : 列番号)
            end_cell : tuple(int,int) :  終了地点のセルのタプル(行番号 : 列番号)

        戻り値 : 二次元配列 : <class 'numpy.ndarray'>
        """
        
        start_row, start_column = start_cell
        end_row, end_column = end_cell

        print(type(self.df.iloc[start_row-1:end_row, start_column-1:end_column].values))
        return self.df.iloc[start_row-1:end_row, start_column-1:end_column].values

    def ReadAll(self):
        """範囲を取得する
        引数 : 

        戻り値 : 二次元配列 : <class 'numpy.ndarray'>
        """
        return self.df.values

    def WriteCell(self, row_index, column_index, value):
        """行のデータを取得する
        引数 : 
            row_index : int : 行番号
            column_index : int : 列番号
            param : str : 書き込む値

        戻り値 : str
        """
        self.df.iloc[row_index - 1, column_index - 1] = value


    def SaveExcel(self):
        self.df.to_excel(self.file_path, index=False, header=False)