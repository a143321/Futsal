import pandas as pd

class ExcelAccess():

    def __init__(self, file_path, header_row = None):
        """
        引数:
            file_path : string : 開くExcelファイルのフルパスを入力
            header_row : int : ヘッダー行を指定(0行) (ヘッダー行がない場合はNone)
        """
        # Excelファイル名
        self.file_path = file_path

        # ヘッダー行の有無
        self.headerExits = (header_row != None)
        
        # ヘッダー行 (ヘッダー行が存在しない場合は-1)
        self.headerLineCount = header_row if self.headerExits else -1

        self.df = self.ReadExcel()
    
    def ReadExcel(self):

        if not self.headerExits:
            return pd.read_excel(self.file_path, dtype=str, encoding='utf8', header=None)
        
        return pd.read_excel(self.file_path, dtype=str, encoding='utf8', header=self.headerLineCount)
        

    def ReadCell(self, row_index, column_index):
        """セルの値を取得する
        引数 : 
            row_index : int : データ行 (最初のデータであれば0)
            column_index : int : データ列(最初のデータであれば0)
        戻り値 : str
        """
        return self.df.iloc[row_index, column_index]


    def ReadRow(self, row_index):
        """行のデータを取得する
        引数 : 
            column_index : int : データ列(最初のデータであれば0)
        戻り値 : str型の一次元配列 : <class 'numpy.ndarray'>
        """
        return self.df.iloc[row_index, 0:]


    def ReadColumn(self, column_index):
        """列のデータを取得する
        引数 : 
            row_index : int : データ行 (最初のデータであれば0)
        戻り値 : 一次元配列 : <class 'numpy.ndarray'>
        """
        return self.df.iloc[0:, column_index]


    def ReadArea(self, start_cell, end_cell):
        """範囲を取得する
        引数 : 
            start_cell : tuple(int,int) : 開始地点のセルのタプル(行番号 : 列番号)
            end_cell : tuple(int,int) :  終了地点のセルのタプル(行番号 : 列番号)
        戻り値 : 二次元配列 : <class 'numpy.ndarray'>
        """
        
        start_row, start_column = start_cell
        end_row, end_column = end_cell
        
        return self.df.iloc[start_row:end_row, start_column:end_column].values

    def ReadAll(self):
        """全範囲を取得する
        引数 : 
        戻り値 : 二次元配列 : <class 'numpy.ndarray'>
        """
        return self.df.values

    def WriteCell(self, row_index, column_index, value):
        """セルに値を設定する
        引数 : 
            row_index : int : データ行 (最初のデータであれば0)
            column_index : int : データ列(最初のデータであれば0)
            param : str : 書き込む値
        戻り値 : str
        """
        self.df.iloc[row_index, column_index] = value


    def SaveExcel(self):
        # ヘッダーがあれば、ヘッダーを付与。ヘッダー行数が2行目の場合は、1行目を空白行とする
        self.df.to_excel(self.file_path, index=False, header=self.headerExits, startrow = self.headerLineCount)

