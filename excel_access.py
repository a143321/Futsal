import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import matplotlib.dates as mdates

class ExcelAccess():

    def __init__(self, file_path):
        self.df = pd.read_excel(file_path, dtype=str, encoding='utf8')


        print(type(self.df))

        for index, row in self.df.iterrows():
            print(f"{index}")
            print(f"[番号 : {row[0]}] Name : {row[1]}] [ID : {row[2]}] [Password : {row[3]}]")

    def get_excel_data(self):
        print(self.df)

        print(self.df.values)

        print(type(self.df.values))

        return self.df.values