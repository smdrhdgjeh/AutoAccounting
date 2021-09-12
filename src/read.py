import os
import pandas as pd

from datetime import datetime, timedelta, date

class My_Read():
    def __init__(self):
        pass
    
    def check_excel_data_date_and_classifying(self, file_location):
        for (path, dir, files) in os.walk(file_location):
            print("path:", path, "dir:", dir, "files:",files)

    def excel_date_type_convert_str(self, excel_date=None): # 엑셀에서 날짜 읽었을때 숫자 형태일시 문자형태로 변화
        if type(excel_date) is not str:
            excel_date = date(1900, 1, 1) + timedelta(int(excel_date - 2))
            excel_date = excel_date.strftime("%Y-%m-%d")

        return excel_date

    def read_excel_file_deposit_inform(self, file=None):
        df = pd.read_excel(file, sheet_name=0)

        # setting the property data information row
        starting_row = 38 #starting property information row

        property_inform = df.loc[starting_row:, list(df.columns)[1:5]]

        self.KB_bank = 0
        self.Hana_bank = 0
        self.Sinhan_bank = 0
        self.Deposit = 0
        self.Installment_savings = 0
        self.House_invest_deposit = 0

        ##########################################
        ############# Need to setting ############
        ##########################################
        i = starting_row
        while property_inform.loc[i, list(property_inform.columns)[0]] != "투자성 자산":
            if property_inform.loc[i, list(property_inform.columns)[1]] == '신한 주거래 S20통장':
                self.Sinhan_bank = property_inform.loc[i, list(property_inform.columns)[3]]
            elif property_inform.loc[i, list(property_inform.columns)[1]] == '저축예금':
                self.Hana_bank = property_inform.loc[i, list(property_inform.columns)[3]]
            elif property_inform.loc[i, list(property_inform.columns)[1]] == '직장인우대통장-저축예금':
                self.KB_bank = property_inform.loc[i, list(property_inform.columns)[3]]
            elif property_inform.loc[i, list(property_inform.columns)[1]] == '주택청약종합저축':
                self.House_invest_deposit = property_inform.loc[i, list(property_inform.columns)[3]]
            elif str(property_inform.loc[i, list(property_inform.columns)[1]]).find("정기예금") != -1:
                self.Deposit += property_inform.loc[i, list(property_inform.columns)[3]]
            elif str(property_inform.loc[i, list(property_inform.columns)[1]]).find("적금") != -1:
                self.Installment_savings += property_inform.loc[i, list(property_inform.columns)[3]]
            i += 1

        ##########################################
        ############ Need to change #############
        ##########################################
        self.pension = 4500000

    def read_excel_file_stock_inform(self, file=None):
        ##########################################
        ############# Need to setting ############
        ##########################################
        stock_data_file1 = file + "5288377410_거래내역.xlsx"
        stock_data_file2 = file + "5932834710_거래내역.xlsx"

        self.stock_invest_money = 0
        self.stock_deposit = 0
        self.stock_revenue = 0
        self.stock_total_deposit = 0
        self.stock_predict_total_deposit = 0

        df = pd.read_excel(stock_data_file1, sheet_name=0)
        stock_data1 = df.loc[0, list(df.columns)[18:]]

        df = pd.read_excel(stock_data_file2, sheet_name=0)
        stock_data2 = df.loc[0, list(df.columns)[18:]]

        self.stock_invest_money = stock_data1[4] + stock_data2[4]
        self.stock_revenue = stock_data1[1] + stock_data2[1]
        self.stock_deposit = stock_data1[5] + stock_data2[5]
        self.stock_total_deposit = stock_data1[6] + stock_data2[6]
        self.stock_predict_total_deposit = stock_data1[7] + stock_data2[7]