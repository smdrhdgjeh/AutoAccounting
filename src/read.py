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

        i = starting_row
        while property_inform.loc[i, list(property_inform.columns)[0]] != "투자성 자산":
            if property_inform.loc[i, list(property_inform.columns)[1]] == '신한 주거래 S20통장':
                self.Sinhan_bank = property_inform.loc[i, list(property_inform.columns)[3]]
            elif property_inform.loc[i, list(property_inform.columns)[1]] == '저축예금':
                self.Hana_bank = property_inform.loc[i, list(property_inform.columns)[3]]
            elif property_inform.loc[i, list(property_inform.columns)[1]] == '직장인우대통장-저축예금':
                self.KB_bank = property_inform.loc[i, list(property_inform.columns)[3]]
            elif str(property_inform.loc[i, list(property_inform.columns)[1]]).find("정기예금") != -1:
                self.Deposit += property_inform.loc[i, list(property_inform.columns)[3]]
            elif str(property_inform.loc[i, list(property_inform.columns)[1]]).find("적금") != -1:
                self.Installment_savings += property_inform.loc[i, list(property_inform.columns)[3]]
            i += 1


