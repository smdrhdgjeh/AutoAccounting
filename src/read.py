import os
import pandas as pd

from datetime import datetime, timedelta, date

class My_Read():
    def __init__(self):
        self.salary_2020 = 1753813
        self.fixed_bonus_2020 = 962686
        self.salary_2021 = 1853813
        self.fixed_bonus_2021 = 992686

        self.year_income_2020 = 0
        self.year_spend_2020 = 0
        self.year_saving_2020 = 0
        self.year_income_2021 = 0
        self.year_spend_2021 = 0
        self.year_saving_2021 = 0

        self.month_info_dict = {}
    
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

    def read_excel_file_transaction_details(self, file=None):
        df = pd.read_excel(file, sheet_name=1)

        month = [g for n, g in df.set_index('날짜').groupby(pd.Grouper(freq='M'))]

        total_income = 0 # 전체 수입
        total_spend = 0 # 전체 지출

        fixed_income = 0 # 고정 수입
        fixed_flag1 = 0 # 월급 flag
        fixed_flag2 = 0 # 고정 상여 flag
        parents_love = 0 # 부모님의 용돈

        fixed_spend = 0 # 고정 지출
        tithe_spend = 0 # 십일조 지출
        invest_spend = 0 # 투자 지출 (주택청약)
        saving_spend = 0 # 적금 지출
        pension_spend = 0 # 연금저출 지출
        cellphon_spend = 0 # 통신비 지출
        insurance_spend = 0 # 보험 지출
        total_trans_spend = 0 # 총 교통비 지출
        car_maintenance_spend = 0 # 자동차 유지비 (주유비 + 톨비)
        trans_spend = 0 # 교통비 지출 (대중교통 + 택시)

        special_income = 0 # 고정 수입 이외 수입

        temp_list = []
        for m in range(len(month)):
            # 월 단위 연산    
            for i in range(len(month[m].index)):
                temp = month[m].iloc[i, 5]
                if temp > 0:
                    total_income += temp

                    # 월 고정 수입 (월급)
                    if month[m].iloc[i, 1] == '이체' and month[m].iloc[i, 4] == '현대모비스(주)':
                        if month[m].index[0].year == 2020:
                            if temp <= self.salary_2020 * 1.1 and temp >= self.salary_2020 * 0.9 and fixed_flag1 == 0:
                                fixed_income += temp
                                fixed_flag1 = 1
                            elif temp <= self.fixed_bonus_2020 * 1.1 and temp >= self.fixed_bonus_2020 * 0.9 and fixed_flag2 == 0:
                                fixed_income += temp
                                fixed_flag2 = 1
                        elif month[m].index[0].year == 2021:
                            if temp <= self.salary_2021 * 1.1 and temp >= self.salary_2021 * 0.9 and fixed_flag1 == 0:
                                fixed_income += temp
                                fixed_flag1 = 1
                            elif temp <= self.fixed_bonus_2021 * 1.1 and temp >= self.fixed_bonus_2021 * 0.9 and fixed_flag2 == 0:
                                fixed_income += temp
                                fixed_flag2 = 1

                    elif month[m].iloc[i, 1] == '이체' and month[m].iloc[i, 4] == '도태영':
                        parents_love += temp
                        fixed_income += temp
                    
                    # 계좌 내 이동은 월 수입에서 제외 (ex: 비자금, 적금 등)
                    if month[m].iloc[i, 7] == 'KB맑은하늘적금':
                        total_income -= temp
                    elif month[m].iloc[i, 7] == '주택청약종합저축':
                        total_income -= temp
                else:
                    total_spend += temp

                    # 월 고정 지출
                    if month[m].iloc[i, 3] == '십일조':
                        tithe_spend += temp
                        fixed_spend += temp
                    elif month[m].iloc[i, 1] == '이체' and month[m].iloc[i, 4][-2:] == '회차':
                        if month[m].index[i].day < 20:
                            invest_spend += temp # 20일 이전은 주택정약
                            fixed_spend += temp
                        else:
                            saving_spend += temp
                            fixed_spend += temp
                    elif month[m].iloc[i, 1] == '이체' and month[m].iloc[i, 4] == '퇴직기일출금':
                        pension_spend += temp
                        fixed_spend += temp
                    elif month[m].iloc[i, 2] == '주거/통신' and month[m].iloc[i, 3] == '휴대폰':
                        cellphon_spend += temp
                        fixed_spend += temp
                    elif month[m].iloc[i, 2] == '보험' and month[m].iloc[i, 4][:3] == '현대해':
                        insurance_spend += temp
                        fixed_spend += temp
                    elif month[m].iloc[i, 2] == '자동차':
                        car_maintenance_spend += temp
                        fixed_spend += temp
                    elif month[m].iloc[i, 2] == '교통':
                        trans_spend += temp
                        fixed_spend += temp
            
            if month[m].index[0].year == 2020:
                self.year_income_2020 += total_income
                self.year_spend_2020 += total_spend
                self.year_saving_2020 += (invest_spend + pension_spend + saving_spend)
            elif month[m].index[0].year == 2021:
                self.year_income_2021 += total_income
                self.year_spend_2021 += total_spend
                self.year_saving_2021 += (invest_spend + pension_spend + saving_spend)

            total_trans_spend = car_maintenance_spend + trans_spend
            special_income = total_income - fixed_income

            # 월 수입 / 월 지출 / 고정 수입 / 용돈 / 고정지출 / 십일조 / 주택청약 / 적금 / 연금저출 / 통신비 / 보험료 / 차 유지비 / 교통비 / 총 교통비 / 월급 외 수입(ex: 보너스 or 출장비 등)
            temp_list.clear()
            temp_list.append(total_income)
            temp_list.append(total_spend)
            temp_list.append(fixed_income)
            temp_list.append(parents_love)
            temp_list.append(fixed_spend)
            temp_list.append(tithe_spend)
            temp_list.append(invest_spend)
            temp_list.append(saving_spend)
            temp_list.append(pension_spend)
            temp_list.append(cellphon_spend)
            temp_list.append(insurance_spend)
            temp_list.append(car_maintenance_spend)
            temp_list.append(trans_spend)
            temp_list.append(total_trans_spend)
            temp_list.append(special_income)

            date_key_temp = str(month[m].index[0].year) + '-' + str(month[m].index[0].month)
            self.month_info_dict[date_key_temp] = temp_list.copy()

            total_income = 0
            total_spend = 0
            fixed_income = 0
            fixed_flag1 = 0
            fixed_flag2 = 0
            parents_love = 0
            fixed_spend = 0
            tithe_spend = 0
            invest_spend = 0
            saving_spend = 0
            pension_spend = 0
            cellphon_spend = 0
            insurance_spend = 0
            car_maintenance_spend = 0
            trans_spend = 0
            total_trans_spend = 0
            special_income = 0
