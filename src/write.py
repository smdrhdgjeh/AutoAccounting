from os import read
import xlsxwriter
from datetime import datetime, timedelta, date

class My_Write():
    def __init__(self):
        pass

    def make_exel_file(self, save_location=None, name=None):
        self.workbook = xlsxwriter.Workbook(save_location + name + '.xlsx')

        # add format
        self.bold_format = self.workbook.add_format({'bold': True,
                                                     'border': True,
                                                    #  'border_color': '#FFFFFF',
                                                     'align': 'center',
                                                     'valign': 'vcenter',
                                                     'fg_color': '#000000',
                                                     'font_color': '#FFFFFF',
                                                     'font_size': 9,
                                                     'text_wrap': True})

        self.word_format = self.workbook.add_format({'border': True,
                                                     'align': 'center',
                                                     'valign': 'vcenter',
                                                     'font_size': 9,
                                                     'text_wrap': True})

        self.money_format = self.workbook.add_format({'num_format': '#,##0',
                                                      'border': True,
                                                      'align': 'right',
                                                      'valign': 'vcenter',
                                                      'font_size': 9})

        self.percent_format = self.workbook.add_format({'num_format': '0.00',
                                                        'border': True,
                                                        'align': 'right',
                                                        'valign': 'vcenter',
                                                        'font_size': 9})

        self.date_format = self.workbook.add_format({'num_format': 'yyyy-mm-dd',
                                                     'border': True,
                                                     'align': 'center',
                                                     'valign': 'vcenter',
                                                     'font_size': 9})

        self.cell_bg_format = self.workbook.add_format({'bg_color': '#FFFF66'})

        self.merge_format = self.workbook.add_format({'bold': True,
                                                      'border': True,
                                                      'align': 'center',
                                                      'valign': 'vcenter',
                                                      'fg_color' : '#000000',
                                                      'font_color': '#FFFFFF',
                                                      'font_size': '9',
                                                      'text_wrap': True})

    def write_excel_total_deposit_inform(self, read_data=None):
        self.worksheet1 = self.workbook.add_worksheet("자산내역")
        self.worksheet1.set_column('A:Z', 10)
        self.worksheet1.set_column('O:O', 20)
        self.worksheet1.set_default_row(24)

        # calculate deposit
        self.total_flow_moeny = read_data.KB_bank\
                            + read_data.Hana_bank\
                            + read_data.Sinhan_bank\
                            + read_data.Deposit\
                            + read_data.Installment_savings
        
        self.total_invest_money = read_data.House_invest_deposit\
                                + read_data.stock_predict_total_deposit\
                                + read_data.pension

        self.total_deposit = self.total_flow_moeny + self.total_invest_money

        # take this month transaction details
        for keys in read_data.month_info_dict:
            last_month = read_data.month_info_dict[keys]
    
        # 월 수입 / 월 지출 / 고정 수입 / 용돈 / 고정지출 / 십일조 / 주택청약 / 적금 / 연금 저축 / 통신비 / 보험료 / 차 유지비 / 교통비 / 총 교통비 / 월급 외 수입(ex: 보너스 or 출장비 등)
        
        # 년 고정 지출 (office 365 + naver membership)
        year_subs_spend = 0
        year_subs_spend += (-89000 / 12)
        year_subs_spend += (-46800 / 12)

        # set data title
        self.worksheet1.write_string(0, 0, '총 자산: ', self.bold_format)
        self.worksheet1.write_number(0, 1, self.total_deposit, self.money_format)
        self.worksheet1.write_string(0, 2, '유동 자산: ', self.bold_format)
        self.worksheet1.write_number(0, 3, self.total_flow_moeny, self.money_format)
        self.worksheet1.write_string(0, 4, '투자 자산: ', self.bold_format)
        self.worksheet1.write_number(0, 5, self.total_invest_money, self.money_format)
        self.worksheet1.write_string(0, 7, '월 고정 수입: ', self.bold_format)
        self.worksheet1.write_number(0, 8, last_month[2], self.money_format)
        self.worksheet1.write_string(0, 9, '월 고정 지출: ', self.bold_format)
        self.worksheet1.write_number(0, 10, last_month[4] + year_subs_spend, self.money_format)
        self.worksheet1.write_string(0, 11, '월 특별 수입: ', self.bold_format)
        self.worksheet1.write_number(0, 12, last_month[15], self.money_format)
        self.worksheet1.write_string(0, 14, '신용카드 혜택: ', self.bold_format)
        self.worksheet1.write_number(0, 15, 22000, self.money_format)
        self.worksheet1.write_string(0, 16, '1년 저축 예상 금액: ', self.bold_format)
        self.worksheet1.write_number(0, 17, 29400000, self.money_format)
        self.worksheet1.write_string(0, 19, '업데이트 날짜: ', self.bold_format)
        self.worksheet1.write(0, 20, datetime.strptime(read_data.last_file_date, '%Y-%m-%d'), self.date_format)

        # set detail data
        self.worksheet1.write_string(1, 2, '예금: ', self.word_format)
        self.worksheet1.write_number(1, 3, read_data.Deposit, self.money_format)
        self.worksheet1.write_string(2, 2, '적금: ', self.word_format)
        self.worksheet1.write_number(2, 3, read_data.Installment_savings, self.money_format)
        self.worksheet1.write_string(3, 2, '국민은행: ', self.word_format)
        self.worksheet1.write_number(3, 3, read_data.KB_bank, self.money_format)
        self.worksheet1.write_string(4, 2, '신한은행: ', self.word_format)
        self.worksheet1.write_number(4, 3, read_data.Sinhan_bank, self.money_format)
        self.worksheet1.write_string(5, 2, '하나은행: ', self.word_format)
        self.worksheet1.write_number(5, 3, read_data.Hana_bank, self.money_format)

        self.worksheet1.write_string(1, 4, '주식: ', self.word_format)
        self.worksheet1.write_number(1, 5, read_data.stock_predict_total_deposit, self.money_format)
        self.worksheet1.write_string(2, 4, '주택청약: ', self.word_format)
        self.worksheet1.write_number(2, 5, read_data.House_invest_deposit, self.money_format)
        self.worksheet1.write_string(3, 4, '연금저축: ', self.word_format)
        self.worksheet1.write_number(3, 5, read_data.pension, self.money_format)
        self.worksheet1.write_string(4, 4, '우리사주: ', self.word_format)
        self.worksheet1.write_number(4, 5, read_data.employee_ownership * read_data.mobis_stock_price, self.money_format)

        self.worksheet1.write_string(1, 7, '월급: ', self.word_format)
        self.worksheet1.write_number(1, 8, (last_month[2] - last_month[3]), self.money_format)
        self.worksheet1.write_string(2, 7, '용돈: ', self.word_format)
        self.worksheet1.write_number(2, 8, last_month[3], self.money_format)

        self.worksheet1.write_string(1, 9, '십일조: ', self.word_format)
        self.worksheet1.write_number(1, 10, last_month[5], self.money_format)
        self.worksheet1.write_string(2, 9, '저축 금액: ', self.word_format)
        self.worksheet1.write_number(2, 10, last_month[7], self.money_format)
        self.worksheet1.write_string(3, 9, '주택 청약: ', self.word_format)
        self.worksheet1.write_number(3, 10, last_month[6], self.money_format)
        self.worksheet1.write_string(4, 9, '연금 저축: ', self.word_format)
        self.worksheet1.write_number(4, 10, last_month[8], self.money_format)
        self.worksheet1.write_string(5, 9, '통신비: ', self.word_format)
        self.worksheet1.write_number(5, 10, last_month[9], self.money_format)
        self.worksheet1.write_string(6, 9, '차 유지비: ', self.word_format)
        self.worksheet1.write_number(6, 10, last_month[11], self.money_format)
        self.worksheet1.write_string(7, 9, '교통비: ', self.word_format)
        self.worksheet1.write_number(7, 10, last_month[12], self.money_format)
        self.worksheet1.write_string(8, 9, '보험료: ', self.word_format)
        self.worksheet1.write_number(8, 10, last_month[10], self.money_format)
        self.worksheet1.write_string(9, 9, '구독료: ', self.word_format)
        self.worksheet1.write_number(9, 10, last_month[14] + year_subs_spend, self.money_format)

        self.worksheet1.write_string(1, 14, '현대카드디지털러버: ' + '\n' + '실적 50만(유튭)', self.word_format)
        self.worksheet1.write_number(1, 15, 10000, self.money_format)
        self.worksheet1.write_string(2, 14, 'KB심플라이프: ' + '\n' + '실적 30만(통신비)', self.word_format)
        self.worksheet1.write_number(2, 15, 12000, self.money_format)
        self.worksheet1.write_string(1, 16, '국민은행 적금: ', self.word_format)
        self.worksheet1.write_number(1, 17, 2000000*12, self.money_format)
        self.worksheet1.write_string(2, 16, '주택 청약: ', self.word_format)
        self.worksheet1.write_number(2, 17, 200000*12, self.money_format)
        self.worksheet1.write_string(3, 16, '연금 저축: ', self.word_format)
        self.worksheet1.write_number(3, 17, 250000*12, self.money_format)


    def write_excel_monthly_transaction_details(self, read_data=None):
        # take data dict key list
        key_list = []
        for keys in read_data.month_info_dict:
            key_list.append(keys)
        key_list.reverse()

        # 월 수입 / 월 지출 / 고정 수입 / 용돈 / 고정지출 / 십일조 / 주택청약 / 적금 / 연금 저축 / 통신비 / 보험료 / 차 유지비 / 교통비 / 총 교통비 / 월급 외 수입(ex: 보너스 or 출장비 등)

        start_row = 15

        merge_range1 = 'A'+str(start_row) + ':' + 'B'+str(start_row)
        merge_range2 = 'I'+str(start_row) + ':' + 'J'+str(start_row)
        self.worksheet1.merge_range(merge_range1, '21년', self.merge_format)
        self.worksheet1.write_string(start_row - 1, 2, '수입', self.bold_format)
        self.worksheet1.write_string(start_row - 1, 4, '지출', self.bold_format)
        self.worksheet1.write_string(start_row - 1, 6, '계', self.bold_format)
        self.worksheet1.merge_range(merge_range2, '비고', self.merge_format)

        if read_data.file_exist_true == True:
            # 새롭게 업데이트 된 데이터 입력
            for i in range(len(key_list)):
                self.worksheet1.write_string(start_row + i, 0, str(key_list[i]) + '월', self.word_format)
                self.worksheet1.write_number(start_row + i, 2, read_data.month_info_dict[key_list[i]][0], self.money_format)
                self.worksheet1.write_number(start_row + i, 4, read_data.month_info_dict[key_list[i]][1], self.money_format)
                self.worksheet1.write_number(start_row + i, 6, read_data.month_info_dict[key_list[i]][0] + read_data.month_info_dict[key_list[i]][1], self.money_format)
                
                # 같은 달 안에서 새롭게 업데이트 하는 건지 확인 (ex: 9/10일 업뎃 후 9/29일 업뎃)
                if str(key_list[i]) + '월' == read_data.exist_file_df.iloc[14,0]:
                    same_month_pass = True
                else:
                    same_month_pass = False
                last_new_data_row = start_row + i + 1
                
            # 기존 데이터 입력
            for i in range(14, read_data.exist_file_df.index.stop):
                if same_month_pass == True:
                    same_month_pass = False
                    last_new_data_row -= 1
                elif read_data.exist_file_df.iloc[i,0][:4] == '2021':
                    self.worksheet1.write_string(last_new_data_row + i - 14, 0, read_data.exist_file_df.iloc[i,0], self.word_format)
                    self.worksheet1.write_number(last_new_data_row + i - 14, 2, read_data.exist_file_df.iloc[i,2], self.money_format)
                    self.worksheet1.write_number(last_new_data_row + i - 14, 4, read_data.exist_file_df.iloc[i,4], self.money_format)
                    self.worksheet1.write_number(last_new_data_row + i - 14, 6, read_data.exist_file_df.iloc[i,6], self.money_format)
                    self.worksheet1.write_string(last_new_data_row + i - 14, 8, read_data.exist_file_df.iloc[i,8], self.word_format)
                    self.worksheet1.write_string(last_new_data_row + i - 14, 9, read_data.exist_file_df.iloc[i,9], self.word_format)
                    last_2021_row = last_new_data_row + i - 14
                elif read_data.exist_file_df.iloc[i,0][:4] == '2020':
                    self.worksheet1.write_string(last_new_data_row + i - 14, 0, read_data.exist_file_df.iloc[i,0], self.word_format)
                    self.worksheet1.write_number(last_new_data_row + i - 14, 2, read_data.exist_file_df.iloc[i,2], self.money_format)
                    self.worksheet1.write_number(last_new_data_row + i - 14, 4, read_data.exist_file_df.iloc[i,4], self.money_format)
                    self.worksheet1.write_number(last_new_data_row + i - 14, 6, read_data.exist_file_df.iloc[i,6], self.money_format)
                    self.worksheet1.write_string(last_new_data_row + i - 14, 8, read_data.exist_file_df.iloc[i,8], self.word_format)
                    self.worksheet1.write_string(last_new_data_row + i - 14, 9, read_data.exist_file_df.iloc[i,9], self.word_format)
                    last_2020_row = last_new_data_row + i - 14 - 3
        else:
            for i in range(len(key_list)):
                if key_list[i][:4] == '2021':
                    self.worksheet1.write_string(start_row + i, 0, str(key_list[i]) + '월', self.word_format)
                    self.worksheet1.write_number(start_row + i, 2, read_data.month_info_dict[key_list[i]][0], self.money_format)
                    self.worksheet1.write_number(start_row + i, 4, read_data.month_info_dict[key_list[i]][1], self.money_format)
                    self.worksheet1.write_formula(start_row + i, 6, ('=' + 'C' + str(start_row + i + 1) + '+' + 'E' + str(start_row + i + 1)), self.money_format)
                    last_2021_row = start_row + i
                elif key_list[i][:4] == '2020':
                    self.worksheet1.write_string(start_row + i + 3, 0, str(key_list[i]) + '월', self.word_format)
                    self.worksheet1.write_number(start_row + i + 3, 2, read_data.month_info_dict[key_list[i]][0], self.money_format)
                    self.worksheet1.write_number(start_row + i + 3, 4, read_data.month_info_dict[key_list[i]][1], self.money_format)
                    self.worksheet1.write_formula(start_row + i + 3, 6, ('=' + 'C' + str(start_row + i + 4) + '+' + 'E' + str(start_row + i + 4)), self.money_format)
                    last_2020_row = start_row + i
        
        merge_range1 = 'A'+str(last_2021_row + 4) + ':' + 'B'+str(last_2021_row + 4)
        merge_range2 = 'I'+str(last_2021_row + 4) + ':' + 'J'+str(last_2021_row + 4)
        self.worksheet1.merge_range(merge_range1, '20년', self.merge_format)
        self.worksheet1.write_string(last_2021_row + 3, 2, '수입', self.bold_format)
        self.worksheet1.write_string(last_2021_row + 3, 4, '지출', self.bold_format)
        self.worksheet1.write_string(last_2021_row + 3, 6, '계', self.bold_format)
        self.worksheet1.merge_range(merge_range2, '비고', self.merge_format)

        # calculate year income and spend
        self.worksheet1.write_formula(start_row - 1, 3, ('=SUM(' + 'C' + str(start_row + 1) + ':' + 'C' + str(last_2021_row + 1) + ')'), self.money_format)
        self.worksheet1.write_formula(start_row - 1, 5, ('=SUM(' + 'E' + str(start_row + 1) + ':' + 'E' + str(last_2021_row + 1) + ')'), self.money_format)
        self.worksheet1.write_formula(start_row - 1, 7, ('=' + 'D' + str(start_row) + '+' + 'F' + str(start_row)), self.money_format)

        self.worksheet1.write_formula(last_2021_row + 3, 3, ('=SUM(' + 'C' + str(last_2021_row + 5) + ':' + 'C' + str(last_2020_row + 4) + ')'), self.money_format)
        self.worksheet1.write_formula(last_2021_row + 3, 5, ('=SUM(' + 'E' + str(last_2021_row + 5) + ':' + 'E' + str(last_2020_row + 4) + ')'), self.money_format)
        self.worksheet1.write_formula(last_2021_row + 3, 7, ('=' + 'D' + str(last_2021_row + 4) + '+' + 'F' + str(last_2021_row + 4)), self.money_format)

    def draw_execl_chart_supply(self):
        self.worksheet8 = self.workbook.add_worksheet('매집수량변동그림')

        chart = self.workbook.add_chart({'type': 'line'})

        chart.add_series({'categories': '=매집수량변동!$A2:$A301',
                         'values': '=매집수량변동!$J2:$J301',
                         'line': {'color': 'red'}})

        chart.set_style(10)

        self.worksheet8.insert_chart('D2', chart, {'x_offset': 25, 'y_offset': 10})

    def save_excel_transaction_history(self, tr_history=None, exist_tr_history=None, account_info=None):
        total_profit = 0
        self.worksheet9 = self.workbook.add_worksheet('거래내역정리')
        self.worksheet9.set_column('A:Z', 10)
        self.worksheet9.set_column('O:P', 15)

        self.worksheet9.conditional_format('O2:P500', {'type': 'data_bar',
                                                       'bar_color': '#FF0000',
                                                       'bar_negative_color': '#0000FF',
                                                       'bar_border_color': '#000000',
                                                       'bar_negative_border_color': '#000000',
                                                       'bar_direction': 'left',
                                                       'bar_axis_position': 'middle',
                                                       'criteria': '>=',
                                                       'value': 0})

        # 셀 조건으로 색깔 입히기
        self.worksheet9.conditional_format('C2:E500', {'type': 'formula',
                                                       'criteria': '=$H2>0',
                                                       'format': self.cell_bg_format})
        self.worksheet9.conditional_format('H2:H500', {'type': 'formula',
                                                       'criteria': '=$H2>0',
                                                       'format': self.cell_bg_format})
        self.worksheet9.conditional_format('K2:K500', {'type': 'formula',
                                                       'criteria': '=$H2>0',
                                                       'format': self.cell_bg_format})
        self.worksheet9.conditional_format('O2:O500', {'type': 'formula',
                                                       'criteria': '=$H2>0',
                                                       'format': self.cell_bg_format})
        self.worksheet9.conditional_format('V1:V2', {'type': 'cell',
                                                     'criteria': 'greater than',
                                                     'value': 0,
                                                     'format': self.cell_bg_format})

        # 데이터 제목 입력
        self.worksheet9.write_string(0, 0, '매도날짜 ', self.bold_format)
        self.worksheet9.write_string(0, 1, '매수날짜', self.bold_format)
        self.worksheet9.write_string(0, 2, '보유기간', self.bold_format)
        self.worksheet9.write_string(0, 3, '종목번호', self.bold_format)
        self.worksheet9.write_string(0, 4, '종목명', self.bold_format)
        self.worksheet9.write_string(0, 5, '매수수량', self.bold_format)
        self.worksheet9.write_string(0, 6, '매도수량', self.bold_format)
        self.worksheet9.write_string(0, 7, '잔여수량', self.bold_format)
        self.worksheet9.write_string(0, 8, '매수단가', self.bold_format)
        self.worksheet9.write_string(0, 9, '매도단가', self.bold_format)
        self.worksheet9.write_string(0, 10, '현재가', self.bold_format)
        self.worksheet9.write_string(0, 11, '매수금액', self.bold_format)
        self.worksheet9.write_string(0, 12, '매도금액', self.bold_format)
        self.worksheet9.write_string(0, 13, '세금(0.25%)', self.bold_format)
        self.worksheet9.write_string(0, 14, '예상 수익금액', self.bold_format)
        self.worksheet9.write_string(0, 15, '실현 수익금액', self.bold_format)
        self.worksheet9.write_string(0, 16, '수익률(%)', self.bold_format)
        self.worksheet9.write_string(0, 18, '실현' + '\n' + '누적수익', self.bold_format)
        self.worksheet9.write_string(0, 19, '실시간' + '\n' + '투자수익', self.bold_format)
        self.worksheet9.write_string(0, 20, '예상' + '\n' + '누적수익', self.bold_format)
        self.worksheet9.write_string(0, 21, '마지막' + '\n' + '업데이트', self.bold_format)
        self.worksheet9.write_string(0, 22, '실시간' + '\n' + '투자금액', self.bold_format)
        self.worksheet9.write_string(0, 23, '실시간' + '\n' + '예수금', self.bold_format)
        self.worksheet9.write_string(0, 24, '키움계좌' + '\n' + '총 자산', self.bold_format)
        self.worksheet9.write_string(0, 25, '키움계좌' + '\n' + '예상 자산', self.bold_format)

        start_row = len(tr_history.keys())
        if exist_tr_history != None:
            for i, key in enumerate(exist_tr_history):
                self.worksheet9.write(i + start_row + 1, 0, exist_tr_history[key]['매도날짜'], self.date_format)
                self.worksheet9.write(i + start_row + 1, 1, exist_tr_history[key]['매수날짜'], self.date_format)
                self.worksheet9.write_number(i + start_row + 1, 2, exist_tr_history[key]['보유기간'], self.money_format)
                self.worksheet9.write(i + start_row + 1, 3, exist_tr_history[key]['종목번호'], self.word_format)
                self.worksheet9.write(i + start_row + 1, 4, exist_tr_history[key]['종목명'], self.word_format)
                self.worksheet9.write_number(i + start_row + 1, 5, exist_tr_history[key]['매수수량'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 6, exist_tr_history[key]['매도수량'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 7, exist_tr_history[key]['잔여수량'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 8, exist_tr_history[key]['매수단가'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 9, exist_tr_history[key]['매도단가'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 10, exist_tr_history[key]['현재가'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 11, exist_tr_history[key]['매수금액'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 12, exist_tr_history[key]['매도금액'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 13, exist_tr_history[key]['세금'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 14, exist_tr_history[key]['예상 수익금액'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 15, exist_tr_history[key]['실현 수익금액'], self.money_format)
                self.worksheet9.write_number(i + start_row + 1, 16, exist_tr_history[key]['수익률'], self.percent_format)
                total_profit = total_profit + exist_tr_history[key]['실현 수익금액']

        for i, key in enumerate(tr_history):
            self.worksheet9.write(i + 1, 0, tr_history[key]['매도날짜'], self.date_format)
            self.worksheet9.write(i + 1, 1, tr_history[key]['매수날짜'], self.date_format)
            self.worksheet9.write_number(i + 1, 2, tr_history[key]['보유기간'], self.money_format)
            self.worksheet9.write(i + 1, 3, tr_history[key]['종목번호'], self.word_format)
            self.worksheet9.write(i + 1, 4, tr_history[key]['종목명'], self.word_format)
            self.worksheet9.write_number(i + 1, 5, tr_history[key]['매수수량'], self.money_format)
            self.worksheet9.write_number(i + 1, 6, tr_history[key]['매도수량'], self.money_format)
            self.worksheet9.write_number(i + 1, 7, tr_history[key]['잔여수량'], self.money_format)
            self.worksheet9.write_number(i + 1, 8, tr_history[key]['매수단가'], self.money_format)
            self.worksheet9.write_number(i + 1, 9, tr_history[key]['매도단가'], self.money_format)
            self.worksheet9.write_number(i + 1, 10, tr_history[key]['현재가'], self.money_format)
            self.worksheet9.write_number(i + 1, 11, tr_history[key]['매수금액'], self.money_format)
            self.worksheet9.write_number(i + 1, 12, tr_history[key]['매도금액'], self.money_format)
            self.worksheet9.write_number(i + 1, 13, tr_history[key]['세금'], self.money_format)
            self.worksheet9.write_number(i + 1, 14, tr_history[key]['예상 수익금액'], self.money_format)
            self.worksheet9.write_number(i + 1, 15, tr_history[key]['실현 수익금액'], self.money_format)
            self.worksheet9.write_number(i + 1, 16, tr_history[key]['수익률'], self.percent_format)
            total_profit = total_profit + tr_history[key]['실현 수익금액']

        self.worksheet9.write_number(1, 18, total_profit, self.money_format) # 실현 누적수익
        self.worksheet9.write_number(1, 19, account_info[0], self.money_format) # 실시간 투자수익
        self.worksheet9.write_number(1, 20, total_profit + account_info[0], self.money_format) # 예상 누적수익
        self.worksheet9.write(1, 21, datetime.today().strftime("%Y-%m-%d"), self.date_format) # 마지막 업데이트
        self.worksheet9.write_number(1, 22, account_info[1], self.money_format) # 실시간 투자금액
        self.worksheet9.write_number(1, 23, account_info[2], self.money_format) # 예수금
        self.worksheet9.write_number(1, 24, account_info[3], self.money_format) # 키움계좌 총 자산
        self.worksheet9.write_number(1, 25, account_info[4], self.money_format) # 키움계좌 예상 자산

    def save_exel_file(self):
        self.workbook.close()