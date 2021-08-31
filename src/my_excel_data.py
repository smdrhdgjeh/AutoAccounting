import os
import sys
import xlrd
import xlsxwriter

from datetime import datetime, timedelta, date

from my_basic_function import My_basic_Func


class My_excel():
    def __init__(self):
        self.func = My_basic_Func() # 자주 사용하는 func 클래스 인스턴스화

    def make_exel_file(self, save_location=None, name=None):
        self.workbook = xlsxwriter.Workbook(save_location + name + '.xlsx')

        # add format
        self.bold_format = self.workbook.add_format({'bold': True,
                                                     'border': True,
                                                     'align': 'center',
                                                     'valign': 'vcenter',
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

    def save_excel_day_kiwoom_db(self, stock_basic_data=None, stock_day_bong=None):
        self.worksheet1 = self.workbook.add_worksheet('일별주가')  # 일봉차트조회

        # 데이터 제목 입력
        self.worksheet1.write_string(0, 0, '일자', self.bold_format)
        self.worksheet1.write_string(0, 1, '시가', self.bold_format)
        self.worksheet1.write_string(0, 2, '고가', self.bold_format)
        self.worksheet1.write_string(0, 3, '저가', self.bold_format)
        self.worksheet1.write_string(0, 4, '현재가', self.bold_format)
        self.worksheet1.write_string(0, 6, '거래량', self.bold_format)
        self.worksheet1.write_string(0, 7, '거래대금(백만)', self.bold_format)

        # 주가 추가 정보
        self.worksheet1.write_string(0, 9, '상장주식', self.bold_format)
        self.worksheet1.write_string(0, 10, '유통주식', self.bold_format)
        self.worksheet1.write_string(0, 11, '유통비율', self.bold_format)

        self.worksheet1.write_number(1, 9, stock_basic_data[2], self.money_format)
        self.worksheet1.write_number(1, 10, stock_basic_data[3], self.money_format)
        self.worksheet1.write_number(1, 11, stock_basic_data[4], self.percent_format)

        for i in range(len(stock_day_bong)):
            # Date 입력
            self.worksheet1.write(i + 1, 0, stock_day_bong[i][3], self.date_format)
            self.worksheet1.write_number(i + 1, 1, stock_day_bong[i][4], self.money_format)
            self.worksheet1.write_number(i + 1, 2, stock_day_bong[i][5], self.money_format)
            self.worksheet1.write_number(i + 1, 3, stock_day_bong[i][6], self.money_format)
            self.worksheet1.write_number(i + 1, 4, stock_day_bong[i][0], self.money_format)
            self.worksheet1.write_number(i + 1, 6, stock_day_bong[i][1], self.money_format)
            self.worksheet1.write_number(i + 1, 7, stock_day_bong[i][2], self.money_format)

    def draw_execl_chart_supply(self):
        self.worksheet8 = self.workbook.add_worksheet('매집수량변동그림')

        chart = self.workbook.add_chart({'type': 'line'})

        chart.add_series({'categories': '=매집수량변동!$A2:$A301',
                         'values': '=매집수량변동!$J2:$J301',
                         'line': {'color': 'red'}})

        chart.set_style(10)

        self.worksheet8.insert_chart('D2', chart, {'x_offset': 25, 'y_offset': 10})

    def read_excel_transaction_history(self, save_location=None, name=None):
        self.exist_tr_history = {}

        if os.path.exists(save_location + name + '.xlsx'):  # 해당 경로에 파일이 있는지 체크한다.
            self.workbook_read = xlrd.open_workbook(save_location + name + '.xlsx')
            self.worksheet_read = self.workbook_read.sheet_by_index(0)

            nrows = self.worksheet_read.nrows

            for i in range(nrows - 1):
                temp_key = self.worksheet_read.row_values(i + 1)[3]

                if self.worksheet_read.row_values(i + 1)[3] in self.exist_tr_history: # 중복 되는 종목 거래 데이터 유지
                    temp_key = temp_key + 'i'

                # 엑셀 형식 날짜 표시 변환
                selldate = self.func.excel_date_type_convert_str(excel_date=self.worksheet_read.row_values(i + 1)[0])
                buydate = self.func.excel_date_type_convert_str(excel_date=self.worksheet_read.row_values(i + 1)[1])

                self.exist_tr_history[temp_key] = {}
                self.exist_tr_history[temp_key].update({"매도날짜": selldate})
                self.exist_tr_history[temp_key].update({"매수날짜": buydate})
                self.exist_tr_history[temp_key].update({"보유기간": self.worksheet_read.row_values(i + 1)[2]})
                self.exist_tr_history[temp_key].update({"종목번호": self.worksheet_read.row_values(i + 1)[3]})
                self.exist_tr_history[temp_key].update({"종목명": self.worksheet_read.row_values(i + 1)[4]})
                self.exist_tr_history[temp_key].update({"매수수량": self.worksheet_read.row_values(i + 1)[5]})
                self.exist_tr_history[temp_key].update({"매도수량": self.worksheet_read.row_values(i + 1)[6]})
                self.exist_tr_history[temp_key].update({"잔여수량": self.worksheet_read.row_values(i + 1)[7]})
                self.exist_tr_history[temp_key].update({"매수단가": self.worksheet_read.row_values(i + 1)[8]})
                self.exist_tr_history[temp_key].update({"매도단가": self.worksheet_read.row_values(i + 1)[9]})
                self.exist_tr_history[temp_key].update({"현재가": self.worksheet_read.row_values(i + 1)[10]})
                self.exist_tr_history[temp_key].update({"매수금액": self.worksheet_read.row_values(i + 1)[11]})
                self.exist_tr_history[temp_key].update({"매도금액": self.worksheet_read.row_values(i + 1)[12]})
                self.exist_tr_history[temp_key].update({"세금": self.worksheet_read.row_values(i + 1)[13]})
                self.exist_tr_history[temp_key].update({"예상 수익금액": self.worksheet_read.row_values(i + 1)[14]})
                self.exist_tr_history[temp_key].update({"실현 수익금액": self.worksheet_read.row_values(i + 1)[15]})
                self.exist_tr_history[temp_key].update({"수익률": self.worksheet_read.row_values(i + 1)[16]})

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