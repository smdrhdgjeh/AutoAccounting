from datetime import datetime, timedelta, date

class My_basic_Func():
    def excel_date_type_convert_str(self, excel_date=None): # 엑셀에서 날짜 읽었을때 숫자 형태일시 문자형태로 변화
        if type(excel_date) is not str:
            excel_date = date(1900, 1, 1) + timedelta(int(excel_date - 2))
            excel_date = excel_date.strftime("%Y-%m-%d")

        return excel_date