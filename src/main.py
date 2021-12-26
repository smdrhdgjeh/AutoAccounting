from datetime import date, datetime, timedelta
import pandas as pd

from read import My_Read
from write import My_Write
from calcul import My_Calcul

if __name__ == "__main__":
    # Class Instance
    excel_read = My_Read()
    excel_write = My_Write()
    data_calcul = My_Calcul()

    # Program file location
    file_location = 'D:/WorkPlan/AutoAccounting/'

    # 가지고 있는 파일 확인
    excel_read.check_excel_data_date_and_classifying(file_location= file_location + 'accounting_data/')

    # 기존 파일 존재 유무 확인 및 업데이트 진행여부 결정
    excel_read.read_excel_file_exist_file(file= file_location + 'dist/자산정리.xlsx')
    if excel_read.file_exist_true:
        last_update_date = excel_read.exist_file_df.columns[-1]
        last_update_date = datetime.strftime(last_update_date, '%Y-%m-%d')[:11]
        if last_update_date < excel_read.last_file_date:
            update_flag = True
        else:
            update_flag = False
    else:
        update_flag = True

    print('updating?: ', update_flag)

    # 가계부 파일 읽은 뒤 자산정리 실행
    if update_flag == True:
        # Read accounting data
        excel_read.read_excel_file_deposit_inform(file=excel_read.last_file_name)
        excel_read.read_excel_file_transaction_details(file=excel_read.last_file_name)

        # Read stockdata
        stock_data_file_dir = "D:/주식/Revolution/"
        excel_read.read_excel_file_stock_inform(file=stock_data_file_dir)

        # Write excel file
        save_location = file_location + 'dist/'
        excel_write.make_exel_file(save_location=save_location, name="자산정리")
        excel_write.write_excel_total_deposit_inform(read_data=excel_read)
        excel_write.write_excel_monthly_transaction_details(read_data=excel_read)
        excel_write.save_exel_file()

    print('Finish')


    ###############################
    ########## Test Area ##########
    ###############################
    # Need to study : https://dandyrilla.github.io/2017-08-12/pandas-10min/#3-%EB%8D%B0%EC%9D%B4%ED%84%B0-%EC%84%A0%ED%83%9D%ED%95%98%EA%B8%B0-selection
    