from datetime import date
import pandas as pd

from read import My_Read
from write import My_Write
from calcul import My_Calcul

if __name__ == "__main__":
    # Class Instance
    excel_read = My_Read()
    excel_write = My_Write()
    data_calcul = My_Calcul()

    # Load section

    # Read accounting data
    deposit_inform_file = "D:/WorkPlan/AutoAccounting/accounting_data/2020-09-19~2021-09-19.xlsx" # Need to change
    excel_read.read_excel_file_deposit_inform(file=deposit_inform_file)
    excel_read.read_excel_file_transaction_details(file=deposit_inform_file)

    # Read stockdata
    stock_data_file_dir = "D:/주식/Revolution/"
    excel_read.read_excel_file_stock_inform(file=stock_data_file_dir)

    # Write excel file
    save_location = "D:/WorkPlan/AutoAccounting/dist/"
    excel_write.make_exel_file(save_location=save_location, name="자산정리")
    excel_write.write_excel_total_deposit_inform(read_data=excel_read)
    excel_write.write_excel_monthly_transaction_details(read_data=excel_read)
    excel_write.save_exel_file()

    print('Finish')

    ###############################
    ############# Test ############
    ###############################
    # Need to study : https://dandyrilla.github.io/2017-08-12/pandas-10min/#3-%EB%8D%B0%EC%9D%B4%ED%84%B0-%EC%84%A0%ED%83%9D%ED%95%98%EA%B8%B0-selection
    