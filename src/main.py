from read import My_Read
from write import My_Write
from calcul import My_Calcul

if __name__ == "__main__":
    # Class Instance
    excel_read = My_Read()
    excel_write = My_Write()
    data_calcul = My_Calcul()

    # Read accounting data
    deposit_inform_file = "D:/WorkPlan/AutoAccounting/accounting_data/2020-08-22~2021-08-22.xlsx" # Need to change
    excel_read.read_excel_file_deposit_inform(file=deposit_inform_file)

    # Read stockdata
    stock_data_file_dir = "D:/주식/Revolution/"
    excel_read.read_excel_file_stock_inform(file=stock_data_file_dir)

    # Write excel file
    save_location = "D:/WorkPlan/AutoAccounting/dist/"
    excel_write.make_exel_file(save_location=save_location, name="자산정리")
    excel_write.write_excel_total_deposit_inform(read_data=excel_read)
    excel_write.save_exel_file()