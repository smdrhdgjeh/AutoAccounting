from read import My_Read
from write import My_Write
from calcul import My_Calcul

if __name__ == "__main__":
    # Class Instance
    excel_read = My_Read()
    excel_write = My_Write()
    data_calcul = My_Calcul()

    # Read accounting data
    deposit_inform_file = "D:/WorkPlan/AutoAccounting/accounting_data/2020-08-22~2021-08-22.xlsx"
    excel_read.read_excel_file_deposit_inform(file=deposit_inform_file)