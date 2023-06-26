from bank_project_txt_utils import *
from bank_project_excel_utils import *
import logging

logging.basicConfig(filename='logfile.log', level=logging.DEBUG, format='%(asctime)s %(levelname)s: %(message)s')

# ------ File Paths ------- 
    # xlsx and txt combo
xlsx_file_path="input\ExcelTestData1.xlsx"
sheet_name="Sheet1"
txt_file_path="input\TextTestData1.txt"

    # txt and txt combo
txt_file1_path = "input\TextTestData2.txt"
txt_file2_path = "input\TextTestData3.txt"


num_rows=200
num_cols=3

# ----------------------- main function -----------------------
def main():
    print("-------------------------- PROGRAM RUNNING --------------------------")
    # Set the status to tell the program what to do
    # status = 0 
    status = 1  # input xlsx, txt
    #status = 2  # input xls, txt
    #status = 3  # input txt, txt
    
    if status == 1:
        excel_data, account_numbers= read_excel(xlsx_file_path, sheet_name, num_rows)
        foloix_data = read_txt_file2(txt_file_path, num_rows)
        write_excel(excel_data, foloix_data, account_numbers, num_rows)
    elif status == 2:
        excel_data, account_numbers= read_excel(xls_to_xlsx(xls_file_path), sheet_name, num_rows)
        foloix_data = read_txt_file2(txt_file_path, num_rows)
        write_excel(excel_data, foloix_data, account_numbers, num_rows)
    elif status == 3:
        erd_data = read_txt_file1(txt_file1_path, num_rows)
        foloix_data = read_txt_file2(txt_file2_path, num_rows)
        write_txt(erd_data, foloix_data)
        txt_to_excel()
    elif status == 0 :
        txt_to_excel()
    else:
        print("what happened????")
    
    
    print("-------------------------- PROGRAM FINSHED --------------------------")
        
main()
