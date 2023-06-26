from bank_project_excel_utils import write_excel_simple
import logging

logging.basicConfig(filename='logfile.log', level=logging.DEBUG, format='%(asctime)s %(levelname)s: %(message)s')
# logging.basicConfig(filename='logfile.log', level=logging.DEBUG, format='%(message)s')

new_files_path="output/OutputTextFile.txt"
txt_to_xlsx_file_path = "OutputTxt-to-xlsx-added-column.txt"

# ----------------------- Reading Txt File 1 ----------------------- 
def read_txt_file1(txt_file1_path, num_rows):
    logging.info("Reading data from the text file ....")
    print("Reading data from the text file ....")
    
    # opens the txt file 
    with open(txt_file1_path, 'r') as file:
        data = []
        count = 0
        # loops through each row which is signaled when it reaches the 'r' at the end of the line, which is not visible in plain text 
        for i in file:
            count+=1
            # reads the line and strips the line from white spaces at the end including the new line characters
            line = i.strip()
            data.append(line)

    logging.info("     STARTED - Sorting data from txt file ....")
    print("     STARTED - Sorting data from txt file ....")
    
    sorted_data = sorted(data,key=lambda x: int(x[:18]) if x else float('inf'))
    
    logging.info("     COMPLETE - Sorting text data Complete ")   
    print("     COMPLETE - Sorting text data Complete ") 
    
    logging.debug(count)
    logging.info("Finished reading from the text file ....")
    print("Finished reading from the text file ....")
    
    return sorted_data

# ----------------------- Reading Txt File 2 ----------------------- 
def read_txt_file2(txt_file2_path, num_rows):
    logging.info("Reading data from the text file ....")
    print("Reading data from the text file ....")
    
    start_index = [1,497]
    end_index = [19,514]
    
    # opens the txt file 
    with open(txt_file2_path, 'r') as file:
        data = []
        count = 0
        # loops through each row which is signaled when it reaches the 'r' at the end of the line, which is not visible in plain text 
        for i in file:
            count+=1
            # reads the line and strips the line from white spaces at the end including the new line characters
            line = i.strip()
            if line: 
                row_data = [line[start:end] for start, end in zip(start_index,end_index)]
                data.append(row_data)

    logging.info("     STARTED - Sorting data from txt file ....")
    print("     STARTED - Sorting data from txt file ....")
    
    sorted_data = sorted(data,key=lambda x: int(x[0]))
    
    logging.info("     COMPLETE - Sorting text data Complete ")   
    print("     COMPLETE - Sorting text data Complete ") 
    
    # print(sorted_data)
    
    logging.debug(count)
    logging.info("Finished reading from the text file ....")
    print("Finished reading from the text file ....")
    
    return sorted_data


#  ----------------------- writing to txt file -----------------------
def write_txt(erd_data, foloix_data):
    # print(erd_data)
    index = 0
    
    with open(new_files_path, "w") as file:
        for line in erd_data:
            
            #logging.info(line[:18]+ " == " +foloix_data[index][0])
            for data in foloix_data:
                if line[:18] == data[0]:
                    logging.info(line[:18]+ " == " +data[0]) 
                    logging.info(True)
                    modified_line = line[:284] + data[1] + line[284:]
                    file.write(modified_line + "\n")   
                # else:
                    # logging.info(line[:18]+ " == " +data[0]) 
                    # logging.info(False)
            # index+=1
            # if index == 100:
            #     break
        

#  ----------------------- Converting Text File to Excel File -----------------------          
def txt_to_excel():
    # start_index = [1,19,54,89,124,159,194,229,264,273,285,   302,312,316,317,338,351,352,365,366,379,380]
    # end_index = [18,53,88,123,158,193,228,263,272,284,301,   311,315,316,337,350,351,364,365,378,379,392]
    
    start_index = [0,18,53,88,123,158,193,228,263,272,284,   301,312,315,316,337,350,351,364,365,378,379]
    end_index = [18,53,88,123,158,193,228,263,272,284,301,   312,315,316,337,350,351,364,365,378,379,392]
       # opens the txt file 
    with open(txt_to_xlsx_file_path, 'r') as file:
        data = []
        count = 0
        # loops through each row which is signaled when it reaches the 'r' at the end of the line, which is not visible in plain text 
        for i in file:
            count+=1
            # reads the line and strips the line from white spaces at the end including the new line characters
            line = i.strip()
            if line: 
                row_data = [line[start:end] for start, end in zip(start_index, end_index)]
                data.append(row_data)
                                
    # logging.info(data)
    
    write_excel_simple(data)
    
    
    
    
    

