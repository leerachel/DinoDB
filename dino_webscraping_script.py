import xlsxwriter
import openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup

# Excel File
path = r'C:\Users\user\Desktop\dino_webscraping_results.xlsx'
book = openpyxl.load_workbook(path)
writer = pd.ExcelWriter(path, engine = 'openpyxl')
writer.book = book

# Dino website to scrap from
URL = 'https://www.thoughtco.com/dinosaurs-a-to-z-1093748'
page = requests.get(URL)
soup = BeautifulSoup(page.content, 'html.parser')
dino_name_list =  []
dino_des_list =  []
id_num = 12
ch = 'A'
id_string = 'mntl-sc-block_1-0-12'

# Function to sift out unrelated paragraphs
def unrelated_set_id_num( x ): 
    if x == 78:
        return_num = 79
    elif x == 427:
        return_num = 432
    elif x == 642:
        return_num = 647
    elif x == 809:
        return_num = 814
    elif x == 1031:
        return_num = 1032
    elif x == 1111:
        return_num = 1114
    elif x == 1407:
        return_num = 1408
    elif x == 1431:
        return_num = 1434
    else:
        return_num = 0
    return return_num

while (True):
    if unrelated_set_id_num(id_num):
        id_num = unrelated_set_id_num(id_num)
        id_string = 'mntl-sc-block_1-0-' + str(id_num)
        continue
    result = soup.find(id=id_string)
    if result is None: # Reached letter Z
        df = pd.DataFrame( {'Description' : dino_des_list, 'Name' : dino_name_list} )
        df.to_excel(writer, sheet_name=str(ch))
        sheet = writer.book.active
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 15
        print("############################## " + id_string)
        writer.save()
        break
    else:
        dino_full = result.find('p') # a dino entry
        if dino_full is not None: # a dino entry
            dino_full_string = dino_full.text.strip()
            if id_num == 727 or id_num == 1327: # missing '-'
                index = dino_full_string.find(' ')
                dino_name = dino_full_string[:index]
                dino_description = dino_full_string[index+1:]
            else:
                index = dino_full_string.find('-')
                dino_name = dino_full_string[:index]
                dino_description = dino_full_string[index+2:]
            print(dino_name)
            print(dino_description)
            print()
            dino_name_list.append(dino_name)
            dino_des_list.append(dino_description)      

        else: 
            df = pd.DataFrame( {'Description' : dino_des_list, 'Name' : dino_name_list} )
            df.to_excel(writer, sheet_name=str(ch))
            sheet = book.get_sheet_by_name(str(ch))
            sheet.column_dimensions['B'].width = 30
            sheet.column_dimensions['C'].width = 15
            ch = chr(ord(ch) + 1) # increment by char + 1
            dino_name_list.clear()
            dino_des_list.clear()
            id_num += 1
            id_string = 'mntl-sc-block_1-0-' + str(id_num)
            writer.save()
            continue

    id_num += 2
    id_string = 'mntl-sc-block_1-0-' + str(id_num)
# end of while loop

writer.save()
writer.close()
print("Should end with id # 1571")