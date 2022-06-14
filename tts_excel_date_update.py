
from openpyxl import load_workbook
from openpyxl.styles.alignment import Alignment
from pathlib import Path
import os

date_update = input(str("Please input the required date, i.e. 18/02/2022: "))
month_update = input('Please enter the new month and year, i.e. Jan_2022: ')


folder = Path('#Insert Path#').glob('**/*.xlsx')
print(folder)

for file in folder:
    tss = str(file)
    book = load_workbook(tss)

    CB_Amounts = book['CB_Amounts']
    CB_Amounts['G2'] = date_update
    CB_Amounts['G2'].alignment = Alignment(horizontal="center")
    
    CB_Labour = book['CB_Labour']
    CB_Labour['H2'] = date_update
    CB_Labour['H2'].alignment = Alignment(horizontal="center")

    replace_date = tss[:67]
    new_name = (replace_date + month_update + '.xlsx') 

    book.title = new_name
 
    book.save(tss)
    book.close()       
    
    os.rename(tss, new_name)
 


