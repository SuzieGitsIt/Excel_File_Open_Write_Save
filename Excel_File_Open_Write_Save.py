# File:     TkinterGUI_2023-04-03
# Version:  0.0.01
# Author:   Susan Haynes
# Online References: 
#   https://xlsxwriter.readthedocs.io/
#   xlsx can write to Excel files, but can NOT read. 

#################################################      CREATE AND WRITE TO EXCEL FILE      ################################################  
import xlsxwriter
import datetime as dt

date = dt.datetime.now()

workbook = xlsxwriter.Workbook('Excel_Write.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Bonjour a tous, aujourdhui, cest le ')
worksheet.write(f"{date:%B-%d-%Y}")

workbook.close()