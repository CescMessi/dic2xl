import openpyxl as xl
import os


dictionary = input('Please input the dictionary: ')
excel_file = input('Please input the Excel file: ')
col_num = int(input('Please input the col number (from 1): '))
row_num = int(input('Please input the start row number (from 1): '))


wb = xl.load_workbook(excel_file)
sheets = wb.get_sheet_names()
sheet0 = wb.get_sheet_by_name(sheets[0])

for root, dirs, files in os.walk(dictionary, topdown=False):
    for name in dirs:
        # print(name)
        # print(os.path.join(root, name))
        path = os.path.join(root, name)
        if name == 'fw' or name == 'fW' or name == 'FW' or name == 'Fw':
            print(path)
            sheet0.cell(row = row_num, column = col_num).value = ('=HYPERLINK("%s")'%path)
            row_num += 1

print('done!')
wb.save(filename = excel_file)

