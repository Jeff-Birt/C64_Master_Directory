import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import xlsxwriter
import csv
import os
import hashlib

root= tk.Tk()

canvasDlg = tk.Canvas(root, width = 300, height = 300,
                    bg = 'lightsteelblue2', relief = 'raised')
canvasDlg.pack()

lblName = tk.Label(root, text='DirCombiner', bg = 'lightsteelblue2')
lblName.config(font=('helvetica', 20))
canvasDlg.create_window(150, 60, window=lblName)


def getCSV ():
    global source_path
    source_path = filedialog.askopenfilename()
    
browseButton_CSV = tk.Button(text="    Source  File    ",
                             command=getCSV, bg='green', fg='white',
                             font=('helvetica', 12, 'bold'))
canvasDlg.create_window(150, 130, window=browseButton_CSV)


def convertToExcel ():
    global source_path
    disk_header = ('Size', 'File Name', 'Type', 'MD5 Hash')
    filename = os.fsdecode(source_path)
    head, tail = os.path.split(os.path.dirname(source_path))
    last_file = ""
    write_row_index  = 1
    worksheet_index = 1

    wb = xlsxwriter.Workbook('Disk Master Directory.xlsx')
    #cell_format1 = wb.add_format({'bold': True, 'font_name': 'C64 Pro Mono'})
    cell_format1 = wb.add_format({'bold': True})
    #cell_format2 = wb.add_format({'bg_color': 'yellow', 'font_name': 'C64 Pro Mono'})
    cell_format2 = wb.add_format({'bg_color': 'yellow'})
    master_header = ('Name', 'Path', 'Sheet#', 'Directory Hash', 'Duplicate')
    master_sheet = wb.add_worksheet("Master_Index")
    master_link = 'internal:' + master_sheet.name + '!A1'
    master_sheet.write_row('A1', master_header, cell_format1)
    master_sheet.set_column('A:A', 22)
    master_sheet.set_column('B:B', 40)
    master_sheet.set_column('C:C', 8)
    master_sheet.set_column('D:D', 34)
    master_sheet.set_column('E:E', 10)
    master_sheet.conditional_format('D1:D1048576', {'type': 'duplicate', 'format':cell_format2})
    master_index = 0
    
    #maybe keep track of disk names and add to 1st 'master' worksheet?
    if filename.endswith(".csv"):
            with open(filename, 'r', encoding="utf8") as f:
                reader = csv.reader(f)
                for r, row in enumerate(reader):
                    write_row_index += 1
                    for c, val in enumerate(row):
                        #skip column A in source
                        if c == 0:
                            if last_file != val:
                                master_index +=1
                                last_file = val
                                m = hashlib.md5()
                                name = os.path.basename(last_file)
                                rel_path = os.path.relpath(last_file, head)
                                worksheet_name = "Image" + str(worksheet_index)

                                master_sheet.write(master_index, 0, name)
                                master_sheet.write(master_index, 1, rel_path)
                                link = 'internal:' + worksheet_name + '!A1'
                                master_sheet.write_url(master_index, 2, link, string=worksheet_name)

                                sh = wb.add_worksheet(worksheet_name)
                                sh.write_url('A1', master_link, string='Index')
                                sh.write('B1', 'Directory Hash')
                                sh.write('A2', 'Name', cell_format1)
                                sh.write('B2', name)
                                sh.write('C2', 'Path', cell_format1)
                                sh.write('D2', rel_path)
                                sh.write_row('A4', disk_header, cell_format1)
                                sh.set_column('A:A', 8)
                                sh.set_column('B:B', 20)
                                sh.set_column('C:C', 10)
                                sh.set_column('D:D', 34)
                                write_row_index = 4
                                worksheet_index += 1
                        else:
                            if c == 1:
                                sh.write_number(write_row_index, c-1, int(val))
                                m.update(val.encode('utf-8'))
                                master_sheet.write(master_index, 3, m.hexdigest())
                            else:
                                sh.write(write_row_index, c-1, val)
                                m.update(val.encode('utf-8'))
                                master_sheet.write(master_index, 3, m.hexdigest())
                                form = '=COUNTIF(D:D, ' + 'D' + str(master_index+1) + ')>1'
                                master_sheet.write_formula(master_index, 4, form)
            
    wb.close()
    

saveAsButton_Excel = tk.Button(text='Convert CSV to Excel',
                               command=convertToExcel, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
canvasDlg.create_window(150, 180, window=saveAsButton_Excel)


root.mainloop()
