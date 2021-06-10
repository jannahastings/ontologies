import os
from pathlib import Path
import re
import openpyxl
from openpyxl.styles import Font
import csv
import codecs
import sys
import argparse
import subprocess
import io




## PROGRAM EXECUTION --- required argument: input file name
if __name__ == '__main__':

    parser=argparse.ArgumentParser()
    parser.add_argument('--inputExcel', '-i',help='Name of the input Excel spreadsheet file')
    
    args=parser.parse_args()

    inputFileName = args.inputExcel
    
    # ROOT_DIR = "/home/tom/Documents/PROGRAMMING/Python/ontologies/Schedule/" #todo: delete, using path from args rather

    if inputFileName is None :
        parser.print_help()
        sys.exit('Not enough arguments. Expected at least -i "Excel file name" ')

    pathpath = str(Path(inputFileName).parents[0])
    basename = str(Path(inputFileName).stem)
    suffix = str(Path(inputFileName).suffix)
    # filename = os.path.basename(inputFileName)
    # basename = os.path.splitext(filename)[0]
    # filename = 
    # wb = openpyxl.load_workbook(ROOT_DIR + inputFileName) 
    wb = openpyxl.load_workbook(inputFileName) #call with full path and filename? Much better
    sheet = wb.active
    data = sheet.rows
    rows = []

    header = [i.value for i in next(data)]
    print("got header: ", header)
    for row in sheet[2:sheet.max_row]:
        values = {}
        for key, cell in zip(header, row):
            values[key] = cell.value
        if any(values.values()):
            rows.append(values)
    
    print("reached rows ", rows)
    for r in range(len(rows)):
        row = [v for v in rows[r].values()]
        if "Aggregate" in header:
            cell = row[header.index("Aggregate")]
            print("cell is: ", cell)
    

    #new sheet to save:

    save_wb = openpyxl.Workbook()
    save_sheet = save_wb.active

    for c in range(len(header)):
        save_sheet.cell(row=1, column=c+1).value=header[c]
        save_sheet.cell(row=1, column=c+1).font = Font(size=12,bold=True)
    for r in range(len(rows)):
        row = [v for v in rows[r].values()]
        # del row[0] # Tabulator-added ID column
        for c in range(len(header)):
            save_sheet.cell(row=r+2, column=c+1).value=row[c]
            

        # Generate identifiers:
        # if 'ID' in first_row:             
        #     if not row[header.index("ID")]: #blank
        #         if 'Label' and 'Parent' and 'Definition' in first_row: #make sure we have the right sheet
        #             if row[header.index("Label")] and row[header.index("Parent")] and row[header.index("Definition")]: #not blank
        #                 #generate ID here: 
        #                 nextIdStr = str(searcher.getNextId(repo_key))
        #                 id = repo_key.upper()+":"+nextIdStr.zfill(app.config['DIGIT_COUNT'])
        #                 new_id = id
        #                 for c in range(len(header)):
        #                     if c==0:
        #                         restart = True
        #                         sheet.cell(row=r+2, column=c+1).value=new_id

    #save:   
    save_wb.save(pathpath + "/" + basename + "_Expanded.xlsx")
    