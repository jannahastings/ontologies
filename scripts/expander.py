import os
import re
import openpyxl
import csv
import codecs
import sys
import argparse
import subprocess




## PROGRAM EXECUTION --- required argument: input file name
if __name__ == '__main__':

    parser=argparse.ArgumentParser()
    parser.add_argument('--inputExcel', '-i',help='Name of the input Excel spreadsheet file')
    
    args=parser.parse_args()

    inputFileName = args.inputExcel
    
    ROOT_DIR = "/home/tom/Documents/PROGRAMMING/Python/ontologies/Schedule/" #todo: need relative path

    if inputFileName is None :
        parser.print_help()
        sys.exit('Not enough arguments. Expected at least -i "Excel file name" ')


    wb = openpyxl.load_workbook(ROOT_DIR + inputFileName) 
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
    