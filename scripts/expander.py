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

    header = [i.value for i in next(data)]
    print("got header: ", header)
    # for row in data:
    #     file_count = file_count+1

    # print (file, ":", file_count)
    # total_count = total_count+file_count