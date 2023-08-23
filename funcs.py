import pandas as pd
import csv
import os


# converts an inputted csv file to an excel file
def csv_to_excel(csv_file, excel_name="Monthly Report.xlsx"):
    reader = pd.read_csv(csv_file)
    reader.to_excel(excel_name, sheet_name=excel_name, index=False)
    return excel_name


# converts an inputted excel file to a csv file
def excel_to_csv(excel_file, csv_name):
    read_file = pd.read_excel(excel_file)
    read_file.to_csv(csv_name, index=None, header=True)
    return csv_name


# takes an absolute path and returns only the last part (file/directory name)
def trunc_path(path):
    lst = path.split("/")
    return lst[len(lst) - 1]


# takes a file and returns the file extension
def get_extension(path):
    lst = path.split(".")
    return lst[len(lst) - 1]

