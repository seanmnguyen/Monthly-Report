from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askopenfile, askopenfilenames
from datetime import datetime
# from funcs import *


FINAL_HEADER = ["Patients", "U & C", "CA", "CK", "CC", "DB", "RB", 
                "INS", "RF", "Total", "G", "C", "I", "D", "T", "F", 
                "O", "L", "DK", "M", "S", "D3", "AO", "N", "P", "T"]
PAY_REP_COLS = 60
FINAL_COLS = 26
DATE = 0
NAME = 1


def parse_pay_rep(pay_rep, final_name="Monthly Report Final"):
    # open the pay rep file using csv reader
    with open(pay_rep, mode='r') as in_file:
        try:
            csv_reader = csv.reader(in_file)
        except(IOError):
            print("invalid csv file for pay rep")
            return
        
        # create new file for the final spreadsheet
        out_file = open(final_name, mode="w", newline="")
        csv_writer = csv.writer(out_file, dialect="excel")

        header = True
        col_check = False
        date = ""
        # iterate over each row, add header to new file, parse info
        for row in csv_reader:
            # error checks input file, compares number of columns
            if not col_check and len(row) < PAY_REP_COLS:
                print("parse_pay_rep --> file column nums mismatch: " + pay_rep)
                in_file.close()
                out_file.close()
                return -1
            else:
                col_check = True
            
            # if first row, add header
            if header is True:
                csv_writer.writerow(FINAL_HEADER)
                header = False
            else:  # not first row, parse info from pay rep to final
                if date == "":  # get first date
                    date = row[DATE]
                else:  # compare for date changes
                    temp = row[DATE]
                    if date != temp:  # new date, insert blank line first
                        date = temp
                        csv_writer.writerow("" * FINAL_COLS)
                csv_writer.writerow(row[0])