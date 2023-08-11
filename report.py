from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askopenfile, askopenfilenames
from datetime import datetime
import csv
# from funcs import *

# header for the final spreadsheet
FINAL_HEADER = ["Patients", "U & C", "CA", "CK", "CC", "DB", "RB", 
                "INS", "RF", "Total", "G", "C", "I", "D", "T", "F", 
                "O", "L", "DK", "M", "S", "D3", "AO", "N", "P", "T"]
PAY_REP_COLS = 60  # number of columns in pay rep sheet
FINAL_COLS = 26  # number of columns in final sheet

# constannts for the index of each column in the pay rep sheet
DATE = 0
NAME = 1
SERVICE_CODES = 2
BILLING_CODES = 3
FEES = 4
FEES_TOTAL = 5
FEES_PAT = 6
FEES_PAT_TOTAL = 7
FEES_INS = 8
PAID = 9
PAID_PAT = 10
PAID_INS = 11
ADJUSTMENTS = 12
INVOICE = 13
BALANCE = 14
PREV_BALANCE = 15
PAY_TYPE = 16
CASH = 17
CHECK = 18
CREDIT_CARD = 19
INSURANCE = 20
DOCTOR = 21
SALES_PERSON = 22
DESCRIPTIONS = 23
FRAME_INFO = 24
LOC_ID = 25
INS_CC = 26




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

        header = True  # True == header is needed, False == header not needed
        col_check = False  # True == number of columns match, False == no match
        date = ""  # "" == initial date not set, keeps track of current date

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
                row_info = []  # holds info for each cell in row

                if date == "":  # get first date
                    date = row[DATE]
                else:  # compare for date changes
                    temp = row[DATE]
                    if date != temp:  # new date --> insert old date + blank line
                        blank_row = [""] * FINAL_COLS
                        csv_writer.writerow(date, blank_row[1:])
                        csv_writer.writerow(blank_row) 
                        date = temp  # update to new date

                # add the rest of the information into the row
                parse_name()