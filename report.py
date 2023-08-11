from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askopenfile, askopenfilenames
from datetime import datetime
import csv
from funcs import *
import pay_rep
import final




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
            if not col_check and len(row) < pay_rep.NUM_COLS:
                print("parse_pay_rep --> file column nums mismatch: " + pay_rep)
                in_file.close()
                out_file.close()
                return -1
            else:
                col_check = True
            
            # if first row, add header
            if header is True:
                csv_writer.writerow(final.HEADER)
                header = False
            else:  # not first row, parse info from pay rep to final
                row_info = []  # holds info for each cell in row

                if date == "":  # get first date
                    date = row[pay_rep.DATE]
                else:  # compare for date changes
                    temp = row[pay_rep.DATE]
                    if date != temp:  # new date --> insert old date + blank line
                        blank_row = [""] * final.NUM_COLS
                        csv_writer.writerow(date, blank_row[1:])
                        csv_writer.writerow(blank_row) 
                        date = temp  # update to new date

                # add the rest of the information into the row
                row_info[final.PATIENT] = parse_patient()
                row_info[final.U_C] = parse_u_c()
                row_info[final.CASH] = parse_cash()
                row_info[final.CHECK] = parse_check()
                row_info[final.CREDIT_CARD] = parse_credit_card()
                row_info[final.DEBIT] = parse_debit()
                row_info[final.REINBURSE] = parse_reinburse()
                row_info[final.INS] = parse_insurance()
                row_info[final.RF] = parse_rf()
                row_info[final.TOTAL] = parse_total()
                row_info[final.G] = parse_g()
                row_info[final.C] = parse_c()
                row_info[final.I] = parse_i()
                row_info[final.D] = parse_d()
                row_info[final.T] = parse_t()
                row_info[final.F] = parse_f()
                row_info[final.O] = parse_o()
                row_info[final.L] = parse_l()
                row_info[final.DK] = parse_dk()
                row_info[final.M] = parse_m()
                row_info[final.S] = parse_s()
                row_info[final.D3] = parse_d3()
                row_info[final.AO] = parse_ao()
                row_info[final.N] = parse_n()
                row_info[final.P] = parse_p()
                row_info[final.T] = parse_t()