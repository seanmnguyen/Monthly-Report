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
                        # TODO: add two lines, first line has date in row[0], and totals in rest of cols
                        # second line = cumulative totals --> make separate function
                        # NOTE: Total col = sum(CA, RB)
                        # NOTE: T col = sum(G + C)

                # add the rest of the information into the row
                row_info[final.PATIENT] = row[pay_rep.PATIENT]
                row_info[final.U_C] = parse_u_c(row)
                row_info[final.CASH] = row[pay_rep.CASH]
                row_info[final.CHECK] = row[pay_rep.CHECK]
                row_info[final.CREDIT_CARD] = parse_credit_card(row)
                row_info[final.DEBIT] = row[pay_rep.DEBIT_CARD]
                row_info[final.REINBURSE] = row[pay_rep.PAID_INS]
                # row_info[final.INS] = parse_insurance()  # from invoice sheet
                row_info[final.RF] = ""
                row_info[final.TOTAL] = ""
                codes = row[pay_rep.BILLING_CODES]
                codes_arr = codes.split("|")
                # row_info[final.G] = parse_g()
                # row_info[final.C] = parse_c()
                # row_info[final.I] = parse_i()
                # row_info[final.D] = parse_d()
                # row_info[final.T] = parse_t()
                # row_info[final.F] = parse_f()
                # row_info[final.O] = parse_o()
                # row_info[final.L] = parse_l()
                # row_info[final.DK] = parse_dk()
                # row_info[final.M] = parse_m()
                # row_info[final.S] = parse_s()
                # row_info[final.D3] = parse_d3()
                # row_info[final.OA] = parse_oa()
                # row_info[final.N] = parse_n()
                # row_info[final.P] = parse_p()
                # row_info[final.T] = parse_t()


# takes 1 row of information, returns the Fees Total column
# but, if the value in Paid Ins is non-zero, returns Fees Total - 70
def parse_u_c(info):
    fees_total = info[pay_rep.FEES_TOTAL]
    paid_ins = info[pay_rep.PAID_INS]
    if paid_ins > 0:  # value in Paid Ins is non-zero
        return fees_total - 70
    return fees_total

# takes 1 row, returns the amount paid via credit card (0 possible)
# there are 4 possible credit cards: VISA, Master, Amex, Discover
# at most 1 will be used, but it is possible for none to be used
def parse_credit_card(info):
    for i in range(pay_rep.VISA, pay_rep.DISCOVER + 1, 1):
        if info[i] > 0:
            return info[i]
    return 0

def parse_g(codes:list):
    return ""

def parse_c(codes:list):
    pass

def parse_i(codes:list):
    pass

def parse_d(codes:list):
    pass

def parse_t(codes:list):
    pass

def parse_f(codes:list):
    pass

def parse_o(codes:list):
    pass

def parse_l(codes:list):
    pass

def parse_dk(codes:list):
    pass

def parse_m(codes:list):
    pass

def parse_s(codes:list):
    pass

def parse_d3(codes:list):
    pass

def parse_oa(codes:list):
    pass

def parse_n(codes:list):
    pass

def parse_p(codes:list):
    pass

def parse_t(codes:list):
    pass