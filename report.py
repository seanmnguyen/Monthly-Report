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
        sheet = []  # 2D matrix holding sheet data
        date = ""  # "" == initial date not set, keeps track of current date
        date_start = 0  # index of the first row for the current date
        date_end = 0  # index of the last row for the current date
        date_indices = []
        first_day = True

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
                row_info = [] * final.NUM_COLS # holds info for each cell in row

                if date == "":  # get first date
                    date = row[pay_rep.DATE]
                else:  # compare for date changes
                    temp = row[pay_rep.DATE]
                    if date != temp:  # new date --> insert old date + blank line
                        # blank_row = [""] * final.NUM_COLS
                        # csv_writer.writerow(date, blank_row[1:])
                        # csv_writer.writerow(blank_row) 
                        # date = temp  # update to new date
                        # # TODO: add two lines, first line has date in row[0], and totals in rest of cols
                        # # second line = cumulative totals --> make separate function
                        # # NOTE: Total col = sum(CA, RB)
                        # # NOTE: T col = sum(G + C)
                        date_indices.append(date_end + 1)
                        if first_day:
                            end_of_date(csv_writer, date, date_start, date_end, True, date_indices)
                            first_day = False
                            # reset day index counter
                            date_start = date_end + 3  # skip day total and line space
                            date_end += 3
                        else:
                            end_of_date(csv_writer, date, date_start, date_end, False, date_indices)
                            # reset day index counter
                            date_start = date_end + 4  # skip day total, cum total, and line space
                            date_end += 4
                    else:  # same date, increment date end index counter
                        date_end += 1

                # add the rest of the information into the row
                row_info[final.PATIENT] = row[pay_rep.PATIENT]
                row_info[final.U_C] = parse_u_c(row)
                row_info[final.CASH] = row[pay_rep.CASH]
                row_info[final.CHECK] = row[pay_rep.CHECK]
                row_info[final.CREDIT_CARD] = parse_credit_card(row)
                row_info[final.DEBIT] = row[pay_rep.DEBIT_CARD]
                row_info[final.REINBURSE] = row[pay_rep.PAID_INS]
                # row_info[final.INS] = parse_insurance()  # from invoice sheet
                row_info[final.REFUND] = ""
                row_info[final.TOTAL] = ""
                # info using billing codes below
                codes = row[pay_rep.BILLING_CODES]
                codes_arr = codes.split("|")
                contacts = parse_c(codes_arr)
                if contacts:  # if parse_c returns true
                    row_info[final.C] = 1  # set c to 1
                    row_info[final.G] = ""  # automatically set g to empty
                    row_info[final.F] = ""  # automatically set F to empty
                else:
                    row_info[final.G] = parse_g(codes_arr)
                    row_info[final.F] = parse_f(codes_arr)
                row_info[final.I] = parse_i(codes_arr)
                row_info[final.D] = parse_d(codes_arr)
                row_info[final.T] = parse_t(codes_arr)
                row_info[final.O] = parse_o()
                row_info[final.L] = parse_l()
                row_info[final.DK] = parse_dk()
                row_info[final.M] = parse_m()
                row_info[final.S] = parse_s()
                row_info[final.D3] = parse_d3()
                row_info[final.OA] = parse_oa()
                row_info[final.N] = parse_n()
                row_info[final.P] = parse_p()
                row_info[final.T] = ""

                # finish parsing row, add to matrix
                sheet.append(row_info)



# takes 1 row of information, returns the Fees Total column
# but, if the value in Paid Ins is non-zero, returns Fees Total - 70
def parse_u_c(info):
    fees_total = int(info[pay_rep.FEES_TOTAL])
    paid_ins = int(info[pay_rep.PAID_INS])
    if paid_ins > 0:  # value in Paid Ins is non-zero
        return fees_total - 70
    return fees_total

# takes 1 row, returns the amount paid via credit card (0 possible)
# there are 4 possible credit cards: VISA, Master, Amex, Discover
# at most 1 will be used, but it is possible for none to be used
def parse_credit_card(info):
    for i in range(pay_rep.VISA, pay_rep.DISCOVER + 1, 1):
        if int(info[i]) > 0:
            return info[i]
    return 0

# takes in a list of the billing codes
# returns true if the list contains 4, 14, s0, or s1 AND
# SPH, PREM, STY, or CUS; returns 0 otherwise
# NOTE: if return 1, then parse_g automatically returns 0 or false
def parse_c(codes:list):
    # key1 = ["4", "14", "s0", "s1"]  # first key checks
    # key2 = ["SPH", "PREM", "STY", "CUS"]  # second key checks
    # matches = [s for s in codes if s in key1]  # make array for matches in key1
    # if len(matches) > 0:  # if there are any matches, check key2
    #     return len([s for s in codes if s in key2]) > 0 
    # return False

    if "4" in codes or "14" in codes or "s0" in codes or "s1" in codes:
        if "SPH" in codes or "PREM" in codes or "STY" in codes or "CUS" in codes:
            return True
    return False

# takes in list of billing codes
# returns str "1" if list contains 4, 14, s0, or s1; empty string "" otherwise
# NOTE: only run if C not set
def parse_g(codes:list):
    return "4" in codes or "14" in codes or "s0" in codes or "s1" in codes

# takes in list of billing codes
# returns str "1" if list contains "SPH, PREM, STY, or CUS"; 
# returns empty string "" otherwise
# NOTE: only run if C not set
def parse_f(codes:list):
    if "SPH" in codes or "PREM" in codes or "STY" in codes or "CUS" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "OPTOS"; empty string "" otherwise
def parse_i(codes:list):
    if "OPTOS" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "DIL"; empty string "" otherwise
def parse_d(codes:list):
    if "DIL" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "Topo"; empty string "" otherwise
def parse_t(codes:list):
    if "Topo" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "201", "202", "203", "211", "212" or "213"
def parse_o(codes:list):
    for i in range(201, 204, 1):
        if str(i) in codes or str(i + 10) in codes:
            return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "LAK"; empty string "" otherwise
def parse_l(codes:list):
    if "LAK" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "DEK"; empty string "" otherwise
def parse_dk(codes:list):
    if "DEK" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "HMask"; empty string "" otherwise
def parse_m(codes:list):
    if "HMask" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "HAS"; empty string "" otherwise
def parse_s(codes:list):
    if "HAS" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "De3"; empty string "" otherwise
def parse_d3(codes:list):
    if "De3" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "OATP"; empty string "" otherwise
def parse_oa(codes:list):
    if "OATP" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "4" or "s0"; empty string "" otherwise
def parse_n(codes:list):
    if "4" in codes or "s0" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "14" or "s1"; empty string "" otherwise
def parse_p(codes:list):
    if "14" in codes or "s1" in codes:
        return "1"
    return ""

def end_of_date(sheet_matrix, date, start_date, end_date, first_day, date_list):
    day_total = [""] * final.NUM_COLS
    day_total[final.PATEINT] = date
    for col in range(final.U_C, final.REINBURSE + 1, 1):
        day_total[col] = sum_col(sheet_matrix, col, start_date, end_date)
    # day_total[final.U_C] = sum_col(sheet_matrix, final.U_C, start_date, end_date)
    # day_total[final.CASH] = sum_col(sheet_matrix, final.CASH, start_date, end_date)
    # day_total[final.CHECK] = sum_col(sheet_matrix, final.CHECK, start_date, end_date)
    # day_total[final.CREDIT_CARD] = sum_col(sheet_matrix, final.CREDIT_CARD, start_date, end_date)
    # day_total[final.DEBIT] = sum_col(sheet_matrix, final.DEBIT, start_date, end_date)
    # day_total[final.REINBURSE] = sum_col(sheet_matrix, final.REINBURSE, start_date, end_date)
    day_total[final.REFUND] = sum_col(sheet_matrix, final.REFUND, start_date, end_date)
    day_total[final.TOTAL] = sum_row(day_total, final.CASH, final.REINBURSE)
    
    for col in range(final.GLASSES, final.PREVIOUS_PAT + 1, 1):
        day_total[col] = sum_col(sheet_matrix, col, start_date, end_date)
    # day_total[final.GLASSES] = sum_col(sheet_matrix, final.GLASSES, start_date, end_date)
    # day_total[final.CONTACTS] = sum_col(sheet_matrix, final.CONTACTS, start_date, end_date)
    # day_total[final.IMAGE] = sum_col(sheet_matrix, final.IMAGE, start_date, end_date)
    # day_total[final.DILATION] = sum_col(sheet_matrix, final.DILATION, start_date, end_date)
    # day_total[final.TOPOGRAPHY] = sum_col(sheet_matrix, final.TOPOGRAPHY, start_date, end_date)
    # day_total[final.FITTING] = sum_col(sheet_matrix, final.FITTING, start_date, end_date)
    # day_total[final.OFFICE_VISIT] = sum_col(sheet_matrix, final.OFFICE_VISIT, start_date, end_date)
    # day_total[final.LASIK] = sum_col(sheet_matrix, final.LASIK, start_date, end_date)
    # day_total[final.DRY_EYE_KIT] = sum_col(sheet_matrix, final.DRY_EYE_KIT, start_date, end_date)
    # day_total[final.MASK] = sum_col(sheet_matrix, final.MASK, start_date, end_date)
    # day_total[final.SPRAY] = sum_col(sheet_matrix, final.SPRAY, start_date, end_date)
    # day_total[final.D3] = sum_col(sheet_matrix, final.D3, start_date, end_date)
    # day_total[final.OA] = sum_col(sheet_matrix, final.OA, start_date, end_date)
    # day_total[final.NEW_PAT] = sum_col(sheet_matrix, final.NEW_PAT, start_date, end_date)
    # day_total[final.PREVIOUS_PAT] = sum_col(sheet_matrix, final.PREVIOUS_PAT, start_date, end_date)
    day_total[final.TOTAL_PAT] = day_total[final.GLASSES] + day_total[final.CONTACTS]

    sheet_matrix.append(day_total)

    if not first_day:  # not first day, add cumlative row
        cum_total = [""] * final.NUM_COLS
        for col in range(final.U_C, final.REINBURSE + 1, 1):
            cum_total[col] = sum_cum(sheet_matrix, col, date_list)
        cum_total[final.REFUND] = sum_cum(sheet_matrix, final.REFUND, date_list)
        cum_total[final.TOTAL] = sum_row(cum_total, final.CASH, final.REINBURSE)

        for col in range(final.GLASSES, final.PREVIOUS_PAT + 1, 1):
            cum_total[col] = sum_cum(sheet_matrix, col, date_list)
        cum_total[final.TOTAL_PAT] = cum_total[final.GLASSES] + cum_total[final.CONTACTS]

        sheet_matrix.append(cum_total)

    # add blank row for formatting
    sheet_matrix.append([""] * final.NUM_COLS)  


def sum_col(sheet_matrix, col, start, end):
    total = 0
    for row in range(start, end + 1, 1):
        if sheet_matrix[row][col] != "":
            total += int(sheet_matrix[row][col])
    return total

def sum_row(row, start, end):
    total = 0
    for col in range(start, end + 1, 1):
        if row[col] != "":
            total += int(row[col])
    return total

def sum_cum(sheet_matrix, col, date_indices):
    total = 0
    for index in date_indices:
        if sheet_matrix[index][col] != "":
            total += int(sheet_matrix[index][col])
    return total