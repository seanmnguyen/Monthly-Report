from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askopenfile, askopenfilenames
from datetime import datetime
import csv
from funcs import *
import pay_rep
import final

DAYS = 1000


def parse_pay_rep(pay_rep_file, invoice, final_name="Monthly Report Final.csv"):
    # open the pay rep file using csv reader
    with open(pay_rep_file, mode='r') as in_file:
        sheet = [[""]] * (len(pd.read_csv(in_file)) + DAYS) # 2D matrix holding sheet data
        in_file.close()

    with open("payRep08_06_2023.csv", mode='r') as in_file:
        csv_reader = csv.reader(in_file)
        # try:
        #     print(in_file.readline())
        #     csv_reader = csv.reader(in_file)
        #     print("open successful: " + pay_rep_file)
        # except(FileNotFoundError):
        #     print("invalid csv file for pay rep")
        #     return str(FileNotFoundError)
        
        # create new file for the final spreadsheet
        try:
            out_file = open(final_name, mode="w", newline="")
        except(FileNotFoundError):
            print("invalid new name for report")
            return str(FileNotFoundError)
        try:
            csv_writer = csv.writer(out_file, dialect="excel")
        except(FileNotFoundError):
            print("error with csv writer for out file")
            return str(FileNotFoundError)

        print("test")
        header = True  # True == header is needed, False == header not needed
        col_check = False  # True == number of columns match, False == no match
        # print(sheet)
        # print(len(sheet))
        date = ""  # "" == initial date not set, keeps track of current date
        date_start = 0  # index of the first row for the current date
        date_end = 0  # index of the last row for the current date
        date_indices = []
        first_day = True
        sheet_index = 0

        # iterate over each row, add header to new file, parse info
        for row in csv_reader:
            # error checks input file, compares number of columns
            if not col_check and len(row) < pay_rep.NUM_COLS:
                print("parse_pay_rep --> file column nums mismatch: " + pay_rep)
                in_file.close()
                out_file.close()
                return None
            else:
                col_check = True
            
            # if first row, add header
            if header is True:
                csv_writer.writerow(final.HEADER)
                header = False
            else:  # not first row, parse info from pay rep to final
                row_info = [""] * final.NUM_COLS # holds info for each cell in row

                if date == "":  # get first date
                    date = row[pay_rep.DATE]
                else:  # compare for date changes
                    temp = row[pay_rep.DATE]
                    if date != temp:  # new date --> insert old date + blank line
                        date_indices.append(date_end + 1)
                        if first_day:
                            end_of_date(sheet, sheet_index, date, date_start, date_end, True, date_indices)
                            first_day = False
                            # reset day index counter
                            date_start = date_end + 3  # skip day total and line space
                            date_end += 3
                            sheet_index += 2
                        else:
                            end_of_date(sheet, sheet_index, date, date_start, date_end, False, date_indices)
                            # reset day index counter
                            date_start = date_end + 4  # skip day total, cum total, and line space
                            date_end += 4
                            sheet_index += 3
                        date = temp
                    else:  # same date, increment date end index counter
                        date_end += 1
                        
                # add the rest of the information into the row
                name = row[pay_rep.PATIENT]
                if name != "Totals":  # check name first to see if at the end of the file
                    row_info[final.PATIENT] = name
                    row_info[final.U_C] = parse_u_c(row)
                    row_info[final.CASH] = row[pay_rep.CASH]
                    row_info[final.CHECK] = row[pay_rep.CHECK]
                    row_info[final.CREDIT_CARD] = parse_credit_card(row)
                    row_info[final.DEBIT] = row[pay_rep.DEBIT_CARD]
                    row_info[final.REINBURSE] = row[pay_rep.PAID_INS]
                    insurance = parse_insurance(invoice, row[pay_rep.PATIENT])
                    if insurance is None:
                        print(row[pay_rep.PATIENT] + ": insurance not found")
                        return None
                    row_info[final.INS] = insurance  # from invoice sheet
                    row_info[final.REFUND] = ""
                    row_info[final.TOTAL] = ""
                    # info using billing codes below
                    codes = row[pay_rep.BILLING_CODES]
                    codes_arr = codes.split("|")
                    contacts = parse_c(codes_arr)
                    if contacts:  # if parse_c returns true
                        row_info[final.CONTACTS] = 1  # set c to 1
                        row_info[final.GLASSES] = ""  # automatically set g to empty
                        row_info[final.FITTING] = ""  # automatically set F to empty
                    else:
                        row_info[final.GLASSES] = parse_g(codes_arr)
                        row_info[final.FITTING] = parse_f(codes_arr)
                    row_info[final.INS] = parse_i(codes_arr)
                    row_info[final.DILATION] = parse_d(codes_arr)
                    row_info[final.TOPOGRAPHY] = parse_t(codes_arr)
                    row_info[final.OFFICE_VISIT] = parse_o(codes_arr)
                    row_info[final.LASIK] = parse_l(codes_arr)
                    row_info[final.DRY_EYE_KIT] = parse_dk(codes_arr)
                    row_info[final.MASK] = parse_m(codes_arr)
                    row_info[final.SPRAY] = parse_s(codes_arr)
                    row_info[final.D3] = parse_d3(codes_arr)
                    row_info[final.OA] = parse_oa(codes_arr)
                    row_info[final.NEW_PAT] = parse_n(codes_arr)
                    row_info[final.PREVIOUS_PAT] = parse_p(codes_arr)
                    row_info[final.TOTAL_PAT] = ""

                # finish parsing row, add to matrix
                sheet[sheet_index] = row_info
                sheet_index += 1
                print(sheet)
                print(len(sheet))
    # write entire sheet matrix, close files, return new sheet name
    csv_writer.writerows(sheet)
    in_file.close()
    out_file.close()
    print("all done")
    return final_name


# takes invoice file name and patient's name
# returns the insurance type for the patient
def parse_insurance(invoice_sheet, pat_name):
    name_col = 1  # column index for patient name
    ins_col = 4  # column index for patient insurance

    with open(invoice_sheet, mode='r') as in_file:
        try:
            invoice_reader = csv.reader(in_file)
        except(FileNotFoundError):
            print("invalid invoice file")
            return str(FileNotFoundError)
        
        for row in invoice_reader:
            if row[name_col] == pat_name:
                return row[ins_col]

    return None

# takes 1 row of information, returns the Fees Total column
# but, if the value in Paid Ins is non-zero, returns Fees Total - 70
def parse_u_c(info):
    fees_total = float(info[pay_rep.FEES_TOTAL].replace(",", ''))
    paid_ins = float(info[pay_rep.PAID_INS].replace(",", ''))
    if paid_ins > 0:  # value in Paid Ins is non-zero
        return int(fees_total) - 70
    return int(fees_total)

# takes 1 row, returns the amount paid via credit card (0 possible)
# there are 4 possible credit cards: VISA, Master, Amex, Discover
# at most 1 will be used, but it is possible for none to be used
def parse_credit_card(info):
    for i in range(pay_rep.VISA, pay_rep.DISCOVER + 1, 1):
        if float(info[i].replace(",", '')) > 0:
            return info[i]
    return 0

# takes in a list of the billing codes
# returns true if the list contains 4, 14, s0, or s1 AND
# SPH, PREM, STY, or CUS; returns 0 otherwise
# NOTE: if return 1, then parse_g automatically returns 0 or false
def parse_c(codes:list):
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

def end_of_date(sheet_matrix, sheet_index, date, start_date, end_date, first_day, date_list):
    day_total = [""] * final.NUM_COLS
    day_total[final.PATIENT] = date
    for col in range(final.U_C, final.REINBURSE + 1, 1):
        day_total[col] = sum_col(sheet_matrix, col, start_date, end_date)
    day_total[final.REFUND] = sum_col(sheet_matrix, final.REFUND, start_date, end_date)
    day_total[final.TOTAL] = sum_row(day_total, final.CASH, final.REINBURSE)
    
    for col in range(final.GLASSES, final.PREVIOUS_PAT + 1, 1):
        day_total[col] = sum_col(sheet_matrix, col, start_date, end_date)
    day_total[final.TOTAL_PAT] = day_total[final.GLASSES] + day_total[final.CONTACTS]

    sheet_matrix[sheet_index] = day_total
    sheet_index += 1

    if not first_day:  # not first day, add cumlative row
        cum_total = [""] * final.NUM_COLS
        for col in range(final.U_C, final.REINBURSE + 1, 1):
            cum_total[col] = sum_cum(sheet_matrix, col, date_list)
        cum_total[final.REFUND] = sum_cum(sheet_matrix, final.REFUND, date_list)
        cum_total[final.TOTAL] = sum_row(cum_total, final.CASH, final.REINBURSE)

        for col in range(final.GLASSES, final.PREVIOUS_PAT + 1, 1):
            cum_total[col] = sum_cum(sheet_matrix, col, date_list)
        cum_total[final.TOTAL_PAT] = cum_total[final.GLASSES] + cum_total[final.CONTACTS]

        sheet_matrix[sheet_index] = cum_total


def sum_col(sheet_matrix, col, start, end):
    total = 0.0
    print("\n\nsheet row count: " + str(len(sheet_matrix)))
    print("start: " + str(start) + ", end: " + str(end))
    for row in range(start, end + 1, 1):
        print("row: " + str(row) + ", col: " + str(col) + "\n\n")
        if sheet_matrix[row][col] != "":
            total += float(sheet_matrix[row][col])
    return int(total)

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


test_invoice = "invoice08_06_2023.csv"
test_pay_rep = "payRep08_06_2023.csv" 
res = parse_pay_rep(test_pay_rep, test_invoice, "test_final.csv")
print(res)
