import csv
import pay_rep
import final
from funcs import *

DAYS = 30

'''STEPS TO CREATE EXECUTABLE FILE
1) 
'''

def parse_pay_rep(pay_rep_file, invoice, final_name="Monthly Report Final.csv"):
    # open the pay rep file using csv reader
    with open(pay_rep_file, mode='r') as in_file:
        sheet = [[""]] * (len(pd.read_csv(in_file)) + DAYS*4) # 2D matrix holding sheet data
        in_file.close()

    with open(pay_rep_file, mode='r') as in_file:
        # csv_reader = csv.reader(in_file)
        try:
            csv_reader = csv.reader(in_file)
            print("open successful: " + pay_rep_file)
        except(FileNotFoundError):
            print("invalid csv file for pay rep")
            return str(FileNotFoundError)
        
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

        header = True  # True == header is needed, False == header not needed
        col_check = False  # True == number of columns match, False == no match
        date = ""  # "" == initial date not set, keeps track of current date
        date_start = 0  # index of the first row for the current date
        date_end = 0  # index of the last row for the current date
        first_day = True
        sheet_index = 0
        prev_cum_index = 0

        # iterate over each row, add header to new file, parse info
        for row in csv_reader:
            # error checks input file, compares number of columns
            if not col_check and len(row) < pay_rep.NUM_COLS:
                print("parse_pay_rep --> file column nums mismatch: " + pay_rep_file)
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
                        if first_day:
                            prev_cum_index = end_of_date(sheet, sheet_index, date, str(date_start+2), str(date_end+2), True, prev_cum_index)
                            first_day = False
                            # reset day index counter
                            date_start = date_end + 3  # skip day total and line space
                            date_end += 3
                            sheet_index += 2
                        else:
                            prev_cum_index = end_of_date(sheet, sheet_index, date, str(date_start+2), str(date_end+2), False, prev_cum_index)
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
                    row_info[final.CASH] = string_to_float(row[pay_rep.CASH])
                    row_info[final.CHECK] = string_to_float(row[pay_rep.CHECK])
                    row_info[final.CREDIT_CARD] = parse_credit_card(row)
                    row_info[final.DEBIT] = string_to_float(row[pay_rep.DEBIT_CARD])
                    if validate_payment(row_info, row):
                        print("default payment used for: " + row_info[final.PATIENT])
                    row_info[final.REINBURSE] = string_to_float(row[pay_rep.PAID_INS])
                    insurance = parse_insurance(invoice, row[pay_rep.PATIENT])
                    if insurance is None:
                        print(row[pay_rep.PATIENT] + ": insurance not found")
                        return None
                    row_info[final.INS] = insurance  # from invoice sheet
                    row_info[final.IW] = ""
                    row_info[final.REFUND] = ""
                    row_info[final.TOTAL] = parse_total(row_info)
                    # refill IW after total calculated
                    row_info[final.IW] = parse_iw(row_info)

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
                    row_info[final.IMAGE] = parse_i(codes_arr)
                    row_info[final.DILATION] = parse_d(codes_arr)
                    row_info[final.TOPOGRAPHY] = parse_t(codes_arr)
                    row_info[final.OFFICE_VISIT] = parse_o(codes_arr)
                    row_info[final.LASIK] = parse_l(codes_arr)
                    row_info[final.DRY_EYE_KIT] = parse_dk(codes_arr)
                    row_info[final.MASK] = parse_m(codes_arr)
                    row_info[final.SPRAY] = parse_s(codes_arr)
                    row_info[final.D3] = parse_d3(codes_arr)
                    row_info[final.D3L] = parse_d3l(codes_arr)  # new additions
                    row_info[final.GT] = parse_gt(codes_arr)
                    row_info[final.GTL] = parse_gtl(codes_arr)
                    row_info[final.OR] = parse_or(codes_arr)
                    row_info[final.OI] = parse_oi(codes_arr)
                    row_info[final.OM] = parse_om(codes_arr)
                    row_info[final.OA] = parse_oa(codes_arr)
                    row_info[final.ON] = parse_on(codes_arr)  # end of new additions
                    row_info[final.OT] = parse_ot(codes_arr)
                    row_info[final.NEW_PAT] = parse_n(codes_arr)
                    row_info[final.PREVIOUS_PAT] = parse_p(codes_arr)
                    row_info[final.TOTAL_PAT] = ""
                    # finish parsing row, add to matrix
                    sheet[sheet_index] = row_info
                    sheet_index += 1

    # write entire sheet matrix, close files, return new sheet name
    csv_writer.writerows(sheet)
    in_file.close()
    out_file.close()
    return final_name


# takes invoice file name and patient's name
# returns the insurance type for the patient
def parse_insurance(invoice_sheet, pat_name):
    # name_col = 1  # column index for patient name
    # ins_col = 6  # column index for patient insurance
    name_col, ins_col = get_invoice_patient_and_ins(invoice_sheet)

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

# takes invoice file name
# returns the indices for the patient column and INS column
def get_invoice_patient_and_ins(invoice_sheet):
    patient_col, ins_col = -1, -1

    with open(invoice_sheet, mode='r') as in_file:
        try:
            invoice_reader = csv.reader(in_file)
        except FileNotFoundError as e:
            print("invalid invoice file in get_invoice_patient_and_ins")
            return str(e)
    
        header_line = next(invoice_reader)
        for index, column_name in enumerate(header_line):
            if column_name.lower().strip() == "patient":
                patient_col = index
            elif column_name.lower().strip() == "ins":
                ins_col = index
            
            # break condition, both columns found
            if patient_col >= 0 and ins_col >= 0:
                break
        return patient_col, ins_col

# takes 1 row of information, returns the Fees Total column
# but, if the value in Paid Ins is non-zero, returns Fees Total - 74
def parse_u_c(info):
    fees_total = float(info[pay_rep.FEES_TOTAL].replace(",", ''))
    paid_ins = float(info[pay_rep.PAID_INS].replace(",", ''))
    if paid_ins > 0:  # value in Paid Ins is non-zero
        return fees_total - 74
    return fees_total

# takes 1 row, returns the amount paid via credit card (0 possible)
# there are 4 possible credit cards: VISA, Master, Amex, Discover
# at most 1 will be used, but it is possible for none to be used
def parse_credit_card(info):
    sum = 0.0
    for i in range(pay_rep.VISA, pay_rep.DISCOVER + 1, 1):
        if float(info[i].replace(",", '')) > 0:
            sum += float(info[i].replace(",", ''))
    return sum

# takes 1 row_info, that is the partially filled row array, 
# returns the U&C - TOTAL
def parse_iw(row_info):
    return str(float(row_info[final.U_C]) - float(row_info[final.TOTAL]))

# takes 1 row_info, that is the partially filled row array,
# returns the sum of column C to column G (CA + CK + CC + DB + RB) 
def parse_total(row_info):
    sum = 0
    for i in range(2, 7):
        sum += row_info[i]
    return sum

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
    if "4" in codes or "14" in codes or "s0" in codes or "s1" in codes:
        return "1"
    return ""

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
    if "Hmask" in codes:
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
# returns str "1" if list contains "De3L"; empty string "" otherwise
def parse_d3l(codes:list):
    if "De3L" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "GTT"; empty string "" otherwise
def parse_gt(codes:list):
    if "GTT" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "GTTL"; empty string "" otherwise
def parse_gtl(codes:list):
    if "GTTL" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "OHR"; empty string "" otherwise
def parse_or(codes:list):
    if "OHR" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "OI"; empty string "" otherwise
def parse_oi(codes:list):
    if "OHI" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "OM"; empty string "" otherwise
def parse_om(codes:list):
    if "OMGD" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "OA"; empty string "" otherwise
def parse_oa(codes:list):
    if "OA" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "OHN"; empty string "" otherwise
def parse_on(codes:list):
    if "OHN" in codes:
        return "1"
    return ""

# takes in list of billing codes
# returns str "1" if list contains "OATP"; empty string "" otherwise
def parse_ot(codes:list):
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

# takes a string representing a number
# returns a float
def string_to_float(num:str) -> float:
    return float(num.replace(",", ''))

# takes sheet and current row info, checks if payment options are all 0
# if so, check alternative credit card column for default
# note, credit card is already a float, no need to cast or replace commas
def validate_payment(curr_row, pay_rep_row):
    if curr_row[final.CASH] == 0 and curr_row[final.CHECK] == 0 and curr_row[final.CREDIT_CARD] == 0 and curr_row[final.DEBIT] == 0:
        alt_payment = float(pay_rep_row[pay_rep.ALT_CREDIT_CARD].replace(",", ''))
        if alt_payment > 0:
            curr_row[final.CREDIT_CARD] = alt_payment
            return True
    return False

# after the last person for a day, sum up all the values from U & C to Reinbursement, along with Glasses to Previous Patient
# then, if the day is not the first day of the month, make a cumulative sum row, combining the sums from every day before
def end_of_date(sheet_matrix, sheet_index, date, start_date: str, end_date: str, first_day, prev_cum):
    day_total = [""] * final.NUM_COLS
    day_total[final.PATIENT] = date  # set column 0 to the date
    for col in range(final.U_C, final.REINBURSE + 1, 1):  # sum columns from U&C to Reinbursement
        day_total[col] = "=SUM(" + COL_LETTERS[col] + start_date + ":" + COL_LETTERS[col] + end_date + ")"
    day_total[final.REFUND] = "=SUM(" + COL_LETTERS[final.REFUND] + start_date + ":" + COL_LETTERS[final.REFUND] + end_date + ")"
    day_total_row = str(int(end_date) + 1)
    day_total[final.TOTAL] = "=SUM(" + COL_LETTERS[final.CASH] + day_total_row + ":" + COL_LETTERS[final.REINBURSE] + day_total_row + ")-" + COL_LETTERS[final.REFUND] + day_total_row

    # sum columns from Glasses to Previous Patient
    for col in range(final.GLASSES, final.PREVIOUS_PAT + 1, 1):
        day_total[col] = "=SUM(" + COL_LETTERS[col] + start_date + ":" + COL_LETTERS[col] + end_date + ")"
    day_total[final.TOTAL_PAT] = "=SUM(" + COL_LETTERS[final.GLASSES] + day_total_row + ":" + COL_LETTERS[final.CONTACTS] + day_total_row + ")"

    # add day total row into matrix
    sheet_matrix[sheet_index] = day_total
    sheet_index += 1

    # not first day, add cumlative row
    if not first_day:  
        cum_total = [""] * final.NUM_COLS
        for col in range(final.U_C, final.REINBURSE + 1, 1):
            cum_total[col] = sum_cum(COL_LETTERS[col], str(prev_cum), day_total_row)

        for col in range(final.REFUND, final.TOTAL_PAT + 1, 1):
            cum_total[col] = sum_cum(COL_LETTERS[col], str(prev_cum), day_total_row)

        sheet_matrix[sheet_index] = cum_total
        prev_cum = int(end_date) + 2
    else:
        prev_cum = int(end_date) + 1
    return prev_cum


# creates the formula for the cumulative sum for a given letter column
def sum_cum(letter: str, prev_cum: str, curr_cum: str):
    return "=" + letter + prev_cum + "+" + letter + curr_cum
