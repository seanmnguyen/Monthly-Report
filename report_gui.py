import os
import sys
import datetime
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfile
from report import *

PAY_REP_PROMPT = "<---- Select a Pay Rep File For The Month"
INVOICE_PROMPT = "<---- Select an Invoice File For The Month"
NAME_PROMPT = "Insert Name For New Monthly Report"
PAY_REP_STATUS = 0
INVOICE_STATUS = 1
CREATE_STATUS = 2
status_num = [False, False, False]  # [pay_rep, invoice, created_successfully]

# takes file from pay rep text and invoice text, creates monthly report
def create_report():
    global status_num
    pay_rep_name = pay_rep_text.get()
    invoice_name = invoice_text.get()

    if pay_rep_name == PAY_REP_PROMPT:
        print("Error: no pay rep file chosen")
        return
    if invoice_name == INVOICE_PROMPT:
        print("Error: no invoice file chosen")
        return
    
    # determine report file name
    # Step 1: get the base directory path
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # Step 2: append the inputted name and extension
    new_name = name_input.get()
    new_name = new_name.split(".")[0]
    if new_name == NAME_PROMPT or new_name == "":
        time = datetime.datetime.now().strftime("%m-%d-%Y %H_%M_%S")
        csv_name = str(time) + ".csv"
        xlsx_name = str(time) + ".xlsx"
    else:
        csv_name = new_name + ".csv"
        xlsx_name = new_name + ".xlsx"

    csv_path = os.path.join(base_dir, csv_name)
    xlsx_path = os.path.join(base_dir, xlsx_name)
    report_csv = parse_pay_rep(pay_rep_name, invoice_name, final_name=csv_path)
    report_xlsx = csv_to_excel(report_csv, xlsx_path, excel_name=xlsx_name)
    print(report_xlsx + " successfully created!")

    # clean up: delete intermediate csv file
    os.remove(csv_path)

    status_num[CREATE_STATUS] = True
    update_status()  # updates the status text field

    return report_xlsx

# opens file explorer, lets user select csv file for pay rep
def select_pay_rep():
    global status_num
    file = askopenfile(parent=root, mode="r", title="Choose a file", filetypes=[("CSV file", "*.csv")])
    if file:
        pay_rep_text.set(file.name)
        status_num[PAY_REP_STATUS] = True 
    else:  # error or file not chosen
        print("pay rep: no file chosen")
    update_status()  # updates status text field


# opens file explorer, lets user select csv file for invoice
def select_invoice():
    global status_num
    file = askopenfile(parent=root, mode="r", title="Choose a file", filetypes=[("CSV file", "*.csv")])
    if file:
        invoice_text.set(file.name)
        status_num[INVOICE_STATUS] = True
    else:  # error or file not chosen
        print("invoice: no file chosen")
    update_status()  # updates status text field


# erases text in name text entry when clicked
def erase_entry(junk):
    if name_text.get() == NAME_PROMPT:
        name_text.set("")
    

# update the status text field
def update_status():
    global status_num
    if status_num[CREATE_STATUS] is True:  # created successfully
        status_info_text.set("Status: report created successfully!")

    elif status_num[PAY_REP_STATUS] is True:
        if status_num[INVOICE_STATUS] is True:  # both true, set to "ready"
            status_info_text.set("Status: ready to create report")
        else:  # only have pay rep, missing invoice
            status_info_text.set("Status: missing invoice")
            
    else:  # pay rep missing
        if status_num[INVOICE_STATUS] is True:  # only invoice, missing pay rep
            status_info_text.set("Status: missing pay rep")
        else:  # missing both
            status_info_text.set("Status: select files")


# resets name entry, pay rep, and invoice text field to default prompt
def reset_fields():
    global status_num
    pay_rep_text.set(PAY_REP_PROMPT)
    invoice_text.set(INVOICE_PROMPT)
    name_text.set(NAME_PROMPT)
    status_num = [False, False, False]
    update_status()


# create the GUI
root = Tk()
root.title("Monthly Report Parser")
root.geometry("850x350")  # Explicit window size for testing
# frame = ttk.Frame(root, height=350, width=850, padding=15)
# frame.grid(rowspan=5, columnspan=4)
root.rowconfigure(0, minsize=60, pad=30)
root.columnconfigure(0, minsize=5, pad=10)
root.columnconfigure(1, minsize=80, pad=10)
root.columnconfigure(2, minsize=80, pad=10)
root.columnconfigure(3, minsize=80, pad=10)

# title text
title = ttk.Label(root, text="Monthly Report Parsers", font=("Helvetica", 30, "bold"))
title.grid(row=0, column=0, columnspan=4)

# button for pay rep file
pay_rep_btn = ttk.Button(root, text="Pay Rep", command=select_pay_rep, padding=10)
pay_rep_btn.grid(row=1, column=0)

# text field for pay rep file
pay_rep_text = StringVar()
pay_rep_text.set(PAY_REP_PROMPT)
pay_rep_lbl = ttk.Label(root, textvariable=pay_rep_text, font=("Helvetica", 15))
pay_rep_lbl.grid(row=1, column=1, columnspan=3)

# button for invoice file
invoice_btn = ttk.Button(root, text="Invoice", command=select_invoice, padding=10)
invoice_btn.grid(row=2, column=0)

# text field for invoice file
invoice_text = StringVar()
invoice_text.set(INVOICE_PROMPT)
invoice_lbl = ttk.Label(root, textvariable=invoice_text, font=("Helvetica", 15))
invoice_lbl.grid(row=2, column=1, columnspan=3)

# text entry for the name of the new report
name_text = StringVar()
name_text.set(NAME_PROMPT)
name_input = ttk.Entry(root, textvariable=name_text, width=75, font=("Helvetica", 12))
name_input.grid(row=3, column=0, columnspan=4)
name_input.bind("<Button-1>", erase_entry)  # makes text disappear when clicked on

# button to parse files, create monthly report
create_btn = ttk.Button(root, text="Create", command=create_report, padding=10)
create_btn.grid(row=4, column=1)

# status label
status_info_text = StringVar()
status_info_text.set("Status: select files")
status_lbl = ttk.Label(root, textvariable=status_info_text, font=("Helvetica", 12, "italic"))
status_lbl.grid(row=4, column=0)

# clear button: resets pay rep, invoice, and name entry text fields to default prompts
clear_btn = ttk.Button(root, text="Clear", command=reset_fields, padding=10)
clear_btn.grid(row=4, column=2)

# quit button
quit_btn = ttk.Button(root, text="Quit", command=root.destroy, padding=10)
quit_btn.grid(row=4, column=3)

root.mainloop()