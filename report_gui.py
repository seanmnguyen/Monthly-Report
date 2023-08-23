import os
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfile
from datetime import datetime
from report import *

PAY_REP_PROMPT = "<---- Select a Pay Rep File For The Month"
INVOICE_PROMPT = "<---- Select an Invoice File For The Month"
NAME_PROMPT = "Insert Name For New Monthly Report"
SUCCESS_STATUS = "success"
NO_PAY_REP_STATUS = "missing pay rep"
NO_INVOICE_STATUS = "missing invoice"
NO_BOTH_STATUS = "missing both files"
ERROR_STATUS = "error"

# takes file from pay rep text and invoice text, creates monthly report
def create_report():
    pay_rep_name = pay_rep_text.get()
    invoice_name = invoice_text.get()

    if pay_rep_name == PAY_REP_PROMPT:
        print("Error: no pay rep file chosen")
        return
    if invoice_name == INVOICE_PROMPT:
        print("Error: no invoice file chosen")
        return
    
    # determine report file name
    new_name = name_input.get()
    new_name = new_name.split(".")[0]
    if new_name == NAME_PROMPT or new_name == "":
        time = datetime.now().strftime("%m-%d-%Y %H_%M_%S")
        csv_name = str(time) + ".csv"
        xlsx_name = str(time) + ".xlsx"
    else:
        csv_name = new_name + ".csv"
        xlsx_name = new_name + ".xlsx"

    report_csv = parse_pay_rep(pay_rep_name, invoice_name, final_name=csv_name)
    report_xlsx = csv_to_excel(report_csv, excel_name=xlsx_name)
    print(report_xlsx + " successfully created!")

    # clean up: delete intermediate csv file
    os.remove(csv_name)

    update_status("success")  # updates the status text field

    return report_xlsx

# opens file explorer, lets user select csv file for pay rep
def select_pay_rep():
    file = askopenfile(parent=root, mode="r", title="Choose a file", filetype=[("CSV file", "*.csv")])
    if file:
        pay_rep_text.set(file.name)
    else:  # error or file not chosen
        print("pay rep: no file chosen")


# opens file explorer, lets user select csv file for invoice
def select_invoice():
    file = askopenfile(parent=root, mode="r", title="Choose a file", filetype=[("CSV file", "*.csv")])
    if file:
        invoice_text.set(file.name)
    else:  # error or file not chosen
        print("invoice: no file chosen")


# erases text in name text entry when clicked
def erase_entry(junk):
    if name_text.get() == NAME_PROMPT:
        name_text.set("")
    

# # update the status text field
# def update_status(status):
#     match status:
#         case SUCCESS_STATUS:
#             status_info_text.set("Status: report created successfully!")
#         case NO_PAY_REP_STATUS:
#             return
#         case NO_INVOICE_STATUS:
#             print("")
#         case NO_BOTH_STATUS:
#             print("")
#         case ERROR_STATUS:
#             print("")
#         case _:
#             return
#     return    


# create the GUI
root = Tk()
frame = ttk.Frame(root, height=350, width=850, padding=15)
frame.grid(rowspan=5, columnspan=4)
root.rowconfigure(0, minsize=60, pad=30)
root.columnconfigure(0, minsize=5, pad=0)
root.columnconfigure(1, minsize=80, pad=1)
root.columnconfigure(2, minsize=80, pad=10)
root.columnconfigure(3, minsize=80, pad=10)

# title text
title = ttk.Label(root, text="Monthly Report Parsers", font=("Helvectia", 30, "bold")).grid(row=0, column=0, columnspan=4)

# button for pay rep file
pay_rep_btn = ttk.Button(root, text="Pay Rep", command=select_pay_rep, padding=10)
pay_rep_btn.grid(row=1, column=0)

# text field for pay rep file
pay_rep_text = StringVar()
pay_rep_text.set(PAY_REP_PROMPT)
pay_rep_lbl = ttk.Label(root, textvariable=pay_rep_text, font=("Helvectia", 15))
pay_rep_lbl.grid(row=1, column=1, columnspan=3)

# button for invoice file
invoice_btn = ttk.Button(root, text="Invoice", command=select_invoice, padding=10)
invoice_btn.grid(row=2, column=0)

# text field for invoice file
invoice_text = StringVar()
invoice_text.set(INVOICE_PROMPT)
invoice_lbl = ttk.Label(root, textvariable=invoice_text, font=("Helvectia", 15))
invoice_lbl.grid(row=2, column=1, columnspan=3)

# text entry for the name of the new report
name_text = StringVar()
name_text.set(NAME_PROMPT)
name_input = ttk.Entry(root, textvariable=name_text, width=45, font=("Helvectia", 12))
name_input.grid(row=3, column=1, columnspan=3)
name_input.bind("<Button-1>", erase_entry)  # makes text disappear when clicked on

# button to parse files, create monthly report
create_btn = ttk.Button(root, text="Create", command=create_report, padding=10)
create_btn.grid(row=3, column=0)

# status label
status_lbl = ttk.Label(root, text="Status: ", font=("Helvectia", 10))
status_lbl.grid(row=4, column=0)

# status info text label
status_info_text = StringVar()
status_info_text.set("Select files")
status_info_lbl = ttk.Label(root, textvariable=status_info_text, font=("Helvectia", 10))
status_info_lbl.grid(row=4, column=1)

# quit button
quit_btn = ttk.Button(root, text="Quit", command=root.destroy, padding=10)
quit_btn.grid(row=4, column=3)

root.mainloop()