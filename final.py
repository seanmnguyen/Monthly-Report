# header for the final spreadsheet
HEADER = ["Patients", "U & C", "CA", "CK", "CC", "DB", "RB", 
                "INS", "RF", "Total", "G", "C", "I", "D", "T", "F", 
                "O", "L", "DK", "M", "S", "D3", "AO", "N", "P", "T"]
NUM_COLS = 26  # number of columns in final sheet

# constannts for the index of each column in the final report sheet
PATIENT = 0
U_C = 1  # U & C
CASH = 2  # CA
CHECK = 3  # CK
CREDIT_CARD = 4  # CC
DEBIT = 5  # DB
REINBURSE = 6  # RB
INS = 7  # INS --> from invoice sheet
REFUND = 8  # RF --> leave blank
TOTAL = 9  # TOTAL --> leave blank
GLASSES = 10  # G --> 4, 14, s0, s1 Billing Codes
CONTACTS = 11  # C
IMAGE = 12  # I --> OPTOS
DILATION = 13  # D --> DIL
TOPOGRAPHY = 14  # T --> Topo
FITTING = 15  # F --> SPH, PREM, STY, or CUS w/o 4, 14, s0, or s1
OFFICE_VISIT = 16  # O
LASIK = 17  # L
DRY_EYE_KIT = 18  # DK
MASK = 19  # M
SPRAY = 20  # S
D3 = 21  # D3
OA = 22  # OA
NEW_PAT = 23  # N
PREVIOUS_PAT = 24  # P
TOTAL_PAT = 25  # T