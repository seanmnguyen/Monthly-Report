# header for the final spreadsheet
HEADER = ["Patients", "U & C", "CA", "CK", "CC", "DB", "RB", "INS",
                "IW", "RF", "Total", "G", "C", "I", "D", "T", "F", 
                "O", "L", "DK", "M", "S", "D3", "D3L", "GT", "GTL", 
                "OR", "OI", "OM", "OA", "ON", "OT", "N", "P", "T"]
NUM_COLS = 35  # number of columns in final sheet

# constannts for the index of each column in the final report sheet
PATIENT = 0
U_C = 1  # U & C
CASH = 2  # CA
CHECK = 3  # CK
CREDIT_CARD = 4  # CC
DEBIT = 5  # DB
REINBURSE = 6  # RB
INS = 7  # INS --> from invoice sheet
IW = 8  # IW --> U&C - TOTAL
REFUND = 9  # RF --> leave blank
TOTAL = 10  # TOTAL --> leave blank
GLASSES = 11  # G --> 4, 14, s0, s1 Billing Codes
CONTACTS = 12  # C
IMAGE = 13  # I --> OPTOS
DILATION = 14  # D --> DIL
TOPOGRAPHY = 15  # T --> Topo
FITTING = 16  # F --> SPH, PREM, STY, or CUS w/o 4, 14, s0, or s1
OFFICE_VISIT = 17  # O
LASIK = 18  # L
DRY_EYE_KIT = 19  # DK
MASK = 20  # M
SPRAY = 21  # S
D3 = 22  # D3
D3L = 23  # D3L --> De3L billing code
GT = 24  # GT --> GTT billing code
GTL = 25  # GTL --> GTTL billing code
OR = 26  # OR --> OHR billing code
OI = 27  # OI --> OHI billing code
OM = 28  # OM --> OMGA billing code
OA = 29  # OA --> OA billing code
ON = 30  # ON --> OHN billing code
OT = 31  # OT --> OATP billing code
NEW_PAT = 32  # N
PREVIOUS_PAT = 33  # P
TOTAL_PAT = 34  # T