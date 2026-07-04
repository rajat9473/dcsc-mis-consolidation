"""
Central configuration for the DCSC MIS Consolidation tool.

The parser normalises airline-specific Excel columns into the
common MIS schema defined by the assignment.
"""

TARGET_COLUMNS = [
    "Direction",
    "Airline",
    "Month",
    "Fortnight",
    "Source File",
    "Flight No",
    "Flight Date",
    "AWB No",
    "Sfx",
    "Origin",
    "Dest",
    "Pcs",
    "Gross Wt",
    "Chg Wt",
    "Billing SHC",
    "Handling Amt (Incl GST)",
    "X-Ray Amt (Incl GST)",
    "Demurrage Amt (Incl GST)",
    "Total Amt (Incl GST)",
]


UNIQUE_KEY = [
    "Direction",
    "Airline",
    "Month",
    "Fortnight",
    "AWB No",
    "Sfx",
]


COLUMN_ALIASES = {
    "Flight No": [
        "FLIGHT NO",
        "FLIGHT NO.",
        "FLT NO",
        "FLT NO.",
        "FLIGHT_NO",
        "Flight_No",
        "Flight No",
    ],

    "Flight Date": [
        "FLIGHT DATE",
        "FLT DATE",
        "FLIGHT DT",
        "DATE OF FLIGHT",
        "FLIGHT_DATE",
        "Flight_Date",
        "Flight Date",
    ],

    "AWB No": [
        "AWB NO",
        "AWB NO.",
        "AWB NUMBER",
        "AWB",
        "MAWB NO",
        "MAWB NO.",
        "MAWB",
        "MAWB NUMBER",
        "AWB_No",
        "MAWB_No",
    ],

    "Sfx": [
        "SFX",
        "AWB SFX",
        "AWB SUFFIX",
        "SUFFIX",
        "HAWB",
        "HAWB NO",
        "HAWB NO.",
        "HAWB NUMBER",
        "Sfx",
    ],

    "Origin": [
        "ORIGIN",
        "ORG",
        "FROM",
        "STATION OF ORIGIN",
        "Origin",
    ],

    "Dest": [
        "DEST",
        "DESTINATION",
        "DST",
        "TO",
        "DESTN",
        "Dest",
    ],

    "Pcs": [
        "PCS",
        "PIECES",
        "NO OF PCS",
        "NO. OF PCS",
        "TOTAL PCS",
        "RECEIVED PCS",
        "Received_Pcs",
        "Pcs",
    ],

    "Gross Wt": [
        "GROSS WT",
        "GROSS WT.",
        "GROSS WGT",
        "GROSS WEIGHT",
        "GR WT",
        "GR. WT.",
        "GRS WG",
        "GRS WGT",
        "RECEIVED GROSS WGT",
        "Received_Gross_Wgt",
        "Gross Wt",
    ],

    "Chg Wt": [
        "CHG WT",
        "CHG WT.",
        "CHG WGT",
        "CHG WGT.",
        "CHARGEABLE WT",
        "CHARGEABLE WEIGHT",
        "CHG WEIGHT",
        "RECEIVED CHG WGT",
        "Received_Chg_Wgt",
        "Chg Wt",
    ],

    "Billing SHC": [
        "BILLING SHC",
        "SHC",
        "SPECIAL HANDLING CODE",
        "BILLING_SHC",
        "Billing SHC",
    ],

    "Handling Amt (Incl GST)": [
        "HANDLING AMT",
        "HANDLING AMOUNT",
        "HANDLING CHARGES",
        "HANDLING CHARGE",
        "HANDLING AMT (INCL GST)",
        "HANDLING AMOUNT (INCL GST)",
        "Handling Amt (Incl GST)",
    ],

    "X-Ray Amt (Incl GST)": [
        "X-RAY AMT",
        "XRAY AMT",
        "X RAY AMT",
        "X-RAY AMOUNT",
        "XRAY AMOUNT",
        "X-RAY CHARGES",
        "XRAY CHARGES",
        "X-RAY AMT (INCL GST)",
        "X-Ray Amt (Incl GST)",
    ],

    "Demurrage Amt (Incl GST)": [
        "DEMURRAGE AMT",
        "DEMURRAGE AMOUNT",
        "DEMURRAGE",
        "DEMURRAGE CHARGES",
        "DEMURRAGE AMT (INCL GST)",
        "Demurrage Amt (Incl GST)",
    ],

    "Total Amt (Incl GST)": [
        "TOTAL AMT",
        "TOTAL AMOUNT",
        "TOTAL",
        "GRAND TOTAL",
        "NET AMOUNT",
        "TOTAL AMT (INCL GST)",
        "TOTAL AMOUNT (INCL GST)",
        "Total Amt (Incl GST)",
    ],
}


AIRLINE_CODES = {
    "AY": "AY",
    "SQ": "SQ",
    "SZ": "SZ",
    "8M": "8M",
    "BZ": "BZ",
}


SHEET_KEYWORDS = [
    "ANNEXURE",
    "ANNEX",
    "ANX",
    "SUMMARY",
    "DATA_INSERT",
    "EXPORT TP DATA",
    "IMPORT TP DATA",
]


TOTAL_SHEET_KEYWORDS = [
    "SERVICE CHALLAN",
    "SERVICE_CHALLAN",
    "SC",
    "TOTAL SC",
    "TP_SC",
]


INVALID_ROW_KEYWORDS = [
    "TOTAL",
    "GRAND TOTAL",
    "SUB TOTAL",
    "SUBTOTAL",
]


AMOUNT_COLUMNS = [
    "Handling Amt (Incl GST)",
    "X-Ray Amt (Incl GST)",
    "Demurrage Amt (Incl GST)",
    "Total Amt (Incl GST)",
]


NUMERIC_COLUMNS = [
    "Pcs",
    "Gross Wt",
    "Chg Wt",
    "Handling Amt (Incl GST)",
    "X-Ray Amt (Incl GST)",
    "Demurrage Amt (Incl GST)",
    "Total Amt (Incl GST)",
]


GST_RATE = 0.18

HEADER_SCAN_LIMIT = 50

RECONCILIATION_TOLERANCE = 1.00

OUTPUT_FILE_NAME = "MIS_Consolidated_Output.xlsx"

LOG_FILE_NAME = "dcsc_mis_run.log"