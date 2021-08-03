Attribute VB_Name = "modConstants"
Option Explicit


Public Const HEADER_ROW_OFFSET = 9

Public Const EXCEL_MAX_ROWS = 2 ^ 20
Public Const EXCEL_MAX_COLS = 2 ^ 14

'REPORT HEADERS
Public Const LEDGER = 1
Public Const BATCH_NAME = 2
Public Const BATCH_CREATED_BY = 3
Public Const BATCH_SUBMITTED_BY = 4
'Public Const CREATED_EQUALS_SUBMITTED = 5
Public Const BATCH_POSTED_BY = 5
Public Const SOURCE_NAME = 6
Public Const DATETIME_POSTED = 7
Public Const APPROVAL_STATUS = 8
Public Const APPROVAL_DESCRIPTION = 9
Public Const APPROVER_ACTION_DATES = 10
Public Const ADI_UPLOAD_ID = 11
Public Const NOTES = 12
Public Const BATCH_NAME_VS_CREATED_BY = 14
Public Const BATCH_NAME_VS_SUBMITTED_BY = 15

'TBLUSERS HEADERS
Public Const USERNAME = 1
Public Const KEY = 2


Public Const TBL_USERS = "tblUsers"
Public Const APPROVAL_NOT_REQUIRED = "Your journal batch does not require approval."
Public Const SHEET_DOES_NOT_EXIST = " does not exist!" & vbNewLine & "Creating it now..."
Public Const MATCH_LIKELY = "Match likely."
'Public Const MATCH_POSSIBLE = "Match possible." 'NOT USED AT THIS TIME.
Public Const MATCH_UNLIKELY = "Discrepancy Found."
Public Const BATCH_NAME_VS_SUBMITTED_BY_HEADER = "Batch Name vs Submitted By"
Public Const BATCH_NAME_VS_CREATED_BY_HEADER = "Batch Name vs Created By"


Public Function LIST_OF_NON_NAMES() As Variant()
    
    Dim strANonNames() As Variant
    strANonNames = Array("DVAAPNP", "DVAAPCP", "RVS", "APEX", "US", "VALUATION", _
        "UTILITIES", "INTANGIBLE", "PAYABLES", "BOOK", "PA", "BANK", "VIRTUAL", _
        "PREMISE", "MIRA", "REVERSE", "REVERSES", "IVY", "DAX")
    
    LIST_OF_NON_NAMES = strANonNames
End Function
