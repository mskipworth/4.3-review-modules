VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sheet_Picker 
   Caption         =   "Sheet Picker"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Sheet_Picker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sheet_Picker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnCompare_Click()

    Dim shtWorkingSheet As Worksheet
    Dim tblUsers As Worksheet
    
    Set tblUsers = createSheetIfNotExists(TBL_USERS)
    
    If cmbSheets.Value <> "" Then
        Set shtWorkingSheet = Sheets(cmbSheets.Value)
    Else
        Debug.Print cmbSheets.Value
        MsgBox "Please select a valid Worksheet and try again."
        Exit Sub
    End If
    
    If shtWorkingSheet.Name <> TBL_USERS Then
        shtWorkingSheet.Activate
    
        Call parseBatchNames(shtWorkingSheet)
        tblUsers.Range(Sheets(TBL_USERS).Cells(1, KEY), Sheets(TBL_USERS).Cells(Sheets(TBL_USERS).UsedRange.Rows.Count, USERNAME)) _
        .Sort key1:=Sheets(TBL_USERS).Range("A1"), Order1:=xlAscending, Header:=xlNo
    End If
    
    MsgBox "Done!"
    Sheet_Picker.Hide
    
    
End Sub

Private Sub btnCompareAll_Click()

    Dim shtWorkingSheet As Worksheet
    For Each shtWorkingSheet In ActiveWorkbook.Sheets
        If shtWorkingSheet.Name <> TBL_USERS Then
            shtWorkingSheet.Activate
            Call parseBatchNames(shtWorkingSheet)
        End If
    Next shtWorkingSheet
    Sheet_Picker.Hide
End Sub

Private Sub UserForm_Initialize()
    
    cmbSheets.List = listAllSheets
    cmbSheets.Value = ActiveSheet.Name
End Sub
