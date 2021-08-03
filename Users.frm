VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Users 
   Caption         =   "User Dictionary"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Users.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

'------------------------------------------------------------------------------
' Procedure Name   : btnAddUser_Click
' Procedure Kind   : Sub
' Procedure Access : Private
' Author           : MSKIPWORTH
' Date             : 7/1/2021
' Purpose          : Adds a new record to tblUsers. Username and Key textbox
'                    fields must be non-null and the record must not already
'                    exist.
'------------------------------------------------------------------------------
Private Sub btnAddUser_Click()

    Dim lngRow As Long
    Dim ws As Worksheet
    Set ws = Sheets("tblUsers")
    

    If userExists(tbxUsername.Value) Then
        MsgBox "An entry already exists with this Username."
        Exit Sub
    End If
    
    lngRow = ws.Cells(Rows.Count, 1) _
        .End(xlUp).Offset(1, 0).Row
    
    If Trim(Me.tbxUsername.Value) = "" Then
        Me.tbxUsername.SetFocus
        MsgBox "Please enter a valid Username."
        Exit Sub
    ElseIf Trim(Me.tbxKey.Value) = "" Then
        Me.tbxKey.SetFocus
        MsgBox "Please enter a valid Key."
        Exit Sub
    Else
        ws.Cells(lngRow, USERNAME).Value = UCase(Me.tbxUsername.Value)
        ws.Cells(lngRow, KEY).Value = UCase(Me.tbxKey.Value)
        ws.Range(ws.Cells(1, KEY), ws.Cells(ws.UsedRange.Rows.Count, USERNAME)).Sort key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlNo
        
    End If
    
    MsgBox "New user added."
    Me.tbxUsername.Value = ""
    Me.tbxKey.Value = ""
    
End Sub



'------------------------------------------------------------------------------
' Procedure Name   : btnDelete_Click
' Procedure Kind   : Sub
' Procedure Access : Private
' Author           : MSKIPWORTH
' Date             : 7/1/2021
' Purpose          : Deletes the record active in the comboBox when Delete
'                    button clicked.
'------------------------------------------------------------------------------
Private Sub btnDelete_Click()
    Dim lngRow As Long
    Dim rngUserNames As Range
    Dim rngKeys As Range
    Dim intAreYouSure As Integer
    
    Set rngUserNames = Sheets("tblUsers").Columns(USERNAME)
    Set rngKeys = Sheets("tblUsers").Columns(KEY)
    
    
    If Not IsEmpty(cmbUsers.Value) Then
        intAreYouSure = MsgBox("Delete " & cmbUsers.Value & "?", vbYesNo)
        If intAreYouSure = vbNo Then
            MsgBox "Action cancelled."
            Exit Sub
        End If
        rngUserNames.Find(Mid(cmbUsers, InStr(cmbUsers, " : ") + Len(" : "))).EntireRow.Delete (xlShiftUp)
        Me.cmbUsers.Value = ""
        Call MultiPage1_Change
    Else:
    MsgBox "No Username selected!"
    End If
    
    
    
End Sub


'------------------------------------------------------------------------------
' Procedure Name   : close_Users
' Procedure Kind   : Sub
' Procedure Access : Private
' Author           : MSKIPWORTH
' Date             : 7/1/2021
' Purpose          : Closes the form.
'------------------------------------------------------------------------------
Private Sub close_Users()
    
    
    Unload Me
End Sub



'------------------------------------------------------------------------------
' Procedure Name   : MultiPage1_Change
' Procedure Kind   : Sub
' Procedure Access : Private
' Author           : MSKIPWORTH
' Date             : 7/1/2021
' Purpose          : Refreshes the list of ComboBox Items when the tabs are
'                    clicked.
'------------------------------------------------------------------------------
Private Sub MultiPage1_Change()


    Dim rngUserName As Range
    Dim rngKeys As Range
    Dim lngRow As Long
    Dim ws As Worksheet
    Dim lngCounter As Long
    lngCounter = 0
    
    Dim lngNumComboBoxEntries As Long
    
    Set ws = Sheets("tblUsers")

    
    lngRow = ws.Cells(Rows.Count, 1) _
        .End(xlUp).Offset(1, 0).Row
        
    If Me.cmbUsers.ListCount > 0 Then
        While Me.cmbUsers.ListCount > 0
            Me.cmbUsers.RemoveItem (Me.cmbUsers.ListCount - 1)
        Wend
    End If
     
    For Each rngUserName In ws.Range(ws.Cells(1, KEY), ws.Cells(lngRow, KEY)).Cells
        With Me.cmbUsers
            .AddItem rngUserName.Value & " : " & ws.Cells(rngUserName.Row, USERNAME)
            .List(.ListCount - 1, 1) = rngUserName.Offset(0, 1).Value
        End With
    Next rngUserName
End Sub

Private Sub Users_Activate()
    Call MultiPage1_Change
End Sub
