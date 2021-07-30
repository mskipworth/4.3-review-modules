Attribute VB_Name = "modUsersDictionary"
Option Explicit

Function userExists(strName As String) As Boolean
    Dim wsTblUsers As Worksheet
    Dim varResult As Variant
    Dim blnFunctionResult As Boolean
    
    If worksheetExists("tblUsers") = False Then
        MsgBox "tblUsers" & SHEET_DOES_NOT_EXIST
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "tblUsers"
    End If
    Set wsTblUsers = Sheets("tblUsers")
    Set varResult = wsTblUsers.Columns(USERNAME).Find(strName)
    
    If varResult Is Nothing Then
        blnFunctionResult = False
    Else
        blnFunctionResult = True
    End If
    
    userExists = blnFunctionResult
    
End Function

Function keyExists(strKey As String) As Boolean
    Dim wsTblUsers As Worksheet
    Dim varResult As Variant
    Dim blnFunctionResult As Boolean
    
    Set wsTblUsers = Sheets("tblUsers")
    Set varResult = wsTblUsers.Columns(KEY).Find(strKey)
    
    If varResult Is Nothing Then
        blnFunctionResult = False
    Else
        blnFunctionResult = True
    End If
    
    keyExists = blnFunctionResult
    
End Function

Function getUsername(strKey As String)
    If keyExists(strKey) = False Then
        MsgBox "Error in function call: getUsername(" & strKey & ")." _
            & vbNewLine & "Key not found."
        Exit Function
    End If
    
    getUsername = Sheets("tblUsers").Cells(Sheets("tblUsers").Columns(KEY).Find(strKey).Row(), USERNAME).Value
    
End Function
Function getKey(strUsername As String) As String
    If userExists(strUsername) = False Then
'        MsgBox "Error in function call: getUsername(" & strUsername & ")." _
'            & vbNewLine & "Key not found."
        Exit Function
    End If
    
    getKey = Sheets("tblUsers").Cells(Sheets("tblUsers").Columns(USERNAME).Find(strUsername).Row(), KEY).Value
    
End Function

'------------------------------------------------------------------------------
' Procedure Name   : addUser
' Procedure Kind   : Function
' Procedure Access : Public
' Parameter strKey (String): New record key.
' Parameter strUsername (String): New record username.
' Author           : MSKIPWORTH
' Date             : 7/1/2021
' Purpose          : Add a record with the supplied key and username.
'------------------------------------------------------------------------------
Function addUser(strKey As String, strUsername As String)
    
    Dim lngRow As Long
    Dim ws As Worksheet
    Set ws = Sheets(TBL_USERS)
    
    

    If userExists(strUsername) Then
        'MsgBox "An entry already exists with this Username."
        Exit Function
    End If
    
    lngRow = ws.Cells(Rows.Count, 1) _
        .End(xlUp).Offset(1, 0).Row
    
'    If Trim(Me.tbxUsername.Value) = "" Then
'        Me.tbxUsername.SetFocus
'        MsgBox "Please enter a valid Username."
'        Exit Sub
'    ElseIf Trim(Me.tbxKey.Value) = "" Then
'        Me.tbxKey.SetFocus
'        MsgBox "Please enter a valid Key."
'        Exit Sub
'    Else

        ws.Cells(lngRow, USERNAME).Value = strUsername
        ws.Cells(lngRow, KEY).Value = strKey
        
        
'    End If
    
'    MsgBox "New user added."
'    Me.tbxUsername.Value = ""
'    Me.tbxKey.Value = ""
    
End Function
