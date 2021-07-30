Attribute VB_Name = "modConstantBuilder"
Option Explicit



Dim nCols As Long

Sub getHeaders()
    On Error GoTo error_handler
    Dim headers() As String
    Dim lngCounter As Long
    
    Dim lngHeaderRow As Long
    Dim strSheetName As String

    
    lngHeaderRow = InputBox(prompt:="Which row are the table headers on?", Title:="Enter Header Row", Default:=1)
    
    strSheetName = InputBox(prompt:="On which Sheet?", Title:="Sheet Name?", Default:=ActiveSheet.Name)
    
    nCols = getNCols(Sheets(strSheetName), lngHeaderRow)
       
    
    headers = constantBuilder(Sheets(strSheetName), lngHeaderRow)
    
   
    createSheetIfNotExists ("CONSTANTS")
    With Sheets("CONSTANTS")
        For lngCounter = 1 To UBound(headers)
            .Cells(lngCounter, 1).Value = "Public Const " & UCase(Replace(headers(lngCounter - 1), " ", "_")) & " = " & lngCounter
        Next lngCounter
    End With
    
    MsgBox "done."
error_handler:
    MsgBox "ERROR OCCURED: " & Err.Description
    
End Sub

Function constantBuilder(ByVal ws As Worksheet, lngHeaderRow As Long) As String()
    
    
    Dim strAConstants() As String
    Dim lngCounter As Long
    Dim lngNumConstants As Long
    
    ReDim strAConstants(nCols - 1)
    
    With Sheet1

        For lngCounter = 1 To nCols
            strAConstants(lngCounter - 1) = .Cells(lngHeaderRow, lngCounter).Value
        Next lngCounter
    End With
    
    constantBuilder = strAConstants
End Function


Function getNCols(ws As Worksheet, lngHeaderRow As Long) As Long
    Dim lngCounter As Long
    Dim nCols As Long


    Dim cellMissCount As Integer
    Dim tooManyMissed As Integer
    
    
    cellMissCount = 0 'incremented when empty cell found. used to end early if tooManyMissed cells are empty. resets if a non-empty cell is found
    tooManyMissed = 3
    
    
    
    nCols = 0
    
    For lngCounter = 1 To 2 ^ 14
        If IsEmpty(ws.Cells(lngHeaderRow, lngCounter)) = False Then
            cellMissCount = 0
            nCols = nCols + 1
        Else
            If cellMissCount = tooManyMissed Then
                GoTo exitEarly
            Else
                cellMissCount = cellMissCount + 1
            End If
        End If
    Next lngCounter
    
    
exitEarly:
    getNCols = nCols
    
End Function
