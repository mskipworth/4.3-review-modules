Attribute VB_Name = "modSheetUtils"
Option Explicit



Public Function createSheetIfNotExists(sheetName As String) As Worksheet

    'if the worksheet does not yet exist create it, else continue.
    If worksheetExists(sheetName) <> True Then
        
        MsgBox (sheetName & SHEET_DOES_NOT_EXIST)
            
        ActiveWorkbook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName
        
    Else
        ActiveWorkbook.Sheets(sheetName).Activate
        
    End If
        Set createSheetIfNotExists = ActiveSheet
End Function



Public Function worksheetExists(worksheetName As String) As Boolean
    
    Dim shtCurrent As Worksheet
    Dim blnResult As Boolean
    
    blnResult = False
    
    For Each shtCurrent In ThisWorkbook.Worksheets
        If shtCurrent.Name = worksheetName Then
            blnResult = True
            GoTo exit_early
        End If
    Next
    
    
exit_early:
    worksheetExists = blnResult
End Function




Public Function rowLast(sht As Worksheet)
'    rowLast = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
'    rowLast = Worksheets(sht.Name).Cells(Rows.Count, 1).End(xlUp).Row
    rowLast = sht.UsedRange.Rows.Count
    
    
    End Function



Public Function colLast(sht As Worksheet)
    colLast = sht.UsedRange.Columns.Count
End Function

Public Function rowLastInCol(sht As Worksheet, colIndex As Integer) As Long

    Dim celA1 As Range
    Dim lngResult  As Long
    Dim lngCount As Long
    
    
    lngResult = 0
    
    Set celA1 = sht.Range("A1")
    
    For lngCount = 1 To EXCEL_MAX_ROWS
        With sht
            If IsEmpty(.Cells(lngCount, colIndex)) = False Then
                lngResult = lngResult + 1
            End If
        End With
    Next
    
    rowLastInCol = lngResult

End Function



Function max(x As Double, y As Double)

    Dim result As Double
    
    If x > y Then
        result = x
    Else
        result = y
    End If
    
    max = result

End Function



Function min(x As Double, y As Double)

    Dim result As Double
    
    If x < y Then
        result = x
    Else
        result = y
    End If
    
    min = result
End Function

Function listAllSheets() As String()
    
    Dim sht As Worksheet
    Dim nSheets As Integer
    Dim strSheetNames() As String
    nSheets = 1
    
    
    
    For Each sht In ThisWorkbook.Worksheets
        ReDim Preserve strSheetNames(nSheets)
        strSheetNames(nSheets - 1) = sht.Name
        nSheets = nSheets + 1
    Next sht
    
    listAllSheets = strSheetNames
End Function


