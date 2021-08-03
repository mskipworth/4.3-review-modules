Attribute VB_Name = "modGetUserNameFromBatchName"
Option Explicit

'------------------------------------------------------------------------------
' Procedure Name   : parseBatchNames
' Procedure Kind   : Function
' Procedure Access : Public
' Parameter sht (Worksheet): WorkSheet to operate on
' Author           : MSKIPWORTH
' Date             : 7/7/2021
' Purpose          : compares the batchname username with the created by and
'                    submitted by usernames.
'------------------------------------------------------------------------------
Function parseBatchNames(shtTarget As Worksheet) 'NO RETURN

    'declare variables
    Dim lngLastDataRowOfSheet As Long
    Dim lngCounter As Long
    Dim usrname As String
    Dim createdByCompareResult As String
    Dim submittedByCompareResult As String
    Dim cel As Range
    
    
    'find the last occupied row of the sheet using the source name column.
    lngLastDataRowOfSheet = rowLastInCol(shtTarget, SOURCE_NAME) + HEADER_ROW_OFFSET - 1
    
    
    'create new column headers
    shtTarget.Cells(HEADER_ROW_OFFSET, BATCH_NAME_VS_CREATED_BY).Value = BATCH_NAME_VS_CREATED_BY_HEADER
    shtTarget.Cells(HEADER_ROW_OFFSET, BATCH_NAME_VS_SUBMITTED_BY).Value = BATCH_NAME_VS_SUBMITTED_BY_HEADER
           
    
    
    For lngCounter = HEADER_ROW_OFFSET + 1 To lngLastDataRowOfSheet
        Set cel = shtTarget.Cells(lngCounter, APPROVAL_DESCRIPTION)
        
        If cel.Value <> APPROVAL_NOT_REQUIRED Then
            usrname = extractUserNameFromBatchName(Cells(cel.Row, BATCH_NAME).Value)
            
            If usrname = "" Or usrname = " " Then
                Cells(cel.Row, BATCH_NAME_VS_CREATED_BY).Value = ""
                Cells(cel.Row, BATCH_NAME_VS_SUBMITTED_BY).Value = ""
                GoTo continue
                
            End If
            
            createdByCompareResult = compareNames(usrname, Cells(cel.Row, BATCH_CREATED_BY))
            submittedByCompareResult = compareNames(usrname, Cells(cel.Row, BATCH_SUBMITTED_BY))
            
            Cells(cel.Row, BATCH_NAME_VS_CREATED_BY).Value = createdByCompareResult
            Cells(cel.Row, BATCH_NAME_VS_SUBMITTED_BY).Value = submittedByCompareResult
            
            
        Else
            Range(Cells(cel.Row, BATCH_NAME_VS_CREATED_BY), Cells(cel.Row, BATCH_NAME_VS_SUBMITTED_BY)).Value = "Approval not required."
            'Debug.Print "Approval not required." & " Row " & cel.Row
            
        End If
        
        
continue:
    Next lngCounter
    
    shtTarget.Columns(BATCH_NAME_VS_CREATED_BY).AutoFit
    shtTarget.Columns(BATCH_NAME_VS_SUBMITTED_BY).AutoFit
    

End Function


'------------------------------------------------------------------------------
' Procedure Name   : extractUserNameFromBatchName
' Procedure Kind   : Function
' Procedure Access : Public
' Parameter strBatchName (String): Batch name to be parsed.
' Return Type      : String
' Author           : MSKIPWORTH
' Date             : 7/7/2021
' Purpose          : parses the batch name and extracts the abbreviated username
'------------------------------------------------------------------------------
Function extractUserNameFromBatchName(strBatchName As String) As String
    Dim usrname As String
    Dim intSliceBegin As Integer ' index of first character in abbreviated username pulled from the batch name.
    Dim intSliceEnd As Integer ' index of last character in abbreviated username pulled from batch name.

    If InStr(strBatchName, ".") > 0 Then
        

' big list of rules.
        If InStr(strBatchName, " 9.") > 0 Then
            intSliceBegin = InStr(strBatchName, " 9.") + 3
            intSliceEnd = InStr(intSliceBegin + 1, strBatchName, ".")
            usrname = Mid(strBatchName, intSliceBegin, intSliceEnd - intSliceBegin) 'space before "9."
            Debug.Print "1"
            
        ElseIf (Left(strBatchName, 2) = "9." And InStr(InStr(strBatchName, "9.") + 2, strBatchName, ".") > 0) Then
            intSliceBegin = InStr(1, strBatchName, "9.") + 2
            intSliceEnd = InStr(intSliceBegin + 1, strBatchName, ".")
            usrname = Trim(Mid(strBatchName, intSliceBegin, intSliceEnd - intSliceBegin)) ' no space before "9."
            Debug.Print "2"
            
        ElseIf (Left(strBatchName, 2) = "9." And InStr(InStr(strBatchName, "9.") + 2, strBatchName, ".") = 0) Then
            intSliceBegin = InStr(1, strBatchName, "9.") + 2
            intSliceEnd = InStr(intSliceBegin + 1, strBatchName, " ")
            usrname = Trim(Mid(strBatchName, intSliceBegin, intSliceEnd - intSliceBegin)) '"9." detected but no "." found after.
            Debug.Print "3"
            
        ElseIf InStr(strBatchName, "Reverses ") Then

                If (InStr(strBatchName, ".") > 0 And InStr(strBatchName, " ") > 0 And _
                    (InStr(strBatchName, ".") < InStr(Len("Reverses  ") + 2, strBatchName, " "))) Then
                    
                    intSliceBegin = Len("Reverses ") + 2
                    intSliceEnd = InStr(intSliceBegin + 1, strBatchName, ".")
                    usrname = Mid(strBatchName, intSliceBegin, intSliceEnd - intSliceBegin)
                    Debug.Print "4"
                Else
                    intSliceBegin = Len("Reverses ") + 2
                    intSliceEnd = InStr(intSliceBegin + 1, strBatchName, " ")
                    usrname = Mid(strBatchName, intSliceBegin, intSliceEnd - intSliceBegin)
                    Debug.Print "5"
                End If
        ElseIf InStr(strBatchName, "REVERSE ") > 0 And InStr(strBatchName, "REVERSE ") < Len("REVERSE ") Then
            intSliceBegin = InStr(strBatchName, "REVERSE") + Len("REVERSE")
            intSliceEnd = InStr(intSliceBegin + 1, strBatchName, ".")
            usrname = Mid(strBatchName, intSliceBegin, intSliceEnd - intSliceBegin)
            Debug.Print "6"
        ElseIf InStr(strBatchName, "REVERSE.") > 0 And InStr(strBatchName, "REVERSE.") < Len("REVERSE.") Then
            intSliceBegin = InStr(strBatchName, "REVERSE.") + Len("REVERSE.")
            intSliceEnd = InStr(intSliceBegin + 1, strBatchName, ".")
            usrname = Mid(strBatchName, intSliceBegin, intSliceEnd - intSliceBegin)
            Debug.Print "7"
        Else
            If InStr(strBatchName, ".") > 0 Then
                usrname = Left(strBatchName, InStr(strBatchName, ".") - 1)
                Debug.Print "USERNAME=" & usrname
                Debug.Print "8"
            End If
        End If
    Else
        usrname = Left(strBatchName, InStr(strBatchName, " "))
        Debug.Print "9"
    End If
    
    usrname = UCase(clearNonAlphas(usrname))
    Debug.Print "USERNAME=" & usrname
    If exists(LIST_OF_NON_NAMES(), usrname) Then
        usrname = ""
        Debug.Print "10"
    End If
    
    Debug.Print usrname & " : " & strBatchName
    
    extractUserNameFromBatchName = usrname
End Function

'------------------------------------------------------------------------------
' Procedure Name   : compareNames
' Procedure Kind   : Function
' Procedure Access : Public
' Parameter strKey (String): 'Key' abbreviated username parsed from batch name
' Parameter strUsername (String): username entered in either created by or
'                    submitted by columns.
'
' Return Type      : String
' Author           : MSKIPWORTH
' Date             : 7/30/2021
' Purpose          : Compares abbreviated username in batch name with submitted
'                    by and created by columns by character frequency.
'                    compareNames("cat", "tackle") for instance results in
'                    "Match likely." because all characters in Key occur in
'                    strUsername.
'------------------------------------------------------------------------------
Function compareNames(strKey As String, strUsername As String) As String
    Dim dblScore As Double
    Dim result As String
    
    
    If ((strKey = "" Or strKey = " ") Or (strUsername = "" Or strUsername = " ")) Then
        'Debug.Print "1"
        compareNames = ""
        Exit Function
    End If
    
    If userExists(strUsername) = True And getKey(strUsername) = strKey Then ' ENTRY EXISTS IN TBL_USERS
        'Debug.Print "2"
        compareNames = MATCH_LIKELY
        Exit Function
        
    Else
        dblScore = arrayCompare(getHash(strKey), getHash(strUsername))
        
        If dblScore = 1 Then
            Call addUser(strKey, strUsername)
            result = MATCH_LIKELY
            
        Else
            result = MATCH_UNLIKELY
            
        End If
    End If
    
    compareNames = result
End Function




