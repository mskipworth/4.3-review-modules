Attribute VB_Name = "modGetNearest"
Option Explicit

Function scoreByToken(set1 As Object, set2 As Object) As Single

    Dim num As Single: num = 0#
    Dim den As Single: den = 0#
    Dim k As Object
    
    For Each k In set1.keys
        den = den + set1(k)
        If set2.exists(k) Then
            num = num + min(set1(k), set2(k)) * 2
        
        End If
    Next
    
    For Each k In set2.keys
        den = den + set2(k)
    Next
    
    scoreByToken = num / den
    
End Function

'
'Function getNearest(myStr As String) As Integer
'
'    'Dim scoreDict As New Scripting.Dictionary
'
'    Dim rng As Range: Set rng = Application.Range("data!A1:A" & rowLast(Sheets("data")))
'
'    Dim idealMatch As Double: idealMatch = 0#
'    Dim temp As Double
'    Dim matchStr As String, tempStr As String, tempStr2 As String
'    Dim matchIdx As Long: matchIdx = 0
'
'    tempStr = clearNonAlphas(myStr)
'
''    MsgBox rowLast(Sheets("data"))
''    MsgBox rng.Cells.Count
'
'    Dim cel As Range
'    For Each cel In rng
'        With cel
'            tempStr2 = clearNonAlphas(cel.Value)
'            'Debug.Print (tempStr & ", " & tempStr2)
'            temp = arrayCompare(getHash(tempStr), getHash(tempStr2)) ' to score by character frequency
'    '        temp = scoreByToken(getTokenFrequency(myStr), getTokenFrequency(cel.Value)) 'to score by token (word) frequency
'    '        If temp = 0 Then
'    '            temp = arrayCompare(getHash(myStr), getHash(cel.Value)) ' to score by character frequency if no token matches are found.
'    '        End If
'    '        Debug.Print (temp)
'            If temp > idealMatch Then
'                idealMatch = temp
'                matchStr = cel.Value
'                matchIdx = cel.Row
'                Debug.Print tempStr & "=" & cel.Row & ", " & cel.Value & ", SCORE: " & temp
'            End If
'        End With
'    Next
'    'MsgBox (matchStr & ", " & matchIdx)
'    'getNearest = matchStr
'    getNearest = matchIdx
'
'End Function


 Function getHash(myStr As String) As Integer()

    Dim letters(128) As Integer
    Dim idx As Integer
    
    
    myStr = UCase(myStr)
    
    For idx = 1 To Len(myStr)
        letters(Asc(Mid(myStr, idx, 1))) = letters(Asc(Mid(myStr, idx, 1))) + 1
    Next
    
    getHash = letters

End Function


