Attribute VB_Name = "modArrayUtils"
Option Explicit


Function displayArr(arr As Variant)
    
    Dim output As String
    Dim idx As Integer
    
    For idx = 0 To UBound(arr)
        If arr(idx) > 0 Then
            output = output & arr(idx)
            Else
            output = output & 0
            End If
        
    Next
        
    MsgBox output
End Function

 Function arrayCompare(arr1 As Variant, arr2 As Variant)
    
    Dim tally As Integer
    Dim baseline As Integer
    Dim idx As Integer
    tally = 0
    
    For idx = 0 To UBound(arr1)
        If arr1(idx) <> 0 Then
            tally = tally + arr1(idx)
        End If
    Next idx
    
    baseline = tally
    

    If UBound(arr1) <> UBound(arr2) Then
        MsgBox ("Array lengths not equal.")
        tally = 0
    Else
        For idx = 0 To UBound(arr1)
            If arr1(idx) = arr2(idx) Then
'                continue
                tally = tally
            ElseIf arr1(idx) > arr2(idx) Then
            
'                subtract from tally
                tally = tally - Abs(arr2(idx) - arr1(idx))
            End If
        Next
    End If
    If baseline = 0 Then
        baseline = 1
    End If
    
    arrayCompare = (tally / baseline)
            
End Function

Public Function GetArrLength(a As Variant) As Long
   If IsEmpty(a) Then
      GetArrLength = 0
   Else
      GetArrLength = UBound(a) - LBound(a) + 1
   End If
End Function

Public Function exists(myArr() As Variant, myStr As String) As Boolean

    Dim varElement As Variant

    For Each varElement In myArr
        If varElement = myStr Then
            exists = True
            Exit Function
        End If
    Next varElement
    
exists = False
End Function
