Attribute VB_Name = "modStringUtils"
Option Explicit


Function clearNonAlphas(myStr As String) As String
    
    myStr = Replace(myStr, "1", "")
    myStr = Replace(myStr, "2", "")
    myStr = Replace(myStr, "3", "")
    myStr = Replace(myStr, "4", "")
    myStr = Replace(myStr, "5", "")
    myStr = Replace(myStr, "6", "")
    myStr = Replace(myStr, "7", "")
    myStr = Replace(myStr, "8", "")
    myStr = Replace(myStr, "9", "")
    myStr = Replace(myStr, "0", "")
    myStr = Replace(myStr, "/", "")
    myStr = Replace(myStr, "-", " ")
    myStr = Replace(myStr, "_", " ")
    myStr = Replace(myStr, ")", "")
    myStr = Replace(myStr, "(", "")
    
    clearNonAlphas = Trim(myStr)
    
End Function

Function removeAbbrMonths(myStr As String) As String

    myStr = Replace(myStr, UCase("jan"), "")
    myStr = Replace(myStr, UCase("feb"), "")
    myStr = Replace(myStr, UCase("mar"), "")
    myStr = Replace(myStr, UCase("apr"), "")
    myStr = Replace(myStr, UCase("may"), "")
    myStr = Replace(myStr, UCase("jun"), "")
    myStr = Replace(myStr, UCase("jul"), "")
    myStr = Replace(myStr, UCase("aug"), "")
    myStr = Replace(myStr, UCase("sep"), "")
    myStr = Replace(myStr, UCase("oct"), "")
    myStr = Replace(myStr, UCase("nov"), "")
    myStr = Replace(myStr, UCase("dec"), "")
    
    removeAbbrMonths = Trim(myStr)
End Function

Function removeMonths(myStr As String) As String

    myStr = Replace(myStr, UCase("january"), "")
    myStr = Replace(myStr, UCase("february"), "")
    myStr = Replace(myStr, UCase("march"), "")
    myStr = Replace(myStr, UCase("april"), "")
    myStr = Replace(myStr, UCase("may"), "")
    myStr = Replace(myStr, UCase("june"), "")
    myStr = Replace(myStr, UCase("july"), "")
    myStr = Replace(myStr, UCase("august"), "")
    myStr = Replace(myStr, UCase("september"), "")
    myStr = Replace(myStr, UCase("october"), "")
    myStr = Replace(myStr, UCase("november"), "")
    myStr = Replace(myStr, UCase("december"), "")
    
    removeMonths = Trim(myStr)
End Function
