Function REVERSE(InString) As String

    Dim str_len, i As Integer

    REVERSE = ""
    str_len = Len(InString)
    
    For i = str_len To 1 Step -1
        REVERSE = REVERSE & Mid(InString, i, 1)
    Next i
    
End Function
