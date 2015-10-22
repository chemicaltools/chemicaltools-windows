Attribute VB_Name = "modCalculate"
Function calElementChoose(x As String) As Integer
    Dim i As Integer, t As Boolean
    i = 0
    t = False
    While i < 118 And t = False
        i = i + 1
        If i = Int(Val(x)) Then
            calElementChoose = i
            t = True
        ElseIf ElementName(i) = x Then
            calElementChoose = i
            t = True
        ElseIf UCase(ElementAbbr(i)) = UCase(x) Then
            calElementChoose = i
            t = True
        Else
            t = False
        End If
    Wend
    If IsNull(calElementChoose) = True Then calelmentchoose = 0
End Function
