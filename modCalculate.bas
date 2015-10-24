Attribute VB_Name = "modCalculate"
Public Type MaterialAtom
    AtomNumber() As Integer
    AtomMass() As Double
    AtomMassPer() As Double
    TotalMass As Double
    Material As String
End Type

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
    If IsNull(calElementChoose) = True Then calElementChoose = 0
End Function

Function calAsc(x As String) As Integer
    Select Case Asc(x)
        Case Asc("A") To Asc("Z")
            calAsc = 1
        Case Asc("a") To Asc("z")
            calAsc = 2
        Case Asc("0") To Asc("9")
            calAsc = 3
        Case 40         '英文括号
            calAsc = 4
        Case 41
            calAsc = 5
        Case 91         '英文方括号
            calAsc = 4
        Case 93
            calAsc = 5
        Case -23640     '中文括号
            calAsc = 4
        Case -23639
            calAsc = 5
        Case Else
            calAsc = 0
    End Select
End Function

Function calAtom(x As String) As MaterialAtom
    ReDim calAtom.AtomNumber(118) As Integer
    Dim AtomNumber(118) As Integer
    Dim i As Integer, l As Integer, y1 As String, y2 As String, y3 As String, y4 As String, t As String, n As Integer, s As Integer, i2 As Integer
    calAtom.Material = x
    l = Len(x)
    If l = 0 Then calAtom.AtomNumber(0) = 1 Else calAtom.AtomNumber(0) = 0
    Dim MulNumber(50) As Integer, MulIf(50) As Integer, MulLeft(50) As Integer, MulRight(50) As Integer, MulNum(50) As Integer
    i = 0
    s = 0
    While i < l
        i = i + 1
        MulNumber(i) = 1
        y1 = Mid(x, i, 1)
        If calAsc(y1) = 4 Then
            MulIf(i) = 1
        ElseIf calAsc(y1) = 5 Then
            MulIf(i) = -1
        Else
            MulIf(i) = 0
        End If
        s = s + MulIf(i)
    Wend
    If s <> 0 Then calAtom.AtomNumber(0) = 1
    i = 1
    n = 0
    While i < l And calAtom.AtomNumber(0) = 0
        If MulIf(i) = 1 Then
            n = n + 1
            c = 1
            i2 = i + 1
            MulLeft(n) = i
            While c > 0
                c = c + MulIf(i2)
                i2 = i2 + 1
            Wend
            i2 = i2 - 1
            MulRight(n) = i2
            If i2 + 1 > l Then y3 = "a" Else y3 = Mid(x, i2 + 1, 1)
            If calAsc(y3) = 3 Then
                If i2 + 2 > l Then y4 = "a" Else y4 = Mid(x, i2 + 2, 1)
                If calAsc(y4) = 3 Then
                    MulNum(n) = Val(y3 & y4)
                Else
                    MulNum(n) = Val(y3)
                End If
            Else
                MulNum(n) = 1
            End If
        End If
        i = i + 1
    Wend
    i = 0
    While i < n And calAtom.AtomNumber(0) = 0
        i = i + 1
        For i2 = MulLeft(i) To MulRight(i)
            MulNumber(i2) = MulNumber(i2) * MulNum(i)
        Next i2
    Wend
    i = 0
    While i < l And calAtom.AtomNumber(0) = 0 And calAtom.AtomNumber(0) = 0
        i = i + 1
        y1 = Mid(x, i, 1)
        If calAsc(y1) = 1 Then '首位为大写字母
            If i >= l Then y2 = "1" Else y2 = Mid(x, i + 1, 1)
            If calAsc(y2) = 2 Then '第2位为小写
                t = y1 & y2
                n = calElementChoose(t)
                If n = 0 Then
                    calAtom.AtomNumber(0) = 1
                Else
                    If i + 1 >= l Then y3 = "1" Else y3 = Mid(x, i + 2, 1)
                    If calAsc(y3) = 3 Then '第3位为数字
                        If i + 2 >= l Then y4 = "a" Else y4 = Mid(x, i + 3, 1)
                        If calAsc(y4) = 3 Then '第4位为数字
                            calAtom.AtomNumber(n) = calAtom.AtomNumber(n) + Val(y3 & y4) * MulNumber(i)
                            i = i + 3
                        Else
                            calAtom.AtomNumber(n) = calAtom.AtomNumber(n) + Val(y3) * MulNumber(i)
                            i = i + 2
                        End If
                    Else
                        calAtom.AtomNumber(n) = calAtom.AtomNumber(n) + MulNumber(i)
                        i = i + 1
                    End If
                End If
            ElseIf calAsc(y2) = 3 Then
                n = calElementChoose(y1)
                If n = 0 Then
                    calAtom.AtomNumber(0) = 1
                Else
                    If i + 1 >= l Then y3 = "a" Else y3 = Mid(x, i + 2, 1)
                    If calAsc(y3) = 3 Then
                        calAtom.AtomNumber(n) = calAtom.AtomNumber(n) + Val(y2 & y3) * MulNumber(i)
                        i = i + 2
                    Else
                        calAtom.AtomNumber(n) = calAtom.AtomNumber(n) + Val(y2) * MulNumber(i)
                        i = i + 1
                    End If
                End If
            Else
                n = calElementChoose(y1)
                If n = 0 Then
                    calAtom.AtomNumber(0) = 1
                Else
                    calAtom.AtomNumber(n) = calAtom.AtomNumber(n) + MulNumber(i)
                End If
            End If
        ElseIf calAsc(y1) = 4 Then
            
        ElseIf calAsc(y1) = 5 Then
            If i >= l Then y2 = "a" Else y2 = Mid(x, i + 1, 1)
            If calAsc(y2) = 3 Then
                If i + 1 >= l Then y2 = "a" Else y3 = Mid(x, i + 2, 1)
                If calAsc(y3) = 3 Then
                    i = i + 1
                End If
                i = i + 1
            End If
        Else
            calAtom.AtomNumber(0) = 1
        End If
    Wend
End Function

Function calMass(x As MaterialAtom) As MaterialAtom
    ReDim calMass.AtomMass(118) As Double
    ReDim calMass.AtomNumber(118) As Integer
    ReDim calMass.AtomMassPer(118) As Double
    calMass.Material = x.Material
    Dim i As Integer, m As Double
    If x.AtomNumber(0) = 1 Then
        m = -1
    Else
        m = 0
        For i = 1 To 118
            m = m + x.AtomNumber(i) * ElementMass(i)
        Next i
        If m > 0 Then
            For i = 1 To 118
                calMass.AtomNumber(i) = x.AtomNumber(i)
                calMass.AtomMass(i) = x.AtomNumber(i) * ElementMass(i)
                calMass.AtomMassPer(i) = calMass.AtomMass(i) / m
            Next i
        End If
    End If
    calMass.TotalMass = m
End Function

Function calTotalMassStr(x As String) As Double
    calTotalMassStr = calMass(calAtom(x)).TotalMass
End Function

Function calMassStr(x As String) As MaterialAtom
    calMassStr = calMass(calAtom(x))
End Function

Function calMassPer(x As MaterialAtom) As String
    Dim i As Integer
    If x.TotalMass = -1 Then
        calMassPer = "您的输入有误，请重新输入！" & Chr(13) & Chr(10) & "请检查：" & Chr(13) & Chr(10) & "1.元素符号是否正确（区分大小写）；" & Chr(13) & Chr(10) & "2.括号是否缺少。"
    Else
        calMassPer = x.Material & "的" & "分子量为" & x.TotalMass & "，其中：" & Chr(13) & Chr(10)
        For i = 1 To 118
            If x.AtomNumber(i) > 0 Then
                calMassPer = calMassPer & ElementName(i) & "（符号：" & ElementAbbr(i) & "），" & x.AtomNumber(i) & "个原子，原子量为" & Format(ElementMass(i), "0.00") & "，质量分数为" & Format(x.AtomMassPer(i), "0.00%") & "；" & Chr(13) & Chr(10)
            End If
        Next i
    End If
End Function

Function calElementStr(n As Integer) As String
    If n > 0 Then
        calElementStr = ElementName(n) & Chr(13) & Chr(10) & "原子序数：" & n & Chr(13) & Chr(10) & "元素符号：" & ElementAbbr(n) & Chr(13) & Chr(10) & "相对原子质量：" & ElementMass(n)
    Else
        calElementStr = "输入错误！" & Chr(13) & Chr(10) & "请检查您的输入是否有误！"
    End If
End Function

Function calMassPerStr(x As String) As String
    calMassPerStr = calMassPer(calMass(calAtom(x)))
End Function


'Function calGas()
'待完成
'End Function

'Function calEquation(x As String)
'完成中
   ' Dim l As Integer, i As Integer
    
'End Function
