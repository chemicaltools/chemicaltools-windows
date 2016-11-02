Attribute VB_Name = "modCalculate"
Public Type MaterialAtom
    AtomNumber() As Integer
    AtomMass() As Double
    AtomMassPer() As Double
    TotalMass As Double
    Material As String
End Type

Function calElementChoose(x As String) As Integer
    Dim i As Integer, T As Boolean
    i = 0
    T = False
    While i < 118 And T = False
        i = i + 1
        If i = Int(Val(x)) Then
            calElementChoose = i
            T = True
        ElseIf ElementName(i) = x Then
            calElementChoose = i
            T = True
        ElseIf UCase(ElementAbbr(i)) = UCase(x) Then
            calElementChoose = i
            T = True
        Else
            T = False
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
    Dim i As Integer, l As Integer, y1 As String, y2 As String, y3 As String, y4 As String, T As String, n As Integer, s As Integer, i2 As Integer
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
    While i < l And calAtom.AtomNumber(0) = 0
        i = i + 1
        y1 = Mid(x, i, 1)
        If calAsc(y1) = 1 Then '首位为大写字母
            If i >= l Then y2 = "1" Else y2 = Mid(x, i + 1, 1)
            If calAsc(y2) = 2 Then '第2位为小写
                T = y1 & y2
                n = calElementChoose(T)
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

Function calpH(pKa() As Double, c As Double, pKw As Double) As Double
    Dim cH As Double, Ka1 As Double, Kw As Double
    Ka1 = 10 ^ (-pKa(0))
    Kw = 10 ^ (-pKw)
    cH = (Sqr(Ka1 ^ 2 + 4 * Ka1 * c + Kw) - Ka1) * 0.5
    If cH > 0 Then calpH = -Log(cH) / Log(10) Else calpH = 1024
End Function

Function calpHtoc(pKa() As Double, c As Double, pH As Double) As Double()
    ReDim calpHtoc(UBound(pKa) + 1)
    Dim D As Double, E As Double, F As Double, G() As Double, H As Double, Ka() As Double, pHtoc() As Double
    ReDim Ka(UBound(pKa) + 1)
    ReDim G(UBound(pKa) + 1)
    ReDim pHtoc(UBound(pKa) + 1)
    n = UBound(pKa) + 1
    D = 0
    E = 1
    H = 10 ^ (-pH)
    F = H ^ n
    For i = LBound(pKa) To UBound(pKa)
        Ka(i) = 10 ^ (-pKa(i))
    Next i
    For i = LBound(pKa) To UBound(pKa) + 1
        G(i) = F * E
        D = D + G(i)
        F = F / H
        E = E * Ka(i)
    Next i
    For i = LBound(pKa) To UBound(pKa) + 1
        pHtoc(i) = c * G(i) / D
    Next i
    calpHtoc = pHtoc
End Function

Function calpHOut(pKa As String, c As String, pKw As String, AorB As Boolean) As String
    Dim strpKa() As String
    Dim valpKa() As Double
    Dim cAB() As Double
    Dim i As Integer, j As Integer
    Dim error As Boolean
    Dim n As Integer
    error = False
    If Val(c) = 0 Or Not IsNumeric(pKw) Then error = True
    calpHOut = "c=" & c & ", "
    If pKa = "" Then pKa = "error"
    strpKa() = Split(pKa)
    ReDim valpKa(UBound(strpKa))
    For i = LBound(strpKa) To UBound(strpKa)
        If Not IsNumeric(strpKa(i)) Then error = True
        valpKa(i) = Val(strpKa(i))
        If AorB Then calpHOut = calpHOut & "pKa" Else calpHOut = calpHOut & "pKb"
        calpHOut = calpHOut & i + 1 & "=" & strpKa(i) & ", "
    Next i
    calpHOut = calpHOut & Chr(13) & Chr(10)
    pH = calpH(valpKa(), Val(c), Val(pKw))
    cAB = calpHtoc(valpKa(), Val(c), Val(pH))
    If Not AorB Then pH = pKw - pH
    H = 10 ^ (-pH)
    calpHOut = calpHOut & "溶液的pH为" & Format(pH, "0.00") & Chr(13) & Chr(10) & "c(H+)=" & Format(H, "Scientific") & Chr(13) & Chr(10)
    For i = LBound(cAB) To UBound(cAB)
        calpHOut = calpHOut & "c("
        If AorB Then
            If i < UBound(cAB) Then
                calpHOut = calpHOut & "H"
                If UBound(cAB) - i > 1 Then calpHOut = calpHOut & UBound(cAB) - i
            End If
            calpHOut = calpHOut & "A"
            If i > 0 Then
                If i > 1 Then calpHOut = calpHOut & i
                calpHOut = calpHOut & "-"
            End If
        Else
            calpHOut = calpHOut & "B"
            If UBound(cAB) - i > 1 Then
                calpHOut = calpHOut & "(OH)" & UBound(cAB) - i
            ElseIf UBound(cAB) - i = 1 Then
                calpHOut = calpHOut & "OH"
            End If
            If i > 0 Then
                If i > 1 Then calpHOut = calpHOut & i
                calpHOut = calpHOut & "+"
            End If
        End If
        calpHOut = calpHOut & ")=" & Format(cAB(i), "Scientific") & Chr(13) & Chr(10)
    Next i
    If error = True Then calpHOut = "输入错误，请重新输入！"
End Function

Function calRelixue(H1 As String, H2 As String, S1 As String, S2 As String) As String
    Dim strH1() As String, strH2() As String, strS1() As String, strS2() As String
    Dim sumH1 As Double, sumH2 As Double, sumS1 As Double, sumS2 As Double
    Dim s As Double
    Dim detH As Double, detS As Double, detG As Double, T As Double, K As Double
    If H1 = "" Then H1 = "0"
    If H2 = "" Then H2 = "0"
    If S1 = "" Then S1 = "0"
    If S2 = "" Then S2 = "0"
    strH1() = Split(H1)
    strH2() = Split(H2)
    strS1() = Split(S1)
    strS2() = Split(S2)
    s = 0
    For i = LBound(strH1) To UBound(strH1)
        s = s + Val(strH1(i))
    Next i
    sumH1 = s
    s = 0
    For i = LBound(strH2) To UBound(strH2)
        s = s + Val(strH2(i))
    Next i
    sumH2 = s
    s = 0
    For i = LBound(strS1) To UBound(strS1)
        s = s + Val(strS1(i))
    Next i
    sumS1 = s
    s = 0
    For i = LBound(strS2) To UBound(strS2)
        s = s + Val(strS2(i))
    Next i
    sumS2 = s
    calRelixue = "反应物的总生成焓为" & Format(sumH1, "0.0") & "kJ/mol，生成物的总生成焓为" & Format(sumH2, "0.0") & "kJ/mol，反应物的总标准熵为" & Format(sumS1, "0.0") & "J/mol，生成物的总标准熵为" & Format(sumS2, "0.0") & "J/mol。" & Chr(13) & Chr(10)
    detH = sumH2 - sumH1
    detS = sumS2 - sumS1
    detG = detH - 298.15 * detS / 1000
    K = Exp(-detG * 1000 / R / 298.15)
    calRelixue = calRelixue & "反应的标准摩尔焓变为" & Format(detH, "0.0") & "kJ/mol，" & "标准摩尔熵变为" & Format(detS, "0.0") & "J/mol" & "，标准摩尔吉布斯自由能为" & Format(detG, "0.0") & "kJ/mol，标准平衡常数为" & Format(K, "Scientific") & "。" & Chr(13) & Chr(10)
    If detH >= 0 Then
        If detS >= 0 Then
            T = detH / detS * 1000
            calRelixue = calRelixue & "温度T<" & Format(T, "0.0") & "K时，该反应能自发进行，" & "温度T>" & Format(T, "0.0") & "K时，该反应不能自发进行。"
        Else
            calRelixue = calRelixue & "在任何温度下，该反应均不能自发进行。"
        End If
    Else
        If detS >= 0 Then
            calRelixue = calRelixue & "在任何温度下，该反应均能自发进行。"
        Else
            T = detH / detS * 1000
            calRelixue = calRelixue & "温度T>" & Format(T, "0.0") & "K时，该反应能自发进行，" & "温度T<" & Format(T, "0.0") & "K时，该反应不能自发进行。"
        End If
    End If
End Function

Function calGasp(v As Double, n As Double, T As Double)
    calGasp = n * R * T / v
End Function

Function calGasV(p As Double, n As Double, T As Double)
    calGasV = n * R * T / p
End Function

Function calGasn(p As Double, v As Double, T As Double)
    calGasn = p * v / R / T
End Function

Function calGasT(p As Double, v As Double, n As Double)
    calGasT = p * v / n / R
End Function
