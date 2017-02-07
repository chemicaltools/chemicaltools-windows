Attribute VB_Name = "modData"
'配置文件
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'数据库
Public DataAdodbConn As ADODB.Connection
Public DataAdodbRs As ADODB.Recordset
'用户
Public DataUsername As String
Public DataUseNumber As Integer
Public DataQQname As String
Public token As String
Public LoginUsername As String
Public LoginPassword As String
Public LoginSavingPassword As Integer
Public LoginAutoLogin As Integer
Public objectID As String
'配置
Public ExamTimeMax As Integer
Public ExamNumberMax As Integer
Public ExamNoMax As Integer
Public ExamScoreMax As Integer
Public ExamScoreMaxAll As Integer
Public ExamScoreNameAll As String
Public ExamTimeIf As Boolean
Public ExamScoreTime As String
Public ExamScoreTimeAll As String
'历史记录
Public HisElement As String
Public HisMass As String
Public HisUsername As String
Public HisPassword As String
Public HisAutoLogin As String
Public Hisc As String
Public HispKa As String
Public HispKw As String
Public HisAB As Boolean
Public HisElementOutput As String
Public HisMassOutput As String
Public HisAcidOutput As String
'元素
Public ElementName(118) As String
Public ElementAbbr(118) As String
Public ElementMass(118) As Double
Public ElementIUPAC(118) As String
Public ElementOrigin(118) As String
Public xmlhttp As Object
Public htmlStr As String

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function dataOpen(x As Integer)
    Dim path As String
    Set DataAdodbConn = New ADODB.Connection
    path = dataBasePath(x)
    DataAdodbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Persist Security Info=False"
    DataAdodbConn.Open
    Set DataAdodbRs = New ADODB.Recordset
    DataAdodbRs.ActiveConnection = DataAdodbConn
    DataAdodbRs.CursorType = adOpenDynamic
    DataAdodbRs.LockType = adLockOptimistic
End Function

Public Function dataClose()
   Set DataAdodbRs = Nothing
   DataAdodbConn.Close
   Set DataAdodbConn = Nothing
End Function

Public Function dataBasePath(x As Integer) As String
    Dim spath As String
    If Right$(App.path, 1) = "\" Then
        spath = App.path + "mdb\"
    Else
        spath = App.path + "\mdb\"
    End If
    If x = 0 Then dataBasePath = spath Else If x = 1 Then dataBasePath = spath + "Data.mdb" Else dataBasePath = spath + "User.mdb"
End Function

Public Function dataSettingPath() As String
    Dim spath As String
    If Right$(App.path, 1) = "\" Then
        spath = App.path
    Else
        spath = App.path + "\"
    End If
    dataSettingPath = spath + "config.ini"
End Function

Public Function dataElement()
    Dim i As Integer, n As Integer
    i = 0
    Call dataOpen(1)
    DataAdodbRs.Open "select * from Element"
    While Not DataAdodbRs.EOF
        i = i + 1
        If Not IsNull(DataAdodbRs!ElementNumber) Then n = CStr(DataAdodbRs!ElementNumber)
        If Not IsNull(DataAdodbRs!ElementName) Then ElementName(n) = CStr(DataAdodbRs!ElementName)
        If Not IsNull(DataAdodbRs!ElementAbbr) Then ElementAbbr(n) = CStr(DataAdodbRs!ElementAbbr)
        If Not IsNull(DataAdodbRs!ElementMass) Then ElementMass(n) = CStr(DataAdodbRs!ElementMass)
        If Not IsNull(DataAdodbRs!Origin) Then ElementOrigin(n) = CStr(DataAdodbRs!Origin)
        If Not IsNull(DataAdodbRs!iupac) Then ElementIUPAC(n) = CStr(DataAdodbRs!iupac)
        DataAdodbRs.MoveNext
        Wend
    Call dataClose
    i = 0
    While i < 118
        i = i + 1
        If IsNull(ElementName(i)) Then ElementName(i) = "None"
        If IsNull(ElementAbbr(i)) Then ElementAbbr(i) = "None"
        If IsNull(ElementMass(i)) Then ElementMass(i) = 0
    Wend
End Function

Public Function dataSettingWrite(x As String, Y As String, z As String)
    Call WritePrivateProfileString(x, Y, z, dataSettingPath)
End Function

Public Function dataBaseWrite(Username As String, Name As String, Value As String)
        Call dataOpen(2)
        DataAdodbRs.Open "select * from [User] where Username='" & Username & "'"
           DataAdodbRs(Name) = Value
        DataAdodbRs.Update
        Call dataClose
End Function

Public Function dataSettingRead(x As String, Y As String) As String
    Dim z As String, p As Long, i As Integer
    z = String(255, vbNullChar)
    p = GetPrivateProfileString(x, Y, "", z, 255, dataSettingPath)
    i = 1
    While Mid(z, i, 1) <> vbNullChar
        i = i + 1
    Wend
    z = Mid(z, 1, i - 1)
    dataSettingRead = z
End Function

Public Function dataDir(x As String) As Boolean
    If Dir(x, vbDirectory) = "" Then
        dataDir = False
    Else
        dataDir = True
    End If
End Function

Public Function dataBaseDir()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim spath As String
    If Right$(App.path, 1) = "\" Then
        spath = App.path + "mdb"
    Else
        spath = App.path + "\mdb"
    End If
    If Not dataDir(spath) Then fso.CreateFolder (spath)
    If dataDir(dataBasePath(1)) = False Or dataDir(dataBasePath(2)) = False Then
        If dataDir(dataBasePath(1)) = False Then
            ret = URLDownloadToFile(0, "http://test-10061032.cos.myqcloud.com/Data.mdb", dataBasePath(1), 0, 0)
            'MsgBox ret
            If ret <> 0 Then
                MsgBox "数据库文件缺失，请连接网络，程序将自动下载！"
                End
            End If
        End If
        If dataDir(dataBasePath(2)) = False Then
            ret = URLDownloadToFile(0, "http://test-10061032.cos.myqcloud.com/User.mdb", dataBasePath(2), 0, 0)
            'MsgBox ret
            If ret <> 0 Then
                MsgBox "数据库文件缺失，请连接网络，程序将自动下载！"
                End
            End If
        End If
    End If
End Function

Public Function dataSettingDir()
    If dataDir(dataSettingPath) = False Then
        Call dataSettingWrite("Exam", "ScoreMax", "0")
        Call dataSettingWrite("Exam", "ScoreName", "N/A")
        Call dataSettingWrite("Exam", "ScoreTime", "N/A")
        Call dataSettingWrite("History", "Username", "")
        Call dataSettingWrite("History", "Password", "")
        Call dataSettingWrite("History", "AutoLogin", "False")
    End If
End Function

Public Function dataSettingSave(Username As String)
    Call dataOpen(2)
    DataAdodbRs.Open "select * from [User] where Username='" & Username & "'"
    DataAdodbRs("TimeMax") = ExamTimeMax
    DataAdodbRs("NumberMax") = ExamNumberMax
    DataAdodbRs("NoMax") = ExamNoMax
    If ExamTimeIf = True Then DataAdodbRs("TimeIf") = "True" Else DataAdodbRs("TimeIf") = "False"
    DataAdodbRs.Update
    Call dataClose
End Function

Public Function dataScoreSave(Username As String)
    Call dataOpen(2)
    DataAdodbRs.Open "select * from [User] where Username='" & Username & "'"
    DataAdodbRs("ScoreMax") = ExamScoreMax
    DataAdodbRs("ScoreTime") = ExamScoreTime
    If ExamTimeIf = True Then DataAdodbRs("TimeIf") = "True" Else DataAdodbRs("TimeIf") = "False"
    DataAdodbRs.Update
    Call dataClose
    If ExamScoreMax > ExamScoreMaxAll Then
        ExamScoreMaxAll = ExamScoreMax
        ExamScoreNameAll = Username
        ExamScoreTimeAll = ExamScoreTime
        Call dataSettingWrite("Exam", "ScoreMax", Trim(str(ExamScoreMaxAll)))
        Call dataSettingWrite("Exam", "ScoreName", ExamScoreNameAll)
        Call dataSettingWrite("Exam", "ScoreTime", ExamScoreTimeAll)
    End If
End Function

Public Function dataHistoryRead()
    If dataDir(dataSettingPath) = True Then
        ExamScoreMaxAll = Int(Val(dataSettingRead("Exam", "ScoreMax")))
        ExamScoreNameAll = dataSettingRead("Exam", "ScoreName")
        ExamScoreTimeAll = dataSettingRead("Exam", "ScoreTime")
        HisUsername = dataSettingRead("History", "Username")
        HisPassword = dataSettingRead("History", "Password")
        HisAutoLogin = dataSettingRead("History", "AutoLogin")
    End If
End Function

Public Function dataLogin(Username As String, Password As String, SavingPassword As Integer, AutoLogin As Integer, json As String) As Boolean
    dataLogin = False
    'json = dataHtmlLogin(Username, Password, SavingPassword, AutoLogin)
    If JSONParse("sessionToken", json) <> "0" Then
        token = JSONParse("sessionToken", json)
        objectID = JSONParse("objectId", json)
        Call dataOpen(2)
        DataAdodbRs.Open "select * from [User]"
        While Not DataAdodbRs.EOF And dataLogin = False
            If Not IsNull(DataAdodbRs!Username) Then
                If Username = CStr(DataAdodbRs!Username) Then
                    dataLogin = True
                Else
                    DataAdodbRs.MoveNext
                End If
            End If
        Wend
        If dataLogin = False Then
            DataAdodbRs.AddNew
            DataAdodbRs("Username") = Username
            DataAdodbRs("UseNumber") = 0
            DataAdodbRs("TimeMax") = 60
            DataAdodbRs("NumberMax") = 100
            DataAdodbRs("NoMax") = 20
            DataAdodbRs("ScoreMax") = 0
            DataAdodbRs("ScoreTime") = "N/A"
            DataAdodbRs("TimeIf") = "True"
            DataAdodbRs("Element") = ""
            DataAdodbRs("Mass") = ""
            DataAdodbRs("c") = ""
            DataAdodbRs("pKw") = "14"
            DataAdodbRs("pKa") = ""
            DataAdodbRs("AB") = "A"
            DataAdodbRs("qqname") = ""
            DataAdodbRs("CorrectNumber") = 0
            DataAdodbRs("IncorrectNumber") = 0
            dataLogin = True
        End If
        If Not (JSONParse("elementnumber_limit", json) = "") Then
            DataAdodbRs("NumberMax") = Val(JSONParse("elementnumber_limit", json))
        End If
        If Not (JSONParse("historyElement", json) = "") Then
            DataAdodbRs("Element") = Val(JSONParse("historyElementNumber", json))
        End If
        If Not ((JSONParse("historyMass", json) = "")) Then
            DataAdodbRs("Mass") = JSONParse("historyMass", json)
        End If
        If Not ((JSONParse("pKw", json) = "")) Then
            DataAdodbRs("pKw") = Val(JSONParse("pKw", json))
        End If
        If Not ((JSONParse("qqname", json) = "")) Then
            DataAdodbRs("qqname") = JSONParse("qqname", json)
        End If
        If Not ((JSONParse("examCorrectNumber", json) = "")) Then
            DataAdodbRs("CorrectNumber") = Val(JSONParse("examCorrectNumber", json))
        End If
        If Not ((JSONParse("examIncorrectnumber", json) = "")) Then
            DataAdodbRs("Incorrectnumber") = Val(JSONParse("examIncorrectnumber", json))
        End If
        If Not ((JSONParse("historyElementOutput", json) = "")) Then
            DataAdodbRs("ElementOutput") = CutHtml(JSONParse("historyElementOutput", json))
        End If
        If Not ((JSONParse("historyMassOutput", json) = "")) Then
            DataAdodbRs("MassOutput") = CutHtml(JSONParse("historyMassOutput", json))
        End If
        If Not ((JSONParse("historyAcidOutput", json) = "")) Then
            DataAdodbRs("AcidOutput") = CutHtml(JSONParse("historyAcidOutput", json))
        End If
        DataUseNumber = CStr(DataAdodbRs!UseNumber)
        DataUseNumber = DataUseNumber + 1
        DataAdodbRs("UseNumber") = DataUseNumber
        DataAdodbRs.Update
        DataUsername = Username
        If Username <> "访客" Then
            HisUsername = Username
            Call dataSettingWrite("History", "Username", Username)
            If SavingPassword = 1 Then
                HisPassword = Password
                Call dataSettingWrite("History", "Password", Password)
            Else
                HisPassword = ""
                Call dataSettingWrite("History", "Password", "")
            End If
            If AutoLogin = 1 Then
                HisAutoLogin = "True"
                Call dataSettingWrite("History", "AutoLogin", "True")
            Else
                HisAutoLogin = "False"
                Call dataSettingWrite("History", "AutoLogin", "False")
            End If
        End If
        ExamTimeMax = CStr(DataAdodbRs!TimeMax)
        ExamNumberMax = CStr(DataAdodbRs!NumberMax)
        ExamNoMax = CStr(DataAdodbRs!NoMax)
        ExamScoreMax = CStr(DataAdodbRs!ScoreMax)
        ExamScoreTime = CStr(DataAdodbRs!ScoreTime)
        examCorrectNumber = CStr(DataAdodbRs!CorrectNumber)
        examIncorrectNumber = CStr(DataAdodbRs!IncorrectNumber)
        HisElement = CStr(DataAdodbRs!Element)
        HisMass = CStr(DataAdodbRs!Mass)
        Hisc = CStr(DataAdodbRs!c)
        HispKw = CStr(DataAdodbRs!pKw)
        HispKa = CStr(DataAdodbRs!pKa)
        DataQQname = CStr(DataAdodbRs!qqname)
        HisElementOutput = CStr(DataAdodbRs!ElementOutput)
        HisMassOutput = CStr(DataAdodbRs!massoutput)
        HisAcidOutput = CStr(DataAdodbRs!AcidOutput)
        If CStr(DataAdodbRs!AB) = "A" Then HisAB = True Else HisAB = False
        If CStr(DataAdodbRs!TimeIf) = "True" Then ExamTimeIf = True Else ExamTimeIf = False
        Call dataClose
    End If
End Function

Public Function dataSignUp(Username As String, Password As String) As Boolean
    Dim json As String
    If Username = "访客" Then
        json = "{'errorcode':'0'}"
    Else
        Randomize
        strUrl = "https://chemapp.njzjz.win/winsignup.php?username=" & Username & "&password=" & Password & "&t=" & Rnd
        Set xmlhttp = CreateObject("Microsoft.XMLHTTP")
        xmlhttp.Open "GET", strUrl, False
        xmlhttp.send
        If xmlhttp.ReadyState = 4 Then
            StrData = StrConv(xmlhttp.ResponseBody, vbUnicode)
        Else
            StrData = "{'errorcode':'-1'}"
        End If
        json = StrData
    End If
    If JSONParse("errorcode", json) = "0" Then
        dataSignUp = True
    Else
        dataSignUp = False
    End If
End Function

Public Function dataPasswordLock(x As String) As String
    Dim i As Integer, l As Integer
        i = 1
        l = Len(x)
        While i < l
            dataPasswordLock = dataPasswordLock & Chr(Asc(Mid(x, i, 1)) * 3 - 5)
            i = i + 1
        Wend
End Function

Public Function dataSignOut()
    HisAutoLogin = "False"
    Call dataSettingWrite("History", "AutoLogin", "False")
End Function

Public Function dataChangePassword(Username As String, Password As String, NewPassword As String) As Boolean
    Call dataOpen(2)
    DataAdodbRs.Open "select * from [User] where Username='" & Username & "'"
    If CStr(DataAdodbRs!Password) = dataPasswordLock(Password) Then
        dataChangePassword = True
        DataAdodbRs("Password") = dataPasswordLock(NewPassword)
        DataAdodbRs.Update
        HisPassword = ""
        Call dataSettingWrite("History", "Password", "")
        HisAutoLogin = "False"
        Call dataSettingWrite("History", "AutoLogin", "False")
    Else
        dataChangePassword = False
    End If
    Call dataClose
End Function

Public Function dataRenew() As Boolean
    dataRenew = True
    Call dataOpen(2)
    DataAdodbRs.Open "select * from [User]"
    While Not DataAdodbRs.EOF And dataRenew = True
        If Not IsNull(DataAdodbRs!Username) Then
            If CStr(DataAdodbRs!Username) = "访客" Then
                dataRenew = False
            Else
                DataAdodbRs.MoveNext
            End If
        End If
    Wend
    If dataRenew = True Then
        DataAdodbRs.AddNew
    End If
    DataAdodbRs("Password") = dataPasswordLock("user")
    DataAdodbRs("UseNumber") = 0
    DataAdodbRs("TimeMax") = 60
    DataAdodbRs("NumberMax") = 100
    DataAdodbRs("NoMax") = 20
    DataAdodbRs("ScoreMax") = 0
    DataAdodbRs("ScoreTime") = "N/A"
    DataAdodbRs("TimeIf") = "True"
    DataAdodbRs("Element") = ""
    DataAdodbRs("Mass") = ""
    DataAdodbRs("c") = ""
    DataAdodbRs("pKw") = "14"
    DataAdodbRs("pKa") = ""
    DataAdodbRs("AB") = "A"
    DataAdodbRs("qqname") = "访客"
    DataAdodbRs("CorrectNumber") = 0
    DataAdodbRs("IncorrectNumber") = 0
    DataAdodbRs.Update
    Call dataClose
End Function

Public Function dataHtmlLogin(Username As String, Password As String, SavingPassword As Integer, AutoLogin As Integer) As String
    frmMain.lblWelcome = "登陆中，请稍后……"
    If Username = "访客" Then
        Dim json As String
        json = "{'errorcode':'0'}"
        dataRenew
        Call dataLogin(Username, Password, SavingPassword, AutoLogin, json)
        frmMain.tmrFangke.Interval = 1
    Else
        Dim strUrl
        Randomize
        strUrl = "https://api.leancloud.cn/1.1/login" & "?t=" & Rnd
        Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
        xmlhttp.Open "POST", strUrl, True
        xmlhttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
        xmlhttp.SetRequestHeader "Content-Type", "application/json"
        xmlhttp.SetRequestHeader "X-LC-Id", "wUzGKF5dp34OqCeaI0VwVG8E-gzGzoHsz"
        xmlhttp.SetRequestHeader "X-LC-Key", "QiyXtJjBHFJCIVYQRbrKFiB7"
        xmlhttp.send "{""username"":""" & Username & """,""password"":""" & Password & """}"
        frmMain.tmrLogin.Interval = 1
    End If
End Function
Public Function JSONParse(ByVal JSONPath As String, ByVal JSONString As String) As Variant
    Dim json As Object
    Set json = CreateObject("MSScriptControl.ScriptControl")
    json.Language = "JScript"
    JSONParse = json.Eval("JSON=" & JSONString & ";JSON." & JSONPath & ";")
    Set json = Nothing
End Function

Public Function getNickname() As String
    If DataUsername = "访客" Or DataUsername = "" Then
        getNickname = "访客"
    ElseIf DataQQname <> "" Then
        getNickname = CStr(DataQQname)
    Else
        getNickname = Mid(DataUsername, 1, InStr(DataUsername, "@") - 1)
    End If
End Function

Public Function dataHtmlChange(Name As String, Value As String)
    Randomize
    Value = Replace(Value, Chr(10), "\n")
    'strUrl = "https://chemapp.njzjz.win/winchange.php?token=" & token & "&name=" & Name & "&value=" & Value & "&t=" & Rnd
    'Set xmlhttp = CreateObject("Microsoft.XMLHTTP")
 'If JSONParse("errorcode", strData) = "0" Then dataHtmlChange = True Else dataHtmlChange = False
    strUrl = "https://api.leancloud.cn/1.1/users/" & objectID & "?t=" & Rnd
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    xmlhttp.Open "PUT", strUrl, True
    xmlhttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
    xmlhttp.SetRequestHeader "X-LC-Id", "wUzGKF5dp34OqCeaI0VwVG8E-gzGzoHsz"
    xmlhttp.SetRequestHeader "X-LC-Key", "QiyXtJjBHFJCIVYQRbrKFiB7"
    xmlhttp.SetRequestHeader "X-LC-Session", token
    xmlhttp.SetRequestHeader "Content-Type", "application/json"
    xmlhttp.send "{""" & Name & """:""" & Value & """}"
    xmlhttp.waitforresponse
End Function

Public Function CutHtml(all As String) As String
    ok = ""
     S1 = InStr(1, all, "<")
    While S1 > 0 And S2 < Len(all)
        ok = ok & Mid(all, S2 + 1, S1 - S2 - 1)
        S2 = InStr(S1 + 1, all, ">")
        S1 = InStr(S2 + 1, all, "<")
    Wend
    ok = ok & Mid(all, S2 + 1, Len(all) - S2)
    ok = Replace(ok, Chr(10), Chr(13) & Chr(10))
    CutHtml = ok
End Function
