Attribute VB_Name = "modData"
'配置文件
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'数据库
Public DataAdodbConn As ADODB.Connection
Public DataAdodbRs As ADODB.Recordset
'内存
'配置
Public ExamTimeMax As Integer
Public ExamNumberMax As Integer
Public ExamNoMax As Integer
Public ExamScoreMax As Integer
Public ExamScoreName As String
Public ExamTimeIf As Boolean
Public ExamScoreTime As String
'历史记录
Public HisElement As String
Public HisMass As String
'元素
Public ElementName(118) As String
Public ElementAbbr(118) As String
Public ElementMass(118) As Double

Public Function dataOpen()
    Dim path As String
    Set DataAdodbConn = New ADODB.Connection
    path = dataBasePath
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

Public Function dataBasePath() As String
    Dim spath As String
    If Right$(App.path, 1) = "\" Then
        spath = App.path + "mdb\"
    Else
        spath = App.path + "\mdb\"
    End If
    dataBasePath = spath + "Data.mdb"
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
    Call dataOpen
    DataAdodbRs.Open "select * from Element"
    While Not DataAdodbRs.EOF
        i = i + 1
        If Not IsNull(DataAdodbRs!ElementNumber) Then n = CStr(DataAdodbRs!ElementNumber)
        If Not IsNull(DataAdodbRs!ElementName) Then ElementName(n) = CStr(DataAdodbRs!ElementName)
        If Not IsNull(DataAdodbRs!ElementAbbr) Then ElementAbbr(n) = CStr(DataAdodbRs!ElementAbbr)
        If Not IsNull(DataAdodbRs!ElementMass) Then ElementMass(n) = CStr(DataAdodbRs!ElementMass)
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

Public Function dataSettingWrite(x As String, y As String, z As String)
    Call WritePrivateProfileString(x, y, z, dataSettingPath)
End Function

Public Function dataSettingRead(x As String, y As String) As String
    Dim z As String, p As Long, i As Integer
    z = String(255, vbNullChar)
    p = GetPrivateProfileString(x, y, "", z, 255, dataSettingPath)
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
    If dataDir(dataBasePath) = False Then MsgBox "数据库文件缺失，请联系团队一号！"
End Function

Public Function dataSettingDir()
    If dataDir(dataSettingPath) = False Then
        Call dataSettingWrite("Exam", "TimeMax", "60")
        Call dataSettingWrite("Exam", "NumberMax", "100")
        Call dataSettingWrite("Exam", "NoMax", "20")
        Call dataSettingWrite("Exam", "ScoreMax", "0")
        Call dataSettingWrite("Exam", "ScoreName", "N/A")
        Call dataSettingWrite("Exam", "ScoreTime", "N/A")
        Call dataSettingWrite("Exam", "TimeIf", "True")
    End If
End Function

Public Function dataSettingLoad()
    If dataDir(dataSettingPath) = True Then
        ExamTimeMax = Int(Val(dataSettingRead("Exam", "TimeMax")))
        ExamNumberMax = Int(Val(dataSettingRead("Exam", "NumberMax")))
        ExamNoMax = Int(Val(dataSettingRead("Exam", "NoMax")))
        ExamScoreMax = Int(Val(dataSettingRead("Exam", "ScoreMax")))
        ExamScoreName = dataSettingRead("Exam", "ScoreName")
        ExamScoreTime = dataSettingRead("Exam", "ScoreTime")
        If dataSettingRead("Exam", "TimeIf") = "True" Then ExamTimeIf = True Else ExamTimeIf = False
    End If
End Function

Public Function dataSettingSave()
    Call dataSettingWrite("Exam", "TimeMax", Trim(str(ExamTimeMax)))
    Call dataSettingWrite("Exam", "NumberMax", Trim(str(ExamNumberMax)))
    Call dataSettingWrite("Exam", "NoMax", Trim(str(ExamNoMax)))
    If ExamTimeIf = True Then Call dataSettingWrite("Exam", "TimeIf", "True") Else Call dataSettingWrite("Exam", "TimeIf", "False")
End Function

Public Function dataScoreSave()
    Call dataSettingWrite("Exam", "ScoreMax", Trim(str(ExamScoreMax)))
    Call dataSettingWrite("Exam", "ScoreName", ExamScoreName)
    Call dataSettingWrite("Exam", "ScoreTime", ExamScoreTime)
End Function

Public Function dataHistoryRead()
    HisElement = dataSettingRead("History", "Element")
    HisMass = dataSettingRead("History", "Mass")
End Function
