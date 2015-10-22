Attribute VB_Name = "modData"
Public DataAdodbConn As ADODB.Connection
Public DataAdodbRs As ADODB.Recordset

Public ElementName(118) As String
Public ElementAbbr(118) As String
Public ElementMass(118) As Double

Public Function dataOpen()
    Dim path As String
    Set DataAdodbConn = New ADODB.Connection
    path = dataBasepath
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

Public Function dataBasepath() As String
    Dim spath As String
    If Right$(App.path, 1) = "\" Then
        spath = App.path
    Else
        spath = App.path + "\"
    End If
    dataBasepath = spath + "Data.mdb"
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
