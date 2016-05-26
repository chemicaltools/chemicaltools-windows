Attribute VB_Name = "modMain"
Sub Main()
    Call dataBaseDir
    Call dataSettingDir
    Call dataHistoryRead
    Load frmLogin
    Call dataElement
    pKw = 14
End Sub
