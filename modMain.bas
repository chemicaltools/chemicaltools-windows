Attribute VB_Name = "modMain"
Sub Main()
    Call dataBaseDir
    Call dataSettingDir
    Call dataHistoryRead
    Load frmLogin
    frmLogin.Show
    Call dataElement
End Sub
