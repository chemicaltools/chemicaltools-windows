Attribute VB_Name = "modMain"
Public R As Double
Sub Main()
    R = 8.314
    Call dataBaseDir
    Call dataSettingDir
    Call dataHistoryRead
    Load frmLogin
    Call dataElement
End Sub
