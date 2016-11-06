Attribute VB_Name = "modMain"
Public R As Double
Public Fangke As Double
Public First As Double
Sub Main()
    R = 8.314
    Call dataBaseDir
    Call dataSettingDir
    Call dataHistoryRead
    First = True
    If Not (HisUsername <> "" And HisUsername <> "·Ã¿Í" And HisPassword <> "" And HisAutoLogin = "True") Then
        Fangke = True
    Else
        Fangke = False
    End If
    Load frmLogin
    Call dataElement
End Sub
