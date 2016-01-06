Attribute VB_Name = "modMain"
Sub main()
    Call dataBaseDir
    Call dataElement
    Call dataSettingDir
    Call dataSettingLoad
    Call dataHistoryRead
    Load frmMain
    frmMain.Show
End Sub

