Attribute VB_Name = "modMain"
Sub main()
    Load frmSplash
    frmSplash.Show
    Call dataBaseDir
    Call dataElement
    Call dataSettingDir
    Call dataSettingLoad
    Call dataHistoryRead
    Load frmMain
    frmSplash.Hide
    frmMain.Show
    Unload frmSplash
End Sub

