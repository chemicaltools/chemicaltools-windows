Attribute VB_Name = "modUI"
Public Function UIMove()

End Function

Public Function UICopy(x As String)
    Clipboard.Clear
    Clipboard.SetText x, vbCFText
End Function

Public Function UITime(x As String) As String
    UITime = Format(x, "yyyyƒÍmm‘¬dd»’hh:mm:ss")
End Function

