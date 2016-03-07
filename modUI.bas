Attribute VB_Name = "modUI"
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Function UICopy(x As String)
    Clipboard.Clear
    Clipboard.SetText x, vbCFText
End Function

Public Function UITime(x As String) As String
    UITime = Format(x, "yyyyƒÍmm‘¬dd»’hh:mm:ss")
End Function

Public Function UIFormLoad(ByRef frmIn As Form) As Boolean
    Dim frm As Form
    For Each frm In Forms
        If frmIn Is frm Then
            UIFormLoad = True
            Exit For
        End If
    Next
End Function
