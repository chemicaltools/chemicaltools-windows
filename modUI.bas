Attribute VB_Name = "modUI"
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
'托盘
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Type NOTIFYICONDATA
       cbSize As Long   'NOTIFYICONDATA类型的大小，用 Len(变量名)获得即可
       hwnd As Long     '窗体句柄
       uId As Long      '图标资源的ID，通常使用 vbNull
       uFlags As Long   '使哪些参数有效，它是以下枚举类型中的 NIF_INFO Or NIF_ICON Or NIF_TIP Or NIF_MESSAGE 四个数的组合
       uCallBackMessage As Long   '接受消息的事件
       hIcon As Long   '图标句柄
       szTip As String * 128      '当鼠标停留在托盘上时显示的文本
       dwState As Long            '通常为 0
       dwStateMask As Long        '通常为 0
       szInfo As String * 256     'Tip 文本正文
       uTimeoutOrVersion As Long  'Tip 文本显示时间，由于 VB 没有 Union 类型，只能以 Long 代替
       szInfoTitle As String * 64 'Tip 文本标题
       dwInfoFlags As Long
End Type
Public Const NIF_INFO = &H10
Public sampleTray As NOTIFYICONDATA
Public Const NIIF_INFO = &H1

Public Function UICopy(x As String)
    Clipboard.Clear
    Clipboard.SetText x, vbCFText
End Function

Public Function UITime(x As String) As String
    UITime = Format(x, "yyyy年mm月dd日hh:mm:ss")
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

Public Function UIAddIcon()
    With sampleTray           '* 设置托盘属性
       .cbSize = Len(sampleTray)
       .cbSize = Len(sampleTray)
       .hwnd = frmMain.hwnd ''
       .uId = vbNull ''
       .uFlags = NIF_INFO Or NIF_ICON Or NIF_TIP Or NIF_MESSAGE
       .hIcon = frmMain.Icon
       .szInfoTitle = "化学小工具" & vbNullChar
       .szTip = szTip & vbNullChar
       .szInfo = "欢迎使用化学小工具！" & vbNullChar
       .dwState = 0
       .dwStateMask = 0
       .uTimeoutOrVersion = 2000
       .dwInfoFlags = NIIF_INFO
       .uCallBackMessage = WM_MOUSEMOVE
    End With
    Call Shell_NotifyIcon(NIM_ADD, sampleTray)
End Function
