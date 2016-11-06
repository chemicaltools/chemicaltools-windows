Attribute VB_Name = "modUI"
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
'����
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Type NOTIFYICONDATA
       cbSize As Long   'NOTIFYICONDATA���͵Ĵ�С���� Len(������)��ü���
       hwnd As Long     '������
       uId As Long      'ͼ����Դ��ID��ͨ��ʹ�� vbNull
       uFlags As Long   'ʹ��Щ������Ч����������ö�������е� NIF_INFO Or NIF_ICON Or NIF_TIP Or NIF_MESSAGE �ĸ��������
       uCallBackMessage As Long   '������Ϣ���¼�
       hIcon As Long   'ͼ����
       szTip As String * 128      '�����ͣ����������ʱ��ʾ���ı�
       dwState As Long            'ͨ��Ϊ 0
       dwStateMask As Long        'ͨ��Ϊ 0
       szInfo As String * 256     'Tip �ı�����
       uTimeoutOrVersion As Long  'Tip �ı���ʾʱ�䣬���� VB û�� Union ���ͣ�ֻ���� Long ����
       szInfoTitle As String * 64 'Tip �ı�����
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
    UITime = Format(x, "yyyy��mm��dd��hh:mm:ss")
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
    With sampleTray           '* ������������
       .cbSize = Len(sampleTray)
       .cbSize = Len(sampleTray)
       .hwnd = frmMain.hwnd
       .uId = vbNull
       .uFlags = NIF_INFO Or NIF_ICON Or NIF_TIP Or NIF_MESSAGE
       .hIcon = frmMain.Icon
       .szInfoTitle = "��ѧС����" & vbNullChar
       .szTip = szTip & vbNullChar
       .szInfo = "��ӭʹ�û�ѧС���ߣ�" & vbNullChar
       .dwState = 0
       .dwStateMask = 0
       .uTimeoutOrVersion = 2000
       .dwInfoFlags = NIIF_INFO
       .uCallBackMessage = WM_MOUSEMOVE
    End With
    Call Shell_NotifyIcon(NIM_ADD, sampleTray)
End Function

Public Sub UIDelIcon() '��ͼ���ϵͳ��������ɾ��
    sampleTray.uFlags = 0
    Shell_NotifyIcon NIM_DELETE, sampleTray
End Sub

Public Function RefreshAll(ByRef formp As Form) As Boolean
     Dim obj As Object
    For Each obj In formp
        obj.Refresh
    Next
End Function
