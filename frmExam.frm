VERSION 5.00
Begin VB.Form frmExam 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ԫ�ؼ������"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   Icon            =   "frmExam.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   6165
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCopyScore 
      Caption         =   "����ս��"
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetting 
      Caption         =   "����"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Timer tmrExam 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdExam 
      Caption         =   "�ύ"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox texExam 
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "��ʼ����"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblScoreAll 
      BackStyle       =   0  'Transparent
      Caption         =   "��ϰģʽ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   5760
      Picture         =   "frmExam.frx":1B692
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "��ʣ10��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblScoreMax 
      BackStyle       =   0  'Transparent
      Caption         =   "��߷�Ϊ��0 ��ZJZ��2015��10��24��17:22����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   4920
      Width           =   6135
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "��ϰģʽ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   5
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lblCorrect 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label lblExamElementName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   48
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   -600
      Picture         =   "frmExam.frx":1CF8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private InTip As String

Function ExamNew()
    ExamElementNumber = ExamRnd(ExamNumberMax)
    lblExamElementName.Caption = ElementName(ExamElementNumber)
End Function

Function ExamStart()
    ExamIf = True
    texExam.SetFocus
    texExam.texT = InTip
    texExam.ForeColor = RGB(128, 128, 128)
    lblScore.Caption = "��ǰ����Ϊ��" & Int(ExamScore)
    If ExamTimeIf = True Then
        tmrExam.Enabled = True
        ExamTime = ExamTimeMax
        lblTime.Visible = True
        lblTime.Caption = "��ʣ" & ExamTime & "��"
    End If
    Call ExamNew
    cmdStart.Caption = "��������"
End Function

Function ExamEnd()
    ExamIf = False
    ExamNo = 0
    tmrExam.Enabled = False
    MsgBox ("�����������ķ���Ϊ��" & Int(ExamScore))
    If ExamScore > ExamScoreMax Then
        ExamScoreTime = Now()
        ExamScoreMax = ExamScore
        dataScoreSave (getNickname)
        ExamScoreGet
    End If
    lblScore.Caption = "��ϰģʽ��" & Chr(13) & Chr(10) & "�ϴη�����" & Int(ExamScore)
    cmdStart.Caption = "��ʼ����"
    lblTime.Visible = False
    ExamScore = 0
End Function

Function ExamScoreGet()
    If ExamScoreMax > 0 Then
        lblScoreMax = "������߷�Ϊ��" & ExamScoreMax & "����" & UITime(ExamScoreTime) & "���죻" & Chr(13) & Chr(10) & "�����û�����߷�Ϊ��" & ExamScoreMaxAll & "��" & Chr(13) & Chr(10) & "��" & ExamScoreNameAll & "��" & UITime(ExamScoreTimeAll) & "���졣"
    Else
        lblScoreMax = ""
    End If
End Function

Private Sub cmdCopyScore_Click()
    Call UICopy("Hello, ����" & UITime(ExamScoreTime) & "������Ԫ�ؼ�����ԣ��õ���" & ExamScoreMax & "�֣���Ҳ��ʹ�û�ѧС�������԰ɣ�")
    MsgBox "�������ս���Ѿ����Ƶ����а壡"
End Sub

Private Sub cmdExam_Click()
    If ExamIf Then ExamNo = ExamNo + 1
    If ExamAbbr(ExamElementNumber, texExam.texT) Then
        lblCorrect.Caption = "��ϲ�㣬����ˣ�" & Chr(13) & Chr(10) & ElementName(ExamElementNumber) & "�ķ���Ϊ��" & ElementAbbr(ExamElementNumber)
        examCorrectNumber = examCorrectNumber + 1
        If Not DataUsername = "�ÿ�" Then
            Call dataHtmlChange("examCorrectNumber", CStr(examCorrectNumber))
        End If
        If ExamIf Then
            ExamScore = ExamScore + 100 / ExamNoMax
            lblScore.Caption = "��ǰ����Ϊ��" & Int(ExamScore)
        End If
    Else
        lblCorrect.Caption = "���ź�������ˣ�" & Chr(13) & Chr(10) & ElementName(ExamElementNumber) & "�ķ���Ϊ��" & ElementAbbr(ExamElementNumber)
        examIncorrectNumber = examIncorrectNumber + 1
        If Not DataUsername = "�ÿ�" Then
            Call dataHtmlChange("examIncorrectnumber", CStr(examIncorrectNumber))
        End If
    End If
    lblScoreAll = "��ս����" & examCorrectNumber & "��" & examIncorrectNumber & "��"
    texExam.SetFocus
    texExam.texT = InTip
    texExam.ForeColor = RGB(128, 128, 128)
    If ExamNo >= ExamNoMax Then ExamEnd
    Call ExamNew
End Sub

Private Sub cmdSetting_Click()
    frmOptions.Show 1, Me
End Sub

Private Sub cmdStart_Click()
    If ExamIf = False Then
        Call ExamStart
    Else
        Call ExamEnd
    End If
End Sub

Private Sub Form_Load()
    InTip = "����������Ԫ�صķ��š�"
    texExam.texT = InTip
    lblScoreAll = "��ս����" & examCorrectNumber & "��" & examIncorrectNumber & "��"
    lblScore.Caption = "��ϰģʽ��"
    lblTime.Caption = ""
    ExamIf = False
    ExamNo = 0
    ExamScore = 0
    Call ExamNew
    ExamScoreGet
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub imgClose_Click()
    Unload Me
End Sub

Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub texExam_Click()
    If texExam.texT = InTip Then
        texExam.texT = ""
        texExam.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texExam_KeyPress(KeyAscii As Integer)
    If texExam.texT = InTip Then
        texExam.texT = ""
        texExam.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub tmrExam_Timer()
    ExamTime = ExamTime - 1
    lblTime = "��ʣ" & ExamTime & "��"
    If ExamTime <= 0 Then
        ExamEnd
    End If
End Sub

