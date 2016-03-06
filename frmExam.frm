VERSION 5.00
Begin VB.Form frmExam 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "元素记忆测试 Designed by 团队一号"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   Icon            =   "frmExam.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCopyScore 
      Caption         =   "分享战绩"
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetting 
      Caption         =   "设置"
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
      Caption         =   "提交"
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
      Caption         =   "开始测试"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
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
      Caption         =   "还剩10秒"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "最高分为：0 由ZJZ于2015年10月24日17:22创造"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "练习模式中"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblCorrect 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "氢"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Picture         =   "frmExam.frx":1CB86
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
    texExam.Text = InTip
    texExam.ForeColor = RGB(128, 128, 128)
    lblScore.Caption = "当前分数为：" & Chr(13) & Chr(10) & Int(ExamScore)
    If ExamTimeIf = True Then
        tmrExam.Enabled = True
        ExamTime = ExamTimeMax
        lblTime.Visible = True
        lblTime.Caption = "还剩" & ExamTime & "秒"
    End If
    Call ExamNew
    cmdStart.Caption = "结束测试"
End Function

Function ExamEnd()
    ExamIf = False
    ExamNo = 0
    tmrExam.Enabled = False
    MsgBox ("答题结束！你的分数为：" & Int(ExamScore))
    If ExamScore > ExamScoreMax Then
        ExamScoreTime = Now()
        ExamScoreMax = ExamScore
        dataScoreSave (DataUsername)
        ExamScoreGet
    End If
    lblScore.Caption = "练习模式中" & Chr(13) & Chr(10) & "上次分数：" & Int(ExamScore)
    cmdStart.Caption = "开始测试"
    lblTime.Visible = False
    ExamScore = 0
End Function

Function ExamScoreGet()
    If ExamScoreMax > 0 Then
        lblScoreMax = "您的最高分为：" & ExamScoreMax & "，于" & UITime(ExamScoreTime) & "创造；" & Chr(13) & Chr(10) & "所有用户的最高分为：" & ExamScoreMaxAll & "，" & Chr(13) & Chr(10) & "由" & ExamScoreNameAll & "于" & UITime(ExamScoreTimeAll) & "创造。"
    Else
        lblScoreMax = ""
    End If
End Function

Private Sub cmdCopyScore_Click()
    Call UICopy("Hello, 我于" & UITime(ExamScoreTime) & "进行了元素记忆测试，得到了" & ExamScoreMax & "分！你也来使用化学小工具试试吧！")
    MsgBox "您的最高战绩已经复制到剪切板！"
End Sub

Private Sub cmdExam_Click()
    If ExamIf Then ExamNo = ExamNo + 1
    If ExamAbbr(ExamElementNumber, texExam.Text) Then
        lblCorrect.Caption = "恭喜你，答对了！" & Chr(13) & Chr(10) & ElementName(ExamElementNumber) & "的符号为：" & ElementAbbr(ExamElementNumber)
        If ExamIf Then
            ExamScore = ExamScore + 100 / ExamNoMax
            lblScore.Caption = "当前分数为：" & Chr(13) & Chr(10) & Int(ExamScore)
        End If
    Else
        lblCorrect.Caption = "很遗憾，答错了！" & Chr(13) & Chr(10) & ElementName(ExamElementNumber) & "的符号为：" & ElementAbbr(ExamElementNumber)
    End If
    texExam.SetFocus
    texExam.Text = InTip
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
    InTip = "请输入所给元素的符号～"
    texExam.Text = InTip
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
    If texExam.Text = InTip Then
        texExam.Text = ""
        texExam.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texExam_KeyPress(KeyAscii As Integer)
    If texExam.Text = InTip Then
        texExam.Text = ""
        texExam.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub tmrExam_Timer()
    ExamTime = ExamTime - 1
    lblTime = "还剩" & ExamTime & "秒"
    If ExamTime <= 0 Then
        ExamEnd
    End If
End Sub
