VERSION 5.00
Begin VB.Form frmExam 
   Caption         =   "元素记忆测试 Designed by 团队一号"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   5790
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCopyScore 
      Caption         =   "分享战绩"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetting 
      Caption         =   "设置"
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   2160
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
      Left            =   3240
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox texExam 
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "开始测试"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblTime 
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
      Left            =   3120
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblScoreMax 
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
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   5295
   End
   Begin VB.Label lblScore 
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
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblCorrect 
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
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblExamElementName 
      Alignment       =   2  'Center
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2655
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
        ExamScoreName = InputBox("请输入你的姓名", "高分榜", ExamScoreName)
        ExamScoreTime = Now()
        ExamScoreMax = ExamScore
        dataScoreSave
        ExamScoreGet
    End If
    lblScore.Caption = "练习模式中" & Chr(13) & Chr(10) & "上次分数：" & Int(ExamScore)
    cmdStart.Caption = "开始测试"
    lblTime.Visible = False
    ExamScore = 0
End Function

Function ExamScoreGet()
    If ExamScoreMax > 0 Then
        lblScoreMax = "最高分为：" & ExamScoreMax & Chr(13) & Chr(10) & "由" & ExamScoreName & "于" & UITime(ExamScoreTime) & "创造"
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
