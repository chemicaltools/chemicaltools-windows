VERSION 5.00
Begin VB.Form frmExam 
   Caption         =   "元素记忆测试 Designed by 团队一号"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
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
      Left            =   3240
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
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
Private ExamElementNumber As Integer
Private ExamNumberMax As Integer
Private InTip As String
Private ExamNo As Integer
Private ExamNoMax As Integer
Private ExamIf As Boolean
Private ExamScore As Integer

Function ExamNew()
    ExamElementNumber = ExamRnd(ExamNumberMax)
    lblExamElementName.Caption = ElementName(ExamElementNumber)
End Function

Private Sub cmdExam_Click()
    If ExamIf Then ExamNo = ExamNo + 1
    If ExamAbbr(ExamElementNumber, texExam.Text) Then
        lblCorrect.Caption = "恭喜你，答对了！" & Chr(13) & Chr(10) & ElementName(ExamElementNumber) & "的符号为：" & ElementAbbr(ExamElementNumber)
        ExamScore = ExamScore + 1
    Else
        lblCorrect.Caption = "很遗憾，答错了！" & Chr(13) & Chr(10) & ElementName(ExamElementNumber) & "的符号为：" & ElementAbbr(ExamElementNumber)
    End If
    texExam.SetFocus
    texExam.Text = InTip
    texExam.ForeColor = RGB(128, 128, 128)
    If ExamNo >= ExamNoMax Then
        ExamIf = False
        ExamNo = 0
        ExamScore = 0
        MsgBox ("答题结束！你的分数为：" & ExamScore)
    End If
    Call ExamNew
End Sub

Private Sub cmdStart_Click()
    ExamIf = True
    texExam.SetFocus
    texExam.Text = InTip
    texExam.ForeColor = RGB(128, 128, 128)
    Call ExamNew
End Sub

Private Sub Form_Load()
    ExamNumberMax = 100 '临时设置
    ExamNoMax = 20 '临时设置
    InTip = "请输入所给元素的符号～"
    texExam.Text = InTip
    ExamIf = False
    ExamNo = 0
    ExamScore = 0
    Call ExamNew
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
