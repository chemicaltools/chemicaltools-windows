VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "选项"
   ClientHeight    =   7515
   ClientLeft      =   2520
   ClientTop       =   1110
   ClientWidth     =   6090
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   6090
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "示例 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "示例 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   10
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "示例 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   9
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   240
      ScaleHeight     =   3780
      ScaleWidth      =   5445
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2640
      Width           =   5445
      Begin VB.Frame fraExamElement 
         BackColor       =   &H00C0FFFF&
         Caption         =   "答题范围设置"
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   4095
         Begin VB.ComboBox cboNumber 
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "frmOptions.frx":1B692
            Left            =   1320
            List            =   "frmOptions.frx":1B6AB
            TabIndex        =   18
            Text            =   "86"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox texNo 
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   16
            Text            =   "20"
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblExamElement 
            BackStyle       =   0  'Transparent
            Caption         =   "元素范围                     号元素                                                   答题数量                     题"
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
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame fraExamTime 
         BackColor       =   &H00C0FFFF&
         Caption         =   "时间设置"
         Height          =   1545
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   4095
         Begin VB.TextBox texTime 
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   13
            Text            =   "60"
            Top             =   840
            Width           =   1335
         End
         Begin VB.CheckBox chkTimeIf 
            BackColor       =   &H00C0FFFF&
            Caption         =   "限时测试"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.Label lblTimeIf 
            BackColor       =   &H00C0FFFF&
            Caption         =   "限时                           秒"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   2415
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   6840
      Width           =   1095
   End
   Begin VB.PictureBox tbsOptions 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   120
      ScaleHeight     =   4185
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   5640
      Picture         =   "frmOptions.frx":1B6C4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   -600
      Picture         =   "frmOptions.frx":1CFBC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function ReadOptions()
    cboNumber.texT = Trim(str(ExamNumberMax))
    texNo.texT = Trim(str(ExamNoMax))
    texTime.texT = Trim(str(ExamTimeMax))
    If ExamTimeIf = True Then chkTimeIf.Value = 1 Else chkTimeIf.Value = 0
    CheckTimeIf
    cmdApply.Enabled = False
End Function

Private Function WriteOptions() As Boolean
    Dim NumMax As Integer, NoMax As Integer, TimeMax As Integer, ErrorInfo As String
    WriteOptions = False
    NumMax = Int(Val(cboNumber.texT))
    NoMax = Int(Val(texNo.texT))
    TimeMax = Int(Val(texTime.texT))
    If Not (NumMax < 119 And NumMax > 0) Then
        WriteOptions = True
        ErrorInfo = "元素范围输入错误！"
    End If
    If Not (NoMax > 0) Then
        WriteOptions = True
        ErrorInfo = "题目数量输入错误！"
    End If
    If Not (TimeMax > 0) Then
        WriteOptions = True
        ErrorInfo = "答题时间输入错误！"
    End If
    If WriteOptions = False Then
        ExamNumberMax = NumMax
        ExamNoMax = NoMax
        ExamTimeMax = TimeMax
        If chkTimeIf.Value = 1 Then ExamTimeIf = True Else ExamTimeIf = False
        Call dataSettingSave(DataUsername)
        If Not DataUsername = "访客" Then
            Call dataHtmlChange("elementnumber_limit", CStr(ExamNumberMax))
        End If
    Else
        MsgBox ErrorInfo
    End If
End Function

Private Function CheckTimeIf()
    If chkTimeIf.Value = 0 Then
        texTime.Enabled = False
    Else
        texTime.Enabled = True
    End If
End Function

Private Sub chkTimeIf_Click()
    CheckTimeIf
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
    If Not WriteOptions Then cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not WriteOptions Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    ReadOptions
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub imgClose_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


Private Sub texNo_KeyPress(KeyAscii As Integer)
    cmdApply.Enabled = True
End Sub

Private Sub texNumber_KeyPress(KeyAscii As Integer)
    cmdApply.Enabled = True
End Sub

Private Sub texTime_KeyPress(KeyAscii As Integer)
    cmdApply.Enabled = True
End Sub
