VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "化学e+ Designed by 团队一号"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7425
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdGas 
      Caption         =   "气体计算"
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdRelixue 
      Caption         =   "热力学计算"
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdpH 
      BackColor       =   &H00C0FFFF&
      Caption         =   "酸碱计算"
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdSignOut 
      Caption         =   "切换用户"
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmAbout 
      Caption         =   "关于"
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmExam 
      Caption         =   "元素记忆"
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmElement 
      BackColor       =   &H00C0FFFF&
      Caption         =   "元素查询"
      Height          =   615
      Left            =   720
      Picture         =   "main.frx":048A
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmMass 
      Caption         =   "质量计算"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   6495
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   6960
      Picture         =   "main.frx":1BB1C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   0
      Picture         =   "main.frx":1D414
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdfrmAbout_Click()
    Load frmAbout
    frmAbout.Show 1, Me
End Sub

Private Sub cmdfrmElement_Click()
    Load frmElement
    frmElement.Show
End Sub

Private Sub cmdfrmExam_Click()
    Load frmExam
    frmExam.Show
End Sub

Private Sub cmdfrmmass_Click()
    Load frmMass
    frmMass.Show
End Sub

Private Sub cmdGas_Click()
    Load frmGas
    frmGas.Show
End Sub

Private Sub cmdNewPassword_Click()
    If DataUsername = "访客" Then
        MsgBox "访客禁止修改密码！", vbOKOnly, "错误"
    Else
        FrmChangePassword.Show 1
    End If
End Sub

Private Sub cmdpH_Click()
    Load frmpH
    frmpH.Show
End Sub

Private Sub cmdRelixue_Click()
    Load frmThermodynamics
    frmThermodynamics.Show
End Sub

Private Sub cmdSignOut_Click()
    'dataSignOut
    If UIFormLoad(frmLogin) Then Unload frmLogin
    Load frmLogin
    frmLogin.Show
End Sub

Private Sub Form_Load()
    lblWelcome = "欢迎" & getNickname() & "第" & str(DataUseNumber) & "次使用化学e+！"
    If DataUsername = "访客" Then cmdSignOut.Caption = "登陆"
    '托盘
    'Call UIAddIcon
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveToCloud
End Sub

Private Sub imgClose_Click()
    SaveToCloud
    'Call UIDelIcon
    End
End Sub

Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Function SaveToCloud()
    Dim frm As Form
    For Each frm In Forms
        frm.Hide
    Next

    'If Not DataUsername = "访客" Then
    '    Call dataHtmlChange("examIncorrectnumber", CStr(examIncorrectNumber))
    '    Call dataHtmlChange("examCorrectNumber", CStr(examCorrectNumber))
    '    Call dataHtmlChange("elementnumber_limit", CStr(ExamNumberMax))
    '    Call dataHtmlChange("pKw", CStr(HispKw))
    '    Call dataHtmlChange("historyElement", CStr(HisElement))
    '    Call dataHtmlChange("historyMass", CStr(HisMass))
    'End If
    
End Function

