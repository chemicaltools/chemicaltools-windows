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
   Begin VB.Timer tmrFangke 
      Left            =   6480
      Top             =   5040
   End
   Begin VB.Timer tmrLogin 
      Left            =   6960
      Top             =   5040
   End
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
      Picture         =   "main.frx":1084A
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
      Picture         =   "main.frx":2BEDC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   0
      Picture         =   "main.frx":2D7D4
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
    If DataUsername = "访客" Or DataUsername = "" Then
        cmdSignOut.Caption = "登陆"
    Else
        cmdSignOut.Caption = "切换用户"
    End If
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
End Function

Private Sub tmrFangke_Timer()
    Form_Load
    tmrFangke.Interval = 0
End Sub

Private Sub tmrLogin_Timer()
    xmlhttp.WaitForResponse
    If xmlhttp.Status = 200 Then
        Dim json As String
        json = xmlhttp.ResponseText
        tmrLogin.Interval = 0
        If dataLogin(LoginUsername, LoginPassword, LoginSavingPassword, LoginAutoLogin, json) Then
            Form_Load
            If First Then
                Fangke = False
                First = False
            Else
                Dim frm As Form
                For Each frm In Forms
                    If Not frm Is frmLogin And Not frm Is frmMain Then
                        Unload frm
                        Load frm
                        frm.Show
                    End If
                Next
            End If
          End If
    ElseIf xmlhttp.Status = 400 Then
        tmrLogin.Interval = 0
        frmLogin.Show
        MsgBox "用户名或密码错误！", vbOKOnly, "登陆失败"
        Form_Load
    End If
End Sub
