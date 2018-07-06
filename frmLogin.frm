VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "µÇÂ½"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7335
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command1 
      Caption         =   "·Ã¿ÍÄ£Ê½"
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox chkAutoLogin 
      BackColor       =   &H00C0FFFF&
      Caption         =   "×Ô¶¯µÇÂ½"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CheckBox chkPassword 
      BackColor       =   &H00C0FFFF&
      Caption         =   "¼Ç×¡ÃÜÂë"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdSignUp 
      Caption         =   "×¢²á"
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "µÇÂ½"
      Default         =   -1  'True
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox texPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3240
      Width           =   4575
   End
   Begin VB.TextBox texUsername 
      Height          =   375
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "ÃÜ   Âë"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "ÓÃ»§Ãû"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   6840
      Picture         =   "frmLogin.frx":048A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   0
      Picture         =   "frmLogin.frx":1D82
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub chkAutoLogin_Click()
    If chkAutoLogin.Value = 1 And chkPassword.Value <> 1 Then chkPassword.Value = 1
End Sub

Private Sub cmdLogin_Click()
    LoginUsername = texUsername
    LoginPassword = texPassword
    LoginSavingPassword = chkPassword.Value
    LoginAutoLogin = chkAutoLogin.Value
    Me.Hide
    frmMain.Show
    Call dataHtmlLogin(LoginUsername, LoginPassword, LoginSavingPassword, LoginAutoLogin)
End Sub

Private Sub cmdSignUp_Click()
    frmSignUp.Show 1
End Sub

Private Sub Command1_Click()
    Me.Hide
    frmMain.Show
    Call dataHtmlLogin("·Ã¿Í", "user", 0, 0)
End Sub

Private Sub Form_Load()
    If HisUsername <> "" And HisUsername <> "·Ã¿Í" Then
        texUsername = HisUsername
        If HisPassword <> "" Then
            texPassword = HisPassword
            chkPassword.Value = 1
        End If
        If HisAutoLogin = "True" Then
            chkAutoLogin.Value = 1
        End If
        If Not First Then
            Me.Show
            texPassword.SetFocus
        End If
        If First And HisAutoLogin = "True" Then cmdLogin_Click
    End If
    If Fangke Then
        Call Command1_Click
    End If
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
