VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "µÇÂ½"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdSignUp 
      Caption         =   "×¢²á"
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "µÇÂ½"
      Default         =   -1  'True
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   3960
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
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   0
      Picture         =   "frmLogin.frx":14F4
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
Private Sub cmdLogin_Click()
    If dataLogin(texUsername, texPassword) = True Then
        Me.Hide
        frmMain.Show
        Unload Me
    Else
        MsgBox "ÓÃ»§Ãû»òÃÜÂë´íÎó£¡", vbOKOnly, "µÇÂ½Ê§°Ü"
    End If
End Sub

Private Sub cmdSignUp_Click()
    frmSignUp.Show 1
End Sub

Private Sub Form_Load()
    If HisUsername <> "" Then
        texUsername = HisUsername
        Me.Show
        texPassword.SetFocus
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub imgClose_Click()
    End
End Sub

Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
