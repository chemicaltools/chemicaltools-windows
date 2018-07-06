VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "ªØ—ße+"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8235
   StartUpPosition =   1  'À˘”–’ﬂ÷––ƒ
   Begin VB.CommandButton cmdfrmOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "πÿ”⁄"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   6240
      Picture         =   "main.frx":1084A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "«–ªª”√ªß"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   4320
      Picture         =   "main.frx":2BEDC
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘™Àÿº«“‰"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   2400
      Picture         =   "main.frx":4756E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "∆¯ÃÂº∆À„"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   480
      Picture         =   "main.frx":62C00
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "»»¡¶—ßº∆À„"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   6240
      Picture         =   "main.frx":7E292
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "À·ºÓº∆À„"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   4320
      Picture         =   "main.frx":99924
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "÷ ¡øº∆À„"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   2400
      Picture         =   "main.frx":B4FB6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Timer tmrFangke 
      Left            =   6000
      Top             =   2160
   End
   Begin VB.Timer tmrLogin 
      Left            =   6480
      Top             =   2160
   End
   Begin VB.CommandButton cmdfrmOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘™Àÿ≤È—Ø"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   480
      Picture         =   "main.frx":D0648
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "ª∂”≠£°"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
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
      Top             =   2520
      Width           =   6495
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   7800
      Picture         =   "main.frx":EBCDA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2385
      Left            =   0
      Picture         =   "main.frx":ED5D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdfrmOpen_Click(Index As Integer)
    Select Case Index
        Case 0
            frmElement.Show
        Case 1
            frmMass.Show
        Case 2
            frmpH.Show
        Case 3
            frmThermodynamics.Show
        Case 4
            frmGas.Show
        Case 5
            frmExam.Show
        Case 6
            frmLogin.Show
        Case 7
            frmAbout.Show 1, Me
    End Select
End Sub


Private Sub cmdSignOut_Click()
    'dataSignOut
    If UIFormLoad(frmLogin) Then Unload frmLogin
    Load frmLogin
    frmLogin.Show
End Sub

Private Sub Form_Load()
    lblWelcome = "ª∂”≠" & getNickname() & "µ⁄" & str(DataUseNumber) & "¥Œ π”√" + App.Title + "£°"
    If DataUsername = "∑√øÕ" Or DataUsername = "" Then
        cmdfrmOpen(6).Caption = "µ«¬Ω"
    Else
        cmdfrmOpen(6).Caption = "«–ªª”√ªß"
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
        MsgBox "”√ªß√˚ªÚ√‹¬Î¥ÌŒÛ£°", vbOKOnly, "µ«¬Ω ß∞‹"
        Form_Load
    End If
End Sub
