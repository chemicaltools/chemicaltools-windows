VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "关于我的应用程序"
   ClientHeight    =   5115
   ClientLeft      =   2295
   ClientTop       =   1545
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3530.464
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   780
      Left            =   240
      Picture         =   "frmAbout.frx":1B692
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   1
      Top             =   2040
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   4080
      TabIndex        =   0
      Top             =   4440
      Width           =   1500
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   5280
      Picture         =   "frmAbout.frx":36D24
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   -720
      Picture         =   "frmAbout.frx":38218
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   7575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   2816.089
      Y2              =   2816.089
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":439B2
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Chemical Tools"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "版本"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   5
      Top             =   2520
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "警告：  本作品受《中华人民共和国版权法》保护，若未经授权复制或传播其中任何部分，将追究法律责任。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   3630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "关于 " & App.Title
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
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
