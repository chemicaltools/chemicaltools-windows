VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "化学小工具 Designed by 团队一号"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdfrmAbout 
      Caption         =   "关于"
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdfrmExam 
      Caption         =   "元素记忆"
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdfrmElement 
      Caption         =   "元素查询"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdfrmMass 
      Caption         =   "质量计算"
      Height          =   615
      Left            =   1920
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
      Picture         =   "main.frx":1B692
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   0
      Picture         =   "main.frx":1CB86
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
    frmAbout.Show
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
    frmMass.Show 1, Me
End Sub

Private Sub Form_Load()
    lblWelcome = "欢迎" & DataUsername & "！这是您第" & str(DataUseNumber) & "次使用化学小工具！"
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
