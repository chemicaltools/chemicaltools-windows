VERSION 5.00
Begin VB.Form frmMass 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "质量计算器"
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   Icon            =   "frmMass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   7785
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制到剪切板"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox texMassOut 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMass.frx":1B692
      Top             =   3000
      Width           =   7695
   End
   Begin VB.CommandButton cmdMass 
      Caption         =   "计算！"
      Default         =   -1  'True
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox texMassIn 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   7320
      Picture         =   "frmMass.frx":1B6A9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   0
      Picture         =   "frmMass.frx":1CFA1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmMass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private InTip As String

Private Sub cmdCopy_Click()
    UICopy (texMassOut.texT)
End Sub

Private Sub cmdMass_Click()
    texMassOut = calMassPerStr(texMassIn.texT)
    HisMass = texMassIn.texT
    Call dataBaseWrite(DataUsername, "Mass", HisMass)
    If Not DataUsername = "访客" Then
        Call dataHtmlChange("historyMass", CStr(HisMass))
    End If
End Sub

Private Sub Form_Load()
    InTip = "请在此处输入物质化学式，例如：K4[Fe(CN)6]"
    texMassIn.texT = InTip
    If HisMass <> "" Then texMassOut = calMassPerStr(HisMass)
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


Private Sub texMassIn_Click()
    If texMassIn.texT = InTip Then
        texMassIn.texT = ""
        texMassIn.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texMassIn_KeyPress(KeyAscii As Integer)
    If texMassIn.texT = InTip Then
        texMassIn.texT = ""
        texMassIn.ForeColor = RGB(0, 0, 0)
    End If
End Sub

