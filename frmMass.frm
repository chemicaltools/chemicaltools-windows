VERSION 5.00
Begin VB.Form frmMass 
   BackColor       =   &H00FFFFFF&
   Caption         =   "质量计算器 Designed by 团队一号"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7950
   Icon            =   "frmMass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7950
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制到剪切板"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
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
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMass.frx":1B692
      Top             =   840
      Width           =   7695
   End
   Begin VB.CommandButton cmdMass 
      Caption         =   "计算！"
      Default         =   -1  'True
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   240
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmMass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private InTip As String

Private Sub cmdCopy_Click()
    UICopy (texMassOut.Text)
End Sub

Private Sub cmdMass_Click()
    texMassOut = calMassPerStr(texMassIn.Text)
    HisMass = texMassIn.Text
    Call dataSettingWrite("History", "Mass", HisMass)
End Sub

Private Sub Form_Load()
    InTip = "请在此处输入物质化学式，例如：K4[Fe(CN)6]"
    texMassIn.Text = InTip
    If HisMass <> "" Then texMassOut = calMassPerStr(HisMass)
End Sub

Private Sub texMassIn_Click()
    If texMassIn.Text = InTip Then
        texMassIn.Text = ""
        texMassIn.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texMassIn_KeyPress(KeyAscii As Integer)
    If texMassIn.Text = InTip Then
        texMassIn.Text = ""
        texMassIn.ForeColor = RGB(0, 0, 0)
    End If
End Sub

