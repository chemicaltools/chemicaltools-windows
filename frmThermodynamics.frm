VERSION 5.00
Begin VB.Form frmThermodynamics 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "热力学计算器"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   Icon            =   "frmThermodynamics.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8655
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox texS2 
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
      Left            =   4200
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox texH2 
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
      Left            =   4200
      TabIndex        =   1
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox texH1 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制到剪切板"
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox texOut 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "frmThermodynamics.frx":1B692
      Top             =   4080
      Width           =   8415
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "计算！"
      Default         =   -1  'True
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox texS1 
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
      Left            =   1200
      TabIndex        =   2
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "生成物"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "反应物"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblS 
      BackStyle       =   0  'Transparent
      Caption         =   "标准熵 J/mol"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblH 
      BackStyle       =   0  'Transparent
      Caption         =   "生成焓 kJ/mol"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   7920
      Picture         =   "frmThermodynamics.frx":1B6AB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2385
      Left            =   -360
      Picture         =   "frmThermodynamics.frx":1CFA3
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   9735
   End
End
Attribute VB_Name = "frmThermodynamics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private InTip1 As String
Private InTip2 As String
Private InTip3 As String
Private InTip4 As String

Private Sub cmdCopy_Click()
    UICopy (texOut)
End Sub

Private Sub cmdCal_Click()
    texOut = calRelixue(texH1, texH2, texS1, texS2)
End Sub

Private Sub Form_Load()
    Dim n As Integer
    InTip1 = "反应物的焓，以空格间隔"
    InTip2 = "生成物的焓，以空格间隔"
    InTip3 = "反应物的熵，以空格间隔"
    InTip4 = "生成物的熵，以空格间隔"
    texH1.Text = InTip1
    texH2.Text = InTip2
    texS1.Text = InTip3
    texS2.Text = InTip4
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

Private Sub texH1_Click()
    If texH1.Text = InTip1 Then
        texH1.Text = ""
        texH1.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texH1_KeyPress(KeyAscii As Integer)
    If texH1.Text = InTip1 Then
        texH1.Text = ""
        texH1.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texH2_Click()
    If texH2.Text = InTip2 Then
        texH2.Text = ""
        texH2.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texH2_KeyPress(KeyAscii As Integer)
    If texH2.Text = InTip2 Then
        texH2.Text = ""
        texH2.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texS1_Click()
    If texS1.Text = InTip3 Then
        texS1.Text = ""
        texS1.ForeColor = RGB(0, 0, 0)
    End If
End Sub
Private Sub texS1_KeyPress(KeyAscii As Integer)
    If texS1.Text = InTip3 Then
        texS1.Text = ""
        texS1.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texS2_Click()
    If texS2.Text = InTip4 Then
        texS2.Text = ""
        texS2.ForeColor = RGB(0, 0, 0)
    End If
End Sub
Private Sub texS2_KeyPress(KeyAscii As Integer)
    If texS2.Text = InTip4 Then
        texS2.Text = ""
        texS2.ForeColor = RGB(0, 0, 0)
    End If
End Sub
