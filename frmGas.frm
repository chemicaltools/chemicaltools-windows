VERSION 5.00
Begin VB.Form frmGas 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "气体计算器"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   Icon            =   "frmGas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   5760
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton opnT 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   375
   End
   Begin VB.OptionButton opnn 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   375
   End
   Begin VB.OptionButton opnV 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   375
   End
   Begin VB.OptionButton opnp 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.TextBox texT 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox texn 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox texp 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CommandButton cmdpH 
      Caption         =   "勾选需要计算的量，并点击计算！"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   5055
   End
   Begin VB.TextBox texV 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label lblT 
      BackStyle       =   0  'Transparent
      Caption         =   "温度T"
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
      Left            =   480
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lbln 
      BackStyle       =   0  'Transparent
      Caption         =   "物质的量n"
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
      Left            =   480
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblp 
      BackStyle       =   0  'Transparent
      Caption         =   "压强p"
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
      Left            =   480
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblV 
      BackStyle       =   0  'Transparent
      Caption         =   "体积V"
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
      Left            =   480
      TabIndex        =   9
      Top             =   2880
      Width           =   855
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   5280
      Picture         =   "frmGas.frx":1B692
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   -600
      Picture         =   "frmGas.frx":1CF8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCopy_Click()
    UICopy (texpHOut)
End Sub

Private Sub cmdpH_Click()
    Dim p As Double, v As Double, n As Double, T As Double
    p = Val(texp)
    v = Val(texV)
    n = Val(texn)
    T = Val(texT)
    If opnp Then
        texp = calGasp(v, n, T)
    ElseIf opnV Then
        texV = calGasV(p, n, T)
    ElseIf opnn Then
        texn = calGasn(p, v, T)
    Else
        texT = calGasT(p, v, n)
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
