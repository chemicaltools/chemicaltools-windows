VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "化学小工具 Designed by 团队一号"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   4875
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdfrmPhy 
      Caption         =   "动力学、热力学计算器"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdfrmAbout 
      Caption         =   "关于"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdfrmpH 
      Caption         =   "pH计算器"
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdfrmGas 
      Caption         =   "气体状态计算器"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdfrmEquation 
      Caption         =   "配平方程式"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdfrmExam 
      Caption         =   "元素记忆测试"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdfrmElement 
      Caption         =   "元素查询器"
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdfrmMass 
      Caption         =   "质量计算器"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
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

