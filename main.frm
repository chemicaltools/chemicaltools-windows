VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8445
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1 = calMassPerStr(Text1)
End Sub

Private Sub Form_Load()
dataElement
End Sub
