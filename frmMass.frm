VERSION 5.00
Begin VB.Form frmMass 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   7950
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox texMassOut 
      Height          =   3135
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMass.frx":0000
      Top             =   960
      Width           =   6855
   End
   Begin VB.CommandButton cmdMass 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox texMassIn 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmMass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMass_Click()
    texMassOut = calMassPerStr(texMassIn)
End Sub

