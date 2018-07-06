VERSION 5.00
Begin VB.Form frmSignUp 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4875
   FillColor       =   &H00C0FFFF&
   ForeColor       =   &H00C0FFFF&
   Icon            =   "frmSignUp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   4875
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSignUp 
      Caption         =   "ע�ᣡ"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CheckBox chkAgree 
      BackColor       =   &H00C0FFFF&
      Caption         =   "��ͬ�����ع����йط��ɷ��档"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox texAgain 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox texUsername 
      Height          =   375
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   1
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox texPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   4440
      Picture         =   "frmSignUp.frx":1B692
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblAgain 
      BackStyle       =   0  'Transparent
      Caption         =   "ȷ������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "��   ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "��   ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   -1320
      Picture         =   "frmSignUp.frx":1CF8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmSignUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAgree_Click()
    If chkAgree.Value = 0 Then
        cmdSignUp.Enabled = False
    Else
        cmdSignUp.Enabled = True
    End If
End Sub

Private Sub cmdSignUp_Click()
    If texUsername = "" Then
        MsgBox "�������û�����", vbOKOnly, "����"
    ElseIf texusename = "�ÿ�" Then
        MsgBox "��ֹʹ�ø��û�����", vbOKOnly, "����"
    ElseIf texPassword = "" Then
        MsgBox "���������룡", vbOKOnly, "����"
    ElseIf texPassword <> texAgain Then
        MsgBox "�����������벻һ�£�", vbOKOnly, "����"
    ElseIf dataSignUp(texUsername, texPassword) = False Then
        MsgBox "�Ѵ��ڴ��û�����", vbOKOnly, "����"
    Else
        MsgBox "ע��ɹ���", vbOKOnly, "ע��ɹ�"
        Me.Hide
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub imgClose_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
