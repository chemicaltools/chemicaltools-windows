VERSION 5.00
Begin VB.Form FrmChangePassword 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "�޸�����"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   Icon            =   "FrmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   4560
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPasswordChange 
      Caption         =   "�޸�����"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox texAgain 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3960
      Width           =   2895
   End
   Begin VB.TextBox texNewPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox texPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label lblAgain 
      BackStyle       =   0  'Transparent
      Caption         =   "ȷ��������"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblNewPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� ��"
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
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "ԭ �� ��"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   4080
      Picture         =   "FrmChangePassword.frx":1B692
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   -1440
      Picture         =   "FrmChangePassword.frx":1CF8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "FrmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPasswordChange_Click()
    If texNewPassword = texAgain Then
        If dataChangePassword(DataUsername, texPassword, texNewPassword) Then
            MsgBox "�޸ĳɹ���", vbOKOnly, "�ɹ�"
            Me.Hide
            If UIFormLoad(frmAbout) Then Unload frmAbout
            If UIFormLoad(frmElement) Then Unload frmElement
            If UIFormLoad(frmExam) Then Unload frmExam
            If UIFormLoad(frmLogin) Then Unload frmLogin
            If UIFormLoad(frmMain) Then Unload frmMain
            If UIFormLoad(frmMass) Then Unload frmMass
            If UIFormLoad(frmOptions) Then Unload frmOptions
            Load frmLogin
            Unload Me
        Else
            MsgBox "�������", vbOKOnly, "����"
        End If
    Else
        MsgBox "�����������벻һ�£�", vbOKOnly, "����"
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

