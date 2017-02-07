VERSION 5.00
Begin VB.Form frmpH 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "À·ºÓº∆À„"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   Icon            =   "frmpH.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   5760
   StartUpPosition =   1  'À˘”–’ﬂ÷––ƒ
   Begin VB.ComboBox comboAB 
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "frmpH.frx":1B692
      Left            =   120
      List            =   "frmpH.frx":1B69C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox texpKw 
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Text            =   "14"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox texc 
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
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
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "∏¥÷∆µΩºÙ«–∞Â"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   5520
      Width           =   5415
   End
   Begin VB.TextBox texpHOut 
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmpH.frx":1B6A8
      Top             =   3480
      Width           =   5415
   End
   Begin VB.CommandButton cmdpH 
      Caption         =   "º∆À„£°"
      Default         =   -1  'True
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox texpKa 
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label lblpKa 
      BackStyle       =   0  'Transparent
      Caption         =   "pKa"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblc 
      BackStyle       =   0  'Transparent
      Caption         =   "∑÷Œˆ≈®∂»"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblpKw 
      BackStyle       =   0  'Transparent
      Caption         =   "pKw"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.Image imgClose 
      Height          =   450
      Left            =   5280
      Picture         =   "frmpH.frx":1B6BF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgTitle 
      Appearance      =   0  'Flat
      Height          =   2145
      Left            =   -600
      Picture         =   "frmpH.frx":1CFB7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmpH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private InTip As String
Private InTipb As String

Private Sub cmdCopy_Click()
    UICopy (texpHOut)
End Sub

Private Sub cmdpH_Click()
    Dim AorB As Boolean, pHOutHtml As String
    If comboAB = "À·" Then AorB = True Else AorB = False
    pHOutHtml = calpHOut(texpKa, texc, texpKw, AorB)
    texpHOut = CutHtml(pHOutHtml)
    Hisc = texc
    HispKa = texpKa
    HispKw = texpKw
    HisAB = AorB
    HisAcidOutput = CStr(texpHOut)
    Call dataBaseWrite(DataUsername, "c", Hisc)
    Call dataBaseWrite(DataUsername, "pKa", HispKa)
    Call dataBaseWrite(DataUsername, "pKw", HispKw)
    Call dataBaseWrite(DataUsername, "AcidOutput", texpHOut)
    If HisAB Then Call dataBaseWrite(DataUsername, "AB", "A") Else Call dataBaseWrite(DataUsername, "AB", "B")
    If Not DataUsername = "∑√øÕ" Then
        Call dataHtmlChange("pKw", CStr(HispKw))
        Call dataHtmlChange("historyAcidOutput", CStr(pHOutHtml))
    End If
End Sub

Private Sub comboAB_click()
    If comboAB = "À·" Then
        lblpKa = "pKa"
        InTip = "«Î ‰»ÎpKa£¨“‘ø’∏Òº‰∏Ù"
        texpKa.texT = InTip
    Else
        lblpKa = "pKb"
        InTip = "«Î ‰»ÎpKb£¨“‘ø’∏Òº‰∏Ù"
        texpKa.texT = InTip
    End If
End Sub

Private Sub Form_Load()
    Dim n As Integer
    InTip = "«Î ‰»ÎpKa£¨“‘ø’∏Òº‰∏Ù"
    InTipb = "«Î ‰»Î∑÷Œˆ≈®∂»"
    texpKa.texT = InTip
    texc.texT = InTipb
    comboAB.ListIndex = 0
    texpKw = HispKw
    If HisAcidOutput <> "" Then
        texpHOut = HisAcidOutput
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

Private Sub texpKa_Click()
    If texpKa.texT = InTip Then
        texpKa.texT = ""
        texpKa.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texpKa_KeyPress(KeyAscii As Integer)
    If texpKa.texT = InTip Then
        texpKa.texT = ""
        texpKa.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texc_Click()
    If texc.texT = InTipb Then
        texc.texT = ""
        texc.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub texc_KeyPress(KeyAscii As Integer)
    If texc.texT = InTipb Then
        texc.texT = ""
        texc.ForeColor = RGB(0, 0, 0)
    End If
End Sub
