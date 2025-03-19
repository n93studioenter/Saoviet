VERSION 5.00
Begin VB.Form FrmThKC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Thùc hiÖn kÕt chuyÓn"
   ClientHeight    =   6210
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmThKC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Monthly Conversion"
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2160
      LinkItem        =   "Sè hiÖu chøng tõ"
      MaxLength       =   20
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Cbo 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "FrmThKC.frx":57E2
      Left            =   600
      List            =   "FrmThKC.frx":57E4
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   32
      Left            =   7080
      TabIndex        =   52
      Tag             =   "0"
      Top             =   480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   33
      Left            =   7080
      TabIndex        =   51
      Tag             =   "0"
      Top             =   840
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   34
      Left            =   7080
      TabIndex        =   50
      Tag             =   "0"
      Top             =   1200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   35
      Left            =   7080
      TabIndex        =   49
      Tag             =   "0"
      Top             =   1560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   36
      Left            =   7080
      TabIndex        =   48
      Tag             =   "0"
      Top             =   1920
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   37
      Left            =   7080
      TabIndex        =   47
      Tag             =   "0"
      Top             =   2280
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   38
      Left            =   7080
      TabIndex        =   46
      Tag             =   "0"
      Top             =   2640
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   39
      Left            =   7080
      TabIndex        =   45
      Tag             =   "0"
      Top             =   3000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   40
      Left            =   7080
      TabIndex        =   44
      Tag             =   "0"
      Top             =   3360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   41
      Left            =   7080
      TabIndex        =   43
      Tag             =   "0"
      Top             =   3720
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   42
      Left            =   7080
      TabIndex        =   42
      Tag             =   "0"
      Top             =   4080
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   43
      Left            =   7080
      TabIndex        =   41
      Tag             =   "0"
      Top             =   4440
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   44
      Left            =   7080
      TabIndex        =   40
      Tag             =   "0"
      Top             =   4800
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   45
      Left            =   7080
      TabIndex        =   39
      Tag             =   "0"
      Top             =   5160
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   46
      Left            =   7080
      TabIndex        =   38
      Tag             =   "0"
      Top             =   5520
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H80000013&
      Caption         =   "KC"
      Height          =   255
      Index           =   47
      Left            =   7080
      TabIndex        =   37
      Tag             =   "0"
      Top             =   5880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   16
      Left            =   3480
      TabIndex        =   36
      Tag             =   "0"
      Top             =   480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   17
      Left            =   3480
      TabIndex        =   35
      Tag             =   "0"
      Top             =   840
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   18
      Left            =   3480
      TabIndex        =   34
      Tag             =   "0"
      Top             =   1200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   19
      Left            =   3480
      TabIndex        =   33
      Tag             =   "0"
      Top             =   1560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   20
      Left            =   3480
      TabIndex        =   32
      Tag             =   "0"
      Top             =   1920
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   21
      Left            =   3480
      TabIndex        =   31
      Tag             =   "0"
      Top             =   2280
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   22
      Left            =   3480
      TabIndex        =   30
      Tag             =   "0"
      Top             =   2640
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   23
      Left            =   3480
      TabIndex        =   29
      Tag             =   "0"
      Top             =   3000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   24
      Left            =   3480
      TabIndex        =   28
      Tag             =   "0"
      Top             =   3360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   25
      Left            =   3480
      TabIndex        =   27
      Tag             =   "0"
      Top             =   3720
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   26
      Left            =   3480
      TabIndex        =   26
      Tag             =   "0"
      Top             =   4080
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   27
      Left            =   3480
      TabIndex        =   25
      Tag             =   "0"
      Top             =   4440
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   28
      Left            =   3480
      TabIndex        =   24
      Tag             =   "0"
      Top             =   4800
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   29
      Left            =   3480
      TabIndex        =   23
      Tag             =   "0"
      Top             =   5160
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   30
      Left            =   3480
      TabIndex        =   22
      Tag             =   "0"
      Top             =   5520
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   31
      Left            =   3480
      TabIndex        =   21
      Tag             =   "0"
      Top             =   5880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   2040
      Picture         =   "FrmThKC.frx":57E6
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "&Done"
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   3480
      Picture         =   "FrmThKC.frx":6C14
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "&Return"
      Top             =   0
      Width           =   1095
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   17
      Tag             =   "0"
      Top             =   5880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   16
      Tag             =   "0"
      Top             =   5520
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   15
      Tag             =   "0"
      Top             =   5160
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   14
      Tag             =   "0"
      Top             =   4800
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   13
      Tag             =   "0"
      Top             =   4440
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   12
      Tag             =   "0"
      Top             =   4080
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   11
      Tag             =   "0"
      Top             =   3720
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Tag             =   "0"
      Top             =   3360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Tag             =   "0"
      Top             =   3000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Tag             =   "0"
      Top             =   2640
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Tag             =   "0"
      Top             =   2280
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Tag             =   "0"
      Top             =   1920
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Tag             =   "0"
      Top             =   1560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Tag             =   "0"
      Top             =   1200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Tag             =   "0"
      Top             =   840
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox ChkKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KC"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Tag             =   "0"
      Top             =   480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label LbKC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tû gi¸"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   53
      Tag             =   "Ex. Rate"
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LbKC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th¸ng"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Tag             =   "Month"
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmThKC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_Click(Index As Integer)
    Dim i As Integer
    Me.MousePointer = 11
    Select Case Index
        Case 0:
            For i = 0 To MaxKC - 1
                If ChkKC(i).tag = 0 Then GoTo KT
                If ChkKC(i).Value = 1 Then KCTuDong Cbo.ItemData(Cbo.ListIndex), ChkKC(i).tag, Cdbl5(txt.Text)
            Next
    End Select
KT:
    Me.MousePointer = 0
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not ChkKC(0).Visible Then Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyT:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyV:
                RFocus Command(1)
                Command_Click 1
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rs As Recordset
    
    AddMonthToCbo Cbo
    Set rs = DBKetoan.OpenRecordset("SELECT * FROM CTKetChuyen ORDER BY STT", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        ChkKC(i).Caption = rs!diengiai
        ChkKC(i).tag = rs!MaSo
        ChkKC(i).FontName = pFontName
        ChkKC(i).FontSize = pFontSize
        ChkKC(i).Visible = True
        ChkKC(i).Value = 1
        i = i + 1
        If i >= MaxKC Then GoTo KT
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    LbKC(1).Visible = (pTygia > 0)
    txt.Visible = (pTygia > 0)
    If pTygia > 0 Then txt.Text = Format(TyGiaCuoi, Mask_0)
    
    SetFont Me
    
KT:
    If i > 0 Then Me.Height = ChkKC(IIf(i > 16, 15, i - 1)).Top + 780
    Me.Width = Me.Width + (Fix(i / 16) + IIf(i Mod 16 > 0, 0, -1)) * 3495
    Command(0).Left = Me.Width - 1335
    Command(1).Left = Me.Width - 1335
End Sub

Private Sub txt_GotFocus()
    AutoSelect txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    KeyProcess txt, KeyAscii
End Sub
