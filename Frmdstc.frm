VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmDsTC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch phiÕu thu chi"
   ClientHeight    =   7080
   ClientLeft      =   255
   ClientTop       =   390
   ClientWidth     =   9855
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frmdstc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cash Voucher List"
   Begin VB.CommandButton Command 
      BackColor       =   &H8000000E&
      Caption         =   "Vµo &sæ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Xem b¸o c¸o"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtshkh 
      Height          =   300
      Left            =   1200
      LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
      MaxLength       =   20
      TabIndex        =   9
      Tag             =   "0"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdkh 
      Height          =   375
      Left            =   2640
      Picture         =   "Frmdstc.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1640
      Width           =   375
   End
   Begin VB.ComboBox CboThang 
      Height          =   315
      Index           =   2
      ItemData        =   "Frmdstc.frx":5C5C
      Left            =   8880
      List            =   "Frmdstc.frx":5C5E
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H8000000E&
      Height          =   375
      Index           =   4
      Left            =   3600
      Picture         =   "Frmdstc.frx":5C60
      Style           =   1  'Graphical
      TabIndex        =   38
      Tag             =   "&View"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   4800
      Picture         =   "Frmdstc.frx":6DD2
      Style           =   1  'Graphical
      TabIndex        =   37
      Tag             =   "&Print"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.ComboBox CboNguon 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
      Top             =   1280
      Width           =   3135
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   975
      Index           =   0
      Left            =   5520
      TabIndex        =   22
      Top             =   960
      Width           =   1575
      Begin VB.OptionButton SSOpt 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ngµy GS, Sè hiÖu"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   25
         Tag             =   "B. Date, V. Code"
         Top             =   0
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton SSOpt 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ngµy CT, Sè hiÖu"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   24
         Tag             =   "V. Date, V. Code"
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton SSOpt 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sè hiÖu CT"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   23
         Tag             =   "V. Code"
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.OptionButton OptLK 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tõ ngµy"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Tag             =   "From"
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton OptLK 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tõ th¸ng"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "From"
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   8400
      Picture         =   "Frmdstc.frx":8234
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   7200
      Picture         =   "Frmdstc.frx":9656
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "&Select"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   6000
      Picture         =   "Frmdstc.frx":AAB8
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "&List"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.ComboBox CboThang 
      Height          =   315
      Index           =   1
      ItemData        =   "Frmdstc.frx":BE82
      Left            =   3240
      List            =   "Frmdstc.frx":BE84
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PhiÕu chi"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Tag             =   "Money out"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PhiÕu thu"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Tag             =   "Money in"
      Top             =   960
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ComboBox CboThang 
      Height          =   315
      Index           =   0
      ItemData        =   "Frmdstc.frx":BE86
      Left            =   1200
      List            =   "Frmdstc.frx":BE88
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSGrid.Grid GrdChungtu 
      Height          =   3975
      Left            =   120
      TabIndex        =   11
      Tag             =   "20"
      Top             =   2520
      Width           =   9615
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   7011
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   20
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      MouseIcon       =   "Frmdstc.frx":BE8A
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label L 
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   40
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "§èi t­îng"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   44
      Tag             =   "Object"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lbkh 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3120
      TabIndex        =   43
      Tag             =   "1"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tµi kho¶n"
      Height          =   255
      Index           =   16
      Left            =   8040
      TabIndex        =   41
      Tag             =   "Account"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sè phiÕu"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   39
      Tag             =   "Number of Voucher"
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label L 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   2
      Left            =   8160
      TabIndex        =   36
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Chªnh lÖch"
      Height          =   255
      Index           =   14
      Left            =   7320
      TabIndex        =   35
      Tag             =   "Result"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label L 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   34
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label L 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00808000&
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   33
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tæng chi"
      Height          =   255
      Index           =   13
      Left            =   7320
      TabIndex        =   32
      Tag             =   "Expense"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tæng thu"
      Height          =   255
      Index           =   12
      Left            =   7320
      TabIndex        =   31
      Tag             =   "Income"
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chi"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   8160
      TabIndex        =   30
      Tag             =   "Expense"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bé phËn"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   29
      Tag             =   "Index"
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thu"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   6840
      TabIndex        =   28
      Tag             =   "Income"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lý do thu/chi"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   5160
      TabIndex        =   27
      Tag             =   "Reason"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§Þa chØ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   26
      Tag             =   "Address"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "®Õn ngµy"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   21
      Tag             =   "to"
      Top             =   480
      Width           =   735
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   120
      X2              =   120
      Y1              =   2160
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   2
      X1              =   9720
      X2              =   9720
      Y1              =   2160
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   9720
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   9720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "®Õn th¸ng"
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   16
      Tag             =   "to"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Thø tù"
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   15
      Tag             =   "Sort by"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hä tªn"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   14
      Tag             =   "Name"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ngµy CT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   13
      Tag             =   "V. Date"
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè CT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Tag             =   "V. Code"
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "FrmDsTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MaCTChon As Long
Dim LietKe As Boolean
Dim Row As Integer
Dim ord As Integer
Dim ngay(0 To 1) As Date
Dim Lsql As String
Dim kh As New ClsKhachHang



'======================================================================================
' LiÖt kª, Chän chøng tõ
'======================================================================================
Private Sub Command_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0:
            If CboThang(1).ListIndex < CboThang(0).ListIndex Then CboThang(1).ListIndex = CboThang(0).ListIndex
            LietKeChungtu
        Case 1:
            GrdChungtu.col = 7
            If Len(GrdChungtu.Text) = 0 Then Exit Sub
            MaCTChon = CLng5(GrdChungtu.Text)
            LietKe = True
            Hide
        Case 2:
            MaCTChon = 0
            LietKe = True
            Hide
        Case 3, 4:
            If CLng5(L(3).Caption) = 0 Then Exit Sub
            If OptLK(1).Value And CboThang(2).Text = "111" Then
                Dim taikhoan As New ClsTaikhoan
                taikhoan.InitTaikhoanSohieu "111"
                SetRptInfo
                InSocaiTk taikhoan, 0, 0, ngay(0), ngay(1), True, "", 0, 0
                Set taikhoan = Nothing
            Else
                InBangKe
            End If
            frmMain.Rpt.Destination = 4 - Index
            InBaoCaoRPT
        Case 5:
            Dim f As Form, ms As Long
            Set f = New FrmChungtu
            Load f
            With GrdChungtu
                .col = 7
                For i = .SelStartRow To .SelEndRow
                    .Row = i
                    ms = CLng5(.Text)
                    If ms > 0 Then
                        pPhieu = 0
                        f.VaoSoNK CLng5(GrdChungtu.Text)
                    End If
                Next
            End With
            Set f = Nothing
            pPhieu = 1
            Command_Click 0
    End Select
    GrdChungtu.RemoveItem (0)
End Sub

Private Sub Form_Activate()
    If LietKe Then
        Me.Refresh
        Command_Click 0
        LietKe = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
            Select Case KeyCode
                Case vbKeyL:
                    RFocus Command(0)
                    Command_Click 0
                Case vbKeyC:
                    RFocus Command(1)
                    Command_Click 1
                Case vbKeyV:
                    RFocus Command(2)
                    Command_Click 2
            End Select
    End If
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyA Then GridSelAll GrdChungtu
    If KeyCode = vbKeyEscape Then Command_Click 2
End Sub
'======================================================================================
' Khëi t¹o cöa sæ
'======================================================================================
Private Sub Form_Load()
    Dim chi_so As Integer
    
    ColumnSetUp GrdChungtu, 0, 1300, 2
    ColumnSetUp GrdChungtu, 1, 820, 2
    ColumnSetUp GrdChungtu, 2, 1420, 0
    ColumnSetUp GrdChungtu, 3, 1420, 0
    ColumnSetUp GrdChungtu, 4, 1660, 0
    ColumnSetUp GrdChungtu, 5, 1300, 1
    ColumnSetUp GrdChungtu, 6, 1300, 1
    ColumnSetUp GrdChungtu, 7, 1, 0
    
    
    AddMonthToCbo CboThang(0)
    AddMonthToCbo CboThang(1)
'    CboThang(2).ListIndex = 0
    
    For chi_so = 0 To 1
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next
    
    Int_RecsetToCbo "SELECT CLng(LEFT(SoHieu,3)) As F2,LEFT(SoHieu,3) As F1 FROM HethongTK WHERE LEFT(SoHieu,3)='111' OR LEFT(SoHieu,3)='112' OR TK_ID2=112 GROUP BY LEFT(SoHieu,3)", CboThang(2)
    Int_RecsetToCbo "SELECT MaSo As F2,DienGiai As F1 FROM DoituongCT ORDER BY DienGiai", CboNguon
    
    MaCTChon = 0
    GrdChungtu.Row = 0
    GrdChungtu_Click
    LietKe = False

    Caption = Caption + " - " + CStr(pNamTC)
    Command(5).Visible = (pPhieu > 0)
    ord = 0
    SetFont Me
LoiNgay:
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set kh = Nothing
End Sub

'======================================================================================
' Nhí m· sè CT ®­îc chän
'======================================================================================
Private Sub GrdChungtu_Click()
    SendKeys "{Home}", True
    Row = GrdChungtu.Row
    SetGridIndex GrdChungtu, Row
    With GrdChungtu
        .col = 7
        If Len(.Text) = 0 Then MaCTChon = 0 Else MaCTChon = CLng5(.Text)
        .col = 0
    End With
End Sub

Private Sub GrdChungtu_KeyPress(KeyAscii As Integer)
    GrdChungtu_Click
    If KeyAscii = 13 Then Command_Click 1
End Sub

Private Sub GrdChungtu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdChungtu, GrdChungtu.col
    End If
End Sub

Private Sub SSOpt_Click(Index As Integer)
    ord = Index
End Sub
'======================================================================================
' LiÖt kª chøng tõ
'======================================================================================
Private Sub LietKeChungtu()
    Dim i As Integer, hoten As String, DiaChi As String, ovr As Integer
    Dim rs_chungtu As Recordset, no As Double, co As Double, sh As String
    
    If ChkLoai(0).Value = 0 And ChkLoai(1).Value = 0 Then Exit Sub
    Me.MousePointer = 11
    
    i = IIf(CboThang(2).Text = "111", 0, 3)
    sh = IIf(pPhieu > 0, "P", "")
    
    Lsql = "SELECT DISTINCTROW Chungtu" + sh + ".MaCT, ChungTu" + sh + ".SoHieu, NgayCT, Sum(SoPS) AS TSoPS,Chungtu" + sh + ".DienGiai" + IIf(pNN = 1, "E", "") + " AS DG,ChungTuLQ" + sh + ".HoTen,ChungTuLQ" + sh + ".DiaChi,ChungTuLQ" + sh + ".SoCTGoc,MaDT,First(HethongTK.SoHieu) AS SHNo,First(TK.SoHieu) AS SHCo,First(DoituongCT.DienGiai) AS BP FROM (((ChungTu" + sh + " LEFT JOIN HethongTK ON Chungtu" + sh + ".MaTKTCNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON Chungtu" + sh + ".MaTKTCCo=TK.MaSo) INNER JOIN DoituongCT ON Chungtu" + sh + ".MaDT=DoituongCT.MaSo) LEFT JOIN ChungTuLQ" + sh + " ON (ChungTu" + sh + ".MaCT=ChungTuLQ" + sh + ".MaCT AND ChungTuLQ" + sh + ".Loai=" + CStr(i) + ") WHERE "
    
    If ChkLoai(0).Value = 1 And ChkLoai(1).Value = 1 Then
        Lsql = Lsql + "(LEFT(HethongTK.SoHieu,3)='" + CboThang(2).Text + "' OR LEFT(TK.SoHieu,3)='" + CboThang(2).Text + "')"
    Else
        If ChkLoai(0).Value = 1 Then Lsql = Lsql + "(LEFT(HethongTK.SoHieu,3)='" + CboThang(2).Text + "')" Else Lsql = Lsql + "(LEFT(TK.SoHieu,3)='" + CboThang(2).Text + "')"
    End If
    
    If CboNguon.ItemData(CboNguon.ListIndex) > 1 Then Lsql = Lsql + " AND ChungTu" + sh + ".MaDT=" + CStr(CboNguon.ItemData(CboNguon.ListIndex))
    
    If OptLK(0).Value Then
        Lsql = Lsql + " AND " + WThang("ThangCT", CboThang(0).ItemData(CboThang(0).ListIndex), CboThang(1).ItemData(CboThang(1).ListIndex))
    Else
        Lsql = Lsql + " AND " + WNgay("NgayGS", ngay(0), ngay(1))
    End If
    
    If kh.MaSo > 0 Then Lsql = Lsql + " AND (ChungTuLQ" + sh + ".MaKH=" + CStr(kh.MaSo) + ")"
    If pProcessMode = 1 Then Lsql = Lsql + " AND XuLy<2 "
    
    Lsql = Lsql + " GROUP BY ChungTu" + sh + ".MaCT,ChungTu" + sh + ".SoHieu,NgayCT,NgayGS,ChungTu" + sh + ".DienGiai" + IIf(pNN = 1, "E", "") + ",ChungTuLQ" + sh + ".HoTen,ChungTuLQ" + sh + ".DiaChi,ChungTuLQ" + sh + ".SoCTGoc,MaDT "
    Select Case ord
        Case 0:
            Lsql = Lsql + "ORDER BY NgayGS DESC,ChungTu" + sh + ".SoHieu DESC"
        Case 1:
            Lsql = Lsql + "ORDER BY NgayCT DESC,ChungTu" + sh + ".SoHieu DESC"
        Case 2:
            Lsql = Lsql + "ORDER BY ChungTu" + sh + ".SoHieu DESC"
    End Select
    
    ClearGrid GrdChungtu, GrdChungtu.tag
    Set rs_chungtu = DBKetoan.OpenRecordset(Lsql, dbOpenSnapshot)
    Do While Not rs_chungtu.EOF
        If Left(rs_chungtu!shno, 3) = CboThang(2).Text Then no = no + rs_chungtu!Tsops
        If Left(rs_chungtu!shco, 3) = CboThang(2).Text Then co = co + rs_chungtu!Tsops
        If IsNull(rs_chungtu!hoten) Then hoten = "..." Else hoten = rs_chungtu!hoten
        If IsNull(rs_chungtu!DiaChi) Then DiaChi = "..." Else DiaChi = rs_chungtu!DiaChi
        If GrdChungtu.Rows < MaxGridRow Then
            GrdChungtu.AddItem rs_chungtu!sohieu + Chr(9) + Format(rs_chungtu!NgayCT, "dd/mm") + Chr(9) _
                + hoten + Chr(9) + DiaChi + Chr(9) + rs_chungtu!dg + Chr(9) + IIf(Left(rs_chungtu!shno, 3) = CboThang(2).Text, Format(rs_chungtu!Tsops, Mask_0), "") _
                + Chr(9) + IIf(Left(rs_chungtu!shco, 3) = CboThang(2).Text, Format(rs_chungtu!Tsops, Mask_0), "") + Chr(9) + CStr(rs_chungtu!MaCT), 0
        Else
            ovr = 1
        End If
        rs_chungtu.MoveNext
    Loop
    If rs_chungtu.RecordCount > 0 Then
        If Row >= rs_chungtu.RecordCount Then Row = rs_chungtu.RecordCount - 1
    Else
        Row = 0
    End If
    GrdChungtu.Rows = IIf(rs_chungtu.RecordCount > GrdChungtu.tag, IIf(rs_chungtu.RecordCount > MaxGridRow, MaxGridRow, rs_chungtu.RecordCount), GrdChungtu.tag)
    GrdChungtu.col = 0
    L(3).Caption = CStr(rs_chungtu.RecordCount)
    rs_chungtu.Close
    Set rs_chungtu = Nothing
    On Error Resume Next
    GrdChungtu.Row = Row
    GrdChungtu_Click
    If Not GrdChungtu.RowIsVisible(Row) Then GrdChungtu.TopRow = Row - 12
    On Error GoTo 0
    L(0).Caption = Format(no, Mask_0)
    L(1).Caption = Format(co, Mask_0)
    L(2).Caption = Format(no - co, Mask_0)
    Me.MousePointer = 0
    If ovr > 0 Then ErrMsg er_NhieuCT
End Sub
'======================================================================================
' Hµm hiÖn cöa sæ danh s¸ch vµ tr¶ vÒ m· sè CT ®­îc chän
'======================================================================================
Public Function ChonCT() As Long
    Show 1
    ChonCT = MaCTChon
End Function

Private Sub MedNgay_GotFocus(Index As Integer)
    AutoSelect MedNgay(Index)
End Sub

Private Sub MedNgay_LostFocus(Index As Integer)
    If IsDate(MedNgay(Index).Text) Then
        ngay(Index) = CDate(MedNgay(Index).Text)
    Else
        RFocus MedNgay(Index)
    End If
End Sub

Private Sub InBangKe()
    If Len(Lsql) = 0 Then Exit Sub

    frmMain.Rpt.ReportFileName = "PHIEUTC.RPT"
    SetRptInfo
    SetSQL "QNhatky", Lsql
    
    If OptLK(0).Value Then
        If CboThang(0).ListIndex <> CboThang(1).ListIndex Then
            frmMain.Rpt.Formulas(3) = "Thoigian=IF PageNumber()=1 THEN 'Tõ th¸ng " + CStr(CboThang(0).ItemData(CboThang(1).ListIndex)) + " ®Õn thang " + CStr(CboThang(0).ItemData(CboThang(1).ListIndex)) + "'"
        Else
            frmMain.Rpt.Formulas(3) = "Thoigian=IF PageNumber()=1 THEN 'Th¸ng " + CStr(CboThang(0).ItemData(CboThang(1).ListIndex)) + "'"
        End If
        RptSetDate NgayCuoiThang(pNamTC, CboThang(0).ItemData(CboThang(1).ListIndex))
    Else
        If ngay(0) <> ngay(1) Then
            frmMain.Rpt.Formulas(3) = "Thoigian='IF PageNumber()=1 THEN Tõ ngµy " + Format(ngay(0), Mask_DR) + " ®Õn ngµy " + Format(ngay(1), Mask_DR) + "'"
        Else
            frmMain.Rpt.Formulas(3) = "Thoigian='IF PageNumber()=1 THEN Ngµy " + Format(ngay(1), Mask_DR) + "'"
        End If
        RptSetDate ngay(1)
    End If
    frmMain.Rpt.Formulas(4) = "BoPhan='" + IIf(CboNguon.ItemData(CboNguon.ListIndex) > 1, CboNguon.Text, "") + "'"
    If kh.MaSo > 0 Then
        frmMain.Rpt.Formulas(5) = "TenKH='" + kh.sohieu + " - " + kh.Ten + "'"
    End If
End Sub

Private Sub txtshkh_GotFocus()
    AutoSelect txtshkh
End Sub

Private Sub txtshkh_LostFocus()
    Dim mvt As Long
    
    kh.InitKhachHangSohieu txtshkh
    lbkh.Caption = kh.Ten
    HienThongBao Me.Caption, 1
End Sub

Private Sub cmdkh_Click()
    Me.MousePointer = 11
    txtshkh.Text = FrmKhachHang.ChonKhachHang(txtshkh.Text)
    Me.MousePointer = 0
    RFocus txtshkh
End Sub

