VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmTongHop 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tæng hîp sè liÖu"
   ClientHeight    =   3825
   ClientLeft      =   855
   ClientTop       =   2415
   ClientWidth     =   9885
   ClipControls    =   0   'False
   Icon            =   "FrmTongHop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Import data form other file"
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chän chøng tõ "
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   21
      Tag             =   "Create sub-account"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "T¸ch riªng c¸c ®èi t­îng cïng sè hiÖu"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   20
      Tag             =   "Create sub-account"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "T¹o chi tiÕt cho c¸c tµi kho¶n"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   18
      Tag             =   "Create sub-account"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.ComboBox CboThang 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      ItemData        =   "FrmTongHop.frx":57E2
      Left            =   4320
      List            =   "FrmTongHop.frx":57E4
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   1
      Left            =   3480
      TabIndex        =   11
      Top             =   3120
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
      Index           =   0
      Left            =   1800
      TabIndex        =   10
      Top             =   3120
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
   Begin VB.OptionButton Opt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tæng hîp tõ ngµy"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Tag             =   "Import from"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.ComboBox CboThang 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      ItemData        =   "FrmTongHop.frx":57E6
      Left            =   2280
      List            =   "FrmTongHop.frx":57E8
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th¸ng tËp hîp ph¸t sinh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Tag             =   "Import data From"
      Top             =   2760
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tµi s¶n"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   14
      Tag             =   "Fixed Assets"
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C«ng nî"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   13
      Tag             =   "Liability"
      Top             =   3480
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "VËt t­"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   12
      Tag             =   "Inventory"
      Top             =   3480
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   8760
      Picture         =   "FrmTongHop.frx":57EA
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Return"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8760
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FrmTongHop.frx":6C0C
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Import"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8760
      Picture         =   "FrmTongHop.frx":6F4E
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "&Delete"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8760
      Picture         =   "FrmTongHop.frx":8430
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&Add"
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox LstDB 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Index           =   1
      ItemData        =   "FrmTongHop.frx":998A
      Left            =   4320
      List            =   "FrmTongHop.frx":998C
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
   Begin VB.ListBox LstDB 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Index           =   0
      ItemData        =   "FrmTongHop.frx":998E
      Left            =   0
      List            =   "FrmTongHop.frx":9990
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "(Ctrl-A ®Ó chän toµn bé)"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   6960
      TabIndex        =   19
      Tag             =   "(Ctrl-A: select all data files)"
      ToolTipText     =   "0"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "®Õn th¸ng"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   17
      Tag             =   "to"
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "®Õn ngµy"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2745
      TabIndex        =   16
      Tag             =   "to"
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tæng hîp c¸c chi tiÕt"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Tag             =   "Include"
      Top             =   3480
      Width           =   1935
   End
End
Attribute VB_Name = "FrmTongHop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const maxfile = 20

Dim ngay(0 To 1) As Date
Dim pTK As Integer
Dim pCT As String
Dim MaCT(1 To Max1000) As Long
Dim SoCTChon As Integer
Dim db As Database

Private Sub cboThang_Click(Index As Integer)
    Chk(5).Visible = (LstDB(0).ListCount = 1)
End Sub

Private Sub Command_Click(Index As Integer)
    Dim i As Integer, j As Integer, fn As String
    
    Select Case Index
        Case 0:
            fn = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
            If Len(fn) > 0 Then
                For i = 0 To LstDB(1).ListCount - 1
                    If LstDB(1).List(i) = fn Then
                        MsgBox "TÖp d÷ liÖu ®· ®­îc chän !", vbExclamation, App.ProductName
                        Exit Sub
                    End If
                Next
                ThemDB fn, True
            End If
        Case 1:
            If LstDB(0).ListIndex >= 0 Then
                LstDB(1).RemoveItem LstDB(0).ListIndex
                LstDB(0).RemoveItem LstDB(0).ListIndex
            End If
        Case 2:
            Me.MousePointer = 11
            pTK = Chk(3).Value
            If Opt(0).Value And CboThang(0).ListIndex > CboThang(1).ListIndex Then CboThang(1).ListIndex = CboThang(0).ListIndex
            If Chk(5).Value = 1 Then
                SoCTChon = FrmA.ChonCTSoA(LstDB(1).List(0), IIf(Opt(0).Value, CboThang(1).ItemData(CboThang(1).ListIndex), 0), ngay(0), ngay(1), MaCT())
                Me.Refresh
                If SoCTChon = 0 Then GoTo KT
            End If
            j = 0
            For i = 0 To LstDB(0).ListCount - 1
                If LstDB(0).Selected(i) Then
                    On Error GoTo KT1
                    Set db = WSpace.OpenDatabase(LstDB(1).List(i), False, False, ";PWD=" + pPSW)
                    On Error Resume Next
                    If NhapPS(db, IIf(Opt(0).Value, CboThang(0).ItemData(CboThang(0).ListIndex), 0), IIf(Opt(0).Value, CboThang(0).ItemData(CboThang(1).ListIndex), 0), Chk(0).Value, Chk(1).Value, Chk(2).Value, i + 1, ngay(0), ngay(1)) Then
                        If (Opt(0).Value And CboThang(0).ItemData(CboThang(0).ListIndex) = pThangDauKy) Or (Opt(1).Value And ngay(0) = NgayDauThang(pNamTC, pThangDauKy)) Then
                            CBDL CStr(LstDB(1).ItemData(i))
                            CongDK db, CStr(LstDB(1).ItemData(i)), LstDB(0).List(i), 0, pTK, pCT
                            If Chk(0).Value = 1 Then CongDKVT db, CStr(LstDB(1).ItemData(i)), LstDB(0).List(i), 0, pTK, pCT
                            If Chk(1).Value = 1 Then CongDKCN db, CStr(LstDB(1).ItemData(i)), LstDB(0).List(i), 0, pTK, pCT
                        End If
                        j = j + 1
                        HienThongBao "§· tæng hîp " + LstDB(0).List(i), 1
                    Else
                        MsgBox "HÖ thèng tµi kho¶n kh«ng hîp lÖ!", vbCritical, App.ProductName
                    End If
                    db.Close
KT1:
                End If
            Next
            If j > 0 Then
                If Chk(0).Value = 0 Then ExecuteSQL5 "UPDATE HethongTK SET TK_ID=0 WHERE TK_ID=" + CStr(TKVT_ID) Else KiemTraVatTu
                If Chk(1).Value = 0 Then ExecuteSQL5 "UPDATE HethongTK SET TK_ID=0 WHERE TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID)
                If Chk(2).Value = 1 Then SoDuTKTS
                KiemTraTaiKhoan
                MsgBox "Ph¸t sinh cña " + CStr(j) + " ®¬n vÞ ®· ®­îc tËp hîp vµo sè liÖu cña " + VString(pTenCn)
                CloseItemList
            End If
        Case 3:
            Unload Me
    End Select
    Chk(5).Visible = (LstDB(0).ListCount = 1)
KT:
    Me.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyT:    RFocus Command(0):    Command_Click 0
            Case vbKeyX:    RFocus Command(1):    Command_Click 1
            Case vbKeyH:    RFocus Command(2):    Command_Click 2
            Case vbKeyV:    RFocus Command(3):    Command_Click 3
        End Select
    End If
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        For i = 0 To LstDB(0).ListCount - 1
            LstDB(0).Selected(i) = True
            LstDB(1).Selected(i) = True
        Next
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim chi_so As Integer
    
    Caption = Caption + " - " + CStr(pNamTC)
    
    For chi_so = 0 To 1
        AddMonthToCbo CboThang(chi_so)
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next
    LietKeDB
    
    Chk(5).Visible = (LstDB(0).ListCount = 1)
    
    SetFont Me
LoiNgay:
End Sub

Private Sub LietKeDB()
    Dim i As Integer, FileName As String
    
    For i = 1 To maxfile
        FileName = GetSetting(IniPath, "ImportFile", "File" + CStr(i))
        If Len(FileName) > 0 Then
            If Len(Dir(FileName)) > 0 Then ThemDB FileName, False
        Else
            Exit For
        End If
    Next
End Sub
'=================================================================================================================
' Thñ tôc thªm tÖp d÷ liÖu ®­îc chän vµo danh s¸ch
'=================================================================================================================
Private Sub ThemDB(file_name As String, thongbao As Boolean)
    Dim db As Database
    Dim rs_license As Recordset, i As Integer
    
    On Error GoTo LoiDB
    Set db = WSpace.OpenDatabase(file_name, False, False, ";PWD=" + pPSW)
    Set rs_license = db.OpenRecordset("SELECT DISTINCTROW * FROM License", dbOpenSnapshot)
    For i = 0 To LstDB(1).ListCount - 1
        If LstDB(0).ItemData(i) = rs_license!TenCty_ID And LstDB(1).ItemData(i) = rs_license!tencn_id Then
            If thongbao Then MsgBox "TÖp d÷ liÖu ®· ®­îc chän !", vbExclamation, App.ProductName
            GoTo KetThuc
        End If
        If rs_license!NamTC <> pNamTC Then
            If thongbao Then MsgBox "TÖp d÷ liÖu kh«ng cïng n¨m tµi chÝnh víi tÖp d÷ liÖu hiÖn thêi!", vbExclamation, App.ProductName
            GoTo KetThuc
        End If
    Next
    LstDB(0).AddItem rs_license!tencn
    LstDB(0).ItemData(LstDB(0).NewIndex) = rs_license!TenCty_ID
    LstDB(1).AddItem file_name
    LstDB(1).ItemData(LstDB(1).NewIndex) = rs_license!tencn_id
KetThuc:
    rs_license.Close
    db.Close
    On Error GoTo 0
    Set rs_license = Nothing
    Set db = Nothing
    Exit Sub
LoiDB:
    If thongbao Then ErrMsg er_DBFile
    GoTo KetThuc
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Dim i As Integer
    Set db = Nothing
    For i = 1 To LstDB(0).ListCount
        SaveSetting IniPath, "ImportFile", "File" + CStr(i), LstDB(1).List(i - 1)
    Next
    SaveSetting IniPath, "ImportFile", "File" + CStr(LstDB(0).ListCount + 1), ""
End Sub

Private Sub LstDB_Click(Index As Integer)
    Dim i As Integer
    
    LstDB(1 - Index).ListIndex = LstDB(Index).ListIndex
    For i = 0 To LstDB(Index).ListCount - 1
        LstDB(1 - Index).Selected(i) = LstDB(Index).Selected(i)
    Next
End Sub

Private Sub CBDL(s As String)
    If Not TruongDaCo("HethongTK", "n" + s) Then ThemTruong "HethongTK", "n" + s, dbDouble
    If Not TruongDaCo("HethongTK", "c" + s) Then ThemTruong "HethongTK", "c" + s, dbDouble
    If Not TruongDaCo("HethongTK", "nt" + s) Then ThemTruong "HethongTK", "nt" + s, dbDouble
    
    If Not TruongDaCo("SoDuKhachHang", "n" + s) Then ThemTruong "SoDuKhachHang", "n" + s, dbDouble
    If Not TruongDaCo("SoDuKhachHang", "c" + s) Then ThemTruong "SoDuKhachHang", "c" + s, dbDouble
    If Not TruongDaCo("SoDuKhachHang", "nt" + s) Then ThemTruong "SoDuKhachHang", "nt" + s, dbDouble
    
    If Not TruongDaCo("TonKho", "n" + s) Then ThemTruong "TonKho", "n" + s, dbDouble
    If Not TruongDaCo("TonKho", "c" + s) Then ThemTruong "TonKho", "c" + s, dbDouble
End Sub
'=================================================================================================================
' Thñ tôc tæng hîp sè liÖu tõ tÖp ®­îc chän vµo tÖp ®ang më
'=================================================================================================================
Private Function NhapPS(db As Database, thang As Integer, tcuoi As Integer, vt As Integer, cn As Integer, ts As Integer, tt As Integer, ndau As Date, ncuoi As Date) As Boolean
Dim i As Integer, rsct As Recordset, rs2 As Recordset, soct As Long
Dim ctu As New ClsChungtu, k As Long
Dim MaCT As Long, mkh As Long, masocu As Long, mactcu As Long
Dim mct As Long, m As Long, st1 As String, st2 As String, id1 As Long, id2 As Long, st As String
Dim mtkn As Long, mtkc As Long, mvt As Long
Dim ml As Long, idcn As Long, tencn As String, LayCTBanHang As Integer
    
    NhapPS = False
    On Error Resume Next
    db.TableDefs("ChungTu").Fields.Append db.TableDefs("ChungTu").CreateField("SHTKNo", dbText, 20)
    db.TableDefs("ChungTu").Fields.Append db.TableDefs("ChungTu").CreateField("SHTKCo", dbText, 20)
    On Error GoTo 0
    
    Set rs2 = db.OpenRecordset("SELECT DISTINCTROW * FROM License", dbOpenSnapshot)
    idcn = rs2!tencn_id
    tencn = rs2!tencn
    pCT = LaySH(rs2!TKVattu, 2, "-")
    If Len(pCT) = 0 Or pCT = "..." Then pCT = CStr(tt)
        
    If LstDB(0).ListCount > 1 Then ExecuteSQL5 "INSERT INTO Users (MaSo, TenNSD, UserRight) VALUES (" + CStr(idcn) + ",'" + IIf(tencn <> "...", tencn, rs2!TenCty) + "',2)", False
    rs2.Close
    
    db.Execute "UPDATE " + ChungTu2TKNC(-1) + " SET SHTKNo = HethongTK.SoHieu"
    db.Execute "UPDATE " + ChungTu2TKNC(1) + " SET SHTKCo = HethongTK.SoHieu"
    ThemTruong "HoaDon", "MaKH1", dbLong
    ThemTruong "CTTaiSan", "ML", dbLong
    ThemTruong "CTTaiSan", "MN", dbLong
    ThemTruong "CTTaiSan", "MT", dbLong
    ThemTruong "ChungTu", "U", dbInteger
    ThemTruong "ChungTu", "U2", dbInteger
    ThemTruong "CTTaiSan", "U", dbInteger
    ThemTruong "CTTaiSan", "U2", dbInteger
    
    mkh = Lng_MaxValue("MaSo", "KhachHang")
    If mkh = 0 Then
        mkh = 1
        ml = SelectSQL("SELECT MaSo AS F1 FROM PhanLoaiKhachHang WHERE LEFT(SoHieu,1)='#'")
        ExecuteSQL5 "INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu) VALUES (1," + CStr(ml) + ",'#1')"
    End If
    ml = Lng_MaxValue("MaSo", "LoaiChungTu")
        
    If thang > 0 Then
        For i = thang To tcuoi
            XoaCTTheoID i, i * 1000 + idcn, 0
        Next
    Else
        For i = Month(ndau) To Month(ncuoi)
            XoaCTTheoID 0, i * 1000 + idcn, 0, ndau, ncuoi
        Next
    End If
    
    Set rs2 = db.OpenRecordset("SELECT * FROM DoituongCT WHERE MaSo>1", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs2.EOF
        k = SelectSQL("SELECT MaSo AS F1 FROM DoituongCT WHERE DienGiai='" + rs2!diengiai + "'")
        If k = 0 Then ExecuteSQL5 "INSERT INTO DoituongCT (MaSo,DienGiai) VALUES (" + CStr(Lng_MaxValue("MaSo", "DoituongCT") + 1) + ",'" + rs2!diengiai + "')"
        rs2.MoveNext
    Loop
    rs2.Close
    'LayCTBanHang = 0
    Set rsct = db.OpenRecordset("SELECT DISTINCTROW ChungTu.*,DoituongCT.DienGiai AS DG FROM ChungTu INNER JOIN DoituongCT ON ChungTu.MaDT=DoituongCT.MaSo WHERE MaCT > 0 AND " + IIf(thang > 0, WThang("ThangCT", thang, tcuoi), WNgay("NgayGS", ndau, ncuoi)) + " ORDER BY MaCT, ChungTu.MaSo", dbOpenSnapshot)
    masocu = Lng_MaxValue("MaSo", "ChungTu")
    mactcu = Lng_MaxValue("MaCT", "ChungTu")
    MaCT = mactcu
    Do While Not rsct.EOF
        If mct <> rsct!MaCT Then
a:
            mct = rsct!MaCT
            If Not MaCTOK(mct) Then
                Do While rsct!MaCT = mct
                    rsct.MoveNext
                    If rsct.EOF Then GoTo B
                Loop
                GoTo a
            End If
            soct = soct + 1
            MaCT = MaCT + 1
        End If
        If IsNull(rsct!shtkno) Then
            mtkn = 0
        Else
            mtkn = MaTKFix(db, rsct!shtkno, IIf(pTK > 0, pCT, ""), tencn)
        End If
        If IsNull(rsct!shtkco) Then
            mtkc = 0
        Else
            mtkc = MaTKFix(db, rsct!shtkco, IIf(pTK > 0, pCT, ""), tencn)
        End If
        'If mtkn = 0 Or mtkc = 0 Then MsgBox "Stop"
        ctu.InitChungtuRS rsct
        ctu.MaDT = SelectSQL("SELECT MaSo AS F1 FROM DoituongCT WHERE DienGiai='" + rsct!dg + "'")
        If vt = 0 Then
            ctu.MaNguon = 0
            ctu.MaKho = 0
            ctu.MaVattu = 0
        End If
        If cn = 0 Then
            ctu.makh = 0
            ctu.MaKHC = 0
        End If
        'ctu.InitChungtu 0, rsct!maloai, rsct!SoHieu, rsct!ThangCT, rsct!NgayCT, rsct!NgayGS, IIf(vt > 0, rsct!MaNguon, 0), IIf(vt > 0, rsct!MaKho, 0), rsct!diengiai, mtkn, mtkc, _
            rsct!sops, rsct!SoPS2No, rsct!SoPS2Co, IIf(vt > 0, rsct!MaVattu, 0), rsct!GhiChu, 1
        If ctu.tkno.tkcon > 0 Then
            If pTK > 0 Then
                ctu.tkno.InitTaikhoanMaSo mtkn
            Else
                st = MaSo2SoHieu(rsct!MaTkNo, "HethongTK")
                st1 = SelectSQLDB(db, "SELECT Ten AS F1, TenE AS F2 FROM HethongTK WHERE SoHieu='" + rsct!shtkno + "'", st2, id1, id2)
                m = ctu.tkno.ThemTKCon(Right(st, Len(st) - Len(ctu.tkno.sohieu)), st1, st2, id1, id2)
                If m > 0 Then ctu.tkno.InitTaikhoanMaSo m
            End If
        End If
        If ctu.TkCo.tkcon > 0 Then
            If pTK > 0 Then
                ctu.TkCo.InitTaikhoanMaSo mtkc
            Else
                st = MaSo2SoHieu(rsct!MaTkCo, "HethongTK")
                st1 = SelectSQLDB(db, "SELECT Ten AS F1, TenE AS F2 FROM HethongTK WHERE SoHieu='" + rsct!shtkco + "'", st2, id1, id2)
                m = ctu.TkCo.ThemTKCon(Right(st, Len(st) - Len(ctu.TkCo.sohieu)), st1, st2, id1, id2)
                If m > 0 Then ctu.TkCo.InitTaikhoanMaSo m
            End If
        End If
        If ctu.tkno.tkcon > 0 Then
            MsgBox "KiÓm tra tÖp " + db.Name + ", chøng tõ " + ctu.sohieu + " cã tµi kho¶n " + ctu.tkno.sohieu + " kh«ng hîp lÖ!"
            'rsct.Close
            'GoTo KT
        End If
        If ctu.TkCo.tkcon > 0 Then
            MsgBox "KiÓm tra tÖp " + db.Name + ", chøng tõ " + ctu.sohieu + " cã tµi kho¶n " + ctu.TkCo.sohieu + " kh«ng hîp lÖ!"
            'rsct.Close
            'GoTo KT
        End If
        ctu.MaCT = MaCT
        ctu.CT_ID = idcn + rsct!ThangCT * 1000
        ctu.User_ID = IIf(LstDB(0).ListCount > 1, idcn, UserID)
        ctu.GhiChungtuTH 0, 1
        Set rs2 = db.OpenRecordset("SELECT * FROM ChungTuLQ WHERE MaCT=" + CStr(rsct!MaCT), dbOpenSnapshot, dbForwardOnly)
        Do While Not rs2.EOF
            If SelectSQL("SELECT MaSo AS F1 FROM ChungTuLQ WHERE MaCT=" + CStr(rsct!MaCT) + " AND Loai=" + CStr(rs2!loai)) = 0 Then
                ExecuteSQL5 "INSERT INTO ChungTuLQ (MaSo,MaCT,Loai,HoTen,DiaChi,SoCTGoc) VALUES (" + CStr(Lng_MaxValue("MaSo", "ChungTuLQ") + 1) + "," + CStr(MaCT) + "," + CStr(rs2!loai) + ",'" + rs2!hoten + "','" + rs2!DiaChi + "','" + rs2!SoCTGoc + "')"
            End If
            rs2.MoveNext
        Loop
        rs2.Close
        If cn > 0 Then
            Set rs2 = db.OpenRecordset("SELECT HoaDon.* FROM HoaDon WHERE HoaDon.MaSo=" + CStr(rsct!MaSo), dbOpenSnapshot, dbForwardOnly)
            Do While Not rs2.EOF
                ExecuteSQL5 "INSERT INTO HoaDon (MaSo,Loai,MaKhachHang,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,TyLe,HD,KCT,HDBL,HTTT,MauSo,MaKH1) VALUES (" + CStr(ctu.MaSo) + "," + CStr(rs2!loai) + "," + CStr(mkh) _
                    + ",'" + rs2!KyHieu + "','" + rs2!sohd + "',#" + Format(rs2!NgayPH, Mask_DB) + "#,'" + rs2!MatHang + "'," + DoiDau(rs2!SoLuong) + "," + DoiDau(rs2!ThanhTien) + "," + CStr(rs2!TyLe) + "," + CStr(rs2!HD) + "," + CStr(rs2!KCT) + "," + CStr(rs2!HDBL) + ",'" + rs2!HTTT + "','" + rs2!MauSo + "'," + CStr(rs2!MaKhachHang) + ")"
                rs2.MoveNext
            Loop
            rs2.Close
        End If
        If ts > 0 And rsct!maloai > 8 Then
            If SelectSQL("SELECT MaSo AS F1 FROM CTTaiSan WHERE MaCTKT=" + CStr(MaCT)) = 0 Then
                Set rs2 = db.OpenRecordset("SELECT TOP 1 CTTaiSan.* FROM CTTaiSan WHERE MaCTKT=" + CStr(rsct!MaCT), dbOpenSnapshot, dbForwardOnly)
                If Not rs2.EOF Then
                    ExecuteSQL5 "INSERT INTO CTTaiSan (MaSo, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, " _
                        & "MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, " _
                        & "CL_NS, CL_TBS, CL_CNK, CL_TD, MaCTKT,ML,MN) VALUES (" + CStr(Lng_MaxValue("MaSo", "CTTaiSan") + 1) + ",'" + rs2!sohieu + "'," + CStr(rs2!thang) _
                        + ",#" + Format(rs2!VaoSo, Mask_DB) + "#,#" + Format(rs2!NgayGhi, Mask_DB) + "#,'" _
                        + rs2!diengiai + "'," + CStr(ml) + "," + CStr(ml) + "," + CStr(rs2!MaTS) + "," _
                        + DoiDau(rs2!NG_NS) + "," + DoiDau(rs2!NG_TBS) + "," + DoiDau(rs2!NG_CNK) + "," + DoiDau(rs2!NG_TD) + "," _
                        + DoiDau(rs2!CL_NS) + "," + DoiDau(rs2!CL_TBS) + "," + DoiDau(rs2!CL_CNK) + "," + DoiDau(rs2!CL_TD) + "," _
                        + CStr(MaCT) + "," + CStr(rs2!maloai) + "," + CStr(rs2!MaNhom) + ")"
                    rs2.MoveNext
                End If
                rs2.Close
            End If
        End If
        DBEngine.Idle
        If rsct.RecordCount Mod 100 = 0 Then
            Label1(3).Caption = "Sè CT ®· tæng hîp: " + CStr(CLng5(Label1(3).ToolTipText) + soct)
            Me.Refresh
        End If
        rsct.MoveNext
    Loop
B:
    rsct.Close
    'If LayCTBanHang > 0 Then GoTo THCT
    'LayCTBanHang = 1
    'Set rsct = db.OpenRecordset("SELECT DISTINCTROW ChungTu.*,DoituongCT.DienGiai AS DG FROM ChungTu INNER JOIN DoituongCT ON ChungTu.MaDT=DoituongCT.MaSo WHERE (MaLoai=8) AND MaCT > 0 AND " + IIf(thang > 0, WThang("ThangCT", thang, tcuoi), WNgay("NgayGS", ndau, ncuoi)) + " ORDER BY MaCT, ChungTu.MaSo", dbOpenSnapshot)
    'GoTo LCT
    If vt > 0 Then CongDDVT db, masocu, tencn, Chk(4).Value, True, pTK, pCT
    If cn > 0 Then CongDDCN db, masocu, Chk(4).Value, True, pTK, pCT
    If ts > 0 Then CongDDTS db, mactcu, tencn, Chk(4).Value, True, pTK, pCT
    
    NhapPS = True
    Label1(3).ToolTipText = CStr(soct)
KT:
    Set rsct = Nothing
    Set ctu = Nothing
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

Private Function MaCTOK(mct As Long) As Boolean
    Dim i As Integer
    
    If SoCTChon = 0 Then
        MaCTOK = True
    Else
        MaCTOK = False
        For i = 1 To SoCTChon
            If MaCT(i) = mct Then
                MaCTOK = True
                Exit For
            End If
        Next
    End If
End Function
