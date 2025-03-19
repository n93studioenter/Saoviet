VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmBCVT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "B¸o c¸o nhËp xuÊt kho"
   ClientHeight    =   3480
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6600
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
   Icon            =   "FrmBCVT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CboLoai 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2520
      Width           =   3375
   End
   Begin VB.ComboBox CboKho 
      Height          =   315
      Index           =   2
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Theo ph©n lo¹i hµng ho¸"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Theo kho"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   0
      Left            =   2160
      TabIndex        =   19
      Top             =   0
      Width           =   4335
      Begin VB.OptionButton OptTG 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tõ ngµy"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Tag             =   "From"
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptTG 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tõ th¸ng"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Tag             =   "From"
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.ComboBox CboThang 
         Height          =   315
         Index           =   1
         ItemData        =   "FrmBCVT.frx":57E2
         Left            =   3120
         List            =   "FrmBCVT.frx":57E4
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.ComboBox CboThang 
         Height          =   315
         Index           =   0
         ItemData        =   "FrmBCVT.frx":57E6
         Left            =   1080
         List            =   "FrmBCVT.frx":57E8
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin MSMask.MaskEdBox MedNgay 
         Height          =   315
         Index           =   0
         Left            =   1080
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
         Left            =   2880
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
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "®Õn ngµy"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   21
         Tag             =   "to"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "®Õn th¸ng"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   20
         Tag             =   "to"
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.ComboBox CboKho 
      Height          =   315
      Index           =   1
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox CboKho 
      Height          =   315
      Index           =   0
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   2880
      Picture         =   "FrmBCVT.frx":57EA
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "&View"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   4080
      Picture         =   "FrmBCVT.frx":695C
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "&Print"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   5280
      Picture         =   "FrmBCVT.frx":7DBE
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "&Return"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "B¶ng kª chøng tõ nhËp kho"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Tag             =   "Journal Ledger"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "B¶ng kª chøng tõ l­u chuyÓn néi bé"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Tag             =   "Journal Ledger"
      Top             =   960
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "®Õn kho"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   18
      Tag             =   "to"
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tõ kho"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   17
      Tag             =   "to"
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "FrmBCVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim baocao As Integer
Dim ngay(0 To 1) As Date

Private Sub Command_Click(Index As Integer)
    Dim tdau As Integer, tcuoi As Integer
    
    Me.MousePointer = 0
    Select Case Index
        Case 0, 1:
            tdau = CboThang(0).ItemData(CboThang(0).ListIndex)
            tcuoi = CboThang(1).ItemData(CboThang(1).ListIndex)
            SetRptInfo
            Select Case baocao
                Case 0:
                    BangKeLuuChuyen IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(0), ngay(1), CboKho(0).ItemData(CboKho(0).ListIndex), CboKho(1).ItemData(CboKho(1).ListIndex)
                Case 1:
                    BangKeNhapKho IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(0), ngay(1), IIf(Chk(0).Value = 1, CboKho(2).ItemData(CboKho(2).ListIndex), 0), IIf(Chk(1).Value = 1, CboLoai.ItemData(CboLoai.ListIndex), 0)
            End Select
            frmMain.Rpt.Destination = Index
            If Not RptOK(frmMain.Rpt.ReportFileName, 0) Then
                MsgBox "MÉu b¸o c¸o ®· bÞ thay ®æi!", vbCritical, App.ProductName
                GoTo KT
            End If
            frmMain.Rpt.WindowTitle = OptBC(baocao).Caption
            On Error GoTo Bad
            frmMain.Rpt.Action = 1
            On Error GoTo 0
            GoTo KT
        Case 2:
            Unload Me
            GoTo KT
    End Select
Bad:
    MsgBox "Error " + CStr(Err.Number) + ": " + Err.Description, vbExclamation, App.ProductName
KT:
    Me.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyX:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyI:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyV:
                RFocus Command(2)
                Command_Click 2
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim chi_so As Integer
    
    Int_RecsetToCbo "SELECT MaSo As F2,TenKho As F1 FROM KhoHang ORDER BY TenKho", CboKho(0)
    CboCopy CboKho(0), CboKho(1)
    CboCopy CboKho(0), CboKho(2)
    Int_RecsetToCbo "SELECT PhanLoaiVattu.MaSo As F2, (PhanLoaiVattu.SoHieu + ' - '+ PhanLoaiVattu.TenPhanLoai) As F1" _
        & " FROM PhanLoaiVattu ORDER BY PhanLoaiVattu.SoHieu", CboLoai
    For chi_so = 0 To 1
        AddMonthToCbo CboThang(chi_so)
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next
    
    SetFont Me
End Sub

Private Sub OptBc_Click(Index As Integer)
    baocao = Index
End Sub

Private Sub BangKeLuuChuyen(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, K1 As Long, k2 As Long)
    Dim sql As String, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    sql = "SELECT MaCT, SoHieu, NgayCT, NgayGS, DienGiai, Sum(SoPS) AS SoTien FROM ChungTu " _
        & " WHERE MaLoai=4 AND MaKho=" + CStr(K1) + " AND MaNguon=" + CStr(k2) + " AND " + wsql _
        + " GROUP BY MaCT, SoHieu, NgayCT, NgayGS, DienGiai ORDER BY NgayCT, MaCT"
    SetSQL "QTheKho", sql
    
    frmMain.Rpt.ReportFileName = "BKLCNB.RPT"
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi), ThoiGianN(ndau, ncuoi)) + "'"
    frmMain.Rpt.Formulas(4) = "DienGiai = IF PageNumber() = 1 THEN 'Tõ kho: " + tenkho(K1) + " ®Õn kho: " + tenkho(k2) + "'"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi)
End Sub

Private Sub BangKeNhapKho(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, mk As Long, ml As Long)
    Dim sql As String, wsql As String, cap As Integer
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi)) + " AND " + IIf(mk > 0, "((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + "))", "(MaLoai=1 OR MaLoai=4)")
    cap = SelectSQL("SELECT Cap AS F1 FROM PhanLoaiVattu WHERE MaSo=" + CStr(ml))
    
    sql = "SELECT MaCT, ChungTu.SoHieu, NgayCT, NgayGS, DienGiai, Sum(SoPS) AS SoTien, First(HethongTK.Sohieu) AS TKDU, First(KhachHang.Sohieu) AS SHKH, First(KhachHang.Ten) AS TenKH " _
        & " FROM (((" + ChungTu2TKNC(1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo) LEFT JOIN KhachHang ON ChungTu.MaKHC=KhachHang.MaSo " _
        & " WHERE " + wsql + IIf(ml > 0, " AND PhanLoaiVattu.SoHieu LIKE '" + MaSo2SoHieu(ml, "PhanLoaiVattu") + "*' AND (PhanLoaiVattu.Cap>" + CStr(cap) + " OR PhanLoaiVattu.MaSo=" + CStr(ml) + ")", "") + " GROUP BY MaCT, ChungTu.SoHieu, NgayCT, NgayGS, DienGiai ORDER BY NgayCT, MaCT"
    SetSQL "QTheKho", sql
    
    frmMain.Rpt.ReportFileName = "BKNK.RPT"
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi), ThoiGianN(ndau, ncuoi)) + "'"
    If mk > 0 Then frmMain.Rpt.Formulas(4) = "DienGiai = IF PageNumber() = 1 THEN 'Kho: " + tenkho(mk) + "'"
    If ml > 0 Then frmMain.Rpt.Formulas(5) = "DienGiai2 = IF PageNumber() = 1 THEN 'Lo¹i hµng ho¸: " + TenPLVT(0, ml) + "'"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi)
End Sub


