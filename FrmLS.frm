VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmLS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "B∂ng l∑i su t"
   ClientHeight    =   7635
   ClientLeft      =   3675
   ClientTop       =   945
   ClientWidth     =   3930
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
   Icon            =   "FrmLS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7635
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtshkh 
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      LinkItem        =   "SË hi÷u vÀt t≠ c«n xem"
      MaxLength       =   20
      TabIndex        =   3
      Tag             =   "0"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdkh 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      Picture         =   "FrmLS.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Danh Æi”m vÀt t≠"
      Top             =   840
      Width           =   375
   End
   Begin VB.ListBox Lst 
      Height          =   2985
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   3735
   End
   Begin VB.ComboBox CboNT 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Left            =   960
      LinkItem        =   "SË hi÷u tµi kho∂n c«n xem"
      MaxLength       =   12
      TabIndex        =   0
      Tag             =   "0"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdTK 
      Height          =   375
      Left            =   2400
      Picture         =   "FrmLS.frx":5C5C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "H÷ thËng tµi kho∂n"
      Top             =   80
      Width           =   375
   End
   Begin MSGrid.Grid GrdNT 
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Tag             =   "20"
      Top             =   1920
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   2990
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
      Cols            =   3
      FixedRows       =   0
      ScrollBars      =   2
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   2760
      Picture         =   "FrmLS.frx":60D6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   2760
      Picture         =   "FrmLS.frx":74F8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   2760
      Picture         =   "FrmLS.frx":89DA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   2760
      Picture         =   "FrmLS.frx":9E08
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   960
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "0"
      Top             =   3600
      Width           =   1455
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   3600
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
   Begin VB.Label lbkh 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Tag             =   "1"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ßËi t≠Óng"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Chi ti’t c„ theo d‚i l∑i su t vµ l∑i su t hi÷n thÍi"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tµi kho∂n"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L∑i su t % th∏ng"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   13
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tı ngµy"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "FrmLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer
Dim ngay As Date
Dim taikhoan As New ClsTaikhoan
Dim ckh As New ClsKhachHang
Dim f1 As Integer

Private Sub CboNT_Click()
    Dim mnt As Long
    If taikhoan.MaSo = 0 Or CboNT.ListIndex < 0 Then Exit Sub
    mnt = CboNT.ItemData(CboNT.ListIndex)
    taikhoan.InitTaikhoanMaSo SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE SoHieu='" + taikhoan.sohieu + "' AND " + IIf(mnt > 0, "MaNT=" + CStr(mnt), "MaNT<=0"))
    LietKeNgte
End Sub

'====================================================================================================
' Th™m, Ghi, X„a nguy™n t÷
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0:
            txtNhap.Text = "0"
            RFocus MedNgay
            ThemMoi = 1
        Case 1:
            If taikhoan.MaSo = 0 Or taikhoan.tkcon > 0 Then
                RFocus txtShTk
                Exit Sub
            End If
            If (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) And ckh.MaSo = 0 Then
                RFocus txtshkh
                Exit Sub
            End If
            Select Case ThemMoi
                Case 0:
                    GrdNT.col = 2
                    If Not IsNumeric(GrdNT.Text) Then Exit Sub
                    If ExecuteSQL5("UPDATE LaiSuat SET Ngay=#" + Format(ngay, Mask_DB) + "#,LS=" + DoiDau(Cdbl5(txtNhap.Text)) + " WHERE MaKH=" + CStr(ckh.MaSo) + " AND MaSo=" + GrdNT.Text) <> 0 Then Exit Sub
                    GrdNT.col = 0
                    GrdNT.Text = MedNgay.Text
                    GrdNT.col = 1
                    GrdNT.Text = txtNhap.Text
                Case 1:
                    If ExecuteSQL5("INSERT INTO LaiSuat (MaTK,Ngay,LS,MaKH) VALUES (" + CStr(taikhoan.MaSo) + ",#" + Format(ngay, Mask_DB) + "#," + DoiDau(Cdbl5(txtNhap.Text)) + "," + CStr(ckh.MaSo) + ")") <> 0 Then Exit Sub
                    GrdNT.AddItem MedNgay.Text + Chr(9) + txtNhap.Text + Chr(9) + CStr(Lng_MaxValue("MaSo", "LaiSuat")), InsertGridRow(GrdNT, 0, MedNgay.Text)
                    ThemMoi = 0
                    GrdNT.Row = GrdNT.Rows - 1
                    GrdNT.col = 0
                    If GrdNT.Text = "" Then GrdNT.RemoveItem GrdNT.Row
                    GrdNT.Row = 0
            End Select
            KeTaiKhoan
        Case 2:
            GrdNT.col = 2
            If Not IsNumeric(GrdNT.Text) Then Exit Sub
            If ExecuteSQL5("DELETE FROM LaiSuat WHERE MaSo=" + GrdNT.Text) <> 0 Then Exit Sub
            GrdNT.RemoveItem GrdNT.Row
            If GrdNT.Rows < GrdNT.tag Then GrdNT.Rows = GrdNT.tag
            KeTaiKhoan
        Case 3:
            Unload Me
    End Select
End Sub

'====================================================================================================
' Xˆ l˝ ph›m n„ng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask Then
        Select Case KeyCode
            Case vbKeyT:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyG:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyX:
                RFocus Command(2)
                Command_Click 2
            Case vbKeyV:
                RFocus Command(3)
                Command_Click 3
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
'====================================================================================================
' KhÎi tπo cˆa sÊ
'====================================================================================================
Private Sub Form_Load()
    ColumnSetUp GrdNT, 0, 820, 2
    ColumnSetUp GrdNT, 1, 1420, 2
    
    InitDateVars MedNgay, ngay
    
    Caption = Caption + " - " + CStr(pNamTC)
    
    KeTaiKhoan
    
    SetFont Me
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set taikhoan = Nothing
    Set ckh = Nothing
End Sub

Private Sub GrdNt_click()
    Dim i As Integer
    
    SendKeys "{Home}", True
    SetGridIndex GrdNT, GrdNT.Row
    With GrdNT
        .col = 2
        If Not IsNumeric(.Text) Then Exit Sub
        .col = 0
        MedNgay.Text = .Text
        ngay = CVDate(.Text)
        .col = 1
        txtNhap.Text = .Text
        ThemMoi = 0
        RFocus MedNgay
    End With
End Sub

Private Sub GrdNt_KeyPress(KeyAscii As Integer)
    SendKeys "{Home}", True
    SetGridIndex GrdNT, GrdNT.Row
    
    If KeyAscii = 13 Then GrdNt_click
End Sub

Private Sub Lst_Click()
    If f1 = 1 Or Lst.ListIndex < 0 Or (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) Then Exit Sub
    f1 = 1
    taikhoan.InitTaikhoanMaSo Lst.ItemData(Lst.ListIndex)
    txtShTk.Text = taikhoan.sohieu
    txtShTk_LostFocus
    f1 = 0
    RFocus txtShTk
End Sub

Private Sub txtNhap_GotFocus()
    AutoSelect txtNhap
End Sub
'====================================================================================================
' Hi”n thﬁ danh s∏ch nguy™n t÷
'====================================================================================================
Private Sub LietKeNgte()
    Dim rs_ngte As Recordset

    ClearGrid GrdNT, GrdNT.tag
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT * FROM LaiSuat WHERE MaTK=" + CStr(taikhoan.MaSo) + " AND MaKH=" + CStr(ckh.MaSo) + " ORDER BY Ngay DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        GrdNT.AddItem Format(rs_ngte!ngay, Mask_D) + Chr(9) + Format(rs_ngte!ls, Mask_2) + Chr(9) + CStr(rs_ngte!MaSo), 0
        rs_ngte.MoveNext
    Loop
    GrdNT.Rows = IIf(rs_ngte.RecordCount > GrdNT.tag, rs_ngte.RecordCount, GrdNT.tag)
    GrdNT.Row = 0
    GrdNT.col = 0
    rs_ngte.Close
    Set rs_ngte = Nothing
    GrdNt_click
End Sub

Private Sub txtNhap_KeyPress(KeyAscii As Integer)
    KeyProcess txtNhap, KeyAscii
End Sub

Private Sub txtNhap_LostFocus()
    txtNhap.Text = Format(txtNhap.Text, Mask_2)
End Sub

Private Sub MedNgay_GotFocus()
    AutoSelect MedNgay
End Sub

Private Sub MedNgay_LostFocus()
    If IsDate(MedNgay.Text) Then
        ngay = CDate(MedNgay.Text)
    Else
        MsgBox "NhÀp sai ngµy !", vbExclamation, App.ProductName
        RFocus MedNgay
    End If
End Sub

Private Sub txtShTk_GotFocus()
    AutoSelect txtShTk
End Sub

Private Sub txtShTk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdtk_Click
End Sub

Private Sub txtShTk_LostFocus()
    taikhoan.InitTaikhoanSohieu txtShTk
    LbTenTk.Caption = taikhoan.Ten
    Int_RecsetToCbo "SELECT HethongTK.MaNT As F2,NguyenTe.KyHieu As F1 FROM HethongTK INNER JOIN" _
                & " NguyenTe ON HethongTK.MaNT = NguyenTe.MaSo WHERE HethongTK.SoHieu = '" + taikhoan.sohieu _
                + "' ORDER BY NguyenTe.KyHieu", CboNT
    If CboNT.ListCount > 0 Then
        CboNT.Enabled = True
        CboNT.AddItem "VND", 0
        CboNT.ItemData(0) = 0
    Else
        CboNT.Enabled = False
    End If
    txtshkh.Enabled = (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID)
    cmdkh.Enabled = (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID)
    If Not (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) Then ckh.InitKhachHangMaSo 0
    KeTaiKhoan
    LietKeNgte
End Sub

Private Sub cmdtk_Click()
    Me.MousePointer = 11
    txtShTk.Text = FrmTaikhoan.ChonTk(txtShTk.Text)
    Me.MousePointer = 0
    RFocus txtShTk
End Sub

Private Sub KeTaiKhoan()
    Dim rs As Recordset, ls As Double, sh As String
        
    Lst.Clear
    If (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) Then
        Set rs = DBKetoan.OpenRecordset("SELECT HethongTK.MaSo,KhachHang.MaSo AS MaKH,KhachHang.SoHieu,KhachHang.Ten FROM (HethongTK INNER JOIN LaiSuat ON HethongTK.MaSo=LaiSuat.MaTK) INNER JOIN KhachHang ON LaiSuat.MaKH=KhachHang.MaSo GROUP BY HethongTK.MaSo,KhachHang.MaSo,KhachHang.SoHieu,KhachHang.Ten ORDER BY KhachHang.SoHieu", dbOpenSnapshot)
        Do While Not rs.EOF
            ls = SelectSQL("SELECT TOP 1 LS AS F1 FROM LaiSuat WHERE MaTK=" + CStr(rs!MaSo) + " AND MaKH=" + CStr(rs!makh) + " ORDER BY Ngay DESC")
            Lst.AddItem rs!sohieu + " - " + rs!Ten + " - LS: " + Format(ls, Mask_2)
            rs.MoveNext
        Loop
    Else
        Set rs = DBKetoan.OpenRecordset("SELECT HethongTK.MaSo,SoHieu,Ten,KyHieu FROM (HethongTK LEFT JOIN NguyenTe ON HethongTK.MaNT=NguyenTe.MaSo) INNER JOIN LaiSuat ON HethongTK.MaSo=LaiSuat.MaTK GROUP BY HethongTK.MaSo,SoHieu,Ten,KyHieu ORDER BY SoHieu, KyHieu", dbOpenSnapshot)
        Do While Not rs.EOF
            ls = SelectSQL("SELECT TOP 1 LS AS F1 FROM LaiSuat WHERE MaTK=" + CStr(rs!MaSo) + " ORDER BY Ngay DESC")
            If IsNull(rs!KyHieu) Then sh = "VND" Else sh = rs!KyHieu
            Lst.AddItem rs!sohieu + " - " + sh + " - LS: " + Format(ls, Mask_2)
            Lst.ItemData(Lst.NewIndex) = rs!MaSo
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub txtshkh_GotFocus()
    AutoSelect txtshkh
End Sub

Private Sub txtshkh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdkh_Click
End Sub

Private Sub txtshkh_LostFocus()
    Dim mvt As Long
    
    ckh.InitKhachHangSohieu txtshkh
    lbkh.Caption = ckh.Ten
    LietKeNgte
    HienThongBao Me.Caption, 1
End Sub

Private Sub cmdkh_Click()
    Me.MousePointer = 11
    txtshkh.Text = FrmKhachHang.ChonKhachHang(txtshkh.Text)
    Me.MousePointer = 0
    RFocus txtshkh
End Sub

