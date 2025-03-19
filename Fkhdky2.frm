VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FKHDauKy2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sè d­ ®Çu kú kh¸ch hµng"
   ClientHeight    =   7095
   ClientLeft      =   870
   ClientTop       =   735
   ClientWidth     =   9885
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
   Icon            =   "Fkhdky2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   8760
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "0"
      Top             =   6240
      Width           =   735
   End
   Begin MSGrid.Grid GrdVT 
      Height          =   5535
      Left            =   120
      TabIndex        =   16
      Tag             =   "30"
      Top             =   720
      Width           =   9615
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   9763
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
      Rows            =   30
      Cols            =   9
      FixedRows       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   6120
      MaxLength       =   20
      TabIndex        =   5
      Text            =   "0"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   7440
      MaxLength       =   20
      TabIndex        =   6
      Text            =   "0"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdct 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      Picture         =   "Fkhdky2.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   255
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "0"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Left            =   8400
      Picture         =   "Fkhdky2.frx":5B84
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   2
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   20
      TabIndex        =   1
      Top             =   6240
      Width           =   1335
   End
   Begin VB.ComboBox CboKho 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H¹n TT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   8760
      TabIndex        =   21
      Top             =   480
      Width           =   735
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   6195
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nguyªn tÖ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   19
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D­ cã"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   18
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè hiÖu TK"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tæng tiÒn"
      Height          =   255
      Index           =   6
      Left            =   3960
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   4755
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D­ nî"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   13
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tªn kh¸ch hµng"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   12
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè hiÖu KH"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ph©n lo¹i"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FKHDauKy2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim taikhoan As New ClsTaikhoan
Dim ckh As New ClsKhachHang
Dim psw As String

Private Sub CboKho_Click()
    LietKeTonKho CboKho.ItemData(CboKho.ListIndex)
End Sub

Private Sub CmdCt_Click()
    Dim tien1 As Double, tien2 As Double, i As Integer, luong As Double, st As String
    
    If CboKho.ListIndex < 0 Then
        MsgBox "H·y ®¨ng ký danh s¸ch kh¸ch hµng!", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    If taikhoan.MaSo = 0 Then
        MsgBox "H·y nhËp sè hiÖu tµi kho¶n!", vbExclamation, App.ProductName
        RFocus txtTon(0)
        Exit Sub
    End If
    
    If taikhoan.tkcon > 0 Then
        MsgBox "H·y nhËp sè hiÖu tµi kho¶n chi tiÕt!", vbExclamation, App.ProductName
        RFocus txtTon(0)
        Exit Sub
    End If
    
    If taikhoan.tk_id <> TKCNKH_ID And taikhoan.tk_id <> TKCNPT_ID Then
        MsgBox "H·y nhËp tµi kho¶n c«ng nî!", vbExclamation, App.ProductName
        RFocus txtTon(0)
        Exit Sub
    End If
        
    If ckh.MaSo = 0 Then
        MsgBox "H·y nhËp sè hiÖu kh¸ch hµng!", vbExclamation, App.ProductName
        RFocus txtTon(1)
        Exit Sub
    End If
    
    For i = 0 To txtTon.count - 1
        txtTon_LostFocus i
    Next

    tien1 = Cdbl5(txtTon(3).Text)
    tien2 = Cdbl5(txtTon(4).Text)
    luong = Cdbl5(txtTon(5).Text)
    
    If tien1 <> 0 And tien2 <> 0 Then
        MsgBox "NhËp d­ nî hoÆc d­ cã!", vbExclamation, App.ProductName
        RFocus txtTon(3)
        Exit Sub
    End If
    
    Me.MousePointer = 0
    With GrdVT
        'If taikhoan.TK_ID <> TKCNKH_ID Then
        '    For i = 0 To .Rows - 1
        '        .col = 7
        '        .Row = i
        '        If .Text = "" Then Exit For
        '        If CLng5(.Text) = taikhoan.MaSo Then
        '            .col = 8
        '            If CLng5(.Text) = ckh.MaSo Then
        '                If psw <> "" Then
        '                    If FPsw.GetPswX() <> psw Then GoTo XongDK
        '                End If
        '                GrdVT.RemoveItem i
        '                GrdVT.AddItem taikhoan.SoHieu + Chr(9) + ckh.SoHieu + Chr(9) + ckh.Ten + Chr(9) + Format(tien1, Mask_0) + Chr(9) + Format(tien2, Mask_0) + Chr(9) + Format(luong, Mask_2) + Chr(9) + CStr(CInt5(txtTon(6).Text)) + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) + CStr(ckh.MaSo), i
        '                ckh.GhiDauKy taikhoan.MaSo
        '                TongTien
        '                ExecuteSQL5 "UPDATE CNDauNam SET DuNo_0=" + DoiDau(tien1) + ",DuCo_0=" + DoiDau(tien2) + ",DuNT_0=" + DoiDau(luong) + ",HanTT=" + CStr(CInt5(txtTon(6).Text)) + " WHERE MaTaiKhoan=" + CStr(taikhoan.MaSo) + " AND MaKhachHang=" + CStr(ckh.MaSo)
        '                If DBKetoan.RecordsAffected = 0 Then
        '                    ExecuteSQL5 "INSERT INTO CNDauNam (MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0,HanTT) VALUES (" + CStr(taikhoan.MaSo) + "," + CStr(ckh.MaSo) + "," + DoiDau(tien1) + "," + CStr(tien2) + "," + CStr(luong) + "," + CStr(CInt5(txtTon(6).Text)) + ")"
        '                End If
        '                RFocus txtTon(0)
        '                GoTo XongDK
        '            End If
        '        End If
        '    Next
        'Else
            ExecuteSQL5 "INSERT INTO CNDauNam (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0,HanTT) VALUES (" + CStr(Lng_MaxValue("MaSo", "CNDauNam") + 1) + "," + CStr(taikhoan.MaSo) + "," + CStr(ckh.MaSo) + "," + DoiDau(tien1) + "," + CStr(tien2) + "," + CStr(luong) + "," + CStr(CInt5(txtTon(6).Text)) + ")"
        'End If
        .AddItem taikhoan.sohieu + Chr(9) + ckh.sohieu + Chr(9) + ckh.Ten + Chr(9) + Format(tien1, Mask_0) + Chr(9) + Format(tien2, Mask_0) + Chr(9) + Format(luong, Mask_2) + Chr(9) + CStr(CInt5(txtTon(6).Text)) + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) + CStr(ckh.MaSo), NewRowIndex(GrdVT, 0)
        ckh.GhiDauKy2 taikhoan.MaSo
        .Row = .Rows - 1
        .col = 0
        If .Text = "" Then .RemoveItem .Row
        .Row = 0
        RFocus txtTon(0)
    End With
    TongTien
    GoTo XongDK
XongDK:
    Me.MousePointer = 0
End Sub

Private Sub Command_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    HienThongBao Caption, 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = vbAltMask And KeyCode = vbKeyV) Or KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    ColumnSetUp GrdVT, 0, 1300, 0
    ColumnSetUp GrdVT, 1, 1300, 0
    ColumnSetUp GrdVT, 2, 2020, 0
    ColumnSetUp GrdVT, 3, 1300, 1
    ColumnSetUp GrdVT, 4, 1300, 1
    ColumnSetUp GrdVT, 5, 1300, 1
    ColumnSetUp GrdVT, 6, 700, 2
    ColumnSetUp GrdVT, 7, 1, 0
    ColumnSetUp GrdVT, 8, 1, 0
    Caption = Caption + " - " + CStr(pNamTC)
    Int_RecsetToCbo "SELECT DISTINCTROW MaSo As F2,SoHieu + ' - '  + TenPhanLoai As F1 FROM PhanLoaiKhachHang WHERE PLCon=0 AND LEFT(SoHieu,1)<>'#' ORDER BY SoHieu", CboKho
    
    psw = GetSetting(IniPath, "Environment", "InvPsw")
    SetFont Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set taikhoan = Nothing
    Set ckh = Nothing
End Sub

Private Sub GrdVT_DblClick()
    Dim i As Integer, ms As Long
    
    With GrdVT
        .col = 0
        If .Text = "" Then Exit Sub
        For i = 0 To 2
            .col = i
            txtTon(i).Text = .Text
        Next
        For i = 3 To 6
            .col = i
            txtTon(i).Text = .Text
        Next

        .col = 0
        txtTon_LostFocus 0
        txtTon_LostFocus 1
        
        ms = SelectSQL("SELECT TOP 1 MaSo AS F1 FROM CNDauNam WHERE MaTaiKhoan=" + CStr(taikhoan.MaSo) _
         + " AND MaKhachHang=" + CStr(ckh.MaSo) + " AND DuNo_0=" + DoiDau(Cdbl5(txtTon(3).Text)) + " AND (HanTT=" + CStr(CInt5(txtTon(6).Text)) + " OR HanTT=0)")
        If ms > 0 Then ExecuteSQL5 "DELETE * FROM CNDauNam WHERE MaSo=" + CStr(ms)
        ckh.GhiDauKy2 taikhoan.MaSo
        .RemoveItem .Row
        If .Rows < .tag Then .Rows = .tag
            
        RFocus txtTon(0)
    End With
End Sub

Private Sub GrdVT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then GrdVT_DblClick
End Sub

Private Sub GrdVT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdVT, GrdVT.col
    End If
End Sub

Private Sub txtTon_GotFocus(Index As Integer)
    AutoSelect txtTon(Index)
End Sub

Private Sub txtTon_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0:
            If KeyAscii = 13 Then
                txtTon(0).Text = FrmTaikhoan.ChonTk(txtTon(0).Text)
            End If
        Case 1:
            If KeyAscii = 13 Then
                txtTon(1).Text = FrmKhachHang.ChonKhachHang(txtTon(1).Text)
            End If
        Case 2:
            KeyAscii = 0
        Case 3, 4, 5, 6:
            If KeyAscii = 13 Then
                txtTon_LostFocus 3
                CmdCt_Click
            Else
                KeyProcess txtTon(Index), KeyAscii, False
            End If
    End Select
End Sub

Private Sub txtTon_LostFocus(Index As Integer)
    Dim luong As Double
    
    Select Case Index
        Case 0:
            If txtTon(0).Text <> "" Then
                taikhoan.InitTaikhoanSohieu txtTon(0).Text
                txtTon(6).Enabled = (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID)
            Else
                taikhoan.InitTaikhoanMaSo 0
                txtTon(6).Enabled = False
            End If
        Case 1:
            If txtTon(1).Text <> "" Then
                ckh.InitKhachHangSohieu txtTon(1).Text
            Else
                ckh.InitKhachHangMaSo 0
            End If
            txtTon(2).Text = ckh.Ten
            txtTon(5).Enabled = (ckh.MaNT > 0)
            If ckh.MaNT = 0 Then txtTon(5).Text = "0"
        Case 3, 4:
            txtTon(Index).Text = Format(txtTon(Index).Text, Mask_0)
            If Cdbl5(txtTon(Index).Text) <> 0 Then txtTon(7 - Index).Text = "0"
        Case 5:
            txtTon(5).Text = Format(txtTon(5).Text, Mask_2)
    End Select
End Sub
'======================================================================================
' Thñ tôc liÖt kª tån kho
'======================================================================================
Private Sub LietKeTonKho(mkho As Long)
    Dim rs_ton As Recordset
    
    Me.MousePointer = 11
    ClearGrid GrdVT, GrdVT.tag
    Set rs_ton = DBKetoan.OpenRecordset("SELECT HethongTK.MaSo,HethongTK.Kieu,HethongTK.SoHieu AS SHTK,CNDauNam.MaKhachHang,KhachHang.SoHieu,KhachHang.Ten," _
        & " CNDauNam.DuNo_0 AS DuNo,CNDauNam.DuCo_0 AS DuCo,CNDauNam.DuNT_0 AS DuNT,HanTT " _
        & " FROM ((CNDauNam INNER JOIN KhachHang ON CNDauNam.MaKhachHang = KhachHang.MaSo) INNER JOIN HethongTK ON CNDauNam.MaTaiKhoan=HethongTK.MaSo) INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo" _
        & " WHERE PhanLoaiKhachHang.SoHieu LIKE '" + MaSo2SoHieu(mkho, "PhanLoaiKhachHang") + "*' AND (CNDauNam.DuNo_0 <> 0 OR CNDauNam.DuCo_0 <> 0) ORDER BY HethongTK.SoHieu DESC, KhachHang.SoHieu DESC", dbOpenSnapshot)
    Do While Not rs_ton.EOF
        GrdVT.AddItem rs_ton!shtk + Chr(9) + rs_ton!sohieu + Chr(9) + rs_ton!Ten + Chr(9) + Format(rs_ton!duno, Mask_0) + Chr(9) + Format(rs_ton!duco, Mask_0) + Chr(9) + Format(rs_ton!dunt, Mask_2) + Chr(9) + CStr(rs_ton!HanTT) + Chr(9) + CStr(rs_ton!MaSo) + Chr(9) + CStr(rs_ton!MaKhachHang), 0
        rs_ton.MoveNext
    Loop
    GrdVT.Rows = IIf(rs_ton.RecordCount > GrdVT.tag, rs_ton.RecordCount, GrdVT.tag)
    rs_ton.Close
    Set rs_ton = Nothing
    GrdVT.Row = 0
    TongTien
    Me.MousePointer = 0
End Sub

Private Sub TongTien()
    Dim duno As Double, duco As Double
    
    If CboKho.ListIndex >= 0 Then
        duno = SelectSQL("SELECT Sum(SoDuKhachHang.DuNo_0) As F1, Sum(SoDuKhachHang.DuCo_0) As F2" _
            & " FROM (SoDuKhachHang INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo) INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo" _
            & " WHERE PhanLoaiKhachHang.SoHieu LIKE '" + MaSo2SoHieu(CboKho.ItemData(CboKho.ListIndex), "PhanLoaiKhachHang") + "*'", duco)
    End If
    LbTien(0).Caption = Format(duno, Mask_0)
    LbTien(1).Caption = Format(duco, Mask_0)
End Sub


