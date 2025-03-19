VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FKHDauKy 
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
   Icon            =   "Fkhdauky.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Liability Opening Balance"
   Begin VB.CommandButton importexel 
      BackColor       =   &H80000009&
      Caption         =   "CËp nhË tõ excel"
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton importdatabase 
      BackColor       =   &H80000009&
      Caption         =   "CËp nhËt tõ data"
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton xoa 
      BackColor       =   &H80000009&
      Caption         =   "Xãa tån ®Çu kú"
      Height          =   375
      Left            =   6840
      TabIndex        =   20
      Top             =   6600
      Width           =   1335
   End
   Begin MSGrid.Grid GrdVT 
      Height          =   5535
      Left            =   120
      TabIndex        =   15
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
      Cols            =   8
      FixedRows       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   6840
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
      Left            =   8160
      MaxLength       =   20
      TabIndex        =   6
      Text            =   "0"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdct 
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      Picture         =   "Fkhdauky.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   255
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   5520
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "0"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Left            =   8400
      Picture         =   "Fkhdauky.frx":5B84
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6240
      Width           =   3135
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1080
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
      Width           =   975
   End
   Begin VB.ComboBox CboKho 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   6915
      TabIndex        =   19
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
      Left            =   8160
      TabIndex        =   18
      Tag             =   "Foreign Currency"
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
      Left            =   6840
      TabIndex        =   17
      Tag             =   "Credit"
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
      TabIndex        =   16
      Tag             =   "Account"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tæng tiÒn"
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   14
      Tag             =   "Total"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   5475
      TabIndex        =   13
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
      Left            =   5520
      TabIndex        =   12
      Tag             =   "Debit"
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
      Left            =   2400
      TabIndex        =   11
      Tag             =   "Description"
      Top             =   480
      Width           =   3135
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
      Left            =   1080
      TabIndex        =   10
      Tag             =   "Liability Code"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ph©n lo¹i"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Tag             =   "Class"
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FKHDauKy"
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
        ErrMsg er_PhanLoai
        Exit Sub
    End If
    
    If taikhoan.MaSo = 0 Or taikhoan.tkcon > 0 Then
        ErrMsg er_SHTaiKhoan1
        RFocus txtTon(0)
        Exit Sub
    End If
        
    If taikhoan.tk_id <> TKCNKH_ID And taikhoan.tk_id <> TKCNPT_ID Then
        ErrMsg er_SHTKCN
        RFocus txtTon(0)
        Exit Sub
    End If
        
    If ckh.MaSo = 0 Then
        ErrMsg er_SHKhachHang
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
        For i = 0 To .Rows - 1
            .col = 6
            .Row = i
            If Len(.Text) = 0 Then Exit For
            If CLng5(.Text) = taikhoan.MaSo Then
                .col = 7
                If CLng5(.Text) = ckh.MaSo Then
                    If Len(psw) > 0 Then
                        If FPsw.GetPswX() <> psw Then GoTo XongDK
                    End If
                    GrdVT.RemoveItem i
                    GrdVT.AddItem taikhoan.sohieu + Chr(9) + ckh.sohieu + Chr(9) + ckh.Ten + Chr(9) + Format(tien1, Mask_0) + Chr(9) + Format(tien2, Mask_0) + Chr(9) + Format(luong, Mask_2) + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) + CStr(ckh.MaSo), i
                    ckh.GhiDauKy taikhoan.MaSo, tien1, tien2, luong
                    TongTien
                    RFocus txtTon(0)
                    GoTo XongDK
                End If
            End If
        Next
        
        .AddItem taikhoan.sohieu + Chr(9) + ckh.sohieu + Chr(9) + ckh.Ten + Chr(9) + Format(tien1, Mask_0) + Chr(9) + Format(tien2, Mask_0) + Chr(9) + Format(luong, Mask_2) + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) + CStr(ckh.MaSo), NewRowIndex(GrdVT, 0)
        ckh.GhiDauKy taikhoan.MaSo, tien1, tien2, luong
        .Row = .Rows - 1
        .col = 0
        If Len(.Text) = 0 Then .RemoveItem .Row
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbAltMask) > 0 And KeyCode = vbKeyV) Or KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    ColumnSetUp GrdVT, 0, 940, 2
    ColumnSetUp GrdVT, 1, 1300, 2
    ColumnSetUp GrdVT, 2, 3100, 0
    ColumnSetUp GrdVT, 3, 1300, 1
    ColumnSetUp GrdVT, 4, 1300, 1
    ColumnSetUp GrdVT, 5, 1300, 1
    ColumnSetUp GrdVT, 6, 1, 0
    ColumnSetUp GrdVT, 7, 1, 0
    
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
    Dim i As Integer
    
    With GrdVT
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        For i = 0 To 5
            .col = i
            txtTon(i).Text = .Text
        Next
        .col = 0
        txtTon_LostFocus 0
        txtTon_LostFocus 1
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

Private Sub importdatabase_Click()
Dim psw As String, fn As String
Dim sql
  Dim rs_chungtu As Recordset
  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
If Len(psw) > 0 Then
        sql = "insert into khachhang select * from [" + psw + ";PWD=" + pPSW + "].khachhang where maso not in (select maso from khachhang)"
        DBKetoan.Execute sql
        sql = "insert into sodukhachhang select * from [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang where maso not in (select maso from SoDuKhachHang )"
        DBKetoan.Execute sql
        sql = " UPDATE SoDuKhachHang INNER JOIN [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang a "
        sql = sql + " on SoDuKhachHang.maso = a.maso "
        sql = sql + " set SoDuKhachHang.DuNo_0  = a.DuNo_12 "
        sql = sql + ",SoDuKhachHang.DuCo_0 = a.DuCo_12 "
        sql = sql + ",SoDuKhachHang.DuNT_0 = a.DuNT_12 "
        DBKetoan.Execute sql
         LietKeTonKho CboKho.ItemData(CboKho.ListIndex)
        MsgBox "B¹n chuyÓn d÷ liÖu thµnh c«ng."
End If
End Sub

Private Sub importexel_Click()
Dim pDataPath As String, fn As String
Dim xlapp
Dim xlsheet
Dim so
Dim sql As String
pDataPath = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.xlsx", 1)
 
 If Len(pDataPath) > 0 Then
 If MsgBox("B¹n chon Yes th× cËp nhËt míi toµn bé." & vbNewLine & "B¹n chän No cËp nhËt bæ sung thªm vµo danh s¸ch.", vbYesNo + vbCritical, App.ProductName) = vbYes Then
    sql = " UPDATE SoDuKhachHang "
    sql = sql + " set DuNo_0  = 0 "
    sql = sql + ",DuCo_0 = 0 "
    sql = sql + ",DuNT_0 = 0"
    ExecuteSQL5 sql
 End If
   LietKeTonKho CboKho.ItemData(CboKho.ListIndex)
 
 Set xlapp = CreateObject("Excel.Application")
 xlapp.Workbooks.Open pDataPath
 Set xlsheet = xlapp.Worksheets(1)
 Dim sodong, MaSo, i As Integer
 
 'sodong = Int(xlsheet.Cells(4, 2))
 Dim MaTaiKhoan As String
 Dim loai As Integer
 Dim sohieu, Ten, DiaChi, mst, Tel, Fax, email, taikhoan, DaiDien, GhiChu, duno, duco, NguyenTe As String
 sodong = Int(xlsheet.Cells(4, 2))
 For i = 6 To sodong + 5
  sohieu = CStr(xlsheet.Cells(i, 1))
  Ten = CStr(xlsheet.Cells(i, 2))
  DiaChi = CStr(xlsheet.Cells(i, 3)) + "..."
  mst = CStr(xlsheet.Cells(i, 4)) + "..."
  Tel = xlsheet.Cells(i, 5) + "..."
  Fax = xlsheet.Cells(i, 6) + "..."
  email = xlsheet.Cells(i, 7) + "..."
  taikhoan = xlsheet.Cells(i, 8) + "..."
  DaiDien = xlsheet.Cells(i, 9) + "..."
  GhiChu = xlsheet.Cells(i, 10) + "..."
  MaTaiKhoan = xlsheet.Cells(i, 11)
  If (i > 4) Then
  
  duno = CStr(xlsheet.Cells(i, 12))
  duco = CStr(xlsheet.Cells(i, 13))
  NguyenTe = CStr(xlsheet.Cells(i, 14))
  Else
  duno = "0"
  duco = "0"
  NguyenTe = "0"
  End If
   sql = ""
    If Left(MaTaiKhoan, 3) = "331" Then
        loai = 2
    ElseIf Left(MaTaiKhoan, 3) = "131" Then
        loai = 3
    Else
        loai = 1
    End If
    
    MaTaiKhoan = SelectSQL("SELECT MaSo AS F1 FROM hethongtk WHERE  SoHieu = '" + CStr(MaTaiKhoan) + "'")
    MaSo = SelectSQL("SELECT MaSo AS F1 FROM KhachHang WHERE  sohieu = '" + sohieu + "'")
    If MaSo > 0 Then
            sql = sql + "update khachhang set ten = '" + CStr(Ten) + "'"
            sql = sql + ",diachi = '" + CStr(DiaChi) + "..." + "'"
            sql = sql + ",MST = '" + CStr(mst) + "..." + "'"
            sql = sql + ",Tel = '" + CStr(Tel) + "..." + "'"
            sql = sql + ",Fax = '" + CStr(Fax) + "..." + "'"
            sql = sql + ",EMail = '" + CStr(email) + "..." + "'"
            sql = sql + ",DaiDien = '" + CStr(DaiDien) + "..." + "'"
            sql = sql + ",TaiKhoan = '" + CStr(taikhoan) + "'"
            sql = sql + ",GhiChu = '" + CStr(GhiChu) + "..." + "'"
            sql = sql + " where maso = " + CStr(MaSo) + ""
            ExecuteSQL5 sql
            If SelectSQL(" select count(*) as f1 from SoDuKhachHang  where mataikhoan = " + CStr(MaTaiKhoan) + " and MaKhachHang = " + CStr(MaSo)) > 0 Then
            
                ExecuteSQL5 "update SoDuKhachHang set DuNo_0 = " + CStr("0" + duno) + ",DuCo_0 = " + CStr("0" + duco) + ",DuNT_0 = " + CStr("0" + NguyenTe) + " where mataikhoan = " + CStr(MaTaiKhoan) + " and MaKhachHang = " + CStr(MaSo) + ""
            Else
                ExecuteSQL5 "INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(MaTaiKhoan) + "," + CStr(MaSo) + "," + CStr("0" + duno) + "," + CStr("0" + duco) + "," + CStr("0" + NguyenTe) + ")"
               End If

      Else
      If Len(sohieu) > 0 And Len(Ten) > 0 And MaTaiKhoan > 0 Then
          sql = sql + "INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu,Ten,DiaChi,MST,Tel,Fax,EMail,DaiDien,TaiKhoan,GhiChu) VALUES (" + CStr(Lng_MaxValue("MaSo", "KhachHang") + 1) + ","
          sql = sql + CStr(loai) + ",'" + sohieu + "','" + Ten + "','" + DiaChi + "','" + mst + "','" + Tel + "','" + Fax + "','" + email + "','" + DaiDien + "','"
          sql = sql + taikhoan + "','" + GhiChu + "')"
          ExecuteSQL5 sql
          ExecuteSQL5 "INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(MaTaiKhoan) + "," + CStr(Lng_MaxValue("MaSo", "KhachHang")) + "," + CStr("0" + duno) + "," + CStr("0" + duco) + "," + CStr("0" + NguyenTe) + ")"
          Else
          MsgBox "Kh«ng thÓ cËp nhËt ®èi t­îng :" + sohieu + " - " + Ten
      End If
   End If
  Next
   xlapp.Workbooks.Close
   LietKeTonKho CboKho.ItemData(CboKho.ListIndex)
    KiemTraVatTu
    KiemTraTaiKhoan
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
        Case 3, 4, 5:
            If KeyAscii = 13 Then
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
            If Len(txtTon(0).Text) > 0 Then
                taikhoan.InitTaikhoanSohieu txtTon(0).Text
            Else
                taikhoan.InitTaikhoanMaSo 0
            End If
        Case 1:
            If Len(txtTon(1).Text) > 0 Then
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
    Set rs_ton = DBKetoan.OpenRecordset("SELECT HethongTK.MaSo,HethongTK.Kieu,HethongTK.SoHieu AS SHTK,SoDuKhachHang.MaKhachHang,KhachHang.SoHieu,KhachHang.Ten," _
        & " SoDuKhachHang.DuNo_0 AS DuNo,SoDuKhachHang.DuCo_0 AS DuCo,SoDuKhachHang.DuNT_0 AS DuNT " _
        & " FROM ((SoDuKhachHang INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang = KhachHang.MaSo) INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo) INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo" _
        & " WHERE PhanLoaiKhachHang.MaSo = " + CStr(mkho) + " AND (SoDuKhachHang.DuNo_0 <> 0 OR SoDuKhachHang.DuCo_0 <> 0) ORDER BY HethongTK.SoHieu DESC, KhachHang.SoHieu DESC", dbOpenSnapshot)
    Do While Not rs_ton.EOF
        GrdVT.AddItem rs_ton!shtk + Chr(9) + rs_ton!sohieu + Chr(9) + rs_ton!Ten + Chr(9) + Format(rs_ton!duno, Mask_0) + Chr(9) + Format(rs_ton!duco, Mask_0) + Chr(9) + Format(rs_ton!dunt, Mask_2) + Chr(9) + CStr(rs_ton!MaSo) + Chr(9) + CStr(rs_ton!MaKhachHang), 0
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
            & " WHERE PhanLoaiKhachHang.MaSo =" + CStr(CboKho.ItemData(CboKho.ListIndex)), duco)
    End If
    LbTien(0).Caption = Format(duno, Mask_0)
    LbTien(1).Caption = Format(duco, Mask_0)
End Sub

Private Sub xoa_Click()
Dim sql
If MsgBox("B¹n cã ch¾c ch¾n xãa tån ®Çu kh«ng?", vbYesNo + vbCritical, App.ProductName) = vbYes Then

sql = " UPDATE SoDuKhachHang "
sql = sql + " set DuNo_0  = 0 "
sql = sql + ",DuCo_0 = 0 "
sql = sql + ",DuNT_0 = 0"
DBKetoan.Execute sql
         LietKeTonKho CboKho.ItemData(CboKho.ListIndex)
                DBKetoan.Execute "update hethongtk set duno_0 = 0,duco_0 = 0 where sohieu like '331*'"
       DBKetoan.Execute "update hethongtk set duno_0 = 0,duco_0 = 0 where sohieu like '131*'"
MsgBox "Xãa thµnh c«ng"
 End If
End Sub
