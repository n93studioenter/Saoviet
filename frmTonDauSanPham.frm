VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmTonDauSanPham 
   BackColor       =   &H00FFFFFF&
   Caption         =   "frmTonDauSanPham"
   ClientHeight    =   825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   Icon            =   "frmTonDauSanPham.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   825
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chon File"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   240
      Width           =   1275
   End
   Begin VB.ComboBox CboLoai 
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   4470
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      Caption         =   "Load"
      Height          =   495
      Index           =   0
      Left            =   3180
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin MSGrid.Grid Grid 
      Height          =   195
      Left            =   5940
      TabIndex        =   0
      Top             =   660
      Width           =   3570
      _Version        =   65536
      _ExtentX        =   6297
      _ExtentY        =   344
      _StockProps     =   77
      Rows            =   1
      Cols            =   7
      FixedRows       =   0
   End
End
Attribute VB_Name = "frmTonDauSanPham"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command_Click()
'Dim pDataPath As String, fn As String
' pDataPath = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.xlsx", 1)
' If Len(pDataPath) > 0 Then
' Set xlapp = CreateObject("Excel.Application")
' xlapp.Workbooks.Open pDataPath ' + "bangkera.XLS"
' Set xlsheet = xlapp.Worksheets(1)
' Dim sodong, MaSo, i As Integer
' Dim sql As String
' 'sodong = Int(xlsheet.Cells(4, 2))
' Dim MaTaiKhoan As String
' Dim loai As Integer
' Dim sohieu, Ten, donvitinh, soluong, dongia, thanhtien
' sodong = Int(xlsheet.Cells(3, 2))
'
' For i = 4 To sodong + 4
'  sohieu = xlsheet.Cells(i, 1)
'  Ten = xlsheet.Cells(i, 2)
'  donvitinh = xlsheet.Cells(i, 3)
'
'  soluong = Format(CStr(xlsheet.Cells(i, 4)), Mask_0)
'  dongia = Format(CStr(xlsheet.Cells(i, 5)), Mask_0)
'  thanhtien = Format(CStr(xlsheet.Cells(i, 6)), Mask_2)
'  MaTaiKhoan = xlsheet.Cells(i, 7)
'  If Len(sohieu) > 0 Then
'     Grid.AddItem CStr(sohieu) + Chr(9) + Ten + Chr(9) + donvitinh + Chr(9) + CStr(soluong) + Chr(9) + CStr(dongia) + Chr(9) + CStr(thanhtien) + Chr(9) + CStr(MaTaiKhoan)
'  End If
' Next
' Grid.RemoveItem (0)
' End If
'End Sub

Private Sub Command1_Click()
Dim NguyenTe As String
Dim sql As String
Dim SoLuong, dongia, ThanhTien, taikhoan, sohieu
 For i = 1 To Grid.Rows - 1
  FVTDauKy.Form_Activate
  
     FrmVattu.Command_Click (0)
     FrmVattu.CboLoai = CboLoai
     Grid.Row = i
     Grid.col = 0
     sohieu = Grid.Text 'vattu.SoHieu
     FrmVattu.txtVT(0).Text = sohieu
     Grid.Row = i
     Grid.col = 1
     FrmVattu.txtVT(1).Text = Grid.Text 'vattu.TenVattu
     Grid.Row = i
     Grid.col = 2
     FrmVattu.txtVT(2).Text = Grid.Text 'vattu.DonVi
     Grid.Row = i
     Grid.col = 3
    ' FrmVattu.TxtVT(3).Text = Grid.Text  ' Format(vattu.GiaHT, Mask_0)
     
     SoLuong = Grid.Text
     Grid.Row = i
     Grid.col = 4
     dongia = Grid.Text
     Grid.Row = i
     Grid.col = 5
     ThanhTien = Grid.Text
     Grid.Row = i
     Grid.col = 6
     taikhoan = Grid.Text
     
    FrmVattu.txtVT(6).Text = ".." 'vattu.GhiChu
    FrmVattu.txtVT(7).Text = 0 'CStr(vattu.VAT)
    FrmVattu.txtVT(8).Text = 0 'GiaBan1
    FrmVattu.txtVT(9).Text = 0 'GiaBan2
    FrmVattu.txtVT(10).Text = 0 'GiaBan3
    FrmVattu.txtVT(12).Text = 0 'CK
    FrmVattu.txtVT(13).Text = 0 'ThueNK
    FrmVattu.txtTon(0).Text = 0 'TonMin
    FrmVattu.txtTon(1).Text = 0 'TonMax
    If Int(SelectSQL("SELECT MaSo AS F1 FROM vattu WHERE  SoHieu = '" + CStr(sohieu) + "'") + "0") <= 0 Then
      FrmVattu.Command_Click (1) ' luu vat tu do lai
    End If
    
    
    FVTDauKy.txtTon(0).Text = taikhoan
    FVTDauKy.txtTon_LostFocus (0)
    FVTDauKy.txtTon(1).Text = sohieu
    FVTDauKy.txtTon_LostFocus (1)
    
 FVTDauKy.txtTon(4).Text = SoLuong
 FVTDauKy.txtTon(5).Text = dongia
 FVTDauKy.txtTon(6).Text = ThanhTien
 FVTDauKy.CmdCt_Click
 Next
 MsgBox "Ban da luu ton dau xong"
End Sub

Private Sub Command22_Click(Index As Integer)

 
  Dim sql As String
 Dim pDataPath As String, fn As String
 Dim thoat
 thoat = 0
 pDataPath = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.xlsx", 1)
 If MsgBox("B¹n chon Yes th× cËp nhËt míi toµn bé" & vbNewLine & "B¹n chän No cËp nhËt bæ sung thªm vµo danh s¸ch.", vbYesNo + vbCritical, App.ProductName) = vbYes Then
    sql = " Update [Tonkho] b  "
        sql = sql + " SET b.tien_0 = 0"
        sql = sql + " ,b.Luong_0 = 0"
    ExecuteSQL5 sql
 End If
  
 If Len(pDataPath) > 0 Then
 Set xlapp = CreateObject("Excel.Application")
 xlapp.Workbooks.Open pDataPath ' + "bangkera.XLS"
 
 Set xlsheet = xlapp.Worksheets(1)
 Dim sodong, MaSo, i As Integer
 'Dim sql As String
 Dim MaTaiKhoan As String
 Dim loai As Integer
 Dim masovattu As String
 Dim sohieu, Ten, donvitinh, SoLuong, dongia, ThanhTien
 sodong = Int(xlsheet.Cells(3, 2))
  
   For i = 5 To sodong + 4
        sohieu = CStr(xlsheet.Cells(i, 1))
        Ten = CStr(xlsheet.Cells(i, 2))
        donvitinh = CStr(xlsheet.Cells(i, 3))
        MaTaiKhoan = xlsheet.Cells(i, 7)
        If (Len(sohieu) <= 0 Or Len(Ten) <= 0 Or Len(donvitinh) <= 0 Or Len(MaTaiKhoan) <= 0) Then
         MsgBox " B¹n ph¶i nhËp ®ñ th«ng tin  file exel: m· s¶n phÈm , tªn, ®¬n vÞ tÝnh, m· tµi kho¶n!"
        xlapp.Workbooks.Close
        thoat = 1
   End If
  Next
If thoat = 0 Then
 For i = 5 To sodong + 4
  FVTDauKy.Form_Activate
 

 
  FrmVattu.CboLoai = CboLoai
  
  sohieu = CStr(xlsheet.Cells(i, 1))
  Ten = CStr(xlsheet.Cells(i, 2))
  donvitinh = CStr(xlsheet.Cells(i, 3))
  
 
  
  SoLuong = CStr(xlsheet.Cells(i, 4))
  dongia = CStr(xlsheet.Cells(i, 5))
  ThanhTien = CStr(xlsheet.Cells(i, 6))
  MaTaiKhoan = xlsheet.Cells(i, 7)
  
    FrmVattu.txtVT(6).Text = ".." 'vattu.GhiChu
    FrmVattu.txtVT(7).Text = 0 'CStr(vattu.VAT)
    FrmVattu.txtVT(8).Text = 0 'GiaBan1
    FrmVattu.txtVT(9).Text = 0 'GiaBan2
    FrmVattu.txtVT(10).Text = 0 'GiaBan3
    FrmVattu.txtVT(12).Text = 0 'CK
    FrmVattu.txtVT(13).Text = 0 'ThueNK
    FrmVattu.txtTon(0).Text = 0 'TonMin
    FrmVattu.txtTon(1).Text = 0 'TonMax
    If Int(SelectSQL("SELECT count(MaSo) AS F1 FROM vattu WHERE  SoHieu = '" + CStr(sohieu) + "'")) <= 0 Then
      FrmVattu.Command_Click (0)
      FrmVattu.txtVT(0).Text = sohieu
      FrmVattu.txtVT(1).Text = Ten
      FrmVattu.txtVT(2).Text = donvitinh
      FrmVattu.Command_Click (1) ' luu vat tu do lai
    End If
    
    
    FVTDauKy.txtTon(0).Text = MaTaiKhoan
    FVTDauKy.txtTon_LostFocus (0)
    FVTDauKy.txtTon(1).Text = sohieu
    FVTDauKy.txtTon_LostFocus (1)

    FVTDauKy.txtTon(4).Text = SoLuong
    FVTDauKy.txtTon(5).Text = dongia
    FVTDauKy.txtTon(6).Text = ThanhTien
    FVTDauKy.CmdCt_Click

 Next
    KiemTraVatTu
    KiemTraTaiKhoan
  xlapp.Workbooks.Close
  frmTonDauSanPham.Hide
    End If
 End If
 
End Sub

Private Sub Form_Load()
  ColumnSetUp Grid, 0, 1000, 1
  ColumnSetUp Grid, 1, 5000, 1
  ColumnSetUp Grid, 2, 1000, 1
  ColumnSetUp Grid, 3, 1000, 1
  ColumnSetUp Grid, 4, 2000, 1
  ColumnSetUp Grid, 5, 2000, 1
  ColumnSetUp Grid, 6, 2000, 1
    Int_RecsetToCbo "SELECT DISTINCTROW PhanLoaiVattu.MaSo As F2, PhanLoaiVattu.SoHieu + ' - '  + PhanLoaiVattu.TenPhanLoai As F1 FROM PhanLoaiVattu WHERE PLCon=0 ORDER BY PhanLoaiVattu.SoHieu", CboLoai
End Sub
