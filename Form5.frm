VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12570
   LinkTopic       =   "Form3"
   ScaleHeight     =   7050
   ScaleWidth      =   12570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command 
      Caption         =   "Command2"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin MSGrid.Grid Grid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      _Version        =   65536
      _ExtentX        =   21405
      _ExtentY        =   11033
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   1
      Cols            =   14
      FixedRows       =   0
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()

Dim NguyenTe As String
Dim sql As String
 For i = 1 To Grid1.Rows - 1
     Grid1.Row = i
     Grid1.col = 0
     sohieu = Grid1.Text
 '     MsgBox "s0hieu " + SoHieu
     Grid1.Row = i
     Grid1.col = 1
     Ten = Grid1.Text
     
     Grid1.Row = i
     Grid1.col = 2
     DiaChi = Grid1.Text
     
     Grid1.Row = i
     Grid1.col = 3
     mst = Grid1.Text
     
     Grid1.Row = i
     Grid1.col = 4
     Tel = Grid1.Text
     
     Grid1.Row = i
     Grid1.col = 5
     Fax = Grid1.Text
     
     Grid1.Row = i
     Grid1.col = 6
     email = Grid1.Text
     
     Grid1.Row = i
     Grid1.col = 7
     taikhoan = Grid1.Text
     
     Grid1.Row = i
     Grid1.col = 8
     DaiDien = Grid1.Text
      
     Grid1.Row = i
     Grid1.col = 9
     GhiChu = Grid1.Text
     
     Grid1.Row = i
     Grid1.col = 10
     MaTaiKhoan = Grid1.Text
    
    ' MsgBox MaTaiKhoan
     Grid1.Row = i
     Grid1.col = 11
     duno = Grid1.Text
     
     Grid1.Row = i
     Grid1.col = 12
     duco = Grid1.Text
     
     Grid1.Row = i
     Grid1.col = 13
     NguyenTe = Grid1.Text
     
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
            sql = sql + "update khachhang set ten = '" + Ten + "'"
            sql = sql + ",diachi = '" + DiaChi + "'"
            sql = sql + ",MST = '" + mst + "'"
            sql = sql + ",Tel = '" + Tel + "'"
            sql = sql + ",Fax = '" + Fax + "'"
            sql = sql + ",EMail = '" + email + "'"
            sql = sql + ",DaiDien = '" + DaiDien + "'"
            sql = sql + ",TaiKhoan = '" + taikhoan + "'"
            sql = sql + ",GhiChu = '" + GhiChu + "'"
            sql = sql + " where maso = " + CStr(MaSo) + ""
            ExecuteSQL5 sql
               ExecuteSQL5 "update SoDuKhachHang set DuNo_0 = " + CStr(duno) + ",DuCo_0 = " + CStr(duco) + ",DuNT_0 = " + CStr(NguyenTe) + " where mataikhoan = " + CStr(MaTaiKhoan) + " and MaKhachHang = " + CStr(MaSo) + ""

      Else
          sql = sql + "INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu,Ten,DiaChi,MST,Tel,Fax,EMail,DaiDien,TaiKhoan,GhiChu) VALUES (" + CStr(Lng_MaxValue("MaSo", "KhachHang") + 1) + ","
          sql = sql + CStr(loai) + ",'" + sohieu + "','" + "Ten" + "','" + "DiaChi" + "','" + "MST" + "','" + "Tel" + "','" + "Fax" + "','" + "EMail" + "','" + "DaiDien" + "','"
          sql = sql + "TaiKhoan" + "','" + "GhiChu" + "')"
          ExecuteSQL5 sql
          ExecuteSQL5 "INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + MaTaiKhoan + "," + CStr(Lng_MaxValue("MaSo", "KhachHang")) + "," + duno + "," + duco + "," + NguyenTe + ")"
   End If
   Next
   MsgBox "Luu thanh cong"
End Sub

Private Sub Command1_Click()
 Dim pDataPath As String, fn As String
 pDataPath = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "gg.xlsx", 1)
 
 Set xlapp = CreateObject("Excel.Application")
 xlapp.Workbooks.Open pDataPath ' + "bangkera.XLS"
 Set xlsheet = xlapp.Worksheets(1)
 Dim sodong, MaSo, i As Integer
 Dim sql As String
 'sodong = Int(xlsheet.Cells(4, 2))
 Dim MaTaiKhoan As String
 Dim loai As Integer
 Dim sohieu, Ten, DiaChi, mst, Tel, Fax, email, taikhoan, DaiDien, GhiChu, duno, duco, NguyenTe
 sodong = Int(xlsheet.Cells(4, 2))
 For i = 5 To sodong + 6
  sohieu = xlsheet.Cells(i, 1)
  Ten = xlsheet.Cells(i, 2)
  DiaChi = xlsheet.Cells(i, 3)
  mst = xlsheet.Cells(i, 4)
  Tel = xlsheet.Cells(i, 5) + "..."
  Fax = xlsheet.Cells(i, 6) + "..."
  email = xlsheet.Cells(i, 7) + "..."
  taikhoan = xlsheet.Cells(i, 8) + "..."
  DaiDien = xlsheet.Cells(i, 9) + "..."
  GhiChu = xlsheet.Cells(i, 10) + "..."
  MaTaiKhoan = xlsheet.Cells(i, 11)
  duno = Format(xlsheet.Cells(i, 12), Mask_0)
  duco = Format(xlsheet.Cells(i, 13), Mask_0)
  NguyenTe = Format(xlsheet.Cells(i, 14), Mask_2)
  
  If Len(sohieu) > 0 Then
     Grid1.AddItem sohieu + Chr(9) + Ten + Chr(9) + DiaChi + Chr(9) + mst + Chr(9) + Tel + Chr(9) + Fax + Chr(9) + email + Chr(9) + taikhoan + Chr(9) + DaiDien + Chr(9) + GhiChu + Chr(9) + CStr(MaTaiKhoan) + Chr(9) + CStr(duno) + Chr(9) + CStr(duco) + Chr(9) + CStr(NguyenTe)
  End If
 Next
 
  Grid1.RemoveItem (0)

End Sub

  ' If ExecuteSQL5("INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(rs_taikhoan!MaTkCo) + "," + CStr(rs_taikhoan!MaKHC) + ")", False) <> 0 Then GoTo d
  '  If SelectSQL("SELECT MaSo AS F1 FROM KhachHang WHERE MST='" + .MST + "' AND MaSo<>" + CStr(.MaSo)) > 0 Then
       ' If MsgBox("M· sè thu? ®· c?, cho ph?p nhËp?", vbYesNo + vbCritical, App.ProductName) = vbNo Then Exit Function
            
' FrmKhachHang.txtVT(0).Text = SoHieu + "..."
'    FrmKhachHang.txtVT(1).Text = Ten + "..."
'    FrmKhachHang.txtVT(2).Text = DiaChi + "..."
'    FrmKhachHang.txtVT(3).Text = MST + "..."
'    FrmKhachHang.txtVT(4).Text = Tel + "..."
'    FrmKhachHang.txtVT(5).Text = Fax + "..."
'    FrmKhachHang.txtVT(6).Text = EMail + "..."
'    FrmKhachHang.txtVT(7).Text = DaiDien + "..."
'    FrmKhachHang.txtVT(8).Text = TaiKhoan + "..."
'    FrmKhachHang.txtVT(9).Text = GhiChu + "..."
'    FrmKhachHang.txtVT(10).Text = 0
'    FrmKhachHang.CboNT.ListIndex = 2

Private Sub Form_Load()
  ColumnSetUp Grid1, 0, 1000, 2
  ColumnSetUp Grid1, 1, 4000, 2
  ColumnSetUp Grid1, 2, 1000, 2
  ColumnSetUp Grid1, 3, 1000, 2
  ColumnSetUp Grid1, 4, 1000, 2
  ColumnSetUp Grid1, 5, 1000, 2
  ColumnSetUp Grid1, 6, 1000, 2
  ColumnSetUp Grid1, 7, 1000, 2
  ColumnSetUp Grid1, 8, 1000, 2
  ColumnSetUp Grid1, 9, 1000, 2
  ColumnSetUp Grid1, 10, 1000, 2
  ColumnSetUp Grid1, 11, 1000, 2
  ColumnSetUp Grid1, 12, 1000, 2
  ColumnSetUp Grid1, 13, 1000, 2
End Sub
