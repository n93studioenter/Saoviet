VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmTonDauKhachHang 
   Caption         =   "frmTonDauKhachHang"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   Icon            =   "frmTonDauKhachHang.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6570
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      Caption         =   "Load"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin MSGrid.Grid Grid1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      _Version        =   65536
      _ExtentX        =   22886
      _ExtentY        =   10186
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
Attribute VB_Name = "frmTonDauKhachHang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()
Dim pDataPath As String, fn As String
 pDataPath = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.xlsx", 1)
 If Len(pDataPath) > 0 Then
 Set xlapp = CreateObject("Excel.Application")
 xlapp.Workbooks.Open pDataPath
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
  If (i > 4) Then
  
  duno = Format(CStr(xlsheet.Cells(i, 12)), Mask_0)
  duco = Format(CStr(xlsheet.Cells(i, 13)), Mask_0)
  NguyenTe = Format(CStr(xlsheet.Cells(i, 14)), Mask_2)
  Else
  duno = "0"
  duco = "0"
  NguyenTe = "0"
  End If
  
  If Len(sohieu) > 0 Then
     Grid1.AddItem CStr(sohieu) + Chr(9) + Ten + Chr(9) + DiaChi + Chr(9) + mst + Chr(9) + Tel + Chr(9) + Fax + Chr(9) + email + Chr(9) + taikhoan + Chr(9) + DaiDien + Chr(9) + GhiChu + Chr(9) + CStr(MaTaiKhoan) + Chr(9) + CStr(duno) + Chr(9) + CStr(duco) + Chr(9) + CStr(NguyenTe)
  End If
  Next
 
  Grid1.RemoveItem (0)
End If
End Sub

Private Sub Command1_Click()
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
               ExecuteSQL5 "update SoDuKhachHang set DuNo_0 = " + CStr("0" + duno) + ",DuCo_0 = " + CStr("0" + duco) + ",DuNT_0 = " + CStr("0" + NguyenTe) + " where mataikhoan = " + CStr(MaTaiKhoan) + " and MaKhachHang = " + CStr(MaSo) + ""

      Else
          sql = sql + "INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu,Ten,DiaChi,MST,Tel,Fax,EMail,DaiDien,TaiKhoan,GhiChu) VALUES (" + CStr(Lng_MaxValue("MaSo", "KhachHang") + 1) + ","
          sql = sql + CStr(loai) + ",'" + sohieu + "','" + "Ten" + "','" + "DiaChi" + "','" + "MST" + "','" + "Tel" + "','" + "Fax" + "','" + "EMail" + "','" + "DaiDien" + "','"
          sql = sql + "TaiKhoan" + "','" + "GhiChu" + "')"
          ExecuteSQL5 sql
          ExecuteSQL5 "INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(MaTaiKhoan) + "," + CStr(Lng_MaxValue("MaSo", "KhachHang")) + "," + CStr("0" + duno) + "," + CStr("0" + duco) + "," + CStr("0" + NguyenTe) + ")"
   End If
   Next
   MsgBox "Luu thanh cong"
End Sub
