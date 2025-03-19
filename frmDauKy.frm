VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form3 
   Caption         =   "Tån ®Çu kú"
   ClientHeight    =   1890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7305
   LinkTopic       =   "Form3"
   ScaleHeight     =   1890
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Taøi khoaûn"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Xoùa toàn ñaàu TSCD"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TSCD"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      ToolTipText     =   "Tep chuyen phai la dau nam tai chinh"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Coâng nôï"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Saûn phaåm"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox suasolieucu 
      Caption         =   "Caäp nhaät toân ñaàu coù saün"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin MSForms.OptionButton chon 
      Height          =   615
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2415
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "4260;1085"
      Value           =   "0"
      Caption         =   "Chuyeån muïc chöa coù"
      FontName        =   "VNI-Times"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton chon 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "2990;661"
      Value           =   "1"
      Caption         =   "Chuyeån môùi"
      FontName        =   "VNI-Times"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chon_Click(Index As Integer)
If (Index = 1) Then
suasolieucu.Enabled = True
Else
suasolieucu.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Dim psw As String, fn As String
  Dim rs_chungtu As Recordset
  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
  Dim sql, st
  Dim i As Integer
  If Len(psw) > 0 Then
  LockDB
 ExecuteSQL5_Themmoi ("ALTER TABLE phanloaivattu  ADD MaPhanLoaiCu Number")
 DBKetoan.Execute "update phanloaivattu set MaPhanLoaiCu = 0"
  sql = "insert into phanloaivattu select * ,maso as MaPhanLoaiCu from [" + psw + ";PWD=" + pPSW + "].phanloaivattu where maso  not in (select maso from phanloaivattu) "
  DBKetoan.Execute sql
  
   ExecuteSQL5_Themmoi ("ALTER TABLE dvtvattu  ADD MadvtvattuCu Number")
   DBKetoan.Execute "update dvtvattu set MadvtvattuCu = 0"
   sql = "insert into dvtvattu select * ,maso as MadvtvattuCu from [" + psw + ";PWD=" + pPSW + "].dvtvattu " 'where sohieu  not in (select sohieu from dvtvattu) "
  DBKetoan.Execute sql
  
  ' truong hop them moi
  If (chon(0).Value = True) Then
            them_moi_san_pham psw, pPSW
  End If
  ' ke thuc truong hop them moi
  'truong hop sua cac muc da co
   If (chon(0).Value = True) Then
            Them_san_pham_chua_co psw, pPSW
  End If
  UnlockDB
  kiem_tra_nhap_xuat_ton_va_tai_khoan
    ' ket thuc sua muc da co
  MsgBox "B¹n ®· chuyÓn d÷ liÖu ®Çu kú thµnh c«ng."
  Dim sttt As String
  End If

End Sub
Public Sub Them_san_pham_chua_co(psw As String, pPSW As String) ' them nhung san pham chua co
                ExecuteSQL5_Themmoi ("ALTER TABLE vattu  ADD sanphamIdCu Number")
                  DBKetoan.Execute "update vattu set sanphamIdCu = 0"
                  sql = "insert into vattu select MaPhanLoai,SoHieu,TenVattu,DonVi,TonMin,TonMax,GiaDuPhong,GiaHT, dvt2,DonVi2,TyLeQD,GhiChu,VAT,GiaBan1,GiaBan2,GiaBan3,CK,ThueNK,L,T,MaSo as sanphamIdCu from [" + psw + ";PWD=" + pPSW + "].vattu where maso  not in (select maso from vattu) "
                  DBKetoan.Execute sql
                  
                  ExecuteSQL5_Themmoi ("ALTER TABLE tonkho  ADD sanphamIdCu Number")
                  DBKetoan.Execute "update tonkho set sanphamIdCu = 0"
                  sql = "insert into [Tonkho] (masokho,mataikhoan,mavattu,Tien_0,Luong_0) select tk.masokho,tk.mataikhoan,vattu.maso as mavattu,tk.Tien_12 as Tien_0,tk.Luong_12 as Luong_0 from  [" + psw + ";PWD=" + pPSW + "].tonkho  tk inner join  vattu on vattu.sanphamIdCu  = tk.mavattu"
                  DBKetoan.Execute sql
                  
                  If (suasolieucu.Value = 1) Then
                     sql = " Update [Tonkho] b INNER JOIN [" + psw + ";PWD=" + pPSW + "].TonKho a ON b.mavattu = a.mavattu and b.MaSoKho = a.MaSoKho and   b.MaTaiKhoan = a.MaTaiKhoan"
                     sql = sql + " SET b.tien_0 = a.tien_12"
                     sql = sql + " ,b.Luong_0 = a.Luong_12"
                 DBKetoan.Execute sql
                End If
End Sub

Public Sub them_moi_san_pham(psw As String, pPSW As String) ' them moi hoan toan
            If (chon(0).Value = True) Then
              ExecuteSQL5_Themmoi ("ALTER TABLE vattu  ADD sanphamIdCu Number")
              DBKetoan.Execute "update vattu set sanphamIdCu = 0"
              sql = "insert into vattu select MaPhanLoai,SoHieu,TenVattu,DonVi,TonMin,TonMax,GiaDuPhong,GiaHT, dvt2,DonVi2,TyLeQD,GhiChu,VAT,GiaBan1,GiaBan2,GiaBan3,CK,ThueNK,L,T,MaSo as sanphamIdCu from [" + psw + ";PWD=" + pPSW + "].vattu "
              DBKetoan.Execute sql
              
              ExecuteSQL5_Themmoi ("ALTER TABLE tonkho  ADD sanphamIdCu Number")
              DBKetoan.Execute "update tonkho set sanphamIdCu = 0"
              sql = "insert into [Tonkho] (masokho,mataikhoan,mavattu,Tien_0,Luong_0) select tk.masokho,tk.mataikhoan,vattu.maso as mavattu,tk.Tien_12 as Tien_0,tk.Luong_12 as Luong_0 from  [" + psw + ";PWD=" + pPSW + "].tonkho  tk inner join  vattu on vattu.sanphamIdCu  = tk.mavattu"
              DBKetoan.Execute sql
              End If
End Sub
Private Sub Command2_Click()
Dim psw As String, fn As String
Dim sql
  Dim rs_chungtu As Recordset
  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
If Len(psw) > 0 Then
      LockDB
        ExecuteSQL5_Themmoi ("ALTER TABLE khachhang  ADD khachhangidCu Number")
        DBKetoan.Execute "update khachhang set khachhangidCu = 0"
        sql = "insert into khachhang(MaPhanLoai , sohieu,  Ten, DiaChi, mst, Tel, Fax, email, taikhoan, DaiDien, DuMax, GhiChu, MaNT,khachhangidCu) select MaPhanLoai , sohieu,  Ten, DiaChi, mst, Tel, Fax, email, taikhoan, DaiDien, DuMax, GhiChu, MaNT, Maso as khachhangidCu from [" + psw + ";PWD=" + pPSW + "].khachhang "
        If chon(1).Value = True Then
        sql = sql + " where maso not in (select maso from khachhang) "
        End If
        DBKetoan.Execute sql
        sql = " insert into sodukhachhang(MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0) select  sodu.MaTaiKhoan,khachhang.Maso as MaKhachHang,sodu.DuNo_0 as DuNo_12,sodu.DuCo_0 as DuCo_12,sodu.DuNT_0 as DuNT_12  from [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang sodu  inner join khachhang on sodu.MaKhachHang = khachhang.khachhangidCu"
        DBKetoan.Execute sql
        
        
      ' neu la sua lai so du dau ky
      If (suasolieucu.Value = 1) Then
        sql = " UPDATE SoDuKhachHang INNER JOIN [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang a "
        sql = sql + " on SoDuKhachHang.maso = a.maso "
        sql = sql + " set SoDuKhachHang.DuNo_0  = a.DuNo_12 "
        sql = sql + ",SoDuKhachHang.DuCo_0 = a.DuCo_12 "
        sql = sql + ",SoDuKhachHang.DuNT_0 = a.DuNT_12 "
        DBKetoan.Execute sql
        End If
      ' ket thuc sua so du dau ky
      kiem_tra_nhap_xuat_ton_va_tai_khoan
      UnlockDB
        MsgBox "B¹n chuyÓn d÷ liÖu thµnh c«ng."
End If
End Sub

Private Sub Command3_Click()
  Dim rs_chungtu As Recordset
  Dim sql
Dim psw As String
  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
If Len(psw) > 0 Then

'     Dua tai san co dinh qua
LockDB


 DBKetoan.Execute "insert into LoaiTaiSan select * from [" + psw + ";PWD=" + pPSW + "].LoaiTaiSan"
 
ExecuteSQL5_Themmoi ("ALTER TABLE Taisan  ADD taisanIdCu Number")
        DBKetoan.Execute "update Taisan set taisanIdCu = 0"
 ExecuteSQL5 " insert into Taisan(sohieu , Ten, NangLuc, GhiChu, MaNuoc, MaTaiKhoan, maloai, MaNhom, ThangTang, ThangGiam, NamSX, NamSD, NamKH, shct, NCT, taisanIdCu) select  sohieu , Ten, NangLuc, GhiChu, MaNuoc, MaTaiKhoan, maloai, MaNhom, ThangTang, ThangGiam, NamSX, NamSD, NamKH, shct, NCT, maso as taisanIdCu  from [" + psw + ";PWD=" + pPSW + "].Taisan "
 
 ExecuteSQL5_Themmoi ("ALTER TABLE ThongSo  ADD taisanIdCu Number")
        DBKetoan.Execute "update ThongSo set taisanIdCu = 0"
        
ExecuteSQL5 "insert into  ThongSo(MaTS , thang, KH_NS, KH_TBS, KH_CNK, KH_TD, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD, MaDTQL, MaDTSD, MaTTSD,taisanIdCu) select taisan.maso as MaTS,Thang,KH_NS,KH_TBS,KH_CNK,KH_TD,NG_NS,NG_TBS,NG_CNK,NG_TD,CL_NS,CL_TBS,CL_CNK,CL_TD,MaDTQL,MaDTSD,MaTTSD ,taisan.maso as taisanIdCu from [" + psw + ";PWD=" + pPSW + "].thongso ts inner join taisan on ts.mats = taisan.taisanIdCu"
      
   ExecuteSQL5_Themmoi ("ALTER TABLE CTTaiSan  ADD taisanIdCu Number")
        DBKetoan.Execute "update CTTaiSan set taisanIdCu = 0"
      
      sql = "insert into CTTaiSan (MaSo, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD,taisanIdCu) select  ctts.MaSo, ctts.SoHieu, ctts.Thang, ctts.VaoSo, ctts.NgayGhi, ctts.DienGiai, "
      sql = sql + "          ctts.MaLoai, ctts.MaNhom, taisan.maso as MaTS, ctts.NG_NS, ctts.NG_TBS, ctts.NG_CNK, ctts.NG_TD, ctts.CL_NS, ctts.CL_TBS, ctts.CL_CNK, ctts.CL_TD,ctts.mats as taisanIdCu "
       sql = sql + "    From [" + psw + ";PWD=" + pPSW + "].CTTaiSan ctts inner join Taisan on ctts.MATS = taisan.taisanIdCu"
      DBKetoan.Execute sql
 
 
 ExecuteSQL5 "UPDATE DISTINCTROW ThongSo LEFT JOIN ThongSoCuoiKy ON ThongSo.MaTS = ThongSoCuoiKy.MaTS SET ThongSo.KH_NS = ThongSoCuoiKy.KH_NS, ThongSo.KH_TBS = ThongSoCuoiKy.KH_TBS, ThongSo.KH_CNK = ThongSoCuoiKy.KH_CNK, ThongSo.KH_TD = ThongSoCuoiKy.KH_TD, ThongSo.MaDTQL = ThongSoCuoiKy.MaDTQL, ThongSo.MaDTSD = ThongSoCuoiKy.MaDTSD, ThongSo.MaTTSD = ThongSoCuoiKy.MaTTSD WHERE thongso.taisanIdCu > 0 and  (((ThongSo.Thang)<12 And (ThongSo.Thang)>0));"
 kiem_tra_nhap_xuat_ton_va_tai_khoan
 UnlockDB
 MsgBox "B¹n chuyÓn d÷ liÖu thµnh c«ng."
 End If
End Sub

Private Sub Command4_Click()
LockDB
    DBKetoan.Execute "delete *  from thongso where MaTS in (select MaTS from CTTaisan where MaCTKT = 0) "
    DBKetoan.Execute "delete *  from taisan where maso in (select mats from CTTaisan where MaCTKT = 0) "
    DBKetoan.Execute "delete *  from CTTaisan where MaCTKT = 0"
    kiem_tra_nhap_xuat_ton_va_tai_khoan
  UnlockDB
End Sub

Private Sub Command5_Click()
Dim rs_chungtu As Recordset
  Dim sql
Dim psw As String
  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
If Len(psw) > 0 Then
'DBKetoan.Execute " insert into HethongTK(MaSo,SoHieu,Cap,Ten,Kieu,Loai,TKCon,MaNT) select MaSo,SoHieu,Cap,Ten,Kieu,Loai,TKCon,MaNT from [" + psw + ";PWD=" + pPSW + "].HethongTK  where sohieu not in (select maso from HethongTK)"
'     Dua tai san co dinh qua
    LockDB
         sql = "UPDATE HethongTK tk1 inner join  [" + psw + ";PWD=" + pPSW + "].HethongTK tk2 on tk1.sohieu = tk2.sohieu SET tk1.DuNo_0 = tk2.DuNo_12, tk1.DuCo_0 = tk2.DuCo_12,tk1.DuNT_0 = tk2.DuNT_12"
           KiemTraTaiKhoan
         DBKetoan.Execute sql
         kiem_tra_nhap_xuat_ton_va_tai_khoan
    UnlockDB
    End If
End Sub
Sub kiem_tra_nhap_xuat_ton_va_tai_khoan()
            KiemTraVatTu
            KiemTraTaiKhoan
End Sub

'Public Sub chuyen_so_du_dau_ky()
'Dim psw As String, fn As String
'  Dim rs_chungtu As Recordset
'  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
'  Dim sql, st
'  Dim i As Integer
'  If Len(psw) > 0 Then
'                    LockDB 'them san pham
'                                  ExecuteSQL5_Themmoi ("ALTER TABLE phanloaivattu  ADD MaPhanLoaiCu Number")
'                                  DBKetoan.Execute "update phanloaivattu set MaPhanLoaiCu = 0"
'
'                                  ExecuteSQL5_Themmoi "ALTER TABLE  [" + psw + ";PWD=" + pPSW + "].phanloaivattu DROP COLUMN MaPhanLoaiCu"
'
'                                   sql = "insert into phanloaivattu select * ,maso as MaPhanLoaiCu from [" + psw + ";PWD=" + pPSW + "].phanloaivattu where maso  not in (select maso from phanloaivattu) "
'                                   DBKetoan.Execute sql
'
'                                    ExecuteSQL5_Themmoi ("ALTER TABLE dvtvattu  ADD MadvtvattuCu Number")
'                                    DBKetoan.Execute "update dvtvattu set MadvtvattuCu = 0"
'                                    ExecuteSQL5_Themmoi "ALTER TABLE  [" + psw + ";PWD=" + pPSW + "].dvtvattu DROP COLUMN MadvtvattuCu"
'
'
'                                    sql = "insert into dvtvattu select * ,maso as MadvtvattuCu from [" + psw + ";PWD=" + pPSW + "].dvtvattu " 'where sohieu  not in (select sohieu from dvtvattu) "
'                                   DBKetoan.Execute sql
'                                   Them_san_pham_chua_co psw, pPSW
'                                   'ket thuc them san pham
'      'chuyen cong no dau ky
'        ExecuteSQL5_Themmoi ("ALTER TABLE khachhang  ADD khachhangidCu Number")
'        DBKetoan.Execute "update khachhang set khachhangidCu = 0"
'        ExecuteSQL5_Themmoi "ALTER TABLE  [" + psw + ";PWD=" + pPSW + "].khachhang DROP COLUMN khachhangidCu"
'
'        sql = "insert into khachhang(MaPhanLoai , sohieu,  Ten, DiaChi, mst, Tel, Fax, email, taikhoan, DaiDien, DuMax, GhiChu, MaNT,khachhangidCu) select MaPhanLoai , sohieu,  Ten, DiaChi, mst, Tel, Fax, email, taikhoan, DaiDien, DuMax, GhiChu, MaNT, Maso as khachhangidCu from [" + psw + ";PWD=" + pPSW + "].khachhang "
'        sql = sql + " where maso not in (select maso from khachhang) "
'
'        DBKetoan.Execute sql
'        sql = " insert into sodukhachhang(MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0) select  sodu.MaTaiKhoan,khachhang.Maso as MaKhachHang,sodu.DuNo_0 as DuNo_12,sodu.DuCo_0 as DuCo_12,sodu.DuNT_0 as DuNT_12  from [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang sodu  inner join khachhang on sodu.MaKhachHang = khachhang.khachhangidCu"
'        DBKetoan.Execute sql
'
'      ' neu la sua lai so du dau ky
'        sql = " UPDATE SoDuKhachHang INNER JOIN [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang a "
'        sql = sql + " on SoDuKhachHang.maso = a.maso "
'        sql = sql + " set SoDuKhachHang.DuNo_0  = a.DuNo_12 "
'        sql = sql + ",SoDuKhachHang.DuCo_0 = a.DuCo_12 "
'        sql = sql + ",SoDuKhachHang.DuNT_0 = a.DuNT_12 "
'        DBKetoan.Execute sql
'      ' ket thuc cong no dau ky
'
'      'cap nhat lai he thong tai khoan
'         sql = "UPDATE HethongTK tk1 inner join  [" + psw + ";PWD=" + pPSW + "].HethongTK tk2 on tk1.sohieu = tk2.sohieu SET tk1.DuNo_0 = tk2.DuNo_12, tk1.DuCo_0 = tk2.DuCo_12,tk1.DuNT_0 = tk2.DuNT_12"
'         DBKetoan.Execute sql
'      'ket thuc cap nhat he thong tai khoan
'
'        'cap nhat tscd
'
'        ' xoa tai khoan co dinh
'        DBKetoan.Execute "delete *  from thongso where MaTS in (select MaTS from CTTaisan where MaCTKT = 0) "
'        DBKetoan.Execute "delete *  from taisan where maso in (select mats from CTTaisan where MaCTKT = 0) "
'        DBKetoan.Execute "delete *  from CTTaisan where MaCTKT = 0"
'        'ket thuc xoa tscd
'            DBKetoan.Execute "insert into LoaiTaiSan select * from [" + psw + ";PWD=" + pPSW + "].LoaiTaiSan"
'
'            ExecuteSQL5_Themmoi ("ALTER TABLE Taisan  ADD taisanIdCu Number")
'                    DBKetoan.Execute "update Taisan set taisanIdCu = 0"
'             ExecuteSQL5 " insert into Taisan(sohieu , Ten, NangLuc, GhiChu, MaNuoc, MaTaiKhoan, maloai, MaNhom, ThangTang, ThangGiam, NamSX, NamSD, NamKH, shct, NCT, taisanIdCu) select  sohieu , Ten, NangLuc, GhiChu, MaNuoc, MaTaiKhoan, maloai, MaNhom, ThangTang, ThangGiam, NamSX, NamSD, NamKH, shct, NCT, maso as taisanIdCu  from [" + psw + ";PWD=" + pPSW + "].Taisan "
'
'             ExecuteSQL5_Themmoi ("ALTER TABLE ThongSo  ADD taisanIdCu Number")
'                    DBKetoan.Execute "update ThongSo set taisanIdCu = 0"
'
'            ExecuteSQL5 "insert into  ThongSo(MaTS , thang, KH_NS, KH_TBS, KH_CNK, KH_TD, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD, MaDTQL, MaDTSD, MaTTSD,taisanIdCu) select taisan.maso as MaTS,Thang,KH_NS,KH_TBS,KH_CNK,KH_TD,NG_NS,NG_TBS,NG_CNK,NG_TD,CL_NS,CL_TBS,CL_CNK,CL_TD,MaDTQL,MaDTSD,MaTTSD ,taisan.maso as taisanIdCu from [" + psw + ";PWD=" + pPSW + "].thongso ts inner join taisan on ts.mats = taisan.taisanIdCu"
'
'            ExecuteSQL5_Themmoi ("ALTER TABLE CTTaiSan  ADD taisanIdCu Number")
'                    DBKetoan.Execute "update CTTaiSan set taisanIdCu = 0"
'
'                  sql = "insert into CTTaiSan (MaSo, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD,taisanIdCu) select  ctts.MaSo, ctts.SoHieu, ctts.Thang, ctts.VaoSo, ctts.NgayGhi, ctts.DienGiai, "
'                  sql = sql + "          ctts.MaLoai, ctts.MaNhom, taisan.maso as MaTS, ctts.NG_NS, ctts.NG_TBS, ctts.NG_CNK, ctts.NG_TD, ctts.CL_NS, ctts.CL_TBS, ctts.CL_CNK, ctts.CL_TD,ctts.mats as taisanIdCu "
'                   sql = sql + "    From [" + psw + ";PWD=" + pPSW + "].CTTaiSan ctts inner join Taisan on ctts.MATS = taisan.taisanIdCu"
'      DBKetoan.Execute sql
'      ExecuteSQL5 "UPDATE DISTINCTROW ThongSo LEFT JOIN ThongSoCuoiKy ON ThongSo.MaTS = ThongSoCuoiKy.MaTS SET ThongSo.KH_NS = ThongSoCuoiKy.KH_NS, ThongSo.KH_TBS = ThongSoCuoiKy.KH_TBS, ThongSo.KH_CNK = ThongSoCuoiKy.KH_CNK, ThongSo.KH_TD = ThongSoCuoiKy.KH_TD, ThongSo.MaDTQL = ThongSoCuoiKy.MaDTQL, ThongSo.MaDTSD = ThongSoCuoiKy.MaDTSD, ThongSo.MaTTSD = ThongSoCuoiKy.MaTTSD WHERE thongso.taisanIdCu > 0 and  (((ThongSo.Thang)<12 And (ThongSo.Thang)>0));"
'
'        'ket thuc cap nhat tai san co dinh
'
'        'chuyen tai khoan 154
'         sql = "insert into PhanLoai154  select *  from [" + psw + ";PWD=" + pPSW + "].PhanLoai154 where maso  not in (select maso from PhanLoai154) "
'         DBKetoan.Execute sql
'         sql = "insert into TP154  select *  from [" + psw + ";PWD=" + pPSW + "].TP154 where maso  not in (select maso from TP154) "
'         DBKetoan.Execute sql
'        sql = "update  TP154 a inner join  [" + psw + ";PWD=" + pPSW + "].TP154  b on a.maso = b.maso set "
'        sql = sql + "  a.MaPhanLoai = b.MaPhanLoai ,"
'        sql = sql + " a.SoHieu = b.SoHieu,"
'        sql = sql + " a.TenVattu = b.TenVattu,"
'        sql = sql + " a.DonVi = b.DonVi,"
'        sql = sql + "  a.GhiChu = b.GhiChu,"
'        sql = sql + "  a.DK = b.DK,"
'        sql = sql + "  a.DK1 = b.DK1,"
'        sql = sql + "  a.CK1 =b.CK1,"
'        sql = sql + "  a.CPNVL = b.CPNVL,"
'        sql = sql + "  a.CPNC = b.CPNC,"
'        sql = sql + "  a.CPM = b.CPM,"
'        sql = sql + "  a.DT = b.DT,"
'        sql = sql + "  a.CPTC = b.CPTC,"
'        sql = sql + "  a.Matk = b.MAtk,"
'        sql = sql + "   a.KPB = b.KPB,"
'        sql = sql + "   a.SanLuong = b.SanLuong"
'
'         sql = sql + cong_so_cot("CPKH", 1)
'         sql = sql + cong_so_cot("CPSXC", 1)
'
'         sql = sql + cong_so_cot("CPBH", 0)
'         sql = sql + cong_so_cot("CPQL", 0)
'         sql = sql + cong_so_cot("CPBHTT", 0)
'         sql = sql + cong_so_cot("CPQLTT", 0)
'         sql = sql + cong_so_cot("CPSXCTT", 0)
'         sql = sql + cong_so_cot("CPNVLPB", 0)
'         sql = sql + cong_so_cot("CPNCPB", 0)
'
'         DBKetoan.Execute sql
'          sql = "update  TP154 a inner join  [" + psw + ";PWD=" + pPSW + "].TP154  b on a.maso = b.maso set "
'          sql = sql + "  a.CPTC = b.CPTC "
'          sql = sql + cong_so_cot("CPMPB", 0)
'          sql = sql + cong_so_cot("CPTC", 0)
'          sql = sql + cong_so_cot("CPTCTT", 0)
'
'           DBKetoan.Execute sql
'        'ket thuc chuyen 154
'
'
'       UnlockDB
'         KiemTraTaiKhoan
'         kiem_tra_nhap_xuat_ton_va_tai_khoan
'
'  End If
'End Sub
'Public Sub chuyen_so_du_dau_ky()
'Dim psw As String, fn As String
'  Dim rs_chungtu As Recordset
'  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
'  Dim sql, st
'  Dim i As Integer
'  If Len(psw) > 0 Then
'  If (SelectSQL("select count(*)  as F1 from License where masothue = (select masothue from [" + psw + ";PWD=" + pPSW + "].License )") > 0) Then
'                      LockDB 'them san pham
'                         '         ExecuteSQL5_Themmoi ("ALTER TABLE phanloaivattu  ADD MaPhanLoaiCu Number")
'                             '     DBKetoan.Execute "update phanloaivattu set MaPhanLoaiCu = 0"
'
'                               '   ExecuteSQL5_Themmoi "ALTER TABLE  [" + psw + ";PWD=" + pPSW + "].phanloaivattu DROP COLUMN MaPhanLoaiCu"
'
'                                   sql = "insert into phanloaivattu select  *  from [" + psw + ";PWD=" + pPSW + "].phanloaivattu where maso  not in (select maso from phanloaivattu) "
'                                   DBKetoan.Execute sql
'
'                                '    ExecuteSQL5_Themmoi ("ALTER TABLE dvtvattu  ADD MadvtvattuCu Number")
'                                 '   DBKetoan.Execute "update dvtvattu set MadvtvattuCu = 0"
'                               '     ExecuteSQL5_Themmoi "ALTER TABLE  [" + psw + ";PWD=" + pPSW + "].dvtvattu DROP COLUMN MadvtvattuCu"
'
'
'                                    sql = "insert into dvtvattu select *  from [" + psw + ";PWD=" + pPSW + "].dvtvattu where maso  not in (select maso from dvtvattu) "
'                                   DBKetoan.Execute sql
'                                  ' Them_san_pham_chua_co psw, pPSW
'
'                                    sql = "insert into vattu select *  from [" + psw + ";PWD=" + pPSW + "].vattu where maso  not in (select maso from vattu) "
'                                    DBKetoan.Execute sql
'
'                                  sql = "insert into [Tonkho] select  * from  [" + psw + ";PWD=" + pPSW + "].tonkho  where maso not in (select maso from tonkho)"
'                                 DBKetoan.Execute sql
'
'                                       sql = " Update [Tonkho] b INNER JOIN [" + psw + ";PWD=" + pPSW + "].TonKho a ON b.mavattu = a.mavattu and b.MaSoKho = a.MaSoKho and   b.MaTaiKhoan = a.MaTaiKhoan"
'                                      sql = sql + " SET b.tien_0 = a.tien_12"
'                                      sql = sql + " ,b.Luong_0 = a.Luong_12"
''                                      For i = 1 To 12
''                                      st = CStr(i)
''                                      sql = sql + ",b.Luong_" + st + " = a.Luong_12, b.Tien_" + st + " = a.Tien_12"
''                                    Next
'                                DBKetoan.Execute sql
'                            '       sql = " Update [Tonkho] b INNER JOIN [" + psw + ";PWD=" + pPSW + "].TonKho a ON b.sanphamIdCu = a.mavattu and b.MaSoKho = a.MaSoKho and   b.MaTaiKhoan = a.MaTaiKhoan"
''                                   sql = " Update [Tonkho] b INNER JOIN [" + psw + ";PWD=" + pPSW + "].TonKho a ON b.maso = a.maso  "
''                                   sql = sql + " SET b.tien_0 = a.tien_12"
''                                   sql = sql + " ,b.Luong_0 = a.Luong_12"
'                                   'ket thuc them san pham
'      'chuyen cong no dau ky
''        ExecuteSQL5_Themmoi ("ALTER TABLE khachhang  ADD khachhangidCu Number")
''        DBKetoan.Execute "update khachhang set khachhangidCu = 0"
''        ExecuteSQL5_Themmoi "ALTER TABLE  [" + psw + ";PWD=" + pPSW + "].khachhang DROP COLUMN khachhangidCu"
'
'       ' sql = "insert into khachhang(MaPhanLoai , sohieu,  Ten, DiaChi, mst, Tel, Fax, email, taikhoan, DaiDien, DuMax, GhiChu, MaNT,khachhangidCu) select MaPhanLoai , sohieu,  Ten, DiaChi, mst, Tel, Fax, email, taikhoan, DaiDien, DuMax, GhiChu, MaNT, Maso as khachhangidCu from [" + psw + ";PWD=" + pPSW + "].khachhang "
'        sql = "insert into khachhang select * from [" + psw + ";PWD=" + pPSW + "].khachhang where maso not in (select maso from khachhang) "
'
'        DBKetoan.Execute sql
'      '  sql = " insert into sodukhachhang(MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0) select  sodu.MaTaiKhoan,khachhang.Maso as MaKhachHang,sodu.DuNo_0 as DuNo_12,sodu.DuCo_0 as DuCo_12,sodu.DuNT_0 as DuNT_12  from [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang sodu  inner join khachhang on sodu.MaKhachHang = khachhang.khachhangidCu"
'        sql = " insert into sodukhachhang select *  from [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang where maso not in (select maso from sodukhachhang)"
'        DBKetoan.Execute sql
'
'      ' neu la sua lai so du dau ky
'        sql = " UPDATE SoDuKhachHang INNER JOIN [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang a "
'        sql = sql + " on SoDuKhachHang.maso = a.maso "
'        sql = sql + " set SoDuKhachHang.DuNo_0  = a.DuNo_12 "
'        sql = sql + ",SoDuKhachHang.DuCo_0 = a.DuCo_12 "
'        sql = sql + ",SoDuKhachHang.DuNT_0 = a.DuNT_12 "
'        DBKetoan.Execute sql
'      ' ket thuc cong no dau ky
'
'      'cap nhat lai he thong tai khoan
'     '    sql = "UPDATE HethongTK tk1 inner join  [" + psw + ";PWD=" + pPSW + "].HethongTK tk2 on tk1.sohieu = tk2.sohieu SET tk1.DuNo_0 = tk2.DuNo_12, tk1.DuCo_0 = tk2.DuCo_12,tk1.DuNT_0 = tk2.DuNT_12"
'      '   DBKetoan.Execute sql
'
'      DBKetoan.Execute "update HethongTK set duno_0 = 0,duco_0 = 0"
'
'         sql = "UPDATE HethongTK tk1 inner join  [" + psw + ";PWD=" + pPSW + "].HethongTK tk2 on tk1.sohieu = tk2.sohieu SET tk1.DuNo_0 = tk2.DuNo_12, tk1.DuCo_0 = tk2.DuCo_12,tk1.DuNT_0 = tk2.DuNT_12"
'         DBKetoan.Execute sql
'           Dim rs_ktra, rs_tinh As Recordset
'           Dim no
'           Dim co
'        Set rs_ktra = DBKetoan.OpenRecordset("select  *  from hethongtk where Cap>0 AND Len(SoHieu)>1 ORDER BY SoHieu", dbOpenSnapshot)
'        Do While Not rs_ktra.EOF
'        If Left(rs_ktra!sohieu, 3) = "144" Then
'        MsgBox "thong bao"
'        End If
'        Set rs_tinh = DBKetoan.OpenRecordset("select  *  from hethongtk where Cap>0 AND SoHieu  = '" + rs_ktra!sohieu + "' ORDER BY SoHieu", dbOpenSnapshot)
'        If (SelectSQL("select sum(duno_0) as f1 from hethongtk where Cap > 0 AND Len(SoHieu)> len('" + rs_ktra!sohieu + "')  and sohieu like '" + rs_ktra!sohieu + "?'") <> rs_tinh!DuNo_0) Or (SelectSQL("select sum(duco_0) as f1 from hethongtk where Cap > 0 AND Len(SoHieu)> len('" + rs_ktra!sohieu + "')  and sohieu like '" + rs_ktra!sohieu + "?'") <> rs_tinh!DuCo_0) Then
'                    If (SelectSQL("select count(*) as f1 from hethongtk where  sohieu like '" + rs_ktra!sohieu + "?'") > 0) Then
'                     no = SelectSQL("select sum(duno_0) as f1 from hethongtk where  sohieu like '" + rs_ktra!sohieu + "?'")
'                     co = SelectSQL("select sum(duco_0) as f1 from hethongtk where    sohieu like '" + rs_ktra!sohieu + "?'")
'                                If (no <> rs_tinh!DuNo_0) Then
'                                 DBKetoan.Execute "update HethongTK set duno_0 =" + str(Abs(rs_tinh!DuNo_0 - no)) + " where val(sohieu) = (select min(val(sohieu)) from HethongTK where val(sohieu) >0 and sohieu like '" + rs_ktra!sohieu + "?" + "')"
'                                End If
'                                 If (co <> rs_tinh!DuCo_0) Then
'                                 DBKetoan.Execute "update HethongTK set duco_0 =" + str(Abs(rs_tinh!DuCo_0 - co)) + " where val(sohieu) = (select min(val(sohieu)) from HethongTK where val(sohieu) > 0 and sohieu like '" + rs_ktra!sohieu + "?" + "')"
'                                End If
'                     End If
'         End If
'         rs_ktra.MoveNext
'        Loop
'
'      'ket thuc cap nhat he thong tai khoan
'
'        'cap nhat tscd
'
'        ' xoa tai khoan co dinh
'        DBKetoan.Execute "delete *  from thongso where MaTS in (select MaTS from CTTaisan where MaCTKT = 0) "
'        DBKetoan.Execute "delete *  from taisan where maso in (select mats from CTTaisan where MaCTKT = 0) "
'        DBKetoan.Execute "delete *  from CTTaisan where MaCTKT = 0"
'        'ket thuc xoa tscd
'            DBKetoan.Execute "insert into LoaiTaiSan select * from [" + psw + ";PWD=" + pPSW + "].LoaiTaiSan"
'
'            ExecuteSQL5_Themmoi ("ALTER TABLE Taisan  ADD taisanIdCu Number")
'                    DBKetoan.Execute "update Taisan set taisanIdCu = 0"
'             ExecuteSQL5 " insert into Taisan(sohieu , Ten, NangLuc, GhiChu, MaNuoc, MaTaiKhoan, maloai, MaNhom, ThangTang, ThangGiam, NamSX, NamSD, NamKH, shct, NCT, taisanIdCu) select  sohieu , Ten, NangLuc, GhiChu, MaNuoc, MaTaiKhoan, maloai, MaNhom, ThangTang, ThangGiam, NamSX, NamSD, NamKH, shct, NCT, maso as taisanIdCu  from [" + psw + ";PWD=" + pPSW + "].Taisan "
'
'             ExecuteSQL5_Themmoi ("ALTER TABLE ThongSo  ADD taisanIdCu Number")
'                    DBKetoan.Execute "update ThongSo set taisanIdCu = 0"
'
'            ExecuteSQL5 "insert into  ThongSo(MaTS , thang, KH_NS, KH_TBS, KH_CNK, KH_TD, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD, MaDTQL, MaDTSD, MaTTSD,taisanIdCu) select taisan.maso as MaTS,Thang,KH_NS,KH_TBS,KH_CNK,KH_TD,NG_NS,NG_TBS,NG_CNK,NG_TD,CL_NS,CL_TBS,CL_CNK,CL_TD,MaDTQL,MaDTSD,MaTTSD ,taisan.maso as taisanIdCu from [" + psw + ";PWD=" + pPSW + "].thongso ts inner join taisan on ts.mats = taisan.taisanIdCu"
'
'            ExecuteSQL5_Themmoi ("ALTER TABLE CTTaiSan  ADD taisanIdCu Number")
'                    DBKetoan.Execute "update CTTaiSan set taisanIdCu = 0"
'
'                  sql = "insert into CTTaiSan (MaSo, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD,taisanIdCu) select  ctts.MaSo, ctts.SoHieu, ctts.Thang, ctts.VaoSo, ctts.NgayGhi, ctts.DienGiai, "
'                  sql = sql + "          ctts.MaLoai, ctts.MaNhom, taisan.maso as MaTS, ctts.NG_NS, ctts.NG_TBS, ctts.NG_CNK, ctts.NG_TD, ctts.CL_NS, ctts.CL_TBS, ctts.CL_CNK, ctts.CL_TD,ctts.mats as taisanIdCu "
'                   sql = sql + "    From [" + psw + ";PWD=" + pPSW + "].CTTaiSan ctts inner join Taisan on ctts.MATS = taisan.taisanIdCu"
'      DBKetoan.Execute sql
'      ExecuteSQL5 "UPDATE DISTINCTROW ThongSo LEFT JOIN ThongSoCuoiKy ON ThongSo.MaTS = ThongSoCuoiKy.MaTS SET ThongSo.KH_NS = ThongSoCuoiKy.KH_NS, ThongSo.KH_TBS = ThongSoCuoiKy.KH_TBS, ThongSo.KH_CNK = ThongSoCuoiKy.KH_CNK, ThongSo.KH_TD = ThongSoCuoiKy.KH_TD, ThongSo.MaDTQL = ThongSoCuoiKy.MaDTQL, ThongSo.MaDTSD = ThongSoCuoiKy.MaDTSD, ThongSo.MaTTSD = ThongSoCuoiKy.MaTTSD WHERE thongso.taisanIdCu > 0 and  (((ThongSo.Thang)<12 And (ThongSo.Thang)>0));"
'
'        'ket thuc cap nhat tai san co dinh
'
'        'chuyen tai khoan 154
'         sql = "insert into PhanLoai154  select *  from [" + psw + ";PWD=" + pPSW + "].PhanLoai154 where maso  not in (select maso from PhanLoai154) "
'         DBKetoan.Execute sql
'         sql = "insert into TP154  select *  from [" + psw + ";PWD=" + pPSW + "].TP154 where maso  not in (select maso from TP154) "
'         DBKetoan.Execute sql
'        sql = "update  TP154 a inner join  [" + psw + ";PWD=" + pPSW + "].TP154  b on a.maso = b.maso set "
'        sql = sql + "  a.MaPhanLoai = b.MaPhanLoai ,"
'        sql = sql + " a.SoHieu = b.SoHieu,"
'        sql = sql + " a.TenVattu = b.TenVattu,"
'        sql = sql + " a.DonVi = b.DonVi,"
'        sql = sql + "  a.GhiChu = b.GhiChu,"
'        sql = sql + "  a.DK = b.DK,"
'        sql = sql + "  a.DK1 = b.DK1,"
'        sql = sql + "  a.CK1 =b.CK1,"
'        sql = sql + "  a.CPNVL = b.CPNVL,"
'        sql = sql + "  a.CPNC = b.CPNC,"
'        sql = sql + "  a.CPM = b.CPM,"
'        sql = sql + "  a.DT = b.DT,"
'        sql = sql + "  a.CPTC = b.CPTC,"
'        sql = sql + "  a.Matk = b.MAtk,"
'        sql = sql + "   a.KPB = b.KPB,"
'        sql = sql + "   a.SanLuong = b.SanLuong"
'
'         sql = sql + cong_so_cot("CPKH", 1)
'         sql = sql + cong_so_cot("CPSXC", 1)
'
'         sql = sql + cong_so_cot("CPBH", 0)
'         sql = sql + cong_so_cot("CPQL", 0)
'         sql = sql + cong_so_cot("CPBHTT", 0)
'         sql = sql + cong_so_cot("CPQLTT", 0)
'         sql = sql + cong_so_cot("CPSXCTT", 0)
'         sql = sql + cong_so_cot("CPNVLPB", 0)
'         sql = sql + cong_so_cot("CPNCPB", 0)
'
'         DBKetoan.Execute sql
'          sql = "update  TP154 a inner join  [" + psw + ";PWD=" + pPSW + "].TP154  b on a.maso = b.maso set "
'          sql = sql + "  a.CPTC = b.CPTC "
'          sql = sql + cong_so_cot("CPMPB", 0)
'          sql = sql + cong_so_cot("CPTC", 0)
'          sql = sql + cong_so_cot("CPTCTT", 0)
'
'           DBKetoan.Execute sql
'        'ket thuc chuyen 154
'
'
'       UnlockDB
'       KiemTraTaiKhoan
'       kiem_tra_nhap_xuat_ton_va_tai_khoan
'       Else
'       MsgBox "Khong cung tep du lieu"
'       End If
'
'
'  End If
'End Sub
Public Sub chuyen_so_du_dau_ky()
Dim psw As String, fn As String
  Dim rs_chungtu As Recordset
  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
  Dim sql, st
  Dim i As Integer
  If Len(psw) > 0 Then
  If (SelectSQL("select count(*)  as F1 from License where masothue = (select masothue from [" + psw + ";PWD=" + pPSW + "].License )") > 0) Then
                      LockDB 'them san pham
                      
       DBKetoan.Execute "update hethongtk set duno_0 = 0,duco_0 = 0 where sohieu like '156*'"
       DBKetoan.Execute "update hethongtk set duno_0 = 0,duco_0 = 0 where sohieu like '154*'"
       DBKetoan.Execute "update hethongtk set duno_0 = 0,duco_0 = 0 where sohieu like '152*'"

                                  ExecuteSQL5_Themmoi ("ALTER TABLE phanloaivattu  ADD MaPhanLoaiCu Number")
                             '     DBKetoan.Execute "update phanloaivattu set MaPhanLoaiCu = 0"
                                  
                               '   ExecuteSQL5_Themmoi "ALTER TABLE  [" + psw + ";PWD=" + pPSW + "].phanloaivattu DROP COLUMN MaPhanLoaiCu"
                                  
                                   sql = "insert into phanloaivattu select  *  from [" + psw + ";PWD=" + pPSW + "].phanloaivattu where maso  not in (select maso from phanloaivattu) "
                                   DBKetoan.Execute sql
                                   
                                   ExecuteSQL5_Themmoi ("ALTER TABLE dvtvattu  ADD MadvtvattuCu Number")
                                 '   DBKetoan.Execute "update dvtvattu set MadvtvattuCu = 0"
                               '     ExecuteSQL5_Themmoi "ALTER TABLE  [" + psw + ";PWD=" + pPSW + "].dvtvattu DROP COLUMN MadvtvattuCu"
                                    
                                    
                                    sql = "insert into dvtvattu select *  from [" + psw + ";PWD=" + pPSW + "].dvtvattu where maso  not in (select maso from dvtvattu) "
                                   DBKetoan.Execute sql
                                  ' Them_san_pham_chua_co psw, pPSW
                                  
                                    sql = "insert into vattu select *  from [" + psw + ";PWD=" + pPSW + "].vattu where maso  not in (select maso from vattu) "
                                    DBKetoan.Execute sql
                  
                                  sql = "insert into [Tonkho] select  * from  [" + psw + ";PWD=" + pPSW + "].tonkho  where maso not in (select maso from tonkho)"
                                 DBKetoan.Execute sql
                  
                                       sql = " Update [Tonkho] b INNER JOIN [" + psw + ";PWD=" + pPSW + "].TonKho a ON b.maso = a.maso  " 'b.mavattu = a.mavattu and b.MaSoKho = a.MaSoKho and   b.MaTaiKhoan = a.MaTaiKhoan"
                                      sql = sql + " SET b.tien_0 = a.tien_12"
                                      sql = sql + " ,b.Luong_0 = a.Luong_12"
'                                      For i = 1 To 12
'                                      st = CStr(i)
'                                      sql = sql + ",b.Luong_" + st + " = a.Luong_12, b.Tien_" + st + " = a.Tien_12"
'                                    Next
                                DBKetoan.Execute sql
                            '       sql = " Update [Tonkho] b INNER JOIN [" + psw + ";PWD=" + pPSW + "].TonKho a ON b.sanphamIdCu = a.mavattu and b.MaSoKho = a.MaSoKho and   b.MaTaiKhoan = a.MaTaiKhoan"
'                                   sql = " Update [Tonkho] b INNER JOIN [" + psw + ";PWD=" + pPSW + "].TonKho a ON b.maso = a.maso  "
'                                   sql = sql + " SET b.tien_0 = a.tien_12"
'                                   sql = sql + " ,b.Luong_0 = a.Luong_12"
                                   'ket thuc them san pham
      'chuyen cong no dau ky
'        ExecuteSQL5_Themmoi ("ALTER TABLE khachhang  ADD khachhangidCu Number")
'        DBKetoan.Execute "update khachhang set khachhangidCu = 0"
'        ExecuteSQL5_Themmoi "ALTER TABLE  [" + psw + ";PWD=" + pPSW + "].khachhang DROP COLUMN khachhangidCu"
         
       ' sql = "insert into khachhang(MaPhanLoai , sohieu,  Ten, DiaChi, mst, Tel, Fax, email, taikhoan, DaiDien, DuMax, GhiChu, MaNT,khachhangidCu) select MaPhanLoai , sohieu,  Ten, DiaChi, mst, Tel, Fax, email, taikhoan, DaiDien, DuMax, GhiChu, MaNT, Maso as khachhangidCu from [" + psw + ";PWD=" + pPSW + "].khachhang "
        sql = "insert into khachhang select * from [" + psw + ";PWD=" + pPSW + "].khachhang where maso not in (select maso from khachhang) "
      
        DBKetoan.Execute sql
      '  sql = " insert into sodukhachhang(MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0) select  sodu.MaTaiKhoan,khachhang.Maso as MaKhachHang,sodu.DuNo_0 as DuNo_12,sodu.DuCo_0 as DuCo_12,sodu.DuNT_0 as DuNT_12  from [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang sodu  inner join khachhang on sodu.MaKhachHang = khachhang.khachhangidCu"
        sql = " insert into sodukhachhang select *  from [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang where maso not in (select maso from sodukhachhang)"
        DBKetoan.Execute sql
        
      ' neu la sua lai so du dau ky
        sql = " UPDATE SoDuKhachHang INNER JOIN [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang a "
        sql = sql + " on SoDuKhachHang.maso = a.maso "
        sql = sql + " set SoDuKhachHang.DuNo_0  = a.DuNo_12 "
        sql = sql + ",SoDuKhachHang.DuCo_0 = a.DuCo_12 "
        sql = sql + ",SoDuKhachHang.DuNT_0 = a.DuNT_12 "
        DBKetoan.Execute sql
      ' ket thuc cong no dau ky
      
      'cap nhat lai he thong tai khoan
       DBKetoan.Execute "update HethongTK set duno_0 = 0,duco_0 = 0,DuNT_0=0"
      
         sql = "UPDATE HethongTK tk1 inner join  [" + psw + ";PWD=" + pPSW + "].HethongTK tk2 on tk1.sohieu = tk2.sohieu SET tk1.DuNo_0 = tk2.DuNo_12, tk1.DuCo_0 = tk2.DuCo_12,tk1.DuNT_0 = tk2.DuNT_12"
         DBKetoan.Execute sql
           Dim rs_ktra, rs_tinh As Recordset
           Dim no
           Dim co
'        Set rs_ktra = DBKetoan.OpenRecordset("select  *  from hethongtk where Cap>0 AND Len(SoHieu)>1 ORDER BY SoHieu", dbOpenSnapshot)
'        Do While Not rs_ktra.EOF
'        If Left(rs_ktra!sohieu, 3) = "112" Then
'        MsgBox "thong bao"
'        End If
'        Set rs_tinh = DBKetoan.OpenRecordset("select  *  from hethongtk where Cap>0 AND SoHieu  = '" + rs_ktra!sohieu + "' ORDER BY SoHieu", dbOpenSnapshot)
'        If (SelectSQL("select sum(duno_0) as f1 from hethongtk where Cap > 0 AND Len(SoHieu)> len('" + rs_ktra!sohieu + "')  and sohieu like '" + rs_ktra!sohieu + "?'") <> rs_tinh!DuNo_0) Or (SelectSQL("select sum(duco_0) as f1 from hethongtk where Cap > 0 AND Len(SoHieu)> len('" + rs_ktra!sohieu + "')  and sohieu like '" + rs_ktra!sohieu + "?'") <> rs_tinh!DuCo_0) Then
'                    If (SelectSQL("select count(*) as f1 from hethongtk where  sohieu like '" + rs_ktra!sohieu + "?'") > 0) Then
'                     no = SelectSQL("select sum(duno_0) as f1 from hethongtk where  sohieu like '" + rs_ktra!sohieu + "?'")
'                     co = SelectSQL("select sum(duco_0) as f1 from hethongtk where    sohieu like '" + rs_ktra!sohieu + "?'")
'                                If (no <> rs_tinh!DuNo_0) Then
'                                 DBKetoan.Execute "update HethongTK set duno_0 =" + str(Abs(rs_tinh!DuNo_0 - no)) + " where maso not in (select maso from [" + psw + ";PWD=" + pPSW + "].Hethongtk) and  val(sohieu) = (select min(val(sohieu)) from HethongTK where val(sohieu) >0 and sohieu like '" + rs_ktra!sohieu + "?" + "')"
'                                End If
'                                 If (co <> rs_tinh!DuCo_0) Then
'                                 DBKetoan.Execute "update HethongTK set duco_0 =" + str(Abs(rs_tinh!DuCo_0 - co)) + " where maso not in (select maso from [" + psw + ";PWD=" + pPSW + "].Hethongtk) and  val(sohieu) = (select min(val(sohieu)) from HethongTK where val(sohieu) > 0 and sohieu like '" + rs_ktra!sohieu + "?" + "')"
'                                End If
'                     End If
'         End If
'         rs_ktra.MoveNext
'        Loop
     
         
      'ket thuc cap nhat he thong tai khoan
      
        'cap nhat tscd
        
        ' xoa tai khoan co dinh
        DBKetoan.Execute "delete *  from thongso where MaTS in (select MaTS from CTTaisan where MaCTKT = 0) "
        DBKetoan.Execute "delete *  from taisan where maso in (select mats from CTTaisan where MaCTKT = 0) "
        DBKetoan.Execute "delete *  from CTTaisan where MaCTKT = 0"
        'ket thuc xoa tscd
            DBKetoan.Execute "insert into LoaiTaiSan select * from [" + psw + ";PWD=" + pPSW + "].LoaiTaiSan"
             
            ExecuteSQL5_Themmoi ("ALTER TABLE Taisan  ADD taisanIdCu Number")
                    DBKetoan.Execute "update Taisan set taisanIdCu = 0"
             ExecuteSQL5 " insert into Taisan(sohieu , Ten, NangLuc, GhiChu, MaNuoc, MaTaiKhoan, maloai, MaNhom, ThangTang, ThangGiam, NamSX, NamSD, NamKH, shct, NCT, taisanIdCu) select  sohieu , Ten, NangLuc, GhiChu, MaNuoc, MaTaiKhoan, maloai, MaNhom, ThangTang, ThangGiam, NamSX, NamSD, NamKH, shct, NCT, maso as taisanIdCu  from [" + psw + ";PWD=" + pPSW + "].Taisan "
             
             ExecuteSQL5_Themmoi ("ALTER TABLE ThongSo  ADD taisanIdCu Number")
                    DBKetoan.Execute "update ThongSo set taisanIdCu = 0"
                    
            ExecuteSQL5 "insert into  ThongSo(MaTS , thang, KH_NS, KH_TBS, KH_CNK, KH_TD, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD, MaDTQL, MaDTSD, MaTTSD,taisanIdCu) select taisan.maso as MaTS,Thang,KH_NS,KH_TBS,KH_CNK,KH_TD,NG_NS,NG_TBS,NG_CNK,NG_TD,CL_NS,CL_TBS,CL_CNK,CL_TD,MaDTQL,MaDTSD,MaTTSD ,taisan.maso as taisanIdCu from [" + psw + ";PWD=" + pPSW + "].thongso ts inner join taisan on ts.mats = taisan.taisanIdCu"
                  
            ExecuteSQL5_Themmoi ("ALTER TABLE CTTaiSan  ADD taisanIdCu Number")
                    DBKetoan.Execute "update CTTaiSan set taisanIdCu = 0"
                  
                  sql = "insert into CTTaiSan (MaSo, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD,taisanIdCu) select  ctts.MaSo, ctts.SoHieu, ctts.Thang, ctts.VaoSo, ctts.NgayGhi, ctts.DienGiai, "
                  sql = sql + "          ctts.MaLoai, ctts.MaNhom, taisan.maso as MaTS, ctts.NG_NS, ctts.NG_TBS, ctts.NG_CNK, ctts.NG_TD, ctts.CL_NS, ctts.CL_TBS, ctts.CL_CNK, ctts.CL_TD,ctts.mats as taisanIdCu "
                   sql = sql + "    From [" + psw + ";PWD=" + pPSW + "].CTTaiSan ctts inner join Taisan on ctts.MATS = taisan.taisanIdCu"
      DBKetoan.Execute sql
      ExecuteSQL5 "UPDATE DISTINCTROW ThongSo LEFT JOIN ThongSoCuoiKy ON ThongSo.MaTS = ThongSoCuoiKy.MaTS SET ThongSo.KH_NS = ThongSoCuoiKy.KH_NS, ThongSo.KH_TBS = ThongSoCuoiKy.KH_TBS, ThongSo.KH_CNK = ThongSoCuoiKy.KH_CNK, ThongSo.KH_TD = ThongSoCuoiKy.KH_TD, ThongSo.MaDTQL = ThongSoCuoiKy.MaDTQL, ThongSo.MaDTSD = ThongSoCuoiKy.MaDTSD, ThongSo.MaTTSD = ThongSoCuoiKy.MaTTSD WHERE thongso.taisanIdCu > 0 and  (((ThongSo.Thang)<12 And (ThongSo.Thang)>0));"
 
        'ket thuc cap nhat tai san co dinh
        
        'chuyen tai khoan 154
         sql = "insert into PhanLoai154  select *  from [" + psw + ";PWD=" + pPSW + "].PhanLoai154 where maso  not in (select maso from PhanLoai154) "
         DBKetoan.Execute sql
         sql = "insert into TP154  select *  from [" + psw + ";PWD=" + pPSW + "].TP154 where maso  not in (select maso from TP154) "
         DBKetoan.Execute sql
        sql = "update  TP154 a inner join  [" + psw + ";PWD=" + pPSW + "].TP154  b on a.maso = b.maso set "
        sql = sql + "  a.MaPhanLoai = b.MaPhanLoai ,"
        sql = sql + " a.SoHieu = b.SoHieu,"
        sql = sql + " a.TenVattu = b.TenVattu,"
        sql = sql + " a.DonVi = b.DonVi,"
        sql = sql + "  a.GhiChu = b.GhiChu,"
        sql = sql + "  a.DK = b.DK,"
        sql = sql + "  a.DK1 = b.DK1,"
        sql = sql + "  a.CK1 =b.CK1,"
        sql = sql + "  a.CPNVL = b.CPNVL,"
        sql = sql + "  a.CPNC = b.CPNC,"
        sql = sql + "  a.CPM = b.CPM,"
        sql = sql + "  a.DT = b.DT,"
        sql = sql + "  a.CPTC = b.CPTC,"
        sql = sql + "  a.Matk = b.MAtk,"
        sql = sql + "   a.KPB = b.KPB,"
        sql = sql + "   a.SanLuong = b.SanLuong"

         sql = sql + cong_so_cot("CPKH", 1)
         sql = sql + cong_so_cot("CPSXC", 1)

         sql = sql + cong_so_cot("CPBH", 0)
         sql = sql + cong_so_cot("CPQL", 0)
         sql = sql + cong_so_cot("CPBHTT", 0)
         sql = sql + cong_so_cot("CPQLTT", 0)
         sql = sql + cong_so_cot("CPSXCTT", 0)
         sql = sql + cong_so_cot("CPNVLPB", 0)
         sql = sql + cong_so_cot("CPNCPB", 0)
         
         DBKetoan.Execute sql
          sql = "update  TP154 a inner join  [" + psw + ";PWD=" + pPSW + "].TP154  b on a.maso = b.maso set "
          sql = sql + "  a.CPTC = b.CPTC "
          sql = sql + cong_so_cot("CPMPB", 0)
          sql = sql + cong_so_cot("CPTC", 0)
          sql = sql + cong_so_cot("CPTCTT", 0)

           DBKetoan.Execute sql
        'ket thuc chuyen 154
        
        
       UnlockDB
      KiemTraTaiKhoan
      kiem_tra_nhap_xuat_ton_va_tai_khoan
       Else
       MsgBox "Khong cung tep du lieu"
       End If
         
                         
  End If
End Sub
Public Function cong_so_cot(st As String, k As Integer) As String
Dim chuoi
chuoi = ""
 For i = 1 To 12
 If k = 1 Then
  chuoi = chuoi + " ,a." + st + "_" + CStr(i) + "=  b." + st + "_" + CStr(i)
 Else
       chuoi = chuoi + " ,a." + st + CStr(i) + "=  b." + st + CStr(i)
 End If
 Next
 cong_so_cot = chuoi
End Function
'
'Sub tinh_them()
'        Set rs_tinh = DBKetoan.OpenRecordset("select  *  from hethongtk where Cap>0 AND SoHieu  = '" + rs_ktra!sohieu + "' ORDER BY SoHieu", dbOpenSnapshot)
'        If (SelectSQL("select sum(duno_0) as f1 from hethongtk where Cap > 0 AND Len(SoHieu)> len('" + rs_ktra!sohieu + "')  and sohieu like '" + rs_ktra!sohieu + "?'") <> rs_tinh!DuNo_0) Or (SelectSQL("select sum(duco_0) as f1 from hethongtk where Cap > 0 AND Len(SoHieu)> len('" + rs_ktra!sohieu + "')  and sohieu like '" + rs_ktra!sohieu + "?'") <> rs_tinh!DuCo_0) Then
'                    If (SelectSQL("select count(*) as f1 from hethongtk where  sohieu like '" + rs_ktra!sohieu + "?'") > 0) Then
'                     no = SelectSQL("select sum(duno_0) as f1 from hethongtk where  Len(SoHieu)> len('" + rs_ktra!sohieu + "') and sohieu like '" + rs_ktra!sohieu + "*' and ")
'                     co = SelectSQL("select sum(duco_0) as f1 from hethongtk where    Len(SoHieu)> len('" + rs_ktra!sohieu + "') and sohieu like '" + rs_ktra!sohieu + "*")
'                                If (no <> rs_tinh!DuNo_0) Then
'                                     DBKetoan.Execute "update HethongTK set duno_0 =" + str(Abs(rs_tinh!DuNo_0 - no)) + " where val(sohieu) = (select min(val(sohieu)) from HethongTK where val(sohieu) >0 and sohieu like '" + rs_ktra!sohieu + "?" + "')"
'                                End If
'                                 If (co <> rs_tinh!DuCo_0) Then
'                                     DBKetoan.Execute "update HethongTK set duco_0 =" + str(Abs(rs_tinh!DuCo_0 - co)) + " where val(sohieu) = (select min(val(sohieu)) from HethongTK where val(sohieu) > 0 and sohieu like '" + rs_ktra!sohieu + "?" + "')"
'                                End If
'                     End If
'         End If
'         rs_ktra.MoveNext
'End Sub

