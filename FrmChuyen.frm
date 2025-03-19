VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form FrmChuyen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ChuyÓn sè d­ n¨m tr­íc sang..."
   ClientHeight    =   630
   ClientLeft      =   4875
   ClientTop       =   5445
   ClientWidth     =   3975
   ClipControls    =   0   'False
   Icon            =   "FrmChuyen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command 
      Caption         =   "Sè liÖu 1998"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      CancelError     =   -1  'True
      DefaultExt      =   "*.MDB"
      DialogTitle     =   "Chän tÖp d÷ liÖu"
      FileName        =   "*.MDB"
      Filter          =   "TÖp d÷ liÖu (*.MDB)|*.MDB|TÊt c¶ (*.*)|*.*"
   End
End
Attribute VB_Name = "FrmChuyen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_Click()
    dlgCommonDialog.Flags = &H4&
    dlgCommonDialog.filename = "*.MDB"
    On Error GoTo QuitSAS
    dlgCommonDialog.ShowOpen
    p1998 = dlgCommonDialog.filename
    On Error GoTo 0
    If p1998 <> "" Then
        Set DB1998 = Wspace.OpenDatabase(p1998, False, False)
        ChuyenSoLieu
    End If
QuitSAS:
    Wspace.Close
    End
End Sub

Private Sub Form_Load()
    pCurDir = App.Path
    If Right(pCurDir, 1) <> "\" Then pCurDir = pCurDir + "\"
    
    IniPath = Dir(App.Path, vbDirectory)
    
    pDataPath = GetSetting(IniPath, "Environment", "Path", "")
    If pDataPath = "" Then End
    Set Wspace = DBEngine.CreateWorkspace(CStr(Time), "Admin", "", dbUseJet)
    Workspaces.Append Wspace
            
    Do While OpenDB(pDataPath) <> 0
        dlgCommonDialog.Flags = &H4&
        dlgCommonDialog.filename = "*.MDB"
        On Error GoTo QuitSAS
        dlgCommonDialog.ShowOpen
        On Error GoTo 0
        pDataPath = dlgCommonDialog.filename
    Loop
    If Lng_MaxValue("MaSo", "HethongTK") > 5150 Then GoTo QuitSAS
    Exit Sub
QuitSAS:
    Wspace.Close
    End
End Sub

Private Sub ChuyenSoLieu()
    ChuyenTK
    ChuyenVT
    ChuyenTS
    ChuyenCN
End Sub

Private Sub ChuyenTK()
Dim rs As Recordset
Dim rs2 As Recordset
Dim tk As New ClsTaikhoan

    Set rs = DB1998.OpenRecordset("SELECT DISTINCTROW HethongTK.* FROM HethongTK WHERE Cap > 0 AND MaNT <=0 ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        Set rs2 = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo, TkCha0 FROM HethongTK WHERE SoHieu = '" + rs!SoHieu + "'", dbOpenSnapshot)
        If rs2.RecordCount = 0 Then
            Set rs2 = DB1998.OpenRecordset("SELECT DISTINCTROW HethongTK.SoHieu FROM HethongTK WHERE MaSo = " + CStr(rs!TkCha0), dbOpenSnapshot)
            tk.InitTaikhoanSohieu rs2!SoHieu
            If tk.MaSo > 0 Then
                tk.SoHieu = rs!SoHieu
                tk.Ten = rs!Ten
                tk.NoDauKy = rs!DuNo_12
                tk.CoDauKy = rs!DuCo_12
                tk.cap = rs!cap
                tk.TkCon = 0
                tk.TkCha5 = tk.TkCha4
                tk.TkCha4 = tk.TkCha3
                tk.TkCha3 = tk.TkCha2
                tk.TkCha2 = tk.TkCha1
                tk.TkCha1 = tk.TkCha0
                tk.TkCha0 = tk.MaSo
                tk.ThemTk
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
'======================================================================================
' Sub OpenDB
'======================================================================================
Private Function OpenDB(file_name As String) As Integer
    Dim retry As Integer
    
    OpenDB = -1
    If Dir(file_name) = "" Then GoTo KhongMo
    retry = 0
On Error GoTo DB_Handle
      ' Mo tep du lieu duoc chi dinh23
      Set DBKetoan = Wspace.OpenDatabase(file_name, False, False)
      
On Error GoTo 0
      pDataPath = file_name
      OpenDB = 0
      Exit Function
' Khong mo duoc du lieu
DB_Handle:
    If retry = 5 Then
        DBEngine.RepairDatabase file_name
    End If
    If retry > 10 Then
KhongMo:
      MsgBox "Kh«ng më ®­îc tÖp d÷ liÖu", vbCritical, "SAS"
    Else
      retry = retry + 1
      DoEvents
      Resume
    End If
End Function

Private Sub ChuyenVT()
Dim sql1 As String, sql2 As String
Dim i As Integer

    ExecuteSQL5 "DELETE FROM TonKho"
    ExecuteSQL5 "DELETE FROM VatTu"
    ExecuteSQL5 "DELETE FROM PhanLoaiVatTu"
    ExecuteSQL5 "DELETE FROM KhoHang"
    ExecuteSQL5 "DELETE FROM NguonNhapXuat"
    
    ExecuteSQL5 "INSERT INTO PhanLoaiVatTu (MaSo, MaTK, SoHieu, TenPhanLoai)" _
        & " SELECT DISTINCTROW MaSo, MaTK, SoHieu, TenPhanLoai FROM PhanLoaiVatTu IN '" + p1998 + "'"
    ExecuteSQL5 "INSERT INTO VatTu (MaSo, MaPhanLoai, SoHieu, TenVatTu, DonVi, TonMin, TonMax, GiaDuPhong)" _
        & " SELECT DISTINCTROW MaSo, MaPhanLoai, SoHieu, TenVatTu, DonVi, TonMin, TonMax, GiaDuPhong FROM VatTu IN '" + p1998 + "'"
    ExecuteSQL5 "INSERT INTO KhoHang (MaSo, TenKho)" _
        & " SELECT DISTINCTROW MaSo, TenKho FROM KhoHang IN '" + p1998 + "'"
    ExecuteSQL5 "INSERT INTO NguonNhapXuat (MaSo, MaLoai, SoHieu, DienGiai)" _
        & " SELECT DISTINCTROW MaSo, MaLoai, SoHieu, DienGiai FROM NguonNhapXuat IN '" + p1998 + "'"
        
    sql1 = "MaSo, MaSoKho, MaVattu"
    sql2 = "MaSo, MaSoKho, MaVattu"
    For i = 0 To 12
        sql1 = sql1 + ", Luong_" + CStr(i) + ", Tien_" + CStr(i)
        sql2 = sql2 + ", Luong_12, Tien_12"
    Next
    ExecuteSQL5 "INSERT INTO TonKho (" + sql1 + ") SELECT DISTINCTROW " + sql2 + " FROM TonKho IN '" + p1998 + "' WHERE Luong_12 <> 0 AND Tien_12 <> 0"
    SoDuTKVT
End Sub

Private Sub SoDuTKVT()
    Dim rs_tk As Recordset, taikhoan As New ClsTaikhoan
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW Sum(TonKho.Tien_" + CStr(pThangDauKy) + ") AS DuDk, PhanLoaiVattu.MaTK" _
        & " FROM PhanLoaiVattu RIGHT JOIN (Vattu RIGHT JOIN TonKho ON Vattu.MaSo = TonKho.MaVatTu) ON PhanLoaiVattu.MaSo = Vattu.MaPhanLoai" _
        & " GROUP BY PhanLoaiVattu.MaTK", dbOpenSnapshot, dbForwardOnly)
    
    Do While Not rs_tk.EOF
        taikhoan.InitTaikhoanMaSo rs_tk!MaTK
        taikhoan.NoDauKy = rs_tk!DuDk
        taikhoan.CapNhatTk
        rs_tk.MoveNext
    Loop
    rs_tk.Close
    Set rs_tk = Nothing
End Sub

Private Sub ChuyenTS()
    ExecuteSQL5 "DELETE FROM TinhTrang"
    ExecuteSQL5 "DELETE FROM ThongSo"
    ExecuteSQL5 "DELETE FROM TaiSan"
    ExecuteSQL5 "DELETE FROM CTTaiSan"
    ExecuteSQL5 "DELETE FROM DTQLy"
    ExecuteSQL5 "DELETE FROM LoaiChungTu"
    ExecuteSQL5 "DELETE FROM LoaiTaiSan"
    ExecuteSQL5 "DELETE FROM QuocGia"
    ExecuteSQL5 "DELETE FROM BaoCaoCP"
    
    ExecuteSQL5 "INSERT INTO TinhTrang (MaSo, Ten)" _
        & " SELECT DISTINCTROW MaSo, Ten FROM TinhTrang IN '" + p1998 + "'"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo, SoHieu)" _
        & " SELECT DISTINCTROW MaSo, TKCP FROM DTSDung IN '" + p1998 + "'"
    ExecuteSQL5 "INSERT INTO DTQLy (MaSo, Ten)" _
        & " SELECT DISTINCTROW MaSo, Ten FROM DTQLy IN '" + p1998 + "'"
    ExecuteSQL5 "INSERT INTO LoaiChungTu (MaSo, Ten, SoHieu, CapTren, Cap)" _
        & " SELECT DISTINCTROW MaSo, Ten, SoHieu, CapTren, Cap FROM LoaiChungTu IN '" + p1998 + "'"
    ExecuteSQL5 "INSERT INTO LoaiTaiSan (MaSo, Ten, SoHieu, CapTren, Cap)" _
        & " SELECT DISTINCTROW MaSo, Ten, SoHieu, CapTren, Cap FROM LoaiTaiSan IN '" + p1998 + "'"
    ExecuteSQL5 "INSERT INTO QuocGia (MaSo, Ten)" _
        & " SELECT DISTINCTROW MaSo, Ten FROM QuocGia IN '" + p1998 + "'"
    ExecuteSQL5 "INSERT INTO TaiSan (MaSo, SoHieu, Ten, NangLuc, GhiChu, MaNuoc, MaTaiKhoan, MaLoai, MaNhom, ThangTang, ThangGiam, NamSX, NamSD)" _
        & " SELECT DISTINCTROW MaSo, SoHieu, Ten, NangLuc, GhiChu, MaNuoc, MaTaiKhoan, MaLoai, MaNhom, 0 AS ThangTang, 13 AS ThangGiam, NamSX, NamSD FROM TaiSan IN '" + p1998 _
        + "' WHERE ThangGiam = 13"
    ExecuteSQL5 "INSERT INTO CTTaiSan (MaSo, MaCTKT, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD)" _
        & " SELECT DISTINCTROW MaSo, 0 AS MaCTKT, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD FROM CTTaiSan IN '" + p1998 + "'"
    ExecuteSQL5 "INSERT INTO ThongSo (MaSo, MaTS, Thang, KH_NS, KH_TBS, KH_CNK, KH_TD, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD, MaDTSD, MaDTQL, MaTTSD)" _
        & " SELECT DISTINCTROW ThongSo.MaSo, MaTS, Thang, KH_NS, KH_TBS, KH_CNK, KH_TD, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD, MaDTSD, MaDTQL, MaTTSD" _
        & " FROM ThongSo INNER JOIN TaiSan ON ThongSo.MaTS = TaiSan.MaSo IN '" _
        + p1998 + "' WHERE TaiSan.ThangGiam = 13"
    ExecuteSQL5 "UPDATE (ThongSo INNER JOIN BaoCaoCP ON ThongSo.MaDTSD = BaoCaoCP.MaSo) INNER JOIN HethongTK ON BaoCaoCP.SoHieu = HethongTK.SoHieu SET ThongSo.MaDTSD = HethongTK.MaSo"
    
    ExecuteSQL5 "UPDATE (ThongSo INNER JOIN BaoCaoCP ON ThongSo.MaDTSD = BaoCaoCP.MaSo) LEFT JOIN HethongTK ON BaoCaoCP.SoHieu = HethongTK.SoHieu" _
        & " SET ThongSo.MaDTSD = " _
        + CStr(SelectSQL("SELECT Top 1 MaSo as ReturnValue FROM HethongTK" _
                & " WHERE TK_ID2 = 3003 AND TKCon = 0 ORDER BY SoHieu")) _
        + " WHERE IsNULL(HethongTK.MaSo) OR (HethongTK.TK_ID2 <> 3003) "
    
    TinhGiaTriCuoiKy
    TaoChungTuKetChuyen
    On Error Resume Next
    DBKetoan.QueryDefs("KetChuyen").Execute
    On Error GoTo 0
    SoDuTKTS
End Sub

Private Sub TinhGiaTriCuoiKy()
Dim rs_giatri As Recordset
      SetSQL "TongKhauHao", "SELECT Sum(ThongSo.KH_NS) AS TKH_NS, Sum(ThongSo.KH_TBS) AS TKH_TBS, Sum(ThongSo.KH_CNK) AS TKH_CNK, Sum(ThongSo.KH_TD) AS TKH_TD, ThongSo.MaTS, First(TaiSan.ThangTang) As ThangT" _
            & ", Max(IIF(ThongSo.Thang = 12, MaDTQL,0)) As DTQL, Max(IIF(ThongSo.Thang = 12, MaDTSD,0)) As DTSD, Max(IIF(ThongSo.Thang = 12, MaTTSD,0)) As TTSD " _
            & "From TaiSan RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS WHERE ThongSo.Thang >= IIF(TaiSan.ThangTang = 0, " + CStr(pThangDauKy) + ", TaiSan.ThangTang) GROUP BY MaTS"
      SetSQL "TongGiaTri", "SELECT Sum(NG_NS) AS TNG_NS, Sum(NG_TBS) AS TNG_TBS, Sum(NG_CNK) AS TNG_CNK, Sum(NG_TD) AS TNG_TD, " _
            & "Sum(CL_NS) AS TCL_NS, Sum(CL_TBS) AS TCL_TBS, Sum(CL_CNK) AS TCL_CNK, Sum(CL_TD) AS TCL_TD, MaTS " _
            & "FROM CTTaiSan WHERE Thang < 13 GROUP BY MaTS"
      SetSQL "GiaTriTaiSan", "SELECT DISTINCTROW TongKhauHao.ThangT, TongKhauHao.DTQL, TongKhauHao.DTSD, TongKhauHao.TTSD, TNG_NS AS NG_NS, TNG_TBS AS NG_TBS, TNG_CNK AS NG_CNK, TNG_TD AS NG_TD, " _
            & "TCL_NS-TKH_NS AS CL_NS, TCL_TBS-TKH_TBS AS CL_TBS, TCL_CNK-TKH_CNK AS CL_CNK, TCL_TD-TKH_TD AS CL_TD, " _
            & "TongGiaTri.MaTS FROM TongKhauHao INNER JOIN TongGiaTri ON TongKhauHao.MaTS = TongGiaTri.MaTS"
      Set rs_giatri = DB1998.OpenRecordset("GiaTriTaiSan", dbOpenSnapshot, dbForwardOnly)
      Do While Not rs_giatri.EOF
            ExecuteSQL5 "UPDATE ThongSo SET NG_NS = " + CStr(rs_giatri!NG_NS) + ", NG_TBS = " + CStr(rs_giatri!NG_TBS) + ", NG_CNK = " + CStr(rs_giatri!NG_CNK) + ", NG_TD = " + CStr(rs_giatri!NG_TD) _
                  & ", CL_NS = " + CStr(rs_giatri!CL_NS) + ", CL_TBS = " + CStr(rs_giatri!CL_TBS) + ", CL_CNK = " + CStr(rs_giatri!CL_CNK) + ", CL_TD = " + CStr(rs_giatri!CL_TD) _
                  + " WHERE MaTS = " + CStr(rs_giatri!MaTS) + " And Thang = 0"
                  
            If rs_giatri!ThangT > 0 Then
                    Dim i As Integer
                    For i = 0 To rs_giatri!ThangT - 1
                            ExecuteSQL5 "INSERT INTO ThongSo (MaTS, Thang, NG_NS, NG_TBS, NG_CNK, NG_TD" _
                                & ", CL_NS, CL_TBS, CL_CNK, CL_TD, MaDTQL, MaDTSD, MaTTSD) VALUES (" + CStr(rs_giatri!MaTS) + "," + CStr(i) + "," + CStr(rs_giatri!NG_NS) _
                                + "," + CStr(rs_giatri!NG_TBS) + "," + CStr(rs_giatri!NG_CNK) + "," + CStr(rs_giatri!NG_TD) _
                                + "," + CStr(rs_giatri!CL_NS) + "," + CStr(rs_giatri!CL_TBS) + "," + CStr(rs_giatri!CL_CNK) _
                                + "," + CStr(rs_giatri!CL_TD) + "," + CStr(rs_giatri!DTQL) + "," + CStr(rs_giatri!DTSD) + "," + CStr(rs_giatri!TTSD) + ")"
                    Next
            End If
            rs_giatri.MoveNext
      Loop
      rs_giatri.Close
      Set rs_giatri = Nothing
    ' Xo¸ c¸c chøng tõ cña n¨m cò
      ExecuteSQL5 "DELETE * FROM CTTaiSan"
End Sub

Private Sub TaoChungTuKetChuyen()
Dim rs_thongso As Recordset
      Set rs_thongso = DBKetoan.OpenRecordset("SELECT DISTINCTROW TaiSan.MaSo, TaiSan.SoHieu, TaiSan.Ten, " _
            & "ThongSo.NG_NS, ThongSo.NG_TBS, ThongSo.NG_CNK, ThongSo.NG_TD, ThongSo.CL_NS, ThongSo.CL_TBS, ThongSo.CL_CNK, ThongSo.CL_TD " _
            & "FROM TaiSan RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS " _
            & "WHERE ThongSo.Thang=0", dbOpenSnapshot)
      Do Until rs_thongso.EOF
            With rs_thongso
            ExecuteSQL5 "INSERT INTO CTTaiSan (SoHieu, Thang, VaoSo, NgayGhi, DienGiai, " _
                  & "MaNhom, MaLoai, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD) " _
                  & "VALUES ('" + !SoHieu + "99', 0" _
                  + ",#" + Format(Date, "mm/dd/yy") + "#,#" + Format(Date, "mm/dd/yy") + "#,'" _
                  + "§Çu n¨m: " + !Ten + "'," + CStr(31) + "," + CStr(30) + "," + CStr(!MaSo) + "," _
                  + CStr(!NG_NS) + "," + CStr(!NG_TBS) + "," + CStr(!NG_CNK) + "," + CStr(!NG_TD) + "," _
                  + CStr(!CL_NS) + "," + CStr(!CL_TBS) + "," + CStr(!CL_CNK) + "," + CStr(!CL_TD) + ")"
            End With
            rs_thongso.MoveNext
      Loop
      rs_thongso.Close
      Set rs_thongso = Nothing
End Sub

Private Sub ChuyenCN()
Dim rs As Recordset
Dim tk As New ClsTaikhoan
    ExecuteSQL5 "DELETE FROM SoDuCongNo"
    ExecuteSQL5 "DELETE FROM DoituongCongNo"
    ExecuteSQL5 "DELETE FROM PhanLoaiCongNo"
    ExecuteSQL5 "DELETE FROM ThongSoCN"
    ExecuteSQL5 "DELETE FROM LoaiVay"
    
    ExecuteSQL5 "INSERT INTO LoaiVay (MaSo, SoHieu, DienGiai, ThoiHan)" _
        & " SELECT DISTINCTROW MaSo, SoHieu, DienGiai, ThoiHan FROM LoaiVay IN '" + p1998 + "'"
    ExecuteSQL5 "DELETE ThongSoCN.* FROM ThongSoCN LEFT JOIN LoaiVay ON ThongSoCN.MaLoaiVay = LoaiVay.MaSo IN '" + p1998 + "' WHERE IsNULL(LoaiVay.MaSo)"
    ExecuteSQL5 "INSERT INTO ThongSoCN (MaSo, Nam, MaLoaiVay, TuNgay, TyLe)" _
        & " SELECT DISTINCTROW MaSo, Nam + 1 AS NamM, MaLoaiVay, TuNgay, TyLe FROM ThongSoCN IN '" + p1998 + "'"
        
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW PhanLoaiCongNo.MaSo, PhanLoaiCongNo.SoHieu, PhanLoaiCongNo.TenPhanLoai, HethongTK.SoHieu AS SHTK" _
        & " FROM PhanLoaiCongNo INNER JOIN HethongTK ON PhanLoaiCongNo.MaTK = HethongTK.MaSo IN '" + p1998 + "'", dbOpenSnapshot)
    tk.SoHieu = "0"
    Do While Not rs.EOF
        If rs!shtk <> tk.SoHieu Then tk.InitTaikhoanSohieu rs!shtk
        ExecuteSQL5 "INSERT INTO PhanLoaiCongNo (MaSo, MaTK, SoHieu, TenPhanLoai)" _
            & " VALUES (" + CStr(rs!MaSo) + "," + CStr(tk.MaSo) + ",'" + rs!SoHieu + "','" + rs!TenPhanLoai + "')"
        rs.MoveNext
    Loop
    rs.Close
    
    ExecuteSQL5 "INSERT INTO DoituongCongNo (MaSo, MaPhanLoai, SoTK, HoTen, DiaChi, DienThoai, CongTy, PhongBan, CMND, NgayCap, NoiCap, NgayDK, GhiChu, SoDuMin)" _
        & " SELECT DISTINCTROW MaSo, MaPhanLoai, SoTK, HoTen, DiaChi, DienThoai, CongTy, PhongBan, CMND, NgayCap, NoiCap, NgayDK, GhiChu, SoDuMin FROM DoituongCongNo IN '" + p1998 + "'"
    ExecuteSQL5 "DELETE SoDuCongNo.* FROM SoDuCongNo LEFT JOIN LoaiVay ON SoDuCongNo.MaLoaiVay = LoaiVay.MaSo IN '" + p1998 + "' WHERE IsNull(LoaiVay.MaSo)"
    ExecuteSQL5 "INSERT INTO SoDuCongNo (MaSo, MaDTCN, MaLoaiVay, Nam, SoDu_0, LaiLK_0)" _
        & " SELECT DISTINCTROW MaSo, MaDTCN, MaLoaiVay, Nam + 1 AS NamM, SoDu_12, LaiLK_12 FROM SoDuCongNo IN '" + p1998 + "'"
    KiemTraCongNo
    SoDuTKCN
End Sub
Private Sub SoDuTKTS()
    Dim rs_tk As Recordset, taikhoan As New ClsTaikhoan
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT HeThongTK.SoHieu, Sum(CTTaiSan.NG_NS + CTTaiSan.NG_TBS + CTTaiSan.NG_CNK + CTTaiSan.NG_TD) As TNG" _
        & " FROM (LoaiTaiSan INNER JOIN (TaiSan INNER JOIN CTTaiSan ON TaiSan.MaSo = CTTaiSan.MaTS) ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) INNER JOIN HeThongTK ON LoaiTaiSan.SoHieu = HeThongTK.SoHieu" _
        & " Where (((CTTaiSan.maloai) = 30)) GROUP BY HeThongTK.SoHieu", dbOpenSnapshot, dbForwardOnly)
    
    Do While Not rs_tk.EOF
        taikhoan.InitTaikhoanSohieu rs_tk!SoHieu
        taikhoan.NoDauKy = rs_tk!TNG
        taikhoan.CapNhatTk
        rs_tk.MoveNext
    Loop
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT LEFT(HeThongTK.SoHieu,3) As SHTK, Sum((CTTaiSan.NG_NS + CTTaiSan.NG_TBS + CTTaiSan.NG_CNK + CTTaiSan.NG_TD) - (CTTaiSan.CL_NS + CTTaiSan.CL_TBS + CTTaiSan.CL_CNK + CTTaiSan.CL_TD)) AS THM" _
        & " FROM (LoaiTaiSan INNER JOIN (TaiSan INNER JOIN CTTaiSan ON TaiSan.MaSo = CTTaiSan.MaTS) ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) INNER JOIN HeThongTK ON LoaiTaiSan.SoHieu = HeThongTK.SoHieu" _
        & " Where (((CTTaiSan.maloai) = 30)) GROUP BY LEFT(HeThongTK.SoHieu,3)", dbOpenSnapshot, dbForwardOnly)
    
    Do While Not rs_tk.EOF
        taikhoan.InitTaikhoanSohieu "214" + Right(rs_tk!shtk, 1)
        taikhoan.CoDauKy = rs_tk!THM
        taikhoan.CapNhatTk
        rs_tk.MoveNext
    Loop
    
    rs_tk.Close
    Set rs_tk = Nothing
End Sub

Private Sub SoDuTKCN()
    Dim rs_tk As Recordset, taikhoan As New ClsTaikhoan
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW Sum(SoDuCongNo.SoDu_" + CStr(pThangDauKy - 1) + ") AS DuDk, PhanLoaiCongNo.MaTK" _
    & " FROM PhanLoaiCongNo RIGHT JOIN (DoituongCongNo RIGHT JOIN SoDuCongNo ON DoituongCongNo.MaSo = SoDuCongNo.MaDTCN) ON PhanLoaiCongNo.MaSo = DoituongCongNo.MaPhanLoai" _
    & " GROUP BY PhanLoaiCongNo.MaTK", dbOpenSnapshot, dbForwardOnly)
    
    Do While Not rs_tk.EOF
        taikhoan.InitTaikhoanMaSo rs_tk!MaTK
        taikhoan.NoDauKy = 0
        taikhoan.CoDauKy = rs_tk!DuDk
        taikhoan.CapNhatTk
        rs_tk.MoveNext
    Loop
    rs_tk.Close
    Set rs_tk = Nothing
End Sub

Private Sub KiemTraCongNo()
    Dim i As Integer, tyle As Double, sql As String
    Dim rs As Recordset
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo FROM LoaiVay", dbOpenSnapshot)
    Do While Not rs.EOF
        sql = "SELECT DISTINCTROW Top 1 TyLe AS ReturnValue FROM ThongSoCN WHERE MaLoaiVay = " + CStr(rs!MaSo) + " ORDER BY TuNgay DESC"
            
        tyle = SelectSQL(sql)
        For i = 1 To 12
            sql = "UPDATE SoDuCongNo SET Lai_" + CStr(i) + " = (SoDu_" + CStr(i - 1) + " + LaiLK_" + CStr(i - 1) + ") * " + DoiDau(Format(tyle, "Standard")) _
                + ", SoDu_" + CStr(i) + " = SoDu_" + CStr(i - 1) + ", LaiLK_" + CStr(i) + " = LaiLK_" + CStr(i - 1) + " + (SoDu_" + CStr(i - 1) + " + LaiLK_" + CStr(i - 1) + ") * " + DoiDau(Format(tyle, "Standard")) _
                + " WHERE MaLoaiVay = " + CStr(rs!MaSo)
            ExecuteSQL5 sql
        Next
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

