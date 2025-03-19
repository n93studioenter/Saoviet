Attribute VB_Name = "KhachHang"
Option Explicit

'======================================================================================
' Hµm tr¶ vÒ m· sè, tªn ph©n lo¹i tõ sè hiÖu
'======================================================================================
Public Function TenPLKH(sh As String, mpl As Long) As String
    Dim rs_tk As Recordset
        
    Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo,TenPhanLoai FROM PhanLoaiKhachHang WHERE SoHieu = '" + sh + "'", dbOpenSnapshot)
    If rs_tk.RecordCount > 0 Then
        mpl = rs_tk!MaSo
        TenPLKH = rs_tk!TenPhanLoai
    Else
        mpl = 0
        TenPLKH = ""
    End If
    
    rs_tk.Close
    Set rs_tk = Nothing
End Function

Public Sub SoDuTKCN()
    Dim rs_tk As Recordset, taikhoan As New ClsTaikhoan
    
    ExecuteSQL5 "UPDATE SoDuKhachHang SET DuNo_0=IIF(DuNo_0-DuCo_0>=0,DuNo_0-DuCo_0,0),DuCo_0=IIF(DuNo_0-DuCo_0<0,-DuNo_0+DuCo_0,0)"
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW SoHieu FROM HethongTK INNER JOIN SoDuKhachHang ON HethongTK.MaSo=SoDuKhachHang.MaTaiKhoan WHERE (TKCon = 0) AND (TK_ID = " + CStr(TKCNKH_ID) + " OR TK_ID = " + CStr(TKCNPT_ID) + ")", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_tk.EOF
        taikhoan.InitTaikhoanSohieu rs_tk!SoHieu
        taikhoan.NoDauKy = 0
        taikhoan.CoDauKy = 0
        taikhoan.CapNhatTk
        rs_tk.MoveNext
    Loop
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT MaTaiKhoan, Sum(DuNo_0) As DuNo, Sum(DuCo_0) As DuCo FROM SoDuKhachHang GROUP BY MaTaiKhoan", dbOpenSnapshot, dbForwardOnly)
    
    Do While Not rs_tk.EOF
        taikhoan.InitTaikhoanMaSo rs_tk!MaTaiKhoan
        If taikhoan.MaSo > 0 Then
            taikhoan.NoDauKy = rs_tk!duno
            taikhoan.CoDauKy = rs_tk!duco
            taikhoan.CapNhatTk
        End If
        rs_tk.MoveNext
    Loop
    
    rs_tk.Close
    Set rs_tk = Nothing
    Set taikhoan = Nothing
End Sub

Public Function TenKH(sh As String, mtk As Long, Optional mst As String) As String
    Dim rs_tk As Recordset
    If mtk > 0 Then
        Set rs_tk = DBKetoan.OpenRecordset("SELECT SoHieu, Ten, MST FROM KhachHang WHERE MaSo=" + CStr(mtk), dbOpenSnapshot)
        TenKH = rs_tk!Ten
        sh = rs_tk!SoHieu
        mst = rs_tk!mst
    Else
        Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo,Ten, MST FROM KhachHang WHERE SoHieu='" + sh + "'", dbOpenSnapshot)
        If rs_tk.RecordCount > 0 Then
            mtk = rs_tk!MaSo
            TenKH = rs_tk!Ten
            mst = rs_tk!mst
        Else
            mtk = 0
            TenKH = ""
            mst = ""
        End If
    End If
    rs_tk.Close
    Set rs_tk = Nothing
End Function

'======================================================================================
' Hµm tr¶ vÒ sè d­ tµi kho¶n cuèi ngµy, chØ cho tµi kho¶n chi tiÕt
'======================================================================================
Public Sub SoDuKHNgay(mkh As Long, ngay As Date, duno As Double, duco As Double, dunt As Double, Optional mtk As Long = 0, Optional shtk As String = "")
    Dim rs As Recordset, thang As Integer, sql As String, X As Double, y1 As Double, y2 As Double, kieu As Boolean
    
    If mtk > 0 Then
        Set rs = DBKetoan.OpenRecordset("SELECT Sum(DuNo_0) AS n,Sum(DuCo_0) AS c,Sum(DuNT_0) AS nt FROM SoDuKhachHang WHERE MaKhachHang=" + CStr(mkh) + IIf(mtk > 0, " AND MaTaiKhoan=" + CStr(mtk), ""), dbOpenSnapshot)
    Else
        Set rs = DBKetoan.OpenRecordset("SELECT Sum(SoDuKhachHang.DuNo_0) AS n,Sum(SoDuKhachHang.DuCo_0) AS c,Sum(SoDuKhachHang.DuNT_0) AS nt FROM SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo WHERE MaKhachHang=" + CStr(mkh) + " AND SoHieu LIKE '" + shtk + "*'", dbOpenSnapshot)
    End If
        
    If Not IsNull(rs!n) Then
        duno = rs!n
        duco = rs!c
        dunt = rs!nt
    Else
        duno = 0
        duco = 0
        dunt = 0
    End If
    kieu = duno >= duco
    X = SelectSQL("SELECT DISTINCTROW Sum(SoPS) AS F1, Sum(SoPS2No) AS F2 FROM " + ChungTu2TKNC(-1) + " WHERE (TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ") AND NgayGS<=#" + Format(ngay, Mask_DB) + "# AND MaKH=" + CStr(mkh) + IIf(mtk > 0, " AND MaTKNo=" + CStr(mtk), ""), y1)
    duno = duno + X
    X = SelectSQL("SELECT DISTINCTROW Sum(SoPS) AS F1, Sum(SoPS2Co) AS F2 FROM " + ChungTu2TKNC(1) + " WHERE (TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ") AND NgayGS<=#" + Format(ngay, Mask_DB) + "# AND MaKHC=" + CStr(mkh) + IIf(mtk > 0, " AND MaTKCo=" + CStr(mtk), ""), y2)
    duco = duco + X
    If duno - duco >= 0 Then
        duno = duno - duco
        duco = 0
        If kieu Then dunt = dunt + y1 - y2 Else dunt = dunt - y1 + y2
    Else
        duco = duco - duno
        duno = 0
        If kieu Then dunt = dunt - y1 + y2 Else dunt = dunt + y1 - y2
    End If
    rs.Close
    Set rs = Nothing
End Sub

Public Sub SoPhatSinhN(mkh As Long, ndau As Date, ncuoi As Date, psn As Double, psc As Double, psn2 As Double, psc2 As Double, Optional mtk As Long)
    psn = SelectSQL("SELECT DISTINCTROW Sum(SoPS) AS F1, Sum(SoPS2No) AS F2 FROM " + ChungTu2TKNC(-1) + " WHERE (TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ") AND " + WNgay("NgayGS", ndau, ncuoi) + " AND MaKH=" + CStr(mkh) + IIf(mtk > 0, " AND MaTKNo=" + CStr(mtk), ""), psn2)
    psc = SelectSQL("SELECT DISTINCTROW Sum(SoPS) AS F1, Sum(SoPS2Co) AS F2 FROM " + ChungTu2TKNC(1) + " WHERE (TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ") AND " + WNgay("NgayGS", ndau, ncuoi) + " AND MaKHC=" + CStr(mkh) + IIf(mtk > 0, " AND MaTKCo=" + CStr(mtk), ""), psc2)
End Sub

Public Sub DanhDiemCN(mpl As Long)
    Dim sql As String
    
    SetSQL "QChitiet", "SELECT DISTINCTROW KhachHang.SoHieu, Ten,PhanLoaiKhachHang.SoHieu AS SHPL,TenPhanLoai,PhanLoaiKhachHang.PLCha,DiaChi,MST,Tel,Fax,TaiKhoan,PhanLoaiKhachHang.MaSo AS MPL" _
        & " FROM KhachHang INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo WHERE LEFT(KhachHang.SoHieu,1)<>'X'"
    SetSQL "QSoCai", "SELECT QChitiet.*, PhanLoaiKhachHang.SoHieu AS SH2,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai2,PhanLoaiKhachHang.PLCha AS PLCha2,PhanLoaiKhachHang.MaSo AS MPL2 FROM QChitiet LEFT JOIN PhanLoaiKhachHang ON QChitiet.PLCha=PhanLoaiKhachHang.MaSo"
    SetSQL "QTongHopCT", "SELECT QSoCai.*, PhanLoaiKhachHang.SoHieu AS SH1,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai1 FROM QSoCai LEFT JOIN PhanLoaiKhachHang ON QSoCai.PLCha2=PhanLoaiKhachHang.MaSo " _
        + IIf(mpl > 0, "WHERE PhanLoaiKhachHang.MaSo=" + CStr(mpl) + " OR MPL=" + CStr(mpl) + " OR MPL2=" + CStr(mpl), "") + " ORDER BY QSoCai.SoHieu"
    frmMain.Rpt.ReportFileName = "DDCN.RPT"
    frmMain.Rpt.WindowTitle = "B¶ng danh ®iÓm c«ng nî"
End Sub

Public Sub SoDuTKCN2(thang As Integer)
    Dim rs_tk As Recordset, st As String
    Dim n As Double, c As Double, TK As New ClsTaikhoan
    
     st = CStr(CThangDB(thang))
    Set rs_tk = DBKetoan.OpenRecordset("SELECT MaTaiKhoan, Sum(SoDuKhachHang.DuNo_" + st + ") As DuNo, Sum(SoDuKhachHang.DuCo_" + st + ") As DuCo FROM SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + pSHPT + "*' OR HethongTK.SoHieu LIKE '331*' GROUP BY MaTaiKhoan", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_tk.EOF
        n = SelectSQL("SELECT DuNo_" + st + " As F1, DuCo_" + st + " As F2 FROM HethongTK WHERE MaSo=" + CStr(rs_tk!MaTaiKhoan), c)
        TK.InitTaikhoanMaSo rs_tk!MaTaiKhoan
        ExecuteSQL5 "UPDATE HethongTK SET DuNo_" + st + " = DuNo_" + st + " -  " + DoiDau(n) + " + " + DoiDau(rs_tk!duno) + ",DuCo_" + st + " = DuCo_" + st + "-" + DoiDau(c) + "+" + DoiDau(rs_tk!duco) + " WHERE MaSo = " + CStr(TK.MaSo) + " OR MaSo = " + CStr(TK.TkCha0) + " OR MaSo = " + CStr(TK.TkCha1) + " OR MaSo = " + CStr(TK.TkCha2) + " OR MaSo = " + CStr(TK.TkCha3) + " OR MaSo = " + CStr(TK.TkCha4) + " OR MaSo = " + CStr(TK.TkCha5)
        rs_tk.MoveNext
    Loop
        
    rs_tk.Close
    Set rs_tk = Nothing
    Set TK = Nothing
End Sub

Public Function KHMaNT(mkh As Long)
    KHMaNT = SelectSQL("SELECT MaNT AS F1 FROM KhachHang WHERE MaSo=" + CStr(mkh))
End Function

Public Function TenNV(sh As String, ms As Long) As String
    TenNV = SelectSQL("SELECT SoHieu AS F2, Ten AS F1, MaSo AS F3 FROM NhanVien WHERE " + IIf(ms > 0, "MaSo=" + CStr(ms), "SoHieu='" + sh + "'"), sh, ms)
End Function

Public Function DoanhThuTK(shtk As String, tdau As Integer, tcuoi As Integer, thue As Double) As Double
    
    SetSQL "MienTru", "SELECT MaCT, First(TyLe) AS TL FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY MaCT"
    DoanhThuTK = SelectSQL("SELECT Sum(SoPS) AS F1, Sum(iif(IsNull(TL),0,Fix(0.5+SoPS*TL/100))) AS F2 FROM (" + ChungTu2TKNC(1) + ") LEFT JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT WHERE HethongTK.SoHieu LIKE '" + shtk + "*' AND " + WThang("ThangCT", tdau, tcuoi), thue)
    DoanhThuTK = DoanhThuTK - SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE MaLoai<>3 AND HethongTK.SoHieu LIKE '" + shtk + "*' AND (Not TK.SoHieu LIKE '3332*') AND " + WThang("ThangCT", tdau, tcuoi))
End Function

Public Sub SoDuKH(mkh As Long, thang As Integer, duno As Double, duco As Double, dunt As Double, Optional mtk As Long = 0, Optional shtk As String = "")
    Dim rs As Recordset, st As String
    
    st = CStr(CThangDB(thang))
    If mtk > 0 Then
        Set rs = DBKetoan.OpenRecordset("SELECT Sum(DuNo_" + st + ") AS n,Sum(DuCo_" + st + ") AS c,Sum(DuNT_" + st + ") AS nt FROM SoDuKhachHang WHERE MaKhachHang=" + CStr(mkh) + IIf(mtk > 0, " AND MaTaiKhoan=" + CStr(mtk), ""), dbOpenSnapshot)
    Else
        Set rs = DBKetoan.OpenRecordset("SELECT Sum(SoDuKhachHang.DuNo_" + st + ") AS n,Sum(SoDuKhachHang.DuCo_" + st + ") AS c,Sum(IIF(SoDuKhachHang.DuCo_" + st + ">0,SoDuKhachHang.DuNT_" + st + ",-SoDuKhachHang.DuNT_" + st + ")) AS nt FROM SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo WHERE MaKhachHang=" + CStr(mkh) + " AND SoHieu LIKE '" + shtk + "*'", dbOpenSnapshot)
    End If
    If Not IsNull(rs!n) Then
        duno = rs!n
        duco = rs!c
        dunt = Abs(rs!nt)
    Else
        duno = 0
        duco = 0
        dunt = 0
    End If
    If duno < 0 Then
        duco = -duno
        duno = 0
    End If
    If duco < 0 Then
        duno = -duco
        duco = 0
    End If
    rs.Close
    Set rs = Nothing
End Sub

Public Function SoNoTheoHoaDon(mtk As Long, mkh As Long, loai As Integer) As Double
    Dim sodu As Double
    
    If pCongNoHD = 0 Then Exit Function
    
    loai = SelectSQL("SELECT Kieu AS F1 FROM HethongTK WHERE MaSo=" + CStr(mtk))
    If loai < 0 Then
        sodu = SelectSQL("SELECT Sum(SoPS-SoXuat) AS F1 FROM ChungTu WHERE MaTKNo=" + CStr(mtk) + " AND MaKH=" + CStr(mkh) + " AND SoPS>SoXuat")
        sodu = sodu + SelectSQL("SELECT Sum(DuNo_0-SoXuat) AS F1 FROM CNDauNam WHERE MaTaiKhoan=" + CStr(mtk) + " AND MaKhachHang=" + CStr(mkh) + " AND DuNo_0>SoXuat")
    Else
        sodu = SelectSQL("SELECT Sum(SoPS-SoXuat) AS F1 FROM ChungTu WHERE MaTKCo=" + CStr(mtk) + " AND MaKHC=" + CStr(mkh) + " AND SoPS>SoXuat")
        sodu = sodu + SelectSQL("SELECT Sum(DuCo_0-SoXuat) AS F1 FROM CNDauNam WHERE MaTaiKhoan=" + CStr(mtk) + " AND MaKhachHang=" + CStr(mkh) + " AND DuCo_0>SoXuat")
    End If
    SoNoTheoHoaDon = sodu
End Function

Public Sub BangKeTichSo2(shtk As String, mcn As Long, ndau As Date, ncuoi As Date, Optional nx As Integer = 0)
    Dim rs As Recordset, n As Date, sql As String, ms As Long, sdn As Double, sdc As Double, n1 As Date
    Dim sh As String, X As Double, rs2 As Recordset, k As Integer, mtk As Long, mbc As Long
    
    If shtk <> "" Then
        sh = shtk
        k = SelectSQL("SELECT TOP 1 Kieu AS F1,MaSo AS F2 FROM HethongTK WHERE SoHieu LIKE '" + sh + "*'", mtk)
    
        sql = "SELECT -1 AS LoaiPS,NgayGS,Sum(SoPS) AS PS FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo WHERE MaKH=" + CStr(mcn) + " AND " + WNgay("NgayGS", ndau, ncuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*' AND CT_ID<>700000000 GROUP BY NgayGS" _
            & " UNION SELECT 1 AS LoaiPS,NgayGS,Sum(SoPS) AS PS FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKNo=TK.MaSo WHERE MaKHC=" + CStr(mcn) + " AND " + WNgay("NgayGS", ndau, ncuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*' AND CT_ID<>700000000 GROUP BY NgayGS"
        SetSQL "MienTru", sql
        Set rs = DBKetoan.OpenRecordset("SELECT * FROM MienTru ORDER BY NgayGS", dbOpenSnapshot)
        SoDuKHNgay mcn, ndau - 1, sdn, sdc, X, , sh
        If nx = 0 Then ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
        mbc = Lng_MaxValue("MaSo", "BaoCaoCP2")
        n = ndau
        ms = mbc + 1
        ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,MaCha,Kq3,Kq4,Ngay1) VALUES (" + CStr(ms) + ",'" + tentk(shtk, 0) + "','" + shtk + "'," + CStr(mtk) + "," + CStr(sdn) + "," + CStr(sdc) + ",#" + Format(n, Mask_DB) + "#)"
        Do While Not rs.EOF
            If rs!loaips < 0 Then sdn = sdn + rs!ps Else sdc = sdc + rs!ps
            If sdn - sdc >= 0 Then
                sdn = sdn - sdc
                sdc = 0
            Else
                sdc = sdc - sdn
                sdn = 0
            End If
            If rs!NgayGS <> n Then
                ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=" + CStr(rs!NgayGS - n) + " WHERE MaSo=" + CStr(ms)
                n = rs!NgayGS
                ms = ms + 1
                ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,MaCha,Kq1,Kq2,Kq3,Kq4,Ngay1) VALUES (" + CStr(ms) + ",'" + CStr(ms) + "','" + shtk + "'," + CStr(mtk) + "," + IIf(rs!loaips < 0, CStr(rs!ps), "0") + "," + IIf(rs!loaips > 0, CStr(rs!ps), "0") + "," + CStr(sdn) + "," + CStr(sdc) + ",#" + Format(n, Mask_DB) + "#)"
            Else
                ExecuteSQL5 "UPDATE BaoCaoCP2 SET " + IIf(rs!loaips < 0, "Kq1=Kq1+" + CStr(rs!ps), "Kq2=Kq2+" + CStr(rs!ps)) + ",Kq3=" + CStr(sdn) + ",Kq4=" + CStr(sdc) + " WHERE MaSo=" + CStr(ms)
            End If
            rs.MoveNext
        Loop
        rs.Close
        
        n1 = ncuoi
        If n <> n1 Then
            ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=" + CStr(n1 - n + 1) + " WHERE MaSo=" + CStr(ms)
        Else
            ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=1 WHERE MaSo=" + CStr(ms)
        End If
        ExecuteSQL5 "DELETE * FROM BaoCaoCP2 WHERE Kq1=0 AND Kq2=0 AND Kq3=0 AND Kq4=0 AND MaSo>" + CStr(mbc)
        n = SelectSQL("SELECT TOP 1 Ngay1 AS F1 FROM BaoCaoCP2 WHERE MaSo>" + CStr(mbc) + " ORDER BY Ngay1")
        X = SelectSQL("SELECT TOP 1 LS AS F1 FROM LaiSuat INNER JOIN HethongTK ON LaiSuat.MaTK=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + sh + "*' AND Ngay<=#" + Format(n, Mask_DB) + "# AND MaKH=" + CStr(mcn) + " ORDER BY Ngay DESC")
        'If X = 0 Then X = SelectSQL("SELECT TOP 1 LS AS F1 FROM LaiSuat INNER JOIN HethongTK ON LaiSuat.MaTK=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + sh + "*' AND Ngay<=#" + Format(n, Mask_DB) + "# ORDER BY Ngay DESC")
        ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq6=" + DoiDau(Format(X, Mask_2)) + " WHERE MaSo>" + CStr(mbc)
        n = ncuoi + 1
        Set rs = DBKetoan.OpenRecordset("SELECT * FROM BaoCaoCP2 WHERE MaSo>" + CStr(mbc) + " ORDER BY Ngay1 DESC", dbOpenSnapshot)
        Do While Not rs.EOF
            Set rs2 = DBKetoan.OpenRecordset("SELECT Ngay,LS FROM LaiSuat INNER JOIN HethongTK ON LaiSuat.MaTK=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + sh + "*' AND Ngay>=#" + Format(rs!ngay1, Mask_DB) + "# AND Ngay<#" + Format(n, Mask_DB) + "# ORDER BY Ngay DESC", dbOpenSnapshot)
            Do While Not rs2.EOF
                ms = ms + 1
                ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,MaCha,Cap,Kq3,Kq4,Ngay1,Kq6) VALUES (" + CStr(ms) + ",'" + CStr(ms) + "','" + shtk + "'," + CStr(mtk) + "," + CStr(n - rs2!ngay) + "," + CStr(rs!Kq3) + "," + CStr(rs!Kq4) + ",#" + Format(rs2!ngay, Mask_DB) + "#," + DoiDau(Format(rs2!ls, Mask_2)) + ")"
                ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=Cap-" + CStr(n - rs2!ngay) + " WHERE MaSo=" + CStr(rs!MaSo)
                n = rs2!ngay
                rs2.MoveNext
            Loop
            n = rs!ngay1
            rs2.Close
            rs.MoveNext
        Loop
        Set rs2 = Nothing
        ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq5=Kq" + IIf(k < 0, "3", "4") + "*Cap WHERE MaSo>" + CStr(mbc)
        ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq7=Fix(0.5+Kq5*Kq6/3000) WHERE MaSo>" + CStr(mbc)
        ExecuteSQL5 "DELETE * FROM BaoCaoCP2 WHERE Cap=0 AND MaSo>" + CStr(mbc)
    Else
        ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
        Set rs = DBKetoan.OpenRecordset("SELECT HethongTK.SoHieu FROM SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo WHERE MaKhachHang=" + CStr(mcn) + " ORDER BY HethongTK.SoHieu", dbOpenSnapshot)
        Do While Not rs.EOF
            BangKeTichSo rs!SoHieu, mcn, ndau, ncuoi, 1
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    
    If nx = 0 Then
        frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\BKTSCN.RPT"
        frmMain.Rpt.WindowTitle = "B¶ng kª tÝch sè"
        frmMain.Rpt.Formulas(4) = "ThoiGian='Tõ ngµy " + Format(ndau, Mask_DR) + " ®Õn " + Format(ncuoi, Mask_DR) + "'"
        'frmMain.Rpt.Formulas(5) = "TaiKhoan='" + sh + " - " + TenTK(sh, 0) + "'"
        frmMain.Rpt.Formulas(6) = "KH='" + MaSo2SoHieu(mcn, "KhachHang") + " - " + TenKH("", mcn, 0) + "'"
        RptSetDate ncuoi
    End If
End Sub

Public Sub BangKeTichSo(shtk As String, mcn As Long, ndau As Date, ncuoi As Date, Optional nx As Integer = 0)
    Dim rs As Recordset, n As Date, sql As String, ms As Long, sdn As Double, sdc As Double, n1 As Date
    Dim sh As String, X As Double, rs2 As Recordset, k As Integer, mtk As Long, mbc As Long
    
    If shtk <> "" Then
        sh = shtk
        k = SelectSQL("SELECT TOP 1 Kieu AS F1,MaSo AS F2 FROM HethongTK WHERE SoHieu LIKE '" + sh + "*'", mtk)
    
        sql = "SELECT -1 AS LoaiPS,NgayGS,Sum(SoPS) AS PS FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo WHERE MaKH=" + CStr(mcn) + " AND " + WNgay("NgayGS", ndau, ncuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*' AND CT_ID<>700000000 GROUP BY NgayGS" _
            & " UNION SELECT 1 AS LoaiPS,NgayGS,Sum(SoPS) AS PS FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKNo=TK.MaSo WHERE MaKHC=" + CStr(mcn) + " AND " + WNgay("NgayGS", ndau, ncuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*' AND CT_ID<>700000000 GROUP BY NgayGS"
        SetSQL "MienTru", sql
        Set rs = DBKetoan.OpenRecordset("SELECT * FROM MienTru ORDER BY NgayGS", dbOpenSnapshot)
        SoDuKHNgay mcn, ndau - 1, sdn, sdc, X, , sh
        If nx = 0 Then ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
        mbc = Lng_MaxValue("MaSo", "BaoCaoCP2")
        n = ndau
        ms = mbc + 1
        ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,MaCha,Kq3,Kq4,Ngay1) VALUES (" + CStr(ms) + ",'" + tentk(shtk, 0) + "','" + shtk + "'," + CStr(mtk) + "," + CStr(sdn) + "," + CStr(sdc) + ",#" + Format(n, Mask_DB) + "#)"
        Do While Not rs.EOF
            If rs!loaips < 0 Then sdn = sdn + rs!ps Else sdc = sdc + rs!ps
            If sdn - sdc >= 0 Then
                sdn = sdn - sdc
                sdc = 0
            Else
                sdc = sdc - sdn
                sdn = 0
            End If
            If rs!NgayGS <> n Then
                ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=" + CStr(rs!NgayGS - n) + " WHERE MaSo=" + CStr(ms)
                n = rs!NgayGS
                ms = ms + 1
                ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,MaCha,Kq1,Kq2,Kq3,Kq4,Ngay1) VALUES (" + CStr(ms) + ",'" + CStr(ms) + "','" + shtk + "'," + CStr(mtk) + "," + IIf(rs!loaips < 0, CStr(rs!ps), "0") + "," + IIf(rs!loaips > 0, CStr(rs!ps), "0") + "," + CStr(sdn) + "," + CStr(sdc) + ",#" + Format(n, Mask_DB) + "#)"
            Else
                ExecuteSQL5 "UPDATE BaoCaoCP2 SET " + IIf(rs!loaips < 0, "Kq1=Kq1+" + CStr(rs!ps), "Kq2=Kq2+" + CStr(rs!ps)) + ",Kq3=" + CStr(sdn) + ",Kq4=" + CStr(sdc) + " WHERE MaSo=" + CStr(ms)
            End If
            rs.MoveNext
        Loop
        rs.Close
        
        n1 = ncuoi
        If n <> n1 Then
            ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=" + CStr(n1 - n + 1) + " WHERE MaSo=" + CStr(ms)
        Else
            ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=1 WHERE MaSo=" + CStr(ms)
        End If
        ExecuteSQL5 "DELETE * FROM BaoCaoCP2 WHERE Kq1=0 AND Kq2=0 AND Kq3=0 AND Kq4=0 AND MaSo>" + CStr(mbc)
        n = SelectSQL("SELECT TOP 1 Ngay1 AS F1 FROM BaoCaoCP2 WHERE MaSo>" + CStr(mbc) + " ORDER BY Ngay1")
        X = SelectSQL("SELECT TOP 1 LS AS F1 FROM LaiSuat INNER JOIN HethongTK ON LaiSuat.MaTK=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + sh + "*' AND Ngay<=#" + Format(n, Mask_DB) + "# AND MaKH=" + CStr(mcn) + " ORDER BY Ngay DESC")
        'If X = 0 Then X = SelectSQL("SELECT TOP 1 LS AS F1 FROM LaiSuat INNER JOIN HethongTK ON LaiSuat.MaTK=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + sh + "*' AND Ngay<=#" + Format(n, Mask_DB) + "# ORDER BY Ngay DESC")
        ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq6=" + DoiDau(Format(X, Mask_2)) + " WHERE MaSo>" + CStr(mbc)
        n = ncuoi + 1
        Set rs = DBKetoan.OpenRecordset("SELECT * FROM BaoCaoCP2 WHERE MaSo>" + CStr(mbc) + " ORDER BY Ngay1 DESC", dbOpenSnapshot)
        Do While Not rs.EOF
            Set rs2 = DBKetoan.OpenRecordset("SELECT Ngay,LS FROM LaiSuat INNER JOIN HethongTK ON LaiSuat.MaTK=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + sh + "*' AND Ngay>=#" + Format(rs!ngay1, Mask_DB) + "# AND Ngay<#" + Format(n, Mask_DB) + "# ORDER BY Ngay DESC", dbOpenSnapshot)
            Do While Not rs2.EOF
                ms = ms + 1
                ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,MaCha,Cap,Kq3,Kq4,Ngay1,Kq6) VALUES (" + CStr(ms) + ",'" + CStr(ms) + "','" + shtk + "'," + CStr(mtk) + "," + CStr(n - rs2!ngay) + "," + CStr(rs!Kq3) + "," + CStr(rs!Kq4) + ",#" + Format(rs2!ngay, Mask_DB) + "#," + DoiDau(Format(rs2!ls, Mask_2)) + ")"
                ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=Cap-" + CStr(n - rs2!ngay) + " WHERE MaSo=" + CStr(rs!MaSo)
                n = rs2!ngay
                rs2.MoveNext
            Loop
            n = rs!ngay1
            rs2.Close
            rs.MoveNext
        Loop
        Set rs2 = Nothing
        ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq5=Kq" + IIf(k < 0, "3", "4") + "*Cap WHERE MaSo>" + CStr(mbc)
        ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq7=Fix(0.5+Kq5*Kq6/3000) WHERE MaSo>" + CStr(mbc)
        ExecuteSQL5 "DELETE * FROM BaoCaoCP2 WHERE Cap=0 AND MaSo>" + CStr(mbc)
    Else
        ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
        Set rs = DBKetoan.OpenRecordset("SELECT HethongTK.SoHieu FROM SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo WHERE MaKhachHang=" + CStr(mcn) + " ORDER BY HethongTK.SoHieu", dbOpenSnapshot)
        Do While Not rs.EOF
            BangKeTichSo rs!SoHieu, mcn, ndau, ncuoi, 1
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    
    If nx = 0 Then
        frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\BKTS.RPT"
        frmMain.Rpt.WindowTitle = "B¶ng kª tÝch sè"
        frmMain.Rpt.Formulas(4) = "ThoiGian='Tõ ngµy " + Format(ndau, Mask_DR) + " ®Õn " + Format(ncuoi, Mask_DR) + "'"
        'frmMain.Rpt.Formulas(5) = "TaiKhoan='" + sh + " - " + TenTK(sh, 0) + "'"
        frmMain.Rpt.Formulas(6) = "KH='" + MaSo2SoHieu(mcn, "KhachHang") + " - " + TenKH("", mcn, 0) + "'"
        RptSetDate ncuoi
    End If
End Sub


