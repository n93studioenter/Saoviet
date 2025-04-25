Attribute VB_Name = "modVattu"
Option Explicit
'======================================================================================
' Thñ tôc tÝnh l¹i tån kho cña mét lo¹i vËt t­ sau khi nhËp xuÊt
'======================================================================================
Public Function TinhTonKho(mkho As Long, mtk As Long, mvt As Long, thang As Integer, loai As Integer, SoLuong As Double, ThanhTien As Double, ThanhTien2 As Double) As Integer
    Dim th As String, i As Integer, stl As String, stt As String, stt2 As String
    Dim sql As String, j As Integer
    
    th = CStr(CThangDB(thang))
        
    stl = DoiDau(SoLuong)
    If loai < 0 Then
        sql = "UPDATE TonKho SET Luong_Nhap_" + th + " = Luong_Nhap_" + th + " + " + stl _
            + ", Tien_Nhap_" + th + " = Tien_Nhap_" + th + " + " + DoiDau(ThanhTien) + IIf(pGiaUSD > 0, ", USDTien_Nhap_" + th + " = USDTien_Nhap_" + th + " + " + DoiDau(ThanhTien2), "")
    Else
        sql = "UPDATE TonKho SET Luong_Xuat_" + th + " = Luong_Xuat_" + th + " + " + stl _
            + ", Tien_Xuat_" + th + " = Tien_Xuat_" + th + " + " + DoiDau(ThanhTien) + IIf(pGiaUSD > 0, ", USDTien_Xuat_" + th + " = USDTien_Xuat_" + th + " + " + DoiDau(ThanhTien2), "")
    End If
    
    If loai < 0 Then
        stl = " + " + stl
    Else
        stl = " - " + stl
    End If
    stt = DoiDau(IIf(loai < 0, ThanhTien, -ThanhTien))
    stt2 = DoiDau(IIf(loai < 0, ThanhTien2, -ThanhTien2))
    
    For i = CThangDB(thang) To 12
        th = CStr(i)
        sql = sql + ", Luong_" + th + " = Luong_" + th + stl + ", Tien_" + th + " = Tien_" + th + " + " + stt
        If pGiaUSD > 0 Then sql = sql + ", USDTien_" + th + " = USDTien_" + th + " + " + stt2
    Next
    TinhTonKho = ExecuteSQL5(sql + " WHERE MaSoKho=" + CStr(mkho) + " AND MaTaiKhoan=" + CStr(mtk) + " AND MaVatTu=" + CStr(mvt))
    
    If DBKetoan.RecordsAffected = 0 Then
        j = CThangDB(thang)
        sql = "INSERT INTO TonKho (MaSo,MaSoKho,MaTaiKhoan,MaVattu,Luong_" + IIf(loai < 0, "Nhap", "Xuat") + "_" + CStr(j) + ",Tien_" + IIf(loai < 0, "Nhap", "Xuat") + "_" + CStr(j) + IIf(pGiaUSD > 0, ",USDTien_" + IIf(loai < 0, "Nhap", "Xuat") + "_" + CStr(j), "")
        For i = j To 12
            sql = sql + ", Luong_" + CStr(i) + ", Tien_" + CStr(i) + IIf(pGiaUSD > 0, ", USDTien_" + CStr(i), "")
        Next
        sql = sql + ") VALUES (" + CStr(Lng_MaxValue("MaSo", "TonKho") + 1) + "," + CStr(mkho) + "," + CStr(mtk) + "," + CStr(mvt) + ",abs(" + stl + "),abs(" + stt + ")" + IIf(pGiaUSD > 0, ",abs(" + stt2 + ")", "")
        For i = j To 12
            sql = sql + ", " + stl + ", " + stt + IIf(pGiaUSD > 0, "," + stt2, "")
        Next
        sql = sql + ")"
        TinhTonKho = ExecuteSQL5(sql, False)
    End If
End Function
'======================================================================================
' Tra ve so luong ton va thanh tien cua mot vat tu trong kho
' ngay:     ngay tinh ton kho
' mkho, mvt:     ma so kho, vattu can tinh ton kho
' thanhtien:gia tri ton kho
' Trave:    so luong ton kho
'======================================================================================
Public Function SoTonKho(thang As Integer, mkho As Long, mtk As Long, mvt As Long, ThanhTien As Double, tien2 As Double) As Double
    Dim sql As String
    
    sql = "SELECT SUM(Tien_" + CStr(CThangDB(thang)) + ") As F1,SUM(Luong_" + CStr(CThangDB(thang)) + ") As F2" + IIf(pGiaUSD > 0, ",SUM(USDTien_" + CStr(CThangDB(thang)) + ") As F3", "") + " FROM TonKho WHERE (True)"
    If mkho > 0 Then sql = sql + " AND MaSoKho=" + CStr(mkho)
    If mtk > 0 Then sql = sql + " And MaTaiKhoan=" + CStr(mtk)
    If mvt > 0 Then sql = sql + " And MaVattu=" + CStr(mvt)
    
    ThanhTien = SelectSQL(sql, SoTonKho, tien2)
End Function

Public Function SoTonKhoN(ngay As Date, mkho As Long, mtk As Long, mvt As Long, ThanhTien As Double, tien2 As Double) As Double
    Dim sql As String, luong As Double, X As Double
    
    SoTonKhoN = SoTonKho(0, mkho, mtk, mvt, ThanhTien, tien2)
    sql = "SELECT Sum(SoPS) As F1,Sum(SoPS2No) As F2" + IIf(pGiaUSD > 0, ",Sum(PSUSD) AS F3", "") + " FROM ChungTu WHERE MaVattu=" + CStr(mvt) + " AND (MaLoai=1" + IIf(mkho > 0, " OR MaLoai=4)", ")") + " AND NgayGS<=#" + Format(ngay, Mask_DB) + "#"
    If mkho > 0 Then sql = sql + " AND ((MaKho=" + CStr(mkho) + " AND MaLoai=1) OR (MaNguon=" + CStr(mkho) + " AND MaLoai=4))"
    If mtk > 0 Then sql = sql + " And MaTKNo=" + CStr(mtk)
    
    ThanhTien = ThanhTien + SelectSQL(sql, luong, X)
    SoTonKhoN = SoTonKhoN + luong
    tien2 = tien2 + X
    
    sql = "SELECT Sum(SoPS) As F1,Sum(SoPS2Co) As F2" + IIf(pGiaUSD > 0, ",Sum(PSUSD) AS F3", "") + " FROM ChungTu WHERE MaVattu=" + CStr(mvt) + " AND (MaLoai=2" + IIf(mkho > 0, " OR MaLoai=4)", ")") + " AND NgayGS<=#" + Format(ngay, Mask_DB) + "#"
    If mkho > 0 Then sql = sql + " AND (MaKho=" + CStr(mkho) + ")"
    If mtk > 0 Then sql = sql + " And MaTKCo=" + CStr(mtk)
    
    ThanhTien = ThanhTien - SelectSQL(sql, luong, X)
    tien2 = tien2 - X
    SoTonKhoN = SoTonKhoN - luong
End Function
'======================================================================================
' Hµm tr¶ vÒ sè hiÖu mÆc ®Þnh cña vËt t­ míi
'======================================================================================
Public Function SoHieuVTMoi(mpl As Long, Optional loai As Integer = 0) As String
    Dim rs_vt As Recordset, tail As String, dai As Integer, i As Integer, sql As String
    
    Select Case loai
        Case 0:
            sql = "SELECT Top 1 SoHieu FROM Vattu WHERE MaPhanLoai = " + CStr(mpl) + " ORDER BY SoHieu DESC"
        Case 1:
            sql = "SELECT Top 1 SoHieu FROM TP154 WHERE MaPhanLoai = " + CStr(mpl) + " ORDER BY SoHieu DESC"
        Case 2:
            sql = "SELECT Top 1 SoHieu FROM KhachHang WHERE MaPhanLoai = " + CStr(mpl) + " ORDER BY SoHieu DESC"
        Case 3:
            sql = "SELECT Top 1 SoHieu FROM NhanVien WHERE MaPhanLoai = " + CStr(mpl) + " ORDER BY SoHieu DESC"
        Case 4:
            sql = "SELECT Top 1 SoHieu FROM DoituongCT ORDER BY SoHieu DESC"
    End Select
    Set rs_vt = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs_vt.recordCount > 0 Then
        i = 1
        Do While IsNumeric(Right(rs_vt!sohieu, i)) And i <= Len(rs_vt!sohieu)
            dai = i
            i = i + 1
        Loop
        If dai = 0 Then GoTo X
                
        On Error Resume Next
        tail = CStr(CLng5(Right(rs_vt!sohieu, dai)) + 1)
        On Error GoTo 0
        
        Do While Len(tail) < dai
            tail = "0" + tail
        Loop
    Else
        GoTo X
    End If
    
    If rs_vt.recordCount > 0 Then
        If Len(rs_vt!sohieu) > dai Then SoHieuVTMoi = Left(rs_vt!sohieu, Len(rs_vt!sohieu) - dai)
        SoHieuVTMoi = SoHieuVTMoi + tail
    Else
X:
        SoHieuVTMoi = "0001"
    End If
    rs_vt.Close
    Set rs_vt = Nothing
End Function
'====================================================================================================
' Thñ tôc in thÎ kho
'====================================================================================================
Public Function InTheKho2(mkho As Long, mvt As Long, tdau As Integer, tcuoi As Integer, thongbao As Boolean, mn As Long, Optional tkdu As String = "", Optional loaitk As Integer = 0, Optional nn As Integer = 0) As Boolean
    Dim dkl As Double, dkt As Double, dkt2 As Double
    Dim sqll As String, st As String, i As Integer, dv As String, ps As Double, Dvt2 As Integer
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqll = sqll + " + Luong_Nhap_" + st + " + Luong_Xuat_" + st + " + Tien_Nhap_" _
        + st + " + Tien_Xuat_" + st
    Next
    
    st = CStr(CThangDB(ThangTruoc(tdau)))
    dkl = SelectSQL("SELECT SUM(Luong_" + st + ") As F1, SUM(Tien_" + st + ") As F2," + IIf(pGiaUSD > 0, "Sum(USDTien_" + st + ")", "0") _
            & " As F3, SUM(" + sqll + ") As F4 FROM TonKho WHERE MaVattu = " + CStr(mvt) + IIf(mkho > 0, " AND  MaSoKho = " + CStr(mkho), ""), dkt, dkt2, ps)
    If dkl = 0 And dkt = 0 And dkt2 = 0 And ps = 0 Then GoTo KhongPS
    
    If ps <> 0 Then
        SetSQL "QTheKho", "SELECT DISTINCTROW ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.MaLoai, ChungTu.SoPS, ChungTu.SoPS2No, ChungTu.SoPS2Co, ChungTu.MaTKTCNo, ChungTu.MaTKTCCo, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.GhiChu, ChungTu.MaKho, ChungTu.MaNguon," + IIf(pGiaUSD > 0, "ChungTu.PSUSD", "0") + " AS PSUSD1, DonVi, TyLeQD, ChungTu.NgayGS" _
            & " FROM (HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) ON HeThongTK_1.MaSo = ChungTu.MaTKCo) LEFT JOIN DVTVattu ON ChungTu.DVT=DVTVattu.MaSo " _
            & " Where (SoPS2No<>0 OR SoPS2Co<>0" + IIf(loaitk = 0, " OR SoPS<>0 ", "") + ") AND (HethongTK.TK_ID=" + CStr(TKVT_ID) + " OR HethongTK_1.TK_ID=" + CStr(TKVT_ID) + ") AND ((ChungTu.MaLoai =1 OR ChungTu.MaLoai =2 OR ChungTu.MaLoai =4) AND " + WThang("ThangCT", tdau, tcuoi) + " And (chungtu.MaVattu = " + CStr(mvt) + ")" + IIf(mkho > 0, " And ((chungtu.MaLoai<4 And chungtu.MaKho = " + CStr(mkho) + ") OR (chungtu.MaLoai=4 And (chungtu.MaKho = " + CStr(mkho) _
            + " OR chungtu.MaNguon = " + CStr(mkho) + ")))", "") + ") " + IIf(mn > 0, "AND ChungTu.MaLoai<>4 AND ChungTu.MaNguon=" + CStr(mn), "") + IIf(Len(tkdu) > 0, "AND (HethongTK.SoHieu LIKE '" + tkdu + "*' OR HethongTK_1.SoHieu LIKE '" + tkdu + "*')", "") + " ORDER BY ChungTu.NgayGS, ChungTu.MaCT"
    Else
        SetSQL "QTheKho", "SELECT DISTINCTROW '' AS SoHieu, Null AS NgayCT, '' AS DienGiai, 0 AS MaLoai, 0 AS SoPS, 0 AS SoPS2No, 0 AS SoPS2Co, 0 AS MaTKTCNo, 0 AS MaTKTCCo, '' AS SoHieu1, '' AS SoHieu2, '' AS GhiChu, 0 AS MaKho, 0 AS MaNguon, 0 AS  PSUSD1,'' AS DonVi, 0 AS TyLeQD, Null AS NgayGS" _
            & " FROM ChungTu Where (chungtu.MaCT = 0)"
    End If
    
    frmMain.Rpt.Formulas(3) = "TenKho='" + IIf(mkho > 0, "Kho: " + tenkho(mkho), ABCtoVNI("Tæng hîp c¸c kho")) + "'"
    frmMain.Rpt.Formulas(4) = "Thang=" + CStr(tdau)
    frmMain.Rpt.Formulas(5) = "ThangCuoi=" + CStr(tcuoi)
    frmMain.Rpt.Formulas(6) = "MaKho=" + CStr(mkho)
    
    sqll = TenVT(st, mvt, dv, Dvt2)
    
    frmMain.Rpt.Formulas(7) = "TenVt = IF PageNumber() = 1 THEN '" + st + " - " + sqll + ABCtoVNI(" - §¬n vÞ tÝnh: ") + dv + "'"
    frmMain.Rpt.Formulas(8) = "DkLuong=" + DoiDau(dkl)
    frmMain.Rpt.Formulas(9) = "DkTien=" + DoiDau(dkt)
    If pGiaUSD > 0 And loaitk = 0 Then frmMain.Rpt.Formulas(11) = "DkTien2=" + DoiDau(dkt2)
    sqll = SelectSQL("SELECT GhiChu AS F1 FROM Vattu WHERE MaSo=" + CStr(mvt))
    If sqll <> "..." Then frmMain.Rpt.Formulas(10) = "GhiChu='" + sqll + "'"
    frmMain.Rpt.ReportFileName = "THEKHO" + IIf(loaitk > 0, "2", "") + IIf(loaitk = 0 And pGiaUSD > 0, "X", "") + IIf(Dvt2 > 0, "B", "") + ".RPT"
    If Dvt2 > 0 Then
        frmMain.Rpt.Formulas(12) = "TyLeQD=" + DoiDau(SelectSQL("SELECT TOP 1 TyLeQD AS F1 FROM DVTVattu WHERE MaVattu=" + CStr(mvt) + " ORDER BY MaSo"))
    End If
    RptSetDate NgayCuoiThang(pNamTC, tcuoi)
    InTheKho2 = True
    Exit Function
KhongPS:
    If thongbao Then MsgBox "VËt t­ kh«ng cã tån kho hoÆc ph¸t sinh !", vbInformation, App.ProductName
    InTheKho2 = False
End Function

Public Function InTheKho3(mkho As Long, tdau As Integer, tcuoi As Integer)
    Dim sqll As String, st As String, i As Integer
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqll = sqll + " + ABS(Luong_Nhap_" + st + ") + ABS(Luong_Xuat_" + st + ") + ABS(Tien_Nhap_" + st + ") + ABS(Tien_Xuat_" + st + ") + ABS(Luong_" + CStr(CThangDB(ThangTruoc(tdau))) + ") + ABS(Tien_" + CStr(CThangDB(ThangTruoc(tdau))) + ")"
    Next
    
    ExecuteSQL5 "DELETE * FROM BKNhomPS"
    'ExecuteSQL5 "INSERT INTO BKNhomPS (MaCT,DienGiai,ShDu,SoPS2No,SoPSNo,ShVT) SELECT DISTINCTROW 0, Vattu.TenVattu, Vattu.DonVi, Luong_" + CStr(CThangDB(ThangTruoc(tdau))) + ", Tien_" + CStr(CThangDB(ThangTruoc(tdau))) + ",Vattu.SoHieu " _
        & " FROM TonKho INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo WHERE TonKho.MaSoKho=" _
        + CStr(mkho) + " AND (" + sqll + ") <>0 ORDER BY Vattu.SoHieu"
    'ExecuteSQL5 "INSERT INTO BKNhomPS (MaCT,DienGiai,ShDu,SoPS2No,SoPSNo,ShVT) SELECT DISTINCTROW 3, Vattu.TenVattu, Vattu.DonVi, Luong_" + CStr(CThangDB(tcuoi)) + ", Tien_" + CStr(CThangDB(tcuoi)) + ",Vattu.SoHieu " _
        & " FROM TonKho INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo WHERE TonKho.MaSoKho=" _
        + CStr(mkho) + " AND (" + sqll + ") <>0 ORDER BY Vattu.SoHieu"
    
    st = "SELECT DISTINCTROW 1 AS Loai, ChungTu.SoHieu, NgayCT, ChungTu.DienGiai,TK.Sohieu AS SHDU,SoPS2No,SoPS AS SoPSNo,0 AS SoPS2Co,0 AS SoPSCo,NgayGS,Vattu.SoHieu AS SHVT,MaCT " _
        & " FROM ((HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo " _
        & " WHERE (SoPS2No<>0 OR SoPS2Co<>0) AND (HethongTK.TK_ID=" + CStr(TKVT_ID) + ") AND ((ChungTu.MaLoai =1) OR (ChungTu.MaLoai =4 AND ChungTu.MaNguon=" + CStr(mkho) + ")) AND " + WThang("ThangCT", tdau, tcuoi) _
        & " UNION SELECT DISTINCTROW 2, ChungTu.SoHieu, NgayCT, ChungTu.DienGiai,TK.Sohieu,0,0,SoPS2Co,SoPS,NgayGS,Vattu.SoHieu,MaCT " _
        & " FROM ((HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKCo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKNo=TK.MaSo " _
        & " WHERE (SoPS2No<>0 OR SoPS2Co<>0) AND (HethongTK.TK_ID=" + CStr(TKVT_ID) + ") AND (ChungTu.MaLoai =2 OR ChungTu.MaLoai =4) AND (ChungTu.MaKho=" + CStr(mkho) + ") AND " + WThang("ThangCT", tdau, tcuoi) _
        & " UNION SELECT DISTINCTROW 0, Vattu.SoHieu, Null AS NgayCT, Vattu.TenVattu, Vattu.DonVi, Luong_" + CStr(CThangDB(ThangTruoc(tdau))) + ", Tien_" + CStr(CThangDB(ThangTruoc(tdau))) + ",0,0, CVDate('1/1/1980') AS NgayGS,Vattu.SoHieu,0 " _
        & " FROM TonKho INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo WHERE TonKho.MaSoKho=" _
        + CStr(mkho) + " AND (" + sqll + ") <>0" _
        & " UNION SELECT DISTINCTROW 3, Vattu.SoHieu, Null AS NgayCT, Vattu.TenVattu, Vattu.DonVi, Luong_" + CStr(CThangDB(tcuoi)) + ", Tien_" + CStr(CThangDB(tcuoi)) + ",0,0, CVDate('1/1/2030') AS NgayGS,Vattu.SoHieu,0 " _
        & " FROM TonKho INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo WHERE TonKho.MaSoKho=" _
        + CStr(mkho) + " AND (" + sqll + ") <>0"
    SetSQL "MienTru", st
                
    ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,SoHieu,Ngay,DienGiai,ShDu,SoPS2No,SoPSNo,SoPS2Co,SoPSCo,NgayGS,ShVT) SELECT MaSo,Loai, SoHieu,NgayCT,DienGiai,ShDU,SoPS2No,SoPSNo,SoPS2Co,SoPSCo,NgayGS,SHVT " _
        & " FROM MienTru ORDER BY ShVT, NgayGS, Loai, MaCT"
                
    frmMain.Rpt.ReportFileName = "THEKHO3.RPT"
    frmMain.Rpt.Formulas(3) = "TenKho='" + IIf(mkho > 0, "Kho: " + tenkho(mkho), ABCtoVNI("Tæng hîp c¸c kho")) + "'"
    frmMain.Rpt.Formulas(4) = "Thang=" + CStr(tdau)
    frmMain.Rpt.Formulas(5) = "ThangCuoi=" + CStr(tcuoi)
        
    RptSetDate NgayCuoiThang(pNamTC, tcuoi)
End Function

Public Function tenkho(mkho As Long) As String
    Dim m As Long
    
    tenkho = SelectSQL("SELECT DISTINCTROW TenKho AS F1, MaSo AS F2 FROM KhoHang WHERE MaSo = " + CStr(mkho), m)
    If m = 0 Then tenkho = ABCtoVNI("Tæng hîp c¸c kho")
End Function

'======================================================================================
' Hµm tr¶ vÒ m· sè, tªn ph©n lo¹i tõ sè hiÖu
'======================================================================================
Public Function TenPLVT(sh As String, mpl As Long) As String
    Dim rs_tk As Recordset
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo,TenPhanLoai FROM PhanLoaiVattu WHERE " + IIf(mpl > 0, "MaSo=" + CStr(mpl), "SoHieu='" + sh + "'"), dbOpenSnapshot)
    If rs_tk.recordCount > 0 Then
        mpl = rs_tk!MaSo
        TenPLVT = rs_tk!TenPhanLoai
    Else
        mpl = 0
        TenPLVT = ""
    End If
    
    rs_tk.Close
    Set rs_tk = Nothing
End Function

Public Function Ten154(sh As String, mpl As Long) As String
    Dim rs_tk As Recordset
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo,TenVattu FROM TP154 WHERE SoHieu = '" + sh + "'", dbOpenSnapshot)
    If rs_tk.recordCount > 0 Then
        mpl = rs_tk!MaSo
        Ten154 = rs_tk!TenVattu
    Else
        mpl = 0
        Ten154 = ""
    End If
    
    rs_tk.Close
    Set rs_tk = Nothing
End Function
'======================================================================================
' Hµm tr¶ vÒ m· sè, tªn VT tõ sè hiÖu
'======================================================================================
Public Function TenVT(sh As String, mvt As Long, dv As String, Optional Dvt2 As Integer) As String
    TenVT = SelectSQL("SELECT SoHieu AS F2, TenVattu AS F1, DonVi AS F3,Dvt2 AS F4 FROM Vattu WHERE " + IIf(mvt > 0, "MaSo=" + CStr(mvt), "SoHieu='" + sh + "'"), sh, dv, Dvt2)
    If mvt = 0 Then mvt = SoHieu2MaSo(sh, "Vattu")
End Function

Public Function TenTP(sh As String, mvt As Long) As String
    TenTP = SelectSQL("SELECT SoHieu AS F2, TenVattu AS F1 FROM TP154 WHERE " + IIf(mvt > 0, "MaSo=" + CStr(mvt), "SoHieu='" + sh + "'"), sh)
    If mvt = 0 Then mvt = SoHieu2MaSo(sh, "TP154")
End Function

Public Function TenDVT(mdv As Long, Optional TyLeQD As Double) As String
    TenDVT = SelectSQL("SELECT DonVi AS F1, TyLeQD AS F2 FROM DVTVattu WHERE MaSo=" + CStr(mdv))
End Function

Public Function TenNguon(ms As Long) As String
    TenNguon = SelectSQL("SELECT SoHieu+' - '+DienGiai AS F1 FROM NguonNhapXuat WHERE MaSo=" + CStr(ms))
End Function


'======================================================================================
' Thñ tôc tÝnh sè d­ ®Çu kú cña c¸c tµi kho¶n vËt t­
'======================================================================================
Public Sub SoDuTKVT()
    Dim rs_tk As Recordset, taikhoan As New ClsTaikhoan, sql As String, m As Long
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaTaiKhoan,Sum(Tien_0) AS DuDk" _
        & " FROM TonKho GROUP BY MaTaiKhoan", dbOpenSnapshot, dbForwardOnly)
    
    Do While Not rs_tk.EOF
        taikhoan.InitTaikhoanMaSo rs_tk!MaTaiKhoan
        If rs_tk!DuDk > 0 Then
            taikhoan.NoDauKy = rs_tk!DuDk
            taikhoan.CoDauKy = 0
        Else
            taikhoan.CoDauKy = rs_tk!DuDk
            taikhoan.NoDauKy = 0
        End If
        taikhoan.CapNhatTk
        rs_tk.MoveNext
    Loop
    rs_tk.Close
    
    If OutCost <> 0 Then
        sql = "SELECT ChungTu.MaSo,MaKho,MaTKCo,ChungTu.MaVattu,SoPS2Co FROM (" + ChungTu2TKNC(1) + ") LEFT JOIN VTDauNam ON ABS(ChungTu.CT_ID)-2000000000=VTDauNam.MaSo WHERE (MaLoai=2 OR MaLoai=4) AND IsNull(VTDauNam.MaSo) AND ABS(ChungTu.CT_ID)>2000000000  AND TK_ID=" + CStr(TKVT_ID)
        Set rs_tk = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rs_tk.EOF
            m = SelectSQL("SELECT MaSo AS F1 FROM VTDauNam WHERE MaSoKho=" + CStr(rs_tk!MaKho) + " AND MaTaiKhoan=" + CStr(rs_tk!MaTkCo) + " AND MaVattu=" + CStr(rs_tk!MaVattu))
            If m > 0 Then
                ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=SoXuat+" + CStr(rs_tk!SoPS2Co) + " WHERE MaSo=" + CStr(m)
                ExecuteSQL5 "UPDATE ChungTu SET CT_ID=" + CStr(-(2000000000 - m)) + " WHERE MaSo=" + CStr(rs_tk!MaSo)
            End If
            rs_tk.MoveNext
         Loop
     End If
     
    Set rs_tk = Nothing
    Set taikhoan = Nothing
End Sub

Public Function GetVAT(mvt As Long) As Integer
    Dim sql As String
    
    sql = "SELECT VAT AS F1 FROM PhanLoaiVattu INNER JOIN Vattu ON PhanLoaiVattu.MaSo=Vattu.MaPhanLoai WHERE Vattu.MaSo=" + CStr(mvt)
    GetVAT = SelectSQL(sql)
End Function

Public Function InTheKho2N(mkho As Long, mvt As Long, ndau As Date, ncuoi As Date, thongbao As Boolean, mn As Long, Optional tkdu As String = "", Optional loaitk As Integer = 0, Optional nn As Integer = 0) As Boolean
    Dim rs_ps As Recordset, dkl As Double, dkt As Double, dkt2 As Double
    Dim sqll As String, st As String, i As Integer, dv As String, Dvt2 As Integer
    
    dkl = SoTonKhoN(ndau - 1, mkho, 0, mvt, dkt, dkt2)

    SetSQL "QTheKho", "SELECT DISTINCTROW ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.MaLoai, ChungTu.SoPS, ChungTu.SoPS2No, ChungTu.SoPS2Co, ChungTu.MaTKTCNo, ChungTu.MaTKTCCo, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.GhiChu, ChungTu.MaKho, ChungTu.MaNguon," + IIf(pGiaUSD > 0, "PSUSD", "0") + " AS PSUSD, DonVi, TyLeQD, ChungTu.NgayGS" _
        & " FROM (HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) ON HeThongTK_1.MaSo = ChungTu.MaTKCo) LEFT JOIN DVTVattu ON ChungTu.DVT=DVTVattu.MaSo " _
        & " Where (SoPS2No<>0 OR SoPS2Co<>0" + IIf(loaitk = 0, " OR SoPS<>0 ", "") + ") AND (HethongTK.TK_ID=" + CStr(TKVT_ID) + " OR HethongTK_1.TK_ID=" + CStr(TKVT_ID) + ") AND ((ChungTu.MaLoai =1 OR ChungTu.MaLoai =2 OR ChungTu.MaLoai =4) AND " + WNgay("NgayGS", ndau, ncuoi) + " AND (chungtu.MaVattu = " + CStr(mvt) + ")" + IIf(mkho > 0, " And ((chungtu.MaLoai<4 And chungtu.MaKho = " + CStr(mkho) + ") OR (chungtu.MaLoai=4 And (chungtu.MaKho = " + CStr(mkho) _
        + " OR chungtu.MaNguon = " + CStr(mkho) + ")))", "") + ") " + IIf(mn > 0, "AND ChungTu.MaLoai<>4 AND ChungTu.MaNguon=" + CStr(mn), "") + IIf(Len(tkdu) > 0, "AND (HethongTK.SoHieu LIKE '" + tkdu + "*' OR HethongTK_1.SoHieu LIKE '" + tkdu + "*')", "") + " ORDER BY ChungTu.NgayGS, ChungTu.MaCT"
    If SelectSQL("SELECT Count(MaKho) AS F1 FROM QTheKho") = 0 Then
        SetSQL "QTheKho", "SELECT DISTINCTROW '' AS SoHieu, Null AS NgayCT, '' AS DienGiai, 0 AS MaLoai, 0 AS SoPS, 0 AS SoPS2No, 0 AS SoPS2Co, 0 AS MaTKTCNo, 0 AS MaTKTCCo, '' AS SoHieu1, '' AS SoHieu2, '' AS GhiChu, 0 AS MaKho, 0 AS MaNguon,0 AS  PSUSD,'' AS DonVi, 0 AS TyLeQD, Null AS NgayGS" _
        & " FROM ChungTu Where (chungtu.MaCT = 0)"
    End If
    
    frmMain.Rpt.Formulas(3) = "TenKho='" + IIf(mkho > 0, "Kho: " + tenkho(mkho), ABCtoVNI("Tæng hîp c¸c kho")) + "'"
    frmMain.Rpt.Formulas(5) = "ThoiGian='" + ThoiGianN(ndau, ncuoi, nn) + "'"
    frmMain.Rpt.Formulas(6) = "MaKho=" + CStr(mkho)
    
    sqll = TenVT(st, mvt, dv, Dvt2)
    
    frmMain.Rpt.Formulas(7) = "TenVt = IF PageNumber() = 1 THEN '" + st + " - " + sqll + ABCtoVNI(" - §¬n vÞ tÝnh: ") + dv + "'"
    frmMain.Rpt.Formulas(8) = "DkLuong=" + DoiDau(dkl)
    frmMain.Rpt.Formulas(9) = "DkTien=" + DoiDau(dkt)
    If pGiaUSD > 0 Then frmMain.Rpt.Formulas(11) = "DkTien2=" + DoiDau(dkt2)
    sqll = SelectSQL("SELECT GhiChu AS F1 FROM Vattu WHERE MaSo=" + CStr(mvt))
    If sqll <> "..." Then frmMain.Rpt.Formulas(10) = "GhiChu='" + sqll + "'"
    frmMain.Rpt.ReportFileName = "THEKHO" + IIf(loaitk > 0, "2", "") + IIf(loaitk = 0 And pGiaUSD > 0, "X", "") + IIf(Dvt2 > 0, "B", "") + ".RPT"
    If Dvt2 > 0 Then
        frmMain.Rpt.Formulas(12) = "TyLeQD=" + DoiDau(SelectSQL("SELECT TOP 1 TyLeQD AS F1 FROM DVTVattu WHERE MaVattu=" + CStr(mvt) + " ORDER BY MaSo"))
    End If
    RptSetDate ncuoi
    InTheKho2N = True
    Exit Function
KhongPS:
    If thongbao Then MsgBox "VËt t­ kh«ng cã tån kho hoÆc ph¸t sinh !", vbInformation, App.ProductName
    InTheKho2N = False
End Function

Public Sub NhapTheoNguon(tdau As Integer, tcuoi As Integer, mvt As Long, mtk As Long, tkdu As String, l1 As Double, T As Double, Optional mkho As Long = 0)
    T = SelectSQL("SELECT SUM(SoPS) AS F1,SUM(SoPS2No) AS F2 FROM " + ChungTu2TKNC(1) _
        & " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND ChungTu.MaLoai=1 AND MaTKNo=" + CStr(mtk) + " AND MaVattu=" + CStr(mvt) + " AND HethongTK.SoHieu LIKE '" + tkdu + "*'" + IIf(mkho > 0, " AND MaKho=" + CStr(mkho), ""), l1)
End Sub

Public Sub XuatTheoNguon(tdau As Integer, tcuoi As Integer, mvt As Long, mtk As Long, tkdu As String, l1 As Double, T As Double, Optional mkho As Long = 0)
    T = SelectSQL("SELECT SUM(SoPS) AS F1,SUM(SoPS2Co) AS F2 FROM " + ChungTu2TKNC(-1) _
        & " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND ChungTu.MaLoai=2 AND MaTKCo=" + CStr(mtk) + " AND MaVattu=" + CStr(mvt) + " AND HethongTK.SoHieu LIKE '" + tkdu + "*'" + IIf(mkho > 0, " AND MaKho=" + CStr(mkho), ""), l1)
End Sub

Public Function GiaXuatKho(mk As Long, mtk As Long, mvt As Long, ngay As Date, SoLuong As Double, Optional tien2 As Double) As Double
    Dim luong As Double, tien As Double, tien1 As Double, t2 As Double
    
    tien2 = 0
    If OutCost = 0 Then
        luong = SoTonKho(ThangTruoc(Month(ngay)), mk, mtk, mvt, tien, tien2)
        luong = luong + SoNhapKho(mk, mtk, mvt, NgayDauThang(pNamTC, Month(ngay)), NgayCuoiThang(pNamTC, Month(ngay)), tien1, t2)
        tien = tien + tien1
        tien2 = tien2 + t2
        If luong > 0 Then
            If luong = SoLuong Then
                GiaXuatKho = tien
            Else
                tien2 = Fix(0.5 + Mask_N * tien2 * SoLuong / luong) / Mask_N
                GiaXuatKho = RoundMoney(tien * SoLuong / luong)
            End If
        End If
    End If
    
End Function

Public Function GiaXuatKhoBQCK(mk As Long, mtk As Long, mvt As Long, thang As Integer, SoLuong As Double, Optional tien2 As Double) As Double
    Dim luong As Double, tien As Double, tien1 As Double, t2 As Double
    
    tien2 = 0
    If OutCost = 0 Then
        luong = SoTonKho(ThangTruoc(thang), mk, mtk, mvt, tien, tien2)
        luong = luong + SoNhapKhoThang(mk, mtk, mvt, thang, thang, tien1, t2)
        tien = tien + tien1
        tien2 = tien2 + t2
        If luong <> 0 Then
            If luong = SoLuong Then
                GiaXuatKhoBQCK = tien
            Else
                tien2 = Fix(0.5 + Mask_N * tien2 * SoLuong / luong) / Mask_N
                GiaXuatKhoBQCK = RoundMoney(tien * SoLuong / luong)
            End If
        End If
    End If
    
End Function

Public Sub GhiXuatNVL(mct As Long, n As Date, thang As Integer, xk As Integer, tp As Cls154, Optional ktra As Integer = 0, Optional tygia As Double = 1)
    Dim rs As Recordset, tien As Double, mtk As Long, tongtien As Double, rs2 As Recordset, st As String, n0 As Date
    Dim ct As New ClsChungtu, MaCT As Long, dgia As Double, luong As Double, thangdm As Integer, sh As String, i As Integer, tcp As Double
        
    n0 = NgayDauThang(pNamTC, thang)
    
    If xk <> 0 And ktra = 0 Then
        XoaCTTheoID thang, 610000000 + mct, 0, n0, n
        ExecuteSQL5 "UPDATE ThanhPham SET CPNVL=0 WHERE MaCT=" + CStr(mct)
        sh = ""
        SetSQL "MienTru", "SELECT DISTINCTROW MaSo,SoHieu,Ngay,MaKhoNVL,MaTKCP,MaTKNC,MaTKSX,MaTP,SoLuong FROM ThanhPham WHERE MaCT=" + CStr(mct)
        Set rs = DBKetoan.OpenRecordset("MienTru", dbOpenSnapshot, dbForwardOnly)
        Do While Not rs.EOF
            thangdm = SelectSQL("SELECT  TOP 1 Thang AS F1 FROM DinhMuc WHERE MaNVL>0 AND MaTP=" + CStr(rs!MaTP) + " AND " + WThang("Thang", 0, thang) + " ORDER BY " + SetMonthOrder("Thang") + " DESC")
            If thangdm > 0 Then
                Set rs2 = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaNVL,DinhMuc.Soluong*MienTru.SoLuong AS Luong FROM MienTru INNER JOIN DinhMuc ON MienTru.MaTP=DinhMuc.MaTP WHERE DinhMuc.Thang=" + CStr(thangdm) + " AND MaNVL>0 AND DinhMuc.SoLuong<>0 AND MienTru.MaSo=" + CStr(rs!MaSo), dbOpenSnapshot, dbForwardOnly)
                Do While Not rs2.EOF
                    mtk = SelectSQL("SELECT MaTaiKhoan AS F1 FROM TonKho WHERE MaSoKho=" + CStr(rs!MaKhoNVL) + " AND MaVattu=" + CStr(rs2!MaNVL))
                    tien = GiaXuatKho(rs!MaKhoNVL, mtk, rs2!MaNVL, rs!ngay, rs2!luong)
                    If tien <> 0 Then
                        If rs2!luong <> 0 Then dgia = tien / rs2!luong Else dgia = 0
                        tongtien = tongtien + tien
                        ct.InitChungtu 0, 2, "CPNVLTT" + tp.sohieu + rs!sohieu, thang, rs!ngay, rs!ngay, 0, rs!MaKhoNVL, ABCtoVNI("XuÊt nguyªn vËt liÖu cho s¶n xuÊt"), rs!MaTKCP, mtk, tien, _
                            0, rs2!luong, rs2!MaNVL, "Taäp hôïp töï ñoäng", 0, "", "", "", ""
                        ct.CT_ID = 610000000 + mct
                        ct.MaTP = tp.MaSo
                        If sh <> rs!sohieu Then
                            sh = rs!sohieu
                            MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
                        End If
                        ct.MaCT = MaCT
                        If pTygia > 0 Then ct.tygia = tygia
                        ct.GhiChungtuTH 0, 1
                        ExecuteSQL5 "UPDATE ThanhPham INNER JOIN DinhMuc ON ThanhPham.MaTP=DinhMuc.MaTP SET CPNVL=CPNVL+ThanhPham.SoLuong*DinhMuc.SoLuong*" + DoiDau(dgia) + " WHERE ThanhPham.MaSo=" + CStr(rs!MaSo) + " AND DinhMuc.Thang=" + CStr(thangdm) + " AND MaNVL=" + CStr(rs2!MaNVL)
                    End If
                    rs2.MoveNext
                Loop
                rs2.Close
            End If
            rs.MoveNext
        Loop
    Else
        XoaCTTheoID thang, 610000000 + mct, 0, n0, n, 2
    End If
    
    sh = ""
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW SoHieu,Ngay,MaTKCP,MaTKSX," + RoundMoneySQL("Sum(CPNVL)") + " AS Tien FROM ThanhPham WHERE CPNVL<>0 AND MaCT=" + CStr(mct) + " GROUP BY SoHieu,Ngay,MaTKCP,MaTKSX", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        ct.InitChungtu 0, 3, "KCCPNVL" + tp.sohieu + rs!sohieu, thang, rs!ngay, rs!ngay, 0, 0, ABCtoVNI("KÕt chuyÓn chi phÝ nguyªn vËt liÖu trùc tiÕp"), rs!MaTKSX, rs!MaTKCP, rs!tien, 0, 0, 0, "KC", 0, "", "", "", ""
        ct.MaTP = tp.MaSo
        If sh <> rs!sohieu Then
            sh = rs!sohieu
            MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
        End If
        ct.MaCT = MaCT
        ct.CT_ID = 610000000 + mct
        If pTygia > 0 Then ct.tygia = tygia
        ct.GhiChungtuTH 0, 1
        rs.MoveNext
    Loop
    
    sh = ""
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW SoHieu,Ngay,MaTKNC,MaTKSX,MaTP,Sum(SoLuong) AS SL FROM ThanhPham WHERE MaCT=" + CStr(mct) + " GROUP BY SoHieu,Ngay,MaTKNC,MaTKSX,MaTP", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        tien = RoundMoney(rs!sl * SelectSQL("SELECT  TOP 1 SoLuong AS F1 FROM DinhMuc WHERE MaNVL=0 AND MaTP=" + CStr(rs!MaTP) + " AND " + WThang("Thang", 0, thang) + " ORDER BY " + SetMonthOrder("Thang") + " DESC"))
        If tien <> 0 Then
            ct.InitChungtu 0, 3, "KCCPNC" + CStr(thang) + rs!sohieu, thang, rs!ngay, rs!ngay, 0, 0, ABCtoVNI("KÕt chuyÓn chi phÝ nh©n c«ng trùc tiÕp"), rs!MaTKSX, rs!MaTKNC, tien, 0, 0, 0, "KC", 0, "", "", "", ""
            If sh <> rs!sohieu Then
                sh = rs!sohieu
                MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
            End If
            ct.MaCT = MaCT
            ct.CT_ID = 610000000 + mct
            If pTygia > 0 Then ct.tygia = tygia
            ct.GhiChungtuTH 0, 1
        End If
        rs.MoveNext
    Loop
    
    st = CStr(CThangDB(thang))
    MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
    mtk = SelectSQL("SELECT Top 1 MaSo AS F1 FROM HethongTK WHERE SoHieu LIKE '334*' AND TKCon=0")
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo,DuCo_" + st + " AS NC FROM HethongTK WHERE DuCo_" + st + "<>0 AND TKCon=0 AND SoHieu LIKE '622*'", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        ct.InitChungtu 0, 0, "CPNC" + CStr(thang), thang, n, n, 0, 0, ABCtoVNI("Chi phÝ nh©n c«ng trùc tiÕp"), rs!MaSo, mtk, rs!nC, 0, 0, 0, "KC töï ñoäng", 0, "", "", "", ""
        ct.MaCT = MaCT
        ct.CT_ID = 610000000 + mct
        If pTygia > 0 Then ct.tygia = tygia
        ct.GhiChungtuTH 0, 1
        rs.MoveNext
    Loop
        
    MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo,DuNo_" + st + "-DuCo_" + st + " AS Tien FROM HethongTK WHERE TKCon=0 AND SoHieu LIKE '627*' AND DuNo_" + st + "-DuCo_" + st + ">0", dbOpenSnapshot, dbForwardOnly)
    Set rs2 = DBKetoan.OpenRecordset("SELECT MaTKSX FROM ThanhPham WHERE MaCT=" + CStr(mct) + " GROUP BY MaTKSX", dbOpenSnapshot, dbForwardOnly)
    Do While (Not rs.EOF) And (Not rs2.EOF)
        ct.InitChungtu 0, 3, "KCCPSXC" + CStr(thang), thang, n, n, 0, 0, ABCtoVNI("KÕt chuyÓn chi phÝ s¶n xuÊt chung"), rs2!MaTKSX, rs!MaSo, rs!tien, 0, 0, 0, "KC töï ñoäng", 0, "", "", "", ""
        ct.MaCT = MaCT
        ct.CT_ID = 610000000 + mct
        If pTygia > 0 Then ct.tygia = tygia
        ct.GhiChungtuTH 0, 1
        rs.MoveNext
    Loop
    
    If pDTTP = 0 Then
        tcp = PSTKCP("627", thang, thang)
        If SelectSQL("SELECT Fix(0.5+Sum(CPSXC)) AS F1 FROM ThanhPham WHERE Thang=" + CStr(thang)) <> tcp Then
            tongtien = SelectSQL("SELECT Fix(0.5+Sum(CPNC)) AS F1 FROM ThanhPham WHERE Thang=" + CStr(thang))
            If tongtien > 0 Then
                thangdm = SelectSQL("SELECT Count(MaSo) AS F1 FROM ThanhPham WHERE Thang=" + CStr(thang))
                ExecuteSQL5 "UPDATE ThanhPham SET CPSXC=CPNC/" + DoiDau(tongtien) + " WHERE Thang=" + CStr(thang)
                Set rs = DBKetoan.OpenRecordset("SELECT * FROM ThanhPham WHERE Thang=" + CStr(thang) + " ORDER BY CPSXC", dbOpenSnapshot)
                i = 0
                tongtien = 0
                Do While Not rs.EOF
                    i = i + 1
                    If i < thangdm Then
                        tien = Fix(0.5 + tcp * rs!cpsxc)
                        tongtien = tongtien + tien
                    Else
                        tien = tcp - tongtien
                    End If
                    ExecuteSQL5 "UPDATE ThanhPham SET CPSXC=" + CStr(tien) + " WHERE MaSo=" + CStr(rs!MaSo)
                    rs.MoveNext
                Loop
            End If
        End If
    End If
    
    MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo,DuNo_" + st + "-DuCo_" + st + " AS Tien FROM HethongTK WHERE TKCon=0 AND SoHieu LIKE '622*' AND DuNo_" + st + "-DuCo_" + st + ">0", dbOpenSnapshot, dbForwardOnly)
    Set rs2 = DBKetoan.OpenRecordset("SELECT MaTKSX FROM ThanhPham WHERE MaCT=" + CStr(mct) + " GROUP BY MaTKSX", dbOpenSnapshot, dbForwardOnly)
    Do While (Not rs.EOF) And (Not rs2.EOF)
        ct.InitChungtu 0, 3, "KCCPNC" + CStr(thang), thang, n, n, 0, 0, ABCtoVNI("KÕt chuyÓn chi phÝ nh©n c«ng"), rs2!MaTKSX, rs!MaSo, rs!tien, 0, 0, 0, "KC töï ñoäng", 0, "", "", "", ""
        ct.MaCT = MaCT
        ct.CT_ID = 610000000 + mct
        If pTygia > 0 Then ct.tygia = tygia
        ct.GhiChungtuTH 0, 1
        rs.MoveNext
    Loop
    
    PBCPKhac "641", thang
    PBCPKhac "642", thang
    
    If ktra = 0 And pGiaHT = 0 Then
        sh = ""
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW SoHieu,Ngay,MaTK,MaTP,MaKhoTP,MaTKSX,Sum(SoLuong) AS Luong," + RoundMoneySQL("Sum(CPNVL+CPNC+CPSXC)") + " AS Tien FROM ThanhPham WHERE MaCT=" + CStr(mct) + " GROUP BY SoHieu,Ngay,MaTK,MaTP,MaKhoTP,MaTKSX", dbOpenSnapshot, dbForwardOnly)
        Do While Not rs.EOF
            If sh <> rs!sohieu Then
                sh = rs!sohieu
                MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
            End If
            ct.InitChungtu 0, 1, "KCTP-" + CStr(thang) + CStr(tp.MaSo) + sh, thang, rs!ngay, rs!ngay, 0, rs!MaKhoTP, ABCtoVNI("KÕt chuyÓn thµnh phÈm hoµn thµnh"), rs!MaTK, rs!MaTKSX, rs!tien, rs!luong, 0, rs!MaTP, "KC töï ñoäng", 0, "", "", "", ""
            ct.MaCT = MaCT
            ct.MaTP = tp.MaSo
            ct.CT_ID = 610000000 + mct
            If pTygia > 0 Then ct.tygia = tygia
            ct.GhiChungtuTH 0, 1
            rs.MoveNext
        Loop
        
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaTP,Vattu.SoHieu FROM ThanhPham INNER JOIN Vattu ON ThanhPham.MaTP=Vattu.MaSo WHERE MaCT=" + CStr(mct) + " GROUP BY MaTP,Vattu.SoHieu", dbOpenSnapshot, dbForwardOnly)
        Do While Not rs.EOF
            'TinhGVBH n0, NgayCuoiThang(pNamTC, thang), vbYes, rs!MaTP, 1
            TinhGXK thang, thang, rs!sohieu, "621", 1
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    On Error Resume Next
    rs2.Close
    Set rs2 = Nothing
    On Error GoTo 0
    Set ct = Nothing
    If ktra = 0 Then
        KiemTraVatTu
        KiemTraTaiKhoan
    End If
End Sub

Private Sub PBCPKhac(shtk As String, thang As Integer)
    Dim sopb As Double, luongtp As Double, dgia As Double, st As String
    
    st = CStr(CThangDB(thang))
    sopb = SelectSQL("SELECT Sum(DuNo_" + st + "-DuCo_" + st + ") AS F1 FROM HethongTK WHERE TKCon=0 AND SoHieu LIKE '" + shtk + "*'")
    luongtp = SelectSQL("SELECT Sum(SoLuong) AS F1 FROM ThanhPham WHERE Thang=" + CStr(thang))
    If luongtp <> 0 Then dgia = sopb / luongtp
    Select Case shtk
        Case "641":
            ExecuteSQL5 "UPDATE ThanhPham SET CPBH=SoLuong*" + DoiDau(dgia)
        Case "642":
            ExecuteSQL5 "UPDATE ThanhPham SET CPQL=SoLuong*" + DoiDau(dgia)
    End Select
End Sub

Public Sub DanhDiemVT(mpl As Long)
    Dim sql As String
        
    sql = "SELECT DISTINCTROW PhanLoaiVattu.MaSo AS MPL3,PLCha AS PLCha3,PhanLoaiVattu.SoHieu AS SHPL3, PhanLoaiVattu.TenPhanLoai AS TenPL3, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, GiaBan1, GiaBan2, GiaBan3, DVTVattu.DonVi AS DVQD, DVTVattu.TyLeQD, DVTVattu.GiaBan" _
        & " FROM (PhanLoaiVattu INNER JOIN Vattu ON PhanLoaiVattu.MaSo = Vattu.MaPhanLoai) LEFT JOIN DVTVattu ON Vattu.MaSo=DVTVattu.MaVattu ORDER BY PhanLoaiVattu.SoHieu, Vattu.SoHieu, DVTVattu.DonVi"
    SetSQL "QChitiet", sql
    sql = "SELECT QChitiet.*,PhanLoaiVattu.MaSo AS MPL2,PhanLoaiVattu.SoHieu AS SHPL2,TenPhanLoai AS TenPL2,PhanLoaiVattu.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoaiVattu ON QChitiet.PLCha3=PhanLoaiVattu.MaSo "
    SetSQL "QDuPhong", sql
    sql = "SELECT QDuPhong.*,PhanLoaiVattu.SoHieu AS SHPL1,TenPhanLoai AS TenPL1 FROM QDuPhong LEFT JOIN PhanLoaiVattu ON QDuPhong.PLCha2=PhanLoaiVattu.MaSo " _
        + IIf(mpl > 0, "WHERE PhanLoaiVattu.MaSo=" + CStr(mpl) + " OR MPL3=" + CStr(mpl) + " OR MPL2=" + CStr(mpl), "") + " ORDER BY QDuPhong.Sohieu"
    SetSQL "QLuyKe", sql
        
    frmMain.Rpt.ReportFileName = "DDVT.RPT"
    frmMain.Rpt.WindowTitle = "B¶ng danh ®iÓm vËt t­, hµng ho¸"
End Sub

Public Function KtraCTNhap(mct As Long) As Boolean
    Dim rs As Recordset, T As Double, t2 As Double
    
    KtraCTNhap = True
    Set rs = DBKetoan.OpenRecordset("SELECT ThangCT,MaKho,MaTKNo,MaVattu FROM ChungTu WHERE MaLoai=1 AND MaVattu>0 AND MaCT=" + CStr(mct), dbOpenSnapshot)
    Do While Not rs.EOF
        If SoTonKho(rs!ThangCT, rs!MaKho, rs!MaTkNo, rs!MaVattu, T, t2) <= 0 Then
            KtraCTNhap = False
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Function

Public Sub TinhGXK(tdau As Integer, tcuoi As Integer, shvt As String, tkno As String, Optional ktra As Integer = 0)
    Dim rs As Recordset, ms As Long, tien As Double, luong As Double, sql As String
    Dim mk As Long, mv As Long, mt As Long, n As Date, tienx As Double, luongx As Double, tien2 As Double, tienx2 As Double
    
    ExecuteSQL5 "UPDATE " + ChungTu2TKNC(0) + " SET MaTKNo=MaTKCo,MaTKTCNo=MaTKTCCo WHERE MaLoai=4 AND HethongTK.Cap=0 AND HethongTK.Loai=0 AND TK.Loai>0"
    ExecuteSQL5 "UPDATE ChungTu SET SoPS=Fix(IIF(SoPS>=0,0.5,-0.5)+SoPS), SoPS2Co=Fix(IIF(SoPS2Co>=0,0.5,-0.5)+SoPS2Co*" + CStr(Mask_N) + ")/" + CStr(Mask_N)
    
    If OutCost > 0 Then Exit Sub
        
    If Len(tkno) > 0 Then
        sql = "SELECT DISTINCTROW ChungTu.MaSo,NgayGS,MaCT,MaKho,MaVattu,MaTKCo,SoPS,SoPS2Co" + IIf(pGiaUSD > 0, ",PSUSD", "") + " FROM (" + ChungTu2TKNC(0) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE HethongTK.SoHieu LIKE '" + tkno + "*' AND (MaLoai=2 OR MaLoai=4) AND MaVattu>0 AND MaTKNo>0 AND TK.TK_ID=" + CStr(TKVT_ID) + " AND SoPS2Co>0 AND " + WThang("ThangCT", tdau, tcuoi) + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " ORDER BY MaKho,MaTKCo,MaVattu,NgayGS,ChungTu.MaCT"
    Else
        sql = "SELECT DISTINCTROW ChungTu.MaSo,NgayGS,MaCT,MaKho,MaVattu,MaTKCo,SoPS,SoPS2Co" + IIf(pGiaUSD > 0, ",PSUSD", "") + " FROM (" + ChungTu2TKNC(1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE (MaLoai=2 OR MaLoai=4) AND MaVattu>0 AND HethongTK.TK_ID=" + CStr(TKVT_ID) + " AND SoPS2Co>0 AND " + WThang("ThangCT", tdau, tcuoi) + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " ORDER BY MaKho,MaTKCo,MaVattu,NgayGS,ChungTu.MaCT"
    End If
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        If mk <> rs!MaKho Or mt <> rs!MaTkCo Or mv <> rs!MaVattu Then
            n = rs!NgayGS
            mk = rs!MaKho
            mt = rs!MaTkCo
            mv = rs!MaVattu
            luong = SoTonKhoN2(n, mk, mt, mv, tien, rs!MaCT, tien2)
        Else
            If n <> rs!NgayGS Then
                If Len(tkno) > 0 Then
                    luong = SoTonKhoN2(rs!NgayGS, mk, mt, mv, tien, rs!MaCT, tien2)
                Else
                    luong = luong + SoNhapKho(mk, mt, mv, n + 1, rs!NgayGS, tienx, tienx2)
                    tien = tien + tienx
                    tien2 = tien2 + tienx2
                End If
                n = rs!NgayGS
            End If
        End If
        
        If tien = 0 Then
            tienx = SelectSQL("SELECT TOP 1 SoPS AS F1,SoPS2Co AS F2 FROM ChungTu WHERE (MaLoai=2 OR MaLoai=4) AND MaKho=" + CStr(mk) + " AND MaTKCo=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND NgayGS<=#" + Format(n, Mask_DB) + "# AND MaCT<" + CStr(rs!MaCT) + " AND SoPS<>0 AND SoPS2Co<>0 ORDER BY NgayGS DESC, MaCT DESC", luongx)
            If tienx = 0 Then
                 tienx = SelectSQL("SELECT TOP 1 SoPS AS F1,SoPS2No AS F2 FROM ChungTu WHERE ((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + ")) AND MaTKNo=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND NgayGS>#" + Format(n, Mask_DB) + "# AND SoPS<>0 AND SoPS2No<>0 ORDER BY NgayGS", luongx)
            End If
            If luongx <> 0 And luongx <> rs!SoPS2Co Then tienx = RoundMoney(tienx * rs!SoPS2Co / luongx)
        Else
            luong = Fix(IIf(luong >= 0, 0.5, -0.5) + luong * Mask_N) / Mask_N
            If luong <> 0 And luong <> rs!SoPS2Co Then tienx = RoundMoney(tien * rs!SoPS2Co / luong) Else tienx = tien
        End If
        
        tienx = Abs(tienx)
        If Len(tkno) = 0 Then
            luong = luong - rs!SoPS2Co
            tien = tien - tienx
        End If
        
        If tienx <> rs!sops Then ExecuteSQL5 "UPDATE ChungTu SET SoPS=" + DoiDau(tienx) + " WHERE MaSo=" + CStr(rs!MaSo)
        
        If pGiaUSD > 0 Then
            If tien2 = 0 Then
                tienx2 = SelectSQL("SELECT TOP 1 PSUSD AS F1,SoPS2Co AS F2 FROM ChungTu WHERE (MaLoai=2 OR MaLoai=4) AND MaKho=" + CStr(mk) + " AND MaTKCo=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND NgayGS<=#" + Format(n, Mask_DB) + "# AND MaCT<" + CStr(rs!MaCT) + " AND SoPS<>0 AND SoPS2Co<>0 ORDER BY NgayGS DESC, MaCT DESC", luongx)
                If tienx2 = 0 Then
                     tienx2 = SelectSQL("SELECT TOP 1 PSUSD AS F1,SoPS2No AS F2 FROM ChungTu WHERE ((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + ")) AND MaTKNo=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND NgayGS>#" + Format(n, Mask_DB) + "# AND SoPS<>0 AND SoPS2No<>0 ORDER BY NgayGS", luongx)
                End If
                If luongx <> 0 And luongx <> rs!SoPS2Co Then tienx2 = Fix(0.5 + Mask_N * tienx2 * rs!SoPS2Co / luongx) / Mask_N
            Else
                luong = Fix(IIf(luong >= 0, 0.5, -0.5) + luong * Mask_N) / Mask_N
                If luong <> 0 And luong <> rs!SoPS2Co Then tienx2 = Fix(0.5 + Mask_N * tien2 * rs!SoPS2Co / luong) / Mask_N Else tienx2 = tien2
            End If
            
            tienx2 = Abs(tienx2)
            If Len(tkno) = 0 Then
                luong = luong - rs!SoPS2Co
                tien2 = tien2 - tienx2
            End If
            
            If tienx2 <> rs!PSUSD Then ExecuteSQL5 "UPDATE ChungTu SET SoPS=" + DoiDau(tienx2) + " WHERE MaSo=" + CStr(rs!MaSo)
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    If mv > 0 And ktra = 0 Then
        KiemTraTaiKhoan 1
        KiemTraVatTu 1
    End If
    
    'Beep
End Sub

Private Function SoTonKhoN2(ngay As Date, mkho As Long, mtk As Long, mvt As Long, ThanhTien As Double, mct As Long, tien2 As Double) As Double
    Dim sql As String, luong As Double, X As Double
    
    SoTonKhoN2 = SoTonKho(0, mkho, mtk, mvt, ThanhTien, tien2)
    sql = "SELECT Sum(SoPS) As F1,Sum(SoPS2No) As F2" + IIf(pGiaUSD > 0, ",Sum(PSUSD) As F3", "") + " FROM ChungTu WHERE MaVattu=" + CStr(mvt) + " AND (MaLoai=1" + IIf(mkho > 0, " OR MaLoai=4)", ")") + " AND NgayGS<=#" + Format(ngay, Mask_DB) + "# AND MaCT<" + CStr(mct)
    If mkho > 0 Then sql = sql + " AND ((MaKho=" + CStr(mkho) + " AND MaLoai=1) OR (MaNguon=" + CStr(mkho) + " AND MaLoai=4))"
    If mtk > 0 Then sql = sql + " And MaTKNo=" + CStr(mtk)
    
    ThanhTien = ThanhTien + SelectSQL(sql, luong, tien2)
    SoTonKhoN2 = SoTonKhoN2 + luong
    
    sql = "SELECT Sum(SoPS) As F1,Sum(SoPS2Co) As F2" + IIf(pGiaUSD > 0, ",Sum(PSUSD) As F3", "") + " FROM ChungTu WHERE MaVattu=" + CStr(mvt) + " AND (MaLoai=2" + IIf(mkho > 0, " OR MaLoai=4)", ")") + " AND NgayGS<#" + Format(ngay, Mask_DB) + "# AND MaCT<" + CStr(mct)
    If mkho > 0 Then sql = sql + " AND (MaKho=" + CStr(mkho) + ")"
    If mtk > 0 Then sql = sql + " And MaTKCo=" + CStr(mtk)
    
    ThanhTien = ThanhTien - SelectSQL(sql, luong, X)
    tien2 = tien2 - X
    SoTonKhoN2 = SoTonKhoN2 - luong
End Function

Public Function SoCPPB(tdau As Integer, tcuoi As Integer, f As String, m As Long) As Double
    Dim sql As String, i As Integer
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sql = sql + "+" + f + CStr(i)
    Next
    sql = "SELECT Sum(" + sql + ") AS F1 FROM TP154"
    If m > 0 Then sql = sql + " WHERE MaSo=" + CStr(m)
    SoCPPB = SelectSQL(sql)
End Function

Private Function SoNhapKho(mkho As Long, mtk As Long, mvt As Long, ndau As Date, ncuoi As Date, tien As Double, tien2 As Double) As Double
    Dim sql As String, luong As Double
    
    sql = "SELECT Sum(SoPS) As F1,Sum(SoPS2No) As F2" + IIf(pGiaUSD > 0, ",Sum(PSUSD) As F3", "") + " FROM ChungTu WHERE MaVattu=" + CStr(mvt) + " AND (MaLoai=1" + IIf(mkho > 0, " OR MaLoai=4)", ")") + " AND " + WNgay("NgayGS", ndau, ncuoi)
    If mkho > 0 Then sql = sql + " AND ((MaKho=" + CStr(mkho) + " AND MaLoai=1) OR (MaNguon=" + CStr(mkho) + " AND MaLoai=4))"
    If mtk > 0 Then sql = sql + " And MaTKNo=" + CStr(mtk)
    
    tien = SelectSQL(sql, luong, tien2)
    SoNhapKho = luong
End Function

Private Function SoNhapKhoThang(mkho As Long, mtk As Long, mvt As Long, tdau As Integer, tcuoi As Integer, tien As Double, tien2 As Double) As Double
    Dim sql As String, luong As Double
    
    sql = "SELECT Sum(SoPS) As F1,Sum(SoPS2No) As F2" + IIf(pGiaUSD > 0, ",Sum(PSUSD) As F3", "") + " FROM ChungTu WHERE MaVattu=" + CStr(mvt) + " AND (MaLoai=1" + IIf(mkho > 0, " OR MaLoai=4)", ")") + " AND " + WThang("ThangCT", tdau, tcuoi)
    If mkho > 0 Then sql = sql + " AND ((MaKho=" + CStr(mkho) + " AND MaLoai=1) OR (MaNguon=" + CStr(mkho) + " AND MaLoai=4))"
    If mtk > 0 Then sql = sql + " And MaTKNo=" + CStr(mtk)
    
    tien = SelectSQL(sql, luong, tien2)
    SoNhapKhoThang = luong
End Function

Public Function DinhMucNC(mvt As Long) As Double
    DinhMucNC = SelectSQL("SELECT DinhMuc.Soluong AS F1 FROM DinhMuc WHERE DinhMuc.MaTP=" + CStr(mvt) + " AND MaNVL=0")
End Function

Public Function QuyDoiTheoDVT1(mvt As Long, dvt As Long, SoLuong As Double) As Double
    Dim TyLe As Double
    If dvt = 0 Then
        QuyDoiTheoDVT1 = SoLuong
    Else
        TyLe = SelectSQL("SELECT TyLeQD AS F1 FROM DVTVattu WHERE MaSo=" + CStr(dvt) + " AND MaVattu=" + CStr(mvt))
        If TyLe = 0 Then
            QuyDoiTheoDVT1 = SoLuong
        Else
            QuyDoiTheoDVT1 = Fix(0.5 + Mask_N * SoLuong * TyLe) / Mask_N
        End If
    End If
End Function

Public Function QuyDoiTheoDVT2(mvt As Long, dvt As Long, SoLuong As Double) As Double
    Dim TyLe As Double
    If dvt = 0 Then
        QuyDoiTheoDVT2 = SoLuong
    Else
        TyLe = SelectSQL("SELECT TyLeQD AS F1 FROM DVTVattu WHERE MaSo=" + CStr(dvt) + " AND MaVattu=" + CStr(mvt))
        If TyLe = 0 Then
            QuyDoiTheoDVT2 = SoLuong
        Else
            QuyDoiTheoDVT2 = Fix(0.5 + Mask_N * SoLuong / TyLe) / Mask_N
        End If
    End If
End Function

Public Sub TinhGVBH(ndau As Date, ncuoi As Date, tl As Integer, mvt As Long, Optional ktra As Integer = 0)
      Dim ct As New ClsChungtu, rs As Recordset, luongtp As Double, m As Long, rs_ktra As Recordset, m1 As Long
      Dim i As Integer, luong() As Double, tien() As Double, id() As Long, c As Integer, tien2() As Double
      
      If OutCost <> 1 Then
        ExecuteSQL5 "UPDATE VTDauNam SET VTDauNam.SoXuat=0 WHERE (True) " + IIf(mvt > 0, " AND VTDauNam.MaVattu=" + CStr(mvt), "")
        ExecuteSQL5 "UPDATE ChungTu SET SoXuat=0 WHERE (True) " + IIf(mvt > 0, " AND MaVattu=" + CStr(mvt), "")
        
        ExecuteSQL5 "DELETE * FROM ChungTu WHERE MaLoai=2 AND RIGHT(SoHieu,2)='GV' AND SoPS=0 AND " + WNgay("NgayGS", ndau, ncuoi) + IIf(mvt > 0, " AND MaVattu=" + CStr(mvt), "")
        If tl = vbYes Then
            ExecuteSQL5 "DELETE * FROM ChungTu WHERE MaLoai=2 AND Right(SoHieu,2)='GV' AND " + WNgay("NgayGS", ndau, ncuoi) + IIf(mvt > 0, " AND MaVattu=" + CStr(mvt), "")
        End If
                
        SetSQL "MienTru", "SELECT DISTINCTROW ChungTu.MaSo,ChungTu.MaCT,ChungTu.NgayGS,ChungTu.SoHieu,ChungTu.MaKho,ChungTu.SoPS2Co,ChungTu.MaVattu,ChungTu.ThangCT FROM " + ChungTu2TKNC(1) _
            & " WHERE (ChungTu.MaLoai=8 OR ChungTu.MaLoai=4 OR (ChungTu.MaLoai=2 AND RIGHT(ChungTu.SoHieu,2)<>'GV')) AND ChungTu.MaVattu>0 AND (TK_ID=" + CStr(TKDT_ID) + " OR ChungTu.MaLoai<>8) AND (Not HethongTK.SoHieu LIKE '5113*') AND " + WNgay("ChungTu.NgayGS", ndau, ncuoi) + IIf(mvt > 0, " AND ChungTu.MaVattu=" + CStr(mvt), "") + " ORDER BY ChungTu.NgayCT,ChungTu.MaCT"
            
        Set rs_ktra = DBKetoan.OpenRecordset("MienTru", dbOpenSnapshot)
        Do While Not rs_ktra.EOF
              m1 = 0
              ct.InitChungtu rs_ktra!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
              If ct.maloai = 8 Then
                  If ct.sops = 0 Then
                        m1 = SelectSQL("SELECT MaTK AS F1 FROM KhoHang WHERE MaSo=" + CStr(ct.MaKho))
                        If m1 > 0 Then ct.tkno.InitTaikhoanMaSo m1
                  End If
                  If m1 = 0 Then
                        m1 = SelectSQL("SELECT MaTKGV AS F1 FROM KhoHang WHERE MaSo=" + CStr(ct.MaKho))
                        If m1 > 0 Then ct.tkno.InitTaikhoanMaSo m1
                  End If
                  If m1 = 0 Then ct.tkno.InitTaikhoanMaSo SelectSQL("SELECT TOP 1 MaSo AS F1 FROM HethongTK WHERE SoHieu LIKE '632*' AND TKCon=0 ORDER BY SoHieu")
                  If ct.tkno.MaSo = 0 Then GoTo KT
                  ct.sohieu = ct.sohieu + "GV"
                  ct.maloai = 2
                  If CTGS_GV > 0 Then ct.CTGS = CTGS_GV
                  ct.User_ID = 0
                  ct.CT_ID = 500000000 + rs_ktra!MaSo
                  If m <> rs_ktra!MaCT Then
                        m = rs_ktra!MaCT
                        m1 = SelectSQL("SELECT TOP 1 MaCT AS F1 FROM ChungTu WHERE MaLoai=2 AND SoHieu='" + rs_ktra!sohieu + "'+'GV' AND NgayGS=#" + Format(rs_ktra!NgayGS, Mask_DB) + "# ")
                        If m1 = 0 Then m1 = Lng_MaxValue("MaCT", "ChungTu") + 1
                  End If
                  ct.MaCT = m1
                  m = SelectSQL("SELECT MaTaiKhoan AS F1 FROM TonKho WHERE MaSoKho=" + CStr(ct.MaKho) + " AND MaVattu=" + CStr(ct.MaVattu) + " AND Luong_" + CStr(ct.ThangCT) + ">0")
                  If m = 0 Then m = SelectSQL("SELECT MaTK AS F1 FROM PhanLoaiVattu INNER JOIN Vattu ON PhanLoaiVattu.MaSo=Vattu.MaPhanLoai WHERE Vattu.MaSo=" + CStr(ct.MaVattu))
                  ct.TkCo.InitTaikhoanMaSo m
            
                  If ct.TkCo.MaSo = 0 Then GoTo KT
            Else
                ExecuteSQL5 "DELETE * FROM ChungTu WHERE MaSo=" + CStr(ct.MaSo)
            End If
                  Select Case OutCost
                        Case 0: ct.sops = GiaXuatKho(ct.MaKho, ct.TkCo.MaSo, ct.MaVattu, ct.NgayGS, ct.SoPS2Co, ct.PSUSD)
                        Case 2: c = GiaXuatKhoFIFO(ct.MaKho, ct.TkCo.MaSo, ct.MaVattu, ct.SoPS2Co, luong, tien, id, tien2)
                        Case 3: c = GiaXuatKhoLIFO(ct.MaKho, ct.TkCo.MaSo, ct.MaVattu, ct.SoPS2Co, luong, tien, id, tien2)
                  End Select
                  
                  If OutCost = 0 Then
                    ct.GhiChungtuTH 0, 1
                  Else
                    For i = 1 To c
                        ct.SoPS2Co = luong(i)
                        ct.sops = tien(i)
                        If pGiaUSD > 0 Then ct.PSUSD = tien2(i)
                        ct.CT_ID = id(i)
                        ct.GhiChungtuTH 0, 1
                    Next
    
                End If
KT:
              rs_ktra.MoveNext
        Loop
        If rs_ktra.recordCount > 0 And ktra = 0 Then
            KiemTraTaiKhoan 1
            KiemTraVatTu 1
        End If
        rs_ktra.Close
        Set rs_ktra = Nothing
    End If
    Erase luong
    Erase tien
    Erase tien2
    Erase id
    Set ct = Nothing
End Sub

Public Sub TinhGVBHBQ(tdau As Integer, tcuoi As Integer, tl As Integer, mvt As Long, loai As Integer)
      Dim m1 As Long, n1 As Long, sql As String
      
      If OutCost = 0 Then
        ExecuteSQL5 "DELETE * FROM ChungTu WHERE MaLoai=2 AND RIGHT(SoHieu,2)='GV' AND SoPS=0 AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mvt > 0, " AND MaVattu=" + CStr(mvt), "")
        If tl = vbYes Then
            ExecuteSQL5 "DELETE * FROM ChungTu WHERE MaLoai=2 AND Right(SoHieu,2)='GV' AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mvt > 0, " AND MaVattu=" + CStr(mvt), "")
        End If
        
        m1 = SelectSQL("SELECT TOP 1 MaSo AS F1 FROM HethongTK WHERE SoHieu LIKE '632*' AND TKCon=0 ORDER BY SoHieu")
        n1 = SelectSQL("SELECT TOP 1 MaSo AS F1 FROM HethongTK WHERE SoHieu LIKE '1561*' AND TKCon=0 ORDER BY SoHieu")
        sql = "INSERT INTO Chungtu (MaCT, MaLoai, SoHieu, ThangCT, NgayCT, NgayGS, MaNguon, MaKho, DienGiai," + IIf(pSongNgu, "DienGiaiE,", "") _
            + "MaTkNo, MaTkCo, SoPS, SoPS2No, SoPS2Co, MaVattu, GhiChu, CT_ID, MaDT, MaDT1, MaDT2, MaDT3,MaKH,CTGS,MaKHC,MaTP,DVT,User_ID,MaNV,HanTT,SH1,T1,TLCK,CK" + IIf(pTygia > 0, ",TyGia", "") + IIf(pGiaUSD > 0, ",PSUSD", "") + ") " _
            & "SELECT DISTINCTROW 1000000+MaCT,2,ChungTu.SoHieu+'GV',ThangCT,NgayCT,NgayGS,MaNguon,MaKho,DienGiai," + IIf(pSongNgu, "DienGiaiE,", "") _
            + "IIF(KhoHang.MaTKGV=0," + CStr(m1) + ",KhoHang.MaTKGV),IIF(PhanLoaiVattu.MaTK=0," + CStr(n1) + ",PhanLoaiVattu.MaTK), SoPS, SoPS2No, SoPS2Co, MaVattu, ChungTu.GhiChu, ChungTu.MaSo+500000000, MaDT, MaDT1, MaDT2, MaDT3,MaKH,CTGS,MaKHC,MaTP,DVT,User_ID,MaNV,HanTT,SH1,T1,TLCK,ChungTu.CK" + IIf(pTygia > 0, ",TyGia", "") + IIf(pGiaUSD > 0, ",PSUSD", "") _
            + " FROM (((" + ChungTu2TKNC(1) + ") INNER JOIN KhoHang ON ChungTu.MaKho=KhoHang.MaSo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo" _
            + " WHERE MaLoai=8 AND MaVattu>0 AND TK_ID=" + CStr(TKDT_ID) + " AND (Not HethongTK.SoHieu LIKE '5113*') AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mvt > 0, " AND MaVattu=" + CStr(mvt), "") + " ORDER BY ChungTu.NgayCT,ChungTu.MaCT"
        ExecuteSQL5 sql
        If loai = 1 Then
            TinhGXKBQ tdau, tcuoi, MaSo2SoHieu(mvt, "Vattu"), "632"
        Else
            TinhGXK tdau, tcuoi, MaSo2SoHieu(mvt, "Vattu"), "632"
        End If
    End If
End Sub

Public Function GiaNVLTheoDM(mtp As Long, sl As Double, mk As Long, ngay As Date) As Double
    Dim rs As Recordset
    Dim thangdm As Integer
    
    thangdm = SelectSQL("SELECT  TOP 1 Thang AS F1 FROM DinhMuc WHERE MaNVL>0 AND MaTP=" + CStr(mtp) + " AND " + WThang("Thang", 0, Month(ngay)) + " ORDER BY Thang DESC")
       
    Set rs = DBKetoan.OpenRecordset("SELECT MaNVL,SoLuong FROM DinhMuc WHERE MaNVL>0 AND MaTP=" + CStr(mtp) + " AND Thang=" + CStr(thangdm), dbOpenSnapshot)
    Do While Not rs.EOF
        GiaNVLTheoDM = GiaNVLTheoDM + GiaXuatKho(mk, 0, rs!MaNVL, ngay, sl * rs!SoLuong)
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Function

Public Function GiaNCTheoDM(mtp As Long, sl As Double, thang As Integer) As Double
    Dim thangdm As Integer
    thangdm = SelectSQL("SELECT  TOP 1 Thang AS F1 FROM DinhMuc WHERE MaNVL>0 AND MaTP=" + CStr(mtp) + " AND " + WThang("Thang", 0, thang) + " ORDER BY Thang DESC")
    GiaNCTheoDM = RoundMoney(sl * SelectSQL("SELECT SoLuong AS F1 FROM DinhMuc WHERE MaNVL=0 AND MaTP=" + CStr(mtp) + " AND Thang=" + CStr(thangdm)))
End Function

Public Function SoPSCPTP(sh As String, m As Cls154, tcuoi As Integer) As Double
    If m.GiaThanhCK(ThangTruoc(tcuoi)) > 0 Then
        SoPSCPTP = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE " + WThang("ThangCT", 0, tcuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*'" + IIf(m.MaSo > 0, " AND MaTP=" + CStr(m.MaSo), ""))
        SoPSCPTP = Abs(SoPSCPTP - SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(1) + " WHERE MaLoai<>3 AND " + WThang("ThangCT", 0, tcuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*'" + IIf(m.MaSo > 0, " AND MaTP=" + CStr(m.MaSo), "")))
    Else
        SoPSCPTP = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE " + WThang("ThangCT", tcuoi, tcuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*'" + IIf(m.MaSo > 0, " AND MaTP=" + CStr(m.MaSo), ""))
        SoPSCPTP = Abs(SoPSCPTP - SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(1) + " WHERE MaLoai<>3 AND " + WThang("ThangCT", tcuoi, tcuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*'" + IIf(m.MaSo > 0, " AND MaTP=" + CStr(m.MaSo), "")))
    End If
End Function

Public Function SoPSCPTP_DM(sh As String, m As Cls154, tcuoi As Integer) As Double
    If m.GiaThanhCK(ThangTruoc(tcuoi)) > 0 Then
        SoPSCPTP_DM = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE (CT_ID < 610000000 OR CT_ID > 620000000) AND " + WThang("ThangCT", 0, tcuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*'" + IIf(m.MaSo > 0, " AND MaTP=" + CStr(m.MaSo), ""))
        SoPSCPTP_DM = SoPSCPTP_DM - SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(1) + " WHERE MaLoai<>3 AND " + WThang("ThangCT", 0, tcuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*'" + IIf(m.MaSo > 0, " AND MaTP=" + CStr(m.MaSo), ""))
    Else
        SoPSCPTP_DM = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE (CT_ID < 610000000 OR CT_ID > 620000000) AND " + WThang("ThangCT", tcuoi, tcuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*'" + IIf(m.MaSo > 0, " AND MaTP=" + CStr(m.MaSo), ""))
        SoPSCPTP_DM = SoPSCPTP_DM - SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(1) + " WHERE MaLoai<>3 AND " + WThang("ThangCT", tcuoi, tcuoi) + " AND HethongTK.SoHieu LIKE '" + sh + "*'" + IIf(m.MaSo > 0, " AND MaTP=" + CStr(m.MaSo), ""))
    End If
End Function

Public Sub ChuyenKho(mk1 As Long, mk2 As Long)
    Dim rs As Recordset
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaTaiKhoan,MaVattu,Luong_0,Tien_0 FROM TonKho WHERE (Luong_0<>0 OR Tien_0<>0) AND MaSoKho=" + CStr(mk1))
    Do While Not rs.EOF
        ExecuteSQL5 "UPDATE TonKho SET Luong_0=Luong_0+" + DoiDau(rs!Luong_0) + ",Tien_0=Tien_0+" + DoiDau(rs!Tien_0) + " WHERE MaSoKho=" + CStr(mk2) + " AND MaTaiKhoan=" + CStr(rs!MaTaiKhoan) + " AND MaVattu=" + CStr(rs!MaVattu)
        If DBKetoan.RecordsAffected = 0 Then ExecuteSQL5 "UPDATE TonKho SET MaSoKho=" + CStr(mk2) + " WHERE MaSoKho=" + CStr(mk1) + " AND MaTaiKhoan=" + CStr(rs!MaTaiKhoan) + " AND MaVattu=" + CStr(rs!MaVattu)
        If OutCost <> 0 Then
            ExecuteSQL5 "UPDATE VTDauNam SET Luong_0=Luong_0+" + DoiDau(rs!Luong_0) + ",Tien_0=Tien_0+" + DoiDau(rs!Tien_0) + " WHERE MaSoKho=" + CStr(mk2) + " AND MaTaiKhoan=" + CStr(rs!MaTaiKhoan) + " AND MaVattu=" + CStr(rs!MaVattu)
            If DBKetoan.RecordsAffected = 0 Then ExecuteSQL5 "UPDATE VTDauNam SET MaSoKho=" + CStr(mk2) + " WHERE MaSoKho=" + CStr(mk1) + " AND MaTaiKhoan=" + CStr(rs!MaTaiKhoan) + " AND MaVattu=" + CStr(rs!MaVattu)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    ExecuteSQL5 "DELETE * FROM TonKho WHERE MaSoKho=" + CStr(mk1)
    If OutCost <> 0 Then ExecuteSQL5 "DELETE * FROM VTDauNam WHERE MaSoKho=" + CStr(mk1)
    ExecuteSQL5 "UPDATE ChungTu SET MaKho=" + CStr(mk2) + " WHERE (MaLoai=1 OR MaLoai=2) AND MaKho=" + CStr(mk1)
    ExecuteSQL5 "UPDATE ChungTu SET MaNguon=" + CStr(mk2) + " WHERE (MaLoai=4) AND MaNguon=" + CStr(mk1)
    KiemTraVatTu
End Sub

Public Function DaTinhGiaThanh154(ms As Long, thang As Integer) As Boolean
    DaTinhGiaThanh154 = (SelectSQL("SELECT TOP 1 MaSo AS F1 FROM ThanhPham WHERE Ma154=" + CStr(ms) + " AND Thang=" + CStr(thang)) > 0)
End Function

Public Sub DieuChinhGiaThanh(tp As Cls154, thang As Integer, Optional ktra As Integer = 0)
    Dim CP As Double, cpa As Double, rs As Recordset, sodong As Double, i As Double, cp1 As Double, cp1a As Double, cpcu As Double, tcp As Double, tcpa As Double, xk As Integer, n As Date
    Dim tp1 As New Cls154, sql As String, cpnc As Double, cpc As Double
    
    If tp.MaSo = 0 Then Exit Sub
    If Not DaTinhGiaThanh154(tp.MaSo, thang) Then Exit Sub
    sql = "SELECT * FROM ThanhPham WHERE Thang=" + CStr(thang) + " AND Ma154=" + CStr(tp.MaSo) + " ORDER BY CPNVL"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    rs.MoveLast
    sodong = rs.recordCount
    xk = rs!xk
    n = rs!ngay
    ' Tinh lai cp NVL
    CP = tp.SoCPNVL(thang, thang) + tp.SoCPNVLPB(thang, thang)
    cpcu = SelectSQL("SELECT Sum(CPNVL) AS F1 FROM ThanhPham WHERE Thang=" + CStr(thang) + " AND Ma154=" + CStr(tp.MaSo))
    If cpcu <> CP Then
        rs.MoveFirst
        For i = 1 To sodong - 1
            If cpcu <> 0 Then
                cp1 = RoundMoney(rs!CPNVL * CP / cpcu)
                ExecuteSQL5 "UPDATE ThanhPham SET GiaThanh=GiaThanh-CPNVL+" + DoiDau(cp1) + ",CPNVL=" + DoiDau(cp1) + " WHERE MaSo=" + CStr(rs!MaSo)
                tcp = tcp + cp1
            End If
            rs.MoveNext
        Next
        ' Dong cuoi
        ExecuteSQL5 "UPDATE ThanhPham SET GiaThanh=GiaThanh-CPNVL+" + DoiDau(CP - tcp) + ",CPNVL=" + DoiDau(CP - tcp) + " WHERE MaSo=" + CStr(rs!MaSo)
    End If
    ExecuteSQL5 "UPDATE ThanhPham SET CPSXC=CPNVL/" + DoiDau(IIf(CP <> 0, CP, 1)) + ",GiaThanh=GiaThanh-CPNC-CPSXC WHERE Thang=" + CStr(thang) + " AND Ma154=" + CStr(tp.MaSo)
    rs.Requery
    ' Tinh lai cp nc
    tcp = 0
    tcpa = 0
    CP = tp.SoCPNC(thang, thang) + tp.SoCPNCPB(thang, thang)
    cpa = tp.SoCPSXC(thang, thang) + tp.SoCPSXCTT(thang, thang)
    rs.MoveFirst
    For i = 1 To sodong - 1
        cp1 = RoundMoney(CP * rs!cpsxc)
        cp1a = RoundMoney(cpa * rs!cpsxc)
        ExecuteSQL5 "UPDATE ThanhPham SET GiaThanh=GiaThanh+" + DoiDau(cp1 + cp1a) + ",CPNC=" + DoiDau(cp1) + ",CPSXC=" + DoiDau(cp1a) + " WHERE MaSo=" + CStr(rs!MaSo)
        tcp = tcp + cp1
        tcpa = tcpa + cp1a
        rs.MoveNext
    Next
    ' Dong cuoi
    ExecuteSQL5 "UPDATE ThanhPham SET GiaThanh=GiaThanh+" + DoiDau(CP - tcp + cpa - tcpa) + ",CPNC=" + DoiDau(CP - tcp) + ",CPSXC=" + DoiDau(cpa - tcpa) + " WHERE MaSo=" + CStr(rs!MaSo)
        
    GhiXuatNVL rs!MaCT, n, thang, IIf(ktra = 0, 0, xk), tp, 1
    rs.Close
    
    SetSQL "MienTru", sql
    Set rs = DBKetoan.OpenRecordset("SELECT ChungTu.MaTP FROM ChungTu INNER JOIN MienTru ON ChungTu.MaVattu=MienTru.MaTP WHERE ChungTu.MaTP>0 AND ChungTu.MaLoai=2 AND ChungTu.ThangCT=" + CStr(thang) + " AND CT_ID>610000000 GROUP BY ChungTu.MaTP", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        tp1.InitTPMaSo rs!MaTP
        DieuChinhGiaThanh tp1, thang, 1
        rs.MoveNext
    Loop
        
    rs.Close
    Set rs = Nothing
    Set tp1 = Nothing
    
    If ktra = 0 Then
        TinhLaiGiaVonTP tp, thang
        KiemTraVatTu
        KiemTraTaiKhoan
        HienThongBao "", 1
    End If
End Sub

Public Sub TinhLaiGiaThanhPham(tdau As Integer, tcuoi As Integer)
    Dim rs As Recordset, tp As New Cls154
        
    Set rs = DBKetoan.OpenRecordset("SELECT ThanhPham.* FROM ThanhPham WHERE " + WThang("Thang", tdau, tcuoi), dbOpenSnapshot)
    Do While Not rs.EOF
        tp.InitTPMaSo rs!Ma154
        DieuChinhGiaThanh tp, rs!thang, 1
        rs.MoveNext
    Loop
    If rs.recordCount > 0 Then
        KiemTraVatTu
        KiemTraTaiKhoan
    End If
    rs.Close
    Set rs = Nothing
    Set tp = Nothing
End Sub

Public Sub TinhLaiGiaVonTP(tp As Cls154, thang As Integer)
    Dim rs As Recordset
    
    If tp.MaSo = 0 Then Exit Sub
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaTP FROM ThanhPham WHERE Ma154=" + CStr(tp.MaSo) + " AND Thang=" + CStr(thang) + " GROUP BY MaTP", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        TinhGVBH NgayDauThang(pNamTC, thang), NgayCuoiThang(pNamTC, thang), vbYes, rs!MaTP, 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Public Sub DieuChinhGiaTP(thang As Integer, mtp As Long, mtk As Long, mvt As Long, dgia As Double, tien As Double)
    Dim dg As Double, m As Long, rs As Recordset, T As Double
    
    If pGiaHT = 0 Then Exit Sub
    If dgia = 0 Then dg = SelectSQL("SELECT GiaHT AS F1 FROM Vattu WHERE MaSo=" + CStr(mvt)) Else dg = dgia
    ExecuteSQL5 "UPDATE " + ChungTu2TKNC(1) + "  SET SoPS=Fix(0.5+SoPS2No*" + DoiDau(dg) + ") " _
        & " WHERE ThangCT=" + CStr(thang) + IIf(mtp > 0, " AND MaTP=" + CStr(mtp), "") + " AND MaTKNo=" + CStr(mtk) _
        + " AND MaVattu=" + CStr(mvt) + " AND HethongTK.SoHieu LIKE '" + ShTkSPDo + "*'"
    
    SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKNC(1) + " WHERE MaLoai=1 AND ThangCT=" + CStr(thang) + " AND MaTKNo=0 AND HethongTK.SoHieu LIKE '" + ShTkSPDo + "*' GROUP BY MaCT"
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo,ChungTu.MaCT,SoPS2No,SoPS FROM ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT WHERE MaTP=" + CStr(mtp) + " AND MaTKNo=" + CStr(mtk) + " AND MaVattu=" + CStr(mvt), dbOpenSnapshot)
    Do While Not rs.EOF
        T = RoundMoney(rs!SoPS2No * dg)
        If T <> rs!sops Then
            ExecuteSQL5 "UPDATE ChungTu SET SoPS=" + DoiDau(T) + " WHERE MaSo=" + CStr(rs!MaSo)
            ExecuteSQL5 "UPDATE " + ChungTu2TKNC(1) + " SET SoPS=SoPS+" + DoiDau(T - rs!sops) + " WHERE MaTP=" + CStr(mtp) + " AND MaCT=" + CStr(rs!MaCT) + " AND HethongTK.SoHieu LIKE '" + ShTkSPDo + "*'"
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    dg = SelectSQL("SELECT Sum(SoPS) AS F1,Max(IIF(MaTKNo>0 And MaTKCo>0,ChungTu.MaSo)) AS F2 FROM ChungTu WHERE MaLoai=1 AND MaTP=" + CStr(mtp) + " AND MaTKNo=" + CStr(mtk) + " AND MaVattu=" + CStr(mvt) + " AND ThangCT=" + CStr(thang), m)
    If dg <> tien Then ExecuteSQL5 "UPDATE ChungTu SET SoPS=SoPS+" + DoiDau(tien - dg) + " WHERE MaSo=" + CStr(m)
    TinhGXK thang, thang, MaSo2SoHieu(mvt, "Vattu"), "", 1
End Sub

Public Sub DatLaiGiaHT(mtp As Long)
    Dim rs As Recordset
    
    If pGiaHT = 0 Then Exit Sub
    Set rs = DBKetoan.OpenRecordset("SELECT * FROM ThanhPham WHERE MaSo=" + CStr(mtp))
    Do While Not rs.EOF
        DieuChinhGiaTP rs!thang, rs!Ma154, rs!MaTK, rs!MaTP, 0, 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Sub TinhGXKFIFO(tdau As Integer, tcuoi As Integer, shvt As String, tkno As String, Optional ktra As Integer = 0)
    Dim rs As Recordset, sql As String, tien As Double, luong As Double, ps As Double, nx As Integer, tien2 As Double, ps2 As Double, soton As Double
    Dim mk As Long, mv As Long, mt As Long, ms As Long, m As Long
    Dim ctu As New ClsChungtu
    
    If OutCost <> 2 Then Exit Sub
    
    If OutCost <> 0 Then
          If SelectSQL("SELECT Max(MaVattu) AS F1 FROM VTDauNam") = 0 Then
              ExecuteSQL5 "INSERT INTO VTDauNam (MaSo,MaSoKho, MaTaiKhoan, MaVattu, Luong_0, Tien_0) SELECT MaSo,MaSoKho, MaTaiKhoan, MaVattu, Luong_0, Tien_0 FROM TonKho WHERE Luong_0<>0 OR Tien_0<>0"
          End If
    End If
    
    If Len(tkno) > 0 Then
        sql = "SELECT DISTINCTROW ChungTu.MaSo,ThangCT,NgayGS,MaCT,MaKho,MaVattu,MaTKCo,SoPS,SoPS2Co," + IIf(pGiaUSD > 0, "PSUSD", "0") + " AS SoPS2 FROM (" + ChungTu2TKNC(0) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE HethongTK.SoHieu LIKE '" + tkno + "*' AND (MaLoai=2 OR MaLoai=4) AND MaVattu>0 AND MaTKNo>0 AND TK.TK_ID=" + CStr(TKVT_ID) + " AND " + WThang("ThangCT", tdau, tcuoi) + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " ORDER BY MaKho,MaTKCo,MaVattu,ThangCT,NgayGS,ChungTu.MaCT,ChungTu.MaSo"
    Else
        sql = "SELECT DISTINCTROW ChungTu.MaSo,ThangCT,NgayGS,MaCT,MaKho,MaVattu,MaTKCo,SoPS,SoPS2Co," + IIf(pGiaUSD > 0, "PSUSD", "0") + " AS SoPS2 FROM (" + ChungTu2TKNC(1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE (MaLoai=2 OR MaLoai=4) AND MaVattu>0 AND HethongTK.TK_ID=" + CStr(TKVT_ID) + " AND " + WThang("ThangCT", tdau, tcuoi) + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " ORDER BY MaKho,MaTKCo,MaVattu,ThangCT,NgayGS,ChungTu.MaCT,ChungTu.MaSo"
    End If
    
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        If mk <> rs!MaKho Or mt <> rs!MaTkCo Or mv <> rs!MaVattu Then
            mk = rs!MaKho
            mt = rs!MaTkCo
            mv = rs!MaVattu
            soton = 0
            ms = 0
            luong = 0
            tien = 0
            ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=0 WHERE MaSoKho=" + CStr(mk) + " AND MaTaiKhoan=" + CStr(mt) + " AND MaVattu=" + CStr(mv)
            ExecuteSQL5 "UPDATE ChungTu SET SoXuat=0 WHERE ((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + ")) AND MaTKNo=" + CStr(mt) + " AND MaVattu=" + CStr(mv)
        End If

        If soton <= 0 Then
            nx = 0
            If rs!SoPS2Co > 0 Then
                luong = SelectSQL("SELECT TOP 1 Luong_0 AS F1,Tien_0 AS F2,MaSo AS F3" + IIf(pGiaUSD > 0, ",USDTien_0 AS F4", "") + ", Luong_0-SoXuat AS F5 FROM VTDauNam WHERE MaSoKho=" + CStr(mk) + " AND MaTaiKhoan=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND Luong_0>SoXuat ORDER BY MaSo", tien, ms, tien2, soton)
                If luong <= 0 Then
                    nx = 1
                    luong = SelectSQL("SELECT TOP 1 SoPS2No AS F1,SoPS AS F2,MaSo AS F3" + IIf(pGiaUSD > 0, ",PSUSD AS F4", "") + ", SoPS2No-SoXuat AS F5 FROM ChungTu WHERE ((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + ")) AND MaTKNo=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND SoPS2No>SoXuat ORDER BY ThangCT, NgayGS, MaCT, MaSo", tien, ms, tien2, soton)
                End If
            Else
                luong = SelectSQL("SELECT TOP 1 Luong_0 AS F1,Tien_0 AS F2,MaSo AS F3" + IIf(pGiaUSD > 0, ",USDTien_0 AS F4", "") + ", Luong_0-SoXuat AS F5 FROM VTDauNam WHERE MaSoKho=" + CStr(mk) + " AND MaTaiKhoan=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND Luong_0=0 AND SoXuat=0 ORDER BY MaSo", tien, ms, tien2, soton)
                If tien <= 0 Then
                    nx = 1
                    luong = SelectSQL("SELECT TOP 1 SoPS2No AS F1,SoPS AS F2,MaSo AS F3" + IIf(pGiaUSD > 0, ",PSUSD AS F4", "") + ", SoPS2No-SoXuat AS F5 FROM ChungTu WHERE ((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + ")) AND MaTKNo=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND SoPS2No=0 AND SoXuat=0 ORDER BY ThangCT, NgayGS, MaCT, MaSo", tien, ms, tien2, soton)
                End If
            End If
            
        End If
        
        If (soton > 0 And luong > 0) Or (rs!SoPS2Co = 0 And tien > 0) Then
            m = rs!MaSo

            If rs!SoPS2Co > soton Then
                ctu.InitChungtu m, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
                ctu.SoPS2Co = rs!SoPS2Co - soton
                ctu.sops = 0
                ctu.MaSo = 0
                ctu.CT_ID = 0
                ctu.GhiChungtu
                
                ExecuteSQL5 "UPDATE ChungTu SET SoPS2Co=" + DoiDau(soton) + ",SoPS=0 WHERE MaSo=" + CStr(m)
                DBEngine.Idle
                rs.Requery
                rs.FindFirst "MaSo=" + CStr(m)
            End If
            
            If rs!SoPS2Co > 0 Then
                soton = soton - rs!SoPS2Co
                ps = RoundMoney(rs!SoPS2Co * tien / luong)
                ps2 = RoundMoney(rs!SoPS2Co * tien2 / luong)
            Else
                ps = tien
                ps2 = tien2
            End If
            
            ExecuteSQL5 "UPDATE ChungTu SET SoPS=" + DoiDau(ps) + ", CT_ID=-" + CStr(IIf(nx = 0, 2000000000, 0) + ms) + IIf(pGiaUSD > 0, ",PSUSD=" + DoiDau(ps2), "") + " WHERE MaSo=" + CStr(m)
            If nx = 0 Then
                ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=SoXuat+" + DoiDau(IIf(rs!SoPS2Co > 0, rs!SoPS2Co, 1)) + " WHERE MaSo=" + CStr(ms)
            Else
                ExecuteSQL5 "UPDATE ChungTu SET SoXuat=SoXuat+" + DoiDau(IIf(rs!SoPS2Co > 0, rs!SoPS2Co, 1)) + " WHERE MaSo=" + CStr(ms)
            End If
            DBEngine.Idle
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set ctu = Nothing
    If mv > 0 And ktra = 0 Then
        KiemTraTaiKhoan
        KiemTraVatTu 1
    End If
End Sub

Public Sub TinhGXKDD(tdau As Integer, tcuoi As Integer, shvt As String, tkno As String, Optional ktra As Integer = 0)
    Dim rs As Recordset, sql As String, tien As Double, luong As Double, ps As Double, nx As Integer, tien2 As Double, ps2 As Double, soton As Double
    Dim mk As Long, mv As Long, mt As Long, ms As Long, m As Long
    Dim ctu As New ClsChungtu
    
    If OutCost <> 1 Then Exit Sub

    ExecuteSQL5 "UPDATE (VTDauNam INNER JOIN HethongTK ON VTDauNam.MaTaiKhoan=HethongTK.MaSo) INNER JOIN Vattu ON VTDauNam.MaVattu=Vattu.MaSo SET SoXuat=0 WHERE HethongTK.SoHieu LIKE '" + tkno + "*'" + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "")
    
    If Len(tkno) > 0 Then
        sql = "SELECT ChungTu.MaSo,ThangCT, ChungTu.SoHieu, MaTKCo,MaVattu,SoPS2Co,SoPS,CT_ID FROM (" + ChungTu2TKNC(0) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE (MaLoai=2 OR MaLoai=4) AND ABS(CT_ID)>2000000000 AND TK.TK_ID=" + CStr(TKVT_ID) + " AND HethongTK.SoHieu LIKE '" + tkno + "*'" + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " ORDER BY ThangCT, ChungTu.SoHieu"
    Else
        sql = "SELECT ChungTu.MaSo,ThangCT, ChungTu.SoHieu, MaTKCo,MaVattu,SoPS2Co,SoPS,CT_ID FROM (" + ChungTu2TKNC(1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE (MaLoai=2 OR MaLoai=4) AND ABS(CT_ID)>2000000000 AND HethongTK.TK_ID=" + CStr(TKVT_ID) + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " ORDER BY ThangCT, ChungTu.SoHieu"
    End If
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
       m = Abs(rs!CT_ID) - 2000000000
       ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=SoXuat+" + DoiDau(rs!SoPS2Co) + " WHERE MaVattu=" + CStr(rs!MaVattu) + " AND MaSo=" + CStr(Abs(rs!CT_ID) - 2000000000) + " AND SoXuat+" + DoiDau(rs!SoPS2Co) + ">=Luong_0"
       If DBKetoan.RecordsAffected = 0 Then ExecuteSQL5 "UPDATE ChungTu SET CT_ID=0 WHERE MaSo=" + CStr(rs!MaSo)
       rs.MoveNext
    Loop
    
    ExecuteSQL5 "UPDATE ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo SET ChungTu.SoXuat=0 WHERE (MaLoai=1 OR MaLoai=4) AND TK_ID=" + CStr(TKVT_ID)
    If Len(tkno) > 0 Then
        sql = "SELECT ChungTu.MaSo,ThangCT, ChungTu.SoHieu, MaTKCo,MaVattu,SoPS2Co,SoPS,CT_ID FROM (" + ChungTu2TKNC(0) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE (MaLoai=2 OR MaLoai=4) AND ABS(CT_ID)<2000000000 AND TK.TK_ID=" + CStr(TKVT_ID) + " AND HethongTK.SoHieu LIKE '" + tkno + "*'" + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " ORDER BY ThangCT, ChungTu.SoHieu"
    Else
        sql = "SELECT ChungTu.MaSo,ThangCT, ChungTu.SoHieu, MaTKCo,MaVattu,SoPS2Co,SoPS,CT_ID FROM (" + ChungTu2TKNC(1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE (MaLoai=2 OR MaLoai=4) AND ABS(CT_ID)<2000000000 AND HethongTK.TK_ID=" + CStr(TKVT_ID) + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " ORDER BY ThangCT, ChungTu.SoHieu"
    End If
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
       ExecuteSQL5 "UPDATE ChungTu SET SoXuat=SoXuat+" + DoiDau(rs!SoPS2Co) + " WHERE (MaLoai=1 OR MaLoai=4) AND MaVattu=" + CStr(rs!MaVattu) + " AND MaSo=" + CStr(Abs(rs!CT_ID)) + " AND SoXuat+" + DoiDau(rs!SoPS2Co) + ">=SoPS2No"
       If DBKetoan.RecordsAffected = 0 Then ExecuteSQL5 "UPDATE ChungTu SET CT_ID=0 WHERE MaSo=" + CStr(rs!MaSo)
       rs.MoveNext
    Loop
    rs.Close
             
    If Len(tkno) > 0 Then
        sql = "SELECT DISTINCTROW ChungTu.MaSo,NgayGS,MaCT,MaKho,MaVattu,MaTKCo,SoPS,SoPS2Co," + IIf(pGiaUSD > 0, "PSUSD", "0") + " AS SoPS2 FROM (" + ChungTu2TKNC(0) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE HethongTK.SoHieu LIKE '" + tkno + "*' AND (MaLoai=2 OR MaLoai=4) AND MaVattu>0 AND MaTKNo>0 AND TK.TK_ID=" + CStr(TKVT_ID) + " AND SoPS2Co>0 AND " + WThang("ThangCT", tdau, tcuoi) + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " AND CT_ID=0 ORDER BY ThangCT,NgayGS,ChungTu.MaCT,MaKho,MaTKCo,MaVattu, ChungTu.MaSo"
    Else
        sql = "SELECT DISTINCTROW ChungTu.MaSo,NgayGS,MaCT,MaKho,MaVattu,MaTKCo,SoPS,SoPS2Co," + IIf(pGiaUSD > 0, "PSUSD", "0") + " AS SoPS2 FROM (" + ChungTu2TKNC(1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE (MaLoai=2 OR MaLoai=4) AND MaVattu>0 AND HethongTK.TK_ID=" + CStr(TKVT_ID) + " AND SoPS2Co>0 AND " + WThang("ThangCT", tdau, tcuoi) + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " AND CT_ID=0 ORDER BY ThangCT,NgayGS,ChungTu.MaCT,MaKho,MaTKCo,MaVattu, ChungTu.MaSo"
    End If
        
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        If mk <> rs!MaKho Or mt <> rs!MaTkCo Or mv <> rs!MaVattu Then
            mk = rs!MaKho
            mt = rs!MaTkCo
            mv = rs!MaVattu
            soton = 0
            ms = 0
            luong = 0
            tien = 0
        End If

        If soton <= 0 Then
            nx = 0
            luong = SelectSQL("SELECT TOP 1 Luong_0 AS F1,Tien_0 AS F2,MaSo AS F3" + IIf(pGiaUSD > 0, ",USDTien_0 AS F4", "") + ", Luong_0-SoXuat AS F5 FROM VTDauNam WHERE MaSoKho=" + CStr(mk) + " AND MaTaiKhoan=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND Luong_0>SoXuat ORDER BY MaSo", tien, ms, tien2, soton)
            If luong <= 0 Then
                nx = 1
                luong = SelectSQL("SELECT TOP 1 SoPS2No AS F1,SoPS AS F2,MaSo AS F3" + IIf(pGiaUSD > 0, ",PSUSD AS F4", "") + ", SoPS2No-SoXuat AS F5 FROM ChungTu WHERE ((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + ")) AND MaTKNo=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND SoPS2No>SoXuat ORDER BY ThangCT, NgayGS, MaCT, MaSo", tien, ms, tien2, soton)
            End If
        End If
        
        If soton > 0 And luong > 0 Then
            m = rs!MaSo

            If rs!SoPS2Co > soton Then
                ctu.InitChungtu m, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
                ctu.SoPS2Co = rs!SoPS2Co - soton
                ctu.sops = 0
                ctu.MaSo = 0
                ctu.CT_ID = 0
                ctu.GhiChungtu
                
                ExecuteSQL5 "UPDATE ChungTu SET SoPS2Co=" + DoiDau(soton) + ",SoPS=0 WHERE MaSo=" + CStr(m)
                DBEngine.Idle
                rs.Requery
                rs.FindFirst "MaSo=" + CStr(m)
            End If
            
            soton = soton - rs!SoPS2Co
            ps = RoundMoney(rs!SoPS2Co * tien / luong)
            ps2 = RoundMoney(rs!SoPS2Co * tien2 / luong)
            
            ExecuteSQL5 "UPDATE ChungTu SET SoPS=" + DoiDau(ps) + ", CT_ID=-" + CStr(IIf(nx = 0, 2000000000, 0) + ms) + IIf(pGiaUSD > 0, ",PSUSD=" + DoiDau(ps2), "") + " WHERE MaSo=" + CStr(m)
            If nx = 0 Then
                ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=SoXuat+" + DoiDau(rs!SoPS2Co) + " WHERE MaSo=" + CStr(ms)
            Else
                ExecuteSQL5 "UPDATE ChungTu SET SoXuat=SoXuat+" + DoiDau(rs!SoPS2Co) + " WHERE MaSo=" + CStr(ms)
            End If
            DBEngine.Idle
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set ctu = Nothing
    If mv > 0 And ktra = 0 Then
        KiemTraTaiKhoan
        KiemTraVatTu 1
    End If
End Sub

Public Function GiaBanQD(mdv As Long, hsqd As Double) As Double
    GiaBanQD = SelectSQL("SELECT GiaBan AS F1, TyLeQD AS F2 FROM DVTVattu WHERE MaSo=" + CStr(mdv), hsqd)
End Function

Public Function GiaXuatKhoFIFO(mk As Long, mt As Long, mv As Long, sl As Double, SoLuong() As Double, tien() As Double, id() As Long, tien2() As Double) As Integer
    Dim rs As Recordset, T As Double, L As Double, c As Integer, slx As Double, t2 As Double
    
    If OutCost <> 2 Then Exit Function
    
    If sl = 0 Then
        Set rs = DBKetoan.OpenRecordset("SELECT MaSo, Tien_0 AS Tien," + IIf(pGiaUSD > 0, "USDTien_0", "0") + " AS Tien2 FROM VTDauNam WHERE MaSoKho=" + CStr(mk) + " AND MaTaiKhoan=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND Luong_0=0 AND SoXuat=0 ORDER BY MaSo", dbOpenSnapshot)
        Do While (Not rs.EOF)
            c = c + 1
            ReDim Preserve SoLuong(1 To c) As Double
            ReDim Preserve tien(1 To c) As Double
            ReDim Preserve tien2(1 To c) As Double
            ReDim Preserve id(1 To c) As Long
            SoLuong(c) = 0
            tien(c) = rs!tien
            tien2(c) = rs!tien2
            id(c) = -(2000000000 + rs!MaSo)
            rs.MoveNext
        Loop
        rs.Close
        
        Set rs = DBKetoan.OpenRecordset("SELECT MaSo, SoPS AS Tien," + IIf(pGiaUSD > 0, "PSUSD", "0") + " AS Tien2 FROM ChungTu WHERE ((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + ")) AND MaTKNo=" + CStr(mt) + " AND SoPS2No=0 AND SoXuat=0 AND MaVattu=" + CStr(mv) + " ORDER BY NgayGS, MaCT, MaSo", dbOpenSnapshot)
        Do While (Not rs.EOF)
            c = c + 1
            ReDim Preserve SoLuong(1 To c) As Double
            ReDim Preserve tien(1 To c) As Double
            ReDim Preserve tien2(1 To c) As Double
            ReDim Preserve id(1 To c) As Long
            SoLuong(c) = 0
            tien(c) = rs!tien
            tien2(c) = rs!tien2
            id(c) = -(rs!MaSo)
            rs.MoveNext
        Loop
        rs.Close
        GoTo KT
    End If
    
    slx = sl
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo, Luong_0-SoXuat AS LuongTon, Luong_0 AS Luong,SoXuat,Tien_0 AS Tien," + IIf(pGiaUSD > 0, "USDTien_0", "0") + " AS Tien2 FROM VTDauNam WHERE MaSoKho=" + CStr(mk) + " AND MaTaiKhoan=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND Luong_0<>0 AND Luong_0>SoXuat ORDER BY MaSo", dbOpenSnapshot)
    Do While (Not rs.EOF) And (slx > 0)
        L = IIf(slx > rs!luongton, rs!luongton, slx)
        T = RoundMoney(L * rs!tien / rs!luong)
        t2 = RoundMoney(L * rs!tien2 / rs!luong)
        c = c + 1
        ReDim Preserve SoLuong(1 To c) As Double
        ReDim Preserve tien(1 To c) As Double
        ReDim Preserve tien2(1 To c) As Double
        ReDim Preserve id(1 To c) As Long
        SoLuong(c) = L
        tien(c) = T
        tien2(c) = t2
        id(c) = -(2000000000 + rs!MaSo)
        slx = slx - L
        rs.MoveNext
    Loop
    rs.Close
    
    If slx = 0 Then GoTo KT
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo, SoPS2No-SoXuat AS Luong,SoPS2No,SoPS AS Tien," + IIf(pGiaUSD > 0, "PSUSD", "0") + " AS Tien2 FROM ChungTu WHERE ((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + ")) AND MaTKNo=" + CStr(mt) + " AND SoPS2No-SoXuat>0 AND MaVattu=" + CStr(mv) + " ORDER BY NgayGS, MaCT, MaSo", dbOpenSnapshot)
    Do While (Not rs.EOF) And (slx > 0)
        L = IIf(slx > rs!luong, rs!luong, slx)
        T = RoundMoney(L * rs!tien / rs!SoPS2No)
        t2 = RoundMoney(L * rs!tien2 / rs!SoPS2No)
        c = c + 1
        ReDim Preserve SoLuong(1 To c) As Double
        ReDim Preserve tien(1 To c) As Double
        ReDim Preserve tien2(1 To c) As Double
        ReDim Preserve id(1 To c) As Long
        SoLuong(c) = L
        tien(c) = T
        tien2(c) = t2
        id(c) = -(rs!MaSo)
        slx = slx - L
        rs.MoveNext
    Loop
    rs.Close
    
    'If c > 0 Then
    '    L = SoTonKho(ThangCuoiNamTC, mk, mt, mv, T, t2)
    '    If L = 0 And T <> 0 Then
    '        tien(c) = tien(c) + T
    '        tien2(c) = tien2(c) + T
    '    End If
    'End If
KT:
    Set rs = Nothing
    GiaXuatKhoFIFO = c
End Function

Public Function GiaXuatKhoLIFO(mk As Long, mt As Long, mv As Long, sl As Double, SoLuong() As Double, tien() As Double, id() As Long, tien2() As Double) As Integer
    Dim rs As Recordset, T As Double, L As Double, c As Integer, slx As Double, t2 As Double
    
    If OutCost <> 3 Or sl = 0 Then Exit Function
        
    slx = sl
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo, SoPS2No-SoXuat AS Luong, SoPS2No,SoPS AS Tien," + IIf(pGiaUSD > 0, "PSUSD", "0") + " AS Tien2 FROM ChungTu WHERE ((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + ")) AND MaTKNo=" + CStr(mt) + " AND SoPS2No-SoXuat>0 AND MaVattu=" + CStr(mv) + " ORDER BY NgayGS DESC, MaCT DESC, MaSo DESC", dbOpenSnapshot)
    Do While (Not rs.EOF) And (slx > 0)
        L = IIf(slx > rs!luong, rs!luong, slx)
        T = RoundMoney(L * rs!tien / rs!SoPS2No)
        t2 = RoundMoney(L * rs!tien2 / rs!SoPS2No)
        c = c + 1
        ReDim Preserve SoLuong(1 To c) As Double
        ReDim Preserve tien(1 To c) As Double
        ReDim Preserve tien2(1 To c) As Double
        ReDim Preserve id(1 To c) As Long
        SoLuong(c) = L
        tien(c) = T
        tien2(c) = t2
        id(c) = -(rs!MaSo)
        slx = slx - L
        rs.MoveNext
    Loop
    rs.Close
    
    If slx = 0 Then GoTo KT
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo, Luong_0 AS Luong,Tien_0 AS Tien," + IIf(pGiaUSD > 0, "USDTien_0", "0") + " AS Tien2 FROM VTDauNam WHERE MaSoKho=" + CStr(mk) + " AND MaTaiKhoan=" + CStr(mt) + " AND MaVattu=" + CStr(mv) + " AND Luong_0<>0 AND Luong_0>SoXuat ORDER BY MaSo DESC", dbOpenSnapshot)
    Do While (Not rs.EOF) And (slx > 0)
        L = IIf(slx > rs!luong, rs!luong, slx)
        T = RoundMoney(L * rs!tien / rs!luong)
        t2 = RoundMoney(L * rs!tien2 / rs!luong)
        c = c + 1
        ReDim Preserve SoLuong(1 To c) As Double
        ReDim Preserve tien(1 To c) As Double
        ReDim Preserve tien2(1 To c) As Double
        ReDim Preserve id(1 To c) As Long
        SoLuong(c) = L
        tien(c) = T
        tien2(c) = t2
        id(c) = -(2000000000 + rs!MaSo)
        slx = slx - L
        rs.MoveNext
    Loop
    rs.Close
    
KT:
    Set rs = Nothing
    GiaXuatKhoLIFO = c
End Function

Public Function GiaXuatKhoDD(mk As Long, mt As Long, mv As Long, sl As Double, SoLuong() As Double, tien() As Double, id() As Long, tien2() As Double) As Integer
    Dim slx As Double, vt As New ClsVattu, L As Double, T As Double, id1 As Long, c As Integer, t2 As Double
    
    If OutCost <> 1 Or sl = 0 Then Exit Function
    vt.InitVattuMaSo mv
    
    Do While (slx < sl)
        FDsNhap.tag = mv
        id1 = FDsNhap.XuatDichDanh(13, vt.sohieu + " - " + vt.TenVattu + ABCtoVNI(" - §.v.t: ") + vt.DonVi, mk, L, T, t2)
        If L > 0 Then
            'slx = slx + L
            
            If slx + L > sl Then
                t2 = sl * t2 / L
                T = sl * T / L
                L = sl - slx
            End If
            slx = slx + L
            
            c = c + 1
            ReDim Preserve SoLuong(1 To c) As Double
            ReDim Preserve tien(1 To c) As Double
            ReDim Preserve tien2(1 To c) As Double
            ReDim Preserve id(1 To c) As Long
            SoLuong(c) = L
            tien(c) = T
            tien2(c) = t2
            id(c) = -(id1)
        End If
    Loop
    
    GiaXuatKhoDD = c
    Set vt = Nothing
End Function

Public Function XuatTheoKho(kx As Long, kn As Long, mvt As Long, ndau As Date, ncuoi As Date, tien As Double) As Double
    XuatTheoKho = SelectSQL("SELECT Sum(SoPS2Co) AS F1, Sum(SoPS) AS F2 FROM ChungTu WHERE MaLoai=4 AND MaKho=" + CStr(kx) + " AND MaNguon=" + CStr(kn) + " AND MaVattu=" + CStr(mvt) + " AND " + WNgay("NgayGS", ndau, ncuoi), tien)
End Function

Public Function XuatBan(shtk As String, kx As Long, mn As Long, mvt As Long, ndau As Date, ncuoi As Date, tien As Double) As Double
    XuatBan = SelectSQL("SELECT Sum(SoPS2Co) AS F1, Sum(SoPS) AS F2 FROM " + ChungTu2TKNC(0) + " WHERE MaLoai=2 AND MaKho=" + CStr(kx) + " AND MaNguon=" + CStr(mn) + " AND HethongTK.SoHieu LIKE '632*' AND TK.SoHieu LIKE '" + shtk + "*' AND MaVattu=" + CStr(mvt) + " AND " + WNgay("NgayGS", ndau, ncuoi), tien)
End Function

Public Function SoCPTT(sh As String, tdau As Integer, tcuoi As Integer) As Double
    SoCPTT = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE MaTP>0 AND HethongTK.SoHieu LIKE '" + sh + "*' AND " + WThang("ThangCT", tdau, tcuoi))
End Function

Public Sub XDTyLeQD(mvt As Long)
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Kq1) SELECT MaVattu,CStr(MaVattu),First(TyLeQD) AS TL FROM DVTVattu " + IIf(mvt > 0, "WHERE MaVattu=" + CStr(mvt), "") + " GROUP BY MaVattu"
    ExecuteSQL5 "UPDATE Vattu INNER JOIN BaoCaoCP2 ON Vattu.MaSo=BaoCaoCP2.MaSo SET TyLeQD=Kq1"
End Sub

Public Function SoTonTheoChungTu(mk As Long, mtk As Long, mvt As Long) As Double
    Dim sodu As Double
    
    If OutCost = 0 Then Exit Function
    
    sodu = SelectSQL("SELECT Sum(SoPS2No-SoXuat) AS F1 FROM ChungTu WHERE ((MaLoai=1 AND MaKho=" + CStr(mk) + ") OR (MaLoai=4 AND MaNguon=" + CStr(mk) + ")) AND MaTKNo=" + CStr(mtk) + " AND MaVattu=" + CStr(mvt) + " AND SoPS2No>SoXuat")
    sodu = sodu + SelectSQL("SELECT Sum(Luong_0-SoXuat) AS F1 FROM VTDauNam WHERE MaSoKho=" + CStr(mk) + " AND MaTaiKhoan=" + CStr(mtk) + " AND MaVattu=" + CStr(mvt) + " AND Luong_0>SoXuat")
    
    SoTonTheoChungTu = sodu
End Function

Public Sub ktraxuatvattu()
    Dim rs As Recordset, tien As Double, loai As Integer, ms As Long, ps As Double, sql As String
      
    If OutCost = 1 Then
        ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=0"
                     
        sql = "SELECT ChungTu.MaSo,MaKho,MaTKCo,MaVattu,SoPS2Co,CT_ID FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE ABS(CT_ID)>2000000000 AND (MaLoai=2 OR MaLoai=4) AND MaVattu>0 AND TK_ID=" + CStr(TKVT_ID)
        Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rs.EOF
           ms = Abs(rs!CT_ID) - 2000000000
           If SelectSQL("SELECT MaSo AS F1 FROM VTDauNam WHERE MaSo=" + CStr(ms)) = 0 Then
               ms = SelectSQL("SELECT TOP 1 MaSo AS F1 FROM VTDauNam WHERE MaSoKho=" + CStr(rs!MaKho) + " AND MaTaiKhoan=" + CStr(rs!MaTkCo) + " AND MaVattu=" + CStr(rs!MaVattu))
               If ms > 0 Then
                   ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=SoXuat+" + DoiDau(rs!SoPS2Co) + " WHERE MaSo=" + CStr(ms)
                   ExecuteSQL5 "UPDATE ChungTu SET CT_ID=" + CStr(-(2000000000 - ms)) + " WHERE MaSo=" + CStr(rs!MaSo)
               End If
           Else
               ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=SoXuat+" + CStr(rs!SoPS2Co) + " WHERE MaSo=" + CStr(ms)
           End If
           rs.MoveNext
        Loop
                
        'ExecuteSQL5 "UPDATE ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo SET ChungTu.SoXuat=0 WHERE (MaLoai=1 OR MaLoai=4) AND MaVattu>0 AND TK_ID=" + CStr(TKVT_ID)
        'sql = "SELECT ChungTu.MaSo,MaKho, SoPS2Co,CT_ID FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE ABS(CT_ID)>0 AND ABS(CT_ID)<2000000000 AND TK_ID=" + CStr(TKVT_ID)
        'Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        'Do While Not rs.EOF
        '   ms = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNKH_ID) + " AND MaCT=" + CStr(Abs(rs!CT_ID)))
        '   ExecuteSQL5 "UPDATE ChungTu SET SoXuat=SoXuat+" + DoiDau(rs!SoPS2Co) + " WHERE MaSo=" + CStr(ms)
        '   rs.MoveNext
        'Loop
        
        'rs.Close
    End If
      
    If OutCost = 2 Then
        Set rs = DBKetoan.OpenRecordset("SELECT MaSoKho, MaTaiKhoan, MaVattu,Luong_12 AS SoDu FROM TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo", dbOpenSnapshot)
        Do While Not rs.EOF
            If rs!sodu <> SoTonTheoChungTu(rs!MaSoKho, rs!MaTaiKhoan, rs!MaVattu) Then
                ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=Luong_0 WHERE MaSoKho=" + CStr(rs!MaSoKho) + " AND MaTaiKhoan=" + CStr(rs!MaTaiKhoan) + " AND MaVattu=" + CStr(rs!MaVattu)
                ExecuteSQL5 "UPDATE ChungTu SET SoXuat=SoPS2No WHERE ((MaLoai=1 AND MaKho=" + CStr(rs!MaSoKho) + ") OR (MaLoai=4 AND MaNguon=" + CStr(rs!MaSoKho) + ")) AND MaTKNo=" + CStr(rs!MaTaiKhoan) + " AND MaVattu=" + CStr(rs!MaVattu)
                tien = 0
                Do While tien < rs!sodu
                    ms = SelectSQL("SELECT TOP 1 MaSo AS F1,SoPS2No AS F2 FROM ChungTu WHERE ((MaLoai=1 AND MaKho=" + CStr(rs!MaSoKho) + ") OR (MaLoai=4 AND MaNguon=" + CStr(rs!MaSoKho) + ")) AND MaTKNo=" + CStr(rs!MaTaiKhoan) + " AND MaVattu=" + CStr(rs!MaVattu) + " AND SoPS2No<=SoXuat AND SoPS2No>0 ORDER BY NgayGS DESC, MaCT DESC", ps)
                    If ms > 0 Then
                        If ps > rs!sodu - tien Then
                            ExecuteSQL5 "UPDATE ChungTu SET SoXuat=SoPS2No-" + DoiDau(rs!sodu - tien) + " WHERE MaSo=" + CStr(ms)
                            tien = rs!sodu
                        Else
                            ExecuteSQL5 "UPDATE ChungTu SET SoXuat=0 WHERE MaSo=" + CStr(ms)
                            tien = tien + ps
                        End If
                    Else
                        ms = SelectSQL("SELECT TOP 1 MaSo AS F1,Luong_0 AS F2 FROM VTDauNam WHERE MaSoKho=" + CStr(rs!MaSoKho) + " AND MaTaiKhoan=" + CStr(rs!MaTaiKhoan) + " AND MaVattu=" + CStr(rs!MaVattu) + " AND Luong_0<=SoXuat AND Luong_0<>0 ORDER BY NgayGS DESC, MaCT DESC", ps)
                        If ms > 0 Then
                            If ps > rs!sodu - tien Then
                                ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=SoXuat-" + DoiDau(rs!sodu - tien) + " WHERE MaSo=" + CStr(ms)
                                tien = rs!sodu
                            Else
                                ExecuteSQL5 "UPDATE VTDauNam SET SoXuat=Luong_0 WHERE MaSo=" + CStr(ms)
                                tien = tien + ps
                            End If
                        Else
                            Exit Do
                        End If
                    End If
                Loop
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Set rs = Nothing
End Sub

Public Function KtraDVT(mvt As Long, mdv As Long, tendv As String) As Boolean
    KtraDVT = SelectSQL("SELECT MaSo AS F1, DonVi AS F2 FROM DVTVattu WHERE MaSo=" + CStr(mdv) + " AND MaVattu=" + CStr(mvt), tendv) > 0
End Function

Public Sub TinhTonVT(thang As Integer, shtk As String)
    Dim rs As Recordset
    
    ExecuteSQL5 "UPDATE Vattu SET L=0,T=0"
    Set rs = DBKetoan.OpenRecordset("SELECT MaVattu, Sum(Luong_" + CStr(CThangDB(thang)) + ") AS L, Sum(Tien_" + CStr(CThangDB(thang)) + ") AS T FROM TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + shtk + "*' GROUP BY MaVattu", dbOpenSnapshot)
    Do While Not rs.EOF
        ExecuteSQL5 "UPDATE Vattu SET L=" + DoiDau(rs!L) + ",T=" + DoiDau(rs!T) + " WHERE MaSo=" + CStr(rs!MaVattu)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Sub TinhGXKBQ(tdau As Integer, tcuoi As Integer, shvt As String, tkno As String, Optional ktra As Integer = 0)
    Dim rs As Recordset, ms As Long, tien As Double, luong As Double, sql As String, i As Integer
    Dim mk As Long, mv As Long, mt As Long, thang As Integer, n As Date, tienx As Double, luongx As Double, tien2 As Double, tienx2 As Double, dgia As Double, dgia2 As Double
    Dim soct As Long, Counter As Long
        
    If shvt = "0" Then shvt = ""
        
    ExecuteSQL5 "UPDATE " + ChungTu2TKNC(0) + " SET MaTKNo=MaTKCo,MaTKTCNo=MaTKTCCo WHERE MaLoai=4 AND HethongTK.Cap=0 AND HethongTK.Loai=0 AND TK.Loai>0"
    ExecuteSQL5 "UPDATE ChungTu SET SoPS=Fix(IIF(SoPS>=0,0.5,-0.5)+SoPS), SoPS2Co=Fix(IIF(SoPS2Co>=0,0.5,-0.5)+SoPS2Co*" + CStr(Mask_N) + ")/" + CStr(Mask_N)
    
    If OutCost > 0 Then Exit Sub
    
    If Len(tkno) > 0 Then
        sql = "SELECT DISTINCTROW ChungTu.MaSo,ThangCT,NgayGS,MaCT,MaKho,MaVattu,MaTKCo,SoPS,SoPS2Co" + IIf(pGiaUSD > 0, ",PSUSD", "") + " FROM (" + ChungTu2TKNC(0) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE HethongTK.SoHieu LIKE '" + tkno + "*' AND (MaLoai=2 OR MaLoai=4) AND MaVattu>0 AND MaTKNo>0 AND TK.TK_ID=" + CStr(TKVT_ID) + " AND SoPS2Co>0 AND " + WThang("ThangCT", tdau, tcuoi) + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " ORDER BY MaKho,MaTKCo,MaVattu,ThangCT,NgayGS,ChungTu.MaCT"
    Else
        sql = "SELECT DISTINCTROW ChungTu.MaSo,ThangCT,NgayGS,MaCT,MaKho,MaVattu,MaTKCo,SoPS,SoPS2Co" + IIf(pGiaUSD > 0, ",PSUSD", "") + " FROM (" + ChungTu2TKNC(1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo WHERE (MaLoai=2 OR MaLoai=4) AND MaVattu>0 AND HethongTK.TK_ID=" + CStr(TKVT_ID) + " AND SoPS2Co>0 AND " + WThang("ThangCT", tdau, tcuoi) + IIf(Len(shvt) > 0, " AND Vattu.SoHieu='" + shvt + "'", "") + " ORDER BY MaKho,MaTKCo,MaVattu,ThangCT,NgayGS,ChungTu.MaCT"
    End If
    
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        soct = rs.recordCount
        rs.MoveFirst
    End If
    Do While Not rs.EOF
        Counter = Counter + 1
        If mk <> rs!MaKho Or mt <> rs!MaTkCo Or mv <> rs!MaVattu Or thang <> rs!ThangCT Then
            Idle
            n = rs!NgayGS
            mk = rs!MaKho
            mt = rs!MaTkCo
            mv = rs!MaVattu
            thang = rs!ThangCT
            luong = SoTonKho(ThangTruoc(thang), mk, mt, mv, tien, tien2)
            luong = luong + SoNhapKhoThang(mk, mt, mv, thang, thang, tienx, tienx2)
            tien = tien + tienx
            tien2 = tien2 + tienx2
            If luong <> 0 Then dgia = Abs(tien / luong) Else dgia = 0
            If luong <> 0 Then dgia2 = Abs(tien2 / luong) Else dgia2 = 0
        End If
                
        If Abs(luong - rs!SoPS2Co) < (1 / Mask_N) Then
            tienx = tien
            luong = 0
        Else
            luong = luong - rs!SoPS2Co
            tienx = RoundMoney(dgia * rs!SoPS2Co)
        End If
        tien = tien - tienx
        If tienx <> rs!sops Then ExecuteSQL5 "UPDATE ChungTu SET SoPS=" + DoiDau(tienx) + " WHERE MaSo=" + CStr(rs!MaSo)
        
        If pGiaUSD > 0 Then
            If luong - rs!SoPS2Co < (1 / Mask_N) Then
                tienx2 = tien2
            Else
                tienx2 = RoundMoney(dgia2 * rs!SoPS2Co)
            End If
            tien2 = tien2 - tienx2
            If tienx2 <> rs!PSUSD Then ExecuteSQL5 "UPDATE ChungTu SET PSUSD=" + DoiDau(tienx2) + " WHERE MaSo=" + CStr(rs!MaSo)
        End If
        If Counter Mod 100 = 0 Then HienThongBao "§· xö lý " + CStr(Fix(100 * Counter / soct)) + "% tæng sè chøng tõ", 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    If mv > 0 And ktra = 0 Then
        KiemTraTaiKhoan 1
        KiemTraVatTu 1
    End If
    HienThongBao "", 1
    'Beep
End Sub

Public Sub KtraCtuGV()
    Dim rs As Recordset, tongso As Integer, chiso As Integer
    
    SetSQL "QChungTuBanHang", "SELECT ChungTu.ThangCT, ChungTu.SoHieu, ChungTu.NgayCT, Sum(ChungTu.SoPS2Co) AS SumOfSoPS2Co, ChungTu.MaVattu, ChungTu.MaKho, Vattu.SoHieu, Vattu.TenVattu " _
        & " FROM (ChungTu INNER JOIN HeThongTK ON ChungTu.MaTKCo = HeThongTK.MaSo) INNER JOIN Vattu ON ChungTu.MaVattu = Vattu.MaSo " _
        & " Where (((chungtu.maloai) = 8) And ((HeThongTK.tk_id) = 5000)) " _
        & " GROUP BY ChungTu.ThangCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.MaVattu, ChungTu.MaKho, Vattu.SoHieu, Vattu.TenVattu, ChungTu.MaCT"
    SetSQL "QChungTuGiaVon", "SELECT ChungTu.ThangCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.MaKho, Sum(ChungTu.SoPS2Co) AS SumOfSoPS2Co, ChungTu.MaTKNo, ChungTu.MaVattu " _
        & " FROM ChungTu INNER JOIN HeThongTK ON ChungTu.MaTKNo = HeThongTK.MaSo" _
        & " WHERE (((ChungTu.MaLoai)=2) AND ((HeThongTK.SoHieu) Like '632*'))" _
        & " GROUP BY ChungTu.MaCT, ChungTu.ThangCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.MaKho, ChungTu.MaTKNo, ChungTu.MaVattu"
    SetSQL "Mientru", "SELECT QChungTuBanHang.MaVattu" _
        & " FROM QChungTuBanHang INNER JOIN QChungtuGiaVon ON (QChungTuBanHang.ChungTu.SoHieu+'GV'=QChungTuGiaVon.SoHieu) AND (QChungTuBanHang.MaKho = QChungtuGiaVon.MaKho) AND (QChungTuBanHang.MaVattu = QChungtuGiaVon.MaVattu) AND (QChungTuBanHang.SumOfSoPS2Co<>QChungTuGiaVon.SumOfSoPS2Co)" _
        & " GROUP BY QChungTuBanHang.MaVattu"
    
    Set rs = DBKetoan.OpenRecordset("MienTru", dbOpenSnapshot)
    On Error Resume Next
    rs.MoveLast
    tongso = rs.recordCount
    rs.MoveFirst
    On Error GoTo 0
    Do While Not rs.EOF
        chiso = chiso + 1
        TinhGVBH NgayDauThang(pNamTC, pThangDauKy), NgayCuoiNam, 1, rs!MaVattu, 1
        HienThongBao "§· xö lý " + CStr(chiso) + " danh ®iÓm vËt t­, " + CStr(Fix(100 * chiso / tongso)) + " % tæng sè cÇn kiÓm tra!", 1
        rs.MoveNext
    Loop
    
    If rs.recordCount > 0 Then
        KiemTraTaiKhoan 1
        KiemTraVatTu 1
    End If
    
    rs.Close
    Set rs = Nothing
    
    HienThongBao "", 1
End Sub
