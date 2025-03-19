Attribute VB_Name = "modTaikhoan"
Option Explicit

Public pXuLyTKTC As Integer
'======================================================================================
' Hµm tr¶ vÒ m· sè tai kho¶n tõ sè hiÖu vµ m· nguyªn tÖ
'======================================================================================
Public Function MaTKNguyenTe(shnt As String, MaNT As Long)
    Dim sql As String
    
    sql = "SELECT MaSo AS F1 FROM HethongTK WHERE SoHieu = '" + shnt + "' AND MaNT " + IIf(MaNT <= 0, " <= 0", " = " + CStr(MaNT))
    MaTKNguyenTe = SelectSQL(sql)
End Function
'======================================================================================
' Hµm tr¶ vÒ sè d­ tµi kho¶n cuèi th¸ng
'======================================================================================
Public Sub SoDuTK(mtk As Long, thang As Integer, duno As Double, duco As Double, Optional nt As Double)
    Dim sql As String
        
    sql = "SELECT DuNo_" + CStr(CThangDB(thang)) + " As F1, DuCo_" + CStr(CThangDB(thang)) + " As F2 FROM HethongTK WHERE MaSo=" + CStr(mtk)
    duno = SelectSQL(sql, duco)
    
    sql = "SELECT DuNT_" + CStr(CThangDB(thang)) + " As F1 FROM HethongTK WHERE MaSo=" + CStr(mtk)
    nt = SelectSQL(sql)
End Sub
'======================================================================================
' Hµm tr¶ vÒ ký hiÖu nguyªn tÖ tõ m· sè
'======================================================================================
Public Function SoHieuNT(nt As Long) As String
    Dim sql As String
    
    sql = "SELECT DISTINCTROW KyHieu AS F1 FROM NguyenTe WHERE MaSo=" + CStr(nt)
    SoHieuNT = SelectSQL(sql)
End Function
'======================================================================================
' Hµm tr¶ vÒ sè d­ tµi kho¶n cuèi ngµy, chØ cho tµi kho¶n chi tiÕt
'======================================================================================
Public Sub SoDuTKNgay(mtk As Long, ngay As Date, duno As Double, duco As Double, Optional dunt As Double)
    Dim sql As String, ps As Double
    
    ' Sè ®Çu th¸ng
    sql = "SELECT DuNo_0 As F1, DuCo_0 As F2,DuNT_0 AS F3 FROM HethongTK WHERE MaSo=" + CStr(mtk)
    duno = SelectSQL(sql, duco, dunt)
    If duno - duco > 0 Then dunt = Abs(dunt) Else dunt = -Abs(dunt)
    
    ' Sè ph¸t sinh nî
    sql = "SELECT Sum(SoPS) As F1,Sum(SoPS2No) AS F2 FROM ChungTu " _
        & "WHERE MaTKNo = " + CStr(mtk) + " AND NgayGS <= #" + Format(ngay, Mask_DB) + "#"
    duno = duno + SelectSQL(sql, ps)
    dunt = dunt + ps
    
    ' Sè ph¸t sinh cã
    sql = "SELECT Sum(SoPS) As F1,Sum(SoPS2Co) AS F2 FROM ChungTu " _
        & "WHERE MaTKCo = " + CStr(mtk) + " AND NgayGS <= #" + Format(ngay, Mask_DB) + "#"
    duco = duco + SelectSQL(sql, ps)
    dunt = dunt - ps
    
    If duno > duco Then
        duno = duno - duco
        duco = 0
    Else
        duco = duco - duno
        duno = 0
    End If
    dunt = Abs(dunt)
End Sub

Public Sub SoDuTKNgaySH(shtk As String, ngay As Date, duno As Double, duco As Double, dunt As Double)
    Dim sql As String, ps As Double
    
    ' Sè ®Çu th¸ng
    sql = "SELECT Sum(DuNo_0) As F1, Sum(DuCo_0) As F2, Sum(DuNT_0) AS F3 FROM HethongTK WHERE SoHieu='" + shtk + "'"
    duno = SelectSQL(sql, duco, dunt)
    If duno - duco > 0 Then dunt = Abs(dunt) Else dunt = -Abs(dunt)
    
    ' Sè ph¸t sinh nî
    sql = "SELECT Sum(SoPS) As F1,Sum(SoPS2No) AS F2 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu = '" + shtk + "' AND NgayGS <= #" + Format(ngay, Mask_DB) + "#"
    duno = duno + SelectSQL(sql, ps)
    dunt = dunt + ps
    
    ' Sè ph¸t sinh cã
    sql = "SELECT Sum(SoPS) As F1,Sum(SoPS2Co) AS F2 FROM " + ChungTu2TKNC(1) + " WHERE HethongTK.SoHieu = '" + shtk + "' AND NgayGS <= #" + Format(ngay, Mask_DB) + "#"
    duco = duco + SelectSQL(sql, ps)
    dunt = dunt - ps
    
    If duno > duco Then
        duno = duno - duco
        duco = 0
    Else
        duco = duco - duno
        duno = 0
    End If
    dunt = Abs(dunt)
End Sub
'====================================================================================================
' Thñ tôc in sæ c¸i tµi kho¶n
'====================================================================================================
Public Function InSocaiTk(taikhoan As ClsTaikhoan, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, thongbao As Boolean, doiung As String, VV As Long, loai As Integer, Optional nn As Integer = 0, Optional mdt1 As Long = 0, Optional mdt2 As Long = 0, Optional mdt3 As Long = 0) As Boolean
    Dim dkn As Double, dkc As Double, dknt As Double, ps As Boolean, sqlw As String
    Dim psn As Double, psc As Double
    
With taikhoan
    If tcuoi > 0 Then
        ps = .TkCoPS(tdau, tcuoi)
        .SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
        sqlw = WThang("ThangCT", tdau, tcuoi)
    Else
        ps = .TkCoPSN(ndau, ncuoi)
        .SoDuNgay ndau - 1, dkn, dkc, dknt
        sqlw = WNgay("NgayGS", ndau, ncuoi)
    End If
    
    ' In bao cao
    If dkn <> 0 Or dkc <> 0 Or ps Then
        
        If ps Then
            SetSQL "QSocai", "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.ThangCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.GhiChu, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.MaTKTCNo, ChungTu.MaTKTCCo, " + IIf(.kieu < 0, "IIF(HethongTK.SoHieu LIKE '" + .SoHieu + "*','0','1')", "IIF(HethongTK.SoHieu LIKE '" + .SoHieu + "*','1','0')") + "+Cstr(10+ChungTu.ThangCT)+ChungTu.SoHieu AS SH1" _
                & " FROM HeThongTK AS HeThongTK_3 RIGHT JOIN (HeThongTK AS HeThongTK_2 RIGHT JOIN (HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo) ON HeThongTK_2.MaSo = ChungTu.MaTKNo) ON HeThongTK_3.MaSo = ChungTu.MaTKCo" _
                & " Where SoPS<>0 AND ((HethongTK.SoHieu LIKE '" + taikhoan.SoHieu + "*'" + IIf(Len(doiung) = 0, "", " And HethongTK_3.SoHieu LIKE '" + doiung + "*'") _
                + ") Or (HethongTK_1.SoHieu LIKE '" + taikhoan.SoHieu + "*'" + IIf(Len(doiung) = 0, "", " And HethongTK_2.SoHieu LIKE '" + doiung + "*'") _
                + ")) And " + sqlw + IIf(VV > 1, " AND (ChungTu.MaDT = " + CStr(VV) + ")", "") + IIf(mdt1 > 0, " AND (ChungTu.MaDT1 = " + CStr(mdt1) + ")", "") + IIf(mdt2 > 0, " AND (ChungTu.MaDT2 = " + CStr(mdt2) + ")", "") + IIf(mdt3 > 0, " AND (ChungTu.MaDT3 = " + CStr(mdt3) + ")", "") + " AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) ORDER BY ThangCT,ChungTu.NgayGS," + IIf(.kieu < 0, "IIF(HethongTK.SoHieu LIKE '" + .SoHieu + "*','0','1')", "IIF(HethongTK.SoHieu LIKE '" + .SoHieu + "*','1','0')") + "+Cstr(10+ChungTu.ThangCT)+ChungTu.SoHieu"
        Else
            SetSQL "QSocai", "SELECT DISTINCTROW 0 AS MaCT,0 AS ThangCT, '' AS SoHieu, Null AS NgayCT, Null AS NgayGS, '' AS DienGiai, 0 AS SoPS, '' AS GhiChu, '' AS SoHieu1, '' AS SoHieu2, 0 AS MaTKTCNo, 0 AS MaTKTCCo" _
                & " FROM ChungTu Where (chungtu.MaCT = 0)"
            frmMain.Rpt.Formulas(9) = "DuCk=" + DoiDau(IIf(.kieu < 0, dkn - dkc, dkc - dkn))
        End If
        
        frmMain.Rpt.Formulas(3) = "SoHieuTK='" + .SoHieu + "'"
        frmMain.Rpt.Formulas(4) = "NoDk=" + DoiDau(dkn - dkc)
        frmMain.Rpt.Formulas(5) = "Kieu=" + CStr(.kieu)
        frmMain.Rpt.Formulas(8) = "TenTk = IF PageNumber() = 1 THEN '" + .SoHieu + " - " + IIf(nn > 0, .TenE, .Ten) + IIf(VV > 1, " - VV: " + TenVV(VV), "") + "'"
        If tcuoi > 0 Then
            frmMain.Rpt.Formulas(6) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
            RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
            psn = SoPSTK(taikhoan.SoHieu, pThangDauKy, tcuoi, -1)
            psc = SoPSTK(taikhoan.SoHieu, pThangDauKy, tcuoi, 1)
            If VV > 1 Then
                taikhoan.SoDuTK tcuoi, dkn, dkc, dknt
                frmMain.Rpt.Formulas(10) = "DuCk = " + DoiDau(IIf(taikhoan.kieu < 0, dkn - dkc, dkc - dkn))
            End If
        Else
            frmMain.Rpt.Formulas(6) = "ThoiGian = '" + ThoiGianN(ndau, ncuoi, nn) + "'"
            RptSetDate ncuoi, nn
            taikhoan.SoPhatSinhN NgayDauThang(pNamTC, pThangDauKy), ncuoi, psn, dkn, psc, dkc
            If VV > 1 Then
                taikhoan.SoDuNgay ncuoi, dkn, dkc, dknt
                frmMain.Rpt.Formulas(10) = "DuCk = " + DoiDau(IIf(taikhoan.kieu < 0, dkn - dkc, dkc - dkn))
            End If
        End If
        frmMain.Rpt.Formulas(11) = "NoLK = " + DoiDau(psn)
        frmMain.Rpt.Formulas(12) = "CoLK = " + DoiDau(psc)
        InSocaiTk = True
'        If loai = 0 Then
'            If Left(taikhoan.SoHieu, 3) = "111" And tcuoi = 0 Then
'                frmMain.Rpt.ReportFileName = "SOQUY.RPT"
'            Else
'                frmMain.Rpt.ReportFileName = "SOCAI.RPT"
'                frmMain.Rpt.Formulas(13) = "LastRow=" + CStr(GetLastRow("QSocai", "MaCT"))
'            End If
'        Else
'            If tdau = tcuoi Then
'                frmMain.Rpt.ReportFileName = "SOCAI2.RPT"
'            Else
              '  frmMain.Rpt.ReportFileName = "SOCAI3.RPT"
 '           End If
  '      End If
  ' them theo cai tien
  If loai = 0 Then
          frmMain.Rpt.ReportFileName = "SOCAI3.RPT"
  Else
          frmMain.Rpt.ReportFileName = "SOCAI.RPT"
          frmMain.Rpt.Formulas(13) = "LastRow=" + CStr(GetLastRow("QSocai", "MaCT"))
  End If
  
    Else
        If thongbao Then MsgBox "Kh«ng cã sè ®Çu kú hoÆc ph¸t sinh !", vbInformation, App.ProductName
        InSocaiTk = False
    End If
End With
End Function
Public Function InSocaiTk_them(taikhoan As ClsTaikhoan, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, thongbao As Boolean, doiung As String, VV As Long, loai As Integer, Optional nn As Integer = 0, Optional mdt1 As Long = 0, Optional mdt2 As Long = 0, Optional mdt3 As Long = 0) As Boolean
    Dim dkn As Double, dkc As Double, dknt As Double, ps As Boolean, sqlw As String
    Dim psn As Double, psc As Double
    
With taikhoan
    If tcuoi > 0 Then
        ps = .TkCoPS(tdau, tcuoi)
        .SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
        sqlw = WThang("ThangCT", tdau, tcuoi)
    Else
        ps = .TkCoPSN(ndau, ncuoi)
        .SoDuNgay ndau - 1, dkn, dkc, dknt
        sqlw = WNgay("NgayGS", ndau, ncuoi)
    End If
    
    ' In bao cao
    If dkn <> 0 Or dkc <> 0 Or ps Then
        
        If ps Then
            SetSQL "QSocai", "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.ThangCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.GhiChu, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.MaTKTCNo, ChungTu.MaTKTCCo, " + IIf(.kieu < 0, "IIF(HethongTK.SoHieu LIKE '" + .SoHieu + "*','0','1')", "IIF(HethongTK.SoHieu LIKE '" + .SoHieu + "*','1','0')") + "+Cstr(10+ChungTu.ThangCT)+ChungTu.SoHieu AS SH1" _
                & " FROM HeThongTK AS HeThongTK_3 RIGHT JOIN (HeThongTK AS HeThongTK_2 RIGHT JOIN (HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo) ON HeThongTK_2.MaSo = ChungTu.MaTKNo) ON HeThongTK_3.MaSo = ChungTu.MaTKCo" _
                & " Where SoPS<>0 AND ((HethongTK.SoHieu LIKE '" + taikhoan.SoHieu + "*'" + IIf(Len(doiung) = 0, "", " And HethongTK_3.SoHieu LIKE '" + doiung + "*'") _
                + ") Or (HethongTK_1.SoHieu LIKE '" + taikhoan.SoHieu + "*'" + IIf(Len(doiung) = 0, "", " And HethongTK_2.SoHieu LIKE '" + doiung + "*'") _
                + ")) And " + sqlw + IIf(VV > 1, " AND (ChungTu.MaDT = " + CStr(VV) + ")", "") + IIf(mdt1 > 0, " AND (ChungTu.MaDT1 = " + CStr(mdt1) + ")", "") + IIf(mdt2 > 0, " AND (ChungTu.MaDT2 = " + CStr(mdt2) + ")", "") + IIf(mdt3 > 0, " AND (ChungTu.MaDT3 = " + CStr(mdt3) + ")", "") + " AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) ORDER BY ThangCT,ChungTu.NgayGS," + IIf(.kieu < 0, "IIF(HethongTK.SoHieu LIKE '" + .SoHieu + "*','0','1')", "IIF(HethongTK.SoHieu LIKE '" + .SoHieu + "*','1','0')") + "+Cstr(10+ChungTu.ThangCT)+ChungTu.SoHieu"
        Else
            SetSQL "QSocai", "SELECT DISTINCTROW 0 AS MaCT,0 AS ThangCT, '' AS SoHieu, Null AS NgayCT, Null AS NgayGS, '' AS DienGiai, 0 AS SoPS, '' AS GhiChu, '' AS SoHieu1, '' AS SoHieu2, 0 AS MaTKTCNo, 0 AS MaTKTCCo" _
                & " FROM ChungTu Where (chungtu.MaCT = 0)"
            frmMain.Rpt.Formulas(9) = "DuCk=" + DoiDau(IIf(.kieu < 0, dkn - dkc, dkc - dkn))
        End If
        
        frmMain.Rpt.Formulas(3) = "SoHieuTK='" + .SoHieu + "'"
        frmMain.Rpt.Formulas(4) = "NoDk=" + DoiDau(dkn - dkc)
        frmMain.Rpt.Formulas(5) = "Kieu=" + CStr(.kieu)
        frmMain.Rpt.Formulas(8) = "TenTk = IF PageNumber() = 1 THEN '" + .SoHieu + " - " + IIf(nn > 0, .TenE, .Ten) + IIf(VV > 1, " - VV: " + TenVV(VV), "") + "'"
        If tcuoi > 0 Then
            frmMain.Rpt.Formulas(6) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
            RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
            psn = SoPSTK(taikhoan.SoHieu, pThangDauKy, tcuoi, -1)
            psc = SoPSTK(taikhoan.SoHieu, pThangDauKy, tcuoi, 1)
            If VV > 1 Then
                taikhoan.SoDuTK tcuoi, dkn, dkc, dknt
                frmMain.Rpt.Formulas(10) = "DuCk = " + DoiDau(IIf(taikhoan.kieu < 0, dkn - dkc, dkc - dkn))
            End If
        Else
            frmMain.Rpt.Formulas(6) = "ThoiGian = '" + ThoiGianN(ndau, ncuoi, nn) + "'"
            RptSetDate ncuoi, nn
            taikhoan.SoPhatSinhN NgayDauThang(pNamTC, pThangDauKy), ncuoi, psn, dkn, psc, dkc
            If VV > 1 Then
                taikhoan.SoDuNgay ncuoi, dkn, dkc, dknt
                frmMain.Rpt.Formulas(10) = "DuCk = " + DoiDau(IIf(taikhoan.kieu < 0, dkn - dkc, dkc - dkn))
            End If
        End If
        frmMain.Rpt.Formulas(11) = "NoLK = " + DoiDau(psn)
        frmMain.Rpt.Formulas(12) = "CoLK = " + DoiDau(psc)
        InSocaiTk_them = True
'        If loai = 0 Then
'            If Left(taikhoan.SoHieu, 3) = "111" And tcuoi = 0 Then
'                frmMain.Rpt.ReportFileName = "SOQUY.RPT"
'            Else
'                frmMain.Rpt.ReportFileName = "SOCAI.RPT"
'                frmMain.Rpt.Formulas(13) = "LastRow=" + CStr(GetLastRow("QSocai", "MaCT"))
'            End If
'        Else
'            If tdau = tcuoi Then
'                frmMain.Rpt.ReportFileName = "SOCAI2.RPT"
'            Else
              '  frmMain.Rpt.ReportFileName = "SOCAI3.RPT"
 '           End If
  '      End If
  ' them theo cai tien
 
   frmMain.Rpt.ReportFileName = "SOQUY.RPT"
   Else
        If thongbao Then MsgBox "Kh«ng cã sè ®Çu kú hoÆc ph¸t sinh !", vbInformation, App.ProductName
        InSocaiTk_them = False
    End If
End With
End Function

'====================================================================================================
' Thñ tôc in sæ chi tiÕt
'====================================================================================================
Public Function InSoChitiet(taikhoan As ClsTaikhoan, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, thongbao As Boolean, doiung As String, VV As Long, Optional loai As Integer = 0, Optional nn As Integer = 0, Optional mdt1 As Long = 0, Optional mdt2 As Long = 0, Optional mdt3 As Long = 0) As Boolean
    Dim dkn As Double, dkc As Double, dknt As Double, ps As Boolean, sqlw As String, psn As Double, psc As Double
    Dim rs_tk As Recordset, sql As String, inso As Boolean, i As Integer, st As String, j As Integer
    
    If tcuoi > 0 Then
        ps = taikhoan.TkCoPS(tdau, tcuoi)
        taikhoan.SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
        sqlw = WThang("ThangCT", tdau, tcuoi)
    Else
        ps = taikhoan.TkCoPSN(ndau, ncuoi)
        sqlw = WNgay("NgayGS", ndau, ncuoi)
        taikhoan.SoDuNgay ndau - 1, dkn, dkc, dknt
    End If
    If dkn = 0 And dkc = 0 And (Not ps) Then GoTo KhongIn
    ' In bao cao
    
    If ps Then
        sql = "((HethongTK_2.SoHieu LIKE '" + taikhoan.SoHieu + "*'" + IIf(Len(doiung) > 0, " And HethongTK_3.SoHieu LIKE '" + doiung + "*'", "") _
            + ") Or (HethongTK_3.SoHieu LIKE '" + taikhoan.SoHieu + "*'" + IIf(Len(doiung) > 0, " And HethongTK_2.SoHieu LIKE '" + doiung + "*'", "") + "))"
        
        SetSQL "QChitiet", "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.SoPS2No, ChungTu.SoPS2Co, ChungTu.GhiChu, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.MaTKNo, ChungTu.MaTKCo, HethongTK_2.SoHieu As SHNo, HethongTK_3.SoHieu As SHCo,ThangCT" _
            & " FROM HeThongTK AS HeThongTK_3 RIGHT JOIN (HeThongTK AS HeThongTK_2 RIGHT JOIN (HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo) ON HeThongTK_2.MaSo = ChungTu.MaTKNo) ON HeThongTK_3.MaSo = ChungTu.MaTKCo" _
            & " Where SoPS<>0 AND (" + sqlw + " And (" + sql + "))" + IIf(VV > 1, " AND (ChungTu.MaDT = " + CStr(VV) + ")", "") + IIf(mdt1 > 0, " AND (ChungTu.MaDT1 = " + CStr(mdt1) + ")", "") + IIf(mdt2 > 0, " AND (ChungTu.MaDT2 = " + CStr(mdt2) + ")", "") + IIf(mdt3 > 0, " AND (ChungTu.MaDT3 = " + CStr(mdt3) + ")", "") + " AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) ORDER BY ChungTu.NgayGS," + IIf(taikhoan.kieu < 0, "IIF(HethongTK_2.SoHieu LIKE '" + taikhoan.SoHieu + "*','0','1')", "IIF(HethongTK_3.SoHieu LIKE '" + taikhoan.SoHieu + "*','0','1')") + "+Cstr(10+ChungTu.ThangCT)+ChungTu.SoHieu, MaCT"
    Else
        SetSQL "QChitiet", "SELECT DISTINCTROW 0 AS MaCT, '' AS SoHieu, Null AS NgayCT, Null AS NgayGS, '' AS DienGiai, 0 AS SoPS, 0 AS SoPS2No, 0 AS SoPS2Co, '' AS GhiChu, '' AS SoHieu0, '' AS SoHieu1, 0 AS MaTKNo, 0 AS MaTKCo, '' As SHNo, '' As SHCo, 0 AS ThangCT" _
            & " FROM ChungTu Where (chungtu.MaCT = 0)"
        frmMain.Rpt.Formulas(9) = "DuCk=" + DoiDau(IIf(taikhoan.kieu < 0, dkn - dkc, dkc - dkn))
    End If
       
    frmMain.Rpt.Formulas(3) = "SoHieuTk='" + taikhoan.SoHieu + "'"
    frmMain.Rpt.Formulas(4) = "NoDk=" + DoiDau(dkn - dkc)
    frmMain.Rpt.Formulas(5) = "Kieu=" + CStr(taikhoan.kieu)
    frmMain.Rpt.Formulas(8) = "TenTk = IF PageNumber() = 1 THEN '" + taikhoan.SoHieu + " - " + IIf(nn > 0, taikhoan.TenE, taikhoan.Ten) + IIf(VV > 1, " - VV: " + TenVV(VV), "") + "'"
    
    If tcuoi > 0 Then
        RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
        frmMain.Rpt.Formulas(6) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
        If VV > 1 Then
            taikhoan.SoDuTK tcuoi, dkn, dkc, dknt
            frmMain.Rpt.Formulas(10) = "DuCk = " + DoiDau(IIf(taikhoan.kieu < 0, dkn - dkc, dkc - dkn))
        End If
        psn = SoPSTK(taikhoan.SoHieu, pThangDauKy, tcuoi, -1)
        psc = SoPSTK(taikhoan.SoHieu, pThangDauKy, tcuoi, 1)
        
        If loai = 0 Then
            frmMain.Rpt.ReportFileName = "CTTK.RPT"
            frmMain.Rpt.Formulas(13) = "LastRow=" + CStr(GetLastRow("QChitiet", "MaCT"))
        Else
            frmMain.Rpt.ReportFileName = "CTTK2.RPT"
        End If
    Else
        RptSetDate ncuoi, nn
        frmMain.Rpt.Formulas(6) = "ThoiGian = '" + ThoiGianN(ndau, ncuoi, nn) + "'"
        If VV > 1 Then
            taikhoan.SoDuNgay ncuoi, dkn, dkc, dknt
            frmMain.Rpt.Formulas(10) = "DuCk = " + DoiDau(IIf(taikhoan.kieu < 0, dkn - dkc, dkc - dkn))
        End If
        taikhoan.SoPhatSinhN NgayDauThang(pNamTC, pThangDauKy), ncuoi, psn, dkn, psc, dkc
        
        Select Case Left(taikhoan.SoHieu, 3)
            Case "112":
                frmMain.Rpt.ReportFileName = "SOTG.RPT"
                frmMain.Rpt.Formulas(8) = "TenTk = IF PageNumber() = 1 THEN '" + taikhoan.GhiChu + "'"
            Case "311", "341":
                frmMain.Rpt.ReportFileName = "SOTV.RPT"
                frmMain.Rpt.Formulas(8) = "TenTk = IF PageNumber() = 1 THEN '" + taikhoan.GhiChu + "'"
            Case Else:
                frmMain.Rpt.Formulas(8) = "TenTk = IF PageNumber() = 1 THEN '" + taikhoan.SoHieu + " - " + taikhoan.Ten + IIf(VV > 1, " - VV: " + TenVV(VV), "") + "'"
                frmMain.Rpt.ReportFileName = "CTTK.RPT"
                frmMain.Rpt.Formulas(13) = "LastRow=" + CStr(GetLastRow("QChitiet", "MaCT"))
        End Select
    End If
    frmMain.Rpt.Formulas(11) = "NoLK = " + DoiDau(psn)
    frmMain.Rpt.Formulas(12) = "CoLK = " + DoiDau(psc)
    
    InSoChitiet = True
    Exit Function
KhongIn:
    InSoChitiet = False
    If thongbao Then MsgBox "Kh«ng cã sè ®Çu kú hoÆc ph¸t sinh !", vbInformation, App.ProductName
End Function

'======================================================================================
' Hµm tr¶ vÒ m· sè, tªn TK tõ sè hiÖu
'======================================================================================
Public Function tentk(sh As String, mtk As Long, Optional CoCon As Integer) As String
    Dim rs_tk As Recordset
    If mtk > 0 Then
        Set rs_tk = DBKetoan.OpenRecordset("SELECT SoHieu,Ten,TKCon FROM HethongTK WHERE MaSo = " + CStr(mtk), dbOpenSnapshot)
        tentk = rs_tk!Ten
        sh = rs_tk!SoHieu
        CoCon = rs_tk!tkcon
    Else
        Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo,Ten,TKCon FROM HethongTK WHERE SoHieu='" + sh + "'", dbOpenSnapshot)
        If rs_tk.RecordCount > 0 Then
            mtk = rs_tk!MaSo
            tentk = rs_tk!Ten
            CoCon = rs_tk!tkcon
        Else
            mtk = 0
            tentk = ""
        End If
    End If
    rs_tk.Close
    Set rs_tk = Nothing
End Function
'======================================================================================
' Hµm tr¶ vÒ m· sè, tªn NTK tõ m· sè
'======================================================================================
Public Function TenNT(mnt As Long) As String
    Dim sql As String
    
    If mnt > 0 Then
        sql = "SELECT KyHieu As F1 FROM NguyenTe WHERE MaSo=" + CStr(mnt)
        TenNT = SelectSQL(sql)
    Else
        TenNT = ""
    End If
End Function

'======================================================================================
' Hµm tr¶ vÒ m· sè tµi kho¶n chitiÕt cña mét tµi kho¶n
'======================================================================================
Public Function TKChitiet(mtk As Long) As Long
    Dim sql As String, MaCT As Long, X As Long
    MaCT = mtk
    Do While True
        sql = "SELECT MaSo AS F1 FROM HethongTK WHERE TkCha0 = " + CStr(MaCT)
        X = SelectSQL(sql)
        If X > 0 Then
            MaCT = SelectSQL(sql)
        Else
            Exit Do
        End If
    Loop
    TKChitiet = MaCT
End Function

'======================================================================================
' Hµm tr¶ vÒ m· sè tµi kho¶n tµi chÝnh cña mét tµi kho¶n
'======================================================================================
Public Function TKTaiChinh(mtk As Long) As Long
    Dim sql As String, ms As Long
    
    If mtk > 0 Then
        sql = "SELECT MaTC AS F1 FROM HethongTK WHERE MaSo = " + CStr(mtk)
        ms = SelectSQL(sql)
        TKTaiChinh = IIf(ms > 0, ms, mtk)
    Else
        TKTaiChinh = 0
    End If
End Function

Public Function TenVV(VV As Long) As String
Dim sql As String
    sql = "SELECT DISTINCTROW DienGiai AS F1 FROM DoituongCT WHERE MaSo = " + CStr(VV)
    TenVV = SelectSQL(sql)
End Function
'====================================================================================================
' Thu tuc xac dinh cac cap tai khoan
'====================================================================================================
Public Sub XDCapTK(shtk As String)
Dim rs As Recordset
    
    ExecuteSQL5 "UPDATE HethongTK SET CapDuoi = 0 WHERE SoHieu LIKE '" + shtk + "*'"
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW TKCha0" _
      & " FROM HethongTK WHERE (TKCon = 1) AND (SoHieu LIKE '" + shtk + "*') GROUP BY TKCha0", dbOpenSnapshot)

    Do While Not rs.EOF
      ExecuteSQL5 "UPDATE HethongTK SET CapDuoi = 1 WHERE TKCha0 = " + CStr(rs!TkCha0)
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

'====================================================================================================
' Thu tuc cong don so du ket chuyen
'====================================================================================================
Public Sub CongDonKC(cap0 As Boolean)
    Dim rs_taikhoan As Recordset, sql As String
    Dim rs_sub As Recordset
    
      Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT HethongTK.MaSo FROM HethongTK" _
          & " WHERE Cap>0 AND TKCon>0 ORDER BY SoHieu DESC", dbOpenSnapshot, dbForwardOnly)
      Do While Not rs_taikhoan.EOF
      
          Set rs_sub = DBKetoan.OpenRecordset("SELECT DISTINCTROW Sum(HethongTK.KC_N) AS KCN, Sum(HethongTK.KC_C) AS KCC, Sum(DuNo) AS DN, Sum(DuCo) AS DC FROM HethongTK WHERE HethongTK.TkCha0 = " + CStr(rs_taikhoan!MaSo), dbOpenSnapshot)
         ' ExecuteSQL5 "UPDATE HethongTK SET KC_N = " + IIf(IsNull(DoiDau(rs_sub!KCN) = True), 0, DoiDau(rs_sub!KCN)) + ", KC_C = " + IIf(IsNull(DoiDau(rs_sub!KCC) = True), 0, (DoiDau(rs_sub!KCC))) + ", DuNo = " + IIf(IsNull(DoiDau(rs_sub!dn) = True), 0, DoiDau(rs_sub!dn)) + ", DuCo = " + IIf(IsNull(DoiDau(rs_sub!DC) = True), 0, DoiDau(rs_sub!DC)) + " WHERE MaSo = " + CStr(rs_taikhoan!MaSo)
          rs_taikhoan.MoveNext
      Loop
    
    If cap0 Then
        Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT HethongTK.MaSo FROM HethongTK WHERE Cap=0", dbOpenSnapshot, dbForwardOnly)
        Do While Not rs_taikhoan.EOF
            Set rs_sub = DBKetoan.OpenRecordset("SELECT DISTINCTROW Sum(HethongTK.KC_N) AS KCN, Sum(HethongTK.KC_C) AS KCC, Sum(DuNo) AS DN, Sum(DuCo) AS DC FROM HethongTK WHERE HethongTK.TkCha0 = " + CStr(rs_taikhoan!MaSo), dbOpenSnapshot)
            
            ExecuteSQL5 "UPDATE HethongTK SET KC_N = " + DoiDau(rs_sub!KCN) + ", KC_C = " + DoiDau(rs_sub!KCC) _
                + ", DuNo = " + DoiDau(rs_sub!dn) + ", DuCo = " + DoiDau(rs_sub!DC) + " WHERE MaSo = " + CStr(rs_taikhoan!MaSo)
            rs_taikhoan.MoveNext
        Loop
    End If
    
    rs_taikhoan.Close
    rs_sub.Close
    Set rs_taikhoan = Nothing
    Set rs_sub = Nothing
End Sub

Public Function MaTKFix(db As Database, shtk As String, shct As String, tenct As String) As Long
    Dim sh As String, m As Long, TK As ClsTaikhoan, tentk As String
        
    sh = shtk
    Do Until m > 0
        m = SelectSQL("SELECT TOP 1 MaSo AS F1 FROM HethongTK WHERE SoHieu='" + sh + "' ORDER BY SoHieu")
        If m = 0 Then
            sh = SelectSQLDB(db, "SELECT TK.SoHieu AS F1, HethongTK.Ten AS F2 FROM HethongTK INNER JOIN HethongTK AS TK ON HethongTK.TKCha0=TK.MaSo WHERE HethongTK.SoHieu='" + sh + "'", tentk)
            If sh = "0" Or sh = "" Then Exit Function
            If shct = "#" Then
                shct = Right(shtk, Len(shtk) - Len(sh))
                tenct = tentk
            End If
        End If
    Loop
    
    If Len(shct) > 0 And shct <> "#" And Left(shtk, 3) <> "211" And Left(shtk, 3) <> "214" Then
        Set TK = New ClsTaikhoan
        TK.InitTaikhoanSohieu sh, shct
        If TK.MaSo > 0 Then
            m = TK.MaSo
        Else
            TK.InitTaikhoanMaSo m
            m = TK.ThemTKCon(shct, tenct, "...", 0, 0)
        End If
        Set TK = Nothing
    End If
    
    MaTKFix = m
End Function

Public Function SoPSTK(sh As String, tdau As Integer, tcuoi As Integer, loaips As Integer, Optional ps2 As Double) As Double
    Dim sql As String, i As Integer, sql2 As String
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        If i > 0 And i < 13 Then sql = sql + "+" + IIf(loaips < 0, "No_" + CStr(i), "Co_" + CStr(i))
        If i > 0 And i < 13 Then sql2 = sql2 + "+" + IIf(loaips < 0, "No_" + CStr(i) + "_NT", "Co_" + CStr(i) + "_NT")
    Next
    If Len(sql) > 0 Then SoPSTK = SelectSQL("SELECT Sum(" + sql + ") AS F1,Sum(" + sql2 + ") AS F2 FROM HethongTK WHERE TKCon=0 AND SoHieu LIKE '" + sh + "*'", ps2)
End Function

Public Function SoPSTKMaSo(ms As Long, tdau As Integer, tcuoi As Integer, loaips As Integer, Optional ps2 As Double) As Double
    Dim sql As String, i As Integer, sql2 As String
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        If i > 0 And i < 13 Then sql = sql + "+" + IIf(loaips < 0, "No_" + CStr(i), "Co_" + CStr(i))
        If i > 0 And i < 13 Then sql2 = sql2 + "+" + IIf(loaips < 0, "No_" + CStr(i) + "_NT", "Co_" + CStr(i) + "_NT")
    Next
    If Len(sql) > 0 Then SoPSTKMaSo = SelectSQL("SELECT Sum(" + sql + ") AS F1,Sum(" + sql2 + ") AS F2 FROM HethongTK WHERE MaSo=" + CStr(ms), ps2)
End Function

Public Function SoDuTKSH(sh As String, tcuoi As Integer, Optional loai As Integer = 0) As Double
    Dim st As String
    
    st = CStr(CThangDB(tcuoi))
    Select Case loai
        Case 0:
            SoDuTKSH = SelectSQL("SELECT Sum(IIF(Kieu<0,DuNo_" + st + "-DuCo_" + st + ",DuCo_" + st + "-DuNo_" + st + ")) AS F1 FROM HethongTK WHERE SoHieu='" + sh + "'")
        Case -1:
            SoDuTKSH = SelectSQL("SELECT Sum(DuNo_" + st + ") AS F1 FROM HethongTK WHERE SoHieu='" + sh + "'")
        Case 1:
            SoDuTKSH = SelectSQL("SELECT Sum(DuCo_" + st + ") AS F1 FROM HethongTK WHERE SoHieu='" + sh + "'")
        End Select
End Function

Public Sub XDSoHieuCap(cap As Integer)
    If cap < 1 Then Exit Sub
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu) SELECT First(MaSo) AS MS,SoHieu FROM HethongTK WHERE (Cap=" + CStr(cap) + ") OR (Cap<" + CStr(cap) + " AND TKCon=0) GROUP BY SoHieu"
End Sub

Public Function PSDu(shno As String, shco As String, tdau As Integer, tcuoi As Integer, Optional shct As String = "", Optional xl As Integer = 0) As Double
    PSDu = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
        & " WHERE HethongTK.SoHieu LIKE '" + shno + "*' AND TK.SoHieu LIKE '" + shco + "*' AND " + WThang("ThangCT", tdau, tcuoi) + IIf(Len(shct) > 0, " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "'", ""))
    If xl = 0 Then PSDu = PSDu + SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) _
        & " WHERE HethongTK.SoHieu LIKE '" + shno + "*' AND InStr(ChungTu.GhiChu,'" + shco + "')>0 AND " + WThang("ThangCT", tdau, tcuoi) + IIf(Len(shct) > 0, " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "'", ""))
End Function

Public Function BKChiTiet(xlapp As Excel.Application, xlsheet As Worksheet, TK As String, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, Optional dg As String = "") As Boolean
    Dim rs As Recordset, sql As String, i As Integer, ctk As New ClsTaikhoan
    Dim dn As Double, DC As Double, dnt As Double, j As Integer, X As String, k As Integer, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    
    sql = "SELECT DISTINCTROW LEFT(HethongTK.SoHieu,3) AS SHN,LEFT(TK.SoHieu,3) AS SHC FROM " + ChungTu2TKNC(0) _
        & " WHERE (HethongTK.SoHieu LIKE '" + TK + "*' OR TK.SoHieu LIKE '" + TK + "*') AND " + wsql + " AND SoPS<>0  GROUP BY LEFT(HethongTK.SoHieu,3),LEFT(TK.SoHieu,3)"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.RecordCount = 0 Then GoTo KgIn
    
    Recycle pCurDir + "CTPS.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\CTPS.XLS", pCurDir + "CTPS.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "CTPS.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    If Len(dg) > 0 Then
        xlsheet.Cells(3, 1) = dg + " " + TK
    Else
        xlsheet.Cells(3, 1) = xlsheet.Cells(3, 1) + " " + TK
    End If
    xlsheet.Cells(4, 1) = IIf(tdau > 0, ThoiGian(tdau, tcuoi), ThoiGianN(ndau, ncuoi))
    
    Do While Not rs.EOF
        X = IIf(rs!SHN <> Left(TK, 3), rs!SHN, rs!shc)
        For j = 1 To i
            If InStr(1, xlsheet.Cells(7, 4 + j), X) > 0 Then GoTo X1
        Next
        i = i + 1
        xlsheet.Cells(7, 4 + i) = IIf(rs!SHN <> Left(TK, 3), rs!SHN, rs!shc)
X1:
        rs.MoveNext
    Loop
    
    xlsheet.Range("E6", XLSCol(4 + i) + "6").MergeCells = True
    
    sql = "SELECT DISTINCTROW LEFT(HethongTK.SoHieu,3) AS SHN,LEFT(TK.SoHieu,3) AS SHC, ChungTu.SoHieu,DienGiai,SoPS FROM " + ChungTu2TKNC(0) _
        & " WHERE (HethongTK.SoHieu LIKE '" + TK + "*' OR TK.SoHieu LIKE '" + TK + "*') AND " + wsql + " AND SoPS<>0 ORDER BY NgayCT"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 0
    Do While Not rs.EOF
        j = j + 1
        xlsheet.Cells(7 + j, 1) = IIf(rs!SHN = Left(TK, 3), rs!SoHieu, "")
        xlsheet.Cells(7 + j, 2) = IIf(rs!shc = Left(TK, 3), rs!SoHieu, "")
        xlsheet.Cells(7 + j, 3) = rs!diengiai
        xlsheet.Cells(7 + j, 4) = rs!sops
        X = IIf(rs!SHN <> Left(TK, 3), rs!SHN, rs!shc)
        For k = 1 To i
            If InStr(1, xlsheet.Cells(7, 4 + k), X) > 0 Then
                xlsheet.Cells(7 + j, 4 + k) = rs!sops
                Exit For
            End If
        Next
        rs.MoveNext
    Loop
    
    For k = 0 To i
        xlsheet.Range(XLSCol(4 + k) + CStr(8 + j)).Formula = "=SUM(" + XLSCol(4 + k) + "8:" + XLSCol(4 + k) + CStr(7 + j) + ")"
    Next
    
    With xlsheet.Range("A6", XLSCol(4 + i) + CStr(8 + j))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 8
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlsheet.Cells(9 + j, 1) = "Sè d­ ®Çu kú"
    xlsheet.Cells(10 + j, 1) = "Sè d­ cuèi kú"
    ctk.InitTaikhoanSohieu TK
    ctk.SoDuNgay ndau - 1, dn, DC, dnt
    xlsheet.Cells(9 + j, 2) = IIf(ctk.kieu < 0, dn - DC, DC - dn)
    ctk.SoDuNgay ncuoi, dn, DC, dnt
    xlsheet.Cells(10 + j, 2) = IIf(ctk.kieu < 0, dn - DC, DC - dn)
    
    xlapp.Workbooks(1).Save
        
    xlapp.Workbooks.Close
    
    CallExcel "CTPS.XLS"
    GoTo KetThuc
KgIn:
    BKChiTiet = False
KetThuc:
    rs.Close
    Set rs = Nothing
    Set ctk = Nothing
End Function

Public Function PSTKCP(shno As String, tdau As Integer, tcuoi As Integer, Optional mtp As Long = 0) As Double
    PSTKCP = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) _
        & " WHERE HethongTK.SoHieu LIKE '" + shno + "*' AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mtp > 0, " AND MaTP=" + CStr(mtp), ""))
    PSTKCP = PSTKCP - SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
        & " WHERE MaLoai<>3 AND TK.SoHieu LIKE '" + shno + "*' AND LEFT(HethongTK.SoHieu,3)<>'154' AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mtp > 0, " AND MaTP=" + CStr(mtp), ""))
End Function

Public Function InSocaiTk2(taikhoan As ClsTaikhoan, tdau As Integer, tcuoi As Integer, thongbao As Boolean, doiung As String, VV As Long, Optional nn As Integer = 0) As Boolean
    Dim dkn As Double, dkc As Double, dknt As Double, ps As Boolean
    
With taikhoan
    ps = .TkCoPS(tdau, tcuoi)
    
    .SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
        
    ' In bao cao
    If dkn <> 0 Or dkc <> 0 Or ps Then
        
        If ps Then
            SetSQL "QSocai", "SELECT DISTINCTROW CTGhiSo.SoHieu, Last(ChungTu.NgayCT) AS NCT, Last(CTGhiSo.DienGiai) AS DG, Sum(IIF(HethongTK.SoHieu LIKE '" + taikhoan.SoHieu + "*', ChungTu.SoPS,0)) AS PSNo, Sum(IIF(HethongTK_1.SoHieu LIKE '" + taikhoan.SoHieu + "*', ChungTu.SoPS,0)) AS PSCo, IIF(HethongTK.SoHieu LIKE '" + taikhoan.SoHieu + "*',HeThongTK_1.SoHieu,HeThongTK.SoHieu) AS TKDu" _
                & " FROM (HeThongTK AS HeThongTK_3 RIGHT JOIN (HeThongTK AS HeThongTK_2 RIGHT JOIN (HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo) ON HeThongTK_2.MaSo = ChungTu.MaTKNo) ON HeThongTK_3.MaSo = ChungTu.MaTKCo) INNER JOIN CTGhiSo ON ChungTu.CTGS=CTGhiSo.MaSo" _
                & " Where SoPS<>0 AND ((HethongTK.SoHieu LIKE '" + taikhoan.SoHieu + "*'" + IIf(Len(doiung) = 0, "", " And HethongTK_3.SoHieu LIKE '" + doiung + "*'") _
                + ") Or (HethongTK_1.SoHieu LIKE '" + taikhoan.SoHieu + "*'" + IIf(Len(doiung) = 0, "", " And HethongTK_2.SoHieu LIKE '" + doiung + "*'") _
                + ")) And " + WThang("ThangCT", tdau, tcuoi) + IIf(VV > 1, " AND (ChungTu.MaDT = " + CStr(VV) + ")", "") + " AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) GROUP BY CTGhiSo.SoHieu, IIF(HethongTK.SoHieu LIKE '" + taikhoan.SoHieu + "*',HeThongTK_1.SoHieu,HeThongTK.SoHieu)"
        Else
            SetSQL "QSocai", "SELECT DISTINCTROW '' AS SoHieu,Null AS NCT, '' AS DienGiai, 0 AS PSNo, 0 AS PSCo, '' AS TKDu" _
            & " FROM ChungTu Where (chungtu.MaCT = 0)"
            frmMain.Rpt.Formulas(9) = "DuCk=" + DoiDau(IIf(.kieu < 0, dkn - dkc, dkc - dkn))
        End If
        
        frmMain.Rpt.Formulas(3) = "SoHieuTK='" + .SoHieu + "'"
        frmMain.Rpt.Formulas(4) = "NoDk=" + DoiDau(dkn - dkc)
        frmMain.Rpt.Formulas(5) = "Kieu=" + CStr(.kieu)
        frmMain.Rpt.Formulas(6) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
        frmMain.Rpt.Formulas(8) = "TenTk = IF PageNumber() = 1 THEN '" + .SoHieu + " - " + IIf(nn > 0, .TenE, .Ten) + IIf(VV > 1, " - VV: " + TenVV(VV), "") + "'"
        If VV > 1 Then
            taikhoan.SoDuTK tcuoi, dkn, dkc, dknt
            frmMain.Rpt.Formulas(9) = "DuCk = " + DoiDau(IIf(taikhoan.kieu < 0, dkn - dkc, dkc - dkn))
        End If
        frmMain.Rpt.Formulas(10) = "NoLK = " + DoiDau(SoPSTK(taikhoan.SoHieu, pThangDauKy, tcuoi, -1))
        frmMain.Rpt.Formulas(11) = "CoLK = " + DoiDau(SoPSTK(taikhoan.SoHieu, pThangDauKy, tcuoi, 1))
        InSocaiTk2 = True
        frmMain.Rpt.ReportFileName = "SOCAI4.RPT"
        RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    Else
        If thongbao Then MsgBox "Kh«ng cã sè ®Çu kú hoÆc ph¸t sinh !", vbInformation, App.ProductName
        InSocaiTk2 = False
    End If
End With
End Function
Public Function SoDuTKSHCP(sh As String, tcuoi As Integer) As Double
    SoDuTKSHCP = SelectSQL("SELECT DuNo_" + CStr(CThangDB(tcuoi)) + "-DuCo_" + CStr(CThangDB(tcuoi)) + " AS F1 FROM HethongTK WHERE SoHieu='" + sh + "'")
    SoDuTKSHCP = SoDuTKSHCP + SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(1) + " WHERE " + WThang("ThangCT", 0, tcuoi) + " AND MaLoai=3 AND HethongTK.SoHieu LIKE '" + sh + "*'")
End Function

Public Function InNhatKy(tdau As Integer, tcuoi As Integer, VV As Long, dktg As Integer, ndau As Date, ncuoi As Date, nn As Integer, Optional p As Integer = 0, Optional mdt1 As Long = 0, Optional mdt2 As Long = 0, Optional mdt3 As Long = 0) As Boolean
    Dim rs_nk As Recordset, sql As String, sqlw As String, sh As String

    sh = IIf(p > 0, "P", "")
    If dktg = 0 Then
        sqlw = WThang("ThangCT", tdau, tcuoi)
        RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    Else
        sqlw = WNgay("NgayGS", ndau, ncuoi)
        RptSetDate ncuoi, nn
    End If
    sql = "SELECT DISTINCTROW First(ChungTu" + sh + ".MaSo) AS MS,ChungTu" + sh + ".MaCT, ChungTu" + sh + ".SoHieu, ChungTu" + sh + ".NgayCT, ChungTu" + sh + ".NgayGS, ChungTu" + sh + ".DienGiai" + IIf(nn > 0, "E", "") + ", Sum(SoPS) AS SumOfSoPS, HeThongTK.SoHieu, HeThongTK.Ten" + IIf(nn > 0, "E", "") + ", -1 AS LoaiPS" _
        & " FROM " + ChungTu2TKNC(-1, p) _
        & " Where SoPS<>0 AND (MaTKTCNo) > 0 AND (((HeThongTK.loai) > 0) And ((chungtu" + sh + ".maloai <> 4) or (chungtu" + sh + ".maloai=4 and matkno<>matkco)) And " + sqlw + ")" + IIf(VV > 1, " AND (ChungTu" + sh + ".MaDT = " + CStr(VV) + ")", "") + IIf(mdt1 > 0, " AND (ChungTu" + sh + ".MaDT1 = " + CStr(mdt1) + ")", "") + IIf(mdt2 > 0, " AND (ChungTu" + sh + ".MaDT2 = " + CStr(mdt2) + ")", "") + IIf(mdt3 > 0, " AND (ChungTu" + sh + ".MaDT3 = " + CStr(mdt3) + ")", "") _
        & " GROUP BY ChungTu" + sh + ".MaCT, ChungTu" + sh + ".SoHieu, ChungTu" + sh + ".NgayCT, ChungTu" + sh + ".NgayGS, ChungTu" + sh + ".DienGiai" + IIf(nn > 0, "E", "") + ", HeThongTK.SoHieu, HeThongTK.Ten" + IIf(nn > 0, "E", "") + "" _
        & " UNION SELECT DISTINCTROW First(ChungTu" + sh + ".MaSo) AS MS,ChungTu" + sh + ".MaCT, ChungTu" + sh + ".SoHieu, ChungTu" + sh + ".NgayCT, ChungTu" + sh + ".NgayGS, ChungTu" + sh + ".DienGiai" + IIf(nn > 0, "E", "") + ", Sum(SoPS) AS SumOfSoPS, HeThongTK.SoHieu, HeThongTK.Ten" + IIf(nn > 0, "E", "") + ", 1 AS LoaiPS" _
        & " FROM " + ChungTu2TKNC(1, p) _
        & " Where SoPS<>0 AND (MaTKTCCo) > 0 AND (((HeThongTK.loai) > 0) And ((chungtu" + sh + ".maloai <> 4) or (chungtu" + sh + ".maloai=4 and matkno<>matkco)) And (" + sqlw + "))" + IIf(VV > 1, " AND (ChungTu" + sh + ".MaDT = " + CStr(VV) + ")", "") + IIf(mdt1 > 0, " AND (ChungTu" + sh + ".MaDT1 = " + CStr(mdt1) + ")", "") + IIf(mdt2 > 0, " AND (ChungTu" + sh + ".MaDT2 = " + CStr(mdt2) + ")", "") + IIf(mdt3 > 0, " AND (ChungTu" + sh + ".MaDT3 = " + CStr(mdt3) + ")", "") _
        & " GROUP BY ChungTu" + sh + ".MaCT, ChungTu" + sh + ".SoHieu, ChungTu" + sh + ".NgayCT, ChungTu" + sh + ".NgayGS, ChungTu" + sh + ".DienGiai" + IIf(nn > 0, "E", "") + ", HeThongTK.SoHieu, HeThongTK.Ten" + IIf(nn > 0, "E", "")
    SetSQL "MienTru", sql
    SetSQL "QNhatky", "SELECT * FROM MienTru ORDER BY NgayGS,MaCT"

    frmMain.Rpt.ReportFileName = "NHATKY.RPT"
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    If VV > 1 Then
        frmMain.Rpt.Formulas(5) = "VV = 'VV: " + TenVV(VV) + "'"
    End If
    frmMain.Rpt.Formulas(6) = "Rows = " + CStr(GetLastRow("QNhatKy", "MaCT"))
    InNhatKy = True
End Function
'====================================================================================================
' Thñ tôc x¸c ®Þnh c¸c ph¸t sinh do kÕt chuyÓn
'====================================================================================================
Public Sub XacdinhKC(tdau As Integer, tcuoi As Integer)
      Dim rs_taikhoan As Recordset, sql As String
      Dim rs_sub As Recordset, i As Integer
            
      ExecuteSQL5 "UPDATE HethongTK SET KC_N = 0, KC_C = 0, DuNo = DuNo_" + CStr(CThangDB(tcuoi)) + ", DuCo = DuCo_" + CStr(CThangDB(tcuoi))
        
      SetSQL "QCdt", "SELECT DISTINCTROW ChungTu.MaTkNo, SUM(SoPS) AS KC_N FROM ChungTu WHERE (ChungTu.MaLoai = 3) AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY ChungTu.MaTkNo"
            
      SetSQL "QChitiet", "SELECT DISTINCTROW ChungTu.MaTkCo, SUM(SoPS) AS KC_C FROM ChungTu WHERE (ChungTu.MaLoai = 3) AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY ChungTu.MaTkCo"
            
      Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT DISTINCTROW  QCdt.* FROM QCdt", dbOpenSnapshot, dbForwardOnly)
      
      Do While Not rs_taikhoan.EOF
            sql = "UPDATE HethongTK SET KC_N = " + DoiDau(rs_taikhoan!KC_N) + ", DuCo = DuCo + " + DoiDau(rs_taikhoan!KC_N) + " WHERE MaSo = " + CStr(rs_taikhoan!MaTkNo)
            ExecuteSQL5 sql
            rs_taikhoan.MoveNext
            i = 1
      Loop
            
      Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT DISTINCTROW  QChitiet.* FROM QChitiet", dbOpenSnapshot, dbForwardOnly)
      
      Do While Not rs_taikhoan.EOF
            sql = "UPDATE HethongTK SET KC_C = " + DoiDau(rs_taikhoan!KC_C) + ", DuNo = DuNo + " + DoiDau(rs_taikhoan!KC_C) + " WHERE MaSo = " + CStr(rs_taikhoan!MaTkCo)
            ExecuteSQL5 sql
            rs_taikhoan.MoveNext
            i = 1
      Loop
      
    ExecuteSQL5 "UPDATE HethongTK SET DuNo = 0, DuCo = 0 WHERE TkCon = 0 AND DuNo = DuCo"
    If i > 0 Then
        Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT HethongTK.MaSo FROM HethongTK" _
              & " WHERE Cap>0 AND TKCon>0 ORDER BY SoHieu DESC", dbOpenSnapshot, dbForwardOnly)
          Do While Not rs_taikhoan.EOF
              Set rs_sub = DBKetoan.OpenRecordset("SELECT DISTINCTROW Sum(HethongTK.KC_N) AS KCN, Sum(HethongTK.KC_C) AS KCC, Sum(DuNo) AS DN, Sum(DuCo) AS DC FROM HethongTK WHERE HethongTK.TkCha0 = " + CStr(rs_taikhoan!MaSo), dbOpenSnapshot)
              ExecuteSQL5 "UPDATE HethongTK SET KC_N = " + DoiDau(rs_sub!KCN) + ", KC_C = " + DoiDau(rs_sub!KCC) + ", DuNo = " + DoiDau(rs_sub!dn) + ", DuCo = " + DoiDau(rs_sub!DC) + " WHERE MaSo = " + CStr(rs_taikhoan!MaSo)
              rs_taikhoan.MoveNext
          Loop
            
        Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT HethongTK.MaSo FROM HethongTK WHERE Cap=0 AND TKCon>0", dbOpenSnapshot, dbForwardOnly)
        Do While Not rs_taikhoan.EOF
            Set rs_sub = DBKetoan.OpenRecordset("SELECT DISTINCTROW Sum(HethongTK.KC_N) AS KCN, Sum(HethongTK.KC_C) AS KCC, Sum(DuNo) AS DN, Sum(DuCo) AS DC FROM HethongTK WHERE HethongTK.TkCha0 = " + CStr(rs_taikhoan!MaSo), dbOpenSnapshot)
            
            ExecuteSQL5 "UPDATE HethongTK SET KC_N = " + DoiDau(rs_sub!KCN) + ", KC_C = " + DoiDau(rs_sub!KCC) _
                + ", DuNo = " + DoiDau(rs_sub!dn) + ", DuCo = " + DoiDau(rs_sub!DC) + " WHERE MaSo = " + CStr(rs_taikhoan!MaSo)
            rs_taikhoan.MoveNext
        Loop
        
        rs_sub.Close
        Set rs_sub = Nothing
    End If
    rs_taikhoan.Close
    Set rs_taikhoan = Nothing
End Sub

Public Function SoPSTK2(TK As ClsTaikhoan, tdau As Integer, tcuoi As Integer) As Double
    Dim sql As String, i As Integer
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        If i > 0 And i < 13 Then sql = sql + "+" + IIf(TK.kieu < 0, "No_" + CStr(i), "Co_" + CStr(i))
    Next
    If Len(sql) > 0 Then SoPSTK2 = SelectSQL("SELECT Sum(" + sql + ") AS F1 FROM HethongTK WHERE SoHieu='" + TK.SoHieu + "'")
    If TK.loai > 4 Then SoPSTK2 = SoPSTK2 - SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(IIf(TK.kieu < 0, 1, -1)) + " WHERE MaLoai<>3 AND HethongTK.SoHieu LIKE '" + TK.SoHieu + "*' AND " + WThang("ThangCT", tdau, tcuoi))
End Function

Public Function GetTK_ID(sh As String, mtk As Long, Optional kieu As Integer, Optional tkcon As Integer) As Long
    GetTK_ID = SelectSQL("SELECT TK_ID AS F1,Kieu AS F2,TKCon AS F3,MaSo AS F4 FROM HethongTK WHERE " + IIf(mtk > 0, "MaSo=" + CStr(mtk), "SoHieu='" + sh + "'"), kieu, tkcon, mtk)
End Function

Public Sub SoPSTKN(shtk As String, ndau As Date, ncuoi As Date, psn As Double, psc As Double)
Dim sql As String
    sql = "SELECT DISTINCTROW Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) _
        & "WHERE HethongTK.SoHieu LIKE '" + shtk + "*' AND " + WNgay("NgayGS", ndau, ncuoi)
    psn = SelectSQL(sql)
    sql = "SELECT DISTINCTROW Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(1) _
        & "WHERE HethongTK.SoHieu LIKE '" + shtk + "*' AND " + WNgay("NgayGS", ndau, ncuoi)
    psc = SelectSQL(sql)
End Sub

Public Function LaySHTKCapTren(shtk As String, cap As Integer) As String
    Dim sh As String, i As Integer, c As Integer
    
    sh = shtk
    If cap > 0 Then
        c = SelectSQL("SELECT Cap AS F1 FROM HethongTK WHERE SoHieu='" + sh + "'")
        For i = cap + 1 To c
            sh = SelectSQL("SELECT TK.SoHieu AS F1 FROM HethongTK INNER JOIN HethongTK AS TK ON HethongTK.TKCha0=TK.MaSo WHERE HethongTK.SoHieu='" + sh + "'")
        Next
    End If
    LaySHTKCapTren = sh
End Function
