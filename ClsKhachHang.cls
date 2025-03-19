VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsKhachHang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MaSo  As Long
Public MaPhanLoai As Long
Public SoHieu As String
Public Ten As String
Public TenE As String
Public DiaChi As String
Public mst As String
Public Tel As String
Public Fax As String
Public email As String
Public DaiDien As String
Public taikhoan  As String
Public DuMax As Double
Public GhiChu As String
Public MaNT As Long
'==========================================================================================================
' Thu tuc khoi tao object tu ma so
' mvt:  Ma so can tham chieu
'       =0 neu moi
'==========================================================================================================
Public Sub InitKhachHangMaSo(mkh As Long)
    Dim rs_vattu As Recordset
    
    If mkh > 0 Then
        Set rs_vattu = DBKetoan.OpenRecordset("SELECT * FROM KhachHang WHERE MaSo=" + CStr(mkh), dbOpenSnapshot)
        If rs_vattu.RecordCount = 0 Then GoTo init
        On Error Resume Next
        MaSo = rs_vattu!MaSo
        MaPhanLoai = rs_vattu!MaPhanLoai
        SoHieu = rs_vattu!SoHieu
        Ten = rs_vattu!Ten
        If pSongNgu Then TenE = rs_vattu!TenE
        DiaChi = rs_vattu!DiaChi
        mst = rs_vattu!mst
        Tel = rs_vattu!Tel
        Fax = rs_vattu!Fax
        email = rs_vattu!email
        DuMax = rs_vattu!DuMax
        DaiDien = rs_vattu!DaiDien
        taikhoan = rs_vattu!taikhoan
        GhiChu = rs_vattu!GhiChu
        MaNT = rs_vattu!MaNT
        On Error GoTo 0
        rs_vattu.Close
        Set rs_vattu = Nothing
    Else
init:
        MaSo = 0
        MaPhanLoai = 0
        SoHieu = "..."
        Ten = "..."
        TenE = "..."
        DiaChi = "..."
        mst = "..."
        Tel = "..."
        Fax = "..."
        email = "..."
        DuMax = 0
        DaiDien = "..."
        taikhoan = "..."
        GhiChu = "..."
        MaNT = 0
    End If
End Sub
'===================================================================================
' Thu tuc khoi tao object tu so hieu vat tu
' mvt:  So hieu vat tu can tham chieu
'===================================================================================
Public Function InitKhachHangSohieu(shkh As String) As String
    Dim sql As String
    
    sql = "SELECT DISTINCTROW MaSo AS F1 FROM KhachHang WHERE Sohieu='" + shkh + "'"
    InitKhachHangMaSo SelectSQL(sql)
    InitKhachHangSohieu = Ten
End Function
'===================================================================================
' Thu tuc dang ky kh¸ch hµng
'===================================================================================
Public Function GhiKhachHang() As Integer
    Dim sql As String
        
    sql = "INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu,Ten,DiaChi,MST,Tel,Fax,EMail,DaiDien,TaiKhoan,DuMax,GhiChu,MaNT) VALUES (" _
        + CStr(Lng_MaxValue("MaSo", "KhachHang") + 1) + "," + CStr(MaPhanLoai) + ",'" + SoHieu + "','" + Ten + "','" + DiaChi + "','" + mst + "','" + Tel + "','" + Fax + "','" + email + "','" + DaiDien + "','" _
        + taikhoan + "'," + DoiDau(DuMax) + ",'" + GhiChu + "'," + CStr(MaNT) + ")"
    If ExecuteSQL5(sql) <> 0 Then
        GhiKhachHang = -1
    Else
        GhiKhachHang = 0
        MaSo = Lng_MaxValue("MaSo", "KhachHang")
    End If
End Function
'===================================================================================
' Hµm tr¶ vÒ sè d­ theo ngµy
'===================================================================================
Public Sub SoDuKH(thang As Integer, duno As Double, duco As Double, dunt As Double, Optional mtk As Long = 0, Optional shtk As String = "")
    Dim rs As Recordset, st As String
    
    st = CStr(CThangDB(thang))
    If mtk > 0 Then
        Set rs = DBKetoan.OpenRecordset("SELECT Sum(DuNo_" + st + ") AS n,Sum(DuCo_" + st + ") AS c,Sum(DuNT_" + st + ") AS nt FROM SoDuKhachHang WHERE MaKhachHang=" + CStr(MaSo) + IIf(mtk > 0, " AND MaTaiKhoan=" + CStr(mtk), ""), dbOpenSnapshot)
    Else
        Set rs = DBKetoan.OpenRecordset("SELECT Sum(SoDuKhachHang.DuNo_" + st + ") AS n,Sum(SoDuKhachHang.DuCo_" + st + ") AS c,Sum(IIF(SoDuKhachHang.DuCo_" + st + ">0,SoDuKhachHang.DuNT_" + st + ",-SoDuKhachHang.DuNT_" + st + ")) AS nt FROM SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo WHERE MaKhachHang=" + CStr(MaSo) + " AND SoHieu LIKE '" + shtk + "*'", dbOpenSnapshot)
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

Public Sub SoPhatSinh(tdau As Integer, tcuoi As Integer, psn As Double, psc As Double, psn2 As Double, psc2 As Double, Optional mtk As Long)
    Dim i As Integer, th As String, sqln As String, sqlc As String, sqln2 As String, sqlc2 As String, sql As String, rs As Recordset
        
    psn = 0:    psc = 0
    For i = CStr(CThangDB(tdau)) To CStr(CThangDB(tcuoi))
        th = CStr(i)
        sqln = sqln + " + No_" + th
        sqlc = sqlc + " + Co_" + th
        sqln2 = sqln2 + " + No_" + th
        sqlc2 = sqlc2 + " + Co_" + th
    Next
    
    sql = "SELECT DISTINCTROW SUM(" + sqln + ") AS N,SUM(" + sqlc + ") AS C,SUM(" + sqln2 + ") AS N2,SUM(" + sqlc2 + ") AS C2 FROM SoDuKhachHang WHERE MaKhachHang=" + CStr(MaSo) + IIf(mtk > 0, " AND MaTaiKhoan=" + CStr(mtk), "")
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If Not IsNull(rs!n) Then
        psn = rs!n
        psc = rs!c
        psn2 = rs!n2
        psc2 = rs!c2
    Else
        psn = 0
        psc = 0
        psn2 = 0
        psc2 = 0
    End If
    rs.Close
    Set rs = Nothing
End Sub

Public Function SuaKH() As Long
    Dim sql As String
    
    sql = "UPDATE KhachHang SET MaPhanLoai=" + CStr(MaPhanLoai) + ",SoHieu='" + SoHieu + "',Ten='" + Ten + "',DiaChi='" _
        + DiaChi + "',MST = '" + mst + "',Tel='" + Tel + "',Fax='" + Fax + "',EMail='" + email + "',DaiDien='" + DaiDien + "',TaiKhoan='" _
        + taikhoan + "',DuMax=" + DoiDau(DuMax) + ",GhiChu='" + GhiChu + "',MaNT=" + CStr(MaNT) + " WHERE MaSo=" + CStr(MaSo)
    SuaKH = ExecuteSQL5(sql)
End Function

Public Function XoaKH() As Integer
    Dim X As Double
    XoaKH = -1
    If SelectSQL("SELECT Count(MaSo) AS F1 FROM ChungTu WHERE MaKH=" + CStr(MaSo) + " OR MaKHC=" + CStr(MaSo)) <> 0 Or SelectSQL("SELECT Count(MaSo) AS F1 FROM HoaDon WHERE MaKhachHang=" + CStr(MaSo)) <> 0 Or SelectSQL("SELECT Count(MaSo) AS F1 FROM ChungTuLQ WHERE MaKH=" + CStr(MaSo)) <> 0 Then Exit Function
    If SelectSQL("SELECT Sum(DuNo_0) AS F1,Sum(DuCo_0) AS F2 FROM SoDuKhachHang WHERE MaKhachHang=" + CStr(MaSo), X) <> 0 Then Exit Function
    If X <> 0 Then Exit Function
    If ExecuteSQL5("DELETE * FROM SoDuKhachHang WHERE MaKhachHang=" + CStr(MaSo)) <> 0 Then Exit Function
    XoaKH = ExecuteSQL5("DELETE FROM KhachHang WHERE MaSo=" + CStr(MaSo))
    If XoaKH = 0 Then SoDuTKCN
End Function

Public Sub GhiDauKy(mtk As Long, sodu1 As Double, sodu2 As Double, nt As Double)
    Dim sql As String, i As Integer, st As String, st3 As String, j As Integer
    
    sql = "UPDATE SoDuKhachHang SET DuNo_0 = " + DoiDau(Abs(sodu1)) + ",DuCo_0 = " + DoiDau(Abs(sodu2)) + ",DuNT_0=" + DoiDau(Abs(nt)) + " WHERE MaTaiKhoan= " + CStr(mtk) + " AND MaKhachHang=" + CStr(MaSo)
    
    ExecuteSQL5 sql
    If DBKetoan.RecordsAffected = 0 Then
        ExecuteSQL5 "INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(mtk) + "," + CStr(MaSo) + ")"
        ExecuteSQL5 sql
    End If
    
    sql = "UPDATE SoDuKhachHang SET MaTaiKhoan = MaTaiKhoan"
    For i = 1 To 12
          st = "DuNo_0 - DuCo_0"
          For j = 1 To i
                st3 = CStr(j)
                st = st + " + No_" + st3 + " - Co_" + st3
          Next
          st3 = CStr(i)
          sql = sql + ", DuNo_" + st3 + " = IIF(" + st + " > 0, " + st + ", 0), DuCo_" + st3 + " =  IIF(" + st + " < 0, -(" + st + "), 0)"
    Next
    ExecuteSQL5 sql + " WHERE MaTaiKhoan= " + CStr(mtk) + " AND MaKhachHang=" + CStr(MaSo)
    
    If MaNT > 0 Then
        sql = "UPDATE SoDuKhachHang  SET "
        For i = 1 To 12
              ExecuteSQL5 sql + "DuNT_" + CStr(i) + "=ABS(DuNT_" + CStr(i - 1) + "+IIF(DuNo_" + CStr(i - 1) + "-DuCo_" + CStr(i - 1) + ">=0,No_" + CStr(i) + "_NT-Co_" + CStr(i) + "_NT,Co_" + CStr(i) + "_NT-No_" + CStr(i) + "_NT)) WHERE MaTaiKhoan= " + CStr(mtk) + " AND MaKhachHang=" + CStr(MaSo)
        Next
    End If
    
    SoDuTKCN
End Sub

Public Sub GhiSoPSKH(TK As ClsTaikhoan, thang As Integer, ps As Double, nt As Double, loai As Integer)
    Dim sql As String, i As Integer, st As String, st3 As String, j As Integer
    
    st = CStr(CThangDB(thang))
    sql = "UPDATE SoDuKhachHang SET " + IIf(loai < 0, "No_", "Co_") + st + "=" + IIf(loai < 0, "No_", "Co_") + st + "+" + DoiDau(ps)
    If MaNT > 0 Then sql = sql + IIf(loai < 0, ",No_", ",Co_") + st + "_NT=" + IIf(loai < 0, "No_", "Co_") + st + "_NT+" + DoiDau(nt)
a:
    ExecuteSQL5 sql + " WHERE MaTaiKhoan= " + CStr(TK.MaSo) + " AND MaKhachHang=" + CStr(MaSo)
    If DBKetoan.RecordsAffected = 0 Then
        ExecuteSQL5 "INSERT INTO SoDuKhachHang (MaSo, MaTaiKhoan,MaKhachHang) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(TK.MaSo) + "," + CStr(MaSo) + ")"
        GoTo a
    End If
    
    sql = "UPDATE SoDuKhachHang SET MaTaiKhoan = MaTaiKhoan"
    For i = CThangDB(thang) To 12
          st3 = CStr(CThangDB(ThangTruoc(thang)))
          st = "DuNo_" + st3 + " - DuCo_" + st3
          For j = CThangDB(thang) To i
                st3 = CStr(j)
                st = st + " + No_" + st3 + " - Co_" + st3
          Next
          st3 = CStr(i)
          sql = sql + ", DuNo_" + st3 + " = IIF(" + st + " > 0, " + st + ", 0), DuCo_" + st3 + " =  IIF(" + st + " < 0, -(" + st + "), 0)"
    Next
    ExecuteSQL5 sql + " WHERE MaTaiKhoan= " + CStr(TK.MaSo) + " AND MaKhachHang=" + CStr(MaSo)
    
    If MaNT > 0 Then
        sql = "UPDATE SoDuKhachHang SET MaTaiKhoan = MaTaiKhoan"
        For i = CThangDB(thang) To 12
              st = "DuNT_" + CStr(CThangDB(ThangTruoc(thang)))
              For j = CThangDB(thang) To i
                    st3 = CStr(j)
                    If TK.kieu < 0 Then
                        st = st + " + No_" + st3 + "_NT - Co_" + st3 + "_NT"
                    Else
                        st = st + " - No_" + st3 + "_NT + Co_" + st3 + "_NT"
                    End If
              Next
              sql = sql + ", DuNT_" + CStr(i) + " = " + st
        Next
        ExecuteSQL5 sql + " WHERE MaTaiKhoan= " + CStr(TK.MaSo) + " AND MaKhachHang=" + CStr(MaSo)
    End If
End Sub

Public Sub GhiDauKy2(mtk As Long)
    Dim sql As String, i As Integer, st As String, st3 As String, j As Integer, sodu1 As Double, sodu2 As Double, nt As Double
        
    sodu1 = SelectSQL("SELECT Sum(DuNo_0) AS F1, Sum(DuCo_0) AS F2 FROM CNDauNam WHERE MaTaiKhoan=" + CStr(mtk) + " AND MaKhachHang=" + CStr(MaSo), sodu2)
    nt = SelectSQL("SELECT Sum(DuNT_0) AS F1 FROM CNDauNam WHERE MaTaiKhoan=" + CStr(mtk) + " AND MaKhachHang=" + CStr(MaSo))
    
    sql = "UPDATE SoDuKhachHang SET DuNo_0 = " + CStr(Abs(sodu1)) + ",DuCo_0 = " + CStr(Abs(sodu2)) + ",DuNT_0=" + DoiDau(Abs(nt)) + " WHERE MaTaiKhoan= " + CStr(mtk) + " AND MaKhachHang=" + CStr(MaSo)
    
    ExecuteSQL5 sql
    If DBKetoan.RecordsAffected = 0 Then
        ExecuteSQL5 "INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(mtk) + "," + CStr(MaSo) + ")"
        ExecuteSQL5 sql
    End If
    
    sql = "UPDATE SoDuKhachHang SET MaTaiKhoan = MaTaiKhoan"
    For i = 1 To 12
          st = "DuNo_0 - DuCo_0"
          For j = 1 To i
                st3 = CStr(j)
                st = st + " + No_" + st3 + " - Co_" + st3
          Next
          st3 = CStr(i)
          sql = sql + ", DuNo_" + st3 + " = IIF(" + st + " > 0, " + st + ", 0), DuCo_" + st3 + " =  IIF(" + st + " < 0, -(" + st + "), 0)"
    Next
    ExecuteSQL5 sql + " WHERE MaTaiKhoan= " + CStr(mtk) + " AND MaKhachHang=" + CStr(MaSo)
    
    If MaNT > 0 Then
        sql = "UPDATE SoDuKhachHang  SET "
        For i = 1 To 12
              ExecuteSQL5 sql + "DuNT_" + CStr(i) + "=ABS(DuNT_" + CStr(i - 1) + "+IIF(DuNo_" + CStr(i - 1) + "-DuCo_" + CStr(i - 1) + ">=0,No_" + CStr(i) + "_NT-Co_" + CStr(i) + "_NT,Co_" + CStr(i) + "_NT-No_" + CStr(i) + "_NT)) WHERE MaTaiKhoan= " + CStr(mtk) + " AND MaKhachHang=" + CStr(MaSo)
        Next
    End If
    
    SoDuTKCN
End Sub

Public Function HDChuaTT(thang As Integer, ngay As Date, shtk As String, mkh As Long) As String
    Dim sql As String, rs_phieu As Recordset
    
    If pCongNoHD = 0 Then Exit Function
    
    sql = "SELECT 'DK' AS SoHieu,#01/01/1980# AS NgayCT,'" + ABCtoVNI("D­ ®Çu kú") + "' AS DienGiai,CNDauNam.SoXuat,CNDauNam.DuNo_0 AS SoPS FROM CNDauNam INNER JOIN HethongTK ON CNDauNam.MaTaiKhoan=HethongTK.MaSo " _
        & " WHERE HethongTK.SoHieu LIKE '" + shtk + "*' AND (MaKhachHang = " + CStr(mkh) + ") AND (CNDauNam.DuNo_0>CNDauNam.SoXuat)" _
        & " UNION SELECT ChungTu.SoHieu, NgayCT, DienGiai, Sum(SoXuat), Sum(SoPS) FROM " + ChungTu2TKNC(-1) _
        & " WHERE HethongTK.SoHieu LIKE '" + shtk + "*' AND (ChungTu.MaKH = " + CStr(mkh) + ") AND " + IIf(thang > 0, WThang("ThangCT", 0, thang), WNgay("NgayGS", CVDate("01/01/80"), ngay)) _
        + " GROUP BY MaCT,ChungTu.SoHieu,NgayCT,DienGiai HAVING Sum(SoPS)>Sum(SoXuat) ORDER BY NgayCT DESC, SoHieu DESC"
    Set rs_phieu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    sql = ""
    Do While Not rs_phieu.EOF
        sql = sql + "CT: " + rs_phieu!SoHieu + ABCtoVNI(" - Ngµy ") + IIf(Year(rs_phieu!NgayCT) > 1980, Format(rs_phieu!NgayCT, Mask_D), "") + ABCtoVNI(" - Sè cßn nî: ") + Format(rs_phieu!sops - rs_phieu!SoXuat, Mask_0) + "; " '+ Chr(13)
        rs_phieu.MoveNext
    Loop
    rs_phieu.Close
    Set rs_phieu = Nothing
    
    HDChuaTT = sql
End Function

