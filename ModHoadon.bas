Attribute VB_Name = "ModHoadon"
Option Explicit

Public Type tpHoaDon
      MaSo As Long
      loai As Integer
      MaKhachHang As Long
      KyHieu As String
      sohd As String
      NgayPH    As Date
      MatHang As String
      SoLuong As Double
      ThanhTien As Double
      TyLe As Integer
      HD As Integer
      KCT As Integer
      HDBL As Integer
      NK As Integer
      DC As Integer
      ts As Integer
      TenKH As String
      DiaChiKH As String
      MSTKH As String
      HTTT As String
      MauSo As String
      tygia As Double
End Type

Public h As tpHoaDon

Public Sub GhiHoaDon(p As Integer)
    Dim sql As String
    
    With h
        If Not KHDetail Then
            sql = "INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu,Ten,DiaChi,MST) VALUES (" + CStr(.MaSo) + "," + CStr(Lng_MaxValue("MaSo", "PhanLoaiKhachHang")) + ",'" + CStr(.MaSo) + "','" + .TenKH + "','" + .DiaChiKH + "','" + .MSTKH + "')"
            ExecuteSQL5 sql
        End If
        sql = "INSERT INTO HoaDon" + IIf(p > 0, "P", "") + " (MaSo,Loai,MaKhachHang,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,TyLe,HD,KCT,HDBL,NK,TS, DC,HTTT,MauSo, TyGia) VALUES (" + CStr(.MaSo) + "," + CStr(.loai) + "," + CStr(IIf(KHDetail, .MaKhachHang, .MaSo)) _
            + ",'" + .KyHieu + "','" + .sohd + "',#" + Format(.NgayPH, Mask_DB) + "#,'" + .MatHang + "'," + DoiDau(.SoLuong) + "," + DoiDau(.ThanhTien) + "," + CStr(.TyLe) + "," + CStr(.HD) + "," + CStr(.KCT) + "," + CStr(.HDBL) + "," + CStr(.NK) + "," + CStr(.ts) + "," + CStr(.DC) + ",'" + .HTTT + "','" + .MauSo + "'," + DoiDau(.tygia) + ")"
        ExecuteSQL5 sql
    End With
End Sub

Public Sub CopyHD(h1 As tpHoaDon, h2 As tpHoaDon)
    With h2
        .MaSo = h1.MaSo
        .loai = h1.loai
        .NgayPH = h1.NgayPH
        .MaKhachHang = h1.MaKhachHang
        .KyHieu = h1.KyHieu
        .sohd = h1.sohd
        .MatHang = h1.MatHang
        .SoLuong = h1.SoLuong
        .ThanhTien = h1.ThanhTien
        .TyLe = h1.TyLe
        .HD = h1.HD
        .KCT = h1.KCT
        .HDBL = h1.HDBL
        .NK = h1.NK
        .ts = h1.ts
        .DC = h1.DC
        .TenKH = h1.TenKH
        .DiaChiKH = h1.DiaChiKH
        .MSTKH = h1.MSTKH
        .HTTT = h1.HTTT
        .MauSo = h1.MauSo
        .tygia = h1.tygia
    End With
End Sub

Public Function LayHoaDon(HD() As tpHoaDon, mc As Integer)
    If mc < 0 Then Exit Function
    CopyHD HD(mc), h
End Function

Public Function BotHoaDon(HD() As tpHoaDon, mc As Integer, sohd As Integer)
    Dim i As Integer
    
    If mc < 0 Then Exit Function
    
    For i = mc To sohd - 1
        CopyHD HD(i), HD(i + 1)
    Next
    sohd = sohd - 1
    
    If sohd >= 0 Then
        ReDim Preserve HD(0 To sohd) As tpHoaDon
    Else
        Erase HD
    End If
End Function

Public Sub XoaHD()
    With h
        .MaSo = 0
        .loai = 0
        .MaKhachHang = 0
        .KyHieu = "..."
        .sohd = "..."
        .MatHang = "..."
        .SoLuong = 0
        .ThanhTien = 0
        .TyLe = 0
        .HD = 1
        .KCT = 0
        .HDBL = 0
        .NK = 0
        .ts = 0
        .DC = 0
        .TenKH = "..."
        .DiaChiKH = "..."
        .MSTKH = "..."
        .HTTT = "..."
        .MauSo = "..."
        .tygia = 0
    End With
End Sub

Public Function PhaiNopVAT(tdau As Integer, tcuoi As Integer, shtk As String, dn As Double) As Double
    Dim kq As Double, pn As Double, sh1 As String
    
    If Len(shtk) > 6 Then
        sh1 = pVATV + Right(shtk, Len(shtk) - 3)
    Else
        sh1 = pVATV
    End If
    
    pn = SoPSTK(shtk, tdau, tcuoi, 1)
    pn = pn - (SoPSTK(sh1, tdau, tcuoi, -1) - PSDu("6", sh1, tdau, tcuoi))
    PhaiNopVAT = pn
    dn = PSDu(shtk, "111", tdau, tcuoi) + PSDu(shtk, "112", tdau, tcuoi)
End Function

