VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsChungtu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MaSo As Long                  ' M· chi tiÕt chøng tõ
Public MaCT As Long                 ' M· chøng tõ 0: Tæng hîp, 1: nhËp, 2: xuÊt, 3: kÕt chuyÓn
Public maloai As Integer           ' Lo¹i chøng tõ
Public sohieu As String              ' Sè hiÖu chøng tõ
Public ThangCT As Integer        ' Th¸ng chøng tõ
Public NgayCT As Date              ' Ngµy chøng tõ
Public NgayGS As Date              ' Ngµy ghi sæ
Public MaNguon As Long          ' M· nguån nhËp/xuÊt
Public MaKho As Long               ' M· kho hµng nhËp xuÊt
Public diengiai As String          ' DiÔn gi¶i chøng tõ
Public DienGiaiE As String          ' DiÔn gi¶i chøng tõ
Public tkno As New ClsTaikhoan            ' M· tµi kho¶n ph¸t sinh nî
Public TkCo As New ClsTaikhoan            ' M· tµi kho¶n ph¸t sinh cã
Public sops As Double                ' Sè ph¸t sinh
Public SoPS2No As Double         ' Sè ph¸t sinh nguyªn tÖ/vËt t­ cña TK nî
Public SoPS2Co As Double         ' Sè ph¸t sinh nguyªn tÖ/vËt t­ cña TK cã
Public MaVattu As Long            ' M· vËt t­ nhËp xuÊt
Public GhiChu As String             ' Ghi chó
Public CT_ID As Long                   ' M· sè phiÕu nhËp khi xuÊt ®Ých danh
Public MaDT As Long
Public MaDT1 As Long
Public MaDT2 As Long
Public MaDT3 As Long
Public makh As Long
Public MaKHC As Long
Public CTGS As Long
Public MaTP As Long
Public dvt As Long
Public User_ID As Long
Public tygia As Double
Public MaNV As Long
Public HanTT As Integer
Public sh1 As String
Public T1 As Integer
Public TLCK As Double
Public CK As Double
Public PSUSD As Double
Public MauSo As String             ' Ghi chó
Public LoaiHD As String             ' Ghi chó
Public solo As String             ' Ghi chó
Public handung As String
Public phantramchietkhau As String ' ap dung cho phieu nhap (giam gia)
Public sotienchietkhau As String ' ap dung cho phieu nhap (giam gia)
'======================================================================================
' Thñ tôc khëi t¹o
'======================================================================================
Public Sub InitChungtu(ma As Long, loai As Integer, sh As String, thang As Integer, ngay1 As Date, ngay2 As Date, _
            nguon As Long, kho As Long, dg As String, MaTkNo As Long, MaTkCo As Long, ps As Double, ps2n As Double, _
            ps2c As Double, mvt As Long, gc As String, DTCT As Long, MauSo1 As String, LoaiHD1 As String, Solo1 As String, handung1 As String)
    Dim rs_ct As Recordset
    
    If ma > 0 Then
        Set rs_ct = DBKetoan.OpenRecordset("SELECT * FROM Chungtu WHERE MaSo=" + CStr(ma), dbOpenSnapshot)
        If rs_ct.RecordCount = 0 Then
            rs_ct.Close
            GoTo iC
        End If
        MaSo = ma
        MaCT = rs_ct!MaCT
        maloai = rs_ct!maloai
        sohieu = rs_ct!sohieu
        ThangCT = rs_ct!ThangCT
        NgayCT = rs_ct!NgayCT
        NgayGS = rs_ct!NgayGS
        MaNguon = rs_ct!MaNguon
        MaKho = rs_ct!MaKho
        diengiai = rs_ct!diengiai
        If pSongNgu Then DienGiaiE = rs_ct!DienGiaiE
        tkno.InitTaikhoanMaSo rs_ct!MaTkNo
        TkCo.InitTaikhoanMaSo rs_ct!MaTkCo
        sops = rs_ct!sops
        SoPS2No = rs_ct!SoPS2No
        SoPS2Co = rs_ct!SoPS2Co
        MaVattu = rs_ct!MaVattu
        GhiChu = rs_ct!GhiChu
        CT_ID = rs_ct!CT_ID
        MaDT = rs_ct!MaDT
        MaDT1 = rs_ct!MaDT1
        MaDT2 = rs_ct!MaDT2
        MaDT3 = rs_ct!MaDT3
        makh = rs_ct!makh
        MaKHC = rs_ct!MaKHC
        CTGS = rs_ct!CTGS
        MaTP = rs_ct!MaTP
        dvt = rs_ct!dvt
        User_ID = rs_ct!User_ID
        MaNV = rs_ct!MaNV
        If pTygia > 0 Then tygia = rs_ct!tygia
        HanTT = rs_ct!HanTT
        sh1 = rs_ct!sh1
        T1 = rs_ct!T1
        TLCK = rs_ct!TLCK
        CK = rs_ct!CK
        solo = IIf(IsNull(rs_ct!solo), "", rs_ct!solo)
        handung = IIf(IsNull(rs_ct!handung), "", rs_ct!handung)
        If pGiaUSD > 0 Then PSUSD = rs_ct!PSUSD
        
        If rs_ct!MauSoHD <> Null Then
        MauSo = rs_ct!MauSoHD
        Else
        MauSo = ""
        End If
        If rs_ct!LoaiHoaDon <> Null Then
        LoaiHD = rs_ct!LoaiHoaDon
        Else
        LoaiHD = ""
        End If
        
        
        rs_ct.Close
    Else
iC:
        MaSo = 0
        MaCT = 0
        maloai = loai
        sohieu = sh
        ThangCT = thang
        NgayCT = ngay1
        NgayGS = ngay2
        MaNguon = nguon
        MaKho = kho
        diengiai = dg
        tkno.InitTaikhoanMaSo MaTkNo
        TkCo.InitTaikhoanMaSo MaTkCo
        sops = ps
        SoPS2No = ps2n
        SoPS2Co = ps2c
        MaVattu = mvt
        GhiChu = IIf(Len(gc) > 0, gc, "...")
        CT_ID = 0
        MaDT = IIf(DTCT > 0, DTCT, 1)
        MaDT1 = 0
        MaDT2 = 0
        MaDT3 = 0
        makh = 0
        MaKHC = 0
        CTGS = 1
        MaTP = 0
        dvt = 0
        User_ID = 0
        MaNV = 0
        tygia = 0
        HanTT = 0
        sh1 = IIf(sh1 <> "", sh1, "...")
        T1 = 0
        TLCK = 0
        CK = 0
        PSUSD = 0
        MauSo = MauSo1
        LoaiHD = LoaiHD1
        solo = Solo1
        handung = handung1
    End If
    Set rs_ct = Nothing
End Sub

'======================================================================================
' Thñ tôc ghi chøng tõ (ch­a ghi sè ph¸t sinh vµo HethongTK)
'======================================================================================
Public Function GhiChungtu(Optional p As Integer = 0, Optional r As Integer = 0) As Integer
    Dim ckh As New ClsKhachHang, tp As Cls154, m As Long, st As String, dv As String
            
    If Len(DienGiaiE) = 0 Then DienGiaiE = "..."
    
    GhiChungtuTH IIf(maloai = 7 Or p <> 0, 1, 0)
    
    If maloai = 7 Or p <> 0 Or (pProcessMode = 1 And r = 0) Then GoTo KT
    MaSo = Lng_MaxValue("MaSo", "ChungTu" + IIf(p > 0 Or maloai = 7, "P", ""))
    
    WSpace.BeginTrans
    GhiChungtu = 0
    If maloai <> 9000 Then ' binh thuong thi <> 4 them chung tu ngan hang
        ' Ghi l¹i c¸c sè d­ hiÖn thêi cña c¸c tµi kho¶n cã ph¸t sinh
        If tkno.MaSo > 0 Then tkno.TinhSoDu ThangCT, sops, -1, IIf(tkno.MaNT > 0, SoPS2No, 0)
        If TkCo.MaSo > 0 Then TkCo.TinhSoDu ThangCT, sops, 1, IIf(TkCo.MaNT > 0, SoPS2Co, 0)
        
        ' TÝnh l¹i tån vËt t­ nÕu cã nhËp xuÊt
        If MaVattu > 0 Then
            If (maloai = 1 And tkno.tk_id = TKVT_ID) Or (maloai = 1 And TkCo.tk_id = TKVT_ID And sops = 0) Then TinhTonKho MaKho, IIf(tkno.MaSo > 0, tkno.MaSo, TkCo.MaSo), MaVattu, ThangCT, -1, IIf(tkno.MaSo > 0, SoPS2No, SoPS2Co), sops, PSUSD
            If ((maloai = 2 Or maloai = 9) And TkCo.tk_id = TKVT_ID) Or (maloai = 2 And tkno.tk_id = TKVT_ID And sops = 0) Then TinhTonKho MaKho, IIf(TkCo.MaSo > 0, TkCo.MaSo, tkno.MaSo), MaVattu, ThangCT, 1, IIf(TkCo.MaSo > 0, SoPS2Co, SoPS2No), sops, PSUSD
            If maloai = 2 And CT_ID < 0 And OutCost <> 0 Then
                If -CT_ID < 2000000000 Then
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat + " + DoiDau(IIf(SoPS2Co = 0, 1, SoPS2Co)) + " WHERE MaSo = " + CStr(-CT_ID)
                Else
                    ExecuteSQL5 "UPDATE VTDauNam SET SoXuat = SoXuat + " + DoiDau(IIf(SoPS2Co = 0, 1, SoPS2Co)) _
                        + " WHERE MaSoKho = " + CStr(MaKho) + " AND MaVattu = " + CStr(MaVattu) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
        End If
        
        If makh > 0 Then
            ckh.InitKhachHangMaSo makh
            If ckh.MaSo > 0 Then
                If tkno.tk_id = TKCNKH_ID Or tkno.tk_id = TKCNPT_ID Then ckh.GhiSoPSKH tkno, ThangCT, sops, SoPS2No, -1
            End If
        End If
        If MaKHC > 0 Then
            ckh.InitKhachHangMaSo MaKHC
            If ckh.MaSo > 0 Then
                If TkCo.tk_id = TKCNKH_ID Or TkCo.tk_id = TKCNPT_ID Then ckh.GhiSoPSKH TkCo, ThangCT, sops, SoPS2Co, 1
            End If
        End If
        
        If pCongNoHD > 0 Then
            If makh > 0 And CT_ID < 0 Then
                If Abs(CT_ID) < 2000000000 Then
                    m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNPT_ID) + " AND MaCT=" + CStr(Abs(CT_ID)))
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat + " + DoiDau(sops) + " WHERE MaSo = " + CStr(m)
                Else
                    ExecuteSQL5 "UPDATE CNDauNam SET SoXuat = SoXuat + " + DoiDau(sops) + " WHERE MaTaiKhoan = " + CStr(tkno.MaSo) + " AND MaKhachHang = " + CStr(makh) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
            If MaKHC > 0 And CT_ID < 0 Then
                If Abs(CT_ID) < 2000000000 Then
                    m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNKH_ID) + " AND MaCT=" + CStr(Abs(CT_ID)))
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat + " + DoiDau(sops) + " WHERE MaSo = " + CStr(m)
                Else
                    ExecuteSQL5 "UPDATE CNDauNam SET SoXuat = SoXuat + " + DoiDau(sops) + " WHERE MaTaiKhoan = " + CStr(TkCo.MaSo) + " AND MaKhachHang = " + CStr(MaKHC) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
        End If
    Else
        If tkno.MaSo <> TkCo.MaSo Then
            If tkno.MaSo > 0 And tkno.tk_id = TKVT_ID Then tkno.TinhSoDu ThangCT, sops, -1, IIf(tkno.MaNT > 0, SoPS2No, 0)
            If TkCo.MaSo > 0 And TkCo.tk_id = TKVT_ID Then TkCo.TinhSoDu ThangCT, sops, 1, IIf(TkCo.MaNT > 0, SoPS2Co, 0)
        End If
        If TkCo.tk_id = TKVT_ID Then TinhTonKho MaKho, TkCo.MaSo, MaVattu, ThangCT, 1, SoPS2Co, sops, PSUSD
        If tkno.tk_id = TKVT_ID Then TinhTonKho MaNguon, tkno.MaSo, MaVattu, ThangCT, -1, SoPS2Co, sops, PSUSD
        If CT_ID < 0 And OutCost <> 0 Then
            If -CT_ID < 2000000000 Then
                ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat + " + DoiDau(SoPS2Co) + " WHERE MaSo = " + CStr(-CT_ID)
            Else
                ExecuteSQL5 "UPDATE VTDauNam SET SoXuat = SoXuat + " + DoiDau(SoPS2Co) _
                    + " WHERE MaSoKho = " + CStr(MaKho) + " AND MaVattu = " + CStr(MaVattu) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
            End If
        End If
    End If
    WSpace.CommitTrans
    MaSo = Lng_MaxValue("MaSo", "ChungTu")
    If (maloai = 8 And TkCo.tk_id = TKDT_ID And Left(TkCo.sohieu, 4) <> "5113") And pGiaVon > 0 And MaVattu > 0 Then
        m = SelectSQL("SELECT MaTaiKhoan AS F1 FROM TonKho WHERE MaSoKho=" + CStr(MaKho) + " AND MaVattu=" + CStr(MaVattu) + " AND Luong_" + CStr(ThangCT) + ">0")
        If m = 0 Then m = SelectSQL("SELECT MaTK AS F1 FROM PhanLoaiVattu INNER JOIN Vattu ON PhanLoaiVattu.MaSo=Vattu.MaPhanLoai WHERE Vattu.MaSo=" + CStr(MaVattu))
        TkCo.InitTaikhoanMaSo m
        m = 0
        If sops = 0 Then
            m = SelectSQL("SELECT MaTK AS F1 FROM KhoHang WHERE MaSo=" + CStr(MaKho))
            If m > 0 Then tkno.InitTaikhoanMaSo m
        End If
        If m = 0 Then
            m = SelectSQL("SELECT MaTKGV AS F1 FROM KhoHang WHERE MaSo=" + CStr(MaKho))
            If m > 0 Then tkno.InitTaikhoanMaSo m
        End If
        If m = 0 Then tkno.InitTaikhoanMaSo SelectSQL("SELECT TOP 1 MaSo AS F1 FROM HethongTK WHERE SoHieu LIKE '632*' AND TKCon=0 ORDER BY SoHieu")
        If TkCo.MaSo = 0 Or tkno.MaSo = 0 Then Exit Function
        sohieu = sohieu + "GV"
        maloai = 2
        CT_ID = 500000000 + MaSo
        MaCT = SelectSQL("SELECT TOP 1 MaCT AS F1 FROM ChungTu WHERE ThangCT=" + CStr(ThangCT) + " AND MaLoai=2 AND SoHieu='" + sohieu + "'")
        If MaCT = 0 Then MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
        User_ID = 0
        If CTGS_GV > 0 Then CTGS = CTGS_GV
        If OutCost = 0 Then
            sops = GiaXuatKho(MaKho, TkCo.MaSo, MaVattu, NgayGS, SoPS2Co, PSUSD)
            
            GhiChungtu
        Else
            Dim i As Integer, luong() As Double, tien() As Double, id() As Long, c As Integer, tien2() As Double
            Select Case OutCost
                Case 1:
                    c = GiaXuatKhoDD(MaKho, TkCo.MaSo, MaVattu, SoPS2Co, luong, tien, id, tien2)
                Case 2:
                    c = GiaXuatKhoFIFO(MaKho, TkCo.MaSo, MaVattu, SoPS2Co, luong, tien, id, tien2)
                Case 3:
                    c = GiaXuatKhoLIFO(MaKho, TkCo.MaSo, MaVattu, SoPS2Co, luong, tien, id, tien2)
            End Select
            For i = 1 To c
                SoPS2Co = luong(i)
                sops = tien(i)
                If pGiaUSD > 0 Then PSUSD = tien2(i)
                CT_ID = id(i)
                GhiChungtu
            Next
            Erase luong
            Erase tien
            Erase tien2
            Erase id
        End If
    End If
    
If MaTP > 0 And (Left(tkno.sohieu, 3) = "154" Or Left(TkCo.sohieu, 3) = "154" Or Left(tkno.sohieu, 3) = "621" Or Left(TkCo.sohieu, 3) = "621") And (CT_ID < 610000000) Then 'CT_ID < 610000000 Or
        If DaTinhGiaThanh154(MaTP, ThangCT) Then
            Set tp = New Cls154
            tp.InitTPMaSo MaTP
            If MsgBox("TÝnh l¹i gi¸ thµnh cña ®èi t­îng " + tp.sohieu + " - " + tp.TenVattu + " ?", vbCritical + vbYesNo, App.ProductName) = vbYes Then
                DieuChinhGiaThanh tp, ThangCT
            End If
           
            Set tp = Nothing
        End If
      
    End If
KT:
    Set ckh = Nothing
End Function

Public Function XuLyCT() As Integer
    Dim ckh As New ClsKhachHang, tp As Cls154, m As Long, st As String, dv As String
                
    WSpace.BeginTrans
    XuLyCT = 0
    If maloai <> 4 Then
        ' Ghi l¹i c¸c sè d­ hiÖn thêi cña c¸c tµi kho¶n cã ph¸t sinh
        If tkno.MaSo > 0 Then tkno.TinhSoDu ThangCT, sops, -1, IIf(tkno.MaNT > 0, SoPS2No, 0)
        If TkCo.MaSo > 0 Then TkCo.TinhSoDu ThangCT, sops, 1, IIf(TkCo.MaNT > 0, SoPS2Co, 0)
        
        ' TÝnh l¹i tån vËt t­ nÕu cã nhËp xuÊt
        If MaVattu > 0 Then
            If (maloai = 1 And tkno.tk_id = TKVT_ID) Or (maloai = 1 And TkCo.tk_id = TKVT_ID And sops = 0) Then TinhTonKho MaKho, IIf(tkno.MaSo > 0, tkno.MaSo, TkCo.MaSo), MaVattu, ThangCT, -1, IIf(tkno.MaSo > 0, SoPS2No, SoPS2Co), sops, PSUSD
            If (maloai = 2 And TkCo.tk_id = TKVT_ID) Or (maloai = 2 And tkno.tk_id = TKVT_ID And sops = 0) Then TinhTonKho MaKho, IIf(TkCo.MaSo > 0, TkCo.MaSo, tkno.MaSo), MaVattu, ThangCT, 1, IIf(TkCo.MaSo > 0, SoPS2Co, SoPS2No), sops, PSUSD
            If maloai = 2 And CT_ID < 0 And OutCost <> 0 Then
                If -CT_ID < 2000000000 Then
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat + " + DoiDau(IIf(SoPS2Co = 0, 1, SoPS2Co)) + " WHERE MaSo = " + CStr(-CT_ID)
                Else
                    ExecuteSQL5 "UPDATE VTDauNam SET SoXuat = SoXuat + " + DoiDau(IIf(SoPS2Co = 0, 1, SoPS2Co)) _
                        + " WHERE MaSoKho = " + CStr(MaKho) + " AND MaVattu = " + CStr(MaVattu) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
        End If
        
        If makh > 0 Then
            ckh.InitKhachHangMaSo makh
            If ckh.MaSo > 0 Then
                If tkno.tk_id = TKCNKH_ID Or tkno.tk_id = TKCNPT_ID Then ckh.GhiSoPSKH tkno, ThangCT, sops, SoPS2No, -1
            End If
        End If
        If MaKHC > 0 Then
            ckh.InitKhachHangMaSo MaKHC
            If ckh.MaSo > 0 Then
                If TkCo.tk_id = TKCNKH_ID Or TkCo.tk_id = TKCNPT_ID Then ckh.GhiSoPSKH TkCo, ThangCT, sops, SoPS2Co, 1
            End If
        End If
        
        If pCongNoHD > 0 Then
            If makh > 0 And CT_ID < 0 Then
                If Abs(CT_ID) < 2000000000 Then
                    m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNPT_ID) + " AND MaCT=" + CStr(Abs(CT_ID)))
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat + " + DoiDau(sops) + " WHERE MaSo = " + CStr(m)
                Else
                    ExecuteSQL5 "UPDATE CNDauNam SET SoXuat = SoXuat + " + DoiDau(sops) + " WHERE MaTaiKhoan = " + CStr(tkno.MaSo) + " AND MaKhachHang = " + CStr(makh) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
            If MaKHC > 0 And CT_ID < 0 Then
                If Abs(CT_ID) < 2000000000 Then
                    m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNKH_ID) + " AND MaCT=" + CStr(Abs(CT_ID)))
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat + " + DoiDau(sops) + " WHERE MaSo = " + CStr(m)
                Else
                    ExecuteSQL5 "UPDATE CNDauNam SET SoXuat = SoXuat + " + DoiDau(sops) + " WHERE MaTaiKhoan = " + CStr(TkCo.MaSo) + " AND MaKhachHang = " + CStr(MaKHC) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
        End If
    Else
        If tkno.MaSo <> TkCo.MaSo Then
            If tkno.MaSo > 0 And tkno.tk_id = TKVT_ID Then tkno.TinhSoDu ThangCT, sops, -1, IIf(tkno.MaNT > 0, SoPS2No, 0)
            If TkCo.MaSo > 0 And TkCo.tk_id = TKVT_ID Then TkCo.TinhSoDu ThangCT, sops, 1, IIf(TkCo.MaNT > 0, SoPS2Co, 0)
        End If
        If TkCo.tk_id = TKVT_ID Then TinhTonKho MaKho, TkCo.MaSo, MaVattu, ThangCT, 1, SoPS2Co, sops, PSUSD
        If tkno.tk_id = TKVT_ID Then TinhTonKho MaNguon, tkno.MaSo, MaVattu, ThangCT, -1, SoPS2Co, sops, PSUSD
        If CT_ID < 0 And OutCost <> 0 Then
            If -CT_ID < 2000000000 Then
                ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat + " + DoiDau(SoPS2Co) + " WHERE MaSo = " + CStr(-CT_ID)
            Else
                ExecuteSQL5 "UPDATE VTDauNam SET SoXuat = SoXuat + " + DoiDau(SoPS2Co) _
                    + " WHERE MaSoKho = " + CStr(MaKho) + " AND MaVattu = " + CStr(MaVattu) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
            End If
        End If
    End If
    ExecuteSQL5 "UPDATE ChungTu SET XuLy=0 WHERE MaSo = " + CStr(MaSo)
    WSpace.CommitTrans
    MaSo = Lng_MaxValue("MaSo", "ChungTu")
    If (maloai = 8 And TkCo.tk_id = TKDT_ID And Left(TkCo.sohieu, 4) <> "5113") And pGiaVon > 0 And MaVattu > 0 Then
        m = SelectSQL("SELECT MaTaiKhoan AS F1 FROM TonKho WHERE MaSoKho=" + CStr(MaKho) + " AND MaVattu=" + CStr(MaVattu) + " AND Luong_" + CStr(ThangCT) + ">0")
        If m = 0 Then m = SelectSQL("SELECT MaTK AS F1 FROM PhanLoaiVattu INNER JOIN Vattu ON PhanLoaiVattu.MaSo=Vattu.MaPhanLoai WHERE Vattu.MaSo=" + CStr(MaVattu))
        TkCo.InitTaikhoanMaSo m
        m = 0
        If sops = 0 Then
            m = SelectSQL("SELECT MaTK AS F1 FROM KhoHang WHERE MaSo=" + CStr(MaKho))
            If m > 0 Then tkno.InitTaikhoanMaSo m
        End If
        If m = 0 Then
            m = SelectSQL("SELECT MaTKGV AS F1 FROM KhoHang WHERE MaSo=" + CStr(MaKho))
            If m > 0 Then tkno.InitTaikhoanMaSo m
        End If
        If m = 0 Then tkno.InitTaikhoanMaSo SelectSQL("SELECT TOP 1 MaSo AS F1 FROM HethongTK WHERE SoHieu LIKE '632*' AND TKCon=0 ORDER BY SoHieu")
        If TkCo.MaSo = 0 Or tkno.MaSo = 0 Then Exit Function
        sohieu = sohieu + "GV"
        maloai = 2
        CT_ID = 500000000 + MaSo
        MaCT = SelectSQL("SELECT TOP 1 MaCT AS F1 FROM ChungTu WHERE ThangCT=" + CStr(ThangCT) + " AND MaLoai=2 AND SoHieu='" + sohieu + "'")
        If MaCT = 0 Then MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
        User_ID = 0
        If CTGS_GV > 0 Then CTGS = CTGS_GV
        If OutCost = 0 Then
            sops = GiaXuatKho(MaKho, TkCo.MaSo, MaVattu, NgayGS, SoPS2Co, PSUSD)
            
            GhiChungtu
        Else
            Dim i As Integer, luong() As Double, tien() As Double, id() As Long, c As Integer, tien2() As Double
            Select Case OutCost
                Case 1:
                    c = GiaXuatKhoDD(MaKho, TkCo.MaSo, MaVattu, SoPS2Co, luong, tien, id, tien2)
                Case 2:
                    c = GiaXuatKhoFIFO(MaKho, TkCo.MaSo, MaVattu, SoPS2Co, luong, tien, id, tien2)
                Case 3:
                    c = GiaXuatKhoLIFO(MaKho, TkCo.MaSo, MaVattu, SoPS2Co, luong, tien, id, tien2)
            End Select
            For i = 1 To c
                SoPS2Co = luong(i)
                sops = tien(i)
                If pGiaUSD > 0 Then PSUSD = tien2(i)
                CT_ID = id(i)
                GhiChungtu
            Next
            Erase luong
            Erase tien
            Erase tien2
            Erase id
        End If
    End If
    If MaTP > 0 And (Left(tkno.sohieu, 3) = "621" Or Left(TkCo.sohieu, 3) = "621") And (CT_ID < 610000000) Then
        If DaTinhGiaThanh154(MaTP, ThangCT) Then
            Set tp = New Cls154
            tp.InitTPMaSo MaTP
            DieuChinhGiaThanh tp, ThangCT
            Set tp = Nothing
        End If
    End If
KT:
    Set ckh = Nothing
End Function

'======================================================================================
' Thñ tôc xãa chøng tõ
'======================================================================================
Public Sub XoaChungtu()
    Dim ckh As New ClsKhachHang, rs As Recordset, tp As Cls154, m As Long

    If maloai = 7 Then Exit Sub
    If pProcessMode = 1 Then
        ExecuteSQL5 "UPDATE ChungTu SET XuLy=2 WHERE MaSo = " + CStr(MaSo)
        Exit Sub
    End If
    WSpace.BeginTrans
    ' Ghi l¹i c¸c sè d­ hiÖn thêi cña c¸c tµi kho¶n cã ph¸t sinh
    If maloai <> 4 Then
        If tkno.MaSo > 0 Then tkno.TinhSoDu ThangCT, -sops, -1, -SoPS2No
        If TkCo.MaSo > 0 Then TkCo.TinhSoDu ThangCT, -sops, 1, -SoPS2Co
        ' TÝnh l¹i tån vËt t­ nÕu cã nhËp xuÊt
        If MaVattu > 0 Then
            If (maloai = 1 And tkno.tk_id = TKVT_ID) Or (maloai = 1 And TkCo.tk_id = TKVT_ID And sops = 0) Then TinhTonKho MaKho, IIf(tkno.MaSo > 0, tkno.MaSo, TkCo.MaSo), MaVattu, ThangCT, -1, IIf(tkno.MaSo > 0, -SoPS2No, -SoPS2Co), -sops, -PSUSD
            If (maloai = 2 And TkCo.tk_id = TKVT_ID) Or (maloai = 2 And tkno.tk_id = TKVT_ID And sops = 0) Then TinhTonKho MaKho, IIf(TkCo.MaSo > 0, TkCo.MaSo, tkno.MaSo), MaVattu, ThangCT, 1, IIf(TkCo.MaSo > 0, -SoPS2Co, -SoPS2No), -sops, -PSUSD

            If maloai = 2 And MaVattu > 0 And CT_ID < 0 And OutCost <> 0 Then
                If -CT_ID < 2000000000 Then
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat - " + DoiDau(SoPS2Co) _
                              + " WHERE MaSo = " + CStr(-CT_ID)
                Else
                    ExecuteSQL5 "UPDATE VTDauNam SET SoXuat = SoXuat - " + DoiDau(SoPS2Co) _
                              + " WHERE MaSoKho = " + CStr(MaKho) + " AND MaVattu = " + CStr(MaVattu) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
        End If
        If makh > 0 Then
            ckh.InitKhachHangMaSo makh
            If ckh.MaSo > 0 Then
                If tkno.tk_id = TKCNKH_ID Or tkno.tk_id = TKCNPT_ID Then ckh.GhiSoPSKH tkno, ThangCT, -sops, -SoPS2No, -1
            End If
        End If
        If MaKHC > 0 Then
            ckh.InitKhachHangMaSo MaKHC
            If ckh.MaSo > 0 Then
                If TkCo.tk_id = TKCNKH_ID Or TkCo.tk_id = TKCNPT_ID Then ckh.GhiSoPSKH TkCo, ThangCT, -sops, -SoPS2Co, 1
            End If
        End If
        If pCongNoHD > 0 Then
            If makh > 0 And CT_ID < 0 Then
                If -CT_ID < 2000000000 Then
                    m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNPT_ID) + " AND MaCT=" + CStr(Abs(CT_ID)))
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat - " + DoiDau(sops) + " WHERE MaSo = " + CStr(m)
                Else
                    ExecuteSQL5 "UPDATE CNDauNam SET SoXuat = SoXuat - " + DoiDau(sops) _
                              + " WHERE MaTaiKhoan = " + CStr(tkno.MaSo) + " AND MaKhachHang = " + CStr(makh) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
            If MaKHC > 0 And CT_ID < 0 Then
                If -CT_ID < 2000000000 Then
                    m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNKH_ID) + " AND MaCT=" + CStr(Abs(CT_ID)))
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat - " + DoiDau(sops) + " WHERE MaSo = " + CStr(m)
                Else
                    ExecuteSQL5 "UPDATE CNDauNam SET SoXuat = SoXuat - " + DoiDau(sops) _
                              + " WHERE MaTaiKhoan = " + CStr(TkCo.MaSo) + " AND MaKhachHang = " + CStr(MaKHC) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
        End If
        ExecuteSQL5 "DELETE FROM Chungtu WHERE MaSo=" + CStr(MaSo)

    Else
        If tkno.MaSo <> TkCo.MaSo Then
            If tkno.MaSo > 0 And tkno.tk_id = TKVT_ID Then tkno.TinhSoDu ThangCT, -sops, -1, -SoPS2No
            If TkCo.MaSo > 0 And TkCo.tk_id = TKVT_ID Then TkCo.TinhSoDu ThangCT, -sops, 1, -SoPS2Co
        End If
        If TkCo.tk_id = TKVT_ID Then TinhTonKho MaKho, TkCo.MaSo, MaVattu, ThangCT, 1, -SoPS2Co, -sops, -PSUSD
        If tkno.tk_id = TKVT_ID Then TinhTonKho MaNguon, tkno.MaSo, MaVattu, ThangCT, -1, -SoPS2Co, -sops, -PSUSD
        If MaVattu > 0 And CT_ID < 0 And OutCost <> 0 Then
            If -CT_ID < 2000000000 Then
                ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat - " + DoiDau(SoPS2Co) _
                          + " WHERE MaSo = " + CStr(-CT_ID)
            Else
                ExecuteSQL5 "UPDATE VTDauNam SET SoXuat = SoXuat - " + DoiDau(SoPS2Co) _
                          + " WHERE MaSoKho = " + CStr(MaKho) + " AND MaVattu = " + CStr(MaVattu) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
            End If
        End If
        ExecuteSQL5 "DELETE FROM Chungtu WHERE MaSo=" + CStr(MaSo)
    End If
    ExecuteSQL5 "DELETE FROM ChungtuLQ WHERE MaCT=" + CStr(MaCT)
    'xoa Import
    'Kiem tra xem den khi nao maso hop le
    Dim Query As String
    Query = "select * from HoaDon where MaSo=" + CStr(MaSo)
    Dim rs_ktra As Recordset
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? các b?n ghi
        Do While Not rs_ktra.EOF
            Query = "SELECT * FROM tbimport WHERE SHDon='" & Trim(CStr(rs_ktra!soHD)) & "' AND KHHDon='" & Trim(CStr(rs_ktra!KyHieu)) & "'"
            Dim rs_ktra2 As Recordset
            Set rs_ktra2 = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
            If Not rs_ktra2.EOF Then
                Do While Not rs_ktra2.EOF
                 ExecuteSQL5 "DELETE FROM tbimportdetail WHERE ParentId='" & Trim(CStr(rs_ktra2!id)) & "'"
                    rs_ktra2.MoveNext
                Loop
            End If

          ExecuteSQL5 "DELETE FROM tbimport WHERE SHDon='" & Trim(CStr(rs_ktra!soHD)) & "' AND KHHDon='" & Trim(CStr(rs_ktra!KyHieu)) & "'"
            rs_ktra.MoveNext
        Loop
    End If


    ExecuteSQL5 "DELETE FROM HoaDon WHERE MaSo=" + CStr(MaSo)
    WSpace.CommitTrans
    If (maloai = 8 And TkCo.tk_id = TKDT_ID) And pGiaVon > 0 And OutCost < 3 And MaVattu > 0 Then
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo FROM ChungTu WHERE SoHieu='" + sohieu + "GV' AND NgayCT=#" + Format(NgayCT, Mask_DB) + "# AND ThangCT=" + CStr(ThangCT), dbOpenSnapshot)
        Do While Not rs.EOF
            InitChungtu rs!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
            If MaSo > 0 Then XoaChungtu
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    If MaTP > 0 And (Left(tkno.sohieu, 3) = "621" Or Left(TkCo.sohieu, 3) = "621") And (CT_ID < 610000000) Then
        If DaTinhGiaThanh154(MaTP, ThangCT) Then
            Set tp = New Cls154
            tp.InitTPMaSo MaTP
            If MsgBox("TÝnh l¹i gi¸ thµnh cña ®èi t­îng " + tp.sohieu + " - " + tp.TenVattu + " ?", vbCritical + vbYesNo, App.ProductName) = vbYes Then
                DieuChinhGiaThanh tp, ThangCT
            End If
            Set tp = Nothing
        End If
    End If
    If maloai = 1 And tkno.tk_id = TKVT_ID Then
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo FROM ChungTu WHERE CT_ID=" + CStr(900000000 + MaSo), dbOpenSnapshot)
        Do While Not rs.EOF
            InitChungtu rs!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
            If MaSo > 0 Then XoaChungtu
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
KT:
    Set ckh = Nothing
End Sub
'======================================================================================
' Thñ tôc dÉn xuÊt
'======================================================================================
Public Sub Duplicate(ct As ClsChungtu)
    With ct
        .MaSo = MaSo
        .MaCT = MaCT
        .maloai = maloai
        .sohieu = sohieu
        .ThangCT = ThangCT
        .NgayCT = NgayCT
        .NgayGS = NgayGS
        .MaNguon = MaNguon
        .MaKho = MaKho
        .diengiai = diengiai
        .DienGiaiE = DienGiaiE
        ct.tkno.DanXuat .tkno
        ct.TkCo.DanXuat .TkCo
        .sops = sops
        .SoPS2No = SoPS2No
        .SoPS2Co = SoPS2Co
        .MaVattu = MaVattu
        .GhiChu = GhiChu
        .CT_ID = CT_ID
        .MaDT = MaDT
        .makh = makh
        .CTGS = CTGS
        .MaKHC = MaKHC
        .MaTP = MaTP
        .dvt = dvt
        .User_ID = User_ID
        .tygia = tygia
        .MaNV = MaNV
        .HanTT = HanTT
        .sh1 = sh1
        .T1 = T1
        .TLCK = TLCK
        .CK = CK
        .PSUSD = PSUSD
        .solo = solo
        .handung = handung
    End With
End Sub

Public Sub GhiThongtinCT(loai As Integer, Ten As String, DiaChi As String, ctgoc As String, makh As Long, Optional p As Integer = 0)
    Dim sh As String
    
    sh = IIf(p > 0 Or maloai = 7, "P", "")
    If SelectSQL("SELECT MaCT AS F1 FROM ChungtuLQ" + sh + " WHERE MaCT=" + CStr(MaCT) + " AND Loai=" + CStr(loai)) = 0 Then
        If Len(Ten) = 0 Then Ten = "..."
        If Len(DiaChi) = 0 Then DiaChi = "..."
        If Len(ctgoc) = 0 Then ctgoc = "..."
        ExecuteSQL5 "INSERT INTO ChungtuLQ" + sh + " (MaSo,MaCT,loai,HoTen,DiaChi,SoCTGoc,MaKH) VALUES (" + CStr(Lng_MaxValue("MaSo", "ChungTuLQ" + sh) + 1) + "," + CStr(MaCT) + "," + CStr(loai) + ",'" + Ten + "','" + DiaChi + "','" + ctgoc + "'," + CStr(makh) + ")"
    End If
End Sub

Public Function GhiChungtuTH(Optional p As Integer = 0, Optional r As Integer = 0) As Integer
    Dim sql As String, sh As String
    
    If User_ID = 0 Then User_ID = UserID
    
    If tkno.tk_id <> TKVT_ID And tkno.tk_id <> TKGT_ID And TkCo.tk_id <> TKVT_ID And TkCo.tk_id <> TKDT_ID And tkno.tk_id <> TKCNKH_ID And TkCo.tk_id <> TKCNKH_ID And tkno.tk_id <> TKCNPT_ID And TkCo.tk_id <> TKCNPT_ID Then MaVattu = 0
    If tkno.tk_id <> TKCNKH_ID And tkno.tk_id <> TKCNPT_ID Then makh = 0
    If TkCo.tk_id <> TKCNKH_ID And TkCo.tk_id <> TKCNPT_ID Then MaKHC = 0
    
    If MaVattu > 0 And sops = 0 Then
        If maloai = 1 And SoPS2No = 0 Then
            SoPS2No = SoPS2Co
            SoPS2Co = 0
        End If
        If maloai = 1 And tkno.MaSo = 0 Then
            TkCo.DanXuat tkno
            TkCo.InitTaikhoanMaSo 0
        End If
        If maloai = 2 And SoPS2Co = 0 Then
            SoPS2Co = SoPS2No
            SoPS2No = 0
        End If
        If maloai = 2 And TkCo.MaSo = 0 Then
            tkno.DanXuat TkCo
            tkno.InitTaikhoanMaSo 0
        End If
    End If
    If Not IsDate(handung) Then handung = Format(NgayCT, Mask_DB)
    sh = IIf(p > 0 Or maloai = 7, "P", "")
        'MAUSO,LOAIHD,MauSo + "," + LoaiHD + ","
'    sql = "INSERT INTO Chungtu" + sh + " (MaCT, MaLoai, SoHieu, ThangCT, NgayCT, NgayGS, MaNguon, MaKho, DienGiai, " + IIf(pSongNgu, "DienGiaiE,", "") _
'            & "MaTkNo, MaTkCo, SoPS, SoPS2No, SoPS2Co, MaTkTCNo, MaTkTCCo, MaVattu, GhiChu, CT_ID, MaDT, MaDT1, MaDT2, MaDT3,MaKH,CTGS,MaKHC,MaTP,DVT,User_ID,MaNV,HanTT,SH1,T1,TLCK,CK" + IIf(pTygia > 0, ",TyGia", "") + IIf(pGiaUSD > 0, ",PSUSD", "") + IIf(pProcessMode = 1 And r = 0, ",XuLy", "") + ",MAUSOHD,LOAIHoaDon,SoLo,HanDung) VALUES (" _
'            + CStr(MaCT) + "," + CStr(maloai) + ",'" + sohieu + "'," + CStr(ThangCT) + ",#" + Format(NgayCT, Mask_DB) _
'            + "#,#" + Format(NgayGS, Mask_DB) + "#," + CStr(MaNguon) + "," + CStr(MaKho) + ",'" + diengiai + "'," + IIf(pSongNgu, "'" + DienGiaiE + "',", "") _
'            + CStr(tkno.MaSo) + "," + CStr(TkCo.MaSo) + "," + DoiDau(sops) + "," + DoiDau(SoPS2No) + "," + DoiDau(SoPS2Co) + "," _
'            + CStr(tkno.MaTC) + "," + CStr(TkCo.MaTC) + "," + CStr(MaVattu) + ",'" + GhiChu + "'," + CStr(CT_ID) + "," + CStr(MaDT) + "," + CStr(MaDT1) + "," + CStr(MaDT2) + "," + CStr(MaDT3) + "," _
'            + CStr(makh) + "," + CStr(CTGS) + "," + CStr(MaKHC) + "," + CStr(MaTP) + "," + CStr(dvt) + "," + CStr(User_ID) + "," + CStr(MaNV) + "," + CStr(HanTT) + ",'" + sh1 + "'," + CStr(T1) + "," + DoiDau(TLCK) + "," + DoiDau(CK) + IIf(pTygia > 0, "," + DoiDau(tygia), "") + IIf(pGiaUSD > 0, "," + DoiDau(PSUSD), "") + IIf(pProcessMode = 1 And r = 0, ",1", "") + ",'" + MauSo + "','" + LoaiHD + "','" + solo + "','" + handung + "')"
            sql = "INSERT INTO Chungtu" + sh + " (MaCT, MaLoai, SoHieu, ThangCT, NgayCT, NgayGS, MaNguon, MaKho, DienGiai, " + IIf(pSongNgu, "DienGiaiE,", "") _
            & "MaTkNo, MaTkCo, SoPS, SoPS2No, SoPS2Co, MaTkTCNo, MaTkTCCo, MaVattu, GhiChu, CT_ID, MaDT, MaDT1, MaDT2, MaDT3,MaKH,CTGS,MaKHC,MaTP,DVT,User_ID,MaNV,HanTT,SH1,T1,TLCK,CK" + IIf(pTygia > 0, ",TyGia", "") + IIf(pGiaUSD > 0, ",PSUSD", "") + IIf(pProcessMode = 1 And r = 0, ",XuLy", "") + ",MAUSOHD,LOAIHoaDon,SoLo,HanDung,phantramchietkhau,sotienchietkhau) VALUES (" _
            + CStr(MaCT) + "," + CStr(maloai) + ",'" + sohieu + "'," + CStr(ThangCT) + ",#" + Format(NgayCT, Mask_DB) _
            + "#,#" + Format(NgayGS, Mask_DB) + "#," + CStr(MaNguon) + "," + CStr(MaKho) + ",'" + diengiai + "'," + IIf(pSongNgu, "'" + DienGiaiE + "',", "") _
            + CStr(tkno.MaSo) + "," + CStr(TkCo.MaSo) + "," + DoiDau(sops) + "," + DoiDau(SoPS2No) + "," + DoiDau(SoPS2Co) + "," _
            + CStr(tkno.MaTC) + "," + CStr(TkCo.MaTC) + "," + CStr(MaVattu) + ",'" + GhiChu + "'," + CStr(CT_ID) + "," + CStr(MaDT) + "," + CStr(MaDT1) + "," + CStr(MaDT2) + "," + CStr(MaDT3) + "," _
            + CStr(makh) + "," + CStr(CTGS) + "," + CStr(MaKHC) + "," + CStr(MaTP) + "," + CStr(dvt) + "," + CStr(User_ID) + "," + CStr(MaNV) + "," + CStr(HanTT) + ",'" + sh1 + "'," + CStr(T1) + "," + DoiDau(TLCK) + "," + DoiDau(CK) + IIf(pTygia > 0, "," + DoiDau(tygia), "") + IIf(pGiaUSD > 0, "," + DoiDau(PSUSD), "") + IIf(pProcessMode = 1 And r = 0, ",1", "") + ",'" + MauSo + "','" + LoaiHD + "','" + solo + "','" + handung + "','" + phantramchietkhau + "','" + sotienchietkhau + "')"

    ExecuteSQL5 sql
    
'    MaSo = Lng_MaxValue("MaSo", "ChungTu" + sh)
    
    If r > 0 Then
         If MaVattu > 0 Then
            If (maloai = 2 Or maloai = 4) And CT_ID < 0 And OutCost <> 0 Then
                If -CT_ID < 2000000000 Then
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat + " + DoiDau(IIf(SoPS2Co = 0, 1, SoPS2Co)) + " WHERE MaSo = " + CStr(-CT_ID)
                Else
                    ExecuteSQL5 "UPDATE VTDauNam SET SoXuat = SoXuat + " + DoiDau(IIf(SoPS2Co = 0, 1, SoPS2Co)) _
                        + " WHERE MaSoKho = " + CStr(MaKho) + " AND MaVattu = " + CStr(MaVattu) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
        End If
    End If
End Function

Private Sub Class_Initialize()
    User_ID = UserID
    DienGiaiE = "..."
    GhiChu = "..."
    sh1 = "..."
End Sub

Private Sub Class_Terminate()
    Set tkno = Nothing
    Set TkCo = Nothing
End Sub

Public Sub XuatThang(TKNo1 As Long, m As Long)
    If maloai <> 1 Or tkno.tk_id <> TKVT_ID Then Exit Sub
    tkno.DanXuat TkCo
    tkno.InitTaikhoanMaSo TKNo1
    SoPS2Co = SoPS2No
    SoPS2No = 0
    If tkno.tk_id = TKCNKH_ID Or tkno.tk_id = TKCNPT_ID Then makh = m
    If tkno.loai = 6 Or tkno.loai = 1 Then
        MaTP = m
    End If
    sohieu = sohieu + "XT"
    maloai = 2
    MaCT = SelectSQL("SELECT TOP 1 MaCT AS F1 FROM ChungTu WHERE ThangCT=" + CStr(ThangCT) + " AND MaLoai=2 AND CT_ID>900000000 AND SoHieu='" + sohieu + "'")
    If MaCT = 0 Then MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
    CT_ID = 900000000 + MaSo
    GhiChungtu
End Sub

Public Sub InitPhieu(ms As Long)
    InitChungtu ms, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
End Sub

Public Sub InitChungtuRS(rs_ct As Recordset)
        MaSo = rs_ct!MaSo
        MaCT = rs_ct!MaCT
        maloai = rs_ct!maloai
        sohieu = rs_ct!sohieu
        ThangCT = rs_ct!ThangCT
        NgayCT = rs_ct!NgayCT
        NgayGS = rs_ct!NgayGS
        MaNguon = rs_ct!MaNguon
        MaKho = rs_ct!MaKho
        diengiai = rs_ct!diengiai
        If pSongNgu Then DienGiaiE = rs_ct!DienGiaiE
        tkno.InitTaikhoanMaSo rs_ct!MaTkNo
        TkCo.InitTaikhoanMaSo rs_ct!MaTkCo
        sops = rs_ct!sops
        SoPS2No = rs_ct!SoPS2No
        SoPS2Co = rs_ct!SoPS2Co
        MaVattu = rs_ct!MaVattu
        GhiChu = rs_ct!GhiChu
        CT_ID = rs_ct!CT_ID
        MaDT = rs_ct!MaDT
        MaDT1 = rs_ct!MaDT1
        MaDT2 = rs_ct!MaDT2
        MaDT3 = rs_ct!MaDT3
        makh = rs_ct!makh
        MaKHC = rs_ct!MaKHC
        CTGS = rs_ct!CTGS
        MaTP = rs_ct!MaTP
        dvt = rs_ct!dvt
        User_ID = rs_ct!User_ID
        MaNV = rs_ct!MaNV
        If pTygia > 0 Then tygia = rs_ct!tygia
        HanTT = rs_ct!HanTT
        sh1 = rs_ct!sh1
        T1 = rs_ct!T1
        TLCK = rs_ct!TLCK
        CK = rs_ct!CK
        solo = IIf(IsNull(rs_ct!solo), "", rs_ct!solo)
        handung = IIf(IsNull(rs_ct!handung), "", rs_ct!handung)
        If pGiaUSD > 0 Then PSUSD = rs_ct!PSUSD
End Sub

Public Sub XuLyCTXoa()
    Dim ckh As New ClsKhachHang, rs As Recordset, tp As Cls154, m As Long
    
    If maloai = 7 Then Exit Sub
    WSpace.BeginTrans
    ExecuteSQL5 "DELETE FROM ChungtuLQ WHERE MaCT=" + CStr(MaCT)
    ExecuteSQL5 "DELETE FROM HoaDon WHERE MaSo=" + CStr(MaSo)
    ExecuteSQL5 "DELETE FROM Chungtu WHERE MaSo=" + CStr(MaSo)
    ' Ghi l¹i c¸c sè d­ hiÖn thêi cña c¸c tµi kho¶n cã ph¸t sinh
    If maloai <> 4 Then
        If tkno.MaSo > 0 Then tkno.TinhSoDu ThangCT, -sops, -1, -SoPS2No
        If TkCo.MaSo > 0 Then TkCo.TinhSoDu ThangCT, -sops, 1, -SoPS2Co
         ' TÝnh l¹i tån vËt t­ nÕu cã nhËp xuÊt
         If MaVattu > 0 Then
            If (maloai = 1 And tkno.tk_id = TKVT_ID) Or (maloai = 1 And TkCo.tk_id = TKVT_ID And sops = 0) Then TinhTonKho MaKho, IIf(tkno.MaSo > 0, tkno.MaSo, TkCo.MaSo), MaVattu, ThangCT, -1, IIf(tkno.MaSo > 0, -SoPS2No, -SoPS2Co), -sops, -PSUSD
            If (maloai = 2 And TkCo.tk_id = TKVT_ID) Or (maloai = 2 And tkno.tk_id = TKVT_ID And sops = 0) Then TinhTonKho MaKho, IIf(TkCo.MaSo > 0, TkCo.MaSo, tkno.MaSo), MaVattu, ThangCT, 1, IIf(TkCo.MaSo > 0, -SoPS2Co, -SoPS2No), -sops, -PSUSD
        
            If maloai = 2 And MaVattu > 0 And CT_ID < 0 And OutCost <> 0 Then
                If -CT_ID < 2000000000 Then
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat - " + DoiDau(SoPS2Co) _
                        + " WHERE MaSo = " + CStr(-CT_ID)
                Else
                    ExecuteSQL5 "UPDATE VTDauNam SET SoXuat = SoXuat - " + DoiDau(SoPS2Co) _
                        + " WHERE MaSoKho = " + CStr(MaKho) + " AND MaVattu = " + CStr(MaVattu) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
        End If
        If makh > 0 Then
            ckh.InitKhachHangMaSo makh
            If ckh.MaSo > 0 Then
                If tkno.tk_id = TKCNKH_ID Or tkno.tk_id = TKCNPT_ID Then ckh.GhiSoPSKH tkno, ThangCT, -sops, -SoPS2No, -1
            End If
        End If
        If MaKHC > 0 Then
            ckh.InitKhachHangMaSo MaKHC
            If ckh.MaSo > 0 Then
                If TkCo.tk_id = TKCNKH_ID Or TkCo.tk_id = TKCNPT_ID Then ckh.GhiSoPSKH TkCo, ThangCT, -sops, -SoPS2Co, 1
            End If
        End If
        If pCongNoHD > 0 Then
            If makh > 0 And CT_ID < 0 Then
                If -CT_ID < 2000000000 Then
                    m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNPT_ID) + " AND MaCT=" + CStr(Abs(CT_ID)))
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat - " + DoiDau(sops) + " WHERE MaSo = " + CStr(m)
                Else
                    ExecuteSQL5 "UPDATE CNDauNam SET SoXuat = SoXuat - " + DoiDau(sops) _
                        + " WHERE MaTaiKhoan = " + CStr(tkno.MaSo) + " AND MaKhachHang = " + CStr(makh) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
            If MaKHC > 0 And CT_ID < 0 Then
                If -CT_ID < 2000000000 Then
                    m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNKH_ID) + " AND MaCT=" + CStr(Abs(CT_ID)))
                    ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat - " + DoiDau(sops) + " WHERE MaSo = " + CStr(m)
                Else
                    ExecuteSQL5 "UPDATE CNDauNam SET SoXuat = SoXuat - " + DoiDau(sops) _
                        + " WHERE MaTaiKhoan = " + CStr(TkCo.MaSo) + " AND MaKhachHang = " + CStr(MaKHC) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
                End If
            End If
        End If
    Else
        If tkno.MaSo <> TkCo.MaSo Then
            If tkno.MaSo > 0 And tkno.tk_id = TKVT_ID Then tkno.TinhSoDu ThangCT, -sops, -1, -SoPS2No
            If TkCo.MaSo > 0 And TkCo.tk_id = TKVT_ID Then TkCo.TinhSoDu ThangCT, -sops, 1, -SoPS2Co
        End If
        If TkCo.tk_id = TKVT_ID Then TinhTonKho MaKho, TkCo.MaSo, MaVattu, ThangCT, 1, -SoPS2Co, -sops, -PSUSD
        If tkno.tk_id = TKVT_ID Then TinhTonKho MaNguon, tkno.MaSo, MaVattu, ThangCT, -1, -SoPS2Co, -sops, -PSUSD
        If MaVattu > 0 And CT_ID < 0 And OutCost <> 0 Then
            If -CT_ID < 2000000000 Then
                ExecuteSQL5 "UPDATE ChungTu SET SoXuat = SoXuat - " + DoiDau(SoPS2Co) _
                    + " WHERE MaSo = " + CStr(-CT_ID)
            Else
                ExecuteSQL5 "UPDATE VTDauNam SET SoXuat = SoXuat - " + DoiDau(SoPS2Co) _
                    + " WHERE MaSoKho = " + CStr(MaKho) + " AND MaVattu = " + CStr(MaVattu) + " AND MaSo=" + CStr(-CT_ID - 2000000000)
            End If
        End If
    End If
        
    WSpace.CommitTrans
    If (maloai = 8 And TkCo.tk_id = TKDT_ID) And pGiaVon > 0 And OutCost < 3 And MaVattu > 0 Then
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo FROM ChungTu WHERE SoHieu='" + sohieu + "GV' AND NgayCT=#" + Format(NgayCT, Mask_DB) + "# AND ThangCT=" + CStr(ThangCT), dbOpenSnapshot)
        Do While Not rs.EOF
            InitChungtu rs!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
            If MaSo > 0 Then XoaChungtu
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    If MaTP > 0 And (Left(tkno.sohieu, 3) = "621" Or Left(TkCo.sohieu, 3) = "621") And (CT_ID < 610000000) Then
        If DaTinhGiaThanh154(MaTP, ThangCT) Then
            Set tp = New Cls154
            tp.InitTPMaSo MaTP
            If MsgBox("TÝnh l¹i gi¸ thµnh cña ®èi t­îng " + tp.sohieu + " - " + tp.TenVattu + " ?", vbCritical + vbYesNo, App.ProductName) = vbYes Then
                DieuChinhGiaThanh tp, ThangCT
            End If
            Set tp = Nothing
        End If
    End If
    If maloai = 1 And tkno.tk_id = TKVT_ID Then
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo FROM ChungTu WHERE CT_ID=" + CStr(900000000 + MaSo), dbOpenSnapshot)
        Do While Not rs.EOF
            InitChungtu rs!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
            If MaSo > 0 Then XoaChungtu
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
KT:
    Set ckh = Nothing
End Sub




