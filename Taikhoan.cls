VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTaikhoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -------- Thong tin ---------------------------------------------------------------------------------------------------------
Public MaSo As Long                              ' M· sè
Public sohieu As String                           ' Sè hiÖu tµi kho¶n
Public Ten As String                                ' Tªn tµi kho¶n
Public TenE As String                                ' Tªn tµi kho¶n
Public NoDauKy As Double                   ' Sè d­ nî ®Çu kú
Public CoDauKy As Double                   ' Sè d­ cã ®Çu kú
Public NTDauKy As Double                  ' Sè d­ nguyªn tÖ ®Çu kú
' -------- Phan loai ---------------------------------------------------------------------------------------------------------
Public cap As Integer                              ' CÊp tµi kho¶n
Public kieu As Integer                            ' KiÓu tµi kho¶n (1: cã, -1: nî).
Public loai As Integer                             ' Lo¹i tµi kho¶n (0..9).
Public MaTC As Long           ' M· tµi kho¶n tµi chÝnh
' -------- Cac doi tuong quan he --------------------------------------------------------------------------------------
Public tkcon As Integer                        ' 1 : Cã tµi kho¶n con cÊp d­íi, 0 : lµ tµi kho¶n chi tiÕt
Public TkCha0 As Long                     ' M· tµi kho¶n cÊp trªn
Public TkCha1 As Long                     ' M· tµi kho¶n cÊp trªn
Public TkCha2 As Long                     ' M· tµi kho¶n cÊp trªn
Public TkCha3 As Long                     ' M· tµi kho¶n cÊp trªn
Public TkCha4 As Long                     ' M· tµi kho¶n cÊp trªn
Public TkCha5 As Long                     ' M· tµi kho¶n cÊp trªn
Public MaNT As Long                             ' M· nguyªn tÖ
Public SoDuMax As Double                  ' Sè d­ tèi ®a cho phÐp ( =0 nÕu bá qua )
Public SoDuMin As Double                  ' Sè d­ tèi thiÓu cho phÐp ( =0 nÕu bá qua )
Public tk_id As Integer                       ' M· nhËn d¹ng tµi kho¶n
Public tk_id2 As Integer                       ' M· nhËn d¹ng tµi kho¶n
Public GhiChu As String

Public TenDA As String
Public NhomDA As String
Public DiaDiem As String
Public DuToan As Double

Public Von1 As Double
Public Von2 As Double
Public Von3 As Double

Public NgayKC As Date
Public NgayHT As Date

Public PSNLK As Double
Public PSCLK As Double

' Thu tuc dang ky tai khoan khong la tai khoan tai chinh
' Tra ve ma so tai khoan duoc dang ky
Public Function ThemTk() As Integer
    Dim i As Integer, sql As String
    
    If MaNT = 0 Then
        Dim rs_taikhoan As Recordset
        sql = "SELECT DISTINCTROW Count(MaSo) As F1 FROM HethongTK WHERE SoHieu='" + sohieu + "' AND MaNT<=0"
        If SelectSQL(sql) > 0 Then
            ThemTk = -1
            Exit Function
        End If
    End If
    If Len(TenE) = 0 Then TenE = "..."
    WSpace.BeginTrans
    If ExecuteSQL5("INSERT INTO HeThongTK (MaSo, SoHieu,Ten" + IIf(pSongNgu, ",TenE", "") + ",Cap,Kieu ,Loai,TKCon,MaNT, TK_ID, TK_ID2," _
        & " SoDuMax,SoDuMin, TkCha0, TkCha1, TkCha2, TkCha3, TkCha4, TkCha5, MaTC, GhiChu, TenDA, NhomDA, DiaDiem, DuToan, Von1, Von2, Von3, NgayKC, NgayHT, PSNLK, PSCLK,THEMMOI)" _
        & " VALUES (" + CStr(Lng_MaxValue("MaSo", "HethongTK") + 1) + ",'" + sohieu + "','" + Ten + "'" + IIf(pSongNgu, ",'" + TenE + "'", "") + "," + CStr(cap) + "," + CStr(kieu) + "," + CStr(loai) + "," + CStr(tkcon) + "," + CStr(MaNT) _
        + "," + CStr(tk_id) + "," + CStr(tk_id2) + "," + DoiDau(SoDuMax) + "," + DoiDau(SoDuMin) + "," + CStr(TkCha0) + "," + CStr(TkCha1) _
        + "," + CStr(TkCha2) + "," + CStr(TkCha3) + "," + CStr(TkCha4) + "," + CStr(TkCha5) + "," + CStr(MaTC) + ",'" + GhiChu + "','" + TenDA + "','" + NhomDA + "','" _
        + DiaDiem + "'," + CStr(DuToan) + "," + CStr(Von1) + "," + CStr(Von2) + "," + CStr(Von3) + ",#" + Format(NgayKC, Mask_DB) + "#,#" + Format(NgayHT, Mask_DB) + "#," + CStr(PSNLK) + "," + CStr(PSCLK) + ",'1')") <> 0 Then
        ThemTk = -1
        WSpace.Rollback
        Exit Function
    End If
    ' Lay ma so tai khoan moi mo
    MaSo = Lng_MaxValue("MaSo", "HethongTK")
    
    If MaNT > 0 Then
        ExecuteSQL5 "Update HethongTK Set MaNT=-1 Where SoHieu='" + sohieu + "' AND MaNT=0"
    End If
    
    ' Neu tai khoan cha chua co tai khoan con thi cap nhat truong TkCon va chuyen so du
    ChuyenSoDu
    ThemTk = 0
       
    ' Nhap so dau ky
    If NoDauKy <> 0 Or CoDauKy <> 0 Or NTDauKy <> 0 Then
        TinhSoDu 0, 0, -1, 0
    End If
    
     ' Sua lai matc khong tai khoan cap 5 ko in ra
    ExecuteSQL5 "UPDATE HeThongTK set MaTC = MaSo where sohieu = '" + sohieu + "'"
   '/////////
    Dim sql_dem
      Dim rs_taikhoan1 As Recordset
      
    sql_dem = "select count(sohieu) as f1 from hethongtk where left(sohieu,'" + str(Len(sohieu) - 1) + "') = '" + Left(sohieu, Len(sohieu) - 1) + "'"
    
        Set rs_taikhoan1 = DBKetoan.OpenRecordset(sql_dem, dbOpenSnapshot)
        If rs_taikhoan1.recordCount > 0 Then
         If (Int(rs_taikhoan1!f1) <= 2) Then
         KiemTraTaiKhoan
         End If
        End If
        
    WSpace.CommitTrans
End Function
'======================================================================================
' Thu tuc cap nhat lai cac thay doi cua tai khoan
' Tra ve true neu co cap nhat lai cac so hieu tai khoan con
'======================================================================================
Public Function CapNhatTk() As Integer
    Dim sql As String
    Dim taikhoan As New ClsTaikhoan
    Dim dn As Double, DC As Double
    
    ' LÊy th«ng tin tµi kho¶n cò
    taikhoan.InitTaikhoanMaSo MaSo
    
    If MaNT = 0 And sohieu <> taikhoan.sohieu Then
        sql = "SELECT DISTINCTROW Count(MaSo) AS F1 FROM HethongTK WHERE SoHieu='" + sohieu + "' AND MaNT=0"
        If SelectSQL(sql) > 0 Then
            CapNhatTk = -1
            GoTo KT
        End If
    End If
    If Len(TenE) = 0 Then TenE = "..."
    WSpace.BeginTrans
    ' Cap nhat thong tin tai khoan moi
    sql = "UPDATE HethongTk SET SoHieu='" + sohieu + "',Ten='" + Ten + "',Kieu=" + CStr(kieu) _
         + ", GhiChu = '" + GhiChu + "',TenDA='" + TenDA + "',NhomDA='" + NhomDA _
        + "',DiaDiem='" + DiaDiem + "',DuToan=" + CStr(DuToan) + ",Von1=" + CStr(Von1) + ",Von2=" + CStr(Von2) + ",Von3=" + CStr(Von3) + ",NgayKC=#" + Format(NgayKC, Mask_DB) _
        + "#,NgayHT=#" + Format(NgayHT, Mask_DB) + "#" + IIf(pSongNgu, ",TenE='" + TenE + "'", "") + " WHERE SoHieu = '" + taikhoan.sohieu + "'"
         
    If ExecuteSQL5(sql) <> 0 Then
        WSpace.Rollback
        CapNhatTk = -1
        GoTo KT
    End If
    
    sql = "UPDATE HethongTk SET MaNT=" + CStr(MaNT) + ",SoDuMax=" + DoiDau(SoDuMax) + ",SoDuMin=" _
        + DoiDau(SoDuMin) + ",TK_ID=" + CStr(tk_id) + " WHERE MaSo=" + CStr(MaSo)
    ExecuteSQL5 sql
    
    ' Neu co thay doi so hieu thi thay doi so hieu cac tai khoan con
    If sohieu <> taikhoan.sohieu And taikhoan.tkcon <> 0 Then
        ExecuteSQL5 "UPDATE HethongTK SET SoHieu='" + sohieu + "' + Right(HethongTK.SoHieu, Len(HethongTK.SoHieu) - " + CStr(Len(taikhoan.sohieu)) + ")" _
                + " WHERE SoHieu LIKE '" + taikhoan.sohieu + "*'"
    End If
    CapNhatTk = 0
    
    If taikhoan.PSNLK <> PSNLK Or taikhoan.PSCLK <> PSCLK Or taikhoan.NoDauKy <> NoDauKy Or taikhoan.CoDauKy <> CoDauKy Or taikhoan.NTDauKy <> NTDauKy Then
        TinhSoDu 0, 0, -1, 0
    End If
    WSpace.CommitTrans
KT:
    Set taikhoan = Nothing
End Function
'======================================================================================
' Thu tuc xoa tai khoan khong phai la tai khoan tai chinh
'======================================================================================
Public Function XoaTk() As Integer
    Dim chuyenps As Boolean, sql As String, i As Integer
    
    If (MaTC = 0 Or MaTC = MaSo) And pXuLyTKTC = 0 Then
        If MsgBox("B¹n ch¾c ch¾n xo¸ tµi kho¶n chuÈn?", vbCritical + vbYesNo, App.ProductName) = vbNo Then Exit Function
        XoaTk = -1      'MsgBox "H·y xo¸ hÕt chi tiÕt cña tµi kho¶n !", 48, "ABC"
       ' Exit Function
    End If
    
    If tkcon > 0 Then
        XoaTk = -2      'MsgBox "H·y xo¸ hÕt chi tiÕt cña tµi kho¶n !", 48, "ABC"
        Exit Function
    End If
    
    If TkCoPS(0, 0) Then
        sql = "SELECT DISTINCTROW Count(MaSo) As F1 FROM HethongTK WHERE MaNT <=0 AND TKCha0 = " + CStr(TkCha0)
        If SelectSQL(sql) > 1 Then
            
            XoaTk = -3
            Exit Function
        Else
            
            chuyenps = True
        End If
    Else
        sql = "SELECT DISTINCTROW Count(MaSo) As F1 FROM TP154 WHERE MaTK = " + CStr(MaSo)
        If (sohieu Like ShTkSPDo + "*") And pDTTP <> 0 And SelectSQL(sql) > 0 Then
            sql = "SELECT DISTINCTROW Count(MaSo) As F1 FROM HethongTK WHERE TKCha0 = " + CStr(TkCha0)
            If SelectSQL(sql) > 1 Then
                XoaTk = -3
                Exit Function
            Else
                chuyenps = True
            End If
        Else
            chuyenps = False
        End If
    End If
    
    If MaNT < 0 Then
        XoaTk = -4
        Exit Function
    End If
    
    If tk_id = TKVT_ID Then
        sql = "DELETE * FROM TonKho WHERE Luong_0=0 AND Tien_0=0 "
        For i = 1 To 12
            sql = sql + " AND Luong_Nhap_" + CStr(i) + "=0 AND Tien_Nhap_" + CStr(i) + "=0 "
        Next
        ExecuteSQL5 sql
    End If
    
    If tk_id = TKCNKH_ID Or tk_id = TKCNPT_ID Then
        sql = "DELETE * FROM SoDuKhachHang WHERE DuNo_0=0 AND DuCo_0=0 AND DuNT_0=0 "
        For i = 1 To 12
            sql = sql + " AND No_" + CStr(i) + "=0 AND Co_" + CStr(i) + "=0"
        Next
        ExecuteSQL5 sql
    End If
    
    WSpace.BeginTrans
    
    NoDauKy = 0
    CoDauKy = 0
    NTDauKy = 0
    TinhSoDu 0, 0, -1, 0
    
    ' Xoa tai khoan
    If ExecuteSQL5("DELETE FROM HethongTk WHERE MaSo=" + CStr(MaSo)) <> 0 Then
        XoaTk = -5
        Exit Function
    End If
    If MaNT > 0 Then
        sql = "SELECT DISTINCTROW Count(MaSo) AS F1 FROM HethongTK WHERE SoHieu='" + sohieu + "' AND MaNT>0"
        If SelectSQL(sql) = 0 Then
            ExecuteSQL5 "UPDATE HethongTk SET MaNT = 0 WHERE SoHieu='" + sohieu + "' AND MaNT=-1"
        End If
    End If
    
    If MaNT = 0 Then
        ' Neu tai khoan cha khong con chi tiet thi cap nhat lai
        sql = "SELECT DISTINCTROW Count(MaSo) AS F1 FROM HethongTK WHERE TkCha0 = " + CStr(TkCha0)
        If SelectSQL(sql) = 0 Then
            ExecuteSQL5 "UPDATE HethongTk SET TkCon = 0,MaTC=IIF(" + CStr(MaTC = MaSo) + ",MaSo,MaTC) WHERE MaSo=" + CStr(TkCha0)
            If chuyenps Then
                ' Thay doi ma tai khoan trong chung tu
                ExecuteSQL5 "UPDATE ThongSo SET MaDTSD = " + CStr(TkCha0) + " WHERE MaDTSD = " + CStr(MaSo)
                ExecuteSQL5 "UPDATE ChungTu SET MaTkNo=" + CStr(TkCha0) + " WHERE MaTkNo=" + CStr(MaSo)
                ExecuteSQL5 "UPDATE ChungTu SET MaTkCo=" + CStr(TkCha0) + " WHERE MaTkCo=" + CStr(MaSo)
                ExecuteSQL5 "UPDATE ChungTuP SET MaTkNo=" + CStr(TkCha0) + " WHERE MaTkNo=" + CStr(MaSo)
                ExecuteSQL5 "UPDATE ChungTuP SET MaTkCo=" + CStr(TkCha0) + " WHERE MaTkCo=" + CStr(MaSo)
                ExecuteSQL5 "UPDATE TonKho SET MaTaiKhoan=" + CStr(TkCha0) + " WHERE MaTaiKhoan=" + CStr(MaSo)
                ExecuteSQL5 "UPDATE KiemKe SET MaTaiKhoan=" + CStr(TkCha0) + " WHERE MaTaiKhoan=" + CStr(MaSo)
                ExecuteSQL5 "UPDATE TP154 SET MaTK=" + CStr(TkCha0) + " WHERE MaTK=" + CStr(MaSo)
                If TruongDaCo("ThanhPham", "MaTKCP") Then
                    ExecuteSQL5 "UPDATE ThanhPham SET MaTKCP=" + CStr(TkCha0) + " WHERE MaTKCP=" + CStr(MaSo)
                    ExecuteSQL5 "UPDATE ThanhPham SET MaTKNC=" + CStr(TkCha0) + " WHERE MaTKNC=" + CStr(MaSo)
                    ExecuteSQL5 "UPDATE ThanhPham SET MaTKSX=" + CStr(TkCha0) + " WHERE MaTKSX=" + CStr(MaSo)
                    ExecuteSQL5 "UPDATE ThanhPham SET MaTKKH=" + CStr(TkCha0) + " WHERE MaTKKH=" + CStr(MaSo)
                End If
            End If
        End If
    End If
    WSpace.CommitTrans
    XoaTk = 0
End Function
'======================================================================================
' Ham kiem tra tµi kho¶n ®· cã sè ph¸t sinh
'======================================================================================
Public Function TkCoPS(tdau As Integer, tcuoi As Integer) As Boolean
    Dim sql1 As String, sql2 As String, sqlw As String
    
    If tcuoi > 0 Then sqlw = " AND " + WThang("ThangCT", tdau, tcuoi)
    sql1 = "SELECT DISTINCTROW Count(ChungTu.MaSo) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '" + sohieu + "*'" + sqlw
    sql2 = "SELECT DISTINCTROW Count(ChungTu.MaSo) AS F1 FROM " + ChungTu2TKNC(1) + " WHERE HethongTK.SoHieu LIKE '" + sohieu + "*'" + sqlw
    TkCoPS = (SelectSQL(sql1) > 0 Or SelectSQL(sql2) > 0)
End Function
'======================================================================================
Public Function TkCoPSN(ndau As Date, ncuoi As Date) As Boolean
    Dim sql1 As String, sql2 As String
    
    sql1 = "SELECT DISTINCTROW Count(ChungTu.MaSo) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '" + sohieu + "*'" + " AND " + WNgay("NgayGS", ndau, ncuoi)
    sql2 = "SELECT DISTINCTROW Count(ChungTu.MaSo) AS F1 FROM " + ChungTu2TKNC(1) + " WHERE HethongTK.SoHieu LIKE '" + sohieu + "*'" + " AND " + WNgay("NgayGS", ndau, ncuoi)
    TkCoPSN = (SelectSQL(sql1) > 0 Or SelectSQL(sql2) > 0)
End Function

' Ham kiem tra so phat sinh co hop le (canh cao neu khong hop le)
' SoPS:     So phat sinh
' LoaiPs:   Loai phat sinh
'======================================================================================
Public Sub KtraPhatsinh(thang As Integer, sops As Double, loaips As Integer)
    Dim ckn As Double, ckc As Double, cknt As Double
    Dim SoDuMoi As Double
    
    If SoDuMax = 0 And SoDuMin = 0 Then Exit Sub
    SoDuTK thang, ckn, ckc, cknt
    
    If loaips < 0 Then
        SoDuMoi = IIf(kieu < 0, ckn - ckc + sops, ckc - ckn - sops)
    Else
        SoDuMoi = IIf(kieu < 0, ckn - ckc - sops, ckc - ckn + sops)
    End If
      
    If SoDuMoi > SoDuMax And SoDuMax > 0 Then
            MsgBox "§· v­ît qu¸ sè d­ cho phÐp !", vbExclamation, App.ProductName
    Else
            If SoDuMoi < SoDuMin And SoDuMin > 0 Then
                  MsgBox "§· xuèng d­íi sè d­ tèi thiÓu !", vbExclamation, App.ProductName
            End If
    End If
End Sub
'======================================================================================
' Thu tuc khoi tao object
' mtk:  Ma so tai khoan can tham chieu
'       =0 neu object rong
'======================================================================================
Public Sub InitTaikhoanMaSo(mtk As Long)
    Dim macha As Long
    If mtk > 0 Then
        Dim rs_taikhoan As Recordset
        Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT * FROM HethongTK WHERE MaSo=" + CStr(mtk), dbOpenSnapshot)
        If rs_taikhoan.recordCount = 0 Then GoTo it
        MaSo = mtk
        sohieu = rs_taikhoan!sohieu
        Ten = rs_taikhoan!Ten
        If pSongNgu Then TenE = rs_taikhoan!TenE
        cap = rs_taikhoan!cap
        kieu = rs_taikhoan!kieu
        loai = rs_taikhoan!loai
        NoDauKy = rs_taikhoan.Fields("DuNo_0")
        CoDauKy = rs_taikhoan.Fields("DuCo_0")
        NTDauKy = rs_taikhoan.Fields("DuNT_0")
        SoDuMax = rs_taikhoan!SoDuMax
        SoDuMin = rs_taikhoan!SoDuMin
        MaNT = rs_taikhoan!MaNT
        tkcon = rs_taikhoan!tkcon
        TkCha0 = rs_taikhoan!TkCha0
        TkCha1 = rs_taikhoan!TkCha1
        TkCha2 = rs_taikhoan!TkCha2
        TkCha3 = rs_taikhoan!TkCha3
        TkCha4 = rs_taikhoan!TkCha4
        TkCha5 = rs_taikhoan!TkCha5
        MaTC = rs_taikhoan!MaTC
        tk_id = rs_taikhoan!tk_id
        tk_id2 = rs_taikhoan!tk_id2
        GhiChu = rs_taikhoan!GhiChu
        TenDA = rs_taikhoan!TenDA
        NhomDA = rs_taikhoan!NhomDA
        DiaDiem = rs_taikhoan!DiaDiem
        DuToan = rs_taikhoan!DuToan
        Von1 = rs_taikhoan!Von1
        Von2 = rs_taikhoan!Von2
        Von3 = rs_taikhoan!Von3
        NgayKC = rs_taikhoan!NgayKC
        NgayHT = rs_taikhoan!NgayHT
        PSNLK = rs_taikhoan!PSNLK
        PSCLK = rs_taikhoan!PSCLK
        rs_taikhoan.Close
        Set rs_taikhoan = Nothing
    Else
it:
        MaSo = 0
        sohieu = ""
        Ten = ""
        TenE = "..."
        cap = 0
        kieu = 0
        loai = 0
        tkcon = 0
        TkCha0 = 0
        TkCha1 = 0
        TkCha2 = 0
        TkCha3 = 0
        TkCha4 = 0
        TkCha5 = 0
        NoDauKy = 0
        CoDauKy = 0
        SoDuMax = 0
        SoDuMin = 0
        MaNT = 0
        MaTC = 0
        tk_id = 0
        tk_id2 = 0
        GhiChu = "..."
        TenDA = "..."
        NhomDA = "..."
        DiaDiem = "..."
        DuToan = 0
        Von1 = 0
        Von2 = 0
        Von3 = 0
        NgayKC = Date
        NgayHT = Date
        PSNLK = 0
        PSCLK = 0
    End If
End Sub
'======================================================================================
' Thu tuc khoi tao object
' sohieu:  So hieu tai khoan can tham chieu
' Tra ve:  Ten tai khoan neu tim thay
'======================================================================================
Public Function InitTaikhoanSohieu(sohieu As String, Optional shct As String = "") As String
    Dim sql As String
    
    sql = "SELECT DISTINCTROW MaSo AS F1 FROM HethongTK WHERE " + IIf(Len(shct) > 0, " Left(SoHieu,Len('" + sohieu + "'))='" + sohieu + "' AND RIGHT(Sohieu," + CStr(Len(shct)) + ")='" + shct + "'", "SoHieu='" + sohieu + "'") + " ORDER BY MaNT DESC"
        
    InitTaikhoanMaSo SelectSQL(sql)
    InitTaikhoanSohieu = Ten
End Function
'======================================================================================
' SUB DanXuat
'======================================================================================
Public Sub DanXuat(doi_tuong As ClsTaikhoan)
    With doi_tuong
      .MaSo = MaSo
      .sohieu = sohieu
      .Ten = Ten
      .TenE = TenE
      .NoDauKy = NoDauKy
      .CoDauKy = CoDauKy
      .NTDauKy = NTDauKy
      ' Phan loai
      .cap = cap
      .kieu = kieu
      .loai = loai
      .MaTC = MaTC
      ' Cac doi tuong quan he
      .tkcon = tkcon
      .TkCha0 = TkCha0
      .TkCha1 = TkCha1
      .TkCha2 = TkCha2
      .TkCha3 = TkCha3
      .TkCha4 = TkCha4
      .TkCha5 = TkCha5
      .MaNT = MaNT
      .SoDuMax = SoDuMax
      .SoDuMin = SoDuMin
      .tk_id = tk_id
      .tk_id2 = tk_id2
      .GhiChu = GhiChu
      .TenDA = TenDA
      .NhomDA = NhomDA
      .DiaDiem = DiaDiem
      .DuToan = DuToan
      .Von1 = Von1
      .Von2 = Von2
      .Von3 = Von3
      .NgayKC = NgayKC
      .NgayHT = NgayHT
      .PSNLK = PSNLK
      .PSCLK = PSCLK
    End With
End Sub
'======================================================================================
' Thu tuc chuyen so du tu tai khoan cha xuong tai khoan con moi mo
'======================================================================================
Private Sub ChuyenSoDu()
    Dim rs_taikhoan As Recordset, i As Integer
    Dim st As String, th As String, cosodu As Boolean
    
    ' Chuyen so du xuong tai khoan con
    Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT * FROM HethongTK WHERE MaSo=" + CStr(TkCha0), dbOpenSnapshot)
    If rs_taikhoan!tkcon > 0 Then
        If rs_taikhoan!MaTC = 0 Then ExecuteSQL5 "Update HethongTK Set MaTC=MaSo Where Maso=" + CStr(MaSo)
        GoTo KetThuc
    End If
    
    ExecuteSQL5 "Update HethongTK Set TKCon=1,MaNT=0,MaTC=IIF(" + CStr(MaSo = MaTC) + ",0,MaTC) Where Maso=" + CStr(TkCha0)
        
    st = "UPDATE HethongTk SET DuNo_0=" + DoiDau(rs_taikhoan!DuNo_0) _
        + ",DuCo_0=" + DoiDau(rs_taikhoan!DuCo_0) + ",DuNT_0=" + DoiDau(rs_taikhoan!DuNT_0)
    For i = 1 To 12
        th = CStr(i)
        st = st + ",No_" + th + "=" + DoiDau(rs_taikhoan.Fields("No_" + th)) + ",Co_" + th + "=" + DoiDau(rs_taikhoan.Fields("Co_" + th)) _
        + ",No_" + th + "_NT=" + DoiDau(rs_taikhoan.Fields("No_" + th + "_NT")) + ",Co_" + th + "_NT=" + DoiDau(rs_taikhoan.Fields("Co_" + th + "_NT")) _
        + ", DuNo_" + th + " = " + DoiDau(rs_taikhoan.Fields("DuNo_" + th)) _
        + ", DuCo_" + th + " = " + DoiDau(rs_taikhoan.Fields("DuCo_" + th)) + ", DuNT_" + th + " = " + DoiDau(rs_taikhoan.Fields("DuNT_" + th))
    Next
    st = st + " WHERE MaSo=" + CStr(MaSo)
    
    ExecuteSQL5 st
        
    ' Thay doi ma tai khoan trong chung tu
    ExecuteSQL5 "UPDATE ThongSo SET MaDTSD = " + CStr(MaSo) + " WHERE MaDTSD = " + CStr(TkCha0)
    ExecuteSQL5 "UPDATE ChungTu SET MaTkNo=" + CStr(MaSo) + " WHERE MaTkNo=" + CStr(TkCha0)
    ExecuteSQL5 "UPDATE ChungTu SET MaTkCo=" + CStr(MaSo) + " WHERE MaTkCo=" + CStr(TkCha0)
    ExecuteSQL5 "UPDATE ChungTuP SET MaTkNo=" + CStr(MaSo) + " WHERE MaTkNo=" + CStr(TkCha0)
    ExecuteSQL5 "UPDATE ChungTuP SET MaTkCo=" + CStr(MaSo) + " WHERE MaTkCo=" + CStr(TkCha0)
    ExecuteSQL5 "UPDATE TonKho SET MaTaiKhoan=" + CStr(MaSo) + " WHERE MaTaiKhoan=" + CStr(TkCha0)
    ExecuteSQL5 "UPDATE KiemKe SET MaTaiKhoan=" + CStr(MaSo) + " WHERE MaTaiKhoan=" + CStr(TkCha0)
    ExecuteSQL5 "UPDATE TP154 SET MaTK=" + CStr(MaSo) + " WHERE MaTK=" + CStr(TkCha0)
    If TruongDaCo("ThanhPham", "MaTKCP") Then
        ExecuteSQL5 "UPDATE ThanhPham SET MaTKCP=" + CStr(MaSo) + " WHERE MaTKCP=" + CStr(TkCha0)
        ExecuteSQL5 "UPDATE ThanhPham SET MaTKNC=" + CStr(MaSo) + " WHERE MaTKNC=" + CStr(TkCha0)
        ExecuteSQL5 "UPDATE ThanhPham SET MaTKSX=" + CStr(MaSo) + " WHERE MaTKSX=" + CStr(TkCha0)
        ExecuteSQL5 "UPDATE ThanhPham SET MaTKKH=" + CStr(MaSo) + " WHERE MaTKKH=" + CStr(TkCha0)
    End If
KetThuc:
    rs_taikhoan.Close
    Set rs_taikhoan = Nothing
End Sub
'=========================================================================================================================================
' Thu tuc tra ve so phat sinh cua tai khoan
'=========================================================================================================================================
Public Sub SoPhatSinh(tdau As Integer, tcuoi As Integer, psn As Double, psc As Double)
    Dim i As Integer, th As String, sqln As String, sqlc As String, sql As String
        
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        th = CStr(i)
        sqln = sqln + " + No_" + th
        sqlc = sqlc + " + Co_" + th
    Next
    
    sql = "SELECT DISTINCTROW SUM(" + sqln + ") AS F1,SUM(" + sqlc + ") AS F2 FROM HethongTK " _
        & "WHERE " + IIf(tkcon > 0 Or MaNT <> 0, "(SoHieu LIKE '" + sohieu + "*' AND TKCon = 0)", "MaSo = " + CStr(MaSo))
    
    psn = SelectSQL(sql, psc)
End Sub
'======================================================================================
' Thñ tôc tÝnh l¹i sè d­ cña tµi kho¶n cÊp d­íi cïng
'======================================================================================
Public Sub TinhSoDu(thang As Integer, sops As Double, loaips As Integer, sopsnt As Double)
    Dim i As Integer, rs_tk As Recordset, sql As String, sqlnt As String, j As Integer, k As Integer, st As String
    Dim dunocu(0 To 12) As Double, ducocu(0 To 12) As Double
    Dim duno(0 To 12) As Double, duco(0 To 12) As Double, dunt As Double, duntmoi As Double
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT * FROM HethongTK WHERE MaSo=" + CStr(MaSo), dbOpenSnapshot)
    
    ' Ghi ph¸t sinh
    If thang > 0 Then
        st = CStr(CThangDB(thang))
        If loaips < 0 Then
            sql = "UPDATE HethongTK SET No_" + st + " = No_" + st + " + " + DoiDau(sops)
        Else
            sql = "UPDATE HethongTK SET Co_" + st + " = Co_" + st + " + " + DoiDau(sops)
        End If
        j = CThangDB(ThangTruoc(thang))
    Else
        sql = "UPDATE HethongTK SET DuNo_0 = DuNo_0 + " + DoiDau(NoDauKy - rs_tk!DuNo_0) + ", DuCo_0 = DuCo_0 + " + DoiDau(CoDauKy - rs_tk!DuCo_0)
        j = 0
    End If
    
    ' Ghi sè d­ cò
    For i = j To 12
        dunocu(i) = rs_tk.Fields("DuNo_" + CStr(i))
        ducocu(i) = rs_tk.Fields("DuCo_" + CStr(i))
    Next
    
    j = j + 1
    
    ' X¸c ®Þnh sè d­ míi
    If thang > 0 Then
        i = CThangDB(thang)
        duno(i) = dunocu(i - 1) - ducocu(i - 1) + rs_tk.Fields("No_" + CStr(i)) - rs_tk.Fields("Co_" + CStr(i)) + IIf(loaips < 0, sops, -sops)
        If duno(i) < 0 Then
            duco(i) = -duno(i)
            duno(i) = 0
        Else
            duco(i) = 0
        End If
        k = i + 1
    Else
        duno(0) = NoDauKy
        duco(0) = CoDauKy
        k = 1
    End If
    
    For i = k To 12
        duno(i) = duno(i - 1) - duco(i - 1) + rs_tk.Fields("No_" + CStr(i)) - rs_tk.Fields("Co_" + CStr(i))
        If duno(i) < 0 Then
            duco(i) = -duno(i)
            duno(i) = 0
        Else
            duco(i) = 0
        End If
    Next
    
    ' Ghi sè d­ míi
    For i = j To 12
        st = CStr(i)
        sql = sql + ", DuNo_" + st + " = DuNo_" + st + " + " + DoiDau(duno(i) - dunocu(i)) + ", DuCo_" + st + " = DuCo_" + st + " + " + DoiDau(duco(i) - ducocu(i))
    Next
    
    If MaNT > 0 Then
        If thang > 0 Then
            st = DoiDau(sopsnt)
            If loaips < 0 Then
                sqlnt = "UPDATE HethongTK SET No_" + CStr(CThangDB(thang)) + "_NT = No_" + CStr(CThangDB(thang)) + "_NT + " + st
            Else
                sqlnt = "UPDATE HethongTK SET Co_" + CStr(CThangDB(thang)) + "_NT = Co_" + CStr(CThangDB(thang)) + "_NT + " + st
            End If
        Else
            sqlnt = "UPDATE HethongTK SET DuNT_0 =  " + DoiDau(NTDauKy)
        End If
        ExecuteSQL5 sqlnt + " WHERE MaSo = " + CStr(MaSo)
        sqlnt = "UPDATE HethongTK SET "
          For i = IIf(thang > 0, CThangDB(thang), 1) To 12
                ExecuteSQL5 sqlnt + "DuNT_" + CStr(i) + "=ABS(DuNT_" + CStr(i - 1) + "+IIF(DuNo_" + CStr(i - 1) + "-DuCo_" + CStr(i - 1) + ">=0,No_" + CStr(i) + "_NT-Co_" + CStr(i) + "_NT,Co_" + CStr(i) + "_NT-No_" + CStr(i) + "_NT)) WHERE MaNT<>0"
          Next
    End If
    rs_tk.Close
    
    ExecuteSQL5 sql + " WHERE MaSo = " + CStr(MaSo) + " OR MaSo = " + CStr(TkCha0) + " OR MaSo = " + CStr(TkCha1) + " OR MaSo = " + CStr(TkCha2) + " OR MaSo = " + CStr(TkCha3) + " OR MaSo = " + CStr(TkCha4) + " OR MaSo = " + CStr(TkCha5)
    If tk_id2 <> TKLT_ID And cap > 0 Then
        sql = "UPDATE HethongTK SET DuNo_1=IIF(DuNo_1>=DuCo_1,DuNo_1-DuCo_1,0),DuCo_1=IIF(DuNo_1<DuCo_1,DuCo_1-DuNo_1,0)"
        For i = 2 To 12
            sql = sql + ",DuNo_" + CStr(i) + "=IIF(DuNo_" + CStr(i) + ">=DuCo_" + CStr(i) + ",DuNo_" + CStr(i) + "-DuCo_" + CStr(i) + ",0),DuCo_" + CStr(i) + "=IIF(DuNo_" + CStr(i) + "<DuCo_" + CStr(i) + ",DuCo_" + CStr(i) + "-DuNo_" + CStr(i) + ",0)"
        Next
        ExecuteSQL5 sql + " WHERE MaSo = " + CStr(MaSo) + " OR MaSo = " + CStr(TkCha0) + " OR MaSo = " + CStr(TkCha1) + " OR MaSo = " + CStr(TkCha2) + " OR MaSo = " + CStr(TkCha3) + " OR MaSo = " + CStr(TkCha4) + " OR MaSo = " + CStr(TkCha5)
    End If
    Set rs_tk = Nothing
End Sub
'======================================================================================
' Hµm tr¶ vÒ sè tµi kho¶n chi tiÕt cña mét tµi kho¶n
'======================================================================================
Public Function DanhsachTkCon(MaTkCon() As Long) As Integer
    Dim rs_tk As Recordset, i As Integer
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE SoHieu LIKE '" + sohieu _
        + "*' AND TkCon=0 ORDER BY SoHieu DESC", dbOpenSnapshot)
    rs_tk.MoveLast
    DanhsachTkCon = rs_tk.recordCount
    ReDim MaTkCon(1 To DanhsachTkCon) As Long
    i = 0
    Do While Not rs_tk.BOF
        i = i + 1
        MaTkCon(i) = rs_tk!MaSo
        rs_tk.MovePrevious
    Loop
    rs_tk.Close
    Set rs_tk = Nothing
End Function
'======================================================================================
' Thñ tôc ghi l¹i sè d­ ®Çu n¨m cña c¸c tµi kho¶n tµi chÝnh l¸ khi th¸ng ®Çu kú > 1
'======================================================================================
Public Sub GhiDauNam(duno As Double, duco As Double)
    ExecuteSQL5 "UPDATE HethongTK SET DuNo_0 = " + DoiDau(duno) + ", DuCo_0 = " + DoiDau(duco) + " WHERE MaSo = " + CStr(MaSo) + " OR MaSo = " + CStr(TkCha0) + " OR MaSo = " + CStr(TkCha1) + " OR MaSo = " + CStr(TkCha2) + " OR MaSo = " + CStr(TkCha3) + " OR MaSo = " + CStr(TkCha4) + " OR MaSo = " + CStr(TkCha5)
End Sub
'======================================================================================
' Ham tra ve so dau, cuoi ky cua tai khoan
' thang: Thang tinh so du, neu truoc thang dau ky chi tra ve so du tai khoan
'======================================================================================
Public Sub SoDuTK(thang As Integer, ckn As Double, ckc As Double, cknt As Double, Optional vnd As Integer = 0)
    Dim th As String, rs_taikhoan As Recordset
    
    ckn = 0
    ckc = 0
    cknt = 0
    
    th = CThangDB(thang)
        ' Tinh tong so du dau nam cua cac tai khoan chi tiet
    Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT DuNo_" + th + " As Dn, DuCo_" + th + " As Dc, DuNT_" + th + " As Dnt FROM HethongTk " _
        & "WHERE " + IIf(tkcon > 0 Or (MaNT <= 0 And vnd = 0), "(SoHieu LIKE '" + sohieu + "*' AND TKCon = 0)", "MaSo = " + CStr(MaSo)), dbOpenSnapshot)
    Do While Not rs_taikhoan.EOF
        ckn = ckn + rs_taikhoan!dn
        ckc = ckc + rs_taikhoan!DC
        cknt = cknt + rs_taikhoan!dnt
        rs_taikhoan.MoveNext
    Loop
    rs_taikhoan.Close
    Set rs_taikhoan = Nothing
End Sub

'======================================================================================
' Thu tuc tra ve so phat sinh cua tai khoan
'======================================================================================
Public Sub SoPhatSinhN(ndau As Date, ncuoi As Date, psn As Double, ntno As Double, psc As Double, ntco As Double)
Dim sql As String
    sql = "SELECT DISTINCTROW Sum(SoPS) AS F1, Sum(SoPS2No) AS F2 FROM " + ChungTu2TKNC(-1) _
        & "WHERE HethongTK.SoHieu LIKE '" + sohieu + "*' AND " + WNgay("NgayGS", ndau, ncuoi)
    psn = SelectSQL(sql, ntno)
    sql = "SELECT DISTINCTROW Sum(SoPS) AS F1, Sum(SoPS2Co) AS F2 FROM " + ChungTu2TKNC(1) _
        & "WHERE HethongTK.SoHieu LIKE '" + sohieu + "*' AND " + WNgay("NgayGS", ndau, ncuoi)
    psc = SelectSQL(sql, ntco)
End Sub
'======================================================================================
' Thu tuc tra ve so du cua tai khoan theo ngay
'======================================================================================
Public Function SoDuNgay(ngay As Date, duno As Double, duco As Double, dunt As Double) As Double
    Dim ckn As Double, ckc As Double, ntno As Double, ntco As Double
    
    SoDuTK 0, duno, duco, dunt
    SoPhatSinhN CVDate("1/1/80"), ngay, ckn, ntno, ckc, ntco
    duno = duno - duco + ckn - ckc
    If duno > 0 Then
        duco = 0
        dunt = dunt + ntno - ntco
    Else
        If duno < 0 Then
            duco = -duno
            duno = 0
            dunt = dunt - ntno + ntco
        End If
    End If
End Function

Public Function SHChiTiet(Optional ct As Integer = 0) As String
    Dim ltc As Integer
    
    If MaTC > 0 Then
        ltc = SelectSQL("SELECT Len(SoHieu) AS F1 FROM HethongTK WHERE MaSo=" + CStr(IIf(ct = 0, MaTC, TkCha0)))
        SHChiTiet = Right(sohieu, Len(sohieu) - ltc)
    Else
        SHChiTiet = ""
    End If
End Function

Private Sub Class_Initialize()
    TenE = "..."
End Sub

Public Sub LayMaNT(mnt As Long)
    Dim id As Long
    
    id = SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE SoHieu='" + sohieu + "' AND MaNT=" + CStr(mnt))
    If id <> 0 Then InitTaikhoanMaSo id
End Sub

Public Function MaxCap() As Integer
    MaxCap = SelectSQL("SELECT Max(Cap) AS F1 FROM HethongTK WHERE SoHieu LIKE '" + sohieu + "*'")
End Function

Public Function ChuyenChiTietSangDoiTuong() As Boolean
    Dim sql As String
    
    ChuyenChiTietSangDoiTuong = False
    If tkcon = 0 Or MaxCap <> cap + 1 Or loai > 4 Or tk_id <> 0 Or MaTC = 0 Then Exit Function
    If MsgBox("ChuyÓn c¸c tµi kho¶n chi tiÕt thµnh ®èi t­îng c«ng nî ?", vbYesNo + vbExclamation) = vbNo Then Exit Function
    
    If SelectSQL("SELECT Max(MaNT) AS F1 FROM HethongTK WHERE SoHieu LIKE '" + sohieu + "*'") > 0 Then Exit Function
    If SelectSQL("SELECT TOP 1 MaSo AS F1 FROM KhachHang WHERE SoHieu LIKE '" + sohieu + "*'") > 0 Then Exit Function
    
    ExecuteSQL5 "INSERT INTO PhanLoaiKhachHang (MaSo,SoHieu,TenPhanLoai,Cap) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhanLoaiKhachHang") + 1) + ",'" + sohieu + "','" + Ten + "',1)"
    ExecuteSQL5 "INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu,Ten,GhiChu,MaNT) SELECT " + CStr(Lng_MaxValue("MaSo", "KhachHang") + 1) + "," + CStr(SoHieu2MaSo(sohieu, "PhanLoaiKhachHang")) + " AS MPL, SoHieu, Ten, GhiChu,MaNT FROM HethongTK WHERE SoHieu LIKE '" + sohieu + "*' AND Cap=" + CStr(cap + 1)
    ExecuteSQL5 "INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan, MaKhachHang, DuNo_0, DuCo_0, DuNT_0) SELECT " + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + ",TKCha0, KhachHang.MaSo, DuNo_0, DuCo_0, DuNT_0 FROM HethongTK INNER JOIN KhachHang ON HethongTK.SoHieu=KhachHang.SoHieu WHERE HethongTK.SoHieu LIKE '" + sohieu + "*' AND Cap=" + CStr(cap + 1) + " AND (DuNo_0<>0 OR DuCo_0<>0 OR DuNT_0<>0)"
    
    ExecuteSQL5 "UPDATE (" + ChungTu2TKNC(-1) + ") INNER JOIN KhachHang ON HethongTK.SoHieu=KhachHang.SoHieu SET MaKH=KhachHang.MaSo, MaTKNo=" + CStr(MaSo) + " WHERE HethongTK.SoHieu LIKE '" + sohieu + "*'"
    ExecuteSQL5 "UPDATE (" + ChungTu2TKNC(1) + ") INNER JOIN KhachHang ON HethongTK.SoHieu=KhachHang.SoHieu SET MaKHC=KhachHang.MaSo, MaTKCo=" + CStr(MaSo) + " WHERE HethongTK.SoHieu LIKE '" + sohieu + "*'"
    
    ExecuteSQL5 "UPDATE HethongTK SET TKCon=0, TK_ID=" + IIf(loai < 3, CStr(TKCNKH_ID), CStr(TKCNPT_ID)) + " WHERE MaSo=" + CStr(MaSo)
    ExecuteSQL5 "DELETE * FROM HethongTK WHERE SoHieu LIKE '" + sohieu + "*' AND Cap>" + CStr(cap)
    
    KiemTraTaiKhoan
    ChuyenChiTietSangDoiTuong = True
End Function

Public Function ThemTKCon(shct As String, tenct As String, tenct2 As String, tkid As Long, tkid2 As Long) As Long
    Dim TK As ClsTaikhoan
    
    If tk_id <> TSCD_ID Or (Not FADetail) Then
        Set TK = New ClsTaikhoan
        DanXuat TK
        TK.sohieu = sohieu + shct
        TK.Ten = tenct
        TK.TenE = tenct2
        TK.TkCha5 = TkCha4
        TK.TkCha4 = TkCha3
        TK.TkCha3 = TkCha2
        TK.TkCha2 = TkCha1
        TK.TkCha1 = TkCha0
        TK.TkCha0 = MaSo
        TK.tkcon = 0
        TK.MaSo = 0
        If tkid <> 0 Then TK.tk_id = tkid
        If tkid2 <> 0 Then TK.tk_id2 = tkid2
        TK.cap = cap + 1
        TK.NoDauKy = 0
        TK.CoDauKy = 0
        TK.NTDauKy = 0
        TK.ThemTk
        ThemTKCon = TK.MaSo
        Set TK = Nothing
    Else
        ThemTKCon = MaSo
    End If
End Function

Public Function GhepCanh(tkc As ClsTaikhoan) As Integer
    Dim sql As String, X As String, i As Integer, j As Integer
    Dim tkcha(0 To max_level - 1)
    
    GhepCanh = -1
    If MaTC = 0 Or tk_id = TKCNKH_ID Or tk_id = TKCNPT_ID Or tk_id = TSCD_ID Or tk_id = TKVT_ID Or (TkCoPS(0, 0) And tkcon = 0 And tkc.tkcon > 0) Then Exit Function
    
    tkcha(0) = TkCha0
    tkcha(1) = TkCha1
    tkcha(2) = TkCha2
    tkcha(3) = TkCha3
    tkcha(4) = TkCha4
    tkcha(5) = TkCha5
    
    ' Dieu chinh ma so cha
    For i = tkc.cap To max_level
        sql = "TKCha" + CStr(i - tkc.cap) + "=" + CStr(MaSo)
        For j = i - tkc.cap + 1 To max_level - 1
            sql = sql + ",TKCha" + CStr(j) + "=" + CStr(tkcha(j - (i - tkc.cap + 1)))
        Next
        ExecuteSQL5 "UPDATE HethongTK SET " + sql + ",Cap=" + CStr(cap + 1 + i - tkc.cap) + " WHERE Cap=" + CStr(i) + " AND SoHieu LIKE '" + tkc.sohieu + "*'"
        If DBKetoan.RecordsAffected = 0 Then Exit For
    Next
    
    ' Dieu chinh so hieu
    If tkcon > 0 Then
        sql = tkc.SHChiTiet(1)
        If SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE TKCha0=" + CStr(MaSo) + " AND Right(SoHieu," + CStr(Len(sql)) + ")='" + sql + "'") > 0 Then X = "X"
    End If
    i = Len(MaSo2SoHieu(tkc.TkCha0, "HethongTK"))
    ExecuteSQL5 "UPDATE HethongTK SET Loai=" + CStr(loai) + ",Kieu=" + CStr(kieu) + ",MaTC=" + CStr(MaTC) + ",SoHieu='" + sohieu + IIf(Len(X) > 0, X, "") + "'+Right(SoHieu,Len(SoHieu)-" + CStr(i) + ") WHERE SoHieu LIKE '" + tkc.sohieu + "*'"
        
    KiemTraTaiKhoan
    GhepCanh = 0
End Function

Public Function PSTuKhoiCong(ncuoi As Date, loaips As Integer) As Double
    Dim ps As Double
    
    ps = IIf(loaips < 0, PSNLK, PSCLK)
    If loaips < 0 Then
        ps = ps + SelectSQL("SELECT DISTINCTROW Sum(SoPS) AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTkNo = HethongTK.MaSo " _
            & "WHERE HethongTK.SoHieu LIKE '" + sohieu + "*' AND ChungTu.NgayGS <= #" + Format(ncuoi, Mask_DB) + "#")
    Else
        ps = ps + SelectSQL("SELECT DISTINCTROW Sum(SoPS) AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTkCo = HethongTK.MaSo " _
            & "WHERE HethongTK.SoHieu LIKE '" + sohieu + "*' AND ChungTu.NgayGS <= #" + Format(ncuoi, Mask_DB) + "#")
    End If
    PSTuKhoiCong = ps
End Function

Public Function PSLKNam(nam As Integer, loaips As Integer) As Double
    If TruongDaCo("HethongTK", "PSNLK" + CStr(nam)) Then
        PSLKNam = SelectSQL("SELECT " + IIf(loaips < 0, "PSNLK" + CStr(nam), "PSCLK" + CStr(nam)) + " AS F1 FROM HethongTK WHERE MaSo=" + CStr(MaSo))
    End If
End Function

Public Function ShCapDuoi(shx As String) As String
    Dim shc As String, i As Integer
    
    shc = sohieu
    For i = Len(sohieu) + 1 To Len(shx)
        shc = shc + Mid(shx, i, 1)
        If SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE SoHieu='" + shc + "'") > 0 Then
            ShCapDuoi = shc
            Exit Function
        End If
    Next
End Function

Public Function InitTaikhoanSohieuCT(sohieu As String, Optional shct As String = "", Optional cap As Integer = 0) As String
    Dim sql As String
    
    sql = "SELECT DISTINCTROW MaSo AS F1 FROM HethongTK WHERE SoHieu LIKE '" + sohieu + "*' " _
    + IIf(shct <> "", " AND RIGHT(Sohieu," + CStr(Len(shct)) + ")='" + shct + "'", "") + IIf(cap > 0, " AND Cap=" + CStr(cap), "") + " ORDER BY MaNT DESC"
        
    InitTaikhoanMaSo SelectSQL(sql)
    InitTaikhoanSohieuCT = Ten
End Function

Public Function CPDaPhanBo(ndau As Date, ncuoi As Date, tkpb As String) As Double
    CPDaPhanBo = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
        & " WHERE HethongTK.SoHieu LIKE '" + tkpb + "*' AND TK.SoHieu LIKE '" + sohieu + "*' AND " + WNgay("NgayGS", ndau, ncuoi))
End Function

Public Function ChoNhap() As Boolean
    If User_Right = 0 Then
        ChoNhap = True
    Else
        ChoNhap = False
        If Not TruongDaCo("HethongTK", "U" + CStr(UserID)) Then Exit Function
        ChoNhap = (SelectSQL("SELECT U" + CStr(UserID) + " AS F1 FROM HethongTK WHERE MaSo=" + CStr(MaSo)) > 0)
    End If
End Function

