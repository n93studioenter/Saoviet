Attribute VB_Name = "modTaiSan"
Option Explicit

Public Const NV_KHONG = 0
Public Const NV_TANG = 32
Public Const NV_GIAM = 33
Public Const NV_DGLAI = 34
Public Const NV_TKHAO = 35

Public Type tpGiaTri
      NG_NS As Double                                                    ' Nguyªn gi¸
      NG_TBS As Double
      NG_CNK As Double
      NG_TD As Double
      CL_NS As Double                                                      ' Gi¸ trÞ cßn l¹i
      CL_TBS As Double
      CL_CNK As Double
      CL_TD As Double
      KH_NS As Double                                                      ' KhÊu hao
      KH_TBS As Double
      KH_CNK As Double
      KH_TD As Double
End Type

Public Type tpPhatSinh
      TK_SoHieu As String
      TS_SoHieu As String
      PS_Loai As Integer
      PS_SoLg As Double
      ShTP As String
End Type

Public Const DT_QUOCGIA = 300
Public Const DT_DOITUONG = 301
Public Const DT_TINHTRANG = 302
' T¸c ®éng gi¶m tµi s¶n
Public Const TD_GIAM = 400
Public Const TD_KHOIPHUC = 401

Public Const KH_CO = 500
Public Const KH_KHONG = 501

Public Const DK_LOAI = 30
Public Const DK_NHOM = 31

Public GiaTri As tpGiaTri
Public pNghiepVu As Integer
Public pThangTacDong As Integer
Public pMaTaiSan As Long
Public pMaChungTu As Long
Public pGhichungtu As Integer

Public arPhatSinh() As tpPhatSinh              ' B¶n d·y c¸c dßng ph¸t sinh ®· ®­îc thµnh lËp s½n
Public parSoPS As Integer
'======================================================================================
' SUB CapNhatGiaTriTaiSan : TÝnh vµ cËp nhËt gi¸ trÞ cho tÊt c¶ c¸c tµi s¶n trong mét th¸ng (Quy t¾c
'                                                           xem thñ tôc TinhGiaTriTaiSan)
'                                      Tham sè : Th¸ng cÇn tÝnh gi¸ trÞ, Gauge Control ®Ó hiÓn thÞ sè % ®· hoµn thµnh
'                                       Sö dông : frmBaoCao vµ frmKhauHao
'======================================================================================
Public Sub CapNhatGiaTriTaiSan(thg As Integer, gauge_control As Object)
Dim rs_giatri As Recordset, i As Integer
     
      gauge_control.Max = 3
     'gauge_control.Refresh
      ' Tæng khÊu hao
      SetSQL "TongKhauHao", "SELECT DISTINCTROW Sum(ThongSo.KH_NS) AS TKH_NS, Sum(ThongSo.KH_TBS) AS TKH_TBS, Sum(ThongSo.KH_CNK) AS TKH_CNK, Sum(ThongSo.KH_TD) AS TKH_TD, ThongSo.MaTS " _
            & "FROM TaiSan RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS " _
            & "WHERE " + WThang("ThangTang", 0, CThangFR(thg)) + " AND " + WThang("ThangGiam", CThangFR(thg), 0) + " AND ThongSo.Thang <= " + CStr(thg) + " GROUP BY ThongSo.MaTS"
      gauge_control.Value = 1
     ' gauge_control.Refresh5
      ' L­îng t¨ng gi¶m
      SetSQL "TongGiaTri", "SELECT Sum(NG_NS) AS TNG_NS, Sum(NG_TBS) AS TNG_TBS, Sum(NG_CNK) AS TNG_CNK, Sum(NG_TD) AS TNG_TD, " _
            & "Sum(CL_NS) AS TCL_NS, Sum(CL_TBS) AS TCL_TBS, Sum(CL_CNK) AS TCL_CNK, Sum(CL_TD) AS TCL_TD, MaTS " _
            & "FROM CTTaiSan WHERE " + WThang("Thang", 0, CThangFR(thg)) + " GROUP BY MaTS"
    
      gauge_control.Value = 2
     ' gauge_control.Refresh
      ' Gi¸ trÞ tµi s¶n
      SetSQL "GiaTriTaiSan", "SELECT DISTINCTROW TNG_NS AS NG_NS, TNG_TBS AS NG_TBS, TNG_CNK AS NG_CNK, TNG_TD AS NG_TD, " _
            & "TCL_NS-TKH_NS AS CL_NS, TCL_TBS-TKH_TBS AS CL_TBS, TCL_CNK-TKH_CNK AS CL_CNK, TCL_TD-TKH_TD AS CL_TD, " _
            & "TongGiaTri.MaTS FROM TongKhauHao INNER JOIN TongGiaTri ON TongKhauHao.MaTS = TongGiaTri.MaTS"
      Set rs_giatri = DBKetoan.OpenRecordset("GiaTriTaiSan", dbOpenSnapshot)
     
      gauge_control.Value = 3
'**************
    '  gauge_control.Refresh
      On Error GoTo Err_NoCurrentRecord
            rs_giatri.MoveLast
      On Error GoTo 0
      gauge_control.Max = rs_giatri.RecordCount
      gauge_control.Value = 0
'**************88888
    '  gauge_control.Refresh
      rs_giatri.MoveFirst
      Do Until rs_giatri.EOF
            ExecuteSQL5 "UPDATE ThongSo SET NG_NS = " + DoiDau(rs_giatri!NG_NS) _
                  + ", NG_TBS = " + DoiDau(rs_giatri!NG_TBS) _
                  + ", NG_CNK = " + DoiDau(rs_giatri!NG_CNK) _
                  + ", NG_TD = " + DoiDau(rs_giatri!NG_TD) _
                  + ", CL_NS = " + DoiDau(RoundMoney(rs_giatri!CL_NS)) _
                  + ", CL_TBS = " + DoiDau(RoundMoney(rs_giatri!CL_TBS)) _
                  + ", CL_CNK = " + DoiDau(RoundMoney(rs_giatri!CL_CNK)) _
                  + ", CL_TD = " + DoiDau(RoundMoney(rs_giatri!CL_TD)) _
                  + " WHERE MaTS = " + CStr(rs_giatri!MaTS) + " AND Thang = " + CStr(thg)
            gauge_control.Value = gauge_control.Value + 1
           ' gauge_control.Refresh
            rs_giatri.MoveNext
      Loop
Err_NoCurrentRecord:
       'KiÓm tra vµ ®iÒu chØnh l¹i l­îng khÊu hao cho tÊt c¶ c¸c tµi s¶n trong th¸ng
      DieuChinhKhauHao thg
      rs_giatri.Close
      Set rs_giatri = Nothing
End Sub
'======================================================================================
' SUB ThanhLapPhatSinh : Thµnh lËp c¸c dßng ph¸t sinh ph¶n ¸nh nghiÖp vô kÕ to¸n t¸c ®éng lªn mét
'                                                     tµi s¶n. §Þnh kho¶n dùa trªn tµi kho¶n tµi s¶n vµ tµi kho¶n chi phÝ khÊu hao
'                                                     t­¬ng øng. Sè ph¸t sinh lÊy tõ biÕn chung GiaTri vµ lo¹i ph¸t sinh (nî hoÆc
'                                                     cã) ®­îc x¸c ®Þnh qua nghiÖp vô.
'                                                     Sè dßng ph¸t sinh ®· thµnh lËp ®­îc cho trong biÕn chung pSoPhatSinh
'                                Tham sè : M· nghiÖp vô, m· lo¹i cña tµi s¶n bÞ t¸c ®éng
'                                       Chó ý : Kh«ng t¹o ra dßng ph¸t sinh thÓ hiÖn l­îng t¨ng gi¶m khÊu hao nÕu nghiÖp
'                                                     vô lµ thay ®æi gi¸ trÞ tµi s¶n
'======================================================================================
Public Sub ThanhLapPhatSinh(nghiep_vu As Long, ma_tkts As Long)
Dim tong_ng As Double, tong_hm As Double
Dim sql As String
      parSoPS = 1
      ReDim arPhatSinh(0 To parSoPS) As tpPhatSinh
      ' TÝnh sè ph¸t sinh
      tong_ng = (GiaTri.NG_NS + GiaTri.NG_TBS + GiaTri.NG_CNK + GiaTri.NG_TD)
      'If nghiep_vu = NV_DGLAI Then
      '      tong_hm = 0
      'Else
            tong_hm = tong_ng - (GiaTri.CL_NS + GiaTri.CL_TBS + GiaTri.CL_CNK + GiaTri.CL_TD)
      'End If
      ' X¸c ®Þnh tµi kho¶n tµi s¶n
      sql = "SELECT SoHieu AS F1 FROM LoaiTaiSan WHERE MaSo = " _
                                                                                                                                                                          + CStr(ma_tkts)
      arPhatSinh(0).TK_SoHieu = CStr(SelectSQL(sql))
      arPhatSinh(0).TS_SoHieu = MaSo2SoHieu(pMaTaiSan, "TaiSan")
      ' Sè hiÖu cña tµi kho¶n chi phÝ khÊu hao x¸c ®Þnh qua lo¹i tµi s¶n
      arPhatSinh(1).TK_SoHieu = "214" + Mid(arPhatSinh(0).TK_SoHieu, 3, 1)
      arPhatSinh(0).PS_SoLg = tong_ng
      arPhatSinh(1).PS_SoLg = tong_hm
      ' Lo¹i  ph¸t sinh x¸c ®Þnh qua nghiÖp vô
      Select Case nghiep_vu
            Case NV_TANG
                  arPhatSinh(0).PS_Loai = -1
                  arPhatSinh(1).PS_Loai = 1
            Case NV_GIAM
                  arPhatSinh(0).PS_Loai = 1
                  arPhatSinh(1).PS_Loai = -1
            Case NV_DGLAI
                  arPhatSinh(0).PS_Loai = -1
                  arPhatSinh(1).PS_Loai = 1
                  'If tong_ng > 0 Then arPhatSinh(0).PS_Loai = -1 Else arPhatSinh(0).PS_Loai = 1
                  'If tong_hm > 0 Then arPhatSinh(1).PS_Loai = 1 Else arPhatSinh(1).PS_Loai = -1
      End Select
      ' Kh«ng chÊp nhËn sè ph¸t sinh nhá h¬n 0
      'If arPhatSinh(0).PS_SoLg < 0 Then arPhatSinh(0).PS_SoLg = -arPhatSinh(0).PS_SoLg
      'If arPhatSinh(1).PS_SoLg < 0 Then arPhatSinh(1).PS_SoLg = -arPhatSinh(1).PS_SoLg
End Sub
'======================================================================================
' SUB TinhGiaTriTaiSan : TÝnh gi¸ trÞ cña tµi s¶n t¹i mét thêi ®iÓm cho tr­íc dùa trªn th«ng tin ®· l­u
'                                                   trong c¸c chøng tõ cã liªn quan vµ l­îng khÊu hao hµng th¸ng.
'                                                   Gi¸ trÞ ®­îc tÝnh b»ng sè ®Çu kú (l­u trong chøng tõ ®Çu kú hoÆc chøng tõ
'                                                   t¨ng t­¬ng øng) céng tæng l­îng t¨ng gi¶m cho ®Õn th¸ng hiÖn t¹i (l­u trong
'                                                   c¸c chøng tõ t¨ng gi¶m) trõ ®i tæng l­îng khÊu hao cho ®Õn th¸ng hiÖn t¹i.
'                                                   KÕt qu¶ tr¶ vÒ ®­îc chøa trong biÕn chung GiaTri
'                              Tham sè : M· sè tµi s¶n, th¸ng cÇn tÝnh gi¸ trÞ, kiÓu tÝnh(cã trÝch khÊu hao hay kh«ng).
'                                    Chó ý : L­îng khÊu hao sÏ ®­îc kiÓm tra vµ kÕt qu¶ cã thÓ bÞ ®iÒu chØnh l¹i.
'                                                     - L­îng khÊu hao sÏ ®­îc ®Æt b»ng gi¸ trÞ cßn l¹i nÕu lín h¬n vµ tæng
'                                                       khÊu hao céng gi¸ trÞ lµ sè d­¬ng
'                                                     - Gi¸ trÞ cßn l¹i nÕu nhá h¬n 0 th× sÏ ®­îc ®Æt l¹i b»ng 0.
'                               Sö dông : Thñ tôc nµy ®­îc gäi tõ frmTaiSan, thñ tôc GiamTaiSan vµ ChiDinh tõ trong
'                                                  clsThongSo.
'======================================================================================
Public Sub TinhGiaTriTaiSan(ma_ts As Long, thg As Integer, khau_hao As Integer)
Dim rs_giatridau As Recordset
Dim rs_tongkhauhao As Recordset
Dim rs_khauhao As Recordset, sql As String
With GiaTri
      If ma_ts = 0 Then Exit Sub
      
    .NG_NS = 0
    .NG_TBS = 0
    .NG_CNK = 0
    .NG_TD = 0
    .CL_NS = 0
    .CL_TBS = 0
    .CL_CNK = 0
    .CL_TD = 0
    
      ' LÊy nguyªn gi¸ vµ gi¸ trÞ cßn l¹i cho ®Õn thêi ®iÓm hiÖn t¹i
      sql = "SELECT Sum(NG_NS) AS TNG_NS, Sum(NG_TBS) AS TNG_TBS, Sum(NG_CNK) AS TNG_CNK, Sum(NG_TD) AS TNG_TD, " _
            & "Sum(CL_NS) AS TCL_NS, Sum(CL_TBS) AS TCL_TBS, Sum(CL_CNK) AS TCL_CNK, Sum(CL_TD) AS TCL_TD " _
            & "FROM CTTaiSan WHERE MaTS = " + CStr(ma_ts) + " AND " + WThang("Thang", 0, thg)
      Set rs_giatridau = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      If IsNull(rs_giatridau!TNG_NS) Then
            sql = "SELECT (NG_NS) AS TNG_NS, (NG_TBS) AS TNG_TBS, (NG_CNK) AS TNG_CNK, (NG_TD) AS TNG_TD, " _
                  & "(CL_NS) AS TCL_NS, (CL_TBS) AS TCL_TBS, (CL_CNK) AS TCL_CNK, (CL_TD) AS TCL_TD " _
                  & "FROM ThongSo WHERE MaTS = " + CStr(ma_ts) + " AND Thang=" + CStr(CThangDB(thg))
            Set rs_giatridau = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      End If
      
        ' Nguyªn gi¸ tµi s¶n
        .NG_NS = rs_giatridau!TNG_NS
        .NG_TBS = rs_giatridau!TNG_TBS
        .NG_CNK = rs_giatridau!TNG_CNK
        .NG_TD = rs_giatridau!TNG_TD
    
If thg > 0 Then
      ' LÊy tæng l­îng khÊu hao cho ®Õn thêi ®iÓm hiÖn t¹i
      sql = "SELECT DISTINCTROW Sum(ThongSo.KH_NS) AS TKH_NS, Sum(ThongSo.KH_TBS) AS TKH_TBS, Sum(ThongSo.KH_CNK) AS TKH_CNK, Sum(ThongSo.KH_TD) AS TKH_TD " _
            & "FROM TaiSan RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS " _
            & "WHERE ThongSo.MaTS = " + CStr(ma_ts) _
            & " AND " + VC("ThongSo.Thang", "IIF(TaiSan.ThangTang=0," + CStr(pThangDauKy) + ",TaiSan.ThangTang)") _
            & " AND ThongSo.Thang <= " + CStr(CThangDB(thg))
      Set rs_tongkhauhao = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      If (Not IsNull(rs_tongkhauhao!TKH_NS)) And (Not IsNull(rs_giatridau!TCL_NS)) Then
            ' Gi¸ trÞ tµi s¶n
            .CL_NS = rs_giatridau!TCL_NS - (rs_tongkhauhao!TKH_NS)
            .CL_TBS = rs_giatridau!TCL_TBS - (rs_tongkhauhao!TKH_TBS)
            .CL_CNK = rs_giatridau!TCL_CNK - (rs_tongkhauhao!TKH_CNK)
            .CL_TD = rs_giatridau!TCL_TD - (rs_tongkhauhao!TKH_TD)
      End If
      rs_tongkhauhao.Close
      Set rs_tongkhauhao = Nothing
      ' L­îng khÊuhao
      sql = "SELECT KH_NS, KH_TBS, KH_CNK, KH_TD FROM ThongSo " _
                              & "WHERE Thang = " + CStr(CThangDB(thg)) + " AND MaTS = " + CStr(ma_ts)
      Set rs_khauhao = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      If rs_khauhao.RecordCount > 0 Then
            .KH_NS = rs_khauhao!KH_NS
            .KH_TBS = rs_khauhao!KH_TBS
            .KH_CNK = rs_khauhao!KH_CNK
            .KH_TD = rs_khauhao!KH_TD
      End If
      rs_khauhao.Close
      Set rs_khauhao = Nothing
      ' KiÓm tra l­îng khÊu hao vµ ®iÒu chØnh l¹i kÕt qu¶ (chó ý r»ng gi¸ trÞ tÝnh
      ' ®­îc ë ®©y lu«n lµ gi¸ trÞ khi ®· trÝch khÊu hao hay lµ gi¸ trÞ cuèi th¸ng)
      If .CL_NS < 0 Then
            If .KH_NS + .CL_NS > 0 Then .KH_NS = .KH_NS + .CL_NS Else .KH_NS = 0
            .CL_NS = 0
      End If
      If .CL_TBS < 0 Then
            If .KH_TBS + .CL_TBS > 0 Then .KH_TBS = .KH_TBS + .CL_TBS Else .KH_TBS = 0
            .CL_TBS = 0
      End If
      If .CL_CNK < 0 Then
            If .KH_CNK + .CL_CNK > 0 Then .KH_CNK = .KH_CNK + .CL_CNK Else .KH_CNK = 0
            .CL_CNK = 0
      End If
      If .CL_TD < 0 Then
            If .KH_TD + .CL_TD > 0 Then .KH_TD = .KH_TD + .CL_TD Else .KH_TD = 0
            .CL_TD = 0
      End If
      ' NÕu tÝnh gi¸ trÞ mµ kh«ng trõ ®i l­îng khÊu hao trong th¸ng (coi nh­ ch­a trÝnh khÊu hao)
      If khau_hao = KH_KHONG Then
            .CL_NS = .CL_NS + .KH_NS
            .CL_TBS = .CL_TBS + .KH_TBS
            .CL_CNK = .CL_CNK + .KH_CNK
            .CL_TD = .CL_TD + .KH_TD
      End If
Else
        .CL_NS = rs_giatridau!TCL_NS
        .CL_TBS = rs_giatridau!TCL_TBS
        .CL_CNK = rs_giatridau!TCL_CNK
        .CL_TD = rs_giatridau!TCL_TD
End If
      rs_giatridau.Close
      Set rs_giatridau = Nothing
End With
End Sub
'======================================================================================
' SUB GiamTaiSan : Thùc hiÖn nghiÖp vô gi¶m tµi s¶n bao gåm :
'                                               - Ghi chøng tõ gi¶m dùa trªn gi¸ trÞ cña tµi s¶n trong th¸ng gi¶m.
'                                               - CËp nhËt th¸ng gi¶m cho tµi s¶n.
'                                               - §Æt l­îng khÊu hao cña tµi s¶n cho c¸c th¸ng sau kÓ tõ th¸ng gi¶m b»ng 0.
'                   Tham sè : M· sè cña tµi s¶n, th¸ng cã t¸c ®éng gi¶m.
'                          Chó ý : Tµi s¶n kh«ng tÝnh khÊu hao trong th¸ng gi¶m.
'                                         Thñ tôc nµy ®­îc gäi duy nhÊt tõ mnHoatDong: "Gi¶m tµi s¶n"
'======================================================================================
Public Sub GiamTaiSan(ma_ts As Long, thg_giam As Integer)
Dim sql As String
      ' TÝnh gi¸ trÞ tµi s¶n trong th¸ng gi¶m (ch­a trÝch khÊu hao)
      TinhGiaTriTaiSan ma_ts, thg_giam + 1, KH_KHONG
      ' LÊy m· tµi kho¶n tµi s¶n
      sql = "SELECT MaTaiKhoan AS F1 FROM TaiSan WHERE MaSo = " _
                                                                                                                                                                        + CStr(ma_ts)
      ' Thµnh lËp ph¸t sinh
      ThanhLapPhatSinh NV_GIAM, CLng5(SelectSQL(sql))
      ' clsChungTu sÏ sö dông c¸c th«ng tin l­u trong biÕn chung GiaTri ®Ó ghi
      ' vµo l­îng t¨ng gi¶m trªn chøng tõ (khi gi¶m cÇn ph¶i cËp nhËt sè ©m).
      With GiaTri
            .NG_NS = .NG_NS * -1
            .NG_TBS = .NG_TBS * -1
            .NG_CNK = .NG_CNK * -1
            .NG_TD = .NG_TD * -1
            .CL_NS = .CL_NS * -1
            .CL_TBS = .CL_TBS * -1
            .CL_CNK = .CL_CNK * -1
            .CL_TD = .CL_TD * -1
      End With
End Sub
'======================================================================================
' SUB TacDongGiamTaiSan : ThÓ hiÖn c¸c thay ®æi trªn d÷ liÖu ®èi víi nghiÖp vô gi¶m tµi s¶n
'                                                                 - L­îng khÊu hao cña c¸c th¸ng kÓ sau tõ th¸ng gi¶m bÞ ®Æt b»ng 0
'                                                                 - Th¸ng gi¶m cña tµi s¶n ®­îc ghi kh¸c 13
'                                                            Kh«i phôc l¹i tr¹ng th¸i tr­íc ®ã (khi xo¸ chøng tõ gi¶m)
'                                                                 - L­îng khÊu hao cña c¸c th¸ng kÓ sau tõ th¸ng gi¶m ®­îc ®Æt l¹i
'                                                                    b»ng gi¸ trÞ cña th¸ng ngay tr­íc ®ã
'                                                                 - Th¸ng gi¶m cña tµi s¶n ®­îc ghi l¹i b»ng 13
'                                      Tham sè : M· tµi s¶n bÞ t¸c ®éng, th¸ng t¸c ®éng, kiÓu t¸c ®éng
'                                       Sö dông : Thñ tôc nµy ®­îc gäi tõ thñ tôc GiamTaiSan vµ frmChungTu
'======================================================================================
Public Sub TacDongGiamTaiSan(ma_ts As Long, thg As Integer, tac_dong As Integer)
Dim sql As String
      If tac_dong = TD_GIAM Then     ' Gi¶m tµi s¶n
            ExecuteSQL5 "UPDATE ThongSo SET KH_NS = 0, KH_TBS = 0, KH_CNK = 0, KH_TD = 0 " _
                                       & "WHERE MaTS = " + CStr(ma_ts) + " AND " + WThang2("Thang", thg, 0)
      Else                                ' Kh«i phôc l¹i d÷ liÖu
            Dim rs_khauhao As Recordset
            ' CËp nhËt l¹i l­îng khÊu hao víi d÷ liÖu cña th¸ng ngay tr­íc th¸ng gi¶m
            sql = "SELECT DISTINCTROW ThongSo.Thang, TaiSan.ThangGiam, ThongSo.KH_NS, ThongSo.KH_TBS, ThongSo.KH_CNK, ThongSo.KH_TD " _
                  & "FROM TaiSan INNER JOIN ThongSo ON (ThongSo.Thang = TaiSan.ThangGiam-1) AND (TaiSan.MaSo = ThongSo.MaTS) " _
                  & "WHERE ThongSo.MaTS = " + CStr(ma_ts)
            Set rs_khauhao = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
            Do While Not rs_khauhao.EOF
                  If rs_khauhao!thang = CThangDB(ThangTruoc(rs_khauhao!ThangGiam)) Then
                        ExecuteSQL5 "UPDATE DISTINCTROW TaiSan INNER JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS " _
                                                                              & "SET ThongSo.KH_NS = " + DoiDau(rs_khauhao!KH_NS) _
                                                                                                        + ", KH_TBS = " + DoiDau(rs_khauhao!KH_TBS) _
                                                                                                        + ", KH_CNK = " + DoiDau(rs_khauhao!KH_CNK) _
                                                                                                        + ", KH_TD = " + DoiDau(rs_khauhao!KH_TD) _
                                                + " WHERE MaTS = " + CStr(ma_ts) + " AND ThongSo.Thang >= " + CStr(rs_khauhao!thang)
                        Exit Do
                    End If
                    rs_khauhao.MoveNext
            Loop
            rs_khauhao.Close
            Set rs_khauhao = Nothing
      End If
      ' Ghi th¸ng gi¶m tµi s¶n
      ExecuteSQL5 "UPDATE TaiSan SET ThangGiam = " + CStr(thg) + " WHERE MaSo = " + CStr(ma_ts)
End Sub
'======================================================================================
' SUB XoaTaiSan : Xo¸ tµi s¶n
'                 Sö dông : frmTaiSan (Khi nhËp ®Çu kú hoÆc khi kh«ng ghi chøng tõ t¨ng hîp lÖ)
'======================================================================================
Public Sub XoaTaiSan(ma_ts As Long)
Dim rs_chungtu As Recordset
Dim sql As String
Dim ctu As New ClsChungtu
     ' Xo¸ chøng tõ
      sql = "SELECT ChungTu.MaSo FROM CTTaiSan INNER JOIN ChungTu ON CTTaiSan.MaCTKT = ChungTu.MaCT WHERE CTTaiSan.MaTS = " + CStr(ma_ts) + " AND CTTaiSan.MaCTKT > 0"
      Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      Do While Not rs_chungtu.EOF
            ctu.InitChungtu rs_chungtu!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
            ctu.XoaChungtu
'            ExecuteSQL5 "DELETE * FROM PhatSinh WHERE MaCT = " + CStr(rs_chungtu!MaSo)
            rs_chungtu.MoveNext
      Loop
      rs_chungtu.Close
      Set rs_chungtu = Nothing
      ExecuteSQL5 "DELETE * FROM CTTaiSan WHERE MaTS = " + CStr(ma_ts)
      ' Xo¸ c¸c dông cô phô tïng kÌm theo
      ExecuteSQL5 "DELETE * FROM DCPTung WHERE MaTS = " + CStr(ma_ts)
      ' Xo¸ c¸c th«ng sè
      ExecuteSQL5 "DELETE * FROM ThongSo WHERE MaTS = " + CStr(ma_ts)
     ' Xo¸ tµi s¶n
      ExecuteSQL5 "DELETE * FROM TaiSan WHERE MaSo = " + CStr(ma_ts)
End Sub
'======================================================================================
' SUB DieuChinhKhauHao : Tù ®éng ®iÒu chØnh l¹i l­îng khÊu hao nÕu lín h¬n gi¸ trÞ cßn l¹i cña tµi s¶n
'                                                      thùc hiÖn cho tÊt c¶ c¸c tµi s¶n trong mét th¸ng.
'                                 Tham sè : Th¸ng kiÓm tra
'                                  Sö dông : §­îc gäi tõ thñ tôc CapNhatGiaTriTaiSan trong frmBaoCao
'======================================================================================
Public Sub DieuChinhKhauHao(thg As Integer)
'      pExecuteSQL = "UPDATE ThongSo SET KH_NS = 0, KH_TBS = 0 WHERE CL_NS < 0 AND Thang = " + CStr(thg)
'      ExecuteSQL5 False
      
      ' Ng©n s¸ch
      ExecuteSQL5 "UPDATE ThongSo SET KH_NS = IIF((KH_NS + CL_NS) < 0, 0 , KH_NS + CL_NS) " _
                                                                                                       & "WHERE CL_NS < 0 AND Thang = " + CStr(thg)
      ' Tù bæ sung
      ExecuteSQL5 "UPDATE ThongSo SET KH_TBS = IIF((KH_TBS + CL_TBS) < 0, 0 , KH_TBS + CL_TBS) " _
                                                                                                       & "WHERE CL_TBS < 0 AND Thang = " + CStr(thg)
      ' C¸c nguån kh¸c
      ExecuteSQL5 "UPDATE ThongSo SET KH_CNK = IIF((KH_CNK + CL_CNK) < 0, 0 , KH_CNK + CL_CNK) " _
                                                                                                       & "WHERE CL_CNK < 0 AND Thang = " + CStr(thg)
      ' TÝn dông
      ExecuteSQL5 "UPDATE ThongSo SET KH_TD = IIF((KH_TD + CL_TD) < 0, 0 , KH_TD + CL_TD) " _
                                                                                                       & "WHERE CL_TD < 0 AND Thang = " + CStr(thg)
End Sub
'======================================================================================
' SUB XoaGiaTri
'======================================================================================
Public Sub XoaGiaTri()
      With GiaTri
            .NG_NS = 0
            .NG_TBS = 0
            .NG_CNK = 0
            .NG_TD = 0
            .CL_NS = 0
            .CL_TBS = 0
            .CL_CNK = 0
            .CL_TD = 0
            .KH_NS = 0
            .KH_TBS = 0
            .KH_CNK = 0
            .KH_TD = 0
      End With
End Sub

Public Function ThangDaKhauHao(tdau As Integer, tcuoi As Integer, loaikh As Long, shtk As String) As Boolean
    
    ThangDaKhauHao = SelectSQL("SELECT DISTINCTROW TOP 1 ChungTu.MaCT AS F1 FROM " + ChungTu2TKNC(-1) _
        & " WHERE HethongTK.SoHieu LIKE '" + shtk + "*' AND " + WThang("ThangCT", tdau, tcuoi) + " AND MaLoai = 12" + IIf(loaikh >= 0, " AND CT_ID = " + CStr(loaikh), "")) > 0
    
End Function

Public Sub XoaChungTuKhauHao(tdau As Integer, tcuoi As Integer, loaikh As Long, ctmoi As Long, shtk As String)
    Dim rs As Recordset, ctu As New ClsChungtu
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW ChungTu.MaSo, NgayCT, NgayGS FROM " + ChungTu2TKNC(-1) _
        & " WHERE HethongTK.SoHieu LIKE '" + shtk + "*' AND MaCT <> " + CStr(ctmoi) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND MaLoai = 12 AND CT_ID = " + CStr(loaikh), dbOpenSnapshot)
    Do While Not rs.EOF
        ctu.InitChungtu rs!MaSo, 0, "", 0, rs!NgayCT, rs!NgayGS, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
        ctu.XoaChungtu
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Sub XoaChungtuTS(loaict As Integer, MaSoCT As Long)
Dim sql As String

        Select Case loaict
            Case 9:
              Dim rs As Recordset
                sql = "SELECT MaTS FROM CTTaiSan  WHERE MaCTKT=" + CStr(MaSoCT)
                    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
                Do While Not rs.EOF
                 XoaTaiSan rs!MaTS
                 rs.MoveNext
                Loop
                 rs.Close
'                sql = "SELECT MaTS AS F1 FROM CTTaiSan  WHERE MaCTKT=" + CStr(MaSoCT)
'                XoaTaiSan SelectSQL(sql)
            Case 10:
                sql = "SELECT MaTS AS F1 FROM CTTaiSan  WHERE MaCTKT=" + CStr(MaSoCT)
                TacDongGiamTaiSan SelectSQL(sql), 13, TD_KHOIPHUC
        End Select
        
        If loaict <> 9 Then ExecuteSQL5 "DELETE FROM CTTaiSan WHERE MaCTKT = " + CStr(MaSoCT)
End Sub
'===================================================================================
' Hµm tr¶ vÒ m· sè, tªn TK tõ sè hiÖu
'===================================================================================
Public Function TenTS(sh As String, mtk As Long) As String
    Dim rs_tk As Recordset
    If mtk > 0 Then
        Set rs_tk = DBKetoan.OpenRecordset("SELECT SoHieu, Ten FROM TaiSan WHERE MaSo=" + CStr(mtk), dbOpenSnapshot)
        TenTS = rs_tk!Ten
        sh = rs_tk!sohieu
    Else
        Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo, Ten FROM TaiSan WHERE SoHieu='" + sh + "'", dbOpenSnapshot)
        If rs_tk.RecordCount > 0 Then
            mtk = rs_tk!MaSo
            TenTS = rs_tk!Ten
        Else
            mtk = 0
            TenTS = ""
        End If
    End If
    rs_tk.Close
    Set rs_tk = Nothing
End Function

Public Sub XoaPSTS(thang As Integer)
Dim rs_chungtu As Recordset

    Set rs_chungtu = DBKetoan.OpenRecordset("SELECT CTTaiSan.* FROM CTTaiSan" _
            & " WHERE Thang = " + CStr(thang), dbOpenSnapshot)
    Do While Not rs_chungtu.EOF
        Select Case rs_chungtu!maloai
            Case NV_TANG:
                XoaTaiSan rs_chungtu!MaTS
            Case NV_GIAM:
                TacDongGiamTaiSan rs_chungtu!MaTS, 13, TD_KHOIPHUC
        End Select
        rs_chungtu.MoveNext
    Loop
    rs_chungtu.Close
    Set rs_chungtu = Nothing
    
    ExecuteSQL5 "DELETE FROM CTTaiSan WHERE Thang = " + CStr(thang)
End Sub
'======================================================================================
' SUB ChuyenNamMoiTS
'======================================================================================
Public Sub ChuyenNamMoiTS()
    Dim i As Integer
      ' Xo¸ c¸c tµi s¶n ®· bÞ gi¶m trong n¨m
      ExecuteSQL5 "DELETE DCPTung.* FROM DCPTung RIGHT JOIN TaiSan " _
                             & "ON DCPTung.MaTS = TaiSan.MaSo WHERE TaiSan.ThangGiam < 13"
      ExecuteSQL5 "DELETE ThongSo.* FROM ThongSo RIGHT JOIN TaiSan " _
                             & "ON ThongSo.MaTS = TaiSan.MaSo WHERE TaiSan.ThangGiam < 13"
      ExecuteSQL5 "DELETE * FROM TaiSan WHERE ThangGiam < 13"
      ExecuteSQL5 "UPDATE TaiSan SET NamKH = 0 WHERE IsNull(NamKH)"
            
      ' TÝnh gi¸ trÞ cho c¸c tµi s¶n vµo cuèi kú
      TinhGiaTriCuoiKy
      ' T¹o chøng tõ kÕt chuyÓn
      TaoChungTuKetChuyen
      ' CËp nhËt l¹i l­îng khÊu hao hµng th¸ng vµ c¸c ®èi t­îng quan hÖ
      ExecuteSQL5 "UPDATE DISTINCTROW ThongSo LEFT JOIN ThongSoCuoiKy ON ThongSo.MaTS = ThongSoCuoiKy.MaTS SET ThongSo.KH_NS = ThongSoCuoiKy.KH_NS, ThongSo.KH_TBS = ThongSoCuoiKy.KH_TBS, ThongSo.KH_CNK = ThongSoCuoiKy.KH_CNK, ThongSo.KH_TD = ThongSoCuoiKy.KH_TD, ThongSo.MaDTQL = ThongSoCuoiKy.MaDTQL, ThongSo.MaDTSD = ThongSoCuoiKy.MaDTSD, ThongSo.MaTTSD = ThongSoCuoiKy.MaTTSD WHERE (((ThongSo.Thang)<12 And (ThongSo.Thang)>0));"
      ' CËp nhËt l¹i thêi gian
      ExecuteSQL5 "UPDATE TaiSan SET ThangTang = 0"
End Sub
'======================================================================================
' SUB TinhGiaTriCuoiKy
'======================================================================================
Private Sub TinhGiaTriCuoiKy()
Dim rs_giatri As Recordset
      SetSQL "TongKhauHao", "SELECT Sum(ThongSo.KH_NS) AS TKH_NS, Sum(ThongSo.KH_TBS) AS TKH_TBS, Sum(ThongSo.KH_CNK) AS TKH_CNK, Sum(ThongSo.KH_TD) AS TKH_TD, ThongSo.MaTS, First(TaiSan.ThangTang) As ThangT" _
            & ", Max(IIF(ThongSo.Thang = 12, MaDTQL,0)) As DTQL, Max(IIF(ThongSo.Thang = 12, MaDTSD,0)) As DTSD, Max(IIF(ThongSo.Thang = 12, MaTTSD,0)) As TTSD " _
            & "From TaiSan RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS WHERE " + VC("ThongSo.Thang", "IIF(TaiSan.ThangTang=0," + CStr(pThangDauKy) + ",TaiSan.ThangTang)") + " GROUP BY MaTS"
      SetSQL "TongGiaTri", "SELECT Sum(NG_NS) AS TNG_NS, Sum(NG_TBS) AS TNG_TBS, Sum(NG_CNK) AS TNG_CNK, Sum(NG_TD) AS TNG_TD, " _
            & "Sum(CL_NS) AS TCL_NS, Sum(CL_TBS) AS TCL_TBS, Sum(CL_CNK) AS TCL_CNK, Sum(CL_TD) AS TCL_TD, MaTS " _
            & "FROM CTTaiSan WHERE Thang < 13 GROUP BY MaTS"
      SetSQL "GiaTriTaiSan", "SELECT DISTINCTROW TongKhauHao.ThangT, TongKhauHao.DTQL, TongKhauHao.DTSD, TongKhauHao.TTSD, TNG_NS AS NG_NS, TNG_TBS AS NG_TBS, TNG_CNK AS NG_CNK, TNG_TD AS NG_TD, " _
            & "TCL_NS-TKH_NS AS CL_NS, TCL_TBS-TKH_TBS AS CL_TBS, TCL_CNK-TKH_CNK AS CL_CNK, TCL_TD-TKH_TD AS CL_TD, " _
            & "TongGiaTri.MaTS FROM TongKhauHao INNER JOIN TongGiaTri ON TongKhauHao.MaTS = TongGiaTri.MaTS"
      Set rs_giatri = DBKetoan.OpenRecordset("GiaTriTaiSan", dbOpenSnapshot, dbForwardOnly)
      Do While Not rs_giatri.EOF
            ExecuteSQL5 "UPDATE ThongSo SET NG_NS = " + DoiDau(rs_giatri!NG_NS) + ", NG_TBS = " + DoiDau(rs_giatri!NG_TBS) + ", NG_CNK = " + DoiDau(rs_giatri!NG_CNK) + ", NG_TD = " + DoiDau(rs_giatri!NG_TD) _
                  & ", CL_NS = " + DoiDau(rs_giatri!CL_NS) + ", CL_TBS = " + DoiDau(rs_giatri!CL_TBS) + ", CL_CNK = " + DoiDau(rs_giatri!CL_CNK) + ", CL_TD = " + DoiDau(rs_giatri!CL_TD) _
                  + " WHERE MaTS = " + CStr(rs_giatri!MaTS) + " And Thang = 0"
                  
            If rs_giatri!ThangT > 0 Then
                    Dim i As Integer
                    For i = 0 To rs_giatri!ThangT - 1
                            ExecuteSQL5 "INSERT INTO ThongSo (MaSo,MaTS, Thang, NG_NS, NG_TBS, NG_CNK, NG_TD" _
                                & ", CL_NS, CL_TBS, CL_CNK, CL_TD, MaDTQL, MaDTSD, MaTTSD) VALUES (" + CStr(Lng_MaxValue("MaSo", "ThongSo") + 1) + "," + CStr(rs_giatri!MaTS) + "," + CStr(i) + "," + DoiDau(rs_giatri!NG_NS) _
                                + "," + DoiDau(rs_giatri!NG_TBS) + "," + DoiDau(rs_giatri!NG_CNK) + "," + DoiDau(rs_giatri!NG_TD) _
                                + "," + DoiDau(rs_giatri!CL_NS) + "," + DoiDau(rs_giatri!CL_TBS) + "," + DoiDau(rs_giatri!CL_CNK) _
                                + "," + DoiDau(rs_giatri!CL_TD) + "," + CStr(rs_giatri!DTQL) + "," + CStr(rs_giatri!DTSD) + "," + CStr(rs_giatri!TTSD) + ")"
                    Next
            End If
            rs_giatri.MoveNext
      Loop
      rs_giatri.Close
      Set rs_giatri = Nothing
    ' Xo¸ c¸c chøng tõ cña n¨m cò
      ExecuteSQL5 "DELETE * FROM CTTaiSan"
End Sub
'======================================================================================
' SUB TaoChungTuKetChuyen
'======================================================================================
Private Sub TaoChungTuKetChuyen()
Dim rs_thongso As Recordset
      Set rs_thongso = DBKetoan.OpenRecordset("SELECT DISTINCTROW TaiSan.MaSo, TaiSan.SoHieu, TaiSan.Ten, " _
            & "ThongSo.NG_NS, ThongSo.NG_TBS, ThongSo.NG_CNK, ThongSo.NG_TD, ThongSo.CL_NS, ThongSo.CL_TBS, ThongSo.CL_CNK, ThongSo.CL_TD " _
            & "FROM TaiSan RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS " _
            & "WHERE ThongSo.Thang=0", dbOpenSnapshot)
      Do Until rs_thongso.EOF
            With rs_thongso
            ExecuteSQL5 "INSERT INTO CTTaiSan (MaSo, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, " _
                  & "MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD) " _
                  & "VALUES (" + CStr(Lng_MaxValue("MaSo", "CTTaiSan") + 1) + ",'" + !sohieu + CStr(pNamTC - 1899) + "', 0" _
                  + ",#" + Format(Date, Mask_DB) + "#,#" + Format(Date, Mask_DB) + "#,'" _
                  + "§Çu n¨m: " + !Ten + "'," + CStr(DK_LOAI) + "," + CStr(DK_NHOM) + "," + CStr(!MaSo) + "," _
                  + DoiDau(!NG_NS) + "," + DoiDau(!NG_TBS) + "," + DoiDau(!NG_CNK) + "," + DoiDau(!NG_TD) + "," _
                  + DoiDau(!CL_NS) + "," + DoiDau(!CL_TBS) + "," + DoiDau(!CL_CNK) + "," + DoiDau(!CL_TD) + ")"
            End With
            rs_thongso.MoveNext
      Loop
      rs_thongso.Close
      Set rs_thongso = Nothing
End Sub
'======================================================================================
' Thñ tôc tÝnh sè d­ ®Çu kú cña c¸c tµi kho¶n vËt t­
'======================================================================================
Public Sub SoDuTKTS()
    Dim rs_tk As Recordset, taikhoan As New ClsTaikhoan
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW SoHieu FROM HethongTK WHERE (TKCon = 0) AND (TK_ID = " + CStr(TSCD_ID) + " OR TK_ID = " + CStr(KHTSCD_ID) + ")", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_tk.EOF
        taikhoan.InitTaikhoanSohieu rs_tk!sohieu
        taikhoan.NoDauKy = 0
        taikhoan.CoDauKy = 0
        taikhoan.CapNhatTk
        rs_tk.MoveNext
    Loop
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT HeThongTK.SoHieu, Sum(CTTaiSan.NG_NS + CTTaiSan.NG_TBS + CTTaiSan.NG_CNK + CTTaiSan.NG_TD) As TNG" _
        & " FROM (LoaiTaiSan INNER JOIN (TaiSan INNER JOIN CTTaiSan ON TaiSan.MaSo = CTTaiSan.MaTS) ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) INNER JOIN HeThongTK ON LEFT(LoaiTaiSan.SoHieu,LEN(HeThongTK.SoHieu)) = HeThongTK.SoHieu" _
        & " Where CTTaiSan.maloai = 30 AND LoaiTaiSan.Cap=1 AND TKCon=0 GROUP BY HeThongTK.SoHieu", dbOpenSnapshot, dbForwardOnly)
    
    Do While Not rs_tk.EOF
        taikhoan.InitTaikhoanSohieu rs_tk!sohieu
        taikhoan.NoDauKy = rs_tk!TNG
        taikhoan.CapNhatTk
        rs_tk.MoveNext
    Loop
    
    Set rs_tk = DBKetoan.OpenRecordset("SELECT LEFT(HeThongTK.SoHieu,3) As SHTK, Sum((CTTaiSan.NG_NS + CTTaiSan.NG_TBS + CTTaiSan.NG_CNK + CTTaiSan.NG_TD) - (CTTaiSan.CL_NS + CTTaiSan.CL_TBS + CTTaiSan.CL_CNK + CTTaiSan.CL_TD)) AS THM" _
        & " FROM (LoaiTaiSan INNER JOIN (TaiSan INNER JOIN CTTaiSan ON TaiSan.MaSo = CTTaiSan.MaTS) ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) INNER JOIN HeThongTK ON (LEFT(LoaiTaiSan.SoHieu,LEN(HeThongTK.SoHieu)) = HeThongTK.SoHieu AND LoaiTaiSan.Cap=1)" _
        & " Where CTTaiSan.maloai = 30 AND LoaiTaiSan.Cap=1 AND TKCon=0 GROUP BY LEFT(HeThongTK.SoHieu,3)", dbOpenSnapshot, dbForwardOnly)
    
    Do While Not rs_tk.EOF
        taikhoan.InitTaikhoanSohieu "214" + Right(rs_tk!shtk, 1)
        taikhoan.CoDauKy = rs_tk!THM
        taikhoan.CapNhatTk
        rs_tk.MoveNext
    Loop
    
    rs_tk.Close
    Set rs_tk = Nothing
End Sub

Public Function GTHaoMon(tkng As String, thang As Integer) As Double
    Dim sql As String
    
    sql = "SELECT Sum(NG_NS+NG_TBS+NG_TD+NG_CNK-CL_NS-CL_TBS-CL_TD-CL_CNK) AS F1 FROM (ThongSo INNER JOIN TaiSan ON ThongSo.MaTS=TaiSan.MaSo) INNER JOIN LoaiTaiSan ON TaiSan.MaLoai=LoaiTaiSan.MaSo " _
        & " WHERE Thang=" + CStr(thang) + " AND LoaiTaiSan.Sohieu LIKE '" + tkng + "*'"
    GTHaoMon = SelectSQL(sql)
End Function

Public Sub DieuChinhKH(mts As Long, thang As Integer)
    Dim ts As New clsTaiSan, i As Integer

    ts.ChiDinh mts, thang
    If ts.NamKH > 0 Then
        With ts.ThongSo
            .KH_NS = RoundMoney(.NG_NS / (12 * ts.NamKH))
            .KH_TBS = RoundMoney(.NG_TBS / (12 * ts.NamKH))
            .KH_CNK = RoundMoney(.NG_CNK / (12 * ts.NamKH))
            .KH_TD = RoundMoney(.NG_TD / (12 * ts.NamKH))
            .SuaDoiQuanHe False
        End With
    End If
End Sub

Public Function KhongDC(ms As Long) As Boolean
    Dim sql As String
    
    sql = "SELECT Count(MaSo) AS F1 FROM CTTaiSan WHERE MaTS=" + CStr(ms) + " AND MaNhom=" + CStr(NV_DGLAI)
    KhongDC = (SelectSQL(sql) > 0)
End Function

Public Function SoTangGiamTS(shtk As String, tdau As Integer, tcuoi As Integer, mnhom As Long) As Double
    SoTangGiamTS = SelectSQL("SELECT SUM(NG_NS+NG_TBS+NG_TD+NG_CNK) AS F1 FROM (CTTaiSan INNER JOIN TaiSan ON CTTaiSan.MaTS=TaiSan.MaSo) INNER JOIN LoaiTaiSan ON TaiSan.MaTaiKhoan=LoaiTaiSan.MaSo WHERE " + WThang("Thang", tdau, tcuoi) + " AND CTTaiSan.MaNhom=" + CStr(mnhom) + " AND LoaiTaiSan.SoHieu LIKE '" + shtk + "*'")
End Function

Public Function SoKHTS(shtk As String, tdau As Integer, tcuoi As Integer)
    SoKHTS = SelectSQL("SELECT SUM(KH_NS+KH_TBS+KH_TD+KH_CNK) AS F1 FROM (ThongSo INNER JOIN TaiSan ON ThongSo.MaTS=TaiSan.MaSo) INNER JOIN LoaiTaiSan ON TaiSan.MaTaiKhoan=LoaiTaiSan.MaSo WHERE " + WThang("Thang", tdau, tcuoi) + " AND LoaiTaiSan.SoHieu LIKE '" + shtk + "*'")
    SoKHTS = SoKHTS + SelectSQL("SELECT SUM(NG_NS+NG_TBS+NG_TD+NG_CNK-CL_NS-CL_TBS-CL_TD-CL_CNK) AS F1 FROM (CTTaiSan INNER JOIN TaiSan ON CTTaiSan.MaTS=TaiSan.MaSo) INNER JOIN LoaiTaiSan ON TaiSan.MaTaiKhoan=LoaiTaiSan.MaSo WHERE " + WThang("Thang", tdau, tcuoi) + " AND CTTaiSan.MaLoai=32 AND LoaiTaiSan.SoHieu LIKE '" + shtk + "*'")
End Function

Public Function NGHetKH(shtk As String, tcuoi As Integer) As Double
    NGHetKH = SelectSQL("SELECT SUM(NG_NS+NG_TBS+NG_TD+NG_CNK) AS F1 FROM (ThongSo INNER JOIN TaiSan ON ThongSo.MaTS=TaiSan.MaSo) INNER JOIN LoaiTaiSan ON TaiSan.MaTaiKhoan=LoaiTaiSan.MaSo WHERE Thang=" + CStr(CThangDB(tcuoi)) + " AND LoaiTaiSan.SoHieu LIKE '" + shtk + "*' AND (CL_NS+CL_TBS+CL_TD+CL_CNK)=0")
End Function

