VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTaiSan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
' Thuéc tÝnh
'======================================================================================
Public MaSo As Long                                                     ' M· sè
Public Ten As String                                                        ' Tªn tµi s¶n
Public sohieu As String                                                 ' Sè hiÖu
Public NangLuc As String                                              ' C«ng suÊt, diÖn tÝch thiÕt kÕ
Public GhiChu As String                                                  ' Ghi chó
Public MaNuoc As Long                                                 ' M· n­íc s¶n xuÊt
Public NamSX As Integer                                               ' N¨m s¶n xuÊt
Public NamSD As Integer                                              ' N¨m sö dông
Public MaTaiKhoan As Long                                        ' M· tµi kho¶n tµi s¶n
Public maloai As Long                                                   ' M· ph©n lo¹i tµi s¶n
Public MaNhom As Long                                                ' M· ph©n nhãm tµi s¶n
Public ThangTang As Integer                                       ' Th¸ng t¨ng (0 nÕu t¨ng tõ c¸c n¨m tr­íc)
Public ThangGiam As Integer                                       ' Th¸ng gi¶m (13 nÕu ch­a gi¶m trong n¨m)
Public ThongSo As New clsThongSo                       ' Th«ng sè thay ®æi theo th¸ng
Public NamKH As Integer
Public shct As String
Public NCT As Date
'======================================================================================
' FUNCTION HopLe : KiÓm tra tÝnh hîp lÖ cña d÷ liÖu
'                         Tr¶ vÒ : 0 nÕu ®óng, -1 nÕu cã lçi
'======================================================================================
Public Function HopLe()
Dim thong_bao As String, sql As String
Dim so_hieu As String
      If Len(sohieu) = 0 Or sohieu = "(...)" Then thong_bao = "ThiÕu sè hiÖu tµi s¶n": GoTo Err_InValidate
      If Len(Ten) = 0 Or Ten = "(...)" Then thong_bao = "ThiÕu tªn tµi s¶n": GoTo Err_InValidate
      ' C¸c ®èi t­îng quan hÖ
      If MaNuoc = 0 Then thong_bao = "ThiÕu hoÆc ch­a ®¨ng ký tªn n­íc s¶n xuÊt": GoTo Err_InValidate
      'If NamSX < 1800 Or NamSX > 3000 Then thong_bao = "Thêi gian s¶n xuÊt cña tµi s¶n thiÕu hoÆc kh«ng chÝnh x¸c": GoTo Err_InValidate
      'If NamSD < 1800 Or NamSD > 3000 Then thong_bao = "Thêi gian sö dông cña tµi s¶n thiÕu hoÆc kh«ng chÝnh x¸c": GoTo Err_InValidate
      If MaTaiKhoan = 0 Then thong_bao = "ThiÕu hoÆc ch­a ®¨ng ký tµi kho¶n tµi s¶n": GoTo Err_InValidate
      If maloai = 0 Then thong_bao = "ThiÕu hoÆc ch­a ®¨ng ký ph©n lo¹i tµi s¶n": GoTo Err_InValidate
'      If MaNhom = 0 Then thong_bao = "ThiÕu hoÆc ch­a ®¨ng ký ph©n nhãm tµi s¶n": GoTo Err_InValidate
      ' KiÓm tra sè hiÖu
      sql = "SELECT SoHieu AS F1 FROM LoaiTaiSan WHERE MaSo = " + CStr(IIf(MaNhom > 0, MaNhom, maloai))
      so_hieu = CStr(SelectSQL(sql))
      If Not Left(sohieu, Len(so_hieu)) = so_hieu Then thong_bao = "Sè hiÖu tµi s¶n kh«ng ®óng quy ®Þnh": GoTo Err_InValidate
      If ThangGiam = 0 Then ThangGiam = 13
      HopLe = ThongSo.HopLe
      Exit Function
Err_InValidate:
      Beep
      MsgBox thong_bao, vbCritical, App.ProductName
      HopLe = -1
End Function
'======================================================================================
' FUNCTION ThemMoi : Ghi ®èi t­îng tµi s¶n hiÖn t¹i vµo c¬ së d÷ liÖu.
'                                                     - Ghi tµi s¶n vµo b¶ng TaiSan
'                                                     - T¹o c¸c b¶n ghi chøa gi¸ trÞ hµng th¸ng (tõ th¸ng 0 cho ®Õn hÕt n¨m)
'                               Tr¶ vÒ : 0 nÕu thµnh c«ng, -1 nÕu cã lçi
'                                 Chó ý : M· sè cña ®èi t­îng tµi s¶n hiÖn t¹i sÏ ®­îc ®Æt kh¸c 0 sau thñ tôc nµy.
'======================================================================================
Public Function ThemMoi(kh1 As Integer) As Integer
Dim chi_so As Integer, sql As String, ts1 As clsThongSo
        
        If Len(shct) = 0 Then shct = "..."
      ' Ghi tµi s¶n
      sql = "INSERT INTO TaiSan (MaSo, Ten, SoHieu, NangLuc, GhiChu, MaNuoc, NamSX, " _
            & "NamSD, MaTaiKhoan, MaLoai, MaNhom, ThangTang, ThangGiam,NamKH, SHCT, NCT) VALUES (" + CStr(Lng_MaxValue("MaSo", "TaiSan") + 1) + ",'" _
            + Ten + "','" + sohieu + "','" + NangLuc + "','" + GhiChu + "'," + CStr(MaNuoc) + "," _
            + CStr(NamSX) + "," + CStr(NamSD) + "," + CStr(MaTaiKhoan) + "," + CStr(maloai) _
            + "," + CStr(MaNhom) + "," + CStr(ThangTang) + "," + CStr(ThangGiam) + "," + CStr(NamKH) + ",'" + shct + "',#" + Format(NCT, Mask_DB) + "#)"
        
      ' Ghi c¸c th«ng sè cho c¸c th¸ng trong n¨m
      If ExecuteSQL5(sql) = 0 Then
            MaSo = Lng_MaxValue("MaSo", "TaiSan")
            ThongSo.MaTS = MaSo
            For chi_so = 0 To 12
                If chi_so >= CThangDB(ThangTang) And chi_so <= CThangDB(ThangGiam) Then
                    ThongSo.thang = chi_so
                    If kh1 = 1 And chi_so = CThangDB(ThangTang) Then
                        Set ts1 = New clsThongSo
                        ThongSo.DanXuat ts1
                        ThongSo.KH_NS = 0
                        ThongSo.KH_TBS = 0
                        ThongSo.KH_TD = 0
                        ThongSo.KH_CNK = 0
                        ThemMoi = ThongSo.ThemMoi
                        ts1.DanXuat ThongSo
                        Set ts1 = Nothing
                    Else
                        ThemMoi = ThongSo.ThemMoi
                    End If
                    If ThemMoi = -1 Then Exit Function
                End If
            Next
      Else
            ThemMoi = -1
      End If
End Function
'======================================================================================
' SUB ChiDinh : G¸n c¸c thuéc tÝnh cña ®èi t­îng tµi s¶n hiÖn t¹i theo gi¸ trÞ l­u trong c¬ së d÷ liÖu
'         Tham sè : M· tµi s¶n cÇn lÊy th«ng tin, thêi gian (®Ó x¸c ®Þnh c¸c th«ng sè)
'          Sö dông : Thñ tôc ChiDinh trong clsThongSo
'                Chó ý : Kh«ng kiÓm tra tham sè khi thùc hiÖn
'======================================================================================
Public Sub ChiDinh(ma_so As Long, thang_cd As Integer)
Dim rs_taisan As Recordset, sql As String
      sql = "SELECT * FROM TaiSan WHERE MaSo = " + CStr(ma_so)
      Set rs_taisan = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
            ' Chi tiÕt
      If rs_taisan.RecordCount > 0 Then
            MaSo = rs_taisan!MaSo
            Ten = rs_taisan!Ten
            sohieu = rs_taisan!sohieu
            NangLuc = rs_taisan!NangLuc
            GhiChu = rs_taisan!GhiChu
            NamSX = rs_taisan!NamSX
            NamSD = rs_taisan!NamSD
            ' C¸c ®èi t­îng quan hÖ
            MaNuoc = rs_taisan!MaNuoc
            MaTaiKhoan = rs_taisan!MaTaiKhoan
            maloai = rs_taisan!maloai
            MaNhom = rs_taisan!MaNhom
            ThangTang = rs_taisan!ThangTang
            ThangGiam = rs_taisan!ThangGiam
            NamKH = rs_taisan!NamKH
            shct = rs_taisan!shct
            NCT = rs_taisan!NCT
            ' Th«ng sè cña th¸ng ®­îc chØ ®Þnh
            If thang_cd > 0 Then ThongSo.ChiDinh ma_so, thang_cd
        Else
            KhoiTao
        End If
      rs_taisan.Close
      Set rs_taisan = Nothing
End Sub

Public Sub ChiDinhSH(sh As String)
    ChiDinh SelectSQL("SELECT MaSo AS F1 FROM TaiSan WHERE SoHieu='" + sh + "'"), 0
End Sub
'======================================================================================
' SUB KhoiTao
'======================================================================================
Public Sub KhoiTao()
      MaSo = 0
      Ten = "(...)"
      sohieu = "(...)"
      NangLuc = "(...)"
      GhiChu = "(...)"
      MaNuoc = 0
      NamSX = 0
      NamSD = 0
      MaTaiKhoan = 0
      maloai = 0
      MaNhom = 0
      ThangTang = 0
      NamKH = 0
      ThangGiam = 13
      shct = "..."
      ThongSo.KhoiTao
End Sub
'======================================================================================
' FUNCTION SuaDoi : Söa ®æi néi dung cña ®èi t­îng tµi s¶n hiÖn thêi
'                                               - C¸c th«ng sè cè ®Þnh
'                                               - C¸c thay ®æi luü tiÕn vÒ l­îng khÊu hao vµ c¸c ®èi t­îng quan hÖ
'                                               - C¸c thay ®æi vÒ gi¸ trÞ tµi s¶n (®­îc cËp nhËt ®ång thêi víi chøng tõ t¨ng
'                                                  hoÆc chøng tõ ®Çu kú t­¬ng øng)
'                              Chó ý : §èi víi tµi s¶n ®· bÞ gi¶m trong n¨m th× c¸c thay ®æi vÒ l­îng khÊu hao vµ
'                                            gi¸ trÞ lµ kh«ng ®­îc phÐp v× sÏ dÉn ®Õn m©u thuÉn víi chøng tõ gi¶m ®· ghi.
'======================================================================================
Public Function SuaDoi() As Integer
    Dim sql As String
      If Len(shct) = 0 Then shct = "..."
      ' Thay ®æi c¸c th«ng sè cè ®Þnh
      sql = "UPDATE TAISAN SET Ten = '" + Ten + "', SoHieu = '" + sohieu + "'" _
            & ", NangLuc = '" + NangLuc + "', GhiChu = '" + GhiChu + "', MaNuoc = " + CStr(MaNuoc) _
            + ", NamSX = " + CStr(NamSX) + ", NamSD = " + CStr(NamSD) + ",NamKH=" + CStr(NamKH) _
            + ", MaTaiKhoan = " + CStr(MaTaiKhoan) + ", MaLoai = " + CStr(maloai) + ", MaNhom = " + CStr(MaNhom) _
            + ", SHCT = '" + shct + "', NCT=#" + Format(NCT, Mask_DB) + "#" _
            + " WHERE MaSo = " + CStr(MaSo)
      If ExecuteSQL5(sql) <> 0 Then
            SuaDoi = -1
            Exit Function
      End If
      WSpace.BeginTrans
      ' NÕu tµi s¶n ®· bÞ ghi chøng tõ gi¶m th× ph¶i xo¸ chøng tõ nµy ®i míi cã thÓ söa ®æi gi¸ trÞ
      If Not ThangGiam = 13 Then
            SuaDoi = -2
            GoTo KhongDC
      End If
      ' Thay ®æi nguyªn gi¸, l­îng hao mßn, khÊu hao vµ c¸c ®èi t­îng quan hÖ
      If ThangTang > 0 And ThongSo.thang = CThangDB(ThangTang) Then
                Dim tong_ng As Double, tong_hm As Double
                Dim tsts As New clsThongSo
                
                tsts.ChiDinh MaSo, ThangTang
                If (tsts.NG_NS + tsts.NG_TBS + tsts.NG_CNK + tsts.NG_TD <> ThongSo.NG_NS + ThongSo.NG_TBS + ThongSo.NG_CNK + ThongSo.NG_TD) Or _
                    (tsts.CL_NS + tsts.CL_TBS + tsts.CL_CNK + tsts.CL_TD <> ThongSo.CL_NS + ThongSo.CL_TBS + ThongSo.CL_CNK + ThongSo.CL_TD) Then
                    tong_ng = (ThongSo.NG_NS + ThongSo.NG_TBS + ThongSo.NG_CNK + ThongSo.NG_TD)
                    tong_hm = tong_ng - (ThongSo.CL_NS + ThongSo.CL_TBS + ThongSo.CL_CNK + ThongSo.CL_TD)
                                            
                    If SetChungtuTang(tong_ng, tong_hm) <> 0 Then GoTo KhongDC
                End If
                Set tsts = Nothing
      End If
      If ThongSo.thang = 1 Or ThongSo.thang = ThangTang Then
            SuaDoi = ThongSo.SuaDoiQuanHe(True)
      Else
            SuaDoi = ThongSo.SuaDoiQuanHe(False)
      End If
      ' Thay ®æi nguyªn gi¸, gi¸ trÞ cßn l¹i l­u trong chøng tõ t¨ng tµi s¶n ®Çu kú vµ trong c¸c th¸ng
      ' tr­íc th¸ng ®Çu kú (Chó ý : c¸c dßng m· d­íi ®©y chØ ®­îc thùc hiÖn ë th¸ng ®Çu kú)
'      If ThongSo.Thang = pThangDauKy Then
      If (ThongSo.thang = 1 Or ThongSo.thang = CThangDB(ThangTang)) And (Not KhongDC(MaSo)) Then
            sql = "SELECT MaSo AS F1 FROM CTTaiSan " _
                                                                        & "WHERE MaTS = " + CStr(MaSo) _
                                                                        + " AND Thang = " + CStr(ThongSo.thang) _
                                                                        + " AND MaLoai = " + CStr(NV_DGLAI)
            If IsNull(SelectSQL(sql)) Then
                  SuaDoi = -3
                  GoTo KhongDC
            Else
                  SuaDoi = ThongSo.SuaDoiGiaTri
                  If SuaDoi = -1 Then GoTo KhongDC
                  sql = "UPDATE CTTaiSan SET NG_NS = " + DoiDau(ThongSo.NG_NS) _
                        + ", NG_TBS = " + DoiDau(ThongSo.NG_TBS) + ", NG_CNK = " + DoiDau(ThongSo.NG_CNK) _
                        + ", NG_TD = " + DoiDau(ThongSo.NG_TD) + ", CL_NS = " + DoiDau(ThongSo.CL_NS) _
                        + ", CL_TBS = " + DoiDau(ThongSo.CL_TBS) + ", CL_CNK = " + DoiDau(ThongSo.CL_CNK) _
                        + ", CL_TD = " + DoiDau(ThongSo.CL_TD) _
                        + " WHERE MaTS = " + CStr(MaSo) + " AND MaLoai = " + CStr(IIf(ThangTang = 0, DK_LOAI, NV_TANG))
                  If ExecuteSQL5(sql) <> 0 Then
                        SuaDoi = -1
                        GoTo KhongDC
                  End If
            End If
      End If
      WSpace.CommitTrans
      Exit Function
KhongDC:
        WSpace.Rollback
End Function
'======================================================================================
' CLASS
'======================================================================================
Public Property Get TongNguyenGia() As Double
      TongNguyenGia = ThongSo.NG_NS + ThongSo.NG_TBS + ThongSo.NG_CNK + ThongSo.NG_TD
End Property

'======================================================================================
' Tæng gi¸ trÞ cßn l¹i
'======================================================================================
Public Property Get TongGiaTri() As Double
      TongGiaTri = ThongSo.CL_NS + ThongSo.CL_TBS + ThongSo.CL_CNK + ThongSo.CL_TD
End Property

'======================================================================================
' §iÒu chØnh l¹i chøng tõ t¨ng
'======================================================================================
Private Function SetChungtuTang(tong_ng As Double, tong_hm As Double) As Integer
    Dim rs_ps As Recordset, ctu As New ClsChungtu
    Dim sopsco As Double, sopsno As Double
    Dim tongpsco As Double
    Dim sopsts As Integer, sopsnots As Double, sohmcu As Double

    SetChungtuTang = -1
        
    Set rs_ps = DBKetoan.OpenRecordset("SELECT ChungTu.MaSo, ChungTu.SoPS, ChungTu.MaTKNo, ChungTu.MaTKCo, HeThongTK.TK_ID As IDNo, HeThongTK_1.TK_ID As IDCo" _
        & " FROM ((CTTaiSan INNER JOIN ChungTu ON CTTaiSan.MaCTKT = ChungTu.MaCT) LEFT JOIN HeThongTK ON ChungTu.MaTKNo = HeThongTK.MaSo) LEFT JOIN HeThongTK AS HeThongTK_1 ON ChungTu.MaTKCo = HeThongTK_1.MaSo" _
        & " WHERE (CTTaiSan.MaTS = " + CStr(MaSo) + ") AND (CTTaiSan.MaLoai = " + CStr(NV_TANG) + ")", dbOpenSnapshot)
    
    If rs_ps!MaTkNo = 0 Or rs_ps!MaTkCo = 0 Then
        Do While Not rs_ps.EOF
            If rs_ps!MaTkNo > 0 Then
                If rs_ps!IDNO = TSCD_ID Then
                    ctu.InitChungtu rs_ps!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
                    ctu.XoaChungtu
                    ctu.sops = tong_ng
                    ctu.GhiChungtu
                Else
                    sopsno = sopsno + rs_ps!sops
                End If
            Else
                If rs_ps!IDCo = KHTSCD_ID Then
                    ctu.InitChungtu rs_ps!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
                    ctu.XoaChungtu
                    If tong_hm > 0 Then
                        ctu.sops = tong_hm
                        ctu.GhiChungtu
                    End If
                Else
                    sopsco = sopsco + rs_ps!sops
                End If
            End If
            rs_ps.MoveNext
        Loop
        tongpsco = tong_ng - tong_hm + sopsno
        rs_ps.MovePrevious
        Do While Not rs_ps.BOF
            If rs_ps!MaTkCo > 0 And rs_ps!IDCo <> KHTSCD_ID Then
                ctu.InitChungtu rs_ps!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
                ctu.XoaChungtu
                ctu.sops = RoundMoney(rs_ps!sops * tongpsco / sopsco)
                ctu.GhiChungtu
            End If
            rs_ps.MovePrevious
        Loop
    Else
        Do While Not rs_ps.EOF
            If rs_ps!IDNO = TSCD_ID Then
                sopsts = sopsts + 1
                sopsnots = sopsnots + rs_ps!sops
            Else
                sopsno = sopsno + rs_ps!sops
            End If
            If rs_ps!IDCo <> KHTSCD_ID Then
                sopsco = sopsco + rs_ps!sops
            Else
                sohmcu = rs_ps!sops
            End If
            rs_ps.MoveNext
        Loop
        rs_ps.MovePrevious
        If sopsts > 1 Then
            Do While Not rs_ps.BOF
                If rs_ps!IDNO = TSCD_ID Then
                    ctu.InitChungtu rs_ps!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
                    ctu.XoaChungtu
                    ctu.sops = RoundMoney(rs_ps!sops * tong_ng / sopsnots)
                    ctu.GhiChungtu
                End If
                rs_ps.MovePrevious
            Loop
        Else
            Do While Not rs_ps.BOF
                If rs_ps!IDNO = TSCD_ID Then
                    ctu.InitChungtu rs_ps!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
                    ctu.XoaChungtu
                    ctu.sops = RoundMoney(rs_ps!sops * tong_ng / sopsnots)
                    ctu.GhiChungtu
                End If
                rs_ps.MovePrevious
            Loop
        End If
    End If
        
    SetChungtuTang = 0
    rs_ps.Close
    Set rs_ps = Nothing
    Set ctu = Nothing
End Function

Private Sub Class_Initialize()
    ThangGiam = 13
End Sub

Private Sub Class_Terminate()
    Set ThongSo = Nothing
End Sub
