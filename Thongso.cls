VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsThongSo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
' Thuéc tÝnh
'======================================================================================
Public MaTS As Long
Public thang As Integer
' Nguyªn gi¸
Public NG_NS As Double
Public NG_TBS As Double
Public NG_CNK As Double
Public NG_TD As Double
' Gi¸ trÞ cßn l¹i
Public CL_NS As Double
Public CL_TBS As Double
Public CL_CNK As Double
Public CL_TD As Double
' L­îng hao mßn
Public HM_NS As Double
Public HM_TBS As Double
Public HM_CNK As Double
Public HM_TD As Double
' L­îng khÊu hao hµng th¸ng
Public KH_NS As Double
Public KH_TBS As Double
Public KH_CNK As Double
Public KH_TD As Double
' M· sè cña c¸c ®èi t­îng quan hÖ
Public MaDTQL As Long
Public MaDTSD As Long
Public MaTTSD As Long
'======================================================================================
' FUNCTION HopLe : KiÓm tra d÷ liÖu
'                          Tr¶ vÒ : 0 nÕu hîp lÖ, -1 nÕu cã lçi
'======================================================================================
Public Function HopLe()
Dim thong_bao As String
      If MaDTQL = 0 Then thong_bao = "ThiÕu hoÆc ch­a ®¨ng ký ®èi t­îng qu¶n lý": GoTo Err_InValidate
      If MaDTSD = 0 Then thong_bao = "ThiÕu hoÆc ch­a ®¨ng ký ®èi t­îng sö dông": GoTo Err_InValidate
      If MaTTSD = 0 Then thong_bao = "ThiÕu hoÆc ch­a ®¨ng ký t×nh tr¹ng sö dông": GoTo Err_InValidate
'      If NG_NS + NG_TBS + NG_CNK + NG_TD = 0 Then thong_bao = "Nguyªn gi¸ cña tµi s¶n thiÕu hoÆc kh«ng chÝnh x¸c": GoTo Err_InValidate
      If (NG_NS = 0 And Not KH_NS = 0) Or (NG_TBS = 0 And Not KH_TBS = 0) Or _
                                          (NG_CNK = 0 And Not KH_CNK = 0) Or (NG_TD = 0 And Not KH_TD = 0) Then _
            thong_bao = "L­îng khÊu hao chØ cã nÕu tµi s¶n cã thµnh phÇn nguån vèn t­¬ng øng": GoTo Err_InValidate
      If (CL_NS < KH_NS) Or (CL_TBS < KH_TBS) Or (CL_CNK < KH_CNK) Or (CL_TD < KH_TD) Then _
            thong_bao = "L­îng khÊu hao lín h¬n gi¸ trÞ cßn l¹i cña tµi s¶n": GoTo Err_InValidate
      HopLe = 0
      Exit Function
Err_InValidate:
      Beep
      MsgBox thong_bao, vbCritical, App.ProductName
      HopLe = -1
End Function
'======================================================================================
' FUNCTION ThemMoi : Thªm mét b¶n ghi chøa gi¸ trÞ hµng th¸ng cña tµi s¶n
'                              Tr¶ vÒ : 0 nÕu thµnh c«ng, -1 nÕu cã lçi
'======================================================================================
Public Function ThemMoi() As Integer
    Dim sql As String
    
    If thang > 0 Then
        sql = "INSERT INTO ThongSo ( MaSo, MaTS, Thang, " _
              & "NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD, " _
              & "KH_NS, KH_TBS, KH_CNK, KH_TD, MaDTQL, MaDTSD, MaTTSD ) " _
              & "VALUES (" + CStr(Lng_MaxValue("MaSo", "ThongSo") + 1) + "," + CStr(MaTS) + "," + CStr(thang) _
              + "," + DoiDau(NG_NS) + "," + DoiDau(NG_TBS) + "," + DoiDau(NG_CNK) + "," + DoiDau(NG_TD) _
              + "," + DoiDau(CL_NS) + "," + DoiDau(CL_TBS) + "," + DoiDau(CL_CNK) + "," + DoiDau(CL_TD) _
              + "," + DoiDau(KH_NS) + "," + DoiDau(KH_TBS) + "," + DoiDau(KH_CNK) + "," + DoiDau(KH_TD) _
              + "," + CStr(MaDTQL) + "," + CStr(MaDTSD) + "," + CStr(MaTTSD) + ")"
    Else
        sql = "INSERT INTO ThongSo ( MaSo, MaTS, Thang, " _
              & "NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD, " _
              & "MaDTQL, MaDTSD, MaTTSD ) " _
              & "VALUES (" + CStr(Lng_MaxValue("MaSo", "ThongSo") + 1) + "," + CStr(MaTS) + ",0," _
              + DoiDau(NG_NS) + "," + DoiDau(NG_TBS) + "," + DoiDau(NG_CNK) + "," + DoiDau(NG_TD) _
              + "," + DoiDau(CL_NS) + "," + DoiDau(CL_TBS) + "," + DoiDau(CL_CNK) + "," + DoiDau(CL_TD) _
              + "," + CStr(MaDTQL) + "," + CStr(MaDTSD) + "," + CStr(MaTTSD) + ")"
    End If
    ThemMoi = ExecuteSQL5(sql)
End Function
'======================================================================================
' SUB ChiDinh : G¸n c¸c thuéc tÝnh cña ®èi t­îng hiÖn t¹i theo gi¸ trÞ l­u trong c¬ së d÷ liÖu
'         Tham sè : M· tµi s¶n, th¸ng cÇn lÊy th«ng tin
'               Chó ý : Ngoµi m· sè cña c¸c ®èi t­îng quan hÖ, c¸c thuéc tÝnh gi¸ trÞ sÏ ®­îc tÝnh to¸n l¹i theo
'                             d÷ liÖu l­u trong chøng tõ cã liªn quan vµ l­îng khÊu hao hµng th¸ng.
'======================================================================================
Public Sub ChiDinh(ma_ts As Long, thang_cd As Integer)
Dim rs_thongso As Recordset, sql As String
      ' C¸c ®èi t­îng quan hÖ
      sql = "SELECT  * FROM ThongSo WHERE MaTS = " + CStr(ma_ts) + " AND Thang = " + CStr(CThangDB(thang_cd))
      Set rs_thongso = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      If rs_thongso.RecordCount > 0 Then
            MaTS = rs_thongso!MaTS
            thang = rs_thongso!thang
            MaDTQL = rs_thongso!MaDTQL
            MaDTSD = rs_thongso!MaDTSD
            MaTTSD = rs_thongso!MaTTSD
        End If
      rs_thongso.Close
      Set rs_thongso = Nothing
      ' Nguyªn gi¸, gi¸ trÞ cßn l¹i, l­îng hao mßn vµ khÊuhao
      ' (gi¸ trÞ ®· ®­îc ®iÒu chØnh vµ kh«ng trÝch khÊu hao)
      TinhGiaTriTaiSan MaTS, thang_cd, KH_KHONG
      With GiaTri
            NG_NS = .NG_NS
            NG_TBS = .NG_TBS
            NG_CNK = .NG_CNK
            NG_TD = .NG_TD
            CL_NS = .CL_NS
            CL_TBS = .CL_TBS
            CL_CNK = .CL_CNK
            CL_TD = .CL_TD
            KH_NS = .KH_NS
            KH_TBS = .KH_TBS
            KH_CNK = .KH_CNK
            KH_TD = .KH_TD
      End With
      ' L­îng hao mßn
      HM_NS = NG_NS - CL_NS
      HM_TBS = NG_TBS - CL_TBS
      HM_CNK = NG_CNK - CL_CNK
      HM_TD = NG_TD - CL_TD
End Sub
'======================================================================================
' SUB KhoiTao
'======================================================================================
Public Sub KhoiTao()
      MaTS = 0
      thang = 0
      ' Nguyªn gi¸
      NG_NS = 0
      NG_TBS = 0
      NG_CNK = 0
      NG_TD = 0
      ' Gi¸ trÞ cßn l¹i
      CL_NS = 0
      CL_TBS = 0
      CL_CNK = 0
      CL_TD = 0
      ' Hao mßn
      HM_NS = 0
      HM_TBS = 0
      HM_CNK = 0
      HM_TD = 0
      ' KhÊu hao
      KH_NS = 0
      KH_TBS = 0
      KH_CNK = 0
      KH_TD = 0
      ' C¸c ®èi t­îng quan hÖ
      MaDTQL = 0
      MaDTSD = 0
      MaTTSD = 0
End Sub
'======================================================================================
' FUNCTION SuaDoiQuanHe : Thay ®æi luü tiÕn l­îng khÊu hao vµ c¸c ®èi t­îng quan hÖ
'                                              Chó ý : C¸c thay ®æi vÒ l­îng khÊu hao sÏ cã thÓ lµm sai lÖch l­îng khÊu hao
'                                                            ®· ®­îc ®iÒu chØnh.
'======================================================================================
Public Function SuaDoiQuanHe(dau_ky As Boolean)
    Dim sql As String
    
      sql = "UPDATE ThongSo SET " _
            + "KH_NS = " + DoiDau(KH_NS) + ", KH_TBS = " + DoiDau(KH_TBS) _
            + ", KH_CNK = " + DoiDau(KH_CNK) + ", KH_TD = " + DoiDau(KH_TD) _
            + ", MaDTQL = " + CStr(MaDTQL) + ", MaDTSD = " + CStr(MaDTSD) + ", MaTTSD = " + CStr(MaTTSD) _
            + " WHERE Thang>=1 AND MaTS = " + CStr(MaTS) + IIf(dau_ky = True, "", " AND Thang >= " + CStr(thang))
      SuaDoiQuanHe = ExecuteSQL5(sql)
End Function
'======================================================================================
' FUNCTION SuaDoiGiaTri : Thay ®æi gi¸ trÞ cho c¸c th¸ng tr­íc th¸ng ®Çu kú hoÆc th¸ng t¨ng.
'                                         Chó ý : Thñ tôc nµy kh«ng cã ý nghÜa ®èi víi c¸c th¸ng sau th¸ng t¨ng hoÆc th¸ng
'                                                        ®Çu kú v× gi¸ trÞ tµi s¶n sÏ ®­îc cËp nhËt l¹i, nh­ng ®èi víi c¸c th¸ng tr­íc
'                                                        th× cÇn thiÕt v× ph¶i duy tr× l­îng tæng nguyªn gi¸ ®Çu kú cho c¸c b¸o c¸o.
'======================================================================================
Public Function SuaDoiGiaTri()
    Dim sql As String
    
      sql = "UPDATE ThongSo SET " _
            + "NG_NS = " + DoiDau(NG_NS) + ", NG_TBS = " + DoiDau(NG_TBS) _
            + ", NG_CNK = " + DoiDau(NG_CNK) + ", NG_TD = " + DoiDau(NG_TD) _
            + ", CL_NS = " + DoiDau(CL_NS) + ", CL_TBS = " + DoiDau(CL_TBS) _
            + ", CL_CNK = " + DoiDau(CL_CNK) + ", CL_TD = " + DoiDau(CL_TD) _
            + " WHERE MaTS = " + CStr(MaTS) + " AND Thang  <= " + CStr(thang)
      SuaDoiGiaTri = ExecuteSQL5(sql)
End Function

Public Sub DanXuat(ts1 As clsThongSo)
    With ts1
        .MaTS = MaTS
        .thang = thang
        ' Nguyªn gi¸
        .NG_NS = NG_NS
        .NG_TBS = NG_TBS
        .NG_CNK = NG_CNK
        .NG_TD = NG_TD
        ' Gi¸ trÞ cßn l¹i
        .CL_NS = CL_NS
        .CL_TBS = CL_TBS
        .CL_CNK = CL_CNK
        .CL_TD = CL_TD
        ' Hao mßn
        .HM_NS = HM_NS
        .HM_TBS = HM_TBS
        .HM_CNK = HM_CNK
        .HM_TD = HM_TD
        ' KhÊu hao
        .KH_NS = KH_NS
        .KH_TBS = KH_TBS
        .KH_CNK = KH_CNK
        .KH_TD = KH_TD
        ' C¸c ®èi t­îng quan hÖ
        .MaDTQL = MaDTQL
        .MaDTSD = MaDTSD
        .MaTTSD = MaTTSD
    End With
End Sub
