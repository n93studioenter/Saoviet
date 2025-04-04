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
' Thu�c t�nh
'======================================================================================
Public MaTS As Long
Public thang As Integer
' Nguy�n gi�
Public NG_NS As Double
Public NG_TBS As Double
Public NG_CNK As Double
Public NG_TD As Double
' Gi� tr� c�n l�i
Public CL_NS As Double
Public CL_TBS As Double
Public CL_CNK As Double
Public CL_TD As Double
' L��ng hao m�n
Public HM_NS As Double
Public HM_TBS As Double
Public HM_CNK As Double
Public HM_TD As Double
' L��ng kh�u hao h�ng th�ng
Public KH_NS As Double
Public KH_TBS As Double
Public KH_CNK As Double
Public KH_TD As Double
' M� s� c�a c�c ��i t��ng quan h�
Public MaDTQL As Long
Public MaDTSD As Long
Public MaTTSD As Long
'======================================================================================
' FUNCTION HopLe : Ki�m tra d� li�u
'                          Tr� v� : 0 n�u h�p l�, -1 n�u c� l�i
'======================================================================================
Public Function HopLe()
Dim thong_bao As String
      If MaDTQL = 0 Then thong_bao = "Thi�u ho�c ch�a ��ng k� ��i t��ng qu�n l�": GoTo Err_InValidate
      If MaDTSD = 0 Then thong_bao = "Thi�u ho�c ch�a ��ng k� ��i t��ng s� d�ng": GoTo Err_InValidate
      If MaTTSD = 0 Then thong_bao = "Thi�u ho�c ch�a ��ng k� t�nh tr�ng s� d�ng": GoTo Err_InValidate
'      If NG_NS + NG_TBS + NG_CNK + NG_TD = 0 Then thong_bao = "Nguy�n gi� c�a t�i s�n thi�u ho�c kh�ng ch�nh x�c": GoTo Err_InValidate
      If (NG_NS = 0 And Not KH_NS = 0) Or (NG_TBS = 0 And Not KH_TBS = 0) Or _
                                          (NG_CNK = 0 And Not KH_CNK = 0) Or (NG_TD = 0 And Not KH_TD = 0) Then _
            thong_bao = "L��ng kh�u hao ch� c� n�u t�i s�n c� th�nh ph�n ngu�n v�n t��ng �ng": GoTo Err_InValidate
      If (CL_NS < KH_NS) Or (CL_TBS < KH_TBS) Or (CL_CNK < KH_CNK) Or (CL_TD < KH_TD) Then _
            thong_bao = "L��ng kh�u hao l�n h�n gi� tr� c�n l�i c�a t�i s�n": GoTo Err_InValidate
      HopLe = 0
      Exit Function
Err_InValidate:
      Beep
      MsgBox thong_bao, vbCritical, App.ProductName
      HopLe = -1
End Function
'======================================================================================
' FUNCTION ThemMoi : Th�m m�t b�n ghi ch�a gi� tr� h�ng th�ng c�a t�i s�n
'                              Tr� v� : 0 n�u th�nh c�ng, -1 n�u c� l�i
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
' SUB ChiDinh : G�n c�c thu�c t�nh c�a ��i t��ng hi�n t�i theo gi� tr� l�u trong c� s� d� li�u
'         Tham s� : M� t�i s�n, th�ng c�n l�y th�ng tin
'               Ch� � : Ngo�i m� s� c�a c�c ��i t��ng quan h�, c�c thu�c t�nh gi� tr� s� ���c t�nh to�n l�i theo
'                             d� li�u l�u trong ch�ng t� c� li�n quan v� l��ng kh�u hao h�ng th�ng.
'======================================================================================
Public Sub ChiDinh(ma_ts As Long, thang_cd As Integer)
Dim rs_thongso As Recordset, sql As String
      ' C�c ��i t��ng quan h�
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
      ' Nguy�n gi�, gi� tr� c�n l�i, l��ng hao m�n v� kh�uhao
      ' (gi� tr� �� ���c �i�u ch�nh v� kh�ng tr�ch kh�u hao)
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
      ' L��ng hao m�n
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
      ' Nguy�n gi�
      NG_NS = 0
      NG_TBS = 0
      NG_CNK = 0
      NG_TD = 0
      ' Gi� tr� c�n l�i
      CL_NS = 0
      CL_TBS = 0
      CL_CNK = 0
      CL_TD = 0
      ' Hao m�n
      HM_NS = 0
      HM_TBS = 0
      HM_CNK = 0
      HM_TD = 0
      ' Kh�u hao
      KH_NS = 0
      KH_TBS = 0
      KH_CNK = 0
      KH_TD = 0
      ' C�c ��i t��ng quan h�
      MaDTQL = 0
      MaDTSD = 0
      MaTTSD = 0
End Sub
'======================================================================================
' FUNCTION SuaDoiQuanHe : Thay ��i lu� ti�n l��ng kh�u hao v� c�c ��i t��ng quan h�
'                                              Ch� � : C�c thay ��i v� l��ng kh�u hao s� c� th� l�m sai l�ch l��ng kh�u hao
'                                                            �� ���c �i�u ch�nh.
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
' FUNCTION SuaDoiGiaTri : Thay ��i gi� tr� cho c�c th�ng tr��c th�ng ��u k� ho�c th�ng t�ng.
'                                         Ch� � : Th� t�c n�y kh�ng c� � ngh�a ��i v�i c�c th�ng sau th�ng t�ng ho�c th�ng
'                                                        ��u k� v� gi� tr� t�i s�n s� ���c c�p nh�t l�i, nh�ng ��i v�i c�c th�ng tr��c
'                                                        th� c�n thi�t v� ph�i duy tr� l��ng t�ng nguy�n gi� ��u k� cho c�c b�o c�o.
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
        ' Nguy�n gi�
        .NG_NS = NG_NS
        .NG_TBS = NG_TBS
        .NG_CNK = NG_CNK
        .NG_TD = NG_TD
        ' Gi� tr� c�n l�i
        .CL_NS = CL_NS
        .CL_TBS = CL_TBS
        .CL_CNK = CL_CNK
        .CL_TD = CL_TD
        ' Hao m�n
        .HM_NS = HM_NS
        .HM_TBS = HM_TBS
        .HM_CNK = HM_CNK
        .HM_TD = HM_TD
        ' Kh�u hao
        .KH_NS = KH_NS
        .KH_TBS = KH_TBS
        .KH_CNK = KH_CNK
        .KH_TD = KH_TD
        ' C�c ��i t��ng quan h�
        .MaDTQL = MaDTQL
        .MaDTSD = MaDTSD
        .MaTTSD = MaTTSD
    End With
End Sub
