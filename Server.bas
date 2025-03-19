Attribute VB_Name = "ModServer"
Option Explicit

Public WSpace As Workspace
Public DBKetoan As Database

Public DEMO As Integer
Private Const EXESize = 0
Public Const pVersion = 1                   ' 0: TNHH,  1: LDCP,    2: Dao tao      3:  HCSN
Public Const pMST = ""

Public Const pRev = 404
Private Const MinFreeSpace = 50            ' MB
Public pPSW As String

Public pNN As Integer
Public pSTOP As Integer
Public pTien As Long
Public pTienStr As String
Public Mask_N As Long
Public pTygia As Integer
Public pTyGiaBQ As Integer
Public pBaoGia As Integer
Public pNVBH As Integer
Public pPhieu As Integer
Public pCongNoHD As Integer
Public pPQTK As Integer
Public pGiaUSD As Integer
Public pChietKhau As Integer
Public pKiemKeNgay As Integer
Public pNoiBo As Integer
Public pNhapKhau As Integer
Public pNhapDoiTuong As Integer
Public pProcessMode As Integer

Public pVV(1 To 3) As String
Public pSoVV As Integer

Public Const TKVT_ID = 1000
Public Const GTGTKT_ID = 1330
Public Const GTGTPN_ID = 3007
Public Const TKCNKH_ID = 3500
Public Const TKVNH_ID = 3800
Public Const TSCD_ID = 2000
Public Const KHTSCD_ID = 2001
Public Const TKCPVL_ID = 3001
Public Const TKCPNC_ID = 3002
Public Const TKCPTB_ID = 3010
Public Const TKCPSX_ID = 3003
Public Const TKCPGIA_ID = 3004
Public Const TKCPQL_ID = 3005
Public Const TKCPBH_ID = 3006
Public Const TKDT_ID = 5000
Public Const TKThue_ID = 3007
Public Const TKPTPN_ID = 4000
Public Const TKTNBT_ID = 7000
Public Const TKCPBT_ID = 8000
Public Const TKXDKQ_ID = 9000
Public Const TKCNPT_ID = 3310
Public Const TTDB_ID = 3332
Public Const TKGT_ID = 5210
Public Const TKLT_ID = 1310

Public pVATV As String
Public pSHPT As String

Public Const MaxGridRow = 16352
Public Const Max1000 = 1000

Public UserID As Long
Public UserName As String
Public User_Right As Long
Public pMaxUser As Integer
Public pOpenCount As Integer
Public pWinDir As String

Public Const max_level = 6

Public pDataPath As String
Public pThang As Integer
Public pThangDauKy As Integer
Public pTKTrunggian As Integer
Public pCurDir As String
Public pSongNgu As Boolean
Public pRpt As Integer

Public pTenCty As String
Public pTenCn As String
' Format
Public Mask_0 As String
Public Mask_2 As String
Public Mask_D As String
Public Const Mask_DB = "mm/dd/yy"
Public Const Mask_DR = "dd/mm/yyyy"

Public pFontName As String
Public pFontSize As Integer
Public sDecimal  As String
Public pSoKT As Integer
Public pKhongDau  As Integer

Public pNamTC As Integer
Public IniPath As String
Public FontFlag As Integer

Public OutCost As Integer
Public STDetail As Boolean
Public FADetail As Boolean
Public KHDetail As Boolean
Public FCost As Boolean
Public pGiaHT As Integer
Public pHachToan As Integer
Public pGiaVon As Integer
Public NgayDauThangMoi As Integer
Public pDTTP As Integer
Public pMaVach As Integer
Public pDinhmuc As Integer
Public CTGS_GV As Long
Public pBarCode As Integer

Public ShTkSPDo As String               ' SË hi÷u tµi kho∂n s∂n ph»m dÎ
Public ShTkTP As String               ' SË hi÷u tµi kho∂n s∂n ph»m dÎ
Public ShTkKQ As String                  ' SË hi÷u tµi kho∂n ket qua

Public SHCT_Len As Integer
Public Const MaxKC = 48

Public pFunction As Integer
Public pCT_ID As Long

Sub Main()
    If App.PrevInstance Then
        End
    End If
    
    If Not CheckMinRez(800, 600) Then
        MsgBox "C«n Æ∆t mµn h◊nh tËi thi”u lµ 800x600", vbCritical, App.ProductName
    End If
    
    Screen.MousePointer = 11
    
    pCurDir = App.Path
    If Right(pCurDir, 1) <> "\" Then pCurDir = pCurDir + "\"
    If EXESize > 0 Then
        If FileLen(pCurDir + App.EXEName + ".EXE") < EXESize - 512 Or FileLen(pCurDir + App.EXEName + ".EXE") > EXESize + 512 Then End
    End If
    IniPath = Dir(App.Path, vbDirectory)
    
    If Not FontDaCo(sFONTNAME) Then Add32Font "VKNT.FON"
    
    pPSW = Chr(8) + Chr(13) + Chr(27) + Chr(27)
    
    'Load frmSplash
    frmSplash.Show
    frmSplash.Refresh
    frmSplash.StartSAS
     
    Load frmMain
    
    Unload frmSplash
    Set frmSplash = Nothing
     
    frmMain.Show 0
    Screen.MousePointer = 0
End Sub
'======================================================================================
' Sub OpenDB
'======================================================================================
Public Function OpenDB(file_name As String) As Integer
    Dim retry As Integer, s As Long, msg As String
    pSTOP = 0
    OpenDB = -1
    If Len(Dir(file_name)) = 0 Then
        msg = "Kh´ng t◊m th y t÷p d˜ li÷u !"
        GoTo KhongMo
    End If
    s = GetAttr(file_name)
    If s Mod 2 = 1 Then
        msg = "T÷p d˜ li÷u bﬁ Æ∆t chÿ Æ‰c - Read Only !"
        GoTo KhongMo
    End If
    retry = 0
    On Error GoTo DB_Handle
      ' Mo tep du lieu duoc chi dinh23
Op:
      Set DBKetoan = WSpace.OpenDatabase(file_name, False, False, ";PWD=" + pPSW)
      On Error GoTo 0
      
      pDataPath = file_name
      If pOpenCount > 0 Then HienThongBao file_name, 2
      s = GetDiskSpace
      OpenDB = 0
      If s > 0 And s < MinFreeSpace Then
            msg = "C«n t®ng dung l≠Óng trËng Æ” l≠u tr˜ sË li÷u, chÿ cﬂn " + CStr(s) + " MB !"
            GoTo KhongMo
      End If
      pOpenCount = pOpenCount + 1
      Exit Function
' Khong mo duoc du lieu
DB_Handle:
    If Err.Number = 3031 Then
        If SetPsw(file_name, "unlock$shark$" + Chr(8) + Chr(13) + Chr(27), pPSW) <> 0 Then GoTo KhongMo
        GoTo Op
    End If
    If retry > 10 Then
        msg = "Kh´ng mÎ Æ≠Óc t÷p d˜ li÷u !"
    Else
      retry = retry + 1
      DoEvents
      Resume
    End If
KhongMo:
    If Len(msg) > 0 Then MsgBox msg, vbCritical, App.ProductName
End Function
'======================================================================================
' Sub CloseUp
'======================================================================================
Public Sub CloseUp(Optional nen As Integer = 0)
    Dim stt As Integer, dpath As String, id As Long, fn As String, dpath2 As String
            
    ExecuteSQL5 "UPDATE Users SET WS='...' WHERE MaSo=" + CStr(UserID), False
    id = SelectSQL("SELECT TenCty_ID AS F1 FROM License")
    DBKetoan.Close
End Sub
'======================================================================================
' ThÒ tÙc th˘c hi÷n cÀp nhÀt
'======================================================================================
Public Function ExecuteSQL5(sql As String, Optional msg As Boolean = True) As Integer
      On Error GoTo ErrLock
      DBKetoan.Execute sql, dbFailOnError
      On Error GoTo 0
      ExecuteSQL5 = 0
      Exit Function
ErrLock:
If msg Then
    Select Case Err.Number
        Case 13:
            MsgBox "Kh´ng nhÀn dπng Æ≠Óc ki”u d˜ li÷u", vbExclamation, App.ProductName
        Case 3040, 3043, 3566, 3572, 3575:
            MsgBox "LÁi Æ‰c ghi Æ‹a. Ki”m tra lπi Æ‹a c¯ng ho∆c mπng !", vbExclamation, App.ProductName
        Case 3006, 3008, 3009, 3202, 3045, 3046, 3186, 3187, 3188, 3189, 3196, 3211, 3212, 3218, 3260, 3261, 3262, 3356, 3576, 3623:
            DBEngine.Idle
            Resume
        Case 3200, 3201:
            MsgBox "ß∑ c„ th´ng tin, kh´ng xo∏ !", vbExclamation, App.ProductName
        Case 3022:
            MsgBox "TrÔng l∆p sË hi÷u, t™n... !", vbExclamation, App.ProductName
        Case 3075:
            MsgBox "Kh´ng sˆ dÙng d u ' trong c∏c chuÁi", vbExclamation, App.ProductName
        Case Else
            MsgBox "Error " + CStr(Err.Number) + ": " + Err.Description, vbExclamation, App.ProductName
    End Select
End If
    ExecuteSQL5 = Err.Number
End Function
'======================================================================================
' FUNCTION SelectSQL : ß‰c th´ng tin tı mÈt tr≠Íng (c„ Alias lµ F1) trong b∂n ghi vÌi Æi“u ki÷n cho trong bi’n chung pSelectSQL.
'                                 Tr∂ v“ : Null n’u kh´ng t◊m th y b∂n ghi tho∂ m∑n Æi“u ki÷n, th´ng tin c«n tr›ch d≠Ìi dπng Variant n’u c„.
'======================================================================================
Public Function SelectSQL(sql As String, Optional SoPhu As Variant = 0, Optional SoPhu2 As Variant = 0, Optional SoPhu3 As Variant = 0, Optional SoPhu4 As Variant = 0) As Variant
    Dim rs As Recordset
    
    On Error Resume Next
    SoPhu = 0
    SoPhu2 = 0
    SoPhu3 = 0
    SoPhu4 = 0
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.RecordCount > 0 Then
            If IsNull(rs!f1) Then
                SelectSQL = 0
            Else
                SelectSQL = rs!f1
                SoPhu = rs!F2
                SoPhu2 = rs!F3
                SoPhu3 = rs!F4
                SoPhu4 = rs!f5
            End If
    Else
            SelectSQL = 0
    End If
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
End Function

Public Function SelectSQLDB(db As Database, sql As String, Optional SoPhu As Variant, Optional SoPhu1 As Variant, Optional SoPhu2 As Variant) As Variant
Dim rs As Recordset
      Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
      On Error Resume Next
      SoPhu = 0
      SoPhu1 = 0
      SoPhu2 = 0
      If rs.RecordCount > 0 Then
            If IsNull(rs!f1) Then
                SelectSQLDB = 0
            Else
                SelectSQLDB = rs!f1
                
                SoPhu = rs!F2
                SoPhu1 = rs!F3
                SoPhu2 = rs!F4
                
            End If
      End If
      rs.Close
      Set rs = Nothing
      On Error GoTo 0
End Function
'====================================================================================================
' Hµm thay ÆÊi nÈi dung Query Æ∑ khai b∏o trong CSDL
'====================================================================================================
Public Sub SetSQL(qname As String, sql As String)
    AddQuery qname, sql
    On Error GoTo QueryErr
    DBKetoan.QueryDefs(qname).sql = sql
    On Error GoTo 0
    Exit Sub
QueryErr:
    Select Case Err.Number
        Case 3006, 3008, 3009, 3202, 3045, 3046, 3186, 3187, 3188, 3189, 3196, 3211, 3212, 3218, 3260, 3261, 3262, 3356, 3576, 3623:
            DBEngine.Idle
            Resume
        Case Else
            MsgBox CStr(Err.Number) + ": " + Err.Description, vbExclamation, App.ProductName
    End Select
End Sub
'====================================================================================================
' Hµm tr∂ v“ sË ph∏t sinh theo ÆËi ¯ng gi˜a chi ti’t vµ tµi kho∂n tµi ch›nh
'====================================================================================================
Private Function PhatSinhDu(mct As Long, mtk As Long, tdau As Integer, tcuoi As Integer, loai As Integer) As Double
    Dim rs_kq As Recordset, sql As String
    
    sql = "SELECT SUM(SoPS) As KetQua FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi)
    If loai = -1 Then
        sql = sql + " AND MaTkNo = " + CStr(mct) + IIf(mtk > 0, " AND MaTkTCCo = " + CStr(mtk), "")
    Else
        sql = sql + " AND MaTkCo = " + CStr(mct) + IIf(mtk > 0, " AND MaTkTCNo = " + CStr(mtk), "")
    End If
    
    Set rs_kq = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If IsNull(rs_kq!KetQua) Then
        PhatSinhDu = 0
    Else
        PhatSinhDu = rs_kq!KetQua
    End If
    
    sql = "SELECT SUM(SoPS) As KetQua FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi)
    If loai = -1 Then
        sql = sql + " AND MaTkCo = " + CStr(mct) + IIf(mtk > 0, " AND MaTkTCNo = " + CStr(mtk), "")
    Else
        sql = sql + " AND MaTkNo = " + CStr(mct) + IIf(mtk > 0, " AND MaTkTCCo = " + CStr(mtk), "")
    End If
    
    Set rs_kq = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If Not IsNull(rs_kq!KetQua) Then
        PhatSinhDu = PhatSinhDu - rs_kq!KetQua
    End If
    
    rs_kq.Close
    Set rs_kq = Nothing
End Function

Public Sub LayThongtinCT(MaCT As Long, loai As Integer, Ten As String, DiaChi As String, Optional ctgoc As String, Optional makh As Long, Optional p As Integer = 0)
    Dim rs As Recordset
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW * FROM ChungTuLQ" + IIf(p > 0, "P", "") + " WHERE MaCT=" + CStr(MaCT) + " AND Loai=" + CStr(loai), dbOpenSnapshot)
    If rs.RecordCount > 0 Then
        Ten = rs!hoten
        DiaChi = rs!DiaChi
        ctgoc = rs!SoCTGoc
        makh = rs!makh
    Else
        Ten = "..."
        DiaChi = "..."
        ctgoc = "..."
        makh = 0
    End If
    rs.Close
    Set rs = Nothing
End Sub

Public Function LaySH(sh As String, stt As Integer, Optional c As String = "#") As String
    Dim pos As Integer, i As Integer, st As String
    
    pos = InStr(sh, c)
    If pos > 0 Then
        st = sh
        i = 1
        Do While i < stt And Len(st) > pos
            i = i + 1
            st = Right(st, Len(st) - pos)
            pos = InStr(st, c)
            If pos = 0 Then Exit Do
        Loop
        If i = stt Then
            If pos > 0 Then LaySH = Left(st, pos - 1) Else LaySH = st
        Else
            LaySH = "..."
        End If
    Else
        If c = "#" Then LaySH = sh
    End If
End Function

Public Function ST2MaSo(f As String, sh As String, tbl As String) As Long
    ST2MaSo = SelectSQL("SELECT MaSo AS F1 FROM " + tbl + " WHERE " + f + "='" + sh + "'")
End Function

Public Function SoHieu2MaSo(sh As String, tbl As String) As Long
    SoHieu2MaSo = SelectSQL("SELECT MaSo AS F1 FROM " + tbl + " WHERE SoHieu='" + sh + "'")
    If SoHieu2MaSo = 0 Then SoHieu2MaSo = SelectSQL("SELECT MaSo AS F1 FROM " + tbl + " WHERE LIKE='" + sh + "*'")
End Function

Public Function MaSo2SoHieu(ms As Long, tbl As String) As String
    MaSo2SoHieu = SelectSQL("SELECT SoHieu AS F1 FROM " + tbl + " WHERE MaSo=" + CStr(ms))
End Function

Public Sub SetDefaultValue()
    Dim i As Integer, j As Integer
    
    On Error Resume Next
    For i = 0 To DBKetoan.TableDefs.Count - 1
        For j = 0 To DBKetoan.TableDefs(i).Fields.Count - 1
                Select Case DBKetoan.TableDefs(i).Fields(j).Type
                    Case dbInteger:
                        ExecuteSQL5 "UPDATE " + DBKetoan.TableDefs(i).Name + " SET " + DBKetoan.TableDefs(i).Fields(j).Name + "=0 WHERE IsNull(" + DBKetoan.TableDefs(i).Fields(j).Name + ")", False
                    Case dbLong, dbDouble:  DBKetoan.TableDefs(i).Fields(j).DefaultValue = 0
                        ExecuteSQL5 "UPDATE " + DBKetoan.TableDefs(i).Name + " SET " + DBKetoan.TableDefs(i).Fields(j).Name + "=0 WHERE IsNull(" + DBKetoan.TableDefs(i).Fields(j).Name + ")", False
                    Case dbText:   DBKetoan.TableDefs(i).Fields(j).DefaultValue = "..."
                        ExecuteSQL5 "UPDATE " + DBKetoan.TableDefs(i).Name + " SET " + DBKetoan.TableDefs(i).Fields(j).Name + "='...' WHERE IsNull(" + DBKetoan.TableDefs(i).Fields(j).Name + ")", False
                    Case dbDate:   DBKetoan.TableDefs(i).Fields(j).DefaultValue = CVDate("#1/1/80#")
                        ExecuteSQL5 "UPDATE " + DBKetoan.TableDefs(i).Name + " SET " + DBKetoan.TableDefs(i).Fields(j).Name + "=#1/1/80# WHERE IsNull(" + DBKetoan.TableDefs(i).Fields(j).Name + ")", False
                End Select
        Next
    Next
    On Error GoTo 0
End Sub

Public Function GetRowNumber(q As String)
    Dim rs As Recordset
    
    Set rs = DBKetoan.OpenRecordset(q, dbOpenSnapshot)
    If Not rs.EOF Then rs.MoveLast
    GetRowNumber = rs.RecordCount
    rs.Close
    Set rs = Nothing
End Function

Public Function SetPsw(file_name As String, psw1 As String, psw2 As String) As Integer
    Dim p As String, i As Integer, px As String
    
    px = psw1
    SetPsw = -1
    On Error Resume Next

    DBKetoan.Close
    
    Do While True
        
        i = InStr(px, "$")
        If i > 0 Then
            p = Left(px, i - 1)
            px = Right(px, Len(px) - i)
        Else
            p = px
            px = ""
        End If
        Err.Number = 0
        Set DBKetoan = WSpace.OpenDatabase(file_name, True, False, ";PWD=" + p)
        If Err.Number = 0 Then
            DBKetoan.NewPassword p, psw2
            DBKetoan.Close
            If Err.Number = 0 Then Exit Do
        End If
        If Len(p) = 0 Then GoTo KT
    Loop
    SetPsw = 0
KT:
    On Error GoTo 0
End Function


Public Function TyGiaCuoi() As Double
    TyGiaCuoi = SelectSQL("SELECT TOP 1 TyGia AS F1 FROM ChungTu ORDER BY NgayCT DESC")
    If TyGiaCuoi = 0 Then TyGiaCuoi = 1
End Function

Private Function QueryDaCo2(qname As String, qname2 As String) As Boolean
    Dim i As Integer, L As Integer
    
    QueryDaCo2 = False
    L = Len(qname)
    For i = 0 To DBKetoan.QueryDefs.Count - 1
        If Left(UCase(DBKetoan.QueryDefs(i).Name), L) = UCase(qname) Then
            qname2 = DBKetoan.QueryDefs(i).Name
            QueryDaCo2 = True
            Exit For
        End If
    Next
End Function

Public Sub CboVVClick(cbo1 As ComboBox, cbo2 As ComboBox)
    Dim m As Long, sql As String

    If cbo1.ListIndex >= 0 And pSoVV > 1 Then
        m = cbo1.ItemData(cbo1.ListIndex)
        If SelectSQL("SELECT Max(MaSo) AS F1 FROM DoituongCT2 WHERE MaKhachHang=" + CStr(m)) > 0 Then
            sql = "SELECT MaSo As F2,DienGiai As F1 FROM DoituongCT2 WHERE MaKhachHang=" + CStr(m) + " ORDER BY SoHieu"
        Else
            sql = "SELECT MaSo As F2,DienGiai As F1 FROM DoituongCT2 ORDER BY SoHieu"
        End If
        Int_RecsetToCbo sql, cbo2
    End If
End Sub

Public Function LayMaPhanLoai(sh As String, tbl As String) As Long
    Dim i As Integer, ms As Long, m2 As Long
    
    For i = Len(sh) - 1 To 1 Step -1
        m2 = SelectSQL("SELECT MaSo AS F1, MaPhanLoai AS F2 FROM " + tbl + " WHERE SoHieu LIKE '" + Left(sh, i) + "*'", ms)
        If m2 > 0 Then Exit For
    Next
    If ms = 0 Then
        For i = Len(sh) - 1 To 1 Step -1
            ms = SelectSQL("SELECT MaSo AS F1 FROM PhanLoai" + tbl + " WHERE SoHieu LIKE '" + Left(sh, i) + "*'")
            If ms > 0 Then Exit For
        Next
    End If
    LayMaPhanLoai = ms
End Function

Public Sub KiemTraTaiKhoan(Optional a As Integer)

End Sub

Public Sub XoaCTTheoID(thang As Integer, id As Long, ml As Integer, Optional ndau As Date, Optional ncuoi As Date, Optional ml2 As Integer = 0)

End Sub

Public Sub KiemTraVatTu(Optional a As Integer)

End Sub
