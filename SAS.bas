Attribute VB_Name = "ModSAS"
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
Public pTrungSoHieuKhacThang As Integer

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

Public ShTkSPDo As String               ' Sè hiÖu tµi kho¶n s¶n phÈm dë
Public ShTkTP As String               ' Sè hiÖu tµi kho¶n s¶n phÈm dë
Public ShTkKQ As String                  ' Sè hiÖu tµi kho¶n ket qua

Public SHCT_Len As Integer
Public Const MaxKC = 48
Public opotion_1 As String
Public pFunction As Integer
Public pCT_ID As Long
Dim n1 As Date, n2 As Date
Public ban_quyen As Integer
Sub Main()
    opotion_1 = "1000"
    ban_quyen = 0
    If Not CheckMinRez(800, 600) Then
        MsgBox "CÇn ®Æt mµn h×nh tèi thiÓu lµ 800x600", vbCritical, App.ProductName
    End If

    Screen.MousePointer = 11

    pCurDir = App.path
    If Right(pCurDir, 1) <> "\" Then pCurDir = pCurDir + "\"
    If EXESize > 0 Then
        If FileLen(pCurDir + App.EXEName + ".EXE") < EXESize - 512 Or FileLen(pCurDir + App.EXEName + ".EXE") > EXESize + 512 Then End
    End If
    IniPath = Dir(App.path, vbDirectory)

    If Not FontDaCo(sFONTNAME) Then Add32Font "VKNT.FON"

    pPSW = "1@35^7*9)1"
    'pPSW = Chr(8) + Chr(13) + Chr(27) + Chr(27) + Chr(8) + Chr(13) + Chr(27) + Chr(27) + Chr(8) + Chr(8) + Chr(13) + Chr(13)

    'Chr(8) + Chr(13) + Chr(27) + Chr(27)
    'Load frmSplash
    frmSplash.Show
    frmSplash.Refresh
    frmSplash.StartSAS

    Load frmMain

    Unload frmSplash
    Set frmSplash = Nothing

    frmMain.Show 0
    Screen.MousePointer = 0

    FrmMatkhau.Show 1
    Set FrmMatkhau = Nothing
End Sub
'======================================================================================
' Sub OpenDB
'======================================================================================
Public Function OpenDB(file_name As String, Optional chonluu As Integer = 0) As Integer
    Dim retry As Integer, s As Long, msg As String, mst As String
    pSTOP = 0
    OpenDB = -1

    Dim st_1, st_2, st_3, st_4


    st_1 = "as"
    st_1 = "1"
    st_2 = "@"
    '    If Len(Dir(file_name)) = 0 Then
    '        msg = "Kh«ng t×m thÊy tÖp d÷ liÖu !"
    '        GoTo LDB
    '    End If
    Dim st_5, st_6, st_7, st_8

    If Len(Dir(file_name)) = 0 Then
        st_8 = "$"
        If Len(Dir(pCurDir + "DATA\QD48.MDB")) = 0 Then

            If Len(Dir(pCurDir + "DATA\QD15.MDB")) = 0 Then
                msg = "Kh«ng t×m thÊy tÖp d÷ liÖu !"
                GoTo LDB
            Else
                file_name = pCurDir + "DATA\QD15.MDB"
            End If
        Else
            file_name = pCurDir + "DATA\QD48.MDB"
        End If
    End If
    s = GetAttr(file_name)
    If s Mod 2 = 1 Then
        msg = "TÖp d÷ liÖu bÞ ®Æt chØ ®äc - Read Only !"
        GoTo KhongMo
    End If
    retry = 0
    On Error GoTo DB_Handle
    ' Mo tep du lieu duoc chi dinh23
Op:
    '   Set DBKetoan = WSpace.OpenDatabase(file_name, False, False, ";PWD= " + pPSW)
    '  Set DBKetoan = WSpace.OpenDatabase(file_name, False, False, ";PWD=" + pPSW)
    'pPSW = "1@35^7*9)"
    'On Error Resume Next    ' B?t l?i


    On Error Resume Next    ' B?t l?i
    Set DBKetoan = WSpace.OpenDatabase(file_name, False, False, ";PWD=" + pPSW)
    If Err.number <> 0 Then
        MsgBox "Database chua duoc cap Licence"
        End
        pPSW = Chr(8) + Chr(13) + Chr(27) + Chr(27) + Chr(8) + Chr(13) + Chr(27) + Chr(27) + Chr(8) + Chr(8) + Chr(13) + Chr(13)
        On Error Resume Next    ' B?t l?i
        Set DBKetoan = WSpace.OpenDatabase(file_name, False, False, ";PWD=" + pPSW)
        If Err.number <> 0 Then
            pPSW = "1@35^7*9)1"
            Set DBKetoan = WSpace.OpenDatabase(file_name, False, False, ";PWD=" + pPSW)
        End If
        On Error Resume Next    ' B?t l?i
    End If
    'On Error Resume Next    ' B?t l?i



    On Error GoTo 0
    Dim pas
    pDataPath = file_name
    If pOpenCount > 0 Then HienThongBao file_name, 2
    s = GetDiskSpace
    OpenDB = 0
    If s > 0 And s < MinFreeSpace Then
        msg = "CÇn t¨ng dung l­îng trèng ®Ó l­u tr÷ sè liÖu, chØ cßn " + CStr(s) + " MB !"
        GoTo KhongMo
    End If
    pOpenCount = pOpenCount + 1
    'Kiem tra co bi ma hoa hay khong


    Exit Function
    ' Khong mo duoc du lieu
    st_5 = "*e"
    st_6 = "e9"
DB_Handle:
    st_3 = "35q"
    st_7 = ")we"
    If Err.number = 3031 Then
        st_4 = "^7555"

        If SetPsw(file_name, Trim(Trim(ModSAS.Federo16Decrypt("dad`dccefucgcqcici", opotion_1))) + "$shark$" + Chr(8) + Chr(13) + Chr(27), pPSW) <> 0 Then GoTo KhongMo
        Dim rs As Recordset
        Exit Function
        Set DBKetoan = WSpace.OpenDatabase(file_name, False, False, ";PWD=" + pPSW)
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW License.* FROM License", dbOpenSnapshot)
        ExecuteSQL5 "update License set TenCty = '" + ModSAS.Federo16Decrypt(rs!TenCty, CStr(rs!NamTC)) + "',DiaChi = '" + ModSAS.Federo16Decrypt(rs!DiaChi, CStr(rs!NamTC)) + "',MaSoThue = '" + ModSAS.Federo16Decrypt(rs!masothue, CStr(rs!NamTC)) + "',CMP = '" + ModSAS.Federo16Decrypt(IIf(IsNull(rs!CMP), "", rs!CMP), CStr(rs!NamTC)) + "'"
        DBKetoan.Close
        GoTo Op
    End If
    If retry > 10 Then
        msg = "Kh«ng më ®­îc tÖp d÷ liÖu !"
LDB:
        If chonluu = 0 Then
            MsgBox msg, vbCritical, App.ProductName
            msg = ""
            ' If MsgBox("Xem danh s¸ch tÖp d÷ liÖu l­u tr÷ tù ®éng ?", vbYesNo, App.ProductName) = vbYes Then
            '     file_name = FrmDB.ChonTepLuu(mst, Year(Date))
            '     If Len(file_name) > 0 Then GoTo Op
            ' End If
        End If
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
    Dim stt As Integer, dpath As String, mst As String, fn As String, dpath2 As String
    
    CloseItemList
    On Error Resume Next
    ExecuteSQL5 "UPDATE Users SET WS='...' WHERE MaSo=" + CStr(UserID), False
    DelTemp
    mst = SelectSQL("SELECT MaSoThue AS F1 FROM License")
    DBKetoan.Close
    On Error GoTo 0
    SetPsw pDataPath, "", pPSW
        
    If nen = 0 Then
        For stt = Len(pDataPath) - 1 To 1 Step -1
        
            If Mid(pDataPath, stt, 1) = "\" Then Exit For
        Next
        dpath2 = "K" + IIf(Month(Date) < 10, "0", "") + CStr(Month(Date)) + IIf(Day(Date) < 10, "0", "") + CStr(Day(Date)) + Right(CStr(pNamTC), 2) + "_" + mst
        fn = dpath2 + ".SAS"
        dpath = Left(pDataPath, stt) + fn
        
        NenTep pDataPath, dpath
        On Error Resume Next
        FileCopy dpath, pWinDir + "\TEMP\" + dpath2
        On Error GoTo 0
        
        dpath2 = GetSetting(IniPath, "Environment", "BackUpPath")
        If Len(dpath2) > 0 Then
            If Right(dpath2, 1) <> "\" Then dpath2 = dpath2 + "\"
            On Error Resume Next
            FileCopy dpath, dpath2 + fn
            On Error GoTo 0
        End If
    End If
       
    stt = 0
    Do While stt < 5
        stt = stt + 1
        If UCase(GetSetting(IniPath, "RecentFiles", "File" + CStr(stt))) = UCase(pDataPath) Then Exit Do
    Loop
    
    Do While stt > 1
        SaveSetting IniPath, "RecentFiles", "File" + CStr(stt), GetSetting(IniPath, "RecentFiles", "File" + CStr(stt - 1))
        stt = stt - 1
    Loop
    
    SaveSetting IniPath, "RecentFiles", "File1", pDataPath
End Sub
'======================================================================================
' Thñ tôc thùc hiÖn cËp nhËt
'======================================================================================
Public Function ExecuteSQL5(sql As String, Optional msg As Boolean = True) As Integer
      On Error GoTo ErrLock
      DBKetoan.Execute sql, dbFailOnError
      On Error GoTo 0
      ExecuteSQL5 = 0
      Exit Function
ErrLock:
If msg Then
    Select Case Err.number
        Case 13:
            MsgBox "Kh«ng nhËn d¹ng ®­îc kiÓu d÷ liÖu", vbExclamation, App.ProductName
        Case 3040, 3043, 3566, 3572, 3575:
            MsgBox "Lçi ®äc ghi ®Üa. KiÓm tra l¹i ®Üa cøng hoÆc m¹ng !", vbExclamation, App.ProductName
        Case 3006, 3008, 3009, 3202, 3045, 3046, 3186, 3187, 3188, 3189, 3196, 3211, 3212, 3218, 3260, 3261, 3262, 3356, 3576, 3623:
            DBEngine.Idle
            Resume
        Case 3200, 3201:
            MsgBox "§· cã th«ng tin, kh«ng xo¸ !", vbExclamation, App.ProductName
        Case 3022:
            MsgBox "Trïng lÆp sè hiÖu, tªn... !", vbExclamation, App.ProductName
        Case 3075:
            MsgBox "Kh«ng sö dông dÊu ' trong c¸c chuçi", vbExclamation, App.ProductName
        Case Else
            MsgBox "Error " + CStr(Err.number) + ": " + Err.Description, vbExclamation, App.ProductName
    End Select
End If
    ExecuteSQL5 = Err.number
End Function
'======================================================================================
' FUNCTION SelectSQL : §äc th«ng tin tõ mét tr­êng (cã Alias lµ F1) trong b¶n ghi víi ®iÒu kiÖn cho trong biÕn chung pSelectSQL.
'                                 Tr¶ vÒ : Null nÕu kh«ng t×m thÊy b¶n ghi tho¶ m·n ®iÒu kiÖn, th«ng tin cÇn trÝch d­íi d¹ng Variant nÕu cã.
'======================================================================================
Public Function SelectSQL(sql As String, Optional SoPhu As Variant = 0, Optional SoPhu2 As Variant = 0, Optional SoPhu3 As Variant = 0, Optional SoPhu4 As Variant = 0) As Variant
    Dim rs As Recordset
    
    'On Error Resume Next
    SoPhu = 0
    SoPhu2 = 0
    SoPhu3 = 0
    SoPhu4 = 0
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.recordCount > 0 Then
            If IsNull(rs!f1) Then
                SelectSQL = 0
            Else
                SelectSQL = rs!f1
                On Error Resume Next
                SoPhu = rs!F2
                SoPhu2 = rs!F3
                SoPhu3 = rs!F4
                SoPhu4 = rs!f5
                On Error GoTo 0
            End If
    Else
            SelectSQL = 0
    End If
    rs.Close
    Set rs = Nothing
    'On Error GoTo 0
End Function
Public Function ExecuteSQL5_Themmoi(sql As String, Optional msg As Boolean = True) As Integer
      On Error GoTo ErrLock
      DBKetoan.Execute sql, dbFailOnError
      On Error GoTo 0
      ExecuteSQL5_Themmoi = 0
      Exit Function
ErrLock:
'MsgBox Err.Description
End Function

Public Function SelectSQLDB(db As Database, sql As String, Optional SoPhu As Variant, Optional SoPhu1 As Variant, Optional SoPhu2 As Variant) As Variant
Dim rs As Recordset
      Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
      On Error Resume Next
      SoPhu = 0
      SoPhu1 = 0
      SoPhu2 = 0
      If rs.recordCount > 0 Then
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
' Hµm thay ®æi néi dung Query ®· khai b¸o trong CSDL
'====================================================================================================
Public Sub SetSQL(qname As String, sql As String)
    AddQuery qname, sql
    On Error GoTo QueryErr
    DBKetoan.QueryDefs(qname).sql = sql
    On Error GoTo 0
    Exit Sub
QueryErr:
    Select Case Err.number
        Case 3006, 3008, 3009, 3202, 3045, 3046, 3186, 3187, 3188, 3189, 3196, 3211, 3212, 3218, 3260, 3261, 3262, 3356, 3576, 3623:
            DBEngine.Idle
            Resume
        Case Else
            MsgBox CStr(Err.number) + ": " + Err.Description, vbExclamation, App.ProductName
    End Select
End Sub
'======================================================================================
' Sub KiemTraDuLieu : KiÓm tra hÖ thèng tµi kho¶n
'======================================================================================
Public Sub KiemTraVatTu(Optional ktraxuat As Integer = 0)
      Dim rs_taikhoan As Recordset, j As Integer, st3 As String, m As Long, n As Long, mk As Long, m1 As Long, st4 As String
      Dim rs_ktra As Recordset, i As Integer, sql As String, st As String, st2 As String, idx As Index
    
      For i = 0 To DBKetoan.TableDefs("Vattu").Indexes.count - 1
        If DBKetoan.TableDefs("Vattu").Indexes(i).Name = "SoHieu" Then GoTo tt
     Next
          
     FixCode "Vattu", "MaSo"
     
      Set idx = DBKetoan.TableDefs("Vattu").CreateIndex("SoHieu")
      idx.Fields.Append idx.CreateField("SoHieu")
      idx.Unique = True
      On Error Resume Next
      DBKetoan.TableDefs("Vattu").Indexes.Append idx
      On Error GoTo 0
      
tt:
      If OutCost <> 0 Then
            If SelectSQL("SELECT Max(MaVattu) AS F1 FROM VTDauNam") = 0 Then
                ExecuteSQL5 "INSERT INTO VTDauNam (MaSo, MaSoKho, MaTaiKhoan, MaVattu, Luong_0, Tien_0) SELECT MaSo, MaSoKho, MaTaiKhoan, MaVattu, Luong_0, Tien_0 FROM TonKho WHERE Luong_0<>0 OR Tien_0<>0"
            End If
      End If
      
      Set rs_ktra = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSoKho,MaTaiKhoan,MaVattu FROM TonKho ORDER BY MaSoKho,MaTaiKhoan,MaVattu", dbOpenSnapshot)
      Do While Not rs_ktra.EOF
            If rs_ktra!MaTaiKhoan = m And rs_ktra!MaVattu = n And rs_ktra!MaSoKho = mk Then
                ExecuteSQL5 "DELETE * FROM TonKho WHERE MaSoKho=" + CStr(mk) + " AND MaTaiKhoan=" + CStr(m) + " AND MaVattu=" + CStr(n) + " AND Luong_0=0 AND Tien_0=0"
            Else
                n = rs_ktra!MaVattu
                mk = rs_ktra!MaSoKho
                If m <> rs_ktra!MaTaiKhoan Then
                    m1 = SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE TKCha0=" + CStr(rs_ktra!MaTaiKhoan))
                    If m1 > 0 Then
                        ExecuteSQL5 "UPDATE TonKho SET MaTaiKhoan=" + CStr(m1) + " WHERE MaTaiKhoan=" + CStr(rs_ktra!MaTaiKhoan)
                        m = m1
                    Else
                        m = rs_ktra!MaTaiKhoan
                    End If
                End If
            End If
            rs_ktra.MoveNext
      Loop
            
      ' KiÓm tra sè ph¸t sinh vËt t­
      HienThongBao "KiÓm tra sè nhËp xuÊt tån ...", 1
      ExecuteSQL5 "DELETE DISTINCTROW TonKho.* FROM Vattu RIGHT JOIN TonKho ON Vattu.MaSo = TonKho.MaVatTu WHERE ((Vattu.MaSo Is Null))"
      sql = "UPDATE TonKho SET MaSoKho=MaSoKho"
      For i = 1 To 12
            sql = sql + ",Luong_Nhap_" + CStr(i) + "=0,Luong_Xuat_" + CStr(i) + "=0,Tien_Nhap_" + CStr(i) + "=0,Tien_Xuat_" + CStr(i) + "=0"
            If pGiaUSD > 0 Then sql = sql + ",USDTien_Nhap_" + CStr(i) + "=0,USDTien_Xuat_" + CStr(i) + "=0"
      Next
      ExecuteSQL5 sql
      
      sql = "SELECT MaTkNo,IIF(MaLoai=1,MaKho,MaNguon) AS MK,ChungTu.MaVattu"
      For i = 1 To 12
            st = CStr(i)
            sql = sql + ", SUM(iif(ThangCT = " + CStr(CThangFR(i)) + ", SoPS,0)) As No_" + st _
              + ", SUM(iif(ThangCT = " + CStr(CThangFR(i)) + ", SoPS2No,0)) As NTNo_" + st
            If pGiaUSD > 0 Then sql = sql + ", SUM(iif(ThangCT = " + CStr(CThangFR(i)) + ", PSUSD,0)) As USDNo_" + st
      Next

      sql = sql + " FROM " + ChungTu2TKNC(-1) + " WHERE (MaLoai=1 OR MaLoai=4) AND (TK_ID=" + CStr(TKVT_ID) + " AND ChungTu.MaVattu > 0 AND MaKho>0) GROUP BY MaTkNo,IIF(MaLoai=1,MaKho,MaNguon),ChungTu.MaVattu"
      Set rs_taikhoan = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
      Do While Not rs_taikhoan.EOF
            sql = "UPDATE TonKho SET MaSoKho=MaSoKho"
            For i = 1 To 12
                  st = CStr(i)
                  sql = sql + ", Tien_Nhap_" + st + " = " + DoiDau(rs_taikhoan.Fields("No_" + st)) _
                        + ", Luong_Nhap_" + st + " = " + DoiDau(rs_taikhoan.Fields("NTNo_" + st))
                  If pGiaUSD > 0 Then sql = sql + ", USDTien_Nhap_" + st + " = " + DoiDau(rs_taikhoan.Fields("USDNo_" + st))
            Next
K1:
            ExecuteSQL5 sql + " WHERE MaSoKho = " + CStr(rs_taikhoan!mk) + " AND MaTaiKhoan=" + CStr(rs_taikhoan!MaTkNo) + " AND MaVattu=" + CStr(rs_taikhoan!MaVattu)
            If DBKetoan.RecordsAffected = 0 Then
                If TinhTonKho(rs_taikhoan!mk, rs_taikhoan!MaTkNo, rs_taikhoan!MaVattu, 12, -1, 0, 0, 0) = 0 Then GoTo K1
            End If
            rs_taikhoan.MoveNext
      Loop
      
      sql = "SELECT MaTkCo,MaKho,ChungTu.MaVattu"
      For i = 1 To 12
            st = CStr(i)
            sql = sql + ", SUM(iif(ThangCT = " + CStr(CThangFR(i)) + ", SoPS,0)) As No_" + st _
              + ", SUM(iif(ThangCT = " + CStr(CThangFR(i)) + ", SoPS2Co,0)) As NTNo_" + st
            If pGiaUSD > 0 Then sql = sql + ", SUM(iif(ThangCT = " + CStr(CThangFR(i)) + ", PSUSD,0)) As USDNo_" + st
      Next

      sql = sql + " FROM " + ChungTu2TKNC(1) + " WHERE (MaLoai=2 OR MaLoai=4) AND (TK_ID=" + CStr(TKVT_ID) + " AND ChungTu.MaVattu > 0 AND MaKho>0) GROUP BY MaTkCo,MaKho,ChungTu.MaVattu"
      Set rs_taikhoan = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
      Do While Not rs_taikhoan.EOF
            sql = "UPDATE TonKho SET MaSoKho=MaSoKho"
            For i = 1 To 12
                  st = CStr(i)
                  sql = sql + ", Tien_Xuat_" + st + " = " + DoiDau(rs_taikhoan.Fields("No_" + st)) _
                        + ", Luong_Xuat_" + st + " = " + DoiDau(rs_taikhoan.Fields("NTNo_" + st))
                  If pGiaUSD > 0 Then sql = sql + ", USDTien_Xuat_" + st + " = " + DoiDau(rs_taikhoan.Fields("USDNo_" + st))
            Next
k2:
            ExecuteSQL5 sql + " WHERE MaSoKho = " + CStr(rs_taikhoan!MaKho) + " AND MaTaiKhoan=" + CStr(rs_taikhoan!MaTkCo) + " AND MaVattu=" + CStr(rs_taikhoan!MaVattu)
            If DBKetoan.RecordsAffected = 0 Then
                If TinhTonKho(rs_taikhoan!MaKho, rs_taikhoan!MaTkCo, rs_taikhoan!MaVattu, 12, -1, 0, 0, 0) = 0 Then GoTo k2
            End If
            rs_taikhoan.MoveNext
      Loop
      
                        
      HienThongBao "KiÓm tra sè tån ...", 1
      ExecuteSQL5 "UPDATE TonKho SET Luong_0 = Fix(iif(Luong_0>0,0.5,-0.5) + Luong_0 * " + CStr(Mask_N) + ") / " + CStr(Mask_N) + ", Tien_0 = " + IIf(pTien = 0, "Fix(iif(Tien_0>0,0.5,-0.5)  + Tien_0)", "Fix(iif(Tien_0>0,0.5,-0.5)  + " + CStr(Mask_N) + "*Tien_0)/" + CStr(Mask_N))
      sql = "UPDATE TonKho SET MaVattu = MaVattu"
      For i = 1 To 12
            st = "Luong_0"
            st2 = "Tien_0"
            st4 = "USDTien_0"
            For j = 1 To i
                  st3 = CStr(j)
                  st = st + " + Luong_Nhap_" + st3 + " - Luong_Xuat_" + st3
                  st2 = st2 + " + Tien_Nhap_" + st3 + " - Tien_Xuat_" + st3
                  st4 = st4 + " + USDTien_Nhap_" + st3 + " - USDTien_Xuat_" + st3
            Next
           sql = sql + ", Luong_" + CStr(i) + " = Fix(iif(" + st + ">0,0.5,-0.5)  + (" + st + ") * " + CStr(Mask_N) + ")/" + CStr(Mask_N) + ", Tien_" + CStr(i) + " = " + st2 + IIf(pGiaUSD > 0, ", USDTien_" + CStr(i) + " = " + st4, "")
      Next
      ExecuteSQL5 sql
      
      sql = "DELETE FROM TonKho WHERE Luong_0=0 And Tien_0=0" + IIf(pGiaUSD > 0, " AND USDTien_0=0", "")
      For i = 1 To 12
          sql = sql + " And Luong_Nhap_" + CStr(i) + "=0 And Luong_Xuat_" + CStr(i) + "=0 And Tien_Nhap_" + CStr(i) + "=0" + IIf(pGiaUSD > 0, " AND USDTien_Nhap_" + CStr(i) + "=0", "")
      Next
      ExecuteSQL5 sql
      
      'If ktraxuat = 0 Then ktraxuatvattu
      
      SoDuTKVT
      
      rs_taikhoan.Close
      Set rs_taikhoan = Nothing
      Set rs_ktra = Nothing
End Sub
'======================================================================================
' Sub KiemTraDuLieu : KiÓm tra hÖ thèng tµi kho¶n
'======================================================================================
Public Sub KiemTraTaiKhoan(Optional ktracn As Integer = 0)
      Dim rs_taikhoan As Recordset, m As Long, n As Long, m1 As Long
      Dim rs_ktra As Recordset, i As Integer, sql As String, st As String, st2 As String, j As Integer, st3 As String
        
      SetDefaultValue
      ExecuteSQL5 "UPDATE " + ChungTu2TKNC(-1) + " SET MaTKTCNo=MaTC"
      ExecuteSQL5 "UPDATE " + ChungTu2TKNC(1) + " SET MaTKTCCo=MaTC"
      ExecuteSQL5 "UPDATE HoaDon SET Loai=1 WHERE Loai=0"
      ExecuteSQL5 "UPDATE KhachHang SET SoHieu='#0000000000' WHERE SoHieu='#'", False
      ExecuteSQL5 "UPDATE TaiSan SET NamKH=0 WHERE IsNull(NamKH)"
      
      HienThongBao "KiÓm tra hÖ thèng tµi kho¶n ...", 1
      
      ' KiÓm tra cÊu tróc tµi kho¶n vµ sè ph¸t sinh
      Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT MaSo, TKCon, TKCha0 FROM HethongTK ORDER BY SoHieu", dbOpenSnapshot, dbForwardOnly)
      Do While Not rs_taikhoan.EOF
            st2 = CStr(rs_taikhoan!MaSo)
            If rs_taikhoan!tkcon > 0 Then
                Set rs_ktra = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTk WHERE TkCha0 = " + st2, dbOpenSnapshot)
                If rs_ktra.recordCount = 0 Then ExecuteSQL5 "UPDATE HethongTk SET TkCon=0 WHERE MaSo=" + st2
            Else
                Set rs_ktra = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTk WHERE TkCha0 = " + st2, dbOpenSnapshot)
                If rs_ktra.recordCount > 0 Then ExecuteSQL5 "UPDATE HethongTk SET TkCon=1 WHERE MaSo=" + st2
            End If
            rs_taikhoan.MoveNext
      Loop
      rs_ktra.Close
    
      HienThongBao "KiÓm tra sè ph¸t sinh ...", 1
      sql = "SELECT ChungTu.MaTkNo"
      For i = 1 To 12
            st = CStr(i)
            sql = sql + ", SUM(iif(ThangCT = " + CStr(CThangFR(i)) + ", SoPS,0)) As No_" + st _
              + ", SUM(iif(ThangCT = " + CStr(CThangFR(i)) + ", SoPS2No,0)) As NTNo_" + st
      Next

      SetSQL "QCdt", sql + " FROM ChungTu WHERE (chungtu.MaTkNo > 0 AND (MaLoai<>4 OR MaTKNo<>MaTKCo)) GROUP BY MaTkNo"
      Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT DISTINCTROW QCdt.*, HethongTK.MaSo As MaSoTK " _
        & " FROM QCdt RIGHT JOIN HethongTK ON QCdt.MaTkNo = HethongTK.MaSo" _
        & " WHERE HethongTK.TkCon = 0", dbOpenSnapshot, dbForwardOnly)
      Do While Not rs_taikhoan.EOF
            sql = "UPDATE HethongTK SET MaTC = MaTC"
            For i = 1 To 12
                  st = CStr(i)
                  sql = sql + ", No_" + st + " = " + DoiDau(IIf(IsNull(rs_taikhoan.Fields("No_" + st)), 0, rs_taikhoan.Fields("No_" + st))) _
                        + ", No_" + st + "_NT = " + DoiDau(IIf(IsNull(rs_taikhoan.Fields("NTNo_" + st)), 0, rs_taikhoan.Fields("NTNo_" + st)))
            Next
            ExecuteSQL5 sql + " WHERE MaSo = " + CStr(rs_taikhoan!MaSoTK)
            rs_taikhoan.MoveNext
      Loop
      
      sql = "SELECT ChungTu.MaTkCo"
      For i = 1 To 12
            st = CStr(i)
            sql = sql + ", SUM(iif(ThangCT = " + CStr(CThangFR(i)) + ", SoPS,0)) As Co_" + st _
              + ", SUM(iif(ThangCT = " + CStr(CThangFR(i)) + ", SoPS2Co,0)) As NTCo_" + st
      Next
      
      SetSQL "QCdt", sql + " FROM ChungTu WHERE (chungtu.MaTkCo > 0 AND (MaLoai<>4 OR MaTKNo<>MaTKCo)) GROUP BY MaTkCo"
      Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT DISTINCTROW QCdt.*, HethongTK.MaSo As MaSoTK " _
        & " FROM QCdt RIGHT JOIN HethongTK ON QCdt.MaTkCo = HethongTK.MaSo" _
        & " WHERE HethongTK.TkCon = 0", dbOpenSnapshot, dbForwardOnly)
      Do While Not rs_taikhoan.EOF
            sql = "UPDATE HethongTK SET MaTC = MaTC"
            For i = 1 To 12
                  st = CStr(i)
                  sql = sql + ", Co_" + st + " = " + DoiDau(IIf(IsNull(rs_taikhoan.Fields("Co_" + st)), 0, rs_taikhoan.Fields("Co_" + st))) _
                        + ", Co_" + st + "_NT = " + DoiDau(IIf(IsNull(rs_taikhoan.Fields("NTCo_" + st)), 0, rs_taikhoan.Fields("NTCo_" + st)))
            Next
            ExecuteSQL5 sql + " WHERE MaSo = " + CStr(rs_taikhoan!MaSoTK)
            rs_taikhoan.MoveNext
      Loop
        
      ' Kiem tra chi tiet hoa don
      m = SoHieu2MaSo("#", "KhachHang")
      If m = 0 Then
            If SelectSQL("SELECT MaSo AS F1 FROM PhanLoaiKhachHang WHERE SoHieu='#'") = 0 Then
                m = Lng_MaxValue("MaSo", "PhanLoaiKhachHang") + 1
                ExecuteSQL5 "INSERT INTO PhanLoaiKhachHang (MaSo,SoHieu) VALUES (" + CStr(m) + ",'#')"
            End If
            m = Lng_MaxValue("MaSo", "KhachHang") + 1
            ExecuteSQL5 "INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu) VALUES (" + CStr(m) + "," + CStr(SoHieu2MaSo("#", "PhanLoaiKhachHang")) + ",'#')"
      End If

      ExecuteSQL5 "INSERT INTO HoaDon (MaSo,Loai,MaKhachHang) SELECT ChungTu.MaSo,-1," + CStr(m) + " FROM (" + ChungTu2TKNC(-1) + ") LEFT JOIN HoaDon ON ChungTu.MaSo=HoaDon.MaSo WHERE IsNull(HoaDon.MaSo) AND TK_ID=" + CStr(GTGTKT_ID)
      ExecuteSQL5 "INSERT INTO HoaDon (MaSo,Loai,MaKhachHang) SELECT ChungTu.MaSo,1," + CStr(m) + " FROM (" + ChungTu2TKNC(1) + ") LEFT JOIN HoaDon ON ChungTu.MaSo=HoaDon.MaSo WHERE IsNull(HoaDon.MaSo) AND TK_ID=" + CStr(GTGTPN_ID)

      HienThongBao "KiÓm tra sè d­ ...", 1
      sql = "UPDATE HethongTK SET Cap = Cap"
      For i = 1 To 12
            st = "DuNo_0 - DuCo_0"
            For j = 1 To i
                  st3 = CStr(j)
                  st = st + " + No_" + st3 + " - Co_" + st3
            Next
            st3 = CStr(i)
            sql = sql + ", DuNo_" + st3 + " = IIF(" + st + " > 0, " + st + ", 0), DuCo_" + st3 + " =  IIF(" + st + " < 0, -(" + st + "), 0)"
      Next
      ExecuteSQL5 sql + " WHERE TKCon = 0"

      sql = "UPDATE HethongTK SET Cap = Cap"
      For i = 1 To 12
            st2 = "DuNT_0"
            For j = 1 To i
                  st3 = CStr(j)
                  st2 = st2 + " + IIF(Kieu < 0, No_" + st3 + "_NT - Co_" + st3 + "_NT, Co_" + st3 + "_NT - No_" + st3 + "_NT)"
            Next
            sql = sql + ", DuNT_" + CStr(i) + " = " + st2
      Next
      ExecuteSQL5 sql + " WHERE TKCon = 0 AND MaNT <> 0"
      
      HienThongBao "KiÓm tra c«ng nî ...", 1
      ExecuteSQL5 "UPDATE " + ChungTu2TKNC(-1) + " SET MaKH=0 WHERE MaTKCo=0 AND TK_ID<>" + CStr(TKCNKH_ID) + " AND TK_ID<>" + CStr(TKCNPT_ID)
      ExecuteSQL5 "UPDATE " + ChungTu2TKNC(1) + " SET MaKH=0 WHERE MaTKNo=0 AND TK_ID<>" + CStr(TKCNKH_ID) + " AND TK_ID<>" + CStr(TKCNPT_ID)
      ExecuteSQL5 "UPDATE " + ChungTu2TKNC(0) + " SET MaKH=0 WHERE MaTKCo=0 AND HethongTK.TK_ID<>" + CStr(TKCNKH_ID) + " AND HethongTK.TK_ID<>" + CStr(TKCNPT_ID) + " AND TK.TK_ID<>" + CStr(TKCNKH_ID) + " AND TK.TK_ID<>" + CStr(TKCNPT_ID)
      
      ExecuteSQL5 "DELETE SoDuKhachHang.* FROM SoDuKhachHang INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo WHERE LEFT(SoHieu,1)='#'"
      
      Set rs_ktra = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaTaiKhoan,MaKhachHang FROM SoDuKhachHang ORDER BY MaTaiKhoan,MaKhachHang", dbOpenSnapshot)
      Do While Not rs_ktra.EOF
            If rs_ktra!MaTaiKhoan = m And rs_ktra!MaKhachHang = n Then
                ExecuteSQL5 "DELETE * FROM SoDuKhachHang WHERE MaTaiKhoan=" + CStr(m) + " AND MaKhachHang=" + CStr(n) + " AND DuNo_0=0 AND DuCo_0=0"
            Else
                n = rs_ktra!MaKhachHang
                If m <> rs_ktra!MaTaiKhoan Then
                    m1 = SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE TKCha0=" + CStr(rs_ktra!MaTaiKhoan))
                    If m1 > 0 Then
                        ExecuteSQL5 "UPDATE SoDuKhachHang SET MaTaiKhoan=" + CStr(m1) + " WHERE MaTaiKhoan=" + CStr(rs_ktra!MaTaiKhoan)
                        m = m1
                    Else
                        m = rs_ktra!MaTaiKhoan
                    End If
                End If
            End If
            rs_ktra.MoveNext
      Loop
      
      If BangDaCo("CNDauNam") Then
            Set rs_ktra = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo, MaTaiKhoan,MaKhachHang FROM CNDauNam ORDER BY MaTaiKhoan,MaKhachHang", dbOpenSnapshot)
            Do While Not rs_ktra.EOF
                  If rs_ktra!MaTaiKhoan = m And rs_ktra!MaKhachHang = n Then
                      ExecuteSQL5 "DELETE * FROM CNDauNam WHERE MaSo=" + CStr(rs_ktra!MaSo)
                  Else
                      n = rs_ktra!MaKhachHang
                      If m <> rs_ktra!MaTaiKhoan Then
                          m1 = SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE TKCha0=" + CStr(rs_ktra!MaTaiKhoan))
                          If m1 > 0 Then
                              ExecuteSQL5 "UPDATE CNDauNam SET MaTaiKhoan=" + CStr(m1) + " WHERE MaTaiKhoan=" + CStr(rs_ktra!MaTaiKhoan)
                              m = m1
                          Else
                              m = rs_ktra!MaTaiKhoan
                          End If
                      End If
                  End If
                  rs_ktra.MoveNext
            Loop
      End If
      
      sql = "UPDATE SoDuKhachHang SET MaKhachHang = MaKhachHang"
      For i = 1 To 12
            sql = sql + ",No_" + CStr(i) + "=0,Co_" + CStr(i) + "=0,No_" + CStr(i) + "_NT=0,Co_" + CStr(i) + "_NT=0"
      Next
      ExecuteSQL5 sql
      
      Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT MaTKNo,MaKH,ThangCT,Sum(SoPS) AS TPS,Sum(SoPS2No) AS TPS2 FROM " + ChungTu2TKNC(-1) + " WHERE MaTKNo>0 AND MaKH >0 AND (TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ") GROUP BY MaTKNo,MaKH,ThangCT", dbOpenSnapshot)
      Do While Not rs_taikhoan.EOF
a:
            ExecuteSQL5 "UPDATE SoDuKhachHang SET No_" + CStr(CThangDB(rs_taikhoan!ThangCT)) + "=" + DoiDau(rs_taikhoan!tps) + ",No_" + CStr(CThangDB(rs_taikhoan!ThangCT)) + "_NT=" + DoiDau(rs_taikhoan!tps2) + " WHERE MaTaiKhoan=" + CStr(rs_taikhoan!MaTkNo) + " AND MaKhachHang=" + CStr(rs_taikhoan!makh)
            If DBKetoan.RecordsAffected = 0 Then
                If ExecuteSQL5("INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(rs_taikhoan!MaTkNo) + "," + CStr(rs_taikhoan!makh) + ")", False) <> 0 Then GoTo c
                GoTo a
            End If
c:
            rs_taikhoan.MoveNext
      Loop
      
      Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT MaTKCo,MaKHC,ThangCT,Sum(SoPS) AS TPS,Sum(SoPS2Co) AS TPS2 FROM " + ChungTu2TKNC(1) + " WHERE MaTKCo>0 AND MaKHC>0 AND (TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ") GROUP BY MaTKCo,MaKHC,ThangCT", dbOpenSnapshot)
      Do While Not rs_taikhoan.EOF
B:
            ExecuteSQL5 "UPDATE SoDuKhachHang SET Co_" + CStr(CThangDB(rs_taikhoan!ThangCT)) + "=" + DoiDau(rs_taikhoan!tps) + ",Co_" + CStr(CThangDB(rs_taikhoan!ThangCT)) + "_NT=" + DoiDau(rs_taikhoan!tps2) + " WHERE MaTaiKhoan=" + CStr(rs_taikhoan!MaTkCo) + " AND MaKhachHang=" + CStr(rs_taikhoan!MaKHC)
            If DBKetoan.RecordsAffected = 0 Then
                If ExecuteSQL5("INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(rs_taikhoan!MaTkCo) + "," + CStr(rs_taikhoan!MaKHC) + ")", False) <> 0 Then GoTo d
                GoTo B
            End If
d:
            rs_taikhoan.MoveNext
      Loop
      
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
      ExecuteSQL5 sql
      
      sql = "UPDATE SoDuKhachHang INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo SET "
      For i = 1 To 12
            ExecuteSQL5 sql + "DuNT_" + CStr(i) + "=ABS(DuNT_" + CStr(i - 1) + "+IIF(DuNo_" + CStr(i - 1) + "-DuCo_" + CStr(i - 1) + ">=0,No_" + CStr(i) + "_NT-Co_" + CStr(i) + "_NT,Co_" + CStr(i) + "_NT-No_" + CStr(i) + "_NT)) WHERE KhachHang.MaNT<>0"
      Next
      SoDuTKCN
     
     If pCongNoHD > 0 And ktracn = 0 Then
        KtraCongNo
        ExecuteSQL5 "UPDATE CNDauNam SET SoXuat=0"
        
        sql = "SELECT ChungTu.MaSo,MaTKNo,MaKH,SoPS,CT_ID FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE ABS(CT_ID)>2000000000 AND TK_ID=" + CStr(TKCNPT_ID)
        Set rs_ktra = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rs_ktra.EOF
           m = Abs(rs_ktra!CT_ID) - 2000000000
           If SelectSQL("SELECT MaSo AS F1 FROM CNDauNam WHERE MaSo=" + CStr(m)) = 0 Then
               m = SelectSQL("SELECT MaSo AS F1 FROM CNDauNam WHERE MaTaiKhoan=" + CStr(rs_ktra!MaTkNo) + " AND MaKhachHang=" + CStr(rs_ktra!makh) + " ORDER BY HanTT DESC")
               If m > 0 Then
                   ExecuteSQL5 "UPDATE CNDauNam SET SoXuat=SoXuat+" + CStr(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
                   ExecuteSQL5 "UPDATE ChungTu SET CT_ID=" + CStr(-(2000000000 - m)) + " WHERE MaSo=" + CStr(rs_ktra!MaSo)
               End If
           Else
               ExecuteSQL5 "UPDATE CNDauNam SET SoXuat=SoXuat+" + CStr(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
           End If
           rs_ktra.MoveNext
        Loop
             
        sql = "SELECT ChungTu.MaSo,MaTKCo,MaKHC,SoPS,CT_ID FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE ABS(CT_ID)>2000000000 AND TK_ID=" + CStr(TKCNKH_ID)
        Set rs_ktra = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rs_ktra.EOF
           m = Abs(rs_ktra!CT_ID) - 2000000000
           If SelectSQL("SELECT MaSo AS F1 FROM CNDauNam WHERE MaSo=" + CStr(m)) = 0 Then
               m = SelectSQL("SELECT MaSo AS F1 FROM CNDauNam WHERE MaTaiKhoan=" + CStr(rs_ktra!MaTkCo) + " AND MaKhachHang=" + CStr(rs_ktra!MaKHC) + " ORDER BY HanTT DESC")
               If m > 0 Then
                   ExecuteSQL5 "UPDATE CNDauNam SET SoXuat=SoXuat+" + DoiDau(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
                   ExecuteSQL5 "UPDATE ChungTu SET CT_ID=" + CStr(-(2000000000 - m)) + " WHERE MaSo=" + CStr(rs_ktra!MaSo)
               End If
           Else
               ExecuteSQL5 "UPDATE CNDauNam SET SoXuat=SoXuat+" + CStr(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
           End If
           rs_ktra.MoveNext
        Loop
        
        ExecuteSQL5 "UPDATE ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo SET ChungTu.SoXuat=0 WHERE TK_ID=" + CStr(TKCNKH_ID)
        sql = "SELECT ChungTu.MaSo,SoPS,CT_ID FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE ABS(CT_ID)>0 AND ABS(CT_ID)<2000000000 AND TK_ID=" + CStr(TKCNKH_ID)
        Set rs_ktra = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rs_ktra.EOF
           m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNKH_ID) + " AND MaCT=" + CStr(Abs(rs_ktra!CT_ID)))
           ExecuteSQL5 "UPDATE ChungTu SET SoXuat=SoXuat+" + DoiDau(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
           rs_ktra.MoveNext
        Loop
        
        ExecuteSQL5 "UPDATE ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo SET ChungTu.SoXuat=0 WHERE TK_ID=" + CStr(TKCNPT_ID)
        sql = "SELECT ChungTu.MaSo,SoPS,CT_ID FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE ABS(CT_ID)>0 AND ABS(CT_ID)<2000000000 AND TK_ID=" + CStr(TKCNPT_ID)
        Set rs_ktra = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not rs_ktra.EOF
           m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNPT_ID) + " AND MaCT=" + CStr(Abs(rs_ktra!CT_ID)))
           ExecuteSQL5 "UPDATE ChungTu SET SoXuat=SoXuat+" + DoiDau(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
           rs_ktra.MoveNext
        Loop
        rs_ktra.Close
        ktracongno2004
      End If
      HienThongBao "Céng dån sè ph¸t sinh ...", 1
      LapCanDoiThu
      
      HienThongBao "", 1
      
      rs_taikhoan.Close
      Set rs_taikhoan = Nothing
      Set rs_ktra = Nothing
End Sub
'======================================================================================
' Thñ tôc xãa ph¸t sinh cña th¸ng
'======================================================================================
Public Sub XoaPSThang(thang As Integer)
    Dim st As String
    Dim rs_ktra As Recordset
    Dim rs_ktra2 As Recordset
    Dim rs_ktradt As Recordset
    st = CStr(thang)

    Dim Query As String
    Dim id As String
    Query = "SELECT * FROM tbimport"
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    While Not rs_ktra.EOF

        Query = "SELECT * FROM HoaDon WHERE SoHD = '" & rs_ktra!SHDon & "' AND KyHieu = '" & rs_ktra!KHHDon & "';"
        Set rs_ktra2 = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
        If Not rs_ktra2.EOF Then
            id = Month(rs_ktra2!NgayPH)
            'Kiem tra neu cung thang thi xoa ca1
            If id = thang Then
                'tim thang con truoc
                'Query = "SELECT * FROM tbimportdetail WHERE ParentId = '" & rs_ktra!id & "';"
                ExecuteSQL5 "DELETE FROM tbimportdetail WHERE ParentId = '" & rs_ktra!id & "';"
                ExecuteSQL5 "DELETE FROM tbimport WHERE ID = " & rs_ktra!id & ";"
                rs_ktra2.MoveNext
            End If
        End If
        rs_ktra.MoveNext
    Wend
    
    ExecuteSQL5 "DELETE FROM Chungtu WHERE (MaLoai <> 5 AND MaLoai <> 6 AND MaLoai <> 7) AND ThangCT = " + st
    ExecuteSQL5 "DELETE ChungTuLQ.* FROM ChungTuLQ LEFT JOIN ChungTu ON ChungTuLQ.MaCT=ChungTu.MaCT WHERE IsNull(ChungTu.MaCT)"
    ExecuteSQL5 "DELETE CTTaiSan.* FROM CTTaiSan LEFT JOIN ChungTu ON CTTaiSan.MaCTKT=ChungTu.MaCT WHERE IsNull(ChungTu.MaCT) AND CTTaiSan.Thang>0"
    ExecuteSQL5 "DELETE HoaDon.* FROM HoaDon LEFT JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo WHERE IsNull(ChungTu.MaSo)"
    ExecuteSQL5 "DELETE ThongSo.* FROM ThongSo LEFT JOIN CTTaiSan ON ThongSo.MaTS= CTTaiSan.MaTS WHERE IsNull(CTTaiSan.MaTS)"
    ExecuteSQL5 "DELETE TaiSan.* FROM TaiSan LEFT JOIN CTTaiSan ON TaiSan.MaSo= CTTaiSan.MaTS WHERE IsNull(CTTaiSan.MaTS)"

    st = CStr(CThangDB(thang))
    ExecuteSQL5 "UPDATE HethongTK SET No_" + st + " = 0, Co_" + st + " = 0, No_" + st + "_NT = 0, Co_" + st + "_NT = 0"
    ExecuteSQL5 "UPDATE TonKho SET Luong_Nhap_" + st + " = 0, Tien_Nhap_" + st + " = 0, Luong_Xuat_" + st + " = 0, Tien_Xuat_" + st + " = 0"
    ExecuteSQL5 "UPDATE SoDuKhachHang SET No_" + st + " = 0, Co_" + st + " = 0, No_" + st + "_NT = 0, Co_" + st + "_NT = 0"

    XoaPSTS thang
    KiemTraTaiKhoan
    If STDetail Then KiemTraVatTu
End Sub
'======================================================================================
' Thñ tôc céng dån sè d­ trong b¶ng c©n ®èi thö
'======================================================================================
Private Sub LapCanDoiThu()
    Dim rs_taikhoan As Recordset, rs_sub As Recordset, sql As String, i As Integer, sql2 As String, st As String, st2 As String
           
    sql = "SELECT DISTINCTROW Sum(HethongTK.DuNo_0) As DN_0, Sum(HethongTK.DuCo_0) As DC_0"
    For i = 1 To 12
            st2 = CStr(i)
            sql = sql + ", Sum(HeThongTK.No_" + st2 + ") As N_" + st2 + ", Sum(HeThongTK.Co_" + st2 + ") As C_" + st2 _
                    + ", Sum(HeThongTK.DuNo_" + st2 + ") As DN_" + st2 + ", Sum(HeThongTK.DuCo_" + st2 + ") As DC_" + st2
    Next
    Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT * FROM HethongTK WHERE Cap>0 AND TkCon>0 ORDER BY SoHieu DESC", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_taikhoan.EOF
        Set rs_sub = DBKetoan.OpenRecordset(sql + " FROM HethongTK WHERE HethongTK.TkCha0 = " + CStr(rs_taikhoan!MaSo), dbOpenSnapshot)
        sql2 = "UPDATE HethongTK SET HethongTK.DuNo_0 = " + DoiDau(rs_sub!DN_0) + ", HethongTK.DuCo_0 = " + DoiDau(Format(rs_sub!DC_0, Mask_2))
        For i = 1 To 12
            st2 = CStr(i)
            sql2 = sql2 + ", HeThongTK.No_" + st2 + " = " + DoiDau(rs_sub.Fields("N_" + st2)) + ", HeThongTK.Co_" + st2 + " = " + DoiDau(rs_sub.Fields("C_" + st2)) _
                    + ", HeThongTK.DuNo_" + st2 + " = " + DoiDau(rs_sub.Fields("DN_" + st2)) + ", HeThongTK.DuCo_" + st2 + " = " + DoiDau(rs_sub.Fields("DC_" + st2))
        Next
        ExecuteSQL5 sql2 + " WHERE MaSo = " + CStr(rs_taikhoan!MaSo)
        rs_taikhoan.MoveNext
    Loop
    
    sql = "UPDATE HethongTK SET DuNo_0=IIF(DuNo_0>=DuCo_0,DuNo_0-DuCo_0,0),DuCo_0=IIF(DuNo_0<DuCo_0,DuCo_0-DuNo_0,0)"
    For i = 1 To 12
        sql = sql + ",DuNo_" + CStr(i) + "=IIF(DuNo_" + CStr(i) + ">=DuCo_" + CStr(i) + ",DuNo_" + CStr(i) + "-DuCo_" + CStr(i) + ",0),DuCo_" + CStr(i) + "=IIF(DuNo_" + CStr(i) + "<DuCo_" + CStr(i) + ",DuCo_" + CStr(i) + "-DuNo_" + CStr(i) + ",0)"
    Next
    ExecuteSQL5 sql + " WHERE TK_ID2<>" + CStr(TKLT_ID)
    
    Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT * FROM HethongTK WHERE Cap=0 AND TKCon>0", dbOpenSnapshot, dbForwardOnly)
    sql = "SELECT DISTINCTROW Sum(HethongTK.DuNo_0) As DN_0, Sum(HethongTK.DuCo_0) As DC_0"
    For i = 1 To 12
            st2 = CStr(i)
            sql = sql + ", Sum(HeThongTK.No_" + st2 + ") As N_" + st2 + ", Sum(HeThongTK.Co_" + st2 + ") As C_" + st2 _
                    + ", Sum(HeThongTK.DuNo_" + st2 + ") As DN_" + st2 + ", Sum(HeThongTK.DuCo_" + st2 + ") As DC_" + st2
    Next
    
    Do While Not rs_taikhoan.EOF
        Set rs_sub = DBKetoan.OpenRecordset(sql + " FROM HethongTK WHERE ((HeThongTK.TKCha0=" + CStr(rs_taikhoan!MaSo) + "))", dbOpenSnapshot)
        sql2 = "UPDATE HethongTK SET HethongTK.DuNo_0 = " + DoiDau(rs_sub!DN_0) + ", HethongTK.DuCo_0 = " + DoiDau(Format(rs_sub!DC_0, Mask_2))
        For i = 1 To 12
            st2 = CStr(i)
            sql2 = sql2 + ", HeThongTK.No_" + st2 + " = " + DoiDau(rs_sub.Fields("N_" + st2)) + ", HeThongTK.Co_" + st2 + " = " + DoiDau(rs_sub.Fields("C_" + st2)) _
                    + ", HeThongTK.DuNo_" + st2 + " = " + DoiDau(rs_sub.Fields("DN_" + st2)) + ", HeThongTK.DuCo_" + st2 + " = " + DoiDau(rs_sub.Fields("DC_" + st2))
        Next
        ExecuteSQL5 sql2 + " WHERE MaSo = " + CStr(rs_taikhoan!MaSo)
        rs_taikhoan.MoveNext
    Loop
    
    sql = "UPDATE HethongTK SET DuNo_0=IIF(DuNo_0>=DuCo_0,DuNo_0-DuCo_0,0),DuCo_0=IIF(DuNo_0<DuCo_0,DuCo_0-DuNo_0,0)"
    For i = 1 To 12
        sql = sql + ",DuNo_" + CStr(i) + "=IIF(DuNo_" + CStr(i) + ">=DuCo_" + CStr(i) + ",DuNo_" + CStr(i) + "-DuCo_" + CStr(i) + ",0),DuCo_" + CStr(i) + "=IIF(DuNo_" + CStr(i) + "<DuCo_" + CStr(i) + ",DuCo_" + CStr(i) + "-DuNo_" + CStr(i) + ",0)"
    Next
    ExecuteSQL5 sql + " WHERE Cap=0 AND TK_ID2<>" + CStr(TKLT_ID)
    
    rs_taikhoan.Close
    Set rs_taikhoan = Nothing
    Set rs_sub = Nothing
End Sub
'======================================================================================
' Thñ tôc thùc hiÖn chuyÓn d÷ liÖu sang n¨m míi
'======================================================================================
Public Sub ChuyenNamMoi()
    Dim i As Integer, sql As String, st As String, NamCu As String, dpath As String, tail As String, ms As Long, sqln As String, sqlc As String
    
    For i = Len(pDataPath) - 1 To 1 Step -1
        If Mid(pDataPath, i, 1) = "\" Then Exit For
    Next
    dpath = Left(pDataPath, i)
    
    NamCu = CStr(pNamTC)
    CloseUp
    
    If Len(Dir(dpath + "KT" + NamCu + ".MDB")) > 0 Then
        For i = 1 To 100
            tail = Chr(i + 64)
            If Len(Dir(dpath + "KT" + NamCu + tail + ".MDB")) = 0 Then Exit For
        Next
    Else
        tail = ""
    End If
    
    DBEngine.CompactDatabase pDataPath, dpath + "KT" + NamCu + tail + ".MDB", , , ";pwd=" + pPSW
    SaveSetting IniPath, "LastYear", CStr(pNamTC), dpath + "KT" + NamCu + tail + ".MDB"
    OpenDB pDataPath
    
    If OutCost <> 0 Then
        ms = Lng_MaxValue("MaSo", "VTDauNam")
        
        ExecuteSQL5 "INSERT INTO VTDauNam (MaSo,MaSoKho,MaTaiKhoan,MaVattu,Luong_0,Tien_0) " _
            & " SELECT DISTINCTROW MaSo+" + CStr(ms) + ",MaSoKho,MaTaiKhoan,MaVattu,Luong_0-SoXuat," + RoundMoneySQL("Tien_0*(Luong_0-SoXuat)/Luong_0") + " FROM VTDauNam WHERE Luong_0>SoXuat"
            
        ExecuteSQL5 "INSERT INTO VTDauNam (MaSo,MaSoKho,MaTaiKhoan,MaVattu,Luong_0,Tien_0) " _
            & " SELECT DISTINCTROW MaSo+" + CStr(ms) + ",MaKho,MaTKNo,MaVattu,SoPS2No-SoXuat," + RoundMoneySQL("SoPS*(SoPS2No-SoXuat)/SoPS2No") + " FROM ChungTu WHERE MaLoai=1 AND MaVattu>0 AND SoPS2No>SoXuat"
        
        ExecuteSQL5 "DELETE * FROM VTDauNam WHERE MaSo<=" + CStr(ms)
    End If
    
    If pCongNoHD <> 0 Then
        ktracongno2004
    
        ms = Lng_MaxValue("MaSo", "CNDauNam")
    
        ExecuteSQL5 "INSERT INTO CNDauNam (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0) " _
            & " SELECT DISTINCTROW MaSo+" + CStr(ms) + ",MaTaiKhoan,MaKhachHang,CNDauNam.DuNo_0-CNDauNam.SoXuat,DuCo_0 FROM CNDauNam INNER JOIN HethonTK ON CNDauNam=HethongTK.MaSo WHERE DuNo_0>SoXuat AND TK_ID=" + CStr(TKCNKH_ID)
        ExecuteSQL5 "INSERT INTO CNDauNam (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0) " _
            & " SELECT DISTINCTROW MaSo+" + CStr(ms) + ",MaTaiKhoan,MaKhachHang,CNDauNam.DuNo_0,DuCo_0-CNDauNam.SoXuat FROM CNDauNam INNER JOIN HethonTK ON CNDauNam=HethongTK.MaSo WHERE DuNo_0>SoXuat AND TK_ID=" + CStr(TKCNPT_ID)
    
        ExecuteSQL5 "INSERT INTO CNDauNam (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0) " _
            & " SELECT DISTINCTROW MaSo+" + CStr(ms) + ",MaTKNo,MaKH,SoPS-SoXuat,0 FROM " + ChungTu2TKNC(-1) + " WHERE MaKH>0 AND SoPS>SoXuat AND TK_ID=" + CStr(TKCNKH_ID)
        ExecuteSQL5 "INSERT INTO CNDauNam (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0) " _
            & " SELECT DISTINCTROW MaSo+" + CStr(ms) + ",MaTKCo,MaKHC,0,SoPS-SoXuat FROM " + ChungTu2TKNC(1) + " WHERE MaKHC>0 AND SoPS>SoXuat AND TK_ID=" + CStr(TKCNPT_ID)
    
        ExecuteSQL5 "DELETE * FROM CNDauNam WHERE MaSo<=" + CStr(ms)
    End If
    
    ' Xãa ph¸t sinh
    HienThongBao "Xãa c¸c ph¸t sinh ...", 1
    ExecuteSQL5 "UPDATE ChungTu SET ThangCT = - ThangCT WHERE (MaLoai = 5 OR MaLoai = 6) AND (MaVattu > 0)"
    ExecuteSQL5 "DELETE FROM ChungTu WHERE MaCT  > 0 AND ThangCT > 0"
    ExecuteSQL5 "DELETE FROM ChungTuLQ"
    ExecuteSQL5 "DELETE FROM HoaDon"
    
    ' Chuyen so du tai khoan cuoi nam ve dau nam
    HienThongBao "ChuyÓn sè d­ tµi kho¶n cuèi n¨m vÒ ®Çu n¨m ...", 1
    ThemTruong "HethongTK", "PSNLK" + CStr(pNamTC), dbDouble
    ThemTruong "HethongTK", "PSCLK" + CStr(pNamTC), dbDouble
    sql = "UPDATE HethongTK SET DuNo_0 = DuNo_12, DuCo_0 = DuCo_12, DuNT_0 = DuNT_12"
    For i = 1 To 12
        st = CStr(i)
        sqln = sqln + "+No_" + CStr(i)
        sqlc = sqlc + "+Co_" + CStr(i)
        sql = sql + ", No_" + st + " = 0, Co_" + st + " = 0, No_" + st + "_NT = 0, Co_" + st + "_NT = 0, DuNo_" _
                    + st + " = DuNo_12, DuCo_" + st + " = DuCo_12, DuNT_" + st + " = DuNT_12"
    Next
    sql = sql + ",PSNLK=" + sqln + ",PSCLK=" + sqlc + ",PSNLK" + CStr(pNamTC) + "=" + sqln + ",PSCLK" + CStr(pNamTC) + "=" + sqlc
    ExecuteSQL5 sql
    sql = "UPDATE SoDuKhachHang SET DuNo_0 = DuNo_12, DuCo_0 = DuCo_12, DuNT_0 = DuNT_12"
    For i = 1 To 12
        st = CStr(i)
        sql = sql + ", No_" + st + " = 0, Co_" + st + " = 0, No_" + st + "_NT = 0, Co_" + st + "_NT = 0, DuNo_" _
                    + st + " = DuNo_12, DuCo_" + st + " = DuCo_12, DuNT_" + st + " = DuNT_12"
    Next
    ExecuteSQL5 sql
    
    ' Chuyen so du vat tu
    HienThongBao "ChuyÓn sè tån kho cuèi n¨m vÒ ®Çu n¨m ...", 1
    sql = "UPDATE TonKho SET Luong_0 = Luong_12, Tien_0 = Tien_12"
    For i = 1 To 12
        st = CStr(i)
        sql = sql + ", Luong_Nhap_" + st + " = 0, Tien_Nhap_" + st + " =0, Luong_Xuat_" + st + " = 0, Tien_Xuat_" _
                    + st + " = 0, Luong_" + st + " = Luong_12, Tien_" + st + " = Tien_12"
    Next
    ExecuteSQL5 sql
    sql = ""
    For i = 0 To 12
        sql = sql + ",Lock" + CStr(i) + "=10*(Lock" + CStr(i) + " \ 10)"
    Next
    ' Ghi nam vao License
    ExecuteSQL5 "UPDATE License SET Thang = 1, NamTC = " + CStr(CInt5(NamCu) + 1) + sql
    ExecuteSQL5 "UPDATE Users SET Psw = Psw + 1"
    
    HienThongBao "ChuyÓn sè d­ TSC§ ...", 1
    ChuyenNamMoiTS
    Dim sttt As String
    'sttt = "INSERT INTO NamTC (Maso,Nam,Path) VALUES (" + CStr(Lng_MaxValue("Maso", "NamTC") + 1) + "," + CStr(pNamTC) + ",'" + dpath + "KT" + NamCu + tail + ".MDB" + "')"
    sttt = "INSERT INTO NamTC (Nam,Path) VALUES (" + CStr(pNamTC) + ",'" + dpath + "KT" + NamCu + tail + ".MDB" + "')"
    ExecuteSQL5 sttt
    pNamTC = pNamTC + 1
    
    MsgBox "D÷ liÖu n¨m " + NamCu + " ®­îc l­u t¹i " + dpath + "KT" + NamCu + tail _
        + ".MDB vµ ch­¬ng tr×nh ®· s½n sµng nhËp d÷ liÖu n¨m " + CStr(pNamTC), vbExclamation, App.ProductName
End Sub
'======================================================================================
' Thñ tôc t¹o chøng tõ kÕt chuyÓn
'======================================================================================
Public Sub KCTuDong(thang As Integer, mkc As Long, tygia As Double)
    Dim rs As Recordset, MaCT As Long, mtkn As Long, mtkc As Long, duno As Double, duco As Double
    Dim oct As New ClsChungtu, TK As New ClsTaikhoan, X As String, s As String, k As Long, K1 As Long, nt As Double
    Dim rskh As Recordset, mk1 As Long, mk2 As Long
    
    XoaCTTheoID thang, 300000000 + mkc, 3
    
    MaCT = Lng_MaxValue("MaCT", "ChungTu")
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW CTKetChuyen.STT,CTKetChuyen.DienGiai,KCChitiet.* FROM CTKetChuyen INNER JOIN KCChitiet ON CTKetChuyen.MaSo=KCChitiet.MaKC WHERE CTKetChuyen.MaSo=" + CStr(mkc), dbOpenSnapshot)
    Do While Not rs.EOF
        If TKChitiet(rs!TK1) <> rs!TK1 Or TKChitiet(rs!tk2) <> rs!tk2 Then
            MsgBox "H·y kiÓm tra l¹i chi tiÕt chøng tõ kÕt chuyÓn sè " + CStr(rs!stt) + " - " + rs!diengiai, vbCritical, App.ProductName
            Exit Do
        End If
        k = rs!TK1
KC1:
'        SoDuTK k, thang, duno, duco
        TK.InitTaikhoanMaSo k
        s = TK.sohieu
        X = CThangDB(thang)
        
    ' them vao sau ket chuyen 154 co chi tiet cong trinh
        
        Dim tondautp
          
        If TK.sohieu Like "154*" Then
          SoDuTK k, thang, duno, duco, nt
        If duno = 0 And duco = 0 Then Exit Sub
'            If duno > 0 Then
'             MsgBox CStr(duno)
'            End If
         
        Set rskh = DBKetoan.OpenRecordset("SELECT *  FROM tp154", dbOpenSnapshot)
            Do While Not rskh.EOF
                tondautp = rskh!dk
                duno = tondautp + SelectSQL("select sum(sops)  as f1 from chungtu where MaTKTCNo =" + CStr(TK.MaSo) + "  and  month(ngaygs) <=" + CStr(thang))
                duco = SelectSQL("select sum(sops)  as f1 from chungtu where MaTKTCCo =" + CStr(TK.MaSo) + "  and month(ngaygs) <=" + CStr(thang))
                duno = duno - duco
                If duno > 0 Then
                    mtkn = rs!tk2
                    mtkc = k
                    mk1 = 0
                    mk2 = rskh!MaSo
              
                oct.InitChungtu 0, 3, "KC_" + CStr(rs!stt), thang, NgayCuoiThang(pNamTC, thang), NgayCuoiThang(pNamTC, thang), 0, 0, rs!diengiai, mtkn, mtkc, duno, IIf(mtkn = k, nt, 0), IIf(mtkc = k, nt, 0), 0, "...", 1, "", "", "", ""
                oct.MaCT = MaCT + rs!stt
                oct.CT_ID = 300000000 + mkc
                oct.MaTP = mk1
                oct.MaTP = mk2
                If pTygia > 0 Then oct.tygia = tygia
                oct.GhiChungtu
                End If
                rskh.MoveNext
          Loop
          Exit Sub
    End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        ' them vao sau ket chuyen 154 co chi tiet cong trinh
        
        If TK.tk_id = TKCNKH_ID Or TK.tk_id = TKCNPT_ID Then
            Set rskh = DBKetoan.OpenRecordset("SELECT MaKhachHang,DuNo_" + X + " AS DN,DuCo_" + X + " AS DC FROM SoDuKhachHang WHERE MaTaiKhoan=" + CStr(k) + " AND (DuNo_" + CStr(thang) + "<>0 OR DuCo_" + CStr(thang) + "<>0)", dbOpenSnapshot)
            Do While Not rskh.EOF
                duno = rskh!dn
                duco = rskh!DC
                If duno > 0 Then
                    mtkn = rs!tk2
                    mtkc = k
                    mk1 = 0
                    mk2 = rskh!MaKhachHang
                Else
                    mtkn = k
                    mtkc = rs!tk2
                    duno = duco - duno
                    mk2 = 0
                    mk1 = rskh!MaKhachHang
                End If
                oct.InitChungtu 0, 3, "KC_" + CStr(rs!stt), thang, NgayCuoiThang(pNamTC, thang), NgayCuoiThang(pNamTC, thang), 0, 0, rs!diengiai, mtkn, mtkc, duno, IIf(mtkn = k, nt, 0), IIf(mtkc = k, nt, 0), 0, "...", 1, "", "", "", ""
                oct.MaCT = MaCT + rs!stt
                oct.CT_ID = 300000000 + mkc
                oct.makh = mk1
                oct.MaKHC = mk2
                If pTygia > 0 Then oct.tygia = tygia
                oct.GhiChungtu
                rskh.MoveNext
            Loop
        Else
            SoDuTK k, thang, duno, duco, nt
        
            If duno = 0 And duco = 0 Then GoTo n
            If duno > 0 Then
                If TK.sohieu Like "154*" Then
                    TK.InitTaikhoanMaSo rs!tk2
                    If (TK.sohieu Like "155*") Or (TK.sohieu Like "632*") Then
                        X = ""
                        Do Until IsNumeric(X)
                            X = InputBox("Tû lÖ kÕt chuyÓn (%)", "KÕt chuyÓn " + s + " sang " + TK.sohieu, "100")
                            If Len(X) = 0 Then X = "0"
                        Loop
                        If CLng5(X) > 100 Then X = "100"
                        duno = RoundMoney(CInt5(X) * duno / 100)
                        If duno = 0 Then GoTo n
                        'TK.InitTaikhoanMaSo rs!TK1
                    End If
                End If
                mtkn = rs!tk2
                mtkc = k
            Else
                mtkn = k
                mtkc = rs!tk2
                duno = duco - duno
            End If
            '///////////////////////////////////// Kiem tra NT
             oct.InitChungtu 0, 3, "KC_" + CStr(rs!stt), thang, NgayCuoiThang(pNamTC, thang), NgayCuoiThang(pNamTC, thang), 0, 0, rs!diengiai, mtkn, mtkc, duno, IIf(mtkn = k, nt, 0), IIf(mtkc = k, nt, 0), 0, "...", 1, "", "", "", ""
             oct.MaCT = MaCT + rs!stt
             oct.CT_ID = 300000000 + mkc
             If pTygia > 0 Then oct.tygia = tygia
             oct.GhiChungtu
n:
             K1 = SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE MaNT<>0 AND SoHieu='" + s + "' AND MaSo>" + CStr(k))
             If K1 > 0 Then
                k = K1
                GoTo KC1:
             End If
                
        End If
         rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set oct = Nothing
    Set TK = Nothing
End Sub
'======================================================================================
' Thñ tôc t¹o chøng tõ ph©n bæ
' tc=1: NVL
' tc=2: NC
' tc=7: M¸y
' tc=4: Doanh thu
'======================================================================================
Public Sub PhanBoCP(tdau As Integer, tcuoi As Integer, shtk As String, tentk As String, tc As Integer, TyLe As Integer)
    Dim sql As String, tongcp As Double, sopb As Integer
    Dim rs As Recordset, i As Integer, tongpb As Double
    Dim rsct As Recordset, ctu As New ClsChungtu, MaCT As Long
    
    Select Case shtk
        Case "641", "642":
            InKetQua tdau, tcuoi, 1, False, 0
        Case Else:
            InChiPhi tdau, tcuoi, False, 0
    End Select
    XDTyLePB tc
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND CT_ID=400000000+" + shtk + " AND MaLoai=3", dbOpenSnapshot)
    If Not rs.EOF Then
        If MsgBox(shtk + " ®· ®­îc kÕt chuyÓn, thùc hiÖn l¹i ?", vbYesNo + vbCritical, App.ProductName) = vbNo Then GoTo KT
    End If
    Do While Not rs.EOF
        ctu.InitChungtu rs!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", "'"
        ctu.XoaChungtu
        rs.MoveNext
    Loop
    MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
    sql = "SELECT Count(MaSo) AS F1 FROM BaoCaoCP"
    sopb = SelectSQL(sql)
    XacdinhKC tdau, tcuoi
    sql = ""
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sql = sql + "+No_" + CStr(i) + "-Co_" + CStr(i)
    Next
    sql = "SELECT DISTINCTROW MaSo,(" + sql + "+KC_C)" + IIf(TyLe < 100, "*" + CStr(TyLe) + "/100", "") + " AS CP FROM HethongTK WHERE SoHieu LIKE '" + shtk + "*' AND TKCon=0 AND (" + sql + "+KC_C)<>0"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        tongpb = 0
        i = 0
        Set rsct = DBKetoan.OpenRecordset("SELECT * FROM BaoCaoCP ORDER BY Kq3", dbOpenSnapshot)
        Do While Not rsct.EOF
            i = i + 1
            If i < sopb Then
                tongcp = RoundMoney(rs!CP * rsct!Kq3)
                tongpb = tongpb + tongcp
                If Abs(tongpb - rs!CP) <= 1 Then
                    tongcp = rs!CP - (tongpb - tongcp)
                    tongpb = rs!CP
                End If
            Else
                tongcp = rs!CP - tongpb
            End If
            If tongcp <> 0 Then
                ctu.InitChungtu 0, 3, "PBCP_" + shtk, tcuoi, NgayCuoiThang(pNamTC, tcuoi), NgayCuoiThang(pNamTC, tcuoi), 0, 0, ABCtoVNI("Ph©n bæ ") + tentk, rsct!MaSo, rs!MaSo, tongcp, 0, 0, 0, "...", 1, "", "", "", ""
                ctu.MaCT = MaCT
                ctu.CT_ID = 400000000 + CLng5(shtk)
                ctu.GhiChungtu
            End If
            rsct.MoveNext
        Loop
        rs.MoveNext
    Loop
KT:
    rs.Close
    Set rs = Nothing
    Set ctu = Nothing
End Sub

Private Sub XDTyLePB(tc As Integer)
    Dim sql As String, cps As String, tongcp As Double
    
    sql = "DELETE FROM BaoCaoCP WHERE CoCon<>0"
    ExecuteSQL5 sql
    
    Select Case tc
        Case 1, 2, 4, 7:  cps = "Kq" + CStr(tc)
        Case 3:               cps = "Kq1+Kq2"
        Case 8:               cps = "Kq1+Kq7"
        Case 9:               cps = "Kq2+Kq7"
        Case 10:            cps = "Kq1+Kq2+Kq7"
    End Select
    
    sql = "SELECT SUM(" + cps + ") AS F1 FROM BaoCaoCP"
    tongcp = SelectSQL(sql)
    If tongcp <= 0 Then Exit Sub
    sql = "UPDATE BaoCaoCP SET Kq3=(" + cps + ")/" + DoiDau(tongcp)
    ExecuteSQL5 sql
End Sub

Public Sub PhanBoCP2(tdau As Integer, tcuoi As Integer, tc As Integer)
    Dim sql As String, tongcp As Double, sopb As Integer, cpkh As Double, tongpbkh As Double, tongcpkh As Double
    Dim i As Integer, tongpb As Double, rsct As Recordset, CP As Double, cp2 As String, kh As String
    
    InChiPhi3 tdau, tcuoi, 0
    ExecuteSQL5 "DELETE FROM BaoCaoCP"
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        cp2 = cp2 + "+CPSXC_" + CStr(i)
        kh = kh + "+CPKH_" + CStr(i)
    Next
    
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo, SoHieu, Kq1,Kq2,Kq3,Kq4,Kq7) SELECT DISTINCTROW MaSo, SoHieu, CPNVL, CPNC, (" _
        + cp2 + ") AS CPSXC, (" + kh + ") AS CPKH, CPM FROM TP154 WHERE (CPNVL<>0 OR CPNC<>0 OR CPM<>0 OR (" + cp2 + ")<>0) AND KPB=0"
    XDTyLePB tc
    
    sql = "SELECT Count(MaSo) AS F1 FROM BaoCaoCP"
    sopb = SelectSQL(sql)
    
    sql = ""
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sql = sql + "+No_" + CStr(i)
    Next
    
    CP = SelectSQL("SELECT DISTINCTROW Sum(" + sql + ") AS F1 FROM HethongTK WHERE SoHieu LIKE '627*' AND TKCon=0") - SoCPTT("627", tdau, tcuoi)
    cpkh = SelectSQL("SELECT DISTINCTROW Sum(" + sql + ") AS F1 FROM HethongTK WHERE SoHieu LIKE '6274*' AND TKCon=0") - SoCPTT("6274", tdau, tcuoi)
    
    For i = CThangDB(tdau) To CThangDB(ThangTruoc(tcuoi))
        ExecuteSQL5 "UPDATE TP154 SET CPSXC_" + CStr(i) + "=0,CPKH_" + CStr(i) + "=0"
    Next
    
    i = 0
    Set rsct = DBKetoan.OpenRecordset("SELECT * FROM BaoCaoCP ORDER BY Kq3", dbOpenSnapshot)
    Do While Not rsct.EOF
        i = i + 1
        If i < sopb Then
            tongcp = RoundMoney(CP * rsct!Kq3)
            tongcpkh = RoundMoney(cpkh * rsct!Kq3)
            tongpb = tongpb + tongcp
            tongpbkh = tongpbkh + tongcpkh
        Else
            tongcp = CP - tongpb
            tongcpkh = cpkh - tongpbkh
        End If
        ExecuteSQL5 "UPDATE TP154 SET CPSXC_" + CStr(CThangDB(tcuoi)) + "=" + DoiDau(tongcp) + ",CPKH_" + CStr(CThangDB(tcuoi)) + "=" + DoiDau(tongcpkh) + " WHERE MaSo=" + CStr(rsct!MaSo)
        rsct.MoveNext
    Loop
    rsct.Close
    Set rsct = Nothing
End Sub
'====================================================================================================
' Thñ tôc in chi phÝ s¶n xuÊt vµ gi¸ thµnh s¶n phÈm
'====================================================================================================
Public Function InChiPhi(tdau As Integer, tcuoi As Integer, msg As Boolean, nn As Integer) As Boolean
    Dim sql As String, i As Integer, rs_tk As Recordset, rs_tktc As Recordset
    
    sql = "DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + "<>0 " + " OR DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + "<>0"
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sql = sql + " OR No_" + CStr(i) + "<>0 " + " OR Co_" + CStr(i) + "<>0"
    Next
    
    ExecuteSQL5 "DELETE FROM BaoCaoCP"
    If ExecuteSQL5("INSERT INTO BaoCaoCP (MaSo, SoHieu, CoCon, MaCha, Cap, Ten, Kq5) SELECT DISTINCTROW MaSo, SoHieu, TkCon, TkCha0, Cap, Ten" + IIf(nn > 0, "E", "") + ", DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) _
        & " FROM HethongTK WHERE (SoHieu LIKE '" + ShTkSPDo + "*' OR SoHieu LIKE '" + ShTkTP + "*') AND (MaTC > 0 AND MaTC <> MaSo) AND (" + sql + ") ORDER BY SoHieu") <> 0 Then GoTo KhongIn
    
    If DBKetoan.RecordsAffected = 0 Then
        If msg Then ErrMsg er_KoPS
        InChiPhi = False
        Exit Function
    End If
        
    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW BaoCaoCP.MaSo, BaoCaoCP.SoHieu, HethongTK.Ten, BaoCaoCP.CoCon" _
        & " FROM BaoCaoCP INNER JOIN HethongTK ON BaoCaoCP.SoHieu = HethongTK.SoHieu ORDER BY BaoCaoCP.SoHieu DESC", dbOpenSnapshot)
    Do While Not rs_tk.EOF
        HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
        If rs_tk!CoCon = 0 Then
            Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE (TK_ID = " + CStr(TKCPVL_ID) + " OR TK_ID = " + CStr(TKVT_ID) + ") AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
            Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1 = Kq1 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, -1)) _
                    & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
            Loop
            
            Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE TK_ID = " + CStr(TKCPNC_ID) + " AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
            Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq2 = Kq2 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, -1)) _
                    & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
            Loop
                        
            Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE TK_ID = " + CStr(TKCPTB_ID) + " AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
            Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq7 = Kq7 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, -1)) _
                    & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
            Loop
            
            Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE TK_ID = " + CStr(TKCPSX_ID) + " AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
            Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq3 = Kq3 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, -1)) _
                    & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
            Loop
            
            Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE TK_ID = " + CStr(TKCPSX_ID) + " AND TK_ID2 = " + CStr(TKCPSX_ID) + " AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
            Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq4 = Kq4 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, -1)) _
                    & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
            Loop
            
            Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE ((TK_ID = " + CStr(TKCPGIA_ID) + " OR TK_ID2 = " + CStr(TKCPGIA_ID) + ") OR (SoHieu LIKE '" + ShTkTP + "*')) AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
            Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq6 = Kq6 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, 1)) _
                    & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
            Loop
        Else
            Set rs_tktc = DBKetoan.OpenRecordset("SELECT SUM(Kq1) As SKq1, SUM(Kq2) As SKq2, SUM(Kq3) As SKq3," _
                & " SUM(Kq4) As SKq4, SUM(Kq5) As SKq5, SUM(Kq6) As SKq6, SUM(Kq7) As SKq7, SUM(Kq8) As SKq8 FROM BaoCaoCP WHERE MaCha = " + CStr(rs_tk!MaSo), dbOpenSnapshot, dbForwardOnly)
            ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1 = " + DoiDau(rs_tktc!SKq1) + ", Kq2 = " + DoiDau(rs_tktc!SKq2) + ", Kq3 = " + DoiDau(rs_tktc!SKq3) _
                & ", Kq4 = " + DoiDau(rs_tktc!SKq4) + ", Kq5 = " + DoiDau(rs_tktc!SKq5) + ", Kq6 = " + DoiDau(rs_tktc!SKq6) + ", Kq7 = " + DoiDau(rs_tktc!SKq7) + ", Kq8 = " + DoiDau(rs_tktc!SKq8) + " WHERE MaSo = " + CStr(rs_tk!MaSo)
        End If
        rs_tktc.Close
        rs_tk.MoveNext
    Loop
    rs_tk.Close
    Set rs_tk = Nothing
    Set rs_tktc = Nothing
    
    XDCapTK ShTkSPDo
    XDCapTK ShTkTP
    sql = "UPDATE HethongTK INNER JOIN BaoCaoCP ON HethongTK.MaSo = BaoCaoCP.MaSo SET BaoCaoCP.CoCon = HethongTK.CapDuoi"
    ExecuteSQL5 sql
    
    On Error GoTo KT
    frmMain.Rpt.ReportFileName = "CHIPHI.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    On Error GoTo 0
    InChiPhi = True
    GoTo KT
KhongIn:
    InChiPhi = False
    ErrMsg er_RWait
KT:
    HienThongBao "", 1
End Function

Public Function InChiPhi2(tdau As Integer, tcuoi As Integer, nn As Integer) As Boolean
    Dim sql As String, wsql As String, i As Integer
    
    TinhTonVT ThangTruoc(tdau), "155"
    For i = tdau To tcuoi
        wsql = wsql + " AND Luong_Nhap_" + CStr(CThangDB(i)) + "=0 "
    Next
    SetSQL "MienTru", "SELECT MaTP FROM ThanhPham WHERE " + WThang("Thang", tdau, tcuoi) + " GROUP BY MaTP"
    sql = "SELECT Vattu.MaSo,Vattu.SoHieu,TenVattu,DonVi,L,T,Sum(SoLuong) AS SL,Sum(CPNVL) AS NVL,Sum(CPNC) AS NC,Sum(CPKH) AS KH,Sum(CPSXC) AS SXC,Sum(CPBH) AS BH,Sum(CPQL) AS QL,Sum(CPTC) AS TC,First(PhanLoaiVattu.SoHieu) AS SHPL3,First(TenPhanLoai) AS TenPL3,First(PLCha) AS PCha FROM (ThanhPham INNER JOIN Vattu ON ThanhPham.MaTP=Vattu.MaSo) LEFT JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo WHERE " + WThang("Thang", tdau, tcuoi) + " GROUP BY Vattu.MaSo,Vattu.SoHieu,TenVattu,DonVi, L, T" _
        & " UNION SELECT Vattu.MaSo,Vattu.SoHieu,TenVattu,DonVi,L,T,0 AS SL,0 AS NVL,0 AS NC,0 AS KH,0 AS SXC,0 AS BH,0 AS QL,0 AS TC,First(PhanLoaiVattu.SoHieu) AS SHPL3,First(TenPhanLoai) AS TenPL3,First(PLCha) AS PCha FROM ((TonKho INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo) LEFT JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo) LEFT JOIN MienTru ON TonKho.MaVattu=MienTru.MaTP WHERE IsNull(MienTru.MaTP) AND L<>0 " + wsql + " GROUP BY Vattu.MaSo,Vattu.SoHieu,TenVattu,DonVi, L, T"
    SetSQL "QChitiet", sql
    sql = "SELECT MaVattu,Sum(SoPS) AS Tien,Sum(SoPS2Co) AS Luong FROM " + ChungTu2TKNC(1) + " WHERE TK_ID=" + CStr(TKDT_ID) + " AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY MaVattu"
    SetSQL "QSoCai", sql
    sql = "SELECT QChitiet.*,QSocai.Tien,QSoCai.Luong,PhanLoaiVattu.SoHieu AS SHPL2,TenPhanLoai AS TenPL2,PLCha FROM (QChitiet LEFT JOIN QSoCai ON QChitiet.MaSo=QSoCai.MaVattu) LEFT JOIN PhanLoaiVattu ON QChitiet.PCha=PhanLoaiVattu.MaSo"
    SetSQL "QDuPhong", sql
    sql = "SELECT QDuPhong.*,PhanLoaiVattu.SoHieu AS SHPL1,TenPhanLoai AS TenPL1 FROM QDuPhong LEFT JOIN PhanLoaiVattu ON QDuPhong.PLCha=PhanLoaiVattu.MaSo ORDER BY QDuPhong.Sohieu"
    SetSQL "QNhatKy", sql
    
    frmMain.Rpt.ReportFileName = "CHIPHI2.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    InChiPhi2 = True
KT:
    HienThongBao "", 1
    Exit Function
KhongIn:
    InChiPhi2 = False
    HienThongBao "", 1
    ErrMsg er_RWait
End Function

Public Function InChiPhi3(tdau As Integer, tcuoi As Integer, nn As Integer) As Boolean
    Dim rs As Recordset, tp As New Cls154, cp2 As String, i As Integer, kh As String
    Dim s621 As String, s622 As String, s623 As String
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo FROM TP154", dbOpenSnapshot)
    Do While Not rs.EOF
        tp.InitTPMaSo rs!MaSo
        HienThongBao tp.TenVattu, 1
        tp.XDDauKy tdau
        tp.GhiCPTT tdau, tcuoi
        tp.XDCuoiKy tcuoi
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        cp2 = cp2 + "+CPSXC_" + CStr(i) + "+CPSXCTT" + CStr(i)
        kh = kh + "+CPKH_" + CStr(i)
        s621 = s621 + "+CPNVLPB" + CStr(i)
        s622 = s622 + "+CPNCPB" + CStr(i)
        s623 = s623 + "+CPMPB" + CStr(i)
    Next
    
'    SetSQL "QChitiet", "SELECT DISTINCTROW TP154.SoHieu, TenVattu, PhanLoai154.SoHieu AS SHPL, TenPhanLoai," _
'        & " DK1, CPNVL+" + s621 + " AS NVL, CPNC+" + s622 + " AS NC, CPM+" + s623 + " AS M, (" + cp2 + ") AS CPSXC, (" + kh + ") AS CPKH, CK1, PhanLoai154.PLCha, SanLuong" _
'        & " FROM TP154 INNER JOIN PhanLoai154 ON TP154.MaPhanLoai=PhanLoai154.MaSo" _
'        & " WHERE DK1<>0 OR CPNVL+" + s621 + "<>0 OR CPNC+" + s622 + "<>0 OR CPM+" + s623 + "<>0 OR (" + cp2 + ")<>0"
'    SetSQL "QSoCai", "SELECT QChitiet.*, PhanLoai154.SoHieu AS SH2,PhanLoai154.TenPhanLoai AS TenPhanLoai2,PhanLoai154.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoai154 ON QChitiet.PLCha=PhanLoai154.MaSo"
'    SetSQL "QTongHopCT", "SELECT QSoCai.*, PhanLoai154.SoHieu AS SH1,PhanLoai154.TenPhanLoai AS TenPhanLoai1 FROM QSoCai LEFT JOIN PhanLoai154 ON QSoCai.PLCha2=PhanLoai154.MaSo"
        
        
         Dim stt As String
         
   SetSQL "QChitiet", "SELECT DISTINCTROW  Str(TP154.DT) as SoHieu,TP154.Sohieu +'                    '+TP154.TenVattu as TenVattu, PhanLoai154.SoHieu AS SHPL, TenPhanLoai," _
        & " DK1, CPNVL+" + s621 + " AS NVL, CPNC+" + s622 + " AS NC, CPM+" + s623 + " AS M, (" + cp2 + ") AS CPSXC, (" + kh + ") AS CPKH, CK1, PhanLoai154.PLCha, SanLuong" _
        & " FROM TP154 INNER JOIN PhanLoai154 ON TP154.MaPhanLoai=PhanLoai154.MaSo" _
        & " WHERE TP154.DT <> 0  or DK1<>0 OR CPNVL+" + s621 + "<>0 OR CPNC+" + s622 + "<>0 OR CPM+" + s623 + "<>0 OR (" + cp2 + ")<>0"
        
   stt = " QChitiet.SoHieu, QChitiet.TenVattu, QChitiet.SHPL, QChitiet.TenPhanLoai,QChitiet.DK1,QChitiet.NVL,QChitiet.NC,QChitiet.M,QChitiet.CPSXC,QChitiet.CPKH,QChitiet.CK1,QChitiet.PLCha,QChitiet.Sanluong"
    
   SetSQL "QSoCai", "SELECT " + stt + ", PhanLoai154.SoHieu AS SH2,PhanLoai154.TenPhanLoai AS TenPhanLoai2,PhanLoai154.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoai154 ON QChitiet.PLCha=PhanLoai154.MaSo"
   stt = "  QSoCai.Sohieu, QSoCai.TenVattu, QSoCai.SHPL, QSoCai.TenPhanLoai,QSoCai.DK1," _
            & "QSoCai.NVL,QSoCai.NC,QSoCai.M,QSoCai.CPSXC,QSoCai.CPKH,QSoCai.CK1,QSoCai.PLCha,QSoCai.Sanluong," _
            & "QSoCai.SH2,QSoCai.TenPhanLoai2,QSoCai.PLCha2 "
      stt = "SELECT " + stt + ",PhanLoai154.SoHieu AS SH1,PhanLoai154.TenPhanLoai AS TenPhanLoai1  FROM QSoCai LEFT JOIN PhanLoai154 ON QSoCai.PLCha2=PhanLoai154.MaSo"
    SetSQL "QTongHopCT", stt
    
    
    On Error GoTo KT
    frmMain.Rpt.ReportFileName = "CHIPHI3.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    On Error GoTo 0
    
    InChiPhi3 = True
    GoTo KT
KhongIn:
    InChiPhi3 = False
    ErrMsg er_RWait
KT:
    Set tp = Nothing
    HienThongBao "", 1
End Function

Public Function InKetQua3(tdau As Integer, tcuoi As Integer, loaibc As Integer, nn As Integer, Optional ct As Integer = 0) As Boolean
    Dim rs As Recordset, tp As New Cls154, cp1 As String, cp2 As String, cp3 As String, sxc As String, i As Integer, s62x As String
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    If ct = 0 Then
        ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo, SoHieu, Ten,Kq1,Kq2) VALUES (10001,'HDTC','" + ABCtoVNI("KÕt qu¶ ho¹t ®éng tµi chÝnh") + "'," + DoiDau(PSDu("911", "635", tdau, tcuoi)) + "," + DoiDau(PSDu("515", "911", tdau, tcuoi)) + ")"
        ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo, SoHieu, Ten,Kq1,Kq2) VALUES (10002,'HDK','" + ABCtoVNI("KÕt qu¶ ho¹t ®éng kh¸c") + "'," + DoiDau(PSDu("911", "811", tdau, tcuoi)) + "," + DoiDau(PSDu("711", "911", tdau, tcuoi)) + ")"
        ExecuteSQL5 "DELETE * FROM BaoCaoCP2 WHERE Kq1=0 AND Kq2 =0"
    End If
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo FROM TP154", dbOpenSnapshot)
    Do While Not rs.EOF
        tp.InitTPMaSo rs!MaSo
        HienThongBao tp.TenVattu, 1
        tp.XDDauKy IIf(loaibc = 0, tdau, pThangDauKy)
        tp.GhiCPTT IIf(loaibc = 0, tdau, pThangDauKy), tcuoi
        tp.XDCuoiKy tcuoi
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    For i = CThangDB(IIf(loaibc = 0, tdau, pThangDauKy)) To CThangDB(tcuoi)
        cp1 = cp1 + "+CPBH" + CStr(i) + "+CPBHTT" + CStr(i)
        cp2 = cp2 + "+CPQL" + CStr(i) + "+CPQLTT" + CStr(i)
        cp3 = cp3 + "+CPTC" + CStr(i) + "+CPTCTT" + CStr(i)
        sxc = sxc + "+CPSXC_" + CStr(i) + "+CPSXCTT" + CStr(i)
        s62x = s62x + "+CPNVLPB" + CStr(i) + "+CPNCPB" + CStr(i) + "+CPMPB" + CStr(i)
    Next
        
   SetSQL "MienTru", "SELECT DISTINCTROW TP154.SoHieu, TenVattu, PhanLoai154.SoHieu AS SHPL, TenPhanLoai," _
        & " (DK1+CPNVL+ CPNC+ CPM+ " + s62x + sxc + "-CK1) AS GT," + cp1 + " AS CPBH," + cp2 + " AS CPQL, DT, PhanLoai154.PLCha," + cp3 + " AS CPTC" _
        & " FROM TP154 INNER JOIN PhanLoai154 ON TP154.MaPhanLoai=PhanLoai154.MaSo" _
        & " WHERE (DK1-CK1)<>0 OR CPNVL<>0 OR CPNC<>0 OR CPM<>0 OR (" + s62x + sxc + ")<>0 OR (" + cp1 + ")<>0 OR (" + cp2 + ")<>0 OR (" + cp3 + ")<>0 OR DT<>0 ORDER BY PhanLoai154.SoHieu, TP154.SoHieu" _
        & " UNION SELECT SoHieu, Ten AS TenVattu, 'x' AS SHPL, '' AS TenPhanLoai, Kq1 AS GT, 0 AS CPBH, 0 AS CPQL, Kq2, 0 AS PLCha,0 AS CPTC FROM BaoCaoCP2"
    SetSQL "QTongHopCT", "SELECT * FROM MienTru ORDER BY SHPL, SoHieu"
    
    On Error GoTo KT
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\KETQUA" + IIf(loaibc = 0, "3", "4") + ".RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(IIf(loaibc = 0, tdau, pThangDauKy), tcuoi, nn) + "'"
    On Error GoTo 0
    InKetQua3 = True
    GoTo KT
KhongIn:
    InKetQua3 = False
    ErrMsg er_RWait
KT:
    Set tp = Nothing
    HienThongBao "", 1
End Function
'==========================================================================================================
' Thñ tôc in chi phÝ s¶n xuÊt vµ gi¸ thµnh s¶n phÈm
'==========================================================================================================
Public Function InKetQua(tdau As Integer, tcuoi As Integer, tag As Integer, msg As Boolean, nn As Integer) As Boolean
    Dim sql As String, i As Integer, rs_tk As Recordset, rs_tktc As Recordset, st As String
    
    st = CStr(CThangDB(ThangTruoc(tdau)))
    sql = "DuNo_" + st + "<>0 OR DuCo_" + st
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sql = sql + " OR No_" + st + "<>0 " + " OR Co_" + st + "<>0 "
    Next
    
    ExecuteSQL5 "DELETE FROM BaoCaoCP"
    If ExecuteSQL5("INSERT INTO BaoCaoCP (MaSo, SoHieu, CoCon, MaCha, Cap, Ten) SELECT DISTINCTROW MaSo, SoHieu, TkCon, TkCha0, Cap, Ten" + IIf(nn > 0, "E", "") _
        & " FROM HethongTK WHERE (SoHieu LIKE '" + ShTkKQ + "*') AND (MaTC > 0 AND MaTC <> MaSo) AND (" + IIf(tag = 0, sql, "TRUE") + ") ORDER BY SoHieu") <> 0 Then GoTo KhongIn
    
    If DBKetoan.RecordsAffected = 0 Then
        If msg Then ErrMsg er_KoPS
        InKetQua = False
        Exit Function
    End If
       
    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW BaoCaoCP.MaSo, BaoCaoCP.SoHieu, HethongTK.Ten, BaoCaoCP.CoCon" _
        & " FROM BaoCaoCP INNER JOIN HethongTK ON BaoCaoCP.SoHieu = HethongTK.SoHieu ORDER BY BaoCaoCP.SoHieu DESC", dbOpenSnapshot)
    Do While Not rs_tk.EOF
        HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
        If rs_tk!CoCon = 0 Then
           Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE (TK_ID = " + CStr(TKCPGIA_ID) + " OR SoHieu LIKE '635*') AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
           Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1 = Kq1 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, -1)) _
                    & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
           Loop
           
           Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE TK_ID = " + CStr(TKCPBT_ID) + " AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
           Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1 = Kq1 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, -1)) _
                    & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
           Loop
                
           Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE TK_ID = " + CStr(TKCPQL_ID) + " AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
           Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq2 = Kq2 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, -1)) _
                    & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
           Loop
                
           Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE TK_ID = " + CStr(TKCPBH_ID) + " AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
           Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq3 = Kq3 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, -1)) _
                    & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
           Loop
            
           Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE (TK_ID = " + CStr(TKDT_ID) + " OR SoHieu LIKE '515*') AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
           Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq4 = Kq4 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, 1)) _
                     & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
           Loop
           
           Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE TK_ID = " + CStr(TKTNBT_ID) + " AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
           Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq4 = Kq4 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, 1)) _
                     & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
           Loop
           Set rs_tktc = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE SoHieu LIKE '635*' AND MaTC = MaSo", dbOpenSnapshot, dbForwardOnly)
           Do While Not rs_tktc.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP SET Kq5 = Kq5 + " + DoiDau(PhatSinhDu(rs_tk!MaSo, rs_tktc!MaSo, tdau, tcuoi, -1)) _
                     & " WHERE MaSo = " + CStr(rs_tk!MaSo)
                rs_tktc.MoveNext
           Loop
        Else
            Set rs_tktc = DBKetoan.OpenRecordset("SELECT SUM(Kq1) As SKq1, SUM(Kq2) As SKq2, SUM(Kq3) As SKq3," _
                & " SUM(Kq4) As SKq4, SUM(Kq5) As SKq5 FROM BaoCaoCP WHERE MaCha = " + CStr(rs_tk!MaSo), dbOpenSnapshot, dbForwardOnly)
            ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1 = " + DoiDau(rs_tktc!SKq1) + ", Kq2 = " + DoiDau(rs_tktc!SKq2) + ", Kq3 = " + DoiDau(rs_tktc!SKq3) _
                & ", Kq4 = " + DoiDau(rs_tktc!SKq4) + ", Kq5 = " + DoiDau(rs_tktc!SKq5) + " WHERE MaSo = " + CStr(rs_tk!MaSo)
        End If
        rs_tktc.Close
        rs_tk.MoveNext
    Loop
    rs_tk.Close
    Set rs_tk = Nothing
    Set rs_tktc = Nothing
    
    XDCapTK ShTkKQ
    sql = "UPDATE HethongTK INNER JOIN BaoCaoCP ON HethongTK.MaSo = BaoCaoCP.MaSo SET BaoCaoCP.CoCon = HethongTK.CapDuoi"
    ExecuteSQL5 sql
    
    On Error GoTo KT
    frmMain.Rpt.ReportFileName = "KETQUA.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    On Error GoTo 0
    InKetQua = True
    GoTo KT
KhongIn:
    InKetQua = False
    ErrMsg er_RWait
KT:
    HienThongBao "", 1
End Function
'====================================================================================================
' Hµm tr¶ vÒ sè ph¸t sinh theo ®èi øng gi÷a chi tiÕt vµ tµi kho¶n tµi chÝnh
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
    If IsNull(rs_kq!ketqua) Then
        PhatSinhDu = 0
    Else
        PhatSinhDu = rs_kq!ketqua
    End If
    
    sql = "SELECT SUM(SoPS) As KetQua FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi)
    If loai = -1 Then
        sql = sql + " AND MaTkCo = " + CStr(mct) + IIf(mtk > 0, " AND MaTkTCNo = " + CStr(mtk), "")
    Else
        sql = sql + " AND MaTkNo = " + CStr(mct) + IIf(mtk > 0, " AND MaTkTCCo = " + CStr(mtk), "")
    End If
    
    Set rs_kq = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If Not IsNull(rs_kq!ketqua) Then
        PhatSinhDu = PhatSinhDu - rs_kq!ketqua
    End If
    
    rs_kq.Close
    Set rs_kq = Nothing
End Function

Public Sub LayThongtinCT(MaCT As Long, loai As Integer, Ten As String, DiaChi As String, Optional ctgoc As String, Optional makh As Long, Optional p As Integer = 0)
    Dim rs As Recordset
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW * FROM ChungTuLQ" + IIf(p > 0, "P", "") + " WHERE MaCT=" + CStr(MaCT) + " AND Loai=" + CStr(loai), dbOpenSnapshot)
    If rs.recordCount > 0 Then
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

Public Sub UpDateDB()
    Dim tdf As TableDef, rel As Relation, TK As New ClsTaikhoan
    Dim i As Integer, sql As String, j As Integer, s As Boolean
    
    If DBKetoan.TableDefs("Users").Fields("VT").Type = dbInteger Then
        DBKetoan.TableDefs("Users").Fields("VT").Name = "OR"
        DBKetoan.TableDefs("Users").Fields.Append DBKetoan.TableDefs("Users").CreateField("VT", dbLong)
        DBKetoan.TableDefs("Users").Fields!vt.DefaultValue = 0
        ExecuteSQL5 "UPDATE Users SET VT=IIF(MaSo=1,1111111111,0)"
    End If
    
    If ThemTruong("License", "FontName", dbText, 30) Then
        ThemTruong "License", "FontSize", dbInteger
        If FontDaCo("VNI-Times") Then
            sql = "VNI-Times"
            i = 10
        Else
            sql = IIf(FontDaCo(sFONTNAME), sFONTNAME, "MS Sans Serif")
            i = 8
        End If
        ExecuteSQL5 "UPDATE License SET FontName='" + sql + "',FontSize=" + CStr(i)
    End If
    
    ThemTruong "License", "App1Path", dbText, 50
    ThemTruong "License", "LoaiTien", dbLong
    ThemTruong "License", "TyGia", dbDouble
    ThemTruong "License", "Flag1", dbLong
    ThemTruong "Users", "WS", dbText, 20
    
    s = pSongNgu
    pSongNgu = False
    ExecuteSQL5 "UPDATE " + ChungTu2TKNC(-1) + " SET SoPS2Co=SoPS2No,SoPS2No=SoPS2Co,MaTKCo=MaTKNo,MaTKTCCo=MaTKTCNo,MaTKNo=MaTKCo,MaTKTCNo=MaTKTCCo WHERE HethongTK.SoHieu LIKE '3331*' AND SoPS=0"
    ExecuteSQL5 "UPDATE " + ChungTu2TKNC(1) + " SET SoPS2Co=SoPS2No,SoPS2No=SoPS2Co,MaTKCo=MaTKNo,MaTKTCCo=MaTKTCNo,MaTKNo=MaTKCo,MaTKTCNo=MaTKTCCo WHERE HethongTK.SoHieu LIKE '" + pVATV + "*' AND SoPS=0"
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID=5000 WHERE SoHieu LIKE '511*'"
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID=0 WHERE SoHieu LIKE '" + pVATV + "112*' OR SoHieu LIKE '" + pVATV + "212*'"
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID2=" + CStr(TKLT_ID) + " WHERE SoHieu LIKE '" + pSHPT + "*' OR SoHieu LIKE '331*' OR SoHieu LIKE '138*' OR SoHieu LIKE '338*' OR Cap=0"
    ExecuteSQL5 "UPDATE License SET TKVattu='4' WHERE Instr(TKVattu,'-')=0"
    ExecuteSQL5 "UPDATE License SET Flag1=Flag1+300000000 WHERE Flag1\100000000=0"
        
    XoaBang "KQKD911"
        
    ThemTruong "Thue", "DaHoan", dbDouble
    ThemTruong "Thue", "MienGiam", dbDouble
    
    ThemTruong "License", "CTGS_GV", dbLong
    ThemTruong "License", "Quan", dbText, 30
    ThemTruong "License", "ThanhPho", dbText, 30
    ThemTruong "License", "SMTP", dbText, 30, 0, "mail.hn.vnn.vn"
    ThemTruong "License", "EMail", dbText, 30, 0, "1@2"
    ThemTruong "License", "EMailDB", dbText, 30, 0, "" ' "unesco@hn.vnn.vn"
    ThemTruong "License", "MV", dbInteger
    If ThemTruong("License", "SoKT", dbInteger) Then ExecuteSQL5 "UPDATE License SET SoKT=111"
    
    ThemTruong "License", "MKUP", dbLong
    ThemTruong "License", "MST_ID", dbLong
    
    For i = 0 To 12
        ThemTruong "License", "Lock" + CStr(i), dbInteger
    Next
    
    If ThemTruong("HethongTK", "TK_ID3", dbLong) Then ExecuteSQL5 "UPDATE HethongTK SET TK_ID3=1 WHERE SoHieu LIKE '336*' Or SoHieu LIKE '334*'"
    ThemTruong "HethongTK", "TenDA", dbText, 50
    ThemTruong "HethongTK", "NhomDA", dbText, 3
    ThemTruong "HethongTK", "DiaDiem", dbText, 50
    
    ThemTruong "HethongTK", "DuToan", dbDouble
    ThemTruong "HethongTK", "Von1", dbDouble
    ThemTruong "HethongTK", "Von2", dbDouble
    ThemTruong "HethongTK", "Von3", dbDouble
    
    ThemTruong "HethongTK", "NgayKC", dbDate
    ThemTruong "HethongTK", "NgayHT", dbDate
    
    ThemTruong "HethongTK", "PSNLK", dbDouble
    ThemTruong "HethongTK", "PSCLK", dbDouble
    
    If Not BangDaCo("NamTC") Then
        CopyTable pCurDir + "UPDATE.MDB", "NamTC"
                
        For i = 9 To 0 Step -1
            sql = GetSetting(IniPath, "LastYear", CStr(pNamTC + i - 10), "")
            If Len(sql) > 0 Then ExecuteSQL5 "INSERT INTO NamTC (MaSo,Nam,Path) VALUES (" + CStr(Lng_MaxValue("MaSo", "NamTC") + 1) + "," + CStr(pNamTC + i - 10) + ",'" + sql + "')"
        Next
    End If
    
    If Not BangDaCo("Reports") Then CopyTable pCurDir + "UPDATE.MDB", "Reports"
        
    ThemTruong "HoaDon", "NK", dbInteger
    If ThemTruong("HoaDon", "TS", dbInteger) Then ExecuteSQL5 "UPDATE " + ChungTu2TKHD(0) + " SET TS=1 WHERE ChungTu.MaLoai=9"
    ThemTruong "HoaDon", "DC", dbInteger
    ThemTruong "PhieuNX", "DonGia", dbDouble
    ThemTruong "PhieuNX", "TyLe", dbInteger
    ThemTruong "PhieuNX", "Thue", dbDouble
    ThemTruong "PhieuNX", "ThanhTien2", dbDouble
    ThemTruong "PhieuNX", "CK", dbDouble
    
    ThemTruong "ChungTu", "User_ID", dbLong, , 1
    
    If ThemTruong("PhanLoaiVattu", "MaTK", dbLong) Then
          ExecuteSQL5 "UPDATE PhanLoaiVattu SET MaTK=" + CStr(SelectSQL("SELECT TOP 1 MaSo AS F1 FROM HethongTK WHERE TKCon=0 AND SoHieu LIKE '156*' ORDER BY SoHieu"))
    End If
    
    ThemTruong "KhoHang", "MaTK", dbLong
    ThemTruong "KhoHang", "MaTKGV", dbLong
            
    ThemTruong "ChungTu", "DVT", dbInteger
    ThemTruong "ChungTu", "XuLy", dbInteger
    
    ThemTruong "PhieuNX", "DVT", dbText, 20
    If ThemTruong("ChungTuLQ", "MaKH", dbLong) Then
        ExecuteSQL5 "UPDATE (ChungTuLQ INNER JOIN ChungTu ON ChungTuLQ.MaCT=ChungTu.MaCT) INNER JOIN HoaDon ON ChungTu.MaSo=HoaDon.MaSo SET ChungTuLQ.MaKH=MaKhachHang WHERE ChungTuLQ.Loai=0"
    End If
    
    ThemTruong "PhanLoaiVattu", "PLCon", dbInteger
    ThemTruong "PhanLoaiVattu", "PLCha", dbInteger
    ThemTruong "PhanLoaiVattu", "Cap", dbInteger, 2, 1
    ThemTruong "TonKho", "SoXuat", dbDouble
    ThemTruong "HoaDon", "HDBL", dbInteger

    ThemTruong "Vattu", "Dvt2", dbInteger
    ThemTruong "Vattu", "Donvi2", dbText, 20
    ThemTruong "Vattu", "TyLeQD", dbSingle
    ThemTruong "Vattu", "GhiChu", dbText, 50
    ThemTruong "Vattu", "VAT", dbInteger
    ThemTruong "Vattu", "GiaBan1", dbDouble
    ThemTruong "Vattu", "GiaBan2", dbDouble
    ThemTruong "Vattu", "GiaBan3", dbDouble
    ThemTruong "Vattu", "CK", dbDouble
    ThemTruong "Vattu", "ThueNK", dbDouble
    ThemTruong "Vattu", "L", dbDouble
    ThemTruong "Vattu", "T", dbDouble
    
    If Not BangDaCo("DVTVattu") Then
        CopyTable pCurDir + "UPDATE.MDB", "DVTVattu"
        
        Set rel = DBKetoan.CreateRelation("DVTVT", "Vattu", "DVTVattu", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaVattu"
        DBKetoan.Relations.Append rel
        
        ExecuteSQL5 "INSERT INTO DVTVattu (MaSo,MaVattu,DonVi,TyLeQD) SELECT " + CStr(Lng_MaxValue("MaSo", "DVTVattu") + 1) + ",MaSo,DonVi2,TyLeQD FROM Vattu WHERE Dvt2=1"
    End If
        
    ThemTruong "ChungTu", "CTGS", dbLong, 0, 1
    ThemTruong "ChungTu", "MaTP", dbLong
    ThemTruong "ChungTu", "TyGia", dbDouble, , 1
    ThemTruong "ChungTu", "MaNV", dbLong
    
    For i = 1 To 3
        ThemTruong "ChungTu", "MaDT" + CStr(i), dbLong
        ThemTruong "ChungTu2", "MaDT" + CStr(i), dbLong
        ThemTruong "ChungTuP", "MaDT" + CStr(i), dbLong
        ThemTruong "CPGVHD", "MaDT" + CStr(i), dbLong
    Next
    
    ThemTruong "HoaDon", "KCT", dbInteger
    ThemTruong "HoaDon", "HTTT", dbText, 20
    ThemTruong "HoaDon", "MauSo", dbText, 20
    ThemTruong "HoaDon", "TyGia", dbDouble
                
    If Not BangDaCo("PhanLoaiNhanVien") Then
        ThemTruong "ChungTu", "MaNV", dbLong
        
        CopyTable pCurDir + "UPDATE.MDB", "PhanLoaiNhanVien"
        CopyTable pCurDir + "UPDATE.MDB", "NhanVien"
        
        Set rel = DBKetoan.CreateRelation("PLNV", "PhanLoaiNhanVien", "NhanVien", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaPhanLoai"
        DBKetoan.Relations.Append rel
    End If
    
    ThemTruong "ChungTu", "HanTT", dbInteger
    ThemTruong "ChungTu", "SH1", dbText, 20
    If ThemTruong("ChungTu", "T1", dbInteger) Then LaySHCT
    
    If ThemTruong("ChungTu", "TLCK", dbDouble) Then
        ThemTruong "ChungTu", "CK", dbDouble
        XoaBang "ChungTuP"
    End If
    
    CopyTable2 "ChungTu", "ChungTuP"
    CopyTable2 "HoaDon", "HoaDonP"
    CopyTable2 "ChungTuLQ", "ChungTuLQP"
    
    ThemTruong "DoituongCT", "NgayTH", dbDate
    
    ThemTruong "License", "NgayDauThang", dbInteger
            
    If Not BangDaCo("CNDauNam") Then
        CopyTable pCurDir + "UPDATE.MDB", "CNDauNam"
        
        Set rel = DBKetoan.CreateRelation("CNDK2", "HethongTK", "CNDauNam", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaTaiKhoan"
        DBKetoan.Relations.Append rel
        Set rel = DBKetoan.CreateRelation("VTDK4", "KhachHang", "CNDauNam", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaKhachHang"
        DBKetoan.Relations.Append rel
    End If
    CopyTable2 "CNDauNam", "CNDauNam2"
    
    If Not BangDaCo("User2") Then CopyTable pCurDir + "UPDATE.MDB", "User2"
    
    ThemTruong "BaoCaoCP", "MK", dbLong
    CopyTable2 "BaoCaoCP", "BaoCaoCP2"
    
    SoDuTKCN
    
    If Not BangDaCo("CTGhiSo") Then
        CopyTable pCurDir + "UPDATE.MDB", "CTGhiSo"
        
        ExecuteSQL5 "INSERT INTO CTGhiSo (MaSo) VALUES (1)"
        ThemTruong "ChungTu", "CTGS", dbLong, , 1

        Set rel = DBKetoan.CreateRelation("CTGS", "CTGhiSo", "ChungTu", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "CTGS"
        DBKetoan.Relations.Append rel
    End If
        
    If pVersion <> 3 Then
        If Not BangDaCo("DinhMuc") Then
            CopyTable pCurDir + "UPDATE.MDB", "DinhMuc"
            
            Set rel = DBKetoan.CreateRelation("DM1", "Vattu", "DinhMuc", dbRelationDeleteCascade)
            rel.Fields.Append rel.CreateField("MaSo")
            rel.Fields!MaSo.ForeignName = "MaTP"
            DBKetoan.Relations.Append rel
        End If
            
        If Not ThemTruong("DinhMuc", "Thang", dbInteger, , 1) Then
            If pDinhmuc = 0 Then ExecuteSQL5 "UPDATE DinhMuc SET Thang=1"
        End If
        ThemTruong "DinhMuc", "ThangHL", dbInteger, , 0
        
        If Not BangDaCo("ThanhPham") Then
            CopyTable pCurDir + "UPDATE.MDB", "ThanhPham"
            
            Set rel = DBKetoan.CreateRelation("MK1", "KhoHang", "ThanhPham", dbRelationDeleteCascade)
            rel.Fields.Append rel.CreateField("MaSo")
            rel.Fields!MaSo.ForeignName = "MaKhoTP"
            DBKetoan.Relations.Append rel
            Set rel = DBKetoan.CreateRelation("MK2", "KhoHang", "ThanhPham", dbRelationDeleteCascade)
            rel.Fields.Append rel.CreateField("MaSo")
            rel.Fields!MaSo.ForeignName = "MaKhoNVL"
            DBKetoan.Relations.Append rel
            Set rel = DBKetoan.CreateRelation("MVT", "Vattu", "ThanhPham", dbRelationDeleteCascade)
            rel.Fields.Append rel.CreateField("MaSo")
            rel.Fields!MaSo.ForeignName = "MaVattu"
            DBKetoan.Relations.Append rel
        End If
    
        If TruongDaCo("ThanhPham", "MaTaiKhoan") Then
            Set tdf = DBKetoan.TableDefs!ThanhPham
            tdf.Fields("MaTaiKhoan").Name = "MaTK"
        End If
        
        If TruongDaCo("ThanhPham", "MaVattu") Then
            Set tdf = DBKetoan.TableDefs!ThanhPham
            tdf.Fields("MaVattu").Name = "MaTP"
        End If
        
        ThemTruong "ThanhPham", "MaTKCP", dbLong
        ThemTruong "ThanhPham", "MaTKNC", dbLong
        ThemTruong "ThanhPham", "MaTKSX", dbLong
        ThemTruong "ThanhPham", "Ma154", dbLong
        ThemTruong "ThanhPham", "GiaThanh", dbDouble
        ThemTruong "ThanhPham", "XK", dbInteger, , GetSetting(IniPath, "Stock", "AutoOut", 1)
        ThemTruong "ThanhPham", "CPTC", dbDouble
        
        If Not BangDaCo("PhanLoai154") Then CopyTable pCurDir + "UPDATE.MDB", "PhanLoai154"
        
        If Not BangDaCo("TP154") Then
            CopyTable pCurDir + "UPDATE.MDB", "TP154"
                    
            Set rel = DBKetoan.CreateRelation("PLTP", "PhanLoai154", "TP154", dbRelationDeleteCascade)
            rel.Fields.Append rel.CreateField("MaSo")
            rel.Fields!MaSo.ForeignName = "MaPhanLoai"
            DBKetoan.Relations.Append rel
        Else
            For i = 1 To 12
                ThemTruong "TP154", "CPKH_" + CStr(i), dbDouble
                ThemTruong "TP154", "CPSXC_" + CStr(i), dbDouble
                ThemTruong "TP154", "CPBH" + CStr(i), dbDouble
                ThemTruong "TP154", "CPQL" + CStr(i), dbDouble
                ThemTruong "TP154", "CPTC" + CStr(i), dbDouble
                ThemTruong "TP154", "CPBHTT" + CStr(i), dbDouble
                ThemTruong "TP154", "CPQLTT" + CStr(i), dbDouble
                ThemTruong "TP154", "CPTCTT" + CStr(i), dbDouble
                ThemTruong "TP154", "CPSXCTT" + CStr(i), dbDouble
                ThemTruong "TP154", "CPNVLPB" + CStr(i), dbDouble
                ThemTruong "TP154", "CPNCPB" + CStr(i), dbDouble
                ThemTruong "TP154", "CPMPB" + CStr(i), dbDouble
            Next
            ThemTruong "TP154", "DT", dbDouble
            ThemTruong "TP154", "CPTC", dbDouble
            ThemTruong "TP154", "MaTK", dbLong, , SelectSQL("SELECT TOP 1 MaSo AS F1 FROM HethongTK WHERE SoHieu LIKE '" + ShTkSPDo + "*' ORDER BY SoHieu")
            ThemTruong "TP154", "KPB", dbInteger
            ThemTruong "TP154", "SanLuong", dbDouble
        End If
    End If
    
    If Not BangDaCo("VTDauNam") And OutCost <> 0 Then
        CopyTable pCurDir + "UPDATE.MDB", "VTDauNam"
        
        Set rel = DBKetoan.CreateRelation("VTDK1", "KhoHang", "VTDauNam", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaSoKho"
        DBKetoan.Relations.Append rel
        Set rel = DBKetoan.CreateRelation("VTDK2", "HethongTK", "VTDauNam", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaTaiKhoan"
        DBKetoan.Relations.Append rel
        Set rel = DBKetoan.CreateRelation("VTDK3", "Vattu", "VTDauNam", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaVattu"
        DBKetoan.Relations.Append rel
        
        ExecuteSQL5 "INSERT INTO VTDauNam (MaSo,MaSoKho,MaTaiKhoan,MaVattu,Tien_0,Luong_0,SoXuat) SELECT MaSo,MaSoKho,MaTaiKhoan,MaVattu,Tien_0,Luong_0,SoXuat FROM TonKho WHERE Tien_0<>0 OR Luong_0<>0"
    End If
    
    If Not BangDaCo("KiemKe") Then
        CopyTable pCurDir + "UPDATE.MDB", "KiemKe"
        
        Set rel = DBKetoan.CreateRelation("VTKK1", "KhoHang", "KiemKe", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaSoKho"
        DBKetoan.Relations.Append rel
        Set rel = DBKetoan.CreateRelation("VTKK2", "HethongTK", "KiemKe", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaTaiKhoan"
        DBKetoan.Relations.Append rel
        Set rel = DBKetoan.CreateRelation("VTKK3", "Vattu", "KiemKe", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaVattu"
        DBKetoan.Relations.Append rel
    End If
    
    If Not BangDaCo("KiemKeN") Then
        CopyTable pCurDir + "UPDATE.MDB", "KiemKeN"
        
        Set rel = DBKetoan.CreateRelation("VTKKN1", "KhoHang", "KiemKeN", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaSoKho"
        DBKetoan.Relations.Append rel
        Set rel = DBKetoan.CreateRelation("VTKKN2", "HethongTK", "KiemKeN", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaTaiKhoan"
        DBKetoan.Relations.Append rel
        Set rel = DBKetoan.CreateRelation("VTKKN3", "Vattu", "KiemKeN", dbRelationDeleteCascade)
        rel.Fields.Append rel.CreateField("MaSo")
        rel.Fields!MaSo.ForeignName = "MaVattu"
        DBKetoan.Relations.Append rel
    End If
    
    CopyTable pCurDir + "UPDATE.MDB", "BKNhomPS"
    
    ThemTruong "CTGhiSo", "DienGiai", dbText, 50
    
    If ThemTruong("TaiSan", "NamKH", dbInteger) Then
        ExecuteSQL5 "UPDATE TaiSan INNER JOIN ThongSo ON TaiSan.MaSo=ThongSo.MaTS SET NamKH=FIX(0.9+(NG_NS+NG_TBS+NG_CNK+NG_TD)/(12*(KH_NS+KH_TBS+KH_CNK+KH_TD))) WHERE Thang=12 AND (KH_NS+KH_TBS+KH_CNK+KH_TD)>0"
    Else
        ExecuteSQL5 "UPDATE TaiSan SET NamKH = 0 WHERE IsNull(NamKH)"
    End If
    
    If ThemTruong("TaiSan", "SHCT", dbText, 20) Then
        ThemTruong "TaiSan", "NCT", dbDate
        ExecuteSQL5 "UPDATE TaiSan INNER JOIN CTTaiSan ON TaiSan.MaSo=CTTaiSan.MaTS SET SHCT=CTTaiSan.SoHieu, NCT=CTTaiSan.VaoSo WHERE CTTaiSan.MaLoai=32"
    End If
        
    If ThemTruong("HoaDon", "GiaTT", dbDouble) Then
        ExecuteSQL5 "UPDATE HethongTK SET TK_ID=3332 WHERE SoHieu LIKE '3332*'"
    End If
    
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID=1330 WHERE SoHieu LIKE '33312*'"
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID=" + CStr(TKGT_ID) + " WHERE SoHieu LIKE '521*' OR SoHieu LIKE '531*'"
    
    If pVersion = 3 Then GoTo E
    TK.InitTaikhoanSohieu "6234"
    If TK.MaSo = 0 Then
        TK.InitTaikhoanSohieu "623"
        If TK.MaSo > 0 Then
            TK.ThemTKCon "4", ABCtoVNI("Chi phÝ khÊu hao TSC§"), "Depreciation of fixed assets", 0, 3003
            ExecuteSQL5 "UPDATE HethongTK SET MaTC=MaSo WHERE Cap=2 AND SoHieu LIKE '623*'"
            ExecuteSQL5 "UPDATE HethongTK SET MaTC=0 WHERE Cap=1 AND SoHieu LIKE '623*'"
        End If
    End If
    TK.InitTaikhoanSohieu "621"
    If TK.MaSo = 0 Then
        ExecuteSQL5 "UPDATE Cdts SET MaTk2=82 WHERE MaSo=114"
        ExecuteSQL5 "UPDATE Cdts SET MaTk2=18 WHERE MaSo=312"
        ExecuteSQL5 "UPDATE Kqkd SET MaSoCha=30 WHERE MaSo=20"
        GoTo B
    End If
    TK.InitTaikhoanSohieu "515"
    If TK.MaSo > 0 Then GoTo B
    
    If MsgBox("Bæ sung c¸c tµi kho¶n 242, 515, 635 vµ ®iÒu chØnh mÉu b¸o c¸o Tµi chÝnh ?", vbYesNo + vbExclamation, App.ProductName) <> vbYes Then GoTo B
    
    ExecuteSQL5 "UPDATE HethongTK SET Ten='" + ABCtoVNI("Thu nhËp kh¸c") + "' WHERE MaSo=168 OR MaSo=169"
    ExecuteSQL5 "UPDATE HethongTK SET Ten='" + ABCtoVNI("Chi phÝ kh¸c") + "' WHERE MaSo=171 OR MaSo=172"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, TkCha1, MaTC,TK_ID) VALUES (9999,'2136',2,'" + ABCtoVNI("GiÊy phÐp vµ giÊy phÐp nh­îng quyÒn") + "',2,-1,57,47,9999,2000)"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, MaTC) VALUES (10000,'242',1,'" + ABCtoVNI("Chi phÝ tr¶ tr­íc dµi h¹n") + "',2,-1,47,10000)"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, MaTC) VALUES (10001,'515',1,'" + ABCtoVNI("Doanh thu ho¹t ®éng tµi chÝnh") + "',5,1,122,10001)"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, MaTC) VALUES (10002,'635',1,'" + ABCtoVNI("Chi phÝ tµi chÝnh") + "',6,-1,137,10002)"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TkCha1, MaTC) VALUES (10003,'63501',2,'" + ABCtoVNI("L·i vay ph¶i tr¶") + "',6,-1,10002,137,10002)"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, MaTC) VALUES (10004,'158',1,'" + ABCtoVNI("Hµng ho¸ kho b¶o thuÕ") + "',1,-1,1,10004)"
    ChuyenCTTKCap1 "711", "515"
    ChuyenCTTKCap1 "721", "711"
    ChuyenCTTKCap1 "811", "635"
    ChuyenCTTKCap1 "821", "811"
    KiemTraTaiKhoan
    ExecuteSQL5 "INSERT INTO Cdts (Maso,TaiSan,DienGiai,MaTK1,NhomCha) VALUES (241,-1,'  V. Chi phÝ tr¶ tr­íc dµi h¹n',10000,200)"
    ExecuteSQL5 "INSERT INTO Cdts (Maso,TaiSan,DienGiai,MaTK1,NhomCha) VALUES (148,-1,'    8. Hµng ho¸ kho b¶o thuÕ',10004,140)"
    ExecuteSQL5 "UPDATE Cdts SET DienGiai='    9. Dù phßng gi¶m gi¸ hµng tån kho (*)' WHERE MaSo=149"
B:
    ExecuteSQL5 "UPDATE HethongTK SET Ten='" + ABCtoVNI("Doanh thu ch­a thùc hiÖn") + "' WHERE SoHieu='3387'"
    TK.InitTaikhoanSohieu "337"
    If TK.MaSo > 0 Then GoTo c
    If MsgBox("Bæ sung c¸c tµi kho¶n 337, 343 vµ ®iÒu chØnh mÉu b¸o c¸o Tµi chÝnh ?", vbYesNo + vbExclamation, App.ProductName) <> vbYes Then GoTo c
    
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, TkCha1, MaTC) VALUES (12000,'337',1,'" + ABCtoVNI("Thanh to¸n theo tiÕn ®é kÕ ho¹ch hîp ®ång x©y dùng") + "',3,1,79,0,12000)"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, TkCha1, MaTC) VALUES (12001,'343',1,'" + ABCtoVNI("Tr¸i phiÕu ph¸t hµnh") + "',3,1,79,0,0)"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, TkCha1, MaTC) VALUES (12002,'3431',2,'" + ABCtoVNI("MÖnh gi¸ tr¸i phiÕu") + "',3,1,12001,79,12002)"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, TkCha1, MaTC) VALUES (12003,'3432',2,'" + ABCtoVNI("ChiÕt khÊu tr¸i phiÕu") + "',3,-1,12001,79,12003)"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, TkCha1, MaTC) VALUES (12004,'3433',2,'" + ABCtoVNI("Phô tréi tr¸i phiÕu") + "',3,1,12001,79,12004)"
    ExecuteSQL5 "DELETE * FROM Cdts WHERE MaSo=137 OR MaSo=319 OR MaSo=323"
    ExecuteSQL5 "INSERT INTO Cdts (Maso,TaiSan,DienGiai,MaTK1,MaTK2,NhomCha) VALUES (137,-1,'    5. Ph¶i thu theo tiÕn ®é kÕ ho¹ch hîp ®ång x©y dùng',12000,-1,130)"
    ExecuteSQL5 "UPDATE Cdts SET DienGiai='    6. C¸c kho¶n ph¶i thu kh¸c' WHERE MaSo=138"
    ExecuteSQL5 "UPDATE Cdts SET DienGiai='    7. Dù phßng c¸c kho¶n ph¶i thu khã ®ßi' WHERE MaSo=139"
    ExecuteSQL5 "INSERT INTO Cdts (Maso,TaiSan,DienGiai,MaTK1,MaTK2,NhomCha) VALUES (319,1,'    9. Ph¶i tr¶ theo tiÕn ®é kÕ ho¹ch hîp ®ång x©y dùng',12000,-1,310)"
    ExecuteSQL5 "INSERT INTO Cdts (Maso,TaiSan,DienGiai,MaTK1,MaTK2,NhomCha) VALUES (323,1,'    3. Tr¸i phiÕu ph¸t hµnh',12001,0,320)"
c:
    TK.InitTaikhoanSohieu "4131"
    If TK.MaSo > 0 Then GoTo d
    If MsgBox("Bæ sung c¸c tµi kho¶n 4131, 4132, 4133 ?", vbYesNo + vbExclamation, App.ProductName) <> vbYes Then GoTo d
    TK.InitTaikhoanSohieu "413"
    If TK.MaSo > 0 Then
        TK.ThemTKCon "1", ABCtoVNI("Chªnh lÖch tû gi¸ ®¸nh gi¸ l¹i cuèi n¨m tµi chÝnh"), "...", 0, 0
        TK.ThemTKCon "2", ABCtoVNI("Chªnh lÖch tû gi¸ trong giai ®o¹n ®Çu t­ c¬ b¶n"), "...", 0, 0
        TK.ThemTKCon "3", ABCtoVNI("Chªnh lÖch tû gi¸ tõ chuyÓn ®æi b¸o c¸o tµi chÝnh"), "...", 0, 0
        
        ExecuteSQL5 "UPDATE HethongTK SET MaTC=MaSo WHERE Cap=2 AND SoHieu LIKE '413*'"
        ExecuteSQL5 "UPDATE HethongTK SET MaTC=0 WHERE Cap=1 AND SoHieu LIKE '413*'"
    End If
d:
    If frmMain.Lb(0).tag < 3 Then
        TK.InitTaikhoanSohieu "1385"
        If TK.MaSo > 0 Then GoTo E
        If MsgBox("Bæ sung c¸c tµi kho¶n 1385, 3385, 417, 419 ?", vbYesNo + vbExclamation, App.ProductName) <> vbYes Then GoTo E
        TK.InitTaikhoanSohieu "138"
        If TK.MaSo > 0 Then TK.ThemTKCon "5", ABCtoVNI("Ph¶i thu vÒ cæ phÇn ho¸"), "...", 0, 0
        TK.InitTaikhoanSohieu "338"
        If TK.MaSo > 0 Then TK.ThemTKCon "5", ABCtoVNI("Ph¶i tr¶ vÒ cæ phÇn ho¸"), "...", 0, 0
        ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, TkCha1, MaTC) VALUES (13000,'417',1,'" + ABCtoVNI("Quü hç trî s¾p xÕp vµ cæ phÇn ho¸ DNNN") + "',4,1,105,0,13000)"
        ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, TkCha1, MaTC) VALUES (13001,'419',1,'" + ABCtoVNI("Cæ phiÕu mua l¹i") + "',4,1,105,0,13001)"
        ExecuteSQL5 "INSERT INTO Cdts (Maso,TaiSan,DienGiai,MaTK1,MaTK2,NhomCha) VALUES (428,1,'    6. Quü hç trî s¾p xÕp vµ cæ phÇn ho¸ DNNN',13000,0,420)"
        ExecuteSQL5 "INSERT INTO Cdts (Maso,TaiSan,DienGiai,MaTK1,MaTK2,NhomCha) VALUES (429,1,'    7. Cæ phiÕu mua l¹i',13001,0,420)"
    End If
E:
    TK.InitTaikhoanSohieu "217"
    If TK.MaSo > 0 Then GoTo f
    
    If MsgBox("Bæ sung c¸c tµi kho¶n 217, 2147, 5117, 1567, 223 vµ ®iÒu chØnh mÉu b¸o c¸o Tµi chÝnh ?", vbYesNo + vbExclamation, App.ProductName) <> vbYes Then GoTo f
   
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, MaTC,TK_ID) VALUES (14000,'217',1,'" + ABCtoVNI("BÊt ®éng s¶n ®Çu t­") + "',2,-1,47,14000," + CStr(TSCD_ID) + ")"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, TkCha1, MaTC,TK_ID) VALUES (14001,'2147',2,'" + ABCtoVNI("Hao mßn bÊt ®éng s¶n ®Çu t­") + "',2,1,64,47,14001," + CStr(KHTSCD_ID) + ")"
   ' ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, TkCha1, MaTC, TK_ID) VALUES (14002,'5117',2,'" + ABCtoVNI("Doanh thu kinh doanh bÊt ®éng s¶n ®Çu t­") + "',5,123,122,14002," + CStr(KHTSCD_ID) + ")"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TkCha1, MaTC) VALUES (14003,'1567',2,'" + ABCtoVNI("Hµng hãa bÊt ®éng s¶n") + "',1,-1,39,1,14003)"
    ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, MaTC) VALUES (14004,'223',1,'" + ABCtoVNI("§Çu t­ vµo c«ng ty liªn kÕt") + "',2,-1,47,14004)"
    ExecuteSQL5 "INSERT INTO LoaiChungTu (MaSo,SoHieu,Cap,Ten,CapTren) VALUES (101,'KHAO-04',2,'" + ABCtoVNI("TrÝch khÊu hao BÊt ®éng s¶n ®Çu t­") + "',35)"
    ExecuteSQL5 "INSERT INTO LoaiTaiSan (MaSo,SoHieu,Cap,Ten,CapTren) VALUES (1001,'217',1,'" + ABCtoVNI("BÊt ®éng s¶n ®Çu t­") + "',0)"
    ExecuteSQL5 "UPDATE HethongTK SET Ten='" + ABCtoVNI("Vèn gãp liªn doanh") + "' WHERE SoHieu='222'"
    ExecuteSQL5 "UPDATE HethongTK SET Ten='" + ABCtoVNI("§Çu t­ vµo c«ng ty con") + "' WHERE SoHieu='221'"
        'cap nhat QD 15/2006/QD-BTC
f:
    TK.InitTaikhoanSohieu "1281"
    If TK.MaSo > 0 Then GoTo k
     If MsgBox("Bæ sung theo Q§ Sè 15/2006/Q§-BTC,B¹n muèn cËp nhËp kh«ng?", vbYesNo + vbExclamation, App.ProductName) <> vbYes Then GoTo m
m:
   frmXuly.Show
     UpdateAcount
   GoTo k
k:
    dieuchinh
    SetDefaultValue
    
    Set TK = Nothing
    Set tdf = Nothing
    Set rel = Nothing
    
    pSongNgu = s
    ExecuteSQL5 "UPDATE License SET MV=(MV MOD 10000)+10000,MKUP=" + CStr(pRev) + ",TenCty_ID = " + CStr(Int_StrToCode(pTenCty)) + ",TenCn_ID = " + CStr(Int_StrToCode(pTenCn)) + ",MST_ID = " + CStr(Int_StrToCode(frmMain.LbCty(8).Caption))
    frmXuly.Gau.Max = 21
    frmXuly.Hide
      MsgBox "CËp nhËt kÕt thóc!", vbInformation, App.ProductName
End Sub

Public Sub XoaCTTheoID(thang As Integer, id As Long, ml As Integer, Optional ndau As Date, Optional ncuoi As Date, Optional ml2 As Integer = 0)
    Dim rs As Recordset, oct As New ClsChungtu
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo FROM ChungTu WHERE CT_ID=" + CStr(id) + " AND " _
        + IIf(thang > 0, "ThangCT=" + CStr(thang), WNgay("NgayGS", ndau, ncuoi)) _
        + IIf(ml > 0, " AND MaLoai=" + CStr(ml), "") + IIf(ml2 > 0, " AND MaLoai<>" + CStr(ml2), ""), dbOpenSnapshot)
    Do While Not rs.EOF
        oct.InitChungtu rs!MaSo, 0, "", 0, Date, Date, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
        oct.XoaChungtu
        rs.MoveNext
    Loop
    
    ExecuteSQL5 "DELETE ChungTuLQ.* FROM ChungTuLQ LEFT JOIN ChungTu ON ChungTuLQ.MaCT=ChungTu.MaCT WHERE IsNull(ChungTu.MaCT)"
    ExecuteSQL5 "DELETE CTTaiSan.* FROM CTTaiSan LEFT JOIN ChungTu ON CTTaiSan.MaCTKT=ChungTu.MaCT WHERE IsNull(ChungTu.MaCT) AND CTTaiSan.Thang>0"
    ExecuteSQL5 "DELETE HoaDon.* FROM HoaDon LEFT JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo WHERE IsNull(ChungTu.MaSo)"
    ExecuteSQL5 "DELETE ThongSo.* FROM ThongSo LEFT JOIN CTTaiSan ON ThongSo.MaTS= CTTaiSan.MaTS WHERE IsNull(CTTaiSan.MaTS)"
    ExecuteSQL5 "DELETE TaiSan.* FROM TaiSan LEFT JOIN CTTaiSan ON TaiSan.MaSo= CTTaiSan.MaTS WHERE IsNull(CTTaiSan.MaTS)"
    
    Set oct = Nothing
    rs.Close
    Set rs = Nothing
End Sub
Public Sub UpdateAcount()
  Dim tdf As TableDef, rel As Relation, TK As New ClsTaikhoan
    Dim i As Integer, sql As String, j As Integer, s As Boolean
     frmXuly.Gau.Max = 21
   'hieu chinh theo thong tu 15
  'loai tai khoan ngan hang loai 1
     'Tim 128 xem co TK con chua
      Dim rs As DAO.Recordset
       
      'Them Tkcon cua TK 128==========================================
       Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '128*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
        'update tk 128
          ExecuteSQL5 "UPDATE HethongTK SET TKcon=1 WHERE SoHieu='128'"
        'them con
         frmXuly.Gau.Value = 1
        TK.InitTaikhoanSohieu "1281"
        If TK.MaSo > 0 Then GoTo t2
          ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, MaTC) VALUES (14028,'1281',2,'" + ABCtoVNI("TiÒn göi cã kú h¹n") + "',1,-1,16,14028)"
t2:
       frmXuly.Gau.Value = 2
         'Them tkcon 511 la 5117 ==============================================================
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '511*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        
        If rs.recordCount <> 0 Then
          TK.InitTaikhoanSohieu "5117"
            If TK.MaSo > 0 Then
            ExecuteSQL5 "Update HethongTK set TKcon=1, cap=2, TKCha0=123,TKcha1=122 where sohieu='5117'"
             End If
            ExecuteSQL5 "Update HethongTK set ten='Doanh thu trôï caáp, trôï giaù' where sohieu='5114'"
        End If

T:
  frmXuly.Gau.Value = 3
       TK.InitTaikhoanSohieu "1288"
       If TK.MaSo > 0 Then GoTo h
          ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, MaTC) VALUES (14029,'1288',2,'" + ABCtoVNI("§Çu t­ ng¾n h¹n kh¸c") + "',1,-1,16,14029)"
        End If
h:
       'thuc hien 138
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '138*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
         Loop
       If rs.recordCount > 0 Then
        TK.InitTaikhoanSohieu "1385"
        If TK.MaSo > 0 Then GoTo h1
        ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,Tkcha1, MaTC,TK_ID2) VALUES (14007,'1385',2,'" + ABCtoVNI("Ph¶i thu vÒ cæ phÇn ho¸") + "',1,-1,22,1,14007,1310)"
       End If
       frmXuly.Gau.Value = 4
      'thuc hien 142 giam TK con
h1:
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '142*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        frmXuly.Gau.Value = 5
        If rs.recordCount <> 0 Then
        TK.InitTaikhoanSohieu "1421"
          If TK.MaSo > 0 Then
           'ExecuteSQL5 "Delete * from HethongTK where Sohieu= '1421'"
          End If
        TK.InitTaikhoanSohieu "1422"
          If TK.MaSo > 0 Then
          ' ExecuteSQL5 "Delete * from HethongTK where Sohieu= '1422'"
          End If
         ExecuteSQL5 "Update HethongTK set TKcon=1 where Sohieu='142'"
       End If
      'thuc hien 153=================================================================
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '153*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        frmXuly.Gau.Value = 6
        If rs.recordCount <> 0 Then
        TK.InitTaikhoanSohieu "1531"
         If TK.MaSo > 0 Then
           ExecuteSQL5 "Delete * from HethongTK where Sohieu= '1531'"
         End If
        TK.InitTaikhoanSohieu "1532"
         If TK.MaSo > 0 Then
          ExecuteSQL5 "Delete * from HethongTK where Sohieu= '1532'"
         End If
        TK.InitTaikhoanSohieu "1533"
         If TK.MaSo > 0 Then
          ExecuteSQL5 "Delete * from HethongTK where Sohieu= '1533'"
         End If
         ExecuteSQL5 "Update HethongTK set TKcon=0 where Sohieu='153'"
       End If
       frmXuly.Gau.Value = 7
       'thuc hien them tai khoan moi 158============================================
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '158*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
         Loop
         If rs.recordCount = 0 Then
          TK.InitTaikhoanSohieu "158"
         If TK.MaSo > 0 Then GoTo h2
          ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0, MaTC) VALUES (14008,'158',1,'" + ABCtoVNI("Hµng ho¸ kho b¶o thuÕ") + "',1,-1,1,14008)"
         End If
        'loai 2
       'thuc hien TK 221 giam TKcon==================================================
h2:
 
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '221*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
         If rs.recordCount <> 0 Then
         'do update sub  acount of acount to sub acount
          TK.InitTaikhoanSohieu "2211*"
         If TK.MaSo > 0 Then
          ExecuteSQL5 "Update HethongTK set TKcon=2 where Sohieu='2211*'"
          End If
         TK.InitTaikhoanSohieu "2211"
         If TK.MaSo > 0 Then
          ExecuteSQL5 "Delete * from HethongTK where Sohieu='2211'"
         End If
         TK.InitTaikhoanSohieu "2212"
         If TK.MaSo > 0 Then
          ExecuteSQL5 "Delete * from HethongTK where Sohieu='2212'"
         End If
          ExecuteSQL5 "Update HethongTK set TKcon=1 where Sohieu='153'"
        End If
        'thuc hien them tk 228==========================================================
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '228*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        frmXuly.Gau.Value = 8
         If rs.recordCount <> 0 Then
         TK.InitTaikhoanSohieu "2281"
         If TK.MaSo > 0 Then GoTo H3
           ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC) VALUES (14009,'2281',2,'" + ABCtoVNI("Cå phiÕu") + "',2,-1,72,47,14009)"
H3:
         TK.InitTaikhoanSohieu "2282"
         If TK.MaSo > 0 Then GoTo H4
           ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC) VALUES (14010,'2282',2,'" + ABCtoVNI("Tr¸i phiÕu") + "',2,-1,72,47,14010)"
H4:
         TK.InitTaikhoanSohieu "2288"
         If TK.MaSo > 0 Then GoTo H5
           ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC) VALUES (14011,'2288',2,'" + ABCtoVNI("§Çu t­ dµi h¹n kh¸c") + "',2,-1,72,47,14011)"
           'ExecuteSQL5 "Update HeThongTK set TKcon = 1, TKcha=47 where Sohieu ='228'"
           ExecuteSQL5 "Update HethongTK set TKcon= 1, TKcha0=47 where Sohieu='228'"
         End If
         frmXuly.Gau.Value = 9
        'thuc hieu them moi tK 243
H5:
 
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '243*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
         If rs.recordCount < 0 Then
           TK.InitTaikhoanSohieu "243"
            If TK.MaSo > 0 Then GoTo H6
             ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC) VALUES (14012,'243',1,'" + ABCtoVNI("Tµi s¶n thuÕ thu nhËp ho¶n l¹i") + "',2,-1,47,0,14012)"
         End If
         frmXuly.Gau.Value = 10
         'thuc hien TK loai 3 1-3
         'giam 331
H6:
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '331*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
        TK.InitTaikhoanSohieu "3311"
           If TK.MaSo > 0 Then
            ExecuteSQL5 "Delete * from HethongTK where Sohieu='3311'"
           End If
         TK.InitTaikhoanSohieu "3312"
           If TK.MaSo > 0 Then
            ExecuteSQL5 "Delete * from HethongTK where Sohieu='3312'"
           End If
        
          ExecuteSQL5 "Update HethongTK set TKcon=1 where sohieu='331'"
        End If
        'thuc hien them tk 334
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '334*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
       If rs.recordCount <> 0 Then
        TK.InitTaikhoanSohieu "3348"
           If TK.MaSo > 0 Then GoTo H7
            ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC,TK_ID,TK_ID2) VALUES (14013,'3348',2,'" + ABCtoVNI("Ph¶i tr¶ ng­êi lao ®éng kh¸c") + "',3,1,82,79,14013,3310,1310)"
            ExecuteSQL5 "Update HethongTK set Ten='" + ABCtoVNI("Ph¶i tr¶ ng­êi lao ®éng") + "'where sohieu='334'"
          End If
          frmXuly.Gau.Value = 11
H7:
        'thu hien TK 338
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '338*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount > 0 Then
         TK.InitTaikhoanSohieu "3385"
           If TK.MaSo > 0 Then GoTo H8
            ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC,TK_ID2) VALUES (14014,'3385',2,'" + ABCtoVNI("Ph¶i tr¶ vÒ cæ phÇn ho¸") + "',3,1,96,79,14014,1310)"
           End If
H8:
        TK.InitTaikhoanSohieu "3386"
           If TK.MaSo > 0 Then GoTo H9
             ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC,TK_ID2) VALUES (14015,'3386',2,'" + ABCtoVNI("NhËt ký quü, ký c­îc ng¾n h¹n") + "',3,1,96,79,14015,1310)"
        
H9:
         'them tK 347,351,352
          Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '347*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount = 0 Then
        TK.InitTaikhoanSohieu "347"
           If TK.MaSo > 0 Then GoTo H10
          ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC) VALUES (14016,'347',1,'" + ABCtoVNI("ThuÕ thu nhËp ho¶n l¹i ph¶i tr¶") + "',3,1,79,0,14016)"
        End If
        frmXuly.Gau.Value = 12
H10:
          Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '351*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount = 0 Then
         TK.InitTaikhoanSohieu "351"
           If TK.MaSo > 0 Then GoTo H11
          ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC) VALUES (14017,'351',1,'" + ABCtoVNI("Quü dù phßng trî cÊp mÊt viÖc") + "',3,1,79,0,14017)"
        End If
H11:
          Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '352*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount = 0 Then
          TK.InitTaikhoanSohieu "352"
           If TK.MaSo > 0 Then GoTo H12
          ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC) VALUES (14018,'352',1,'" + ABCtoVNI("Dù phßng ph¶i tr¶") + "',3,1,79,0,14018)"
         End If
H12:
        'loai 4
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '411*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
          TK.InitTaikhoanSohieu "4118"
           If TK.MaSo > 0 Then GoTo H13
             ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC) VALUES (14019,'4118',2,'" + ABCtoVNI("Vèn kh¸c") + "',4,1,117,105,14019)"
        End If
        frmXuly.Gau.Value = 13
H13:
      ' Thuc hien tk411 giam TKcon cap 3, sua ten tai khoan cap 2
      Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '411*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
          TK.InitTaikhoanSohieu "411101"
           If TK.MaSo > 0 Then
               ExecuteSQL5 "Delete * from HethongTK where Sohieu='411101'"
           End If
             TK.InitTaikhoanSohieu "411102"
           If TK.MaSo > 0 Then
               ExecuteSQL5 "Delete * from HethongTK where Sohieu='411102'"
               
           End If
             TK.InitTaikhoanSohieu "411201"
           If TK.MaSo > 0 Then
               ExecuteSQL5 "Delete * from HethongTK where Sohieu='411201'"
           End If
             TK.InitTaikhoanSohieu "411202"
           If TK.MaSo > 0 Then
             ExecuteSQL5 "Delete * from HethongTK where Sohieu='411202'"
           End If
             TK.InitTaikhoanSohieu "4111"
           If TK.MaSo > 0 Then
             ExecuteSQL5 "Update HethongTK set ten='Voán ñaàu tö cuûa chuû sôû höõu', MaTC =0  where sohieu='4111'"
           End If
           TK.InitTaikhoanSohieu "4112"
           If TK.MaSo > 0 Then
             ExecuteSQL5 "Update HethongTK set ten='Thaëng dö voán coå phaàn', MaTC = 0  where sohieu='4112'"
           End If
        End If
      frmXuly.Gau.Value = 14
      ' thuchien TK 412 giam TKcon
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '412*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
          TK.InitTaikhoanSohieu "4121"
           If TK.MaSo > 0 Then
               ExecuteSQL5 "Delete * from HethongTK where Sohieu='4121'"
           End If
           TK.InitTaikhoanSohieu "4122"
           If TK.MaSo > 0 Then
               ExecuteSQL5 "Delete * from HethongTK where Sohieu='4122'"
           End If
             ExecuteSQL5 "Update HethongTK set TKcon=1 where sohieu='412'"
        End If
        'thuc  hien 413================================================
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '413*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
          TK.InitTaikhoanSohieu "4133"
           If TK.MaSo > 0 Then
               ExecuteSQL5 "Delete * from HethongTK where Sohieu='4133'"
           End If
        End If
    'thuc hien  414============================================================
     Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '414*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
           TK.InitTaikhoanSohieu "4141"
           If TK.MaSo > 0 Then
               ExecuteSQL5 "Delete * from HethongTK where Sohieu='4141'"
           End If
           TK.InitTaikhoanSohieu "4142"
           If TK.MaSo > 0 Then
               ExecuteSQL5 "Delete * from HethongTK where Sohieu='4142'"
           End If
           ExecuteSQL5 "Update HethongTK set TKcon=1 where sohieu='414'"
        End If
        frmXuly.Gau.Value = 15
    'thuc hien Xoa 416, 417 =======================================================================
       Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '416*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
           TK.InitTaikhoanSohieu "416"
           If TK.MaSo > 0 Then
               ExecuteSQL5 "Delete * from HethongTK where Sohieu='416'"
           End If
           TK.InitTaikhoanSohieu "417"
           If TK.MaSo > 0 Then
               ExecuteSQL5 "Delete * from HethongTK where Sohieu='417'"
           End If
        End If
    'thuc hien Them moi 418 =========================================================
       Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '418*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount = 0 Then
            TK.InitTaikhoanSohieu "418"
           If TK.MaSo > 0 Then GoTo H14
            ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC) VALUES (14020,'418',1,'" + ABCtoVNI("C¸c quü kh¸c thuéc vèn chñ së h÷u") + "',4,1,0,0,14020)"
         End If
         frmXuly.Gau.Value = 16
H14:
    'thuc hien them moi TK 419=============================================================
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '419*'", dbOpenSnapshot)
          Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount = 0 Then
            TK.InitTaikhoanSohieu "419"
           If TK.MaSo > 0 Then GoTo H15
            ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TKCha1, MaTC) VALUES (14021,'419',1,'" + ABCtoVNI("Cæ phiÕu quü") + "',4,1,0,0,14021)"
         End If
    
    ' Thuc hien doi ten tai khoan cap 2 cua tk 421 ==================================
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '421*'", dbOpenSnapshot)
          Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount > 0 Then
            TK.InitTaikhoanSohieu "4211"
            ExecuteSQL5 "Update HethongTK set ten='Lôïi nhuaän chöa phaân phoái naêm tröôùc'  where sohieu='4211'"
            ExecuteSQL5 "Update HethongTK set ten='Lôïi nhuaän chöa phaân phoái naêm nay'  where sohieu='4212'"
        End If
        frmXuly.Gau.Value = 17
H15:
     'thuc hien 431=======================================================
       Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '431*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
            TK.InitTaikhoanSohieu "4314"
           If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='4314'"
           End If
        End If
    'thuc hien 441 giam con=====================================
       Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '441*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
            TK.InitTaikhoanSohieu "4411"
           If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='4411'"
           End If
            TK.InitTaikhoanSohieu "4412"
           If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='4412'"
           End If
           ExecuteSQL5 "Update HethongTK set TKcon=1 where sohieu='441'"
        End If
     frmXuly.Gau.Value = 18
     'thuc hien TK 451 xoa
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '451*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
            TK.InitTaikhoanSohieu "451"
            If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='451'"
           End If
        End If
        'thuc hien loai 5 ========================================================
        'thuc hien xoa Tkcon 512=================================
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '521*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
         Loop
          If rs.recordCount <> 0 Then
            TK.InitTaikhoanSohieu "5211"
            If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='5211'"
           End If
            TK.InitTaikhoanSohieu "5212"
            If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='5212'"
           End If
            TK.InitTaikhoanSohieu "5213"
            If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='5213'"
           End If
             ExecuteSQL5 "Update HethongTK set TKcon=1 where sohieu='521'"
        End If
        '515============
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '515*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
         Loop
       
       frmXuly.Gau.Value = 19
        
        'thuchien TK loai 6
        'thuc hien 623==========================================================
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '623*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
          TK.InitTaikhoanSohieu "6237"
            If TK.MaSo > 0 Then GoTo H16
            ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TK_ID,TKCha1, MaTC) VALUES (14022,'6237',2,'" + ABCtoVNI("Chi phÝ dÞch vô mua ngoµi") + "',6,-1,5126,3010,137,14022)"
H16:
            TK.InitTaikhoanSohieu "6238"
            If TK.MaSo > 0 Then GoTo H17
            ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TK_ID,TKCha1, MaTC) VALUES (14023,'6238',2,'" + ABCtoVNI("Chi phÝ b»ng tiÒn kh¸c") + "',6,-1,5126,3010,137,14023)"
       End If
H17:
        'thu hien 635 xoa conTK
          Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '635*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount <> 0 Then
           TK.InitTaikhoanSohieu "63501"
            If TK.MaSo > 0 Then
            '  ExecuteSQL5 "Delete * from HethongTK where Sohieu='63501'"
           End If
          ExecuteSQL5 "Update HethongTK set TKcon=1 where sohieu='635'"
        End If
        'thuc hien TK loai 8
          Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '821*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount = 0 Then
           TK.InitTaikhoanSohieu "821"
            If TK.MaSo > 0 Then GoTo h18
             ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TK_ID,TKCha1, MaTC) VALUES (14024,'821',1,'" + ABCtoVNI("Chi phÝ thuÕ thu nhËp doanh nghiÖp") + "',8,-1,0,0,0,1024)"
        End If
        frmXuly.Gau.Value = 20
h18:
     'thieun hien tK con
        Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '821*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount > 0 Then
           TK.InitTaikhoanSohieu "8211"
            If TK.MaSo > 0 Then GoTo H19
             ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TK_ID,TKCha1, MaTC) VALUES (14025,'8211',2,'" + ABCtoVNI("Chi phÝ thuÕ TNDN hiÖn hµnh") + "',8,-1,0,0,0,14025)"
H19:
            TK.InitTaikhoanSohieu "8212"
            If TK.MaSo > 0 Then GoTo h20
             ExecuteSQL5 "INSERT INTO HethongTK (MaSo,SoHieu,Cap,Ten,Loai,Kieu,TkCha0,TK_ID,TKCha1, MaTC) VALUES (14026,'8212',2,'" + ABCtoVNI("Chi phÝ thuÕ TNDN ho·n l¹i") + "',8,-1,0,0,0,14026)"

        End If
h20:
        'thuc hien 008=======================================
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '008*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount > 0 Then
           TK.InitTaikhoanSohieu "0081"
            If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='0081'"
           End If
           TK.InitTaikhoanSohieu "0082"
            If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='0082'"
              
           End If
            ExecuteSQL5 "Update HethongTK set TKcon=1, Ten='" + ABCtoVNI("Dù to¸n chi sù nghiÖp, dù ¸n") + "' where sohieu='008'"
        End If
        'thuc hien 007=====================
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '007*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        frmXuly.Gau.Value = 21
        If rs.recordCount > 0 Then
           TK.InitTaikhoanSohieu "00701"
            If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='00701'"
           End If
            ExecuteSQL5 "Update HethongTK set TKcon=1 where sohieu='007'"
        End If
        ' thuchien bo 009
         Set rs = DBKetoan.OpenRecordset("SELECT SoHieu as F1 FROM HeThongTK WHERE SoHieu Like '009*'", dbOpenSnapshot)
         Do While Not rs.EOF
          rs.MoveNext
        Loop
        If rs.recordCount > 0 Then
            TK.InitTaikhoanSohieu "009"
            If TK.MaSo > 0 Then
              ExecuteSQL5 "Delete * from HethongTK where Sohieu='009'"
           End If
        End If
        
End Sub
Public Function XoaCTOK(MaCT As Long) As Boolean
    Dim rs As Recordset, tien As Double, tien2 As Double
    
    XoaCTOK = True
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW * FROM ChungTu WHERE SoPS2No>0 AND MaLoai=1 AND MaVattu>0 AND MaCT=" + CStr(MaCT), dbOpenSnapshot)
    Do While Not rs.EOF
        If SoTonKho(rs!ThangCT, rs!MaKho, rs!MaTkNo, rs!MaVattu, tien, tien2) = 0 Then
            XoaCTOK = False
            GoTo KT
        End If
        rs.MoveNext
    Loop
KT:
    rs.Close
    Set rs = Nothing
End Function

Public Function KtraMKAdmin() As Boolean
     Dim mk As Long, pstr_psw As String
    
    mk = SelectSQL("SELECT PSW AS F1 FROM Users WHERE MaSo=1") - pNamTC
    If mk = 0 Then
        KtraMKAdmin = True
        Exit Function
    End If
    pstr_psw = FPsw.GetPswX()
    KtraMKAdmin = (Int_StrToCode(pstr_psw) = mk)
End Function

Private Sub ChuyenCTTKCap1(sh1 As String, sh2 As String)
    Dim TK As New ClsTaikhoan
    TK.InitTaikhoanSohieu sh1
    If TK.MaSo = 0 Or TK.tkcon = 0 Then Exit Sub
    TK.InitTaikhoanSohieu sh2
    If TK.MaSo = 0 Then Exit Sub
    ExecuteSQL5 "UPDATE HethongTK SET TKCon=1 WHERE SoHieu='" + sh2 + "'"
    ExecuteSQL5 "UPDATE HethongTK SET SoHieu='" + sh2 + "'+RIGHT(SoHieu,Len(Sohieu)-3),Loai=" + CStr(TK.loai) + ",MaTC=" + CStr(TK.MaTC) + " WHERE Cap>1 AND SoHieu LIKE '" + sh1 + "*'"
    ExecuteSQL5 "UPDATE HethongTK SET TKCha0=" + CStr(TK.MaSo) + ",TKCha1=" + CStr(TK.TkCha0) + " WHERE Cap=2 AND SoHieu LIKE '" + sh2 + "*'"
    ExecuteSQL5 "UPDATE HethongTK SET TKCha1=" + CStr(TK.MaSo) + ",TKCha2=" + CStr(TK.TkCha0) + " WHERE Cap=3 AND SoHieu LIKE '" + sh2 + "*'"
    ExecuteSQL5 "UPDATE HethongTK SET TKCha2=" + CStr(TK.MaSo) + ",TKCha3=" + CStr(TK.TkCha0) + " WHERE Cap=4 AND SoHieu LIKE '" + sh2 + "*'"
    ExecuteSQL5 "UPDATE HethongTK SET TKCha3=" + CStr(TK.MaSo) + ",TKCha4=" + CStr(TK.TkCha0) + " WHERE Cap=5 AND SoHieu LIKE '" + sh2 + "*'"
    ExecuteSQL5 "UPDATE HethongTK SET TKCha4=" + CStr(TK.MaSo) + ",TKCha5=" + CStr(TK.TkCha0) + " WHERE Cap=6 AND SoHieu LIKE '" + sh2 + "*'"
    Set TK = Nothing
End Sub

Public Function ST2MaSo(f As String, sh As String, tbl As String) As Long
    ST2MaSo = SelectSQL("SELECT MaSo AS F1 FROM " + tbl + " WHERE " + f + "='" + sh + "'")
End Function

Public Function SoHieu2MaSo(sh As String, tbl As String) As Long
    SoHieu2MaSo = SelectSQL("SELECT MaSo AS F1 FROM " + tbl + " WHERE SoHieu='" + sh + "'")
    If SoHieu2MaSo = 0 Then SoHieu2MaSo = SelectSQL("SELECT MaSo AS F1 FROM " + tbl + " WHERE SoHieu LIKE '" + sh + "*'")
End Function

Public Function MaSo2SoHieu(ms As Long, tbl As String) As String
    MaSo2SoHieu = SelectSQL("SELECT SoHieu AS F1 FROM " + tbl + " WHERE MaSo=" + CStr(ms))
End Function

Public Sub SetDefaultValue()
    Dim i As Integer, j As Integer
    
    On Error Resume Next
    For i = 0 To DBKetoan.TableDefs.count - 1
        For j = 0 To DBKetoan.TableDefs(i).Fields.count - 1
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

Public Sub PhanBoCP3(tdau As Integer, tcuoi As Integer, tc As Integer, shtk As String)
    Dim sql As String, tongcp As Double, sopb As Integer
    Dim i As Integer, tongpb As Double, rsct As Recordset, CP As Double, cp2 As String, f As String
    
    Select Case shtk
        Case "621": f = "CPNVLPB"
        Case "622": f = "CPNCPB"
        Case "623": f = "CPMPB"
        Case Else: Exit Sub
    End Select
    InChiPhi3 tdau, tcuoi, 0
    ExecuteSQL5 "DELETE FROM BaoCaoCP"
    For i = tdau To tcuoi
        cp2 = cp2 + f + CStr(i)
     Next
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,Kq1,Kq2,Kq4,Kq6,Kq7) SELECT DISTINCTROW MaSo,SoHieu,CPNVL,CPNC,DT," + cp2 + " AS SXC,CPM FROM TP154 WHERE KPB=0"                 ' WHERE CPNVL<>0 OR CPNC<>0 OR CPM<>0 OR DT<>0"
    XDTyLePB tc
    
    sql = "SELECT Count(MaSo) AS F1 FROM BaoCaoCP"
    sopb = SelectSQL(sql)
       
    CP = SelectSQL("SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '" + shtk + "*' AND " + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=0")
    
    For i = CThangDB(tdau) To CThangDB(ThangTruoc(tcuoi))
        ExecuteSQL5 "UPDATE TP154 SET " + f + CStr(i) + "=0"
    Next
        
    i = 0
    Set rsct = DBKetoan.OpenRecordset("SELECT * FROM BaoCaoCP ORDER BY Kq3", dbOpenSnapshot)
    Do While Not rsct.EOF
        i = i + 1
        If i < sopb Then
            tongcp = RoundMoney(CP * rsct!Kq3)
            tongpb = tongpb + tongcp
        Else
            tongcp = CP - tongpb
        End If
        ExecuteSQL5 "UPDATE TP154 SET " + f + CStr(tcuoi) + "=" + DoiDau(tongcp) + " WHERE MaSo=" + CStr(rsct!MaSo)
        rsct.MoveNext
    Loop
    rsct.Close
    Set rsct = Nothing
End Sub

Public Sub PhanBoCP64(tdau As Integer, tcuoi As Integer, tc As Integer, shtk As String)
    Dim sql As String, tongcp As Double, sopb As Integer
    Dim i As Integer, tongpb As Double, rsct As Recordset, CP As Double, sxc As String, f As String
    Dim s621 As String, s622 As String, s623 As String
    
    Select Case shtk
        Case "641":     f = "CPBH"
        Case "642":     f = "CPQL"
        Case "635":     f = "CPTC"
    End Select
    
    InChiPhi3 tdau, tcuoi, 0
    ExecuteSQL5 "DELETE FROM BaoCaoCP"
    For i = tdau To tcuoi
        sxc = sxc + "+CPSXC_" + CStr(i) + "+CPSXCTT" + CStr(i)
        s621 = s621 + "+CPNVLPB" + CStr(i)
        s622 = s622 + "+CPNCPB" + CStr(i)
        s623 = s623 + "+CPMPB" + CStr(i)
    Next
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo, SoHieu, Kq1,Kq2,Kq4,Kq6,Kq7) SELECT DISTINCTROW MaSo, SoHieu, CPNVL" + s621 + ", CPNC" + s622 + ",DT, (" _
        + sxc + ") AS CPSXC, CPM" + s623 + " FROM TP154 WHERE (CPNVL<>0 OR CPNC<>0 OR CPM<>0 OR DT<>0 OR (" + sxc + ")<>0) AND KPB=0"
    XDTyLePB tc
    
    sql = "SELECT Count(MaSo) AS F1 FROM BaoCaoCP"
    sopb = SelectSQL(sql)
    
    CP = SelectSQL("SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE HethongTK.SoHieu LIKE '" + shtk + "*' AND " + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=0")
    
    For i = CThangDB(tdau) To CThangDB(ThangTruoc(tcuoi))
        ExecuteSQL5 "UPDATE TP154 SET " + f + CStr(i) + "=0"
    Next
        
    i = 0
    Set rsct = DBKetoan.OpenRecordset("SELECT * FROM BaoCaoCP ORDER BY Kq3", dbOpenSnapshot)
    Do While Not rsct.EOF
        i = i + 1
        If i < sopb Then
            tongcp = RoundMoney(CP * rsct!Kq3)
            tongpb = tongpb + tongcp
        Else
            tongcp = CP - tongpb
        End If
        ExecuteSQL5 "UPDATE TP154 SET " + f + CStr(tcuoi) + "=" + DoiDau(tongcp) + " WHERE MaSo=" + CStr(rsct!MaSo)
        rsct.MoveNext
    Loop
End Sub

Public Function GetRowNumber(q As String)
    Dim rs As Recordset
    
    Set rs = DBKetoan.OpenRecordset(q, dbOpenSnapshot)
    If Not rs.EOF Then rs.MoveLast
    GetRowNumber = rs.recordCount
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
        Err.number = 0
        Set DBKetoan = WSpace.OpenDatabase(file_name, True, False, ";PWD=" + p)
        If Err.number = 0 Then
            'psw2 = "1@35^7*9)"
            DBKetoan.NewPassword p, psw2
            DBKetoan.Close

            If Err.number = 0 Then Exit Do
        End If
        If Len(p) = 0 Then GoTo KT
    Loop
    SetPsw = 0
KT:
    On Error GoTo 0
End Function

Public Sub ThemSongNgu()
    If ThemTruong("ChungTu", "DienGiaiE", dbText, DBKetoan.TableDefs("ChungTu").Fields("DienGiai").Size) Then ExecuteSQL5 "UPDATE ChungTu SET DienGiaiE=DienGiai"
    If ThemTruong("CTTaiSan", "DienGiaiE", dbText, DBKetoan.TableDefs("CTTaiSan").Fields("DienGiai").Size) Then ExecuteSQL5 "UPDATE CTTaiSan SET DienGiaiE=DienGiai"
End Sub

Private Sub dieuchinh()
    If Len(Dir(pCurDir + "UPDATE.MDB")) > 0 Then
        ThemTruong "HethongTK", "TenE", dbText, DBKetoan.TableDefs("HethongTK").Fields("Ten").Size
        If pVersion <> 3 Then
            ThemTruong "CDTS", "DienGiaiE", dbText, DBKetoan.TableDefs("CDTS").Fields("DienGiai").Size
            
            If Not BangDaCo("LaiSuat") Then CopyTable pCurDir + "UPDATE.MDB", "LaiSuat"
            
            CopyTable pCurDir + "UPDATE.MDB", "CDTS2004"
            CopyTable pCurDir + "UPDATE.MDB", "CDTS2005"
            CopyTable pCurDir + "UPDATE.MDB", "KQKD"
            CopyTable pCurDir + "UPDATE.MDB", "LCTT"
            CopyTable pCurDir + "UPDATE.MDB", "KQKDCT"
            CopyTable pCurDir + "UPDATE.MDB", "TS"
            CopyTable pCurDir + "UPDATE.MDB", "TK"
            
            ExecuteSQL5 "UPDATE HethongTK INNER JOIN TK ON HethongTK.SoHieu=TK.SoHieu SET HethongTK.TenE=TK.TenE"
            XoaBang "TS"
            XoaBang "TK"
            CopyTable2 "KQKD", "KQKD911"
            CopyTable pCurDir + "UPDATE.MDB", "THUE"
            CopyTable pCurDir + "UPDATE.MDB", "VAT"
        End If
        CopyTable pCurDir + "UPDATE.MDB", "BARCODE"
    End If
End Sub

Public Function ChoXemBC(bc As Integer, c As String) As Boolean
    If User_Right = 0 Then
        ChoXemBC = True
    Else
        If TruongDaCo("Users", c + CStr(bc)) Then
            ChoXemBC = (SelectSQL("SELECT " + c + CStr(bc) + " AS F1 FROM Users WHERE MaSo=" + CStr(UserID)) > 0)
        Else
            ChoXemBC = False
        End If
    End If
End Function

Public Function TyGiaCuoi() As Double
    TyGiaCuoi = SelectSQL("SELECT TOP 1 TyGia AS F1 FROM ChungTu ORDER BY NgayCT DESC")
    If TyGiaCuoi = 0 Then TyGiaCuoi = 1
End Function

Public Sub DoiTyGiaDB()
    Dim tygia As Double, sql As String, i As Integer
    
    tygia = SelectSQL("SELECT TyGia AS F1 FROM License")
    If tygia = 0 Then tygia = 1
    sql = IIf(pTien = 0, "*", "/") + DoiDau(tygia)
    ExecuteSQL5 "UPDATE HethongTK SET DuNo_0=DuNo_0" + sql + ",DuCo_0=DuCo_0" + sql
    ExecuteSQL5 "UPDATE SoDuKhachHang SET DuNo_0=DuNo_0" + sql + ",DuCo_0=DuCo_0" + sql
    ExecuteSQL5 "UPDATE TonKho SET Tien_0=Tien_0" + sql
        
    ExecuteSQL5 "UPDATE CTTaiSan SET NG_NS=NG_NS" + sql + ",NG_TBS=NG_TBS" + sql + ",NG_TD=NG_TD" + sql + ",NG_CNK=NG_CNK" + sql + ",CL_NS=CL_NS" + sql + ",CL_TBS=CL_TBS" + sql + ",CL_TD=CL_TD" + sql + ",CL_CNK=CL_CNK" + sql + " WHERE Thang=0"
    ExecuteSQL5 "UPDATE ThongSo SET NG_NS=NG_NS" + sql + ",NG_TBS=NG_TBS" + sql + ",NG_TD=NG_TD" + sql + ",NG_CNK=NG_CNK" + sql + ",CL_NS=CL_NS" + sql + ",CL_TBS=CL_TBS" + sql + ",CL_TD=CL_TD" + sql + ",CL_CNK=CL_CNK" + sql + ",KH_NS=KH_NS" + sql + ",KH_TBS=KH_TBS" + sql + ",KH_TD=KH_TD" + sql + ",KH_CNK=KH_CNK" + sql + " WHERE Thang=0"
    
    sql = IIf(pTien = 0, "*", "/") + "IIF(ChungTu.TyGia>0,ChungTu.TyGia,1)"
    ExecuteSQL5 "UPDATE ChungTu SET SoPS=SoPS" + sql
    ExecuteSQL5 "UPDATE " + ChungTu2TKHD(0) + " SET ThanhTien=ThanhTien" + sql
    ExecuteSQL5 "UPDATE CTTaiSan INNER JOIN ChungTu ON CTTaiSan.MaCTKT=ChungTu.MaCT SET NG_NS=NG_NS" + sql + ",NG_TBS=NG_TBS" + sql + ",NG_TD=NG_TD" + sql + ",NG_CNK=NG_CNK" + sql + ",CL_NS=CL_NS" + sql + ",CL_TBS=CL_TBS" + sql + ",CL_TD=CL_TD" + sql + ",CL_CNK=CL_CNK" + sql
    For i = 12 To 1 Step -1
        ExecuteSQL5 "UPDATE ThongSo INNER JOIN CTTaiSan ON ThongSo.MaTS=CTTaiSan.MaTS SET ThongSo.NG_NS=CTTaiSan.NG_NS,ThongSo.NG_TBS=CTTaiSan.NG_TBS,ThongSo.NG_TD=CTTaiSan.NG_TD,ThongSo.NG_CNK=CTTaiSan.NG_CNK,ThongSo.CL_NS=CTTaiSan.CL_NS,ThongSo.CL_TBS=CTTaiSan.CL_TBS,ThongSo.CL_TD=CTTaiSan.CL_TD,ThongSo.CL_CNK=CTTaiSan.CL_CNK WHERE CTTaiSan.MaLoai=" + CStr(NV_TANG) + " AND CTTaiSan.Thang=" + CStr(CThangFR(i)) + " AND ThongSo.Thang=" + CStr(i)
        tygia = SelectSQL("SELECT TyGia AS F1 FROM ChungTu INNER JOIN CTTaiSan ON ChungTu.MaCT=CTTaiSan.MaCTKT WHERE CTTaiSan.MaLoai=" + CStr(NV_TKHAO) + " AND Thang=" + CStr(CThangFR(i)))
        If tygia > 0 Then
            sql = IIf(pTien = 0, "*", "/") + DoiDau(tygia)
            ExecuteSQL5 "UPDATE ThongSo SET KH_NS=KH_NS" + sql + ",KH_TBS=KH_TBS" + sql + ",KH_TD=KH_TD" + sql + ",KH_CNK=KH_CNK" + sql + " WHERE Thang<=" + CStr(i)
        End If
    Next
    
    If pTien = 0 Then
        ExecuteSQL5 "UPDATE HethongTK SET DuNo_0=Fix(0.5+DuNo_0),DuCo_0=Fix(0.5+DuCo_0)"
        ExecuteSQL5 "UPDATE SoDuKhachHang SET DuNo_0=Fix(0.5+DuNo_0),DuCo_0=Fix(0.5+DuCo_0)"
        ExecuteSQL5 "UPDATE TonKho SET Tien_0=Fix(0.5+Tien_0)"
        ExecuteSQL5 "UPDATE CTTaiSan SET NG_NS=Fix(0.5+NG_NS),NG_TBS=Fix(0.5+NG_TBS),NG_TD=Fix(0.5+NG_TD),NG_CNK=Fix(0.5+NG_CNK),CL_NS=Fix(0.5+CL_NS),CL_TBS=Fix(0.5+CL_TBS),CL_TD=Fix(0.5+CL_TD),CL_CNK=Fix(0.5+CL_CNK)"
        ExecuteSQL5 "UPDATE ThongSo SET NG_NS=Fix(0.5+NG_NS),NG_TBS=Fix(0.5+NG_TBS),NG_TD=Fix(0.5+NG_TD),NG_CNK=Fix(0.5+NG_CNK),CL_NS=Fix(0.5+CL_NS),CL_TBS=Fix(0.5+CL_TBS),CL_TD=Fix(0.5+CL_TD),CL_CNK=Fix(0.5+CL_CNK),KH_NS=Fix(0.5+KH_NS),KH_TBS=Fix(0.5+KH_TBS),KH_TD=Fix(0.5+KH_TD),KH_CNK=Fix(0.5+KH_CNK)"
    Else
        sql = CStr(Mask_N)
        ExecuteSQL5 "UPDATE HethongTK SET DuNo_0=Fix(0.5+" + sql + "*DuNo_0)/" + sql + ",DuCo_0=Fix(0.5+" + sql + "*DuCo_0)/" + sql
        ExecuteSQL5 "UPDATE SoDuKhachHang SET DuNo_0=Fix(0.5+" + sql + "*DuNo_0)/" + sql + ",DuCo_0=Fix(0.5+" + sql + "*DuCo_0)/" + sql
        ExecuteSQL5 "UPDATE TonKho SET Tien_0=Fix(0.5+" + sql + "*Tien_0)/" + sql
        ExecuteSQL5 "UPDATE CTTaiSan SET NG_NS=Fix(0.5+" + sql + "*NG_NS)/" + sql + ",NG_TBS=Fix(0.5+" + sql + "*NG_TBS)/" + sql + ",NG_TD=Fix(0.5+" + sql + "*NG_TD)/" + sql + ",NG_CNK=Fix(0.5+" + sql + "*NG_CNK)/" + sql + ",CL_NS=Fix(0.5+" + sql + "*CL_NS)/" + sql + ",CL_TBS=Fix(0.5+" + sql + "*CL_TBS)/" + sql + ",CL_TD=Fix(0.5+" + sql + "*CL_TD)/" + sql + ",CL_CNK=Fix(0.5+" + sql + "*CL_CNK)/" + sql
        ExecuteSQL5 "UPDATE ThongSo SET NG_NS=Fix(0.5+" + sql + "*NG_NS)/" + sql + ",NG_TBS=Fix(0.5+" + sql + "*NG_TBS)/" + sql + ",NG_TD=Fix(0.5+" + sql + "*NG_TD)/" + sql + ",NG_CNK=Fix(0.5+" + sql + "*NG_CNK)/" + sql + ",CL_NS=Fix(0.5+" + sql + "*CL_NS)/" + sql + ",CL_TBS=Fix(0.5+" + sql + "*CL_TBS)/" + sql + ",CL_TD=Fix(0.5+" + sql + "*CL_TD)/" + sql + ",CL_CNK=Fix(0.5+" + sql + "*CL_CNK)/" + sql + ",KH_NS=Fix(0.5+" + sql + "*KH_NS)/" + sql + ",KH_TBS=Fix(0.5+" + sql + "*KH_TBS)/" + sql + ",KH_TD=Fix(0.5+" + sql + "*KH_TD)/" + sql + ",KH_CNK=Fix(0.5+" + sql + "*KH_CNK)/" + sql
    End If
    
    ExecuteSQL5 "UPDATE License SET TyGia=0"
    KiemTraVatTu
    KiemTraTaiKhoan
End Sub

Public Function KiemTraMaSoThue(mst As String, mask As String, Optional f1 As Integer = 0) As Boolean
    Dim st As String, i As Integer, s1 As String, F2 As Integer, s2 As String
            
    If DEMO = 1 Or Cdbl5(Left(mst, 10)) = 0 Or (Len(pMST) > 0 And Left(mst, Len(pMST)) = pMST) Then GoTo KT2
    KiemTraMaSoThue = False
    
    If Len(mst) < 12 Then
        st = Left(App.Comments, 2)
        If CInt5(st) > 0 Then
            i = 0
            st = Left(mst, 2)
            Do While Len(App.Comments) > 2 * i + 1
                If st = Mid(App.Comments, 2 * i + 1, 2) Then GoTo K1
                i = i + 1
            Loop
            Exit Function
        End If
    End If
    
K1:
    st = Left(mst, 2)
    If CInt5(st) = 0 Or Len(mask) = 0 Or Len(mst) > 11 Then GoTo KT
    i = 0
    Do While Len(mask) > 2 * i + 1
        If st = Mid(mask, 2 * i + 1, 2) Then GoTo KT
        i = i + 1
    Loop
    
    Exit Function
KT:
    If f1 > 0 And CInt5(st) > 0 And Len(mst) < 12 Then
        For i = 1 To 99
            s1 = IIf(i < 10, "0", "") + CStr(i)
            If QueryDaCo2(s1, s2) Then
                If s2 = Left(mst, Len(s2)) Then
                    F2 = 0
                    Exit For
                Else
                    F2 = 1
                End If
            End If
        Next
    End If
    If F2 = 0 Then
KT2:
        KiemTraMaSoThue = True
    End If
End Function

Private Function QueryDaCo2(qname As String, qname2 As String) As Boolean
    Dim i As Integer, L As Integer
    
    QueryDaCo2 = False
    L = Len(qname)
    For i = 0 To DBKetoan.QueryDefs.count - 1
        If Left(UCase(DBKetoan.QueryDefs(i).Name), L) = UCase(qname) Then
            qname2 = DBKetoan.QueryDefs(i).Name
            QueryDaCo2 = True
            Exit For
        End If
    Next
End Function

Public Function ChoNhapTiep() As Boolean
    If DEMO = 1 Or frmMain.Lb(0).tag = 5 Or (CInt5(Left(frmMain.LbCty(8).Caption, 2)) = 0 And pNoiBo = 0 And pVersion <> 3) Or pVersion = 2 Then
        SetSQL "MienTru", "SELECT MaCT FROM ChungTu WHERE MaCT>0 GROUP BY MaCT"
        ChoNhapTiep = SelectSQL("SELECT Count(MaCT) AS F1 FROM MienTru") < 200
    Else
        ChoNhapTiep = True
    End If
End Function

Public Function TenUser(ms As Long) As String
    TenUser = SelectSQL("SELECT TenNSD AS F1 FROM Users WHERE MaSo=" + CStr(ms))
End Function

Public Sub CloseItemList()
    On Error Resume Next
    Unload FrmTaikhoan
    Unload FrmVattu
    Unload FrmKhachHang
    Unload FrmTP
    On Error GoTo 0
End Sub

Public Sub DelTemp()
    ExecuteSQL5 "DELETE * FROM BaoCaoCP", False
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2", False
    ExecuteSQL5 "DELETE * FROM BKNhomPS", False
    'ExecuteSQL5 "DELETE * FROM ChungTuP", False
    ExecuteSQL5 "DELETE * FROM CNDauNam2", False
End Sub

Public Function ChoDieuChinhDauKy() As Boolean
    Dim c As Boolean
    
    ChoDieuChinhDauKy = True
    If User_Right = 0 Then Exit Function
    c = (SelectSQL("SELECT Lock0 AS F1 FROM License") = 0)
    ChoDieuChinhDauKy = c
    If Not c Then MsgBox "Kh«ng ®­îc phÐp ®iÒu chØnh sè d­ ®Çu kú!", vbCritical, App.ProductName
End Function

Public Sub XoaDK()
    
    If SelectSQL("SELECT Sum(Tien_0) AS F1 FROM TonKho") <> 0 Then
        If MsgBox("Xo¸ sè d­ ®Çu kú cña vËt t­, hµng ho¸ ?", vbInformation + vbYesNo, App.ProductName) = vbYes Then
            ExecuteSQL5 "DELETE * FROM TonKho WHERE Luong_0<>0 OR Tien_0<>0"
            KiemTraVatTu
        End If
    End If
    
    If SelectSQL("SELECT ABS(Sum(DuNo_0))+ABS(Sum(DuCo_0))+ABS(Sum(DuNT_0)) AS F1 FROM SoDuKhachHang") > 0 Then
        If MsgBox("Xo¸ sè d­ ®Çu kú cña c¸c danh ®iÓm c«ng nî ?", vbInformation + vbYesNo, App.ProductName) = vbYes Then
            ExecuteSQL5 "DELETE * FROM SoDuKhachHang WHERE DuNo_0<>0 OR DuCo_0<>0 OR DuNT_0<>0"
            SoDuTKCN
        End If
    End If
    
    If SelectSQL("SELECT Count(MaTS) AS F1 FROM CTTaiSan WHERE Thang=0 AND MaLoai=30") > 0 Then
        If MsgBox("Xo¸ danh môc TSC§ ®Çu kú ?", vbInformation + vbYesNo, App.ProductName) = vbYes Then
            ExecuteSQL5 "DELETE ThongSo.* FROM ThongSo INNER JOIN CTTaiSan ON ThongSo.MaTS=CTTaiSan.MaTS WHERE CTTaiSan.Thang=0 AND CTTaiSan.MaLoai=30"
            ExecuteSQL5 "DELETE * FROM CTTaiSan WHERE Thang=0 AND MaLoai=30"
            ExecuteSQL5 "DELETE TaiSan.* FROM TaiSan LEFT JOIN ThongSo ON TaiSan.MaSo=ThongSo.MaTS WHERE IsNull(ThongSo.MaTS)"
            SoDuTKTS
        End If
    End If
                
    ExecuteSQL5 "UPDATE HethongTK SET DuNo_0=0,DuCo_0=0,DuNT_0=0 WHERE TK_ID<>" + CStr(TKVT_ID) + " AND TK_ID<>" + CStr(TKCNKH_ID) + " AND TK_ID<>" + CStr(TKCNPT_ID) + " AND TK_ID<>" + CStr(TSCD_ID)
    
    KiemTraTaiKhoan
End Sub

Public Sub XoaPhieu(mct As Long)
 
    ExecuteSQL5 "DELETE * FROM ChungTuP WHERE MaCT=" + CStr(mct)
    ExecuteSQL5 "DELETE HoaDonP.* FROM HoaDonP LEFT JOIN ChungTuP ON HoaDonP.MaSo=ChungTuP.MaSo WHERE IsNull(ChungTuP.MaSo)"
    ExecuteSQL5 "DELETE FROM ChungtuLQP WHERE MaCT=" + CStr(mct)
End Sub
'======================================================================================
' Thñ tôc in B¶ng c©n ®èi kÕ to¸n
'======================================================================================
Public Sub InCdts(tdau As Integer, tcuoi As Integer, loai As Integer, Optional nn As Integer)
    Dim rs_cdts As Recordset, rs As Recordset, sql As String, st As String, st1 As String, i As Integer
    Dim SoDK As Double, sodu As Double
    Dim TK As ClsTaikhoan, dau As Integer, dn1 As Double, dc1 As Double, nt As Double, dn2 As Double, dc2 As Double
    
    SoDuTKCN2 ThangTruoc(tdau)
    SoDuTKCN2 tcuoi
    KiemTraKetChuyen tcuoi
    
    st1 = CStr(CThangDB(ThangTruoc(tdau)))
    st = CStr(CThangDB(tcuoi))
    
    If loai > 0 Then
        XoaBang "CDTS"
        CopyTable2 "CDTS2005", "CDTS", 1
    ExecuteSQL5 ("delete from CDTS")
     ExecuteSQL5 ("insert into CDTS select * from CDTS2005 order by maso")
        Set TK = New ClsTaikhoan
        Set rs_cdts = DBKetoan.OpenRecordset("SELECT * FROM Cdts WHERE Len(ShTK1)>0", dbOpenSnapshot)
        Do While Not rs_cdts.EOF
            sodu = 0
            SoDK = 0
            For i = 1 To 10
                If Len(rs_cdts.Fields("ShTK" + CStr(i))) > 0 Then
                    TK.InitTaikhoanSohieu rs_cdts.Fields("ShTK" + CStr(i))
                    If TK.MaSo > 0 Then
                        TK.SoDuTK ThangTruoc(tdau), dn1, dc1, nt
                        TK.SoDuTK tcuoi, dn2, dc2, nt
                        
                        If Left(TK.sohieu, 3) = "131" Or Left(TK.sohieu, 3) = "136" Or Left(TK.sohieu, 3) = "138" Or Left(TK.sohieu, 3) = "331" Or Left(TK.sohieu, 3) = "336" Or Left(TK.sohieu, 3) = "338" Or Left(TK.sohieu, 3) = "333" Then
                            If rs_cdts!TaiSan < 0 Then
                                SoDK = SoDK + dn1
                                sodu = sodu + dn2
                            Else
                                SoDK = SoDK + dc1
                                sodu = sodu + dc2
                            End If
                        Else
                            If rs_cdts!TaiSan < 0 Then
                                SoDK = SoDK + dn1 - dc1
                                sodu = sodu + dn2 - dc2
                            Else
                                SoDK = SoDK + dc1 - dn1
                                sodu = sodu + dc2 - dn2
                            End If
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next
            ExecuteSQL5 "UPDATE Cdts SET DauNam = " + DoiDau(SoDK) + ", CuoiKy=" + DoiDau(sodu) + " WHERE MaSo=" + CStr(rs_cdts!MaSo)
            rs_cdts.MoveNext
        Loop
        rs_cdts.Close
        Set TK = Nothing
        
        'sodu = SelectSQL("SELECT DauNam AS F1 FROM Cdts WHERE MaSo=152")
        'If sodu > 0 Then
        '    sodu = SoDuTKSH("3331", ThangTruoc(tdau))
        '    ExecuteSQL5 "UPDATE Cdts SET DauNam=DauNam-" + DoiDau(sodu) + " WHERE MaSo=314"
        'Else
        '    ExecuteSQL5 "UPDATE Cdts SET DauNam=0 WHERE MaSo=152"
        'End If
        
        'sodu = SelectSQL("SELECT CuoiKy AS F1 FROM Cdts WHERE MaSo=152")
        'If sodu > 0 Then
        '    sodu = SoDuTKSH("3331", tcuoi)
        '    ExecuteSQL5 "UPDATE Cdts SET CuoiKy=CuoiKy-" + DoiDau(sodu) + " WHERE MaSo=314"
        'Else
        '    ExecuteSQL5 "UPDATE Cdts SET CuoiKy=0 WHERE MaSo=152"
        'End If
    Else
        XoaBang "CDTS"
        CopyTable2 "CDTS2004", "CDTS", 1
        
        sql = "SELECT DISTINCTROW CdTs.MaSo, CdTs.MaTk1, CdTs.MaTk2, CdTs.MaTk3, CdTs.TaiSan, HeThongTK.DuNo_" + st1 + " As DNo1, HeThongTK.DuCo_" + st1 + " As DCo1, HeThongTK_1.DuNo_" + st1 + " As DNo2, HeThongTK_1.DuCo_" + st1 + " As DCo2, HeThongTK_2.DuNo_" + st1 + " As DNo3, HeThongTK_2.DuCo_" + st1 + " As DCo3, " _
            & " HeThongTK.DuNo_" + st + " As CNo1, HeThongTK.DuCo_" + st + " As CCo1, HeThongTK_1.DuNo_" + st + " As CNo2, HeThongTK_1.DuCo_" + st + " As CCo2, HeThongTK_2.DuNo_" + st + " As CNo3, HeThongTK_2.DuCo_" + st + " As CCo3" _
            & " FROM ((CdTs LEFT JOIN HeThongTK ON CdTs.MaTk1 = HeThongTK.MaSo) LEFT JOIN HeThongTK AS HeThongTK_1 ON CdTs.MaTk2 = HeThongTK_1.MaSo) LEFT JOIN HeThongTK AS HeThongTK_2 ON CdTs.MaTk3 = HeThongTK_2.MaSo WHERE CdTs.CoChitiet = 0 And CdTs.MaTk1 > 0"
                      
        ' Tinh so cuoi ky
        Set rs_cdts = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        If rs_cdts.recordCount > 0 Then
            rs_cdts.MoveLast
            'GauGe.Max = Fix(0.5 + 1.5 * rs_cdts.RecordCount)
        End If
        Do While Not rs_cdts.BOF
            SoDK = 0
            sodu = 0
            If rs_cdts!MaTk1 > 0 And (Not IsNull(rs_cdts!DNo1)) Then
                If rs_cdts!TaiSan < 0 Then
                    If rs_cdts!MaTk2 <> 0 Then
                        SoDK = rs_cdts!DNo1
                        sodu = rs_cdts!CNo1
                    Else
                        SoDK = rs_cdts!DNo1 - rs_cdts!DCo1
                        sodu = rs_cdts!CNo1 - rs_cdts!CCo1
                    End If
                Else
                    If rs_cdts!MaTk2 <> 0 Then
                        SoDK = rs_cdts!DCo1
                        sodu = rs_cdts!CCo1
                    Else
                        SoDK = -rs_cdts!DNo1 + rs_cdts!DCo1
                        sodu = -rs_cdts!CNo1 + rs_cdts!CCo1
                    End If
                End If
            End If
            If rs_cdts!MaTk2 > 0 And (Not IsNull(rs_cdts!DNo2)) Then
                If rs_cdts!TaiSan < 0 Then
                    SoDK = SoDK + rs_cdts!DNo2
                    sodu = sodu + rs_cdts!CNo2
                Else
                    SoDK = SoDK + rs_cdts!DCo2
                    sodu = sodu + rs_cdts!CCo2
                End If
            End If
            If rs_cdts!MaTk3 > 0 And (Not IsNull(rs_cdts!DNo3)) Then
                If rs_cdts!TaiSan < 0 Then
                    SoDK = SoDK - rs_cdts!DNo3
                    sodu = sodu - rs_cdts!CNo3
                Else
                    SoDK = SoDK - rs_cdts!DCo3
                    sodu = sodu - rs_cdts!CCo3
                End If
            End If
            ExecuteSQL5 "UPDATE Cdts SET DauNam = " + DoiDau(SoDK) + ", CuoiKy=" + DoiDau(sodu) + " WHERE MaSo=" + CStr(rs_cdts!MaSo)
            'If GauGe.Value < GauGe.Max - 1 Then GauGe.Value = GauGe.Value + 1
            rs_cdts.MovePrevious
        Loop
        
        If SelectSQL("SELECT Top 1 MaSo AS F1 FROM HethongTK WHERE SoHieu LIKE '621*'") = 0 Then
            sodu = SelectSQL("SELECT SUM(DuNo_" + st + "-DuCo_" + st + ") AS F1, SUM(DuNo_0-DuCo_0) AS F2 FROM HethongTK WHERE MaSo=37 OR MaSo=38 OR MaSo=39 OR MaSo=42", SoDK)
            ExecuteSQL5 "UPDATE Cdts SET DauNam=DauNam+" + DoiDau(SoDK) + ",CuoiKy=CuoiKy+" + DoiDau(sodu) + " WHERE MaSo=118"
        End If
        
        SoDK = SelectSQL("SELECT DISTINCTROW DauNam AS F1, CuoiKy AS F2 FROM Cdts WHERE MaSo=318", sodu)
        If SoDK < 0 Then
            ExecuteSQL5 "UPDATE Cdts SET DauNam=DauNam+" + DoiDau(-SoDK) + " WHERE MaSo=138"
            ExecuteSQL5 "UPDATE Cdts SET DauNam=0 WHERE MaSo=318"
        End If
        If sodu < 0 Then
            ExecuteSQL5 "UPDATE Cdts SET CuoiKy=CuoiKy+" + DoiDau(-sodu) + " WHERE MaSo=138"
            ExecuteSQL5 "UPDATE Cdts SET CuoiKy=0 WHERE MaSo=318"
        End If
    End If
    ' Cong don
    Set rs_cdts = DBKetoan.OpenRecordset("SELECT MaSo FROM CdTs WHERE CoChiTiet=1 ORDER BY MaSo DESC", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_cdts.EOF
        Set rs = DBKetoan.OpenRecordset("SELECT SUM(DauNam) as dn,SUM(CuoiKy) as ck FROM CdTs WHERE NhomCha=" + CStr(rs_cdts!MaSo), dbOpenSnapshot)
        ExecuteSQL5 "UPDATE Cdts SET DauNam=" + DoiDau(rs!dn) + ",CuoiKy=" + DoiDau(rs!CK) + " WHERE MaSo=" + CStr(rs_cdts!MaSo)
        rs.Close
        'If GauGe.Value < GauGe.Max - 1 Then GauGe.Value = GauGe.Value + 1
        rs_cdts.MoveNext
    Loop
    rs_cdts.Close
    Set rs_cdts = Nothing
    Set rs = Nothing
    frmMain.Rpt.ReportFileName = "CDTS.RPT"
    frmMain.Rpt.Formulas(3) = "Thang=" + CStr(tcuoi)
    If tdau <> pThangDauKy Then frmMain.Rpt.Formulas(4) = "DKLB='" + IIf(nn = 0, "Sè ®Çu kú ", "Opening Balance of ") + CStr(tdau) + "/" + CStr(IIf(tdau < pThangDauKy, pNamTC + 1, pNamTC)) + "'"
    'GauGe.Value = GauGe.Max
End Sub

Public Sub KiemTraKetChuyen(thang As Integer)
    Dim sql As String, sodu As Double, sodu2 As Double
    Dim rs As Recordset, st As String
    
    sodu = SelectSQL("SELECT Sum(DuNo_0-DuCo_0) AS F1 FROM HethongTK WHERE Cap=0 AND (Loai=1 OR Loai=2)")
    sodu2 = SelectSQL("SELECT Sum(DuCo_0-DuNo_0) AS F1 FROM HethongTK WHERE Cap=0 AND (Loai=3 OR Loai=4)")
    
    If Fix((sodu - sodu2) * Mask_N) <> 0 Then MsgBox "Sè d­ ®Çu kú chªnh lÖch :" + DoiDau(Abs(sodu - sodu2)) + Chr(13) + "Tµi s¶n:" + DoiDau(sodu) + Chr(13) + "Nguån vèn:" + DoiDau(sodu2), vbCritical, App.ProductName
    
    st = CStr(CThangDB(thang))
    sql = "SELECT SoHieu,Loai,DuNo_" + st + " AS DuNo,DuCo_" + st + " AS DuCo FROM HethongTK WHERE MaTC=MaSo AND Loai>=" + CStr(pTKTrunggian) + " AND (DuNo_" + st + "<>0 OR DuCo_" + st + "<>0) AND (DuNo_" + st + "<>DuCo_" + st + ")"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    sql = ""
    If rs.recordCount > 0 Then
        Do While Not rs.EOF
            sql = sql + Chr(13) + "Sè d­ tµi kho¶n " + rs!sohieu + ":" + Format(IIf(rs!loai > 0, rs!duco - rs!duno, rs!duno - rs!duco), Mask_0)
            rs.MoveNext
        Loop
        MsgBox "C¸c tµi kho¶n cßn sè d­:" + sql, vbCritical, App.ProductName
    End If
    rs.Close
    Set rs = Nothing
End Sub
'====================================================================================================
' Thñ tôc in l·i lç
'====================================================================================================
Public Sub InKqkd(tdau As Integer, tcuoi As Integer, cap As Integer, nn As Integer)
    Dim rs As Recordset, sql As String, KT As Double, kn As Double
    
        ' Tinh so ky truoc
        
    ExecuteSQL5 "UPDATE KQKD SET KyNay=0, KyTruoc=0"
    
    sql = "SELECT DISTINCTROW KQKD.Ma, Sum(IIF(" + WThang2("ThangCT", 0, tdau) + ",ChungTu.SoPS,0)) AS Kqua1, Sum(IIF(" + WThang("ThangCT", tdau, 0) + ",ChungTu.SoPS,0)) AS Kqua2 " _
        & " FROM (" + ChungTu2TKNC(0) + ") INNER JOIN KQKD ON HethongTK.SoHieu LIKE KQKD.SHTKNo+'*' AND TK.SoHieu LIKE KQKD.SHTKCo+'*' " _
        & " Where " + WThang("ThangCT", 0, tcuoi) + " AND KQKD.SHTKNo<>'0' AND KQKD.SHTKCo<>'0' GROUP BY KQKD.Ma"
    
    sql = sql + " UNION SELECT DISTINCTROW KQKD.Ma, Sum(IIF(" + WThang2("ThangCT", 0, tdau) + ",ChungTu.SoPS,0)) AS Kqua1, Sum(IIF(" + WThang("ThangCT", tdau, 0) + ",ChungTu.SoPS,0)) AS Kqua2" _
        & " FROM (" + ChungTu2TKNC(1) + ") INNER JOIN KQKD ON HethongTK.SoHieu LIKE KQKD.SHTKCo+'*'" _
        & " Where (MaLoai<>3 AND MaLoai<>1) And " + WThang("ThangCT", 0, tcuoi) + " And (CLng(KQKD.SHTKNo) = 0) AND KQKD.SHTKCo<>'0' GROUP BY KQKD.Ma"
    
    'sql = sql + " UNION SELECT DISTINCTROW KQKD.Ma, Sum(IIF(" + WThang2("ThangCT", 0, tdau) + ",ChungTu.SoPS,0)) AS Kqua1, Sum(IIF(" + WThang("ThangCT", tdau, 0) + ",ChungTu.SoPS,0)) AS Kqua2 " _
        & " FROM (((" + ChungTu2TKNC(0) + ") INNER JOIN KQKD ON HethongTK.SoHieu LIKE KQKD.SHTKNo+'*') INNER JOIN KQKD ON HethongTK_1.SoHieu LIKE KQKD.SHTKCo+'*' " _
        & " Where ((KQKD.MaSoCha=0) AND (KQKD.MaTKTCNo > 0) And (KQKD.MaTKTCCo > 0) And " + WThang("ThangCT", 0, tcuoi) + ") GROUP BY KQKD.Ma"

    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
            ExecuteSQL5 "UPDATE KQKD SET KyTruoc=" + DoiDau(IIf(IsNull(rs!kqua1), 0, rs!kqua1)) + ", KyNay = " + DoiDau(IIf(IsNull(rs!kqua2), 0, rs!kqua2)) + " WHERE Ma=" + CStr(rs!ma)
            rs.MoveNext
    Loop

    
    sql = "SELECT DISTINCTROW KQKD.Ma, Sum(IIF(" + WThang2("ThangCT", 0, tdau) + ",ChungTu.SoPS,0)) AS Kqua1, Sum(IIF(" + WThang("ThangCT", tdau, 0) + ",ChungTu.SoPS,0)) AS Kqua2 " _
        & " FROM (" + ChungTu2TKNC(0) + ") INNER JOIN KQKD ON HethongTK.SoHieu LIKE KQKD.SHTKCo+'*' AND TK.SoHieu LIKE KQKD.SHTKNo+'*' " _
        & " Where " + WThang("ThangCT", 0, tcuoi) + " GROUP BY KQKD.Ma"
        
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
            ExecuteSQL5 "UPDATE KQKD SET KyTruoc = KyTruoc - " + DoiDau(IIf(IsNull(rs!kqua1), 0, rs!kqua1)) + ", KyNay = KyNay - " + DoiDau(IIf(IsNull(rs!kqua2), 0, rs!kqua2)) + " WHERE Ma=" + CStr(rs!ma)
            rs.MoveNext
    Loop
            
    ' Cong don
    Set rs = DBKetoan.OpenRecordset("SELECT KyTruoc,KyNay,MaSoCha FROM KQKD WHERE CLng(SHTKNo)<>0 or CLng(SHTKCo)<>0 ORDER BY MaSo", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
         If rs!MaSoCha > 0 Then
                 ExecuteSQL5 "UPDATE KQKD SET KyTruoc=KyTruoc+" + DoiDau(rs!KyTruoc) + ",KyNay=KyNay+" + DoiDau(rs!KyNay) + " WHERE MaSo=" + CStr(rs!MaSoCha)
        Else
                 ExecuteSQL5 "UPDATE KQKD SET KyTruoc=KyTruoc-" + DoiDau(rs!KyTruoc) + ",KyNay=KyNay-" + DoiDau(rs!KyNay) + " WHERE MaSo=" + CStr(-rs!MaSoCha)
        End If
        rs.MoveNext
    Loop
    
    Set rs = DBKetoan.OpenRecordset("SELECT KyTruoc,KyNay,MaSoCha FROM KQKD WHERE CLng(SHTKNo)=0 And CLng(SHTKCo)=0 And (TongHop=1) ORDER BY MaSo", dbOpenDynaset, dbForwardOnly)
    Do While Not rs.EOF
        If rs!MaSoCha > 0 Then
            ExecuteSQL5 "UPDATE KQKD SET KyTruoc=KyTruoc+" + DoiDau(rs.Fields("KyTruoc")) + ",KyNay=KyNay+" + DoiDau(rs!KyNay) + " WHERE MaSo=" + CStr(rs!MaSoCha)
        Else
            ExecuteSQL5 "UPDATE KQKD SET KyTruoc=KyTruoc-" + DoiDau(rs.Fields("KyTruoc")) + ",KyNay=KyNay-" + DoiDau(rs!KyNay) + " WHERE MaSo=" + CStr(-rs!MaSoCha)
        End If
        rs.MoveNext
    Loop
    rs.Close
            
        
        If cap > 0 Then
            ExecuteSQL5 "DELETE * FROM KQKD911"
            ExecuteSQL5 "INSERT INTO KQKD911 SELECT * FROM KQKD WHERE MaSo<>23"
    
            ExecuteSQL5 "INSERT INTO KQKD911 (Ma, MaSo, Ten, TenE, SHTKNo, SHTKCo, MaTK, Cap, TongHop, MaSoCha) SELECT 5700000+MaSo, 10 AS MS, Space(4*(Cap-1))+Ten, Space(4*(Cap-1))+TenE, '511', SoHieu, MaSo, Cap, 1, 20 FROM HethongTK WHERE SoHieu LIKE '911*' AND Cap>1 AND Cap<=" + CStr(cap)
            ExecuteSQL5 "INSERT INTO KQKD911 (Ma, MaSo, Ten, TenE, SHTKNo, SHTKCo, MaTK, Cap, TongHop, MaSoCha) SELECT 5750000+MaSo, 10 AS MS, Space(4*(Cap-1))+Ten, Space(4*(Cap-1))+TenE, '512', SoHieu, MaSo, Cap, 1, 20 FROM HethongTK WHERE SoHieu LIKE '911*' AND Cap>1 AND Cap<=" + CStr(cap)
            
            Set rs = DBKetoan.OpenRecordset("SELECT * FROM KQKD911 WHERE MaTK=0 AND MaSo > 10", dbOpenSnapshot)
            Do While Not rs.EOF
                ExecuteSQL5 "INSERT INTO KQKD911 (Ma, MaSo, Ten, TenE, SHTKNo, SHTKCo, MaTK, Cap, TongHop, MaSoCha) SELECT " + CStr(rs!ma * 100000) + "+MaSo, " + CStr(rs!MaSo) + " AS MS, Space(4*(Cap-1))+Ten, Space(4*(Cap-1))+TenE, IIF(LEFT('" + rs!shtkno + "',3)='911',SoHieu,'" + rs!shtkno + "'), IIF(LEFT('" + rs!shtkco + "',3)='911',SoHieu,'" + rs!shtkco + "'), MaSo, Cap," + CStr(rs!TongHop) + "," + CStr(rs!MaSoCha) + " FROM HethongTK WHERE SoHieu LIKE '911*' AND Cap>1 AND Cap<=" + CStr(cap)
                rs.MoveNext
            Loop
            rs.Close
                                    
            sql = "SELECT DISTINCTROW KQKD911.Ma, Sum(IIF(" + WThang2("ThangCT", 0, tdau) + ",ChungTu.SoPS,0)) AS Kqua1, Sum(IIF(" + WThang("ThangCT", tdau, 0) + ",ChungTu.SoPS,0)) AS Kqua2 " _
                & " FROM (" + ChungTu2TKNC(0) + ") INNER JOIN KQKD911 ON HethongTK.SoHieu LIKE KQKD911.SHTKNo+'*' AND TK.SoHieu LIKE KQKD911.SHTKCo+'*' " _
                & " Where MaTK>0 AND " + WThang("ThangCT", 0, tcuoi) + " GROUP BY KQKD911.Ma"
            Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
            Do While Not rs.EOF
                    ExecuteSQL5 "UPDATE KQKD911 SET KyTruoc=" + DoiDau(IIf(IsNull(rs!kqua1), 0, rs!kqua1)) + ", KyNay = " + DoiDau(IIf(IsNull(rs!kqua2), 0, rs!kqua2)) + " WHERE Ma=" + CStr(rs!ma)
                    rs.MoveNext
            Loop
            rs.Close
            
            sql = "SELECT DISTINCTROW KQKD911.Ma, Sum(IIF(" + WThang2("ThangCT", 0, tdau) + ",ChungTu.SoPS,0)) AS Kqua1, Sum(IIF(" + WThang("ThangCT", tdau, 0) + ",ChungTu.SoPS,0)) AS Kqua2 " _
                & " FROM (" + ChungTu2TKNC(0) + ") INNER JOIN KQKD911 ON HethongTK.SoHieu LIKE KQKD911.SHTKCo+'*' AND TK.SoHieu LIKE KQKD911.SHTKNo+'*' " _
                & " Where MaTK>0 AND " + WThang("ThangCT", 0, tcuoi) + " GROUP BY KQKD911.Ma"
            Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
            Do While Not rs.EOF
                    ExecuteSQL5 "UPDATE KQKD911 SET KyTruoc=KyTruoc-" + DoiDau(IIf(IsNull(rs!kqua1), 0, rs!kqua1)) + ", KyNay = KyNay-" + DoiDau(IIf(IsNull(rs!kqua2), 0, rs!kqua2)) + " WHERE Ma=" + CStr(rs!ma)
                    rs.MoveNext
            Loop
            rs.Close
            
            ' Cong don
            Set rs = DBKetoan.OpenRecordset("SELECT MaTK, KyTruoc,KyNay,MaSoCha FROM KQKD911 WHERE MaTK>0 AND Cap>0 AND (SHTKNo<>'0' or SHTKCo<>'0') ORDER BY MaSo, MaTK", dbOpenSnapshot, dbForwardOnly)
            Do While Not rs.EOF
                 If rs!MaSoCha > 0 Then
                         ExecuteSQL5 "UPDATE KQKD911 SET KyTruoc=KyTruoc+" + DoiDau(rs!KyTruoc) + ",KyNay=KyNay+" + DoiDau(rs!KyNay) + " WHERE MaSo=" + CStr(rs!MaSoCha) + " AND MaTK=" + CStr(rs!MaTK)
                Else
                         ExecuteSQL5 "UPDATE KQKD911 SET KyTruoc=KyTruoc-" + DoiDau(rs!KyTruoc) + ",KyNay=KyNay-" + DoiDau(rs!KyNay) + " WHERE MaSo=" + CStr(-rs!MaSoCha) + " AND MaTK=" + CStr(rs!MaTK)
                End If
                rs.MoveNext
            Loop
            
            Set rs = DBKetoan.OpenRecordset("SELECT MaTK, KyTruoc,KyNay,MaSoCha FROM KQKD911 WHERE MaTK>0 AND Cap>0 AND SHTKNo='0' And SHTKCo='0' And (TongHop=1) ORDER BY MaSo", dbOpenDynaset, dbForwardOnly)
            Do While Not rs.EOF
                If rs!MaSoCha > 0 Then
                    ExecuteSQL5 "UPDATE KQKD911 SET KyTruoc=KyTruoc+" + DoiDau(rs.Fields("KyTruoc")) + ",KyNay=KyNay+" + DoiDau(rs!KyNay) + " WHERE MaSo=" + CStr(rs!MaSoCha) + " AND MaTK=" + CStr(rs!MaTK)
                Else
                    ExecuteSQL5 "UPDATE KQKD911 SET KyTruoc=KyTruoc-" + DoiDau(rs.Fields("KyTruoc")) + ",KyNay=KyNay-" + DoiDau(rs!KyNay) + " WHERE MaSo=" + CStr(-rs!MaSoCha) + " AND MaTK=" + CStr(rs!MaTK)
                End If
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
            
            ExecuteSQL5 "UPDATE KQKD911 SET TongHop=0 WHERE MaTK>0"
            SetSQL "QKqkd", "SELECT MaSo, First(Ten" + IIf(nn > 0, "E", "") + ") AS FirstOfTen, First(TongHop) AS FirstOfTongHop, Sum(KyTruoc) AS SumOfKyTruoc, Sum(KyNay) AS SumOfKyNay,First(Cap) AS C From KQKD911 GROUP BY MaSo, MaTK"
        Else
            
            SetSQL "QKqkd", "SELECT MaSo, First(Ten" + IIf(nn > 0, "E", "") + ") AS FirstOfTen, First(TongHop) AS FirstOfTongHop, Sum(KyTruoc) AS SumOfKyTruoc, Sum(KyNay) AS SumOfKyNay,First(Cap) AS C From KQKD GROUP BY MaSo"
        End If
        
        frmMain.Rpt.ReportFileName = "KQKD.RPT"
        RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
        frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
End Sub

Public Sub XemBaoCao(sh As String, ndau As Date, ncuoi As Date)
    Dim sql As String, rs As Recordset
    
    ExecuteSQL5 "UPDATE BaoCao SET KyNay=0, KyTruoc=0 WHERE SoHieu='" + sh + "'"
    
    sql = "SELECT DISTINCTROW BaoCao.MaSo, Sum(IIF(" + WNgay2("ThangCT", 0, ndau) + ",ChungTu.SoPS,0)) AS Kqua1, Sum(IIF(" + WNgay("ThangCT", ndau, 0) + ",ChungTu.SoPS,0)) AS Kqua2 " _
        & " FROM (" + ChungTu2TKNC(0) + ") INNER JOIN BaoCao ON HethongTK.SoHieu LIKE BaoCao.SHTKNo+'*' AND TK.SoHieu LIKE BaoCao.SHTKCo+'*' " _
        & " Where " + WNgay("ThangCT", 0, ncuoi) + " AND BaoCao.SoHieu='" + sh + "' GROUP BY BaoCao.MaSo"
    
    sql = sql + " UNION SELECT DISTINCTROW BaoCao.MaSo, Sum(IIF(" + WNgay2("ThangCT", 0, ndau) + ",ChungTu.SoPS,0)) AS Kqua1, Sum(IIF(" + WNgay("ThangCT", ndau, 0) + ",ChungTu.SoPS,0)) AS Kqua2" _
        & " FROM (" + ChungTu2TKNC(1) + ") INNER JOIN BaoCao ON HethongTK.SoHieu LIKE BaoCao.SHTKCo+'*'" _
        & " Where (MaLoai<>3 AND MaLoai<>1) And " + WNgay("ThangCT", 0, ncuoi) + " And (CLng(BaoCao.SHTKNo) = 0) AND BaoCao.SoHieu='" + sh + "' GROUP BY BaoCao.MaSo"
    
    sql = sql + " UNION SELECT DISTINCTROW BaoCao.MaSo, Sum(IIF(" + WNgay2("ThangCT", 0, ndau) + ",ChungTu.SoPS,0)) AS Kqua1, Sum(IIF(" + WNgay("ThangCT", ndau, 0) + ",ChungTu.SoPS,0)) AS Kqua2" _
        & " FROM (" + ChungTu2TKNC(1) + ") INNER JOIN BaoCao ON HethongTK.SoHieu LIKE BaoCao.SHTKNo+'*'" _
        & " Where (MaLoai<>3 AND MaLoai<>1) And " + WNgay("ThangCT", 0, ncuoi) + " And (CLng(BaoCao.SHTKCo) = 0) AND BaoCao.SoHieu='" + sh + "' GROUP BY BaoCao.MaSo"

    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
            ExecuteSQL5 "UPDATE KQKD BaoCao KyTruoc=" + DoiDau(IIf(IsNull(rs!kqua1), 0, rs!kqua1)) + ", KyNay = " + DoiDau(IIf(IsNull(rs!kqua2), 0, rs!kqua2)) + " WHERE MaSo=" + CStr(rs!MaSo)
            rs.MoveNext
    Loop
    
    sql = "SELECT DISTINCTROW BaoCao.MaSo, Sum(IIF(" + WNgay2("ThangCT", 0, ndau) + ",ChungTu.SoPS,0)) AS Kqua1, Sum(IIF(" + WNgay("ThangCT", ndau, 0) + ",ChungTu.SoPS,0)) AS Kqua2 " _
        & " FROM (" + ChungTu2TKNC(0) + ") INNER JOIN BaoCao ON HethongTK.SoHieu LIKE BaoCao.SHTKCo+'*' AND TK.SoHieu LIKE BaoCao.SHTKNo+'*' " _
        & " Where " + WNgay("ThangCT", 0, ncuoi) + " AND BaoCao.SoHieu='" + sh + "' GROUP BY KQKD.MaSo"
        
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
            ExecuteSQL5 "UPDATE BaoCao SET KyTruoc = KyTruoc - " + DoiDau(IIf(IsNull(rs!kqua1), 0, rs!kqua1)) + ", KyNay = KyNay - " + DoiDau(IIf(IsNull(rs!kqua2), 0, rs!kqua2)) + " WHERE MaSo=" + CStr(rs!MaSo)
            rs.MoveNext
    Loop
            
    ' Cong don
    Set rs = DBKetoan.OpenRecordset("SELECT KyTruoc,KyNay,TongHop FROM BaoCao WHERE BaoCao.SoHieu='" + sh + "' AND TongHop>0 GROUP BY TongHop", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        ExecuteSQL5 "UPDATE BaoCao SET KyTruoc=KyTruoc+" + DoiDau(rs!KyTruoc) + ",KyNay=KyNay+" + DoiDau(rs!KyNay) + " WHERE MaSo=" + CStr(rs!TongHop)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
End Sub

Public Sub KtraCongNo()
    Dim sql As String, rs As Recordset, m As Long
    
    SetSQL "MienTru", "SELECT MaCT FROM ChungTu GROUP BY MaCT"
    sql = "SELECT ChungTu.MaSo, ChungTu.SH1, ChungTu.T1, ChungTu.MaKH, ChungTu.MaKHC FROM ChungTu LEFT JOIN MienTru ON ABS(ChungTu.CT_ID)=MienTru.MaCT WHERE (ChungTu.MaKH>0 OR ChungTu.MaKHC>0) AND IsNull(MienTru.MaCT) AND ChungTu.CT_ID<0 AND Abs(ChungTu.CT_ID)<2000000000"
    
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        m = SelectSQL("SELECT MaCT AS F1 FROM ChungTu WHERE ThangCT=" + CStr(rs!T1) + " AND SoHieu='" + rs!sh1 + "' AND ((MaKH>0 AND MaKH=" + CStr(rs!MaKHC) + ") OR (MaKHC>0 AND MaKHC=" + CStr(rs!makh) + "))")
        ExecuteSQL5 "UPDATE ChungTu SET CT_ID=" + CStr(-m) + " WHERE MaSo=" + CStr(rs!MaSo)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub LaySHCT()
    ExecuteSQL5 "UPDATE ChungTu INNER JOIN ChungTu AS CT ON ABS(ChungTu.CT_ID)=CT.MaCT SET ChungTu.SH1=CT.SoHieu,ChungTu.T1=CT.ThangCT WHERE (ChungTu.MaKH>0 OR ChungTu.MaKHC>0) AND ChungTu.CT_ID<0 AND Abs(ChungTu.CT_ID)<2000000000"
    KtraCongNo
    KiemTraTaiKhoan
End Sub

Public Sub LockDB()
    Dim T As Long, k As Integer
    
    Do While True
        k = SelectSQL("SELECT Lock12 AS F1 FROM License") Mod 100
        T = T + 1
        If T > 5000 Or k \ 10 = 0 Then Exit Do
        DoEvents
    Loop
    
    ExecuteSQL5 "UPDATE License SET Lock12=10+ Lock12 Mod 10 + Lock12 \100"
End Sub

Public Sub UnlockDB()
    ExecuteSQL5 "UPDATE License SET Lock12= Lock12 Mod 10 + Lock12 \100"
End Sub

Private Sub ktracongno2004()
    Dim rs As Recordset, tien As Double, loai As Integer, ms As Long, ps As Double
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaTaiKhoan, MaKhachHang,IIF(Kieu<0,SoDuKhachHang.DuNo_12-SoDuKhachHang.DuCo_12,SoDuKhachHang.DuCo_12-SoDuKhachHang.DuNo_12) AS SoDu FROM SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo", dbOpenSnapshot)
    Do While Not rs.EOF
        If rs!sodu <> SoNoTheoHoaDon(rs!MaTaiKhoan, rs!MaKhachHang, loai) Then
            ExecuteSQL5 "UPDATE CNDauNam SET SoXuat=Abs(DuNo_0-DuCo_0) WHERE MaTaiKhoan=" + CStr(rs!MaTaiKhoan) + " AND MaKhachHang=" + CStr(rs!MaKhachHang)
            ExecuteSQL5 "UPDATE ChungTu SET SoXuat=SoPS WHERE " + IIf(loai < 0, "MaTKNo=" + CStr(rs!MaTaiKhoan) + " AND MaKH=" + CStr(rs!MaKhachHang), "MaTKCo=" + CStr(rs!MaTaiKhoan) + " AND MaKHC=" + CStr(rs!MaKhachHang))
            tien = 0
            Do While tien < rs!sodu
                ms = SelectSQL("SELECT TOP 1 MaSo AS F1,SoPS AS F2 FROM ChungTu WHERE " + IIf(loai < 0, "MaTKNo=" + CStr(rs!MaTaiKhoan) + " AND MaKH=" + CStr(rs!MaKhachHang), "MaTKCo=" + CStr(rs!MaTaiKhoan) + " AND MaKHC=" + CStr(rs!MaKhachHang)) + " AND SoPS<=SoXuat AND SoPS>0 ORDER BY NgayGS DESC, MaCT DESC", ps)
                If ms > 0 Then
                    If ps > rs!sodu - tien Then
                        ExecuteSQL5 "UPDATE ChungTu SET SoXuat=SoPS-" + DoiDau(rs!sodu - tien) + " WHERE MaSo=" + CStr(ms)
                        tien = rs!sodu
                    Else
                        ExecuteSQL5 "UPDATE ChungTu SET SoXuat=0 WHERE MaSo=" + CStr(ms)
                        tien = tien + ps
                    End If
                Else
                    ms = SelectSQL("SELECT TOP 1 MaSo AS F1,Abs(DuNo_0-DuCo_0) AS F2 FROM CNDauNam WHERE MaTaiKhoan=" + CStr(rs!MaTaiKhoan) + " AND MaKhachHang=" + CStr(rs!MaKhachHang) + " AND Abs(DuNo_0-DuCo_0)<=SoXuat AND DuNo_0-DuCo_0<>0 ORDER BY NgayGS DESC, MaCT DESC", ps)
                    If ms > 0 Then
                        If ps > rs!sodu - tien Then
                            ExecuteSQL5 "UPDATE CNDauNam SET SoXuat=SoXuat-" + DoiDau(rs!sodu - tien) + " WHERE MaSo=" + CStr(ms)
                            tien = rs!sodu
                        Else
                            ExecuteSQL5 "UPDATE CNDauNam SET SoXuat=Abs(DuNo_0-DuCo_0) WHERE MaSo=" + CStr(ms)
                            tien = tien + ps
                        End If
                    Else
                        Exit Do
                    End If
                End If
            Loop
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Sub LaySoDauNam(fn As String)
    Dim db As Database, Err As Integer, sql As String, st As String, ms As Long
    
    On Error GoTo KT
    Set db = WSpace.OpenDatabase(fn, False, False, ";PWD=" + pPSW)
    On Error Resume Next
    
    If SelectSQLDB(db, "SELECT NamTC AS F1,TenCty AS F2, TenCn AS F3 FROM License", sql, st) <> pNamTC - 1 Then
        Err = 1
        GoTo KT
    End If
    
    If sql <> frmMain.LbCty(0).Caption Or st <> frmMain.LbCty(1).Caption Then
        If MsgBox("Tªp d÷ liÖu kh¸c tªn c«ng ty hoÆc tªn chi nh¸nh, tiÕp tôc ?", vbCritical + vbYesNo, App.ProductName) = vbNo Then GoTo KT
    End If
    
    CongDK db, "", "", 12, 0, ""
    CongDDCN db, 0, 0, False, 0, ""
    CongDDVT db, 0, "", 0, False, 0, ""
    CongDDTS db, 0, "", 0, False, 0, ""
    
    db.Close
    Set db = Nothing
        
    If OutCost <> 0 Or pCongNoHD <> 0 Then
        If SetPsw(fn, pPSW, "") Then
            If OutCost <> 0 Then
                ExecuteSQL5 "DELETE * FROM VTDauNam"
                
                ExecuteSQL5 "INSERT INTO VTDauNam (MaSo,MaSoKho,MaTaiKhoan,MaVattu,Luong_0,Tien_0) " _
                    & " SELECT DISTINCTROW MaSo,MaSoKho,MaTaiKhoan,MaVattu,Luong_0-SoXuat," + RoundMoneySQL("Tien_0*(Luong_0-SoXuat)/Luong_0") + " FROM VTDauNam IN '" + fn + "' WHERE Luong_0>SoXuat"
                    
                ExecuteSQL5 "INSERT INTO VTDauNam (MaSo,MaSoKho,MaTaiKhoan,MaVattu,Luong_0,Tien_0) " _
                    & " SELECT DISTINCTROW MaSo,MaKho,MaTKNo,MaVattu,SoPS2No-SoXuat," + RoundMoneySQL("SoPS*(SoPS2No-SoXuat)/SoPS2No") + " FROM ChungTu IN '" + fn + "' WHERE MaLoai=1 AND MaVattu>0 AND SoPS2No>SoXuat"
            End If
            
            If pCongNoHD <> 0 Then
                ms = Lng_MaxValue("MaSo", "CNDauNam")
                ExecuteSQL5 "INSERT INTO CNDauNam (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0) " _
                    & " SELECT DISTINCTROW MaSo+" + CStr(ms) + ",MaTaiKhoan,MaKhachHang,CNDauNam.DuNo_0-CNDauNam.SoXuat,DuCo_0 FROM CNDauNam INNER JOIN HethongTK ON CNDauNam=HethongTK.MaSo IN '" + fn + "' WHERE DuNo_0>SoXuat AND TK_ID=" + CStr(TKCNKH_ID)
                ExecuteSQL5 "INSERT INTO CNDauNam (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0) " _
                    & " SELECT DISTINCTROW MaSo+" + CStr(ms) + ",MaTaiKhoan,MaKhachHang,CNDauNam.DuNo_0,DuCo_0-CNDauNam.SoXuat FROM CNDauNam INNER JOIN HethongTK ON CNDauNam=HethongTK.MaSo IN '" + fn + "' WHERE DuNo_0>SoXuat AND TK_ID=" + CStr(TKCNPT_ID)
            
                ExecuteSQL5 "INSERT INTO CNDauNam (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0) " _
                    & " SELECT DISTINCTROW MaSo+" + CStr(ms) + ",MaTKNo,MaKH,SoPS-SoXuat,0 FROM " + ChungTu2TKNC(-1) + " IN '" + fn + "' WHERE MaKH>0 AND SoPS>SoXuat AND TK_ID=" + CStr(TKCNKH_ID)
                ExecuteSQL5 "INSERT INTO CNDauNam (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0) " _
                    & " SELECT DISTINCTROW MaSo+" + CStr(ms) + ",MaTKCo,MaKHC,0,SoPS-SoXuat FROM " + ChungTu2TKNC(1) + " IN '" + fn + "' WHERE MaKHC>0 AND SoPS>SoXuat AND TK_ID=" + CStr(TKCNPT_ID)
            End If
            SetPsw fn, "", pPSW
        End If
    End If
    KiemTraVatTu
    KiemTraTaiKhoan
    
    CloseItemList
    
    Exit Sub
KT:
    Select Case Err
        Case 0:     sql = "Kh«ng më ®­îc tÖp d÷ liÖu !"
        Case 1:     sql = "TÖp kh«ng ®óng n¨m tµi chÝnh " + CStr(pNamTC - 1)
        Case Else:  sql = ""
    End Select
    If Len(sql) > 0 Then MsgBox sql, vbCritical, App.ProductName
    If Err > 0 Then db.Close
    Set db = Nothing
End Sub

Public Sub CongDDVT(FrmDB As Database, masocu As Long, tencn As String, tachsh As Integer, cnct As Boolean, pTK As Integer, pCT As String)
    Dim rs As Recordset
    Dim mcha As Long, sh As String, i As Integer
        
    If cnct Then ExecuteSQL5 "UPDATE ChungTu SET U=0"
    Set rs = FrmDB.OpenRecordset("SELECT * FROM KhoHang ORDER BY TenKho", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        sh = rs!tenkho
        i = 0
        mcha = SelectSQL("SELECT MaSo AS F1 FROM KhoHang WHERE TenKho='" + sh + "'")
        If mcha > 0 And tachsh > 0 Then
            Do While mcha > 0
                i = i + 1
                sh = sh + Chr(64 + i)
                mcha = SelectSQL("SELECT MaSo AS F1 FROM KhoHang WHERE TenKho='" + sh + "'")
            Loop
        End If
        If mcha = 0 Then
            ExecuteSQL5 "INSERT INTO KhoHang (MaSo,TenKho) VALUES (" + CStr(Lng_MaxValue("MaSo", "TenKho") + 1) + ",'" + sh + "')"
            mcha = SelectSQL("SELECT MaSo AS F1 FROM KhoHang WHERE TenKho='" + sh + "'")
        End If
        If cnct Then ExecuteSQL5 "UPDATE ChungTu SET MaKho=" + CStr(mcha) + ",U=1 WHERE U=0 AND MaKho=" + CStr(rs!MaSo) + " AND MaLoai<>4 AND MaKho>0 AND MaSo>" + CStr(masocu)
        If cnct Then ExecuteSQL5 "UPDATE ChungTu SET MaNguon=" + CStr(mcha) + ",U=1 WHERE U=0 AND MaNguon=" + CStr(rs!MaSo) + " AND MaLoai=4 AND MaSo>" + CStr(masocu)
        rs.MoveNext
    Loop
    
    If cnct Then ExecuteSQL5 "UPDATE ChungTu SET U=0"
    Set rs = FrmDB.OpenRecordset("SELECT * FROM NguonNhapXuat ORDER BY SoHieu", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        sh = rs!sohieu
        i = 0
        mcha = SoHieu2MaSo(sh, "NguonNhapXuat")
        If mcha > 0 And tachsh > 0 Then
            Do While mcha > 0
                i = i + 1
                sh = sh + Chr(64 + i)
                mcha = SoHieu2MaSo(sh, "NguonNhapXuat")
            Loop
        End If
        If mcha = 0 Then
            ExecuteSQL5 "INSERT INTO NguonNhapXuat (MaSo,SoHieu,DienGiai) VALUES (" + CStr(Lng_MaxValue("MaSo", "NguonNhapXuat") + 1) + ",'" + sh + "','" + rs!diengiai + "-" + CStr(i) + "')"
            mcha = SoHieu2MaSo(sh, "NguonNhapXuat")
        End If
        If cnct Then ExecuteSQL5 "UPDATE ChungTu SET MaNguon=" + CStr(mcha) + ",U=1 WHERE U=0 AND MaNguon=" + CStr(rs!MaSo) + " AND MaLoai<>4 AND MaNguon>0 AND MaSo>" + CStr(masocu)
        rs.MoveNext
    Loop
    
    Set rs = FrmDB.OpenRecordset("SELECT PhanLoaiVattu.*,PL.SoHieu AS SHCha,HethongTK.SoHieu AS SHTK FROM (PhanLoaiVattu INNER JOIN HethongTK ON PhanLoaiVattu.MaTK=HethongTK.MaSo) LEFT JOIN PhanLoaiVattu AS PL ON PhanLoaiVattu.PLCha=PL.MaSo ORDER BY PhanLoaiVattu.SoHieu", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        sh = rs!sohieu
        i = 0
        mcha = SoHieu2MaSo(sh, "PhanLoaiVattu")
        If mcha > 0 And tachsh > 0 Then
            Do While mcha > 0
                i = i + 1
                sh = sh + Chr(64 + i)
                mcha = SoHieu2MaSo(sh, "PhanLoaiVattu")
            Loop
        End If
        If mcha = 0 Then
            If rs!plcha > 0 Then
                mcha = SoHieu2MaSo(rs!shcha, "PhanLoaiVattu")
            Else
                mcha = 0
            End If
            ExecuteSQL5 "INSERT INTO PhanLoaiVattu (MaSo,SoHieu,TenPhanLoai,PLCon,PLCha,Cap,MaTK) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhanLoaiVattu") + 1) + ",'" + sh + "','" + rs!TenPhanLoai + "'," + CStr(rs!plcon) + "," + CStr(mcha) + "," + CStr(rs!cap) + "," + CStr(MaTKFix(FrmDB, rs!shtk, IIf(pTK > 0, pCT, ""), tencn)) + ")"
        End If
        rs.MoveNext
    Loop
    
    If cnct Then ExecuteSQL5 "UPDATE ChungTu SET U=0"
    Set rs = FrmDB.OpenRecordset("SELECT Vattu.*,PhanLoaiVattu.SoHieu AS SHPL FROM Vattu INNER JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo ORDER BY Vattu.SoHieu", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        sh = rs!sohieu
        i = 0
        mcha = SoHieu2MaSo(sh, "Vattu")
        If mcha > 0 And tachsh > 0 Then
            Do While mcha > 0
                i = i + 1
                sh = sh + Chr(64 + i)
                mcha = SoHieu2MaSo(sh, "Vattu")
            Loop
        End If
        If mcha = 0 Then
            ExecuteSQL5 "INSERT INTO Vattu (MaSo,MaPhanLoai,SoHieu,TenVattu,DonVi,TonMin,TonMax,GiaDuPhong,GiaHT,dvt2,DonVi2,TyLeQD,GhiChu) VALUES (" + CStr(Lng_MaxValue("MaSo", "Vattu") + 1) + "," + CStr(SoHieu2MaSo(rs!shpl, "PhanLoaiVattu")) _
                + ",'" + sh + "','" + rs!TenVattu + "','" + CStr(rs!DonVi) + "'," + DoiDau(rs!TonMin) + "," + DoiDau(rs!TonMax) + "," + DoiDau(rs!GiaDuPhong) + "," + DoiDau(rs!GiaHT) + "," + CStr(rs!Dvt2) + ",'" + rs!DonVi2 + "'," + DoiDau(rs!TyLeQD) + ",'" + rs!GhiChu + "')"
            mcha = SoHieu2MaSo(sh, "Vattu")
        End If
        If cnct Then ExecuteSQL5 "UPDATE ChungTu SET MaVattu=" + CStr(mcha) + ",U=1 WHERE U=0 AND MaVattu=" + CStr(rs!MaSo) + " AND MaSo>" + CStr(masocu)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Sub CongDDCN(FrmDB As Database, masocu As Long, tachsh As Integer, cnct As Boolean, pTK As Integer, pCT As String)
    Dim rs As Recordset
    Dim mcha As Long, sh As String, i As Integer
        
    Set rs = FrmDB.OpenRecordset("SELECT PhanLoaiKhachHang.*,PL.SoHieu AS SHCha FROM PhanLoaiKhachHang LEFT JOIN PhanLoaiKhachHang AS PL ON PhanLoaiKhachHang.PLCha=PL.MaSo ORDER BY PhanLoaiKhachHang.SoHieu", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        sh = rs!sohieu
        i = 0
        mcha = SoHieu2MaSo(sh, "PhanLoaiKhachHang")
        If mcha > 0 And tachsh > 0 Then
            Do While mcha > 0
                i = i + 1
                sh = sh + Chr(64 + i)
                mcha = SoHieu2MaSo(sh, "PhanLoaiKhachHang")
            Loop
        End If
        If mcha = 0 Then
            If rs!plcha > 0 And (Not IsNull(rs!shcha)) Then
                mcha = SoHieu2MaSo(rs!shcha, "PhanLoaiKhachHang")
            Else
                mcha = 0
            End If
            ExecuteSQL5 "INSERT INTO PhanLoaiKhachHang (MaSo,SoHieu,TenPhanLoai,PLCon,PLCha,Cap) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhanLoaiKhachHang") + 1) + ",'" + sh + "','" + rs!TenPhanLoai + "'," + CStr(rs!plcon) + "," + CStr(mcha) + "," + CStr(rs!cap) + ")"
        End If
        rs.MoveNext
    Loop
    
    If cnct Then ExecuteSQL5 "UPDATE ChungTu SET U=0, U2=0"
    Set rs = FrmDB.OpenRecordset("SELECT KhachHang.*,PhanLoaiKhachHang.SoHieu AS SHPL FROM KhachHang INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo ORDER BY KhachHang.SoHieu", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        sh = rs!sohieu
        i = 0
        mcha = SoHieu2MaSo(sh, "KhachHang")
        If mcha > 0 And tachsh > 0 Then
            Do While mcha > 0
                i = i + 1
                sh = sh + Chr(64 + i)
                mcha = SoHieu2MaSo(sh, "KhachHang")
            Loop
        End If
        If mcha = 0 Then
            ExecuteSQL5 "INSERT INTO KhachHang (MaSo,MaPhanLoai,SoHieu,Ten,DiaChi,MST,Tel,Fax,EMail,DaiDien,TaiKhoan,DuMax,GhiChu,MaNT) VALUES (" + CStr(Lng_MaxValue("MaSo", "KhachHang") + 1) + "," _
                + CStr(SoHieu2MaSo(rs!shpl, "PhanLoaiKhachHang")) + ",'" + sh + "','" + rs!Ten + "','" + rs!DiaChi + "','" + rs!mst + "','" + rs!Tel + "','" + rs!Fax + "','" + rs!email + "','" + rs!DaiDien + "','" _
                + rs!taikhoan + "'," + DoiDau(rs!DuMax) + ",'" + rs!GhiChu + "'," + CStr(rs!MaNT) + ")"
            mcha = SoHieu2MaSo(sh, "KhachHang")
        End If
        If cnct Then
            ExecuteSQL5 "UPDATE ChungTu SET MaKH=" + CStr(mcha) + ", U=1 WHERE U=0 AND MaKH=" + CStr(rs!MaSo) + " AND MaSo>" + CStr(masocu)
            ExecuteSQL5 "UPDATE ChungTu SET MaKHC=" + CStr(mcha) + ", U2=1 WHERE U2=0 AND MaKHC=" + CStr(rs!MaSo) + " AND MaSo>" + CStr(masocu)
            ExecuteSQL5 "UPDATE HoaDon SET MaKhachHang=" + CStr(mcha) + " WHERE MaKH1=" + CStr(rs!MaSo) + " AND MaSo>" + CStr(masocu)
        End If
        rs.MoveNext
    Loop
        
    rs.Close
    Set rs = Nothing
End Sub

Public Sub CongDDTS(FrmDB As Database, mactcu As Long, tencn As String, tachsh As Integer, cnct As Boolean, pTK As Integer, pCT As String)
    Dim rs As Recordset, rs2 As Recordset
    Dim mcha As Long, m1 As Long, matscu As Long, i As Integer, sh As String
    
    If cnct Then ExecuteSQL5 "UPDATE CTTaiSan SET U=0, U2=0"
    matscu = Lng_MaxValue("MaSo", "TaiSan")
    Set rs = FrmDB.OpenRecordset("SELECT LoaiChungTu.*,PL.SoHieu AS SHCha FROM LoaiChungTu LEFT JOIN LoaiChungTu AS PL ON LoaiChungTu.CapTren=PL.MaSo ORDER BY LoaiChungTu.SoHieu", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        sh = rs!sohieu
        mcha = SoHieu2MaSo(sh, "LoaiChungTu")
        If mcha = 0 Then
            If rs!CapTren > 0 Then
                m1 = SoHieu2MaSo(rs!shcha, "LoaiChungTu")
            Else
                m1 = 0
            End If
            ExecuteSQL5 "INSERT INTO LoaiChungTu (MaSo,SoHieu,Ten,CapTren,Cap) VALUES (" + CStr(Lng_MaxValue("MaSo", "LoaiChungTu") + 1) + ",'" + sh + "','" + rs!Ten + "'," + CStr(m1) + "," + CStr(rs!cap) + ")"
            mcha = SoHieu2MaSo(sh, "LoaiChungTu")
        End If
        If cnct Then
            ExecuteSQL5 "UPDATE CTTaiSan SET MaLoai=" + CStr(mcha) + ", U=1 WHERE U=0 AND ML=" + CStr(rs!MaSo) + " AND MaCTKT>" + CStr(mactcu)
            ExecuteSQL5 "UPDATE CTTaiSan SET MaNhom=" + CStr(mcha) + ", U2=1 WHERE U2=0 AND MN=" + CStr(rs!MaSo) + " AND MaCTKT>" + CStr(mactcu)
        End If
        rs.MoveNext
    Loop
    
    Set rs = FrmDB.OpenRecordset("SELECT LoaiTaiSan.*,PL.SoHieu AS SHCha FROM LoaiTaiSan LEFT JOIN LoaiTaiSan AS PL ON LoaiTaiSan.CapTren=PL.MaSo ORDER BY LoaiTaiSan.SoHieu", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        sh = rs!sohieu
        i = 0
        mcha = SoHieu2MaSo(sh, "LoaiTaiSan")
        If mcha > 0 And tachsh > 0 Then
            Do While mcha > 0
                i = i + 1
                sh = sh + Chr(64 + i)
                mcha = SoHieu2MaSo(sh, "LoaiTaiSan")
            Loop
        End If
        If mcha = 0 Then
            If rs!CapTren > 0 Then
                m1 = SoHieu2MaSo(rs!shcha, "LoaiTaiSan")
            Else
                m1 = 0
            End If
            If ExecuteSQL5("INSERT INTO LoaiTaiSan (MaSo,SoHieu,Ten,CapTren,Cap) VALUES (" + CStr(Lng_MaxValue("MaSo", "LoaiTaiSan") + 1) + ",'" + sh + "','" + rs!Ten + "'," + CStr(m1) + "," + CStr(rs!cap) + ")", False) <> 0 Then
                i = 1
                Do While ExecuteSQL5("INSERT INTO LoaiTaiSan (MaSo,SoHieu,Ten,CapTren,Cap) VALUES (" + CStr(Lng_MaxValue("MaSo", "LoaiTaiSan") + 1) + ",'" + sh + "','" + rs!Ten + "-" + CStr(i) + "'," + CStr(m1) + "," + CStr(rs!cap) + ")", False) <> 0
                    i = i + 1
                Loop
            End If
        End If
        rs.MoveNext
    Loop
    
    Set rs = FrmDB.OpenRecordset("SELECT * FROM QuocGia ORDER BY Ten", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        mcha = SelectSQL("SELECT MaSo AS F1 FROM QuocGia WHERE Ten='" + rs!Ten + "'")
        If mcha = 0 Then
            ExecuteSQL5 "INSERT INTO QuocGia (MaSo,Ten) VALUES (" + CStr(Lng_MaxValue("MaSo", "QuocGia") + 1) + ",'" + rs!Ten + "')"
        End If
        rs.MoveNext
    Loop
    
    Set rs = FrmDB.OpenRecordset("SELECT * FROM DTQLy ORDER BY Ten", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        mcha = SelectSQL("SELECT MaSo AS F1 FROM DTQLy WHERE Ten='" + rs!Ten + "'")
        If mcha = 0 Then
            ExecuteSQL5 "INSERT INTO DTQLy (MaSo,Ten) VALUES (" + CStr(Lng_MaxValue("MaSo", "DTQLy") + 1) + ",'" + rs!Ten + "')"
        End If
        rs.MoveNext
    Loop
    
    Set rs = FrmDB.OpenRecordset("SELECT * FROM TinhTrang ORDER BY Ten", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        mcha = SelectSQL("SELECT MaSo AS F1 FROM TinhTrang WHERE Ten='" + rs!Ten + "'")
        If mcha = 0 Then
            ExecuteSQL5 "INSERT INTO TinhTrang (MaSo,Ten) VALUES (" + CStr(Lng_MaxValue("MaSo", "TinhTrang") + 1) + ",'" + rs!Ten + "')"
        End If
        rs.MoveNext
    Loop
    If cnct Then
        ExecuteSQL5 "UPDATE CTTaiSan SET U=0"
        FrmDB.Execute "UPDATE TaiSan SET NamKH=0 WHERE IsNull(NamKH)"
    End If
    Set rs = FrmDB.OpenRecordset("SELECT TaiSan.*,LoaiTaiSan.SoHieu AS SHT,L1.SoHieu AS SHL, L2.SoHieu AS SHN,QuocGia.Ten AS QG FROM (((TaiSan INNER JOIN QuocGia ON TaiSan.MaNuoc=QuocGia.MaSo) INNER JOIN LoaiTaiSan ON TaiSan.MaTaiKhoan=LoaiTaiSan.MaSo) INNER JOIN LoaiTaiSan AS L1 ON TaiSan.MaLoai=L1.MaSo) LEFT JOIN LoaiTaiSan AS L2 ON TaiSan.MaNhom=L2.MaSo ORDER BY TaiSan.SoHieu", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        sh = rs!sohieu
        i = 0
        mcha = SoHieu2MaSo(sh, "TaiSan")
        If mcha > 0 And tachsh > 0 Then
            Do While mcha > 0
                i = i + 1
                sh = sh + Chr(64 + i)
                mcha = SoHieu2MaSo(sh, "TaiSan")
            Loop
        End If
        If mcha = 0 Then
            ExecuteSQL5 "INSERT INTO TaiSan (MaSo,Ten, SoHieu, NangLuc, GhiChu, MaNuoc, NamSX, " _
                & "NamSD, MaTaiKhoan, MaLoai, MaNhom, ThangTang, ThangGiam,NamKH) VALUES (" + CStr(Lng_MaxValue("MaSo", "TaiSan") + 1) + ",'" _
                + rs!Ten + "','" + sh + "','" + rs!NangLuc + "','" + rs!GhiChu + "'," + CStr(ST2MaSo("Ten", rs!QG, "QuocGia")) + "," _
                + CStr(rs!NamSX) + "," + CStr(rs!NamSD) + "," + CStr(SoHieu2MaSo(rs!SHT, "LoaiTaiSan")) + "," + CStr(SoHieu2MaSo(rs!SHL, "LoaiTaiSan")) _
                + "," + CStr(SoHieu2MaSo(IIf(IsNull(rs!SHN), "", rs!SHN), "LoaiTaiSan")) + "," + CStr(rs!ThangTang) + "," + CStr(rs!ThangGiam) + "," + CStr(rs!NamKH) + ")"
            mcha = SoHieu2MaSo(sh, "TaiSan")
            Set rs2 = FrmDB.OpenRecordset("SELECT CTTaiSan.*,LoaiChungTu.SoHieu AS SHL,LCT.SoHieu AS SHN FROM (CTTaiSan INNER JOIN LoaiChungTu ON CTTaiSan.MaLoai=LoaiChungTu.MaSo) INNER JOIN LoaiChungTu AS LCT ON CTTaiSan.MaNhom=LCT.MaSo WHERE (Thang=0) AND MaTS=" + CStr(rs!MaSo), dbOpenSnapshot)
            If rs2.recordCount > 0 Then
                ExecuteSQL5 "INSERT INTO CTTaiSan (MaSo,SoHieu, Thang, VaoSo, NgayGhi, DienGiai, " _
                    & "MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, " _
                    & "CL_NS, CL_TBS, CL_CNK, CL_TD)VALUES (" + CStr(Lng_MaxValue("MaSo", "CTTaiSan") + 1) + ",'" + rs2!sohieu + "'," + CStr(rs2!thang) _
                    + ",#" + Format(rs2!VaoSo, Mask_DB) + "#,#" + Format(rs2!NgayGhi, Mask_DB) + "#,'" _
                    + rs2!diengiai + "'," + CStr(SoHieu2MaSo(rs2!SHL, "LoaiChungTu")) + "," + CStr(SoHieu2MaSo(rs2!SHN, "LoaiChungTu")) + "," + CStr(mcha) + "," _
                    + DoiDau(rs2!NG_NS) + "," + DoiDau(rs2!NG_TBS) + "," + DoiDau(rs2!NG_CNK) + "," + DoiDau(rs2!NG_TD) + "," _
                    + DoiDau(rs2!CL_NS) + "," + DoiDau(rs2!CL_TBS) + "," + DoiDau(rs2!CL_CNK) + "," + DoiDau(rs2!CL_TD) + ")"
            End If
            rs2.Close
            Set rs2 = FrmDB.OpenRecordset("SELECT ThongSo.*,HethongTK.SoHieu AS SHTK,DTQLy.Ten AS QL,TinhTrang.Ten AS TT FROM ((ThongSo INNER JOIN HethongTK ON ThongSo.MaDTSD=HethongTK.MaSo) INNER JOIN DTQLy ON ThongSo.MaDTQL=DTQLy.MaSo) INNER JOIN TinhTrang ON ThongSo.MaTTSD=TinhTrang.MaSo WHERE MaTS=" + CStr(rs!MaSo), dbOpenSnapshot, dbForwardOnly)
            Do While Not rs2.EOF
                ExecuteSQL5 "INSERT INTO ThongSo ( MaSo, MaTS, Thang, NG_NS, NG_TBS, NG_CNK, NG_TD, CL_NS, CL_TBS, CL_CNK, CL_TD, " _
                    & "KH_NS, KH_TBS, KH_CNK, KH_TD, MaDTQL, MaDTSD, MaTTSD ) " _
                    & "VALUES (" + CStr(Lng_MaxValue("MaSo", "ThongSo") + 1) + "," + CStr(mcha) + "," + CStr(rs2!thang) + "," + DoiDau(rs2!NG_NS) + "," + DoiDau(rs2!NG_TBS) + "," + DoiDau(rs2!NG_CNK) + "," + DoiDau(rs2!NG_TD) _
                    + "," + DoiDau(rs2!CL_NS) + "," + DoiDau(rs2!CL_TBS) + "," + DoiDau(rs2!CL_CNK) + "," + DoiDau(rs2!CL_TD) + "," + DoiDau(rs2!KH_NS) + "," + DoiDau(rs2!KH_TBS) + "," + DoiDau(rs2!KH_CNK) + "," + DoiDau(rs2!KH_TD) _
                    + "," + CStr(ST2MaSo("Ten", rs2!QL, "DTQLy")) + "," + CStr(MaTKFix(FrmDB, rs2!shtk, IIf(pTK > 0, pCT, ""), tencn)) + "," + CStr(ST2MaSo("Ten", rs2!tt, "TinhTrang")) + ")"
                rs2.MoveNext
            Loop
        End If
        If cnct Then ExecuteSQL5 "UPDATE CTTaiSan SET MaTS=" + CStr(mcha) + ",U=1 WHERE U=0 AND MaTS=" + CStr(rs!MaSo) + " AND MaCTKT>" + CStr(mactcu)
        rs.MoveNext
    Loop
    
    On Error Resume Next
    rs2.Close
    On Error GoTo 0
    rs.Close
    Set rs = Nothing
    
    If cnct Then SoDuTKTS Else ChuyenNamMoiTS
End Sub

Public Sub CongDK(db As Database, s As String, tencn As String, thangdb As Integer, pTK As Integer, pCT As String)
Dim mk As Long, dkn As Double, dkc As Double, dknt As Double, thang As String
Dim rs As Recordset, rs2 As Recordset
    
    thang = CStr(thangdb)
    If thangdb = 0 Then ExecuteSQL5 "UPDATE HethongTK SET DuNo_0=DuNo_0-n" + s + ",DuCo_0=DuCo_0-c" + s + ",DuNT_0=DuNT_0-nt" + s + ",n" + s + "=0,c" + s + "=0,nt" + s + "=0"
    Set rs = db.OpenRecordset("SELECT SoHieu,Sum(DuNo_" + thang + ") AS ndk,Sum(DuCo_" + thang + ") AS cdk,Sum(DuNT_" + thang + ") AS dknt FROM HethongTK WHERE TKCon=0 GROUP BY SoHieu HAVING Sum(DuNo_" + thang + ")<>0 OR Sum(DuCo_" + thang + ")<>0 OR Sum(DuNT_" + thang + ")<>0", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        mk = MaTKFix(db, rs!sohieu, IIf(pTK > 0, pCT, IIf(thangdb = 0, "", "#")), tencn)
        If mk > 0 Then
            If thangdb = 0 Then
                Set rs2 = DBKetoan.OpenRecordset("SELECT DISTINCTROW DuNo_0 AS dkn,DuCo_0 AS dkc,DuNT_0 AS dknt FROM HethongTK WHERE MaSo=" + CStr(mk), dbOpenSnapshot, dbForwardOnly)
                
                dkn = rs2!dkn + rs!ndk
                dkc = rs2!dkc + rs!cdk
                dknt = rs2!dknt + rs!dknt
                
                If dkn >= dkc Then
                    dkn = dkn - dkc
                    dkc = 0
                Else
                    dkc = dkc - dkn
                    dkn = 0
                End If
                
                ExecuteSQL5 "UPDATE HethongTK SET DuNo_0=" + DoiDau(dkn) + ",DuCo_0=" + DoiDau(dkc) + ",DuNT_0=" + DoiDau(dknt) + ",n" + s + "=n" + s + "+" + DoiDau(rs!ndk) + ",c" + s + "=c" + s + "+" + DoiDau(rs!cdk) + ",nt" + s + "=nt" + s + "+" + DoiDau(rs!dknt) + " WHERE MaSo=" + CStr(mk)
                rs2.Close
            Else
                ExecuteSQL5 "UPDATE HethongTK SET DuNo_0=" + DoiDau(rs!ndk) + ",DuCo_0=" + DoiDau(rs!cdk) + ",DuNT_0=" + DoiDau(rs!dknt) + " WHERE MaSo=" + CStr(mk)
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set rs2 = Nothing
End Sub

Public Sub CongDKVT(db As Database, s As String, tencn As String, thangdb As Integer, pTK As Integer, pCT As String)
Dim mk As Long, mv As Long, m2 As Long, dkn As Double, dkc As Double, dknt As Double, thang As String
Dim rs As Recordset, rs2 As Recordset
    
    thang = CStr(thangdb)
    If thangdb = 0 Then ExecuteSQL5 "UPDATE TonKho SET Luong_0=Luong_0-n" + s + ",Tien_0=Tien_0-c" + s + ",n" + s + "=0,c" + s + "=0"
    Set rs = db.OpenRecordset("SELECT KhoHang.TenKho AS Kho,HethongTK.SoHieu AS SHTK, Vattu.SoHieu AS SHVT,Luong_" + thang + " AS ndk,Tien_" + thang + " AS cdk FROM ((TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo) INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo) INNER JOIN KhoHang ON TonKho.MaSoKho=KhoHang.MaSo WHERE Luong_" + thang + "<>0 OR Tien_" + thang + "<>0", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        mk = MaTKFix(db, rs!shtk, IIf(pTK > 0, pCT, ""), tencn)
        mv = SoHieu2MaSo(rs!shvt, "Vattu")
        m2 = ST2MaSo("TenKho", rs!kho, "KhoHang")
        If mk > 0 And mv > 0 And m2 > 0 Then
            Set rs2 = DBKetoan.OpenRecordset("SELECT DISTINCTROW Luong_0 AS dkn,Tien_0 AS dkc FROM TonKho WHERE MaTaiKhoan=" + CStr(mk) + " AND MaVattu=" + CStr(mv) + " AND MaSoKho=" + CStr(m2), dbOpenSnapshot)
            If rs2.recordCount > 0 Then
                If thangdb = 0 Then
                    dkn = rs2!dkn + rs!ndk
                    dkc = rs2!dkc + rs!cdk
                    ExecuteSQL5 "UPDATE TonKho SET Luong_0=" + DoiDau(dkn) + ",Tien_0=" + DoiDau(dkc) + ",n" + s + "=n" + s + "+" + DoiDau(rs!ndk) + ",c" + s + "=c" + s + "+" + DoiDau(rs!cdk) + " WHERE MaTaiKhoan=" + CStr(mk) + " AND MaVattu=" + CStr(mv) + " AND MaSoKho=" + CStr(m2)
                Else
                    ExecuteSQL5 "UPDATE TonKho SET Luong_0=" + DoiDau(rs!ndk) + ",Tien_0=" + DoiDau(rs!cdk) + " WHERE MaTaiKhoan=" + CStr(mk) + " AND MaVattu=" + CStr(mv) + " AND MaSoKho=" + CStr(m2)
                End If
            Else
                If thangdb = 0 Then
                    ExecuteSQL5 "INSERT INTO TonKho (MaSo,MaSoKho,MaTaiKhoan,MaVattu,Luong_0,Tien_0,n" + s + ",c" + s + ") VALUES (" + CStr(Lng_MaxValue("MaSo", "TonKho") + 1) + "," + CStr(m2) + "," + CStr(mk) + "," + CStr(mv) + "," + DoiDau(rs!ndk) + "," + DoiDau(rs!cdk) + "," + DoiDau(rs!ndk) + "," + DoiDau(rs!cdk) + ")"
                Else
                    ExecuteSQL5 "INSERT INTO TonKho (MaSo,MaSoKho,MaTaiKhoan,MaVattu,Luong_0,Tien_0) VALUES (" + CStr(Lng_MaxValue("MaSo", "TonKho") + 1) + "," + CStr(m2) + "," + CStr(mk) + "," + CStr(mv) + "," + DoiDau(rs!ndk) + "," + DoiDau(rs!cdk) + ")"
                End If
            End If
            rs2.Close
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set rs2 = Nothing
End Sub

Public Sub CongDKCN(db As Database, s As String, tencn As String, thangdb As Integer, pTK As Integer, pCT As String)
Dim mk As Long, m2 As Long, dkn As Double, dkc As Double, thang As String
Dim rs As Recordset, rs2 As Recordset, n As Double, c As Double, dknt As Double
    
    thang = CStr(thangdb)
    If thangdb = 0 Then ExecuteSQL5 "UPDATE SoDuKhachHang SET DuNo_0=DuNo_0-n" + s + ",DuCo_0=DuCo_0-c" + s + ",DuNT_0=DuNT_0-nt" + s + ",n" + s + "=0,c" + s + "=0,nt" + s + "=0"
    Set rs = db.OpenRecordset("SELECT HethongTK.SoHieu AS SHTK, KhachHang.SoHieu AS SHKH,SoDuKhachHang.DuNo_" + thang + " AS ndk,SoDuKhachHang.DuCo_" + thang + " AS cdk,SoDuKhachHang.DuNT_" + thang + " AS ntdk FROM (SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo) INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo WHERE SoDuKhachHang.DuNo_" + thang + "<>0 OR SoDuKhachHang.DuCo_" + thang + "<>0 OR SoDuKhachHang.DuNT_" + thang + "<>0", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        mk = MaTKFix(db, rs!shtk, IIf(pTK > 0, pCT, ""), tencn)
        m2 = SoHieu2MaSo(rs!shkh, "KhachHang")
        If mk > 0 And m2 > 0 Then
            Set rs2 = DBKetoan.OpenRecordset("SELECT DISTINCTROW DuNo_0 AS dkn,DuCo_0 AS dkc,DuNT_0 AS dknt FROM SoDuKhachHang WHERE MaTaiKhoan=" + CStr(mk) + " AND MaKhachHang=" + CStr(m2), dbOpenSnapshot)
            If rs2.recordCount > 0 Then
                If thangdb = 0 Then
                    dkn = rs2!dkn + rs!ndk
                    dkc = rs2!dkc + rs!cdk
                    dknt = rs2!dknt + rs!ntdk
                    
                    If dkn >= dkc Then
                        dkn = dkn - dkc
                        dkc = 0
                    Else
                        dkc = dkc - dkn
                        dkn = 0
                    End If
                    
                    ExecuteSQL5 "UPDATE SoDuKhachHang SET DuNo_0=" + DoiDau(dkn) + ",DuCo_0=" + DoiDau(dkc) + ",DuNT_0=" + DoiDau(dknt) + ",n" + s + "=n" + s + "+" + DoiDau(rs!ndk) + ",c" + s + "=c" + s + "+" + DoiDau(rs!cdk) + ",nt" + s + "=nt" + s + "+" + DoiDau(rs!ntdk) + " WHERE MaTaiKhoan=" + CStr(mk) + " AND MaKhachHang=" + CStr(m2)
                Else
                    ExecuteSQL5 "UPDATE SoDuKhachHang SET DuNo_0=" + DoiDau(rs!ndk) + ",DuCo_0=" + DoiDau(rs!cdk) + ",DuNT_0=" + DoiDau(rs!ntdk) + " WHERE MaTaiKhoan=" + CStr(mk) + " AND MaKhachHang=" + CStr(m2)
                End If
            Else
                If thangdb = 0 Then
                    ExecuteSQL5 "INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0,n" + s + ",c" + s + ",nt" + s + ") VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(mk) + "," + CStr(m2) + "," + DoiDau(rs!ndk) + "," + DoiDau(rs!cdk) + "," + DoiDau(rs!ntdk) + "," + DoiDau(rs!ndk) + "," + DoiDau(rs!cdk) + "," + DoiDau(rs!ntdk) + ")"
                Else
                    ExecuteSQL5 "INSERT INTO SoDuKhachHang (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0) VALUES (" + CStr(Lng_MaxValue("MaSo", "SoDuKhachHang") + 1) + "," + CStr(mk) + "," + CStr(m2) + "," + DoiDau(rs!ndk) + "," + DoiDau(rs!cdk) + "," + DoiDau(rs!ntdk) + ")"
                End If
            End If
            rs2.Close
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set rs2 = Nothing
End Sub

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

Public Sub XuLyChungtu()
    Dim rs_chungtu As Recordset
    Dim chungtu As New ClsChungtu
    
    Set rs_chungtu = DBKetoan.OpenRecordset("SELECT MaSo FROM ChungTu WHERE XuLy=1 ORDER BY MaCT, MaSo", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_chungtu.EOF
        chungtu.InitChungtu rs_chungtu!MaSo, 0, "", 0, n1, n2, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
        chungtu.XuLyCT
        rs_chungtu.MoveNext
    Loop
    rs_chungtu.Close
    
    Set rs_chungtu = Nothing
    Set chungtu = Nothing
End Sub


'Ma hoa
Function Federo16(ByVal src As String, ByVal FStr As String) As String
    Dim result As String
    Dim temp$, h$, i%, j%
    j = 1
    For i = 1 To Len(src)
        h = Hex$(Asc(Mid$(src, i, 1)))
        If Len(h) = 1 Then h = "0" & h
        temp = temp & h
    Next i
    For i = 1 To Len(temp)
        result = result & Chr(Asc(Mid$(temp, i, 1)) + Asc(Mid$(FStr, j, 1)))
        j = IIf((j = Len(FStr)), 1, j + 1)
    Next i
    Federo16 = result
End Function
 
'mh
Function Federo16Decrypt(ByVal src As String, ByVal FStr As String) As String
    Dim result As String
    Dim temp$, i%, j%
    j = 1
    For i = 1 To Len(src)
        temp = temp & Chr(Asc(Mid$(src, i, 1)) - Asc(Mid$(FStr, j, 1)))
        j = IIf((j = Len(FStr)), 1, j + 1)
    Next i
    For i = 1 To Len(temp) Step 2
        result = result & Chr(CLng("&H" & Mid$(temp, i, 2)))
    Next i
    Federo16Decrypt = result
End Function
