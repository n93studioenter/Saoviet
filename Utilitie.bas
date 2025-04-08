Attribute VB_Name = "modUtilities"
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
' Maintenance string for PSS usage
End Type
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Const ERROR_SUCCESS = 0&

Private Const MAX_PATH = 260
Private Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_FONTCHANGE = &H1D
Private Declare Function GetVolumeSerialNumber Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As Long, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, ByVal lpMaximumComponentLength As Long, ByVal lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As Long, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFilename As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const REG_SZ = 1 ' Unicode null terminated string
Private Const VER_PLATFORM_WIN32_NT = 2
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetDiskFreeSpace Lib "Kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const er_SoHieu = 1
Public Const er_PhanLoai = 2
Public Const er_KhoHang = 3
Public Const er_NguonNX = 4
Public Const er_SHKhachHang = 5
Public Const er_SHTaiKhoan = 6
Public Const er_SHTaiKhoan1 = 7
Public Const er_SHVattu = 8
Public Const er_SHTaiSan = 9
Public Const er_SHThanhPham = 10
Public Const er_SHTKVT = 11
Public Const er_SHTKCN = 12
Public Const er_SHChTu = 13
Public Const er_Ten = 14

Public Const er_CoPS = 101
Public Const er_CoPS1 = 102
Public Const er_KoPS = 10003
Public Const er_KoPS1 = 10004

Public Const er_KoSD = 201

Public Const er_KoVT = 301
Public Const er_KoTS = 302

Public Const er_KoXem = 10001
Public Const er_RWait = 10002

Public Const er_NhieuCT = 11001

Public Const er_VTKoTon = 401

Public Const er_DBFile = 501
Public Const er_Connection = 502

Public Const er_Version = 901

Public Const SPI_GETNONCLIENTMETRICS = 41
Public Const SPI_SETNONCLIENTMETRICS = 42
Public Const SPI_GETICONTITLELOGFONT = 31
Public Const SPI_SETICONTITLELOGFONT = 34
Public Const sFONTNAME = "VK Sans Serif"

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type


Public Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSmCaptionWidth As Long
    iSmCaptionHeight As Long
    lfSmCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uActicon As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Boolean
Const LOCALE_SSHORTDATE As Long = &H1F
Private Declare Function GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function LoadLibraryRegister Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibfName As String) As Long
Private Declare Function GetProcAddressRegister Lib "Kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThreadForRegister Lib "Kernel32" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetExitCodeThread Lib "Kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function FreeLibraryRegister Lib "Kernel32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Sub ExitThread Lib "Kernel32" (ByVal xc As Long)

Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const gw_hwndnext = 2
Private Const fwp_startswith = 0
Private Const fwp_contains = 1

Public blnMDSettingsChanged As Boolean

Private Declare Function GetUserDefaultLCID Lib "Kernel32" () As Long
Private Declare Function GetSystemDefaultLCID Lib "Kernel32" () As Long
Private Declare Function GetThreadLocale Lib "Kernel32" () As Long

'Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "Kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long

Private Const WM_SETTINGCHANGE As Long = &H1A

'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private old_LOCALE_SSHORTDATE As String

Dim LCID As Long, iRet As Long, lpLCDataVar As String, Symbol As String
Dim iRet2 As Long, pos As Integer

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Const READ_CONTROL = &H20000
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const HKEY_CURRENT_USER = &H80000001
Public m_IgnoreEvents As Boolean

Public Sub ErrMsg(errnum As Integer)
    Dim msg As String
    
    Select Case errnum
        Case er_SoHieu:         msg = "Sè hiÖu kh«ng hîp lÖ !"
        Case er_Ten:         msg = "ThiÕu diÔn gi¶i hoÆc kh«ng hîp lÖ !"
        Case er_PhanLoai:     msg = "ThiÕu ph©n lo¹i !"
        Case er_KhoHang:        msg = "ThiÕu kho hµng !"
        Case er_NguonNX:        msg = "ThiÕu nguån nhËp xuÊt !"
        Case er_SHKhachHang:        msg = "ThiÕu sè hiÖu kh¸ch hµng !"
        Case er_SHTaiKhoan:        msg = "ThiÕu sè hiÖu tµi kho¶n!"
        Case er_SHTaiKhoan1:        msg = "ThiÕu sè hiÖu tµi kho¶n hoÆc tµi kho¶n cã chi tiÕt !"
        Case er_SHVattu:        msg = "ThiÕu sè hiÖu vËt t­ !"
        Case er_SHTaiSan:        msg = "ThiÕu sè hiÖu tµi s¶n !"
        Case er_SHThanhPham:        msg = "ThiÕu sè hiÖu thµnh phÈm !"
        Case er_SHTKVT:            msg = "ThiÕu sè hiÖu tµi kho¶n theo dâi chi tiÕt vËt t­ !"
        Case er_SHTKCN:            msg = "ThiÕu sè hiÖu tµi kho¶n theo dâi chi tiÕt c«ng nî !"
        Case er_SHChTu:             msg = "Sè hiÖu chøng tõ ®· cã !"
        
        Case er_CoPS:               msg = "§èi t­îng ®· cã ph¸t sinh !"
        Case er_CoPS1:               msg = "§èi t­îng ®· cã ph¸t sinh hoÆc sè d­ ®Çu kú !"
        Case er_KoPS:               msg = "Kh«ng cã ph¸t sinh !"
        Case er_KoPS1:               msg = "Kh«ng cã ph¸t sinh hoÆc sè d­ ®Çu kú !"
        
        Case er_KoSD:               msg = "Kh«ng cã quyÒn sö dông chøc n¨ng nµy !"
        
        Case er_KoVT:               msg = "Kh«ng khai b¸o theo dâi chi tiÕt vËt t­ !"
        Case er_KoTS:               msg = "Kh«ng khai b¸o theo dâi chi tiÕt tµi s¶n !"
        
        Case er_KoXem:               msg = "ChØ in ra m¸y in, kh«ng xem tr­íc !"
        Case er_RWait:             msg = "Xin chê m¸y tÝnh kh¸c trong m¹ng in xong b¸o c¸o nµy !"
        
        Case er_DBFile:                 msg = "TÖp d÷ liÖu kh«ng hîp lÖ !"
        Case er_Connection:     msg = "KÕt nèi v­ît sè m¸y cho phÐp !"
        
        Case er_VTKoTon:        msg = "Kh«ng cã tån kho !"
        Case er_NhieuCT:         msg = "Kh«ng hiÓn thÞ qu¸ " + CStr(MaxGridRow) + " chøng tõ, h·y läc chøng tõ theo th¸ng hoÆc lo¹i chøng tõ!"
        
        Case er_Version:        msg = "Kh«ng hç trî trong phiªn b¶n nµy, liªn hÖ Ban gi¶i ph¸p cho khèi doanh nghiÖp Nhµ n­íc, Cæ phÇn, Liªn doanh"
    End Select
    If pKhongDau = 1 Then msg = ABCtoKDau(msg)
    MsgBox msg, IIf(errnum < 1000, vbCritical, vbInformation), App.ProductName
End Sub
'======================================================================================
' Ham liet ke item tu Recordset vao Combo hoac List co kem Ma so
'======================================================================================
Public Function Int_RecsetToCbo(pstr_sql As String, Cbo As Object, Optional id As Integer = 0) As Integer
    Dim recset As Recordset
    Set recset = DBKetoan.OpenRecordset(pstr_sql, dbOpenSnapshot)
    Cbo.Clear
    If recset.RecordCount > 0 Then
        Do While Not recset.EOF
            If Not IsNull(recset!f1) Then
                Cbo.AddItem recset!f1
                Cbo.ItemData(Cbo.NewIndex) = recset!F2
            End If
            recset.MoveNext
        Loop
        If Cbo.ListCount > 0 And id >= 0 Then Cbo.ListIndex = id
    End If
    recset.Close
    Set recset = Nothing
    Int_RecsetToCbo = Cbo.ListCount
End Function
'======================================================================================
' Function Int_StrToCode : H¡m tr¢ vÌ m£ sä cía måt chuãi
'======================================================================================
Public Function Int_StrToCode(str As String) As Long
    Dim i As Long, Length As Integer, k As Long, kq As Long
    
    Length = Len(str)
    If Length > 0 Then
        If Length > 12 Then
            For i = Length To 1 Step -1
                k = Asc(Right(str, i))
                kq = kq + 2 * i * (k ^ 2)
            Next
        Else
            For i = Length To 1 Step -1
                k = Asc(Right(str, i))
                kq = kq + 8 * i * (k ^ 3)
            Next
        End If
    End If
    Int_StrToCode = kq
End Function

Public Function Int_StrToCodes(str As String) As Long
    Dim i As Long, Length As Integer, k As Long, kq As Long
    
    Length = Len(str)
    If Length > 0 Then
        If Length > 12 Then
            For i = Length To 1 Step -1
                k = Asc(Right(str, i))
                If (k Mod 2 = 0) Then
                    kq = kq + 4 * i * (k ^ 2)
                Else
                    kq = kq + 5 * i * (k ^ 2)
                End If
            Next
        Else
            For i = Length To 1 Step -1
                k = Asc(Right(str, i))
                If (k Mod 2 = 0) Then
                    kq = kq + 8 * i * (k ^ 3)
                Else
                    kq = kq + 9 * i * (k ^ 3)
                End If
            Next
        End If
    End If
    Int_StrToCodes = kq
End Function
'==========================================================
'kiem tra da a
Public Function boolean_kiemtra() As Boolean
    Dim rs As Recordset, i As Integer
    Dim KT As Boolean
    Dim st As String
    Dim so As Integer
    
    KT = False
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW License.* FROM License", dbOpenSnapshot)
    st = rs!CMP
   ' If (Int_StrToCodes(st) = rs!CMG) And (Int_StrToCodes(str(rs!nam)) = rs!namCode) Then KT = True
    If (Int_StrToCodes(st) = rs!CMG) Then
                    KT = True
             
    End If
    boolean_kiemtra = KT
End Function

'======================================================================================
' Thu tuc tra ve gia tri max cua mot truong so trong bang
'======================================================================================
Public Function Lng_MaxValue(pstr_field As String, pstr_table As String) As Long
    Lng_MaxValue = CLng(SelectSQL("SELECT Max(" + pstr_field + ") As F1 FROM " + pstr_table))
End Function
'======================================================================================
' Thu tuc xoa cac o nhap tren form
'======================================================================================
Public Sub ClearText(frm As Form)
    Dim i As Integer
    With frm
        For i = 0 To .Controls.count - 1
            If TypeOf .Controls(i) Is TextBox Then .Controls(i).Text = ""
            If (TypeOf .Controls(i) Is Label) And (.Controls(i).tag <> Null) Then .Controls(i).Caption = ""
        Next
    End With
End Sub
'======================================================================================
' Thu tuc dat thong tin license tren bao cao
'======================================================================================
Public Sub SetRptInfo()
    Dim i As Integer
        
    frmMain.Rpt.Formulas(0) = "TenCty='" + pTenCty + "'"
    If Len(Trim(pTenCn)) = 0 Or Left(pTenCn, 1) = "." Then
        frmMain.Rpt.Formulas(1) = "TenCn='MST: " + frmMain.LbCty(8).Caption + "'"
    Else
        frmMain.Rpt.Formulas(1) = "TenCn='" + pTenCn + " - MST: " + frmMain.LbCty(8).Caption + "'"
    End If
    frmMain.Rpt.Formulas(2) = "Nam=" + CStr(pNamTC)
    For i = 3 To 128
        frmMain.Rpt.Formulas(i) = ""
    Next
    frmMain.Rpt.DataFiles(0) = pDataPath
    frmMain.Rpt.connect = "DSN=;PWD=" + pPSW + ";UID=;DSQ="
    
    frmMain.Rpt.WindowShowPrintSetupBtn = True
End Sub
'======================================================================================
' SUB ColumnSetUp
'======================================================================================
Public Sub ColumnSetUp(Grid_control As Grid, col_index As Integer, col_Width As Integer, col_alignment As Integer)
      Grid_control.Row = 0
      Grid_control.col = col_index
      Grid_control.ColWidth(col_index) = col_Width
      Grid_control.FixedAlignment(col_index) = col_alignment
      If col_index >= Grid_control.FixedCols Then Grid_control.ColAlignment(col_index) = col_alignment
End Sub
'======================================================================================
' SUB ClearGrid
'======================================================================================
Public Sub ClearGrid(Grid_control As Grid, visible_rows As Integer)
      Dim i As Integer, str_row As String
      
      str_row = ""
      
      For i = 2 To Grid_control.Cols
            str_row = Chr(9) + str_row
      Next
      
      For i = 1 To visible_rows
            Grid_control.AddItem str_row, 0
      Next
      
      Do Until Grid_control.Rows = visible_rows
            Grid_control.RemoveItem Grid_control.Rows - 1
      Loop
      Grid_control.Row = 0
End Sub
'======================================================================================
' SUB SetListIndex
'======================================================================================
Public Sub SetListIndex(combo_box As Object, ma_so As Long)
Dim n As Integer
      If combo_box.ListCount = 0 Then Exit Sub
      Do Until (n = (combo_box.ListCount - 1)) Or (combo_box.ItemData(n) = ma_so)
            n = n + 1
      Loop
      combo_box.ListIndex = n
End Sub
'======================================================================================
' FUNCTION KeyProcess
'======================================================================================
Public Function KeyProcess(o As TextBox, key As Integer, Optional signenable = False) As Integer
    Dim X As Double
    
      If key = 46 Or key = 44 Then
            If sDecimal = "," Then key = 44
            If sDecimal = "." Then key = 46
      Else
            If key < 48 Or key > 57 Then                                            ' Kh«ng ph¶i c¸c phÝm sè
                  If Not (key = 8 Or (signenable And key = 45)) Then     ' DÊu "." , ",", "-"  vµ BaskSpace
                    
                        Beep
                        key = 0                                                                        ' Huû phÝm bÊm
                        
                        X = FrmCal.Calc
                        If X <> 0 Then o.Text = Format(X, Mask_2)
                  End If
            End If
      End If
      KeyProcess = key
End Function
'======================================================================================
' SUB AutoSelect : Tù ®éng ®¸nh dÊu ®o¹n Text trªn c¸c TextBox khi nhËn Focus
'======================================================================================
Public Sub AutoSelect(text_box As Object)
      text_box.SelStart = 0
      text_box.SelLength = Len(text_box.Text)
End Sub
'======================================================================================
' Hµm ®æi chuçi sè cã 2 sè thËp ph©n thµnh chuçi cã dÊu thËp ph©n lµ '.'
'======================================================================================
Public Function DoiDau(so As Double)
    Dim Length As Integer, pos As Integer, st As String
    so = Fix(IIf(so >= 0, 0.5, -0.5) + so * 100) / 100
    st = CStr(so)
    pos = IIf(sDecimal = ",", InStr(st, "."), InStr(st, ","))
    Do While pos > 0
        st = Left(st, pos - 1) + Right(st, Len(st) - pos)
        pos = IIf(sDecimal = ",", InStr(st, "."), InStr(st, ","))
    Loop
    pos = InStr(st, sDecimal)
    If pos > 0 Then
        Length = Len(st)
        DoiDau = Left(st, pos - 1) + "." + Right(st, Length - pos)
    Else
        DoiDau = st
    End If
End Function

'======================================================================================
' SUB SetGridIndex : §Æt v¹ch ®¸nh dÊu dßng hiÖn thêi trªn Grid.
'                    Tham sè : Grid Control, chØ sè dßng hiÖn thêi, sè dßng cã thÓ thÊy trªn l­íi vµ sè cét.
'                           Chó ý : Grid sÏ bÞ cuén lªn hoÆc xuèng nÕu dßng hiÖn thêi kh«ng n»m trong vïng nh×n
'                                         thÊy ®­îc, v× vËy thñ tôc nµy ph¶i lu«n ®­îc gäi víi mçi phÝm Ên trªn Grid.
'                                         Chøc n¨ng ®¸nh dÊu nhiÒu « trªn Grid bÞ sÏ bá qua nÕu sö dông thñ tôc nµy.
'======================================================================================
Public Sub SetGridIndex(Grid_control As Grid, cur_row As Integer)
With Grid_control
      ' KiÓm tra dßng hîp lÖ
      Select Case cur_row
            Case Is < 0: cur_row = 0
            Case Is > (.Rows - 1): cur_row = .Rows - 1
      End Select
      ' §¸nh dÊu dßng
      .col = 0
      .Row = cur_row

      .SelStartRow = cur_row
      .SelEndRow = cur_row
      .SelStartCol = .FixedCols
      .SelEndCol = .Cols - 1
      ' §iÒu chØnh vïng nh×n thÊy trªn l­íi nÕu v­ît qu¸
      Select Case cur_row
            Case Is < .TopRow: .TopRow = cur_row
           Case Is > (.TopRow + .Rows - 1): .TopRow = (cur_row + 1) - .Rows
     End Select
End With
End Sub

'======================================================================================
' Ham tr¶ vÒ sè ngµy trong mét th¸ng
'======================================================================================
Public Function SoNgayTrongThang(nam As Integer, thang As Integer) As Integer
    Select Case thang
        Case 1, 3, 5, 7, 8, 10, 12:     SoNgayTrongThang = 31
        Case 4, 6, 9, 11:                    SoNgayTrongThang = 30
        Case Else
            SoNgayTrongThang = IIf(nam Mod 4 = 0, 29, 28)
    End Select
End Function
'======================================================================================
' Ham tra ve ngay dau thang
'======================================================================================
Public Function NgayDauThang(n As Integer, thang As Integer) As Date
    Dim nam As Integer
    
    nam = IIf(thang < pThangDauKy, n + 1, n)
    If NgayDauThangMoi > 0 And thang <> pThangDauKy Then
        NgayDauThang = CVDate(CStr(NgayDauThangMoi) + "/" + CStr(ThangTruoc(thang)) + "/" + CStr(nam))
    Else
        NgayDauThang = CVDate("1/" + CStr(thang) + "/" + CStr(nam))
    End If
End Function
'======================================================================================
' Ham tra ve ngay cuoi thang
'======================================================================================
Public Function NgayCuoiThang(n As Integer, thang As Integer) As Date
    Dim nam As Integer
    
    If thang = 0 Or n = 0 Then Exit Function
    nam = IIf(thang < pThangDauKy, n + 1, n)
    If NgayDauThangMoi = 0 Or thang = pThangDauKy - 1 Or (thang = 12 And pThangDauKy = 1) Then
        NgayCuoiThang = CVDate(CStr(SoNgayTrongThang(nam, thang)) + "/" + CStr(thang) + "/" + CStr(nam))
    Else
        NgayCuoiThang = CVDate(CStr(NgayDauThangMoi - 1) + "/" + CStr(thang) + "/" + CStr(nam))
    End If
End Function
'======================================================================================
' SUB InitGrid : Khëi t¹o sè dßng vµ sè cét cho Grid Control.
'       Tham sè : Grid Control, sè dßng (b»ng sè dßng cã thÓ thÊy trªn Grid), sè cét (tÝnh tõ 1)
'              Chó ý : Grid t­¬ng øng ®­îc quy ®Þnh kh«ng chøa dßng vµ cét Fixed
'======================================================================================
Public Sub InitGrid(Grid_control As Grid, row_num As Integer, col_num As Integer)
      Grid_control.Rows = row_num
      Grid_control.Cols = col_num
End Sub
'======================================================================================
' Hien thong bao tren thanh trang thai
'======================================================================================
Public Sub HienThongBao(thong_bao As String, tabid As Integer)
      frmMain.sbStatusBar.Panels(tabid).Text = thong_bao
      frmMain.sbStatusBar.Refresh
End Sub

Public Function InsertGridRow(Grd As Grid, col As Integer, ref As String) As Integer
Dim i As Integer
    
    With Grd
        .col = col
        For i = 0 To .Rows - 1
            .Row = i
            If Len(.Text) = 0 Then Exit For
            If ref < .Text Then Exit For
        Next
        InsertGridRow = i
    End With
End Function

Public Function ToVNText(so As Double) As String
    Dim nst As String, stlen As Integer, i As Integer, j As Integer, pre As Integer, suf As Integer, suf1 As Integer, dau As String
        
    If so < 0 Then
        dau = "¢m "
        so = -so
    End If
    nst = CStr(Fix(so))
    stlen = Len(nst)
'    If stlen > 10 Then Exit Function
    For i = 1 To stlen
        j = CInt5(Mid(nst, i, 1))
        If i > 1 Then
            pre = CInt5(Mid(nst, i - 1, 1))
        Else
            pre = 0
        End If
        If i < stlen Then
            suf = CInt5(Mid(nst, i + 1, 1))
        Else
            suf = 0
        End If
        If i < stlen - 1 Then
            suf1 = CInt5(Mid(nst, i + 2, 1))
        Else
            suf1 = 0
        End If
        ToVNText = ToVNText + DonVi(pre, j, suf, suf1, stlen - i + 1)
    Next
    ToVNText = LTrim(ToVNText)
    If Len(ToVNText) > 0 Then
        ToVNText = dau + UCase(Left(ToVNText, 1)) + Right(ToVNText, Len(ToVNText) - 1)
    End If
End Function


Private Function DonVi(pre As Integer, i As Integer, suf As Integer, suf1 As Integer, pos As Integer) As String
    Select Case i
        Case 0: 'If ((pos - 1) Mod 3 = 0) And pre > 0 Then DonVi = " m­¬i" + IIf(suf = 0, Vitri(pos, suf, suf1,i), 0)
                        If ((pos - 2) Mod 3 = 0) And pre > 0 And suf > 0 Then DonVi = " lÎ"
                        If pos = 10 Then DonVi = " tû"
        Case 1: If ((pos - 1) Mod 3 = 0) Then DonVi = IIf(pre > 1, " mèt", " mét") + Vitri(pos, pre, suf, suf1, i)
                        If ((pos - 2) Mod 3 = 0) Then DonVi = " m­êi" + IIf(suf = 0 And (i > 1 Or pos > 2), Vitri(pos, pre, suf, suf1, i), "")
                        If (pos Mod 3 = 0) Then DonVi = " mét" + Vitri(pos, pre, suf, suf1, i)
        Case 2: DonVi = " hai" + Vitri(pos, pre, suf, suf1, i)
        Case 3: DonVi = " ba" + Vitri(pos, pre, suf, suf1, i)
        Case 4: DonVi = " bèn" + Vitri(pos, pre, suf, suf1, i)
        Case 5:  DonVi = IIf((pos - 1) Mod 3 = 0 And pre > 0, " l¨m", " n¨m") + Vitri(pos, pre, suf, suf1, i)
        Case 6: DonVi = " s¸u" + Vitri(pos, pre, suf, suf1, i)
        Case 7: DonVi = " b¶y" + Vitri(pos, pre, suf, suf1, i)
        Case 8: DonVi = " t¸m" + Vitri(pos, pre, suf, suf1, i)
        Case 9: DonVi = " chÝn" + Vitri(pos, pre, suf, suf1, i)
    End Select
End Function

Private Function Vitri(i As Integer, pre As Integer, suf As Integer, suf1 As Integer, v As Integer) As String
    Dim k As Integer
    k = IIf(i < 11, (i - 1) Mod 10, i Mod 10)
    Select Case k
        Case 0: Vitri = ""
        Case 1: Vitri = " m­¬i"
        Case 2: Vitri = " tr¨m"
        Case 3: Vitri = " ngh×n"
        Case 4: Vitri = IIf(v > 1, " m­¬i", "") + IIf(suf = 0, " ngh×n", "")
        Case 5: Vitri = " tr¨m" + IIf(suf = 0 And suf1 = 0, " ngh×n", "")
        Case 6: Vitri = " triÖu"
        Case 7: Vitri = IIf(v > 1, " m­¬i", "") + IIf(suf = 0, " triÖu", "")
        Case 8: Vitri = " tr¨m" + IIf(suf = 0 And suf1 = 0, " triÖu", "")
        Case 9: Vitri = " tû"
    End Select
End Function

Public Function XLSCol(c As Integer) As String
    Dim i As Integer
            
    If c < 27 Then
        XLSCol = Chr(c + 64)
    Else
        i = Fix((c - 1) / 26)
        XLSCol = Chr(i + 64) + Chr((c - 1) Mod 26 + 65)
    End If
End Function

Public Function BangDaCo(T As String) As Boolean
    Dim i As Integer
    
    BangDaCo = False
    For i = 0 To DBKetoan.TableDefs.count - 1
        If UCase(DBKetoan.TableDefs(i).Name) = UCase(T) Then
            BangDaCo = True
            Exit For
        End If
    Next
End Function

Public Function TruongDaCo(T As String, f As String) As Boolean
    Dim i As Integer
    
    TruongDaCo = False
    If BangDaCo(T) Then
        For i = 0 To DBKetoan.TableDefs(T).Fields.count - 1
            If UCase(DBKetoan.TableDefs(T).Fields(i).Name) = UCase(f) Then
                TruongDaCo = True
                Exit For
            End If
        Next
    End If
End Function

Public Function GetNumber(st As String) As Long
    Dim i As Integer, s As String, L As Integer
    
    L = Len(st)
    For i = 1 To L
        If IsNumeric(Mid(st, i, 1)) Then
            s = s + Mid(st, i, 1)
        Else
            If Len(s) > 0 Then Exit For
        End If
    Next
    If Len(s) > 0 Then GetNumber = CLng5(s)
End Function

Public Sub RptSetDate(ngay As Date, Optional nn As Integer)
    If nn = 0 Then
        frmMain.Rpt.Formulas(71) = "Ngay='Ngµy " + CStr(Day(ngay)) + " th¸ng " + CStr(Month(ngay)) + " n¨m " + CStr(Year(ngay)) + "'"
    Else
        frmMain.Rpt.Formulas(71) = "Ngay='" + Format(ngay, "dddd, mmm dd yyyy") + "'"
    End If
End Sub

Public Sub RFocus(obj As Object)
    On Error Resume Next
    obj.SetFocus
    On Error GoTo 0
End Sub


Public Sub CboCopy(FCbo As ComboBox, TCbo As ComboBox)
    Dim i As Integer
    For i = 0 To FCbo.ListCount - 1
        TCbo.AddItem FCbo.List(i)
        TCbo.ItemData(TCbo.NewIndex) = FCbo.ItemData(i)
    Next
    TCbo.ListIndex = FCbo.ListIndex
End Sub

Public Function VolumeSerial(DriveLetter) As Long
    Dim Serial As Long
    Call GetVolumeSerialNumber(UCase(DriveLetter) & ":\", 0&, 0&, Serial, 0&, 0&, 0&, 0&)
    VolumeSerial = Serial
End Function

Public Sub CallExcel(f As String)
    Dim expath As String
    Dim d As String
    
    d = CurrentDrive
    If Len(Dir(d + "\Program Files\Microsoft Office\Office\EXCEL.EXE")) > 0 Then
        expath = d + "\Program Files\Microsoft Office\Office\EXCEL.EXE"
    Else
        If Len(Dir(d + "\Program Files\Microsoft Office\Office10\EXCEL.EXE")) > 0 Then
            expath = d + "\Program Files\Microsoft Office\Office10\EXCEL.EXE"
        Else
            If Len(Dir(d + "\Program Files\Microsoft Office\Office11\EXCEL.EXE")) > 0 Then
                expath = d + "\Program Files\Microsoft Office\Office11\EXCEL.EXE"
            Else
                expath = "EXCEL.EXE"
            End If
        End If
    End If
    On Error Resume Next
    Shell GetSetting(IniPath, "Environment", "ExcelPath", expath) + " " + pCurDir + f, vbMaximizedFocus
    On Error GoTo 0
End Sub

Public Function GetLastRow(q As String, f As String) As Variant
    Dim rs As Recordset
    
    Set rs = DBKetoan.OpenRecordset(q, dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        GetLastRow = rs.Fields(f)
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function RptOK(fname As String, nn As Integer) As Boolean
    Dim Fs As Long, fn As String
    
    If InStr(fname, "\") = 0 Then
        fn = fname
        fname = pCurDir + "REPORTS" + IIf(nn > 0, "E", "") + IIf(pTien > 0, "2", "") + "\" + fname
    Else
        fn = GetFileName(fname)
    End If
    frmMain.Rpt.ReportFileName = fname
    
    Fs = SelectSQL("SELECT FileSize AS F1 FROM Reports WHERE FileName='" + fn + IIf(nn > 0, "E", "") + "'")
    If Fs > 0 Then
        RptOK = (FileLen(fname) = Fs)
    Else
        RptOK = True
    End If
End Function

Public Function GetFileName(fname As String) As String
    Dim i As Integer, st As String
    
    st = fname
    i = InStr(st, "\")
    Do While i > 0
        st = Right(st, Len(st) - i)
        i = InStr(st, "\")
    Loop
    GetFileName = st
End Function

Public Sub SetFont(frm As Form, Optional c As Integer = 0)
    Dim i As Integer, s As String, j As Integer
        
    On Error Resume Next
    With frm
        If pNN = 1 Or c = 1 Then
            If Len(frm.tag) > 0 And (Not IsNumeric(frm.tag)) Then
                s = frm.Caption
                frm.Caption = frm.tag
                frm.tag = s
            Else
                If Len(frm.LinkTopic) > 0 Then
                    frm.Caption = frm.LinkTopic
                End If
            End If
        End If
        For i = 0 To .Controls.count - 1
            If ((TypeOf .Controls(i) Is Grid Or TypeOf .Controls(i) Is Outline) And FontFlag > 0) Or TypeOf .Controls(i) Is TextBox Or TypeOf .Controls(i) Is ComboBox Or TypeOf .Controls(i) Is ListBox Then
                .Controls(i).FontName = pFontName
                .Controls(i).FontSize = pFontSize
            End If
            If IsNumeric(.Controls(i).tag) Then
                If TypeOf .Controls(i) Is Label And .Controls(i).tag = 1 Then
                    .Controls(i).FontName = pFontName
                    .Controls(i).FontSize = pFontSize
                End If
            End If
            If pNN = 1 Or c = 1 Then
                If (TypeOf .Controls(i) Is OptionButton Or TypeOf .Controls(i) Is Label Or TypeOf .Controls(i) Is CheckBox Or TypeOf .Controls(i) Is Menu Or TypeOf .Controls(i) Is Frame) Then
                    If Len(.Controls(i).tag) > 0 And (Not IsNumeric(.Controls(i).tag)) Then
                        s = .Controls(i).tag
                        .Controls(i).tag = .Controls(i).Caption
                        .Controls(i).Caption = s
                    End If
                    If Not TypeOf .Controls(i) Is Menu Then
                        If Len(.Controls(i).ToolTipText) > 0 And (Not IsNumeric(.Controls(i).ToolTipText)) And (.Controls(i).ForeColor <> &HFF0000) Then
                            s = .Controls(i).ToolTipText
                            .Controls(i).ToolTipText = .Controls(i).Caption
                            .Controls(i).Caption = s
                        End If
                    End If
                End If
                If TypeOf .Controls(i) Is CommandButton And Len(.Controls(i).tag) > 0 Then
                    s = pCurDir + frm.Name + "_" + .Controls(i).Name + "_" + CStr(.Controls(i).Index) + ".BMP"
                    If .Controls(i).Picture <> 0 Then
                        If UCase(Left(frm.Name, 3)) = "FBC" Or UCase(Right(frm.Name, 4)) = "MAIN" Then SavePicture .Controls(i).Picture, s
                        Set .Controls(i).Picture = LoadPicture()
                        .Controls(i).Caption = .Controls(i).tag
                    Else
                        If Len(Dir(s)) > 0 Then
                            Set .Controls(i).Picture = LoadPicture(s)
                            .Controls(i).Caption = ""
                        Else
                            .Controls(i).Caption = .Controls(i).tag
                        End If
                    End If
                End If
                If TypeOf .Controls(i) Is SSTab Then
                    If Len(.Controls(i).tag) > 0 Then
                        s = ""
                        For j = 0 To .Controls(i).Tabs - 1
                            s = s + .Controls(i).TabCaption(j) + "#"
                            .Controls(i).TabCaption(j) = LaySH(.Controls(i).tag, j + 1)
                        Next
                        .Controls(i).tag = Left(s, Len(s) - 1)
                    End If
                End If
                If TypeOf .Controls(i) Is Toolbar Then
                    For j = 1 To .Controls(i).Buttons.count
                        If Len(.Controls(i).Buttons(j).tag) > 0 Then
                            s = .Controls(i).Buttons(j).tag
                            .Controls(i).Buttons(j).tag = .Controls(i).Buttons(j).ToolTipText
                            .Controls(i).Buttons(j).ToolTipText = s
                        End If
                    Next
                End If
            End If
            If pKhongDau = 1 Then
                If TypeOf .Controls(i) Is Menu Then .Controls(i).Caption = ABCtoKDau(.Controls(i).Caption)
                If TypeOf .Controls(i) Is Label And Len(.Controls(i).ToolTipText) > 0 Then .Controls(i).ToolTipText = ABCtoKDau(.Controls(i).ToolTipText)
            End If
        Next
        If pKhongDau = 1 Then .Caption = ABCtoKDau(.Caption)
    End With
    On Error GoTo 0
End Sub

Public Function CurrentDrive() As String
    Dim retValue As Long, Buffer As String * 255
    
    retValue = GetWindowsDirectory(Buffer, 255)
    CurrentDrive = Left(Buffer, 2)
End Function

Public Function ABCtoVNI(st As String) As String
    Dim i As Integer, L As Integer, c As Integer, C1 As Integer, c2 As Integer
    
    If FontFlag <> 2 Then
        ABCtoVNI = st
        Exit Function
    End If
    
    L = Len(st)
    For i = 1 To L
        c = Asc(Mid(st, i, 1))
        c2 = 0
        Select Case c
            Case 181:   C1 = 97
                                   c2 = 248
            Case 184:   C1 = 97
                                   c2 = 249
            Case 182:   C1 = 97
                                   c2 = 251
            Case 183:   C1 = 97
                                   c2 = 245
            Case 185:   C1 = 97
                                   c2 = 239
            Case 169:   C1 = 97
                                   c2 = 226
            Case 199:   C1 = 97
                                   c2 = 224
            Case 202:   C1 = 97
                                   c2 = 225
            Case 200:   C1 = 97
                                   c2 = 229
            Case 201:   C1 = 97
                                   c2 = 227
            Case 203:   C1 = 97
                                   c2 = 228
            Case 174:   C1 = 241
            Case 204:   C1 = 101
                                   c2 = 248
            Case 208:   C1 = 101
                                   c2 = 249
            Case 206:   C1 = 101
                                   c2 = 251
            Case 207:   C1 = 101
                                   c2 = 245
            Case 209:   C1 = 101
                                   c2 = 239
            Case 170:   C1 = 101
                                   c2 = 226
            Case 210:   C1 = 101
                                   c2 = 224
            Case 213:   C1 = 101
                                   c2 = 225
            Case 211:   C1 = 101
                                   c2 = 229
            Case 212:   C1 = 101
                                   c2 = 227
            Case 214:   C1 = 101
                                   c2 = 228
            Case 215:   C1 = 236
            Case 221:   C1 = 237
            Case 216:   C1 = 230
            Case 220:   C1 = 243
            Case 222:   C1 = 242
            Case 223:   C1 = 111
                                   c2 = 248
            Case 227:   C1 = 111
                                   c2 = 249
            Case 225:   C1 = 111
                                   c2 = 251
            Case 226:   C1 = 111
                                   c2 = 245
            Case 228:   C1 = 111
                                   c2 = 239
            Case 171:   C1 = 111
                                   c2 = 226
            Case 229:   C1 = 111
                                   c2 = 224
            Case 232:   C1 = 111
                                   c2 = 225
            Case 230:   C1 = 111
                                   c2 = 229
            Case 231:   C1 = 111
                                   c2 = 227
            Case 233:   C1 = 111
                                   c2 = 228
            Case 172:   C1 = 244
            Case 234:   C1 = 244
                                   c2 = 248
            Case 237:   C1 = 244
                                   c2 = 249
            Case 235:   C1 = 244
                                   c2 = 251
            Case 236:   C1 = 244
                                   c2 = 245
            Case 238:   C1 = 244
                                   c2 = 239
            Case 239:   C1 = 117
                                   c2 = 248
            Case 243:   C1 = 117
                                   c2 = 249
            Case 241:   C1 = 117
                                   c2 = 251
            Case 242:   C1 = 117
                                   c2 = 245
            Case 244:   C1 = 117
                                   c2 = 239
            Case 173:   C1 = 249
                                    c2 = 246
            Case 245:   C1 = 246
                                   c2 = 248
            Case 248:   C1 = 246
                                   c2 = 249
            Case 246:   C1 = 246
                                   c2 = 251
            Case 247:   C1 = 246
                                   c2 = 245
            Case 249:   C1 = 246
                                   c2 = 239
            Case 250:   C1 = 121
                                   c2 = 248
            Case 253:   C1 = 121
                                   c2 = 249
            Case 251:   C1 = 121
                                   c2 = 251
            Case 252:   C1 = 121
                                   c2 = 245
            Case 254:   C1 = 238
            Case 168:   C1 = 97
                                   c2 = 234
            Case 187:   C1 = 97
                                   c2 = 232
            Case 190:   C1 = 97
                                   c2 = 233
            Case 188:   C1 = 97
                                   c2 = 250
            Case 189:   C1 = 97
                                   c2 = 252
            Case 198:   C1 = 97
                                   c2 = 235
            Case 162:   C1 = 65
                                   c2 = 194
            Case 161:   C1 = 65
                                    c2 = 202
            Case 167:   C1 = 209
            Case 163:   C1 = 69
                                    c2 = 194
            Case 164:   C1 = 79
                                    c2 = 194
            Case 165:   C1 = 212
            Case 166:   C1 = 214
            Case Else
                                    C1 = c
        End Select
        ABCtoVNI = ABCtoVNI + Chr(C1) + IIf(c2 > 0, Chr(c2), "")
    Next
End Function

Public Function ABCtoKDau(st As String) As String
    Dim i As Integer, L As Integer, c As Integer, C1 As Integer
        
    L = Len(st)
    For i = 1 To L
        c = Asc(Mid(st, i, 1))
        Select Case c
            Case 181, 184, 182, 183, 185, 169, 199, 202, 200, 201, 203: C1 = 97
            Case 204, 208, 206, 207, 209, 170, 210, 213, 211, 212, 214: C1 = 101
            Case 215, 221, 216, 221, 222:  C1 = 105
            Case 223, 227, 225, 226, 228, 171, 229, 232, 230, 231, 233, 172, 234, 235, 236, 237, 238: C1 = 111
            Case 239, 243, 241, 242, 244, 173, 245, 246, 247, 248, 249: C1 = 117
            Case 174:   C1 = 100
            Case 250, 253, 251, 252: C1 = 121
            Case 254:   C1 = 238
            Case 168, 187, 190, 188, 189, 198: C1 = 97
            Case 162, 161:  C1 = 65
            Case 167:   C1 = 68
            Case 163:   C1 = 69
            Case 164, 165:  C1 = 79
            Case 166:   C1 = 85
            Case Else
                C1 = c
        End Select
        ABCtoKDau = ABCtoKDau + Chr(C1)
    Next
End Function

Public Function FontDaCo(st As String) As Boolean
    Dim i As Integer
    FontDaCo = False
    For i = 0 To Screen.FontCount - 1
        If UCase(Screen.Fonts(i)) = UCase(st) Then
            FontDaCo = True
            Exit For
        End If
    Next
End Function

Public Function VNItoABC(st As String, Optional ktra As Integer = 0) As String
    Dim i As Integer, L As Integer, C1 As Integer, c2 As Integer, c As Integer
    
    If FontFlag = 2 And ktra = 0 Then
        VNItoABC = st
        Exit Function
    End If
    
    L = Len(st)
    i = 1
        Do While i <= L
        C1 = CInt5(Asc(Mid(st, i, 1)))
        c = 0
        If i = L Then GoTo a
        c2 = CInt5(Asc(Mid(st, i + 1, 1)))
        If C1 = 97 Then
            Select Case c2
                Case 248:   c = 181
                Case 249:   c = 184
                Case 251:   c = 182
                Case 245:   c = 183
                Case 239:   c = 185
                Case 226:   c = 169
                Case 224:   c = 199
                Case 225:   c = 202
                Case 229:   c = 200
                Case 227:   c = 201
                Case 228:   c = 203
                Case 234:   c = 168
                Case 232:   c = 187
                Case 233:   c = 190
                Case 250:   c = 188
                Case 252:   c = 189
                Case 235:   c = 198
            End Select
        End If
        If C1 = 65 Then
            Select Case c2
                Case 216:   c = 181
                Case 217:   c = 184
                Case 219:   c = 182
                Case 213:   c = 183
                Case 207:   c = 185
                Case 194:   c = 162
                Case 192:   c = 199
                Case 193:   c = 202
                Case 197:   c = 200
                Case 195:   c = 201
                Case 196:   c = 203
                Case 202:   c = 161
                Case 200:   c = 187
                Case 201:   c = 190
                Case 218:   c = 188
                Case 220:   c = 189
                Case 203:   c = 198
                Case 202:   c = 161
            End Select
        End If
        If C1 = 69 Then
            Select Case c2
                Case 216:   c = 204
                Case 217:   c = 208
                Case 219:   c = 206
                Case 213:   c = 207
                Case 207:   c = 209
                Case 194:   c = 163
                Case 192:   c = 210
                Case 193:   c = 213
                Case 197:   c = 211
                Case 195:   c = 212
                Case 196:   c = 214
            End Select
        End If
        If C1 = 79 Then
            Select Case c2
                Case 216:   c = 223
                Case 217:   c = 227
                Case 219:   c = 225
                Case 213:   c = 226
                Case 207:   c = 228
                Case 194:   c = 164
                Case 192:   c = 229
                Case 193:   c = 232
                Case 197:   c = 230
                Case 195:   c = 231
                Case 196:   c = 233
            End Select
        End If
        If C1 = 212 Then
            Select Case c2
                Case 216:   c = 234
                Case 217:   c = 237
                Case 219:   c = 235
                Case 213:   c = 236
                Case 207:   c = 238
            End Select
        End If
        If C1 = 85 Then
            Select Case c2
                Case 216:   c = 239
                Case 217:   c = 243
                Case 219:   c = 241
                Case 213:   c = 242
                Case 207:   c = 244
            End Select
        End If
        If C1 = 214 Then
            Select Case c2
                Case 216:   c = 245
                Case 217:   c = 248
                Case 219:   c = 246
                Case 213:   c = 247
                Case 207:   c = 249
            End Select
        End If
        If C1 = 89 Then
            Select Case c2
                Case 216:   c = 250
                Case 217:   c = 253
                Case 219:   c = 251
                Case 213:   c = 252
                Case 206:   c = 254
            End Select
        End If
        If C1 = 101 Then
            Select Case c2
                Case 248:   c = 204
                Case 249:   c = 208
                Case 251:   c = 206
                Case 245:   c = 207
                Case 239:   c = 209
                Case 226:   c = 170
                Case 224:   c = 210
                Case 225:   c = 213
                Case 229:   c = 211
                Case 227:   c = 212
                Case 228:   c = 214
            End Select
        End If
        If C1 = 111 Then
            Select Case c2
                Case 248:   c = 223
                Case 249:   c = 227
                Case 251:   c = 225
                Case 245:   c = 226
                Case 239:   c = 228
                Case 226:   c = 171
                Case 224:   c = 229
                Case 225:   c = 232
                Case 229:   c = 230
                Case 227:   c = 231
                Case 228:   c = 233
            End Select
        End If
        If C1 = 244 Then
            Select Case c2
                Case 248:   c = 234
                Case 249:   c = 237
                Case 251:   c = 235
                Case 245:   c = 236
                Case 239:   c = 238
            End Select
        End If
        If C1 = 117 Then
            Select Case c2
                Case 248:   c = 239
                Case 249:   c = 243
                Case 251:   c = 241
                Case 245:   c = 242
                Case 239:   c = 244
            End Select
        End If
        If C1 = 249 And c2 = 246 Then c = 173
        If C1 = 246 Then
            Select Case c2
                Case 248:   c = 245
                Case 249:   c = 248
                Case 251:   c = 246
                Case 245:   c = 247
                Case 239:   c = 249
            End Select
        End If
        If C1 = 121 Then
            Select Case c2
                Case 248:   c = 250
                Case 249:   c = 253
                Case 251:   c = 251
                Case 245:   c = 252
            End Select
        End If
        If c > 0 Then
            i = i + 2
            GoTo KT
        End If
a:
        Select Case C1
            Case 241:   c = 174
            Case 236, 204:  c = 215
            Case 237, 205:  c = 221
            Case 230, 198:  c = 216
            Case 243, 211:  c = 220
            Case 242, 210:  c = 222
            Case 244:   c = 172
            Case 246:   c = 173
            Case 238:   c = 254
            Case 209:   c = 167
            Case 212:   c = 165
            Case 214:   c = 166
        End Select
        If c > 0 Then
            i = i + 1
            GoTo KT
        End If
        If c = 0 Then
            i = i + 1
            c = C1
        End If
KT:
        VNItoABC = VNItoABC + Chr(c)
    Loop
End Function

Public Sub ChuyenDoiFont(ABC2VNI As Boolean)
    Dim i As Integer, j As Integer, rs As Recordset, st As String
    
    For i = 0 To DBKetoan.TableDefs.count - 1
        If Left(DBKetoan.TableDefs(i).Name, 4) <> "MSys" And UCase(DBKetoan.TableDefs(i).Name) <> "CDTS" And UCase(DBKetoan.TableDefs(i).Name) <> "KQKD" And UCase(DBKetoan.TableDefs(i).Name) <> "LCTT" And UCase(DBKetoan.TableDefs(i).Name) <> "THUE" And UCase(DBKetoan.TableDefs(i).Name) <> "VAT" Then
            For j = 0 To DBKetoan.TableDefs(i).Fields.count - 1
                If DBKetoan.TableDefs(i).Fields(j).Type = dbText Then
                    Set rs = DBKetoan.OpenRecordset(DBKetoan.TableDefs(i).Name, dbOpenTable, dbForwardOnly)
                    Do While Not rs.EOF
                        If Not IsNull(rs.Fields(DBKetoan.TableDefs(i).Fields(j).Name)) Then
                            If ABC2VNI Then
                                st = ABCtoVNI(rs.Fields(DBKetoan.TableDefs(i).Fields(j).Name))
                            Else
                                st = VNItoABC(rs.Fields(DBKetoan.TableDefs(i).Fields(j).Name))
                            End If
                            rs.Edit
                            rs.Fields(DBKetoan.TableDefs(i).Fields(j).Name).Value = Left(st, DBKetoan.TableDefs(i).Fields(j).Size)
                            rs.Update
                        End If
                        rs.MoveNext
                    Loop
                End If
            Next
        End If
    Next
    
    If ABC2VNI Then
        pTenCty = ABCtoVNI(pTenCty)
        pTenCn = ABCtoVNI(pTenCn)
    Else
        pTenCty = VNItoABC(pTenCty)
        pTenCn = VNItoABC(pTenCn)
    End If
    
    ExecuteSQL5 "UPDATE License SET TenCty_ID = " + CStr(Int_StrToCode(pTenCty)) + ",TenCn_ID = " + CStr(Int_StrToCode(pTenCn))
                
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
End Sub

Public Function ABCtoUNICODE(st As String) As String
    Dim i As Integer, L As Integer, c As Integer, C1 As Integer
    
    L = Len(st)
    For i = 1 To L
        c = Asc(Mid(st, i, 1))
        C1 = c
        Select Case c
            Case 181:   C1 = 224
            Case 184:   C1 = 225
            Case 182:   C1 = 7843
            Case 183:   C1 = 227
            Case 185:   C1 = 7841
            Case 162:   C1 = 194
            Case 169:   C1 = 226
            Case 199:   C1 = 7847
            Case 202:   C1 = 7845
            Case 200:   C1 = 7849
            Case 201:   C1 = 7851
            Case 203:   C1 = 7853
            Case 161:   C1 = 258
            Case 168:   C1 = 259
            Case 187:   C1 = 7857
            Case 190:   C1 = 7855
            Case 188:   C1 = 7859
            Case 189:   C1 = 7861
            Case 198:   C1 = 7863
            Case 167:   C1 = 272
            Case 174:   C1 = 273
            Case 204:   C1 = 232
            Case 208:   C1 = 233
            Case 206:   C1 = 7867
            Case 207:   C1 = 7869
            Case 209:   C1 = 7865
            Case 163:   C1 = 202
            Case 170:   C1 = 234
            Case 210:   C1 = 7873
            Case 213:   C1 = 7871
            Case 211:   C1 = 7875
            Case 212:   C1 = 7877
            Case 214:   C1 = 7879
            Case 215:   C1 = 236
            Case 221:   C1 = 237
            Case 216:   C1 = 7881
            Case 220:   C1 = 297
            Case 222:   C1 = 7883
            Case 223:   C1 = 242
            Case 227:   C1 = 243
            Case 225:   C1 = 7887
            Case 226:   C1 = 245
            Case 228:   C1 = 7885
            Case 164:   C1 = 212
            Case 171:   C1 = 244
            Case 229:   C1 = 7891
            Case 232:    C1 = 7889
            Case 230:   C1 = 7893
            Case 231:   C1 = 7895
            Case 233:   C1 = 7897
            Case 165:   C1 = 416
            Case 172:   C1 = 417
            Case 234:   C1 = 7901
            Case 237:   C1 = 7899
            Case 235:   C1 = 7903
            Case 236:   C1 = 7905
            Case 238:   C1 = 7907
            Case 239:   C1 = 249
            Case 243:   C1 = 250
            Case 241:   C1 = 7911
            Case 242:   C1 = 361
            Case 244:   C1 = 7909
            Case 166:   C1 = 431
            Case 173:   C1 = 432
            Case 245:   C1 = 7915
            Case 248:   C1 = 7913
            Case 246:   C1 = 7917
            Case 247:   C1 = 7919
            Case 249:   C1 = 7921
            Case 250:   C1 = 7923
            Case 251:   C1 = 7927
            Case 252:   C1 = 7929
            Case 254:   C1 = 7925
        End Select
        ABCtoUNICODE = ABCtoUNICODE + ChrW(C1)
    Next
End Function

Public Function UNICODEtoABC(st As String) As String
    Dim i As Integer, L As Integer, c As Integer, C1 As Integer
    
    L = Len(st)
    For i = 1 To L
        c = AscW(Mid(st, i, 1))
        C1 = c
        Select Case c
            Case 224, 192:  C1 = 181
            Case 225, 193:  C1 = 184
            Case 7843, 7842:  C1 = 182
            Case 227, 195:  C1 = 183
            Case 7841, 7840:  C1 = 185
            Case 194:   C1 = 162
            Case 226:   C1 = 169
            Case 7847, 7846: C1 = 199
            Case 7845, 7844:  C1 = 202
            Case 7849, 7848:  C1 = 200
            Case 7851, 7850:  C1 = 201
            Case 7853, 7852:  C1 = 203
            Case 258:   C1 = 161
            Case 259:   C1 = 168
            Case 7857, 7856:  C1 = 187
            Case 7855, 7854:  C1 = 190
            Case 7859, 7858:  C1 = 188
            Case 7861, 7860:  C1 = 189
            Case 7863, 7862:  C1 = 198
            Case 272:   C1 = 167
            Case 273:   C1 = 174
            Case 232, 200:  C1 = 204
            Case 233, 201:  C1 = 208
            Case 7867, 7866:  C1 = 206
            Case 7869, 7868:  C1 = 207
            Case 7865, 7864:  C1 = 209
            Case 202:   C1 = 163
            Case 234:   C1 = 170
            Case 7873, 7872:  C1 = 210
            Case 7871, 7870:  C1 = 213
            Case 7875, 7874:  C1 = 211
            Case 7877, 7876:  C1 = 212
            Case 7879, 7878:  C1 = 214
            Case 236, 204:  C1 = 215
            Case 237, 205:  C1 = 221
            Case 7881, 7880:  C1 = 216
            Case 297, 296:  C1 = 220
            Case 7883, 7882:  C1 = 222
            Case 242, 210:  C1 = 223
            Case 243, 211:  C1 = 227
            Case 7887, 7886:  C1 = 225
            Case 245, 213:  C1 = 226
            Case 7885, 7884:  C1 = 228
            Case 212:   C1 = 164
            Case 244:   C1 = 171
            Case 7891, 7890:  C1 = 229
            Case 7889, 7888:   C1 = 232
            Case 7893, 7892:  C1 = 230
            Case 7895, 7894:  C1 = 231
            Case 7897, 7896:  C1 = 233
            Case 416:   C1 = 165
            Case 417:   C1 = 172
            Case 7901, 7900:  C1 = 234
            Case 7899, 7898:  C1 = 237
            Case 7903, 7902:  C1 = 235
            Case 7905, 7904:  C1 = 236
            Case 7907, 7906:  C1 = 238
            Case 249, 217:  C1 = 239
            Case 250, 218:  C1 = 243
            Case 7911, 7910:  C1 = 241
            Case 361, 360:  C1 = 242
            Case 7909, 7908:  C1 = 244
            Case 431:   C1 = 166
            Case 432:   C1 = 173
            Case 7915, 7914:  C1 = 245
            Case 7913, 7912:  C1 = 248
            Case 7917, 7916:  C1 = 246
            Case 7919, 7918:  C1 = 247
            Case 7921, 7920:  C1 = 249
            Case 7923, 7922:  C1 = 250
            Case 221:   C1 = 253
            Case 7927, 7926:  C1 = 251
            Case 7929, 7928:  C1 = 252
            Case 7925, 7924:  C1 = 254
        End Select
        If C1 > 255 Then C1 = 255
        UNICODEtoABC = UNICODEtoABC + Chr(C1)
    Next
End Function

Public Function VNItoUNICODE(st As String) As String
    Dim i As Integer, L As Integer, C1 As Integer, c2 As Integer, c As Integer
        
    L = Len(st)
    i = 1
    Do While i <= L
        C1 = CInt5(Asc(Mid(st, i, 1)))
        c = 0
        If i = L Then GoTo a
        c2 = CInt5(Asc(Mid(st, i + 1, 1)))
        If C1 = 97 Then
            Select Case c2
                Case 248:   c = 224
                Case 249:   c = 225
                Case 251:   c = 7843
                Case 245:   c = 227
                Case 239:   c = 7841
                Case 226:   c = 226
                Case 224:   c = 7847
                Case 225:   c = 7845
                Case 229:   c = 7849
                Case 227:   c = 7851
                Case 228:   c = 7853
                Case 234:   c = 259
                Case 232:   c = 7857
                Case 233:   c = 7855
                Case 250:   c = 7859
                Case 252:   c = 7861
                Case 235:   c = 7863
            End Select
        End If
        If C1 = 65 Then
            Select Case c2
                Case 216:   c = 192
                Case 217:   c = 193
                Case 219:   c = 7842
                Case 213:   c = 195
                Case 207:   c = 7840
                Case 194:   c = 194
                Case 192:   c = 7846
                Case 193:   c = 7844
                Case 197:   c = 7848
                Case 195:   c = 7850
                Case 196:   c = 7852
                Case 200:   c = 7856
                Case 201:   c = 7854
                Case 218:   c = 7858
                Case 220:   c = 7860
                Case 203:   c = 7862
                Case 202:   c = 258
            End Select
        End If
        If C1 = 69 Then
            Select Case c2
                Case 216:   c = 200
                Case 217:   c = 201
                Case 219:   c = 7866
                Case 213:   c = 7868
                Case 207:   c = 7864
                Case 194:   c = 202
                Case 192:   c = 7872
                Case 193:   c = 7870
                Case 197:   c = 7874
                Case 195:   c = 7876
                Case 196:   c = 7878
            End Select
        End If
        If C1 = 79 Then
            Select Case c2
                Case 216:   c = 210
                Case 217:   c = 211
                Case 219:   c = 7886
                Case 213:   c = 213
                Case 207:   c = 7884
                Case 194:   c = 212
                Case 192:   c = 7890
                Case 193:   c = 7888
                Case 197:   c = 7892
                Case 195:   c = 7894
                Case 196:   c = 7896
            End Select
        End If
        If C1 = 212 Then
            Select Case c2
                Case 216:   c = 7900
                Case 217:   c = 7898
                Case 219:   c = 7902
                Case 213:   c = 7904
                Case 207:   c = 7906
            End Select
        End If
        If C1 = 85 Then
            Select Case c2
                Case 216:   c = 217
                Case 217:   c = 218
                Case 219:   c = 7910
                Case 213:   c = 360
                Case 207:   c = 7908
            End Select
        End If
        If C1 = 214 Then
            Select Case c2
                Case 216:   c = 7914
                Case 217:   c = 7912
                Case 219:   c = 7916
                Case 213:   c = 7918
                Case 207:   c = 7920
            End Select
        End If
        If C1 = 89 Then
            Select Case c2
                Case 216:   c = 7922
                Case 217:   c = 221
                Case 219:   c = 7926
                Case 213:   c = 7928
            End Select
        End If
        If C1 = 101 Then
            Select Case c2
                Case 248:   c = 232
                Case 249:   c = 233
                Case 251:   c = 7867
                Case 245:   c = 7869
                Case 239:   c = 7865
                Case 226:   c = 234
                Case 224:   c = 7873
                Case 225:   c = 7871
                Case 229:   c = 7875
                Case 227:   c = 7877
                Case 228:   c = 7879
            End Select
        End If
        If C1 = 111 Then
            Select Case c2
                Case 248:   c = 242
                Case 249:   c = 243
                Case 251:   c = 7887
                Case 245:   c = 245
                Case 239:   c = 7885
                Case 226:   c = 244
                Case 224:   c = 7891
                Case 225:   c = 7889
                Case 229:   c = 7893
                Case 227:   c = 7895
                Case 228:   c = 7897
            End Select
        End If
        If C1 = 244 Then
            Select Case c2
                Case 248:   c = 7901
                Case 249:   c = 7899
                Case 251:   c = 7903
                Case 245:   c = 7905
                Case 239:   c = 7907
            End Select
        End If
        If C1 = 117 Then
            Select Case c2
                Case 248:   c = 249
                Case 249:   c = 250
                Case 251:   c = 7911
                Case 245:   c = 361
                Case 239:   c = 7909
            End Select
        End If
        If C1 = 249 And c2 = 246 Then c = 432
        If C1 = 246 Then
            Select Case c2
                Case 248:   c = 7915
                Case 249:   c = 7913
                Case 251:   c = 7917
                Case 245:   c = 7919
                Case 239:   c = 7921
            End Select
        End If
        If C1 = 121 Then
            Select Case c2
                Case 248:   c = 7923
                Case 249:   c = 253
                Case 251:   c = 7927
                Case 245:   c = 7929
            End Select
        End If
        If c > 0 Then
            i = i + 2
            GoTo KT
        End If
a:
        Select Case C1
            Case 241:   c = 273
            Case 236:  c = 236
            Case 204:   c = 204
            Case 237:  c = 237
            Case 205:   c = 205
            Case 230:  c = 7881
            Case 198:    c = 7880
            Case 243:  c = 297
            Case 211:   c = 296
            Case 242:  c = 7883
            Case 210:   c = 7882
            Case 244:   c = 417
            Case 246:   c = 432
            Case 238:   c = 7925
            Case 209:   c = 272
            Case 212:   c = 416
            Case 214:   c = 431
            Case 206:   c = 7924
        End Select
        If c > 0 Then
            i = i + 1
            GoTo KT
        End If
        If c = 0 Then
            i = i + 1
            c = C1
        End If
KT:
        VNItoUNICODE = VNItoUNICODE + ChrW(c)
    Loop
End Function

Public Function ThemTruong(tbl As String, fld As String, tp As Integer, Optional s As Integer = 0, Optional dv As Integer = 0, Optional gt As String = "...") As Boolean
    Dim tdf As TableDef, sql As String
    
    ThemTruong = False
    If Not BangDaCo(tbl) Then GoTo KT
    If TruongDaCo(tbl, fld) Then
        Set tdf = DBKetoan.TableDefs(tbl)
        If tdf.Fields(fld).Type <> tp Then
            XoaTruong tbl, fld
        Else
            GoTo KT
        End If
    Else
        Set tdf = DBKetoan.TableDefs(tbl)
    End If
    tdf.Fields.Append tdf.CreateField(fld, tp, s)
    Select Case tp
        Case dbInteger, dbLong, dbDouble, dbSingle:
                    tdf.Fields(fld).DefaultValue = dv
                    sql = "UPDATE " + tbl + " SET " + fld + "=" + CStr(dv)
        Case dbText:
                    tdf.Fields(fld).DefaultValue = "..."
                    sql = "UPDATE " + tbl + " SET " + fld + "='" + gt + "'"
        Case dbDate
                    tdf.Fields(fld).DefaultValue = CVDate("1/1/80")
                    sql = "UPDATE " + tbl + " SET " + fld + "=#1/1/80#"
    End Select
    If Len(sql) > 0 Then ExecuteSQL5 sql
    ThemTruong = True
KT:
    Set tdf = Nothing
End Function

Public Sub XoaTruong(tbl As String, fld As String)
    Dim tdf As TableDef
    
    If TruongDaCo(tbl, fld) Then
        Set tdf = DBKetoan.TableDefs(tbl)
        tdf.Fields.Delete fld
    End If
End Sub

Public Sub InitDateVars(MedNgay As MaskEdBox, ngay As Date)
    Dim i As Integer, m As String, c As Integer
    
    For i = 1 To Len(Mask_D)
        c = Asc(Mid(Mask_D, i, 1))
        If (c > 64 And c < 91) Or (c > 96 And c < 123) Then m = m + "9" Else m = m + Chr(c)
    Next
    MedNgay.MaxLength = Len(Mask_D)
    MedNgay.mask = m
    ngay = Date
    On Error Resume Next
    MedNgay.Text = Format(ngay, Mask_D)
    On Error GoTo 0
End Sub

Public Function Recycle(ByVal fileName As String) As Integer
    On Error GoTo KT
    Kill fileName
    On Error GoTo 0
    Exit Function
KT:
    Recycle = -1
End Function

Public Function CheckMinRez(pixelWidth As Long, pixelHeight As Long) As Boolean

'

Dim lngTwipsX As Long

Dim lngTwipsY As Long

'

' convert pixels to twips

lngTwipsX = pixelWidth * 15

lngTwipsY = pixelHeight * 15

'

' check against current settings

If lngTwipsX > Screen.Width Then

CheckMinRez = False

Else

If lngTwipsY > Screen.Height Then

CheckMinRez = False

Else

CheckMinRez = True

End If

End If

'

End Function

Public Function Cdbl5(st As String) As Double
    If IsNumeric(st) Then Cdbl5 = CDbl(st) Else Cdbl5 = 0
End Function

Public Function CInt5(st As String) As Double
    Dim X As Double
    If IsNumeric(st) Then
        X = CDbl(st)
        If X >= -32768 And X <= 32767 Then
            CInt5 = CInt(X)
        Else
            CInt5 = 0
        End If
    Else
        CInt5 = 0
    End If
End Function

Public Function CLng5(st As String) As Long
    If IsNumeric(st) Then CLng5 = CLng(st) Else CLng5 = 0
End Function

Public Sub FCenter(f As Form)
    f.Top = (Screen.Height - f.Height) / 2
    f.Left = (Screen.Width - f.Width) / 2
End Sub

Public Sub Add32Font(fileName As String)
    Dim lResult As Long
    Dim strFontPath As String, strFontname As String
    Dim hKey As Long
    
    If Len(Dir(pWinDir + "\FONTS\" + fileName)) = 0 Then
        On Error Resume Next
        FileCopy pCurDir + "FONTS\" + fileName, pWinDir + "\FONTS\" + fileName
        On Error GoTo 0
    End If
    
    If Len(Dir(pWinDir + "\FONTS\" + fileName)) = 0 Then Exit Sub
    
    'This is the font name and path
    strFontPath = Space$(MAX_PATH)
    strFontname = fileName
    If nt Then
    'Windows NT - Call and get the path to the
    '\windows\system directory
    lResult = GetWindowsDirectory(strFontPath, _
    MAX_PATH)
    If lResult <> 0 Then Mid$(strFontPath, _
    lResult + 1, 1) = "\"
    strFontPath = RTrim$(strFontPath)
    Else
    'Win95 - Call and get the path to the
    '\windows\fonts directory
    lResult = GetWindowsDirectory(strFontPath, _
    MAX_PATH)
    If lResult <> 0 Then Mid$(strFontPath, _
    lResult + 1) = "\fonts\"
    strFontPath = RTrim$(strFontPath)
    End If
    'This Actually adds the font to the system's available
    'fonts for this windows session
    lResult = AddFontResource(strFontPath + strFontname)
    ' If lResult = 0 Then MsgBox "Error Occured " & _
    "Calling AddFontResource"
    'Write the registry value to permanently install the
    'font
    lResult = RegOpenKey(HKEY_LOCAL_MACHINE, _
    "software\microsoft\windows\currentversion\" & _
    "fonts", hKey)
    lResult = RegSetValueEx(hKey, "Proscape Font " & strFontname & _
    " (TrueType)", 0, REG_SZ, ByVal strFontname, _
    Len(strFontname))
    lResult = RegCloseKey(hKey)
    'This call broadcasts a message to let all top-level
    'windows know that a font change has occured so they
    'can reload their font list
    lResult = PostMessage(HWND_BROADCAST, WM_FONTCHANGE, _
    0, 0)
    ' MsgBox "Font Added!"
End Sub

Private Function nt() As Boolean
    Dim lResult As Long
    Dim vi As OSVERSIONINFO
    vi.dwOSVersionInfoSize = Len(vi)
    lResult = GetVersionEx(vi)
    If vi.dwPlatformId And VER_PLATFORM_WIN32_NT Then
    nt = True
    Else
    nt = False
    End If
End Function

Public Function GetWinDir() As String
    ' returns Windows directory
    Dim Buffer As String * 254, r As Long, sDir As String
    r = GetWindowsDirectory(Buffer, 254)
    sDir = Left(Buffer, r)
    If Right(sDir, 1) = "\" Then sDir = Left(sDir, Len(sDir) - 1)
    GetWinDir = sDir
End Function

Public Function VString(st As String) As String
    Select Case FontFlag
        Case 0:
            VString = UNICODEtoABC(st)
        Case 2:
            VString = VNItoABC(st, 1)
        Case Else
            VString = st
    End Select
End Function
' Them 12 thang trong nam tai chinh vao combo
Public Sub AddMonthToCbo(Cbo As ComboBox)
    Dim i As Integer
    
    Cbo.Clear
    For i = pThangDauKy To 12
        Cbo.AddItem CStr(i) + "/" + CStr(pNamTC)
        Cbo.ItemData(Cbo.NewIndex) = i
    Next
    For i = 1 To pThangDauKy - 1
        Cbo.AddItem CStr(i) + "/" + CStr(pNamTC + 1)
        Cbo.ItemData(Cbo.NewIndex) = i
    Next
    Cbo.ListIndex = IIf(pThang >= pThangDauKy, pThang - pThangDauKy, pThang - pThangDauKy + 12)
End Sub
' Kiem tra dieu kien thang x nam trong khoang tu tdau den tcuoi
Public Function InMonth(X As Integer, tdau As Integer, tcuoi As Integer) As Boolean
    If tdau <= tcuoi Then
        InMonth = (tdau <= X And X <= tcuoi)
    Else
        InMonth = (tdau <= X And X <= 12) Or (1 <= X And X <= tcuoi)
    End If
End Function
' Tra ve chuoi xac dinh manh de f>=tdau and f<=tcuoi
Public Function WThang(f As String, tdau As Integer, tcuoi As Integer) As String
    If tdau = 0 And tcuoi = 0 Then
        WThang = " (" + f + "=0) "
        Exit Function
    End If
    If tdau <> 0 And tcuoi <> 0 Then
        If tdau <= tcuoi Then
            If tdau < pThangDauKy And pThangDauKy < tcuoi Then
                WThang = " (" + f + ">=" + CStr(tdau) + " AND " + f + "<=" + CStr(pThangDauKy - 1) + ") "
            Else
                WThang = " (" + f + ">=" + CStr(tdau) + " AND " + f + "<=" + CStr(tcuoi) + ") "
            End If
        Else
            WThang = " ((" + f + ">=" + CStr(tdau) + " AND " + f + "<=13) OR (" + f + ">=1 AND " + f + "<=" + CStr(tcuoi) + ")) "
        End If
    Else
        If tdau = 0 Then
            If tcuoi >= pThangDauKy Then
                WThang = " ((" + f + ">=" + CStr(pThangDauKy) + " OR " + f + "=0) AND " + f + "<=" + CStr(tcuoi) + ") "
            Else
                If tcuoi <> 0 Then
                    WThang = " ((" + f + ">=" + CStr(pThangDauKy) + " AND " + f + "<=13) OR (" + f + ">=0 AND " + f + "<=" + CStr(tcuoi) + ")) "
                Else
                    WThang = " (" + f + "=0)"
                End If
            End If
        Else
            If tdau >= pThangDauKy Then
                WThang = " ((" + f + ">=" + CStr(tdau) + " AND " + f + "<=13) OR (" + f + ">=1 AND " + f + "<=" + CStr(pThangDauKy - 1) + ")) "
            Else
                WThang = " ((" + f + ">=" + CStr(tdau) + " AND " + f + "<=" + CStr(pThangDauKy - 1) + ") OR (" + f + "=13))"
            End If
        End If
    End If
End Function
' Tra ve chuoi xac dinh manh de f>tdau and f<tcuoi
Public Function WThang2(f As String, tdau As Integer, tcuoi As Integer)
    If tdau = 0 And tcuoi = 0 Then
        WThang2 = " (False) "
        Exit Function
    End If
    If tdau <> 0 And tcuoi <> 0 Then
        If tdau <= tcuoi Then
            If tdau < pThangDauKy And pThangDauKy < tcuoi Then
                WThang2 = " (" + f + ">" + CStr(tdau) + " AND " + f + "<=" + CStr(pThangDauKy - 1) + ") "
            Else
                WThang2 = " (" + f + ">" + CStr(tdau) + " AND " + f + "<" + CStr(tcuoi) + ") "
            End If
        Else
            WThang2 = " ((" + f + ">" + CStr(tdau) + " AND " + f + "<=13) OR (" + f + ">=1 AND " + f + "<" + CStr(tcuoi) + ")) "
        End If
    Else
        If tdau = 0 Then
            If tcuoi >= pThangDauKy Then
                WThang2 = " ((" + f + ">=" + CStr(pThangDauKy) + " OR " + f + "=0) AND " + f + "<" + CStr(tcuoi) + ") "
            Else
                If tcuoi <> 0 Then
                    WThang2 = " ((" + f + ">=" + CStr(pThangDauKy) + " AND " + f + "<=13) OR (" + f + ">=0 AND " + f + "<" + CStr(tcuoi) + ")) "
                Else
                    WThang2 = " (" + f + "=0) "
                End If
            End If
        Else
            If tdau >= pThangDauKy Then
                WThang2 = " ((" + f + ">" + CStr(tdau) + " AND " + f + "<=13) OR (" + f + ">=1 AND " + f + "<=" + CStr(pThangDauKy - 1) + ")) "
            Else
                WThang2 = " ((" + f + ">" + CStr(tdau) + " AND " + f + "<=" + CStr(pThangDauKy - 1) + ") OR (" + f + "=13))"
            End If
        End If
    End If
End Function

Public Function CThangDB(thang As Integer) As Integer
    If thang <> 0 Then
        If thang < 13 Then
            CThangDB = IIf(thang >= pThangDauKy, thang - pThangDauKy + 1, 13 - pThangDauKy + thang)
        Else
            CThangDB = 13
        End If
    Else
        CThangDB = 0
    End If
End Function

Public Function CThangFR(thang As Integer) As Integer
    If thang <> 0 Then
        CThangFR = IIf(thang <= 13 - pThangDauKy, thang + pThangDauKy - 1, thang + pThangDauKy - 13)
    Else
        CThangFR = 0
    End If
End Function

Public Function ThangTruoc(thang As Integer) As Integer
    If thang = pThangDauKy Then
        ThangTruoc = 0
    Else
        ThangTruoc = thang - 1
        If ThangTruoc = 0 Then ThangTruoc = 12
    End If
End Function

Public Function ThangSau(thang As Integer) As Integer
    If (pThangDauKy = 1 And thang = 12) Or (pThangDauKy > 1 And thang = pThangDauKy - 1) Then
        ThangSau = 13
    Else
        ThangSau = thang + 1
    End If
End Function
' Xet dieu kien thang Vx >= Cx
Public Function VC(Vx As String, cx As String) As String
    VC = "IIF(" + cx + ">=" + CStr(pThangDauKy) + "," + Vx + ">=" + cx + " OR " + Vx + "<" + CStr(pThangDauKy) + "," + Vx + ">=" + cx + " AND " + Vx + "<" + CStr(pThangDauKy) + ")"
End Function

Public Function SetMonthOrder(f As String) As String
    SetMonthOrder = "IIF(" + f + "<=" + CStr(pThangDauKy) + "," + f + "+12," + f + ")"
End Function

Public Function ThoiGian(tdau As Integer, tcuoi As Integer, Optional nn As Integer = 0) As String
If nn = 0 Then
    If tdau <> tcuoi Then
        If tdau = pThangDauKy And ((pThangDauKy <> 1 And tcuoi = pThangDauKy - 1) Or (pThangDauKy = 1 And tcuoi = 12)) Then
            ThoiGian = "N¨m " + CStr(pNamTC)
        Else
            ThoiGian = "Tõ th¸ng " + CStr(tdau) + "/" + IIf(tdau >= pThangDauKy, CStr(pNamTC), CStr(pNamTC + 1))
            ThoiGian = ThoiGian + " ®Õn th¸ng " + CStr(tcuoi) + "/" + IIf(tcuoi >= pThangDauKy, CStr(pNamTC), CStr(pNamTC + 1))
        End If
    Else
        ThoiGian = "Th¸ng " + CStr(tdau) + "/" + IIf(tdau >= pThangDauKy, CStr(pNamTC), CStr(pNamTC + 1))
    End If
Else
    If tdau <> tcuoi Then
        If tdau = pThangDauKy And ((pThangDauKy <> 1 And tcuoi = pThangDauKy - 1) Or (pThangDauKy = 1 And tcuoi = 12)) Then
            ThoiGian = "Year " + CStr(pNamTC)
        Else
            ThoiGian = "From " + CStr(tdau) + "/" + IIf(tdau >= pThangDauKy, CStr(pNamTC), CStr(pNamTC + 1))
            ThoiGian = ThoiGian + " to " + CStr(tcuoi) + "/" + IIf(tcuoi >= pThangDauKy, CStr(pNamTC), CStr(pNamTC + 1))
        End If
    Else
        ThoiGian = "Month " + CStr(tdau) + "/" + IIf(tdau >= pThangDauKy, CStr(pNamTC), CStr(pNamTC + 1))
    End If
End If
End Function

Public Function ThoiGianN(ndau As Date, ncuoi As Date, Optional nn As Integer = 0) As String
If nn = 0 Then
    If ndau <> ncuoi Then
        ThoiGianN = "Tõ ngµy " + Format(ndau, Mask_DR) + " ®Õn ngµy " + Format(ncuoi, Mask_DR)
    Else
        ThoiGianN = "Ngµy " + Format(ncuoi, Mask_DR)
    End If
Else
    If ndau <> ncuoi Then
        ThoiGianN = "From " + Format(ndau, Mask_DR) + " to " + Format(ncuoi, Mask_DR)
    Else
        ThoiGianN = "Date " + Format(ncuoi, Mask_DR)
    End If
End If
End Function

Public Function QueryDaCo(qname As String) As Boolean
    Dim i As Integer
    QueryDaCo = False
    For i = 0 To DBKetoan.QueryDefs.count - 1
        If UCase(DBKetoan.QueryDefs(i).Name) = UCase(qname) Then
            QueryDaCo = True
            Exit For
        End If
    Next
End Function

Public Sub AddQuery(qname As String, Optional sql As String = "SELECT * FROM License")
    If QueryDaCo(qname) Then Exit Sub
    On Error Resume Next
    DBKetoan.QueryDefs.Append DBKetoan.CreateQueryDef(qname, sql)
    On Error GoTo 0
End Sub

Public Sub KTTBL(tbl As String)
    Dim m As Long, rs As Recordset

    Set rs = DBKetoan.OpenRecordset("SELECT * FROM " + tbl + " ORDER BY MaSo", dbOpenDynaset)
    Do While Not rs.EOF
        If m <> rs!MaSo Then
            m = rs!MaSo
        Else
            rs.Delete
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Sub InBaoCaoRPT(Optional nn As Integer = 0)
    If Not RptOK(frmMain.Rpt.ReportFileName, nn) Then Exit Sub
    On Error GoTo LoiIn
    frmMain.Rpt.WindowShowPrintSetupBtn = True
    
    frmMain.Rpt.Action = 1
    On Error GoTo 0
    Exit Sub
LoiIn:
    MsgBox "Error " + CStr(Err.number) + ": " + Err.Description, vbExclamation, App.ProductName
End Sub

Public Function ThangCuoiNamTC() As Integer
    ThangCuoiNamTC = IIf(pThangDauKy = 1, 12, pThangDauKy - 1)
End Function

Public Sub CopyTable(fname As String, tbl As String)
    Dim db As Database, tdf As TableDef, i As Integer
    
    If Len(Dir(fname)) = 0 Then Exit Sub
    If BangDaCo(tbl) Then XoaBang tbl
    Set db = WSpace.OpenDatabase(fname, True, True)
    
    Set tdf = DBKetoan.CreateTableDef(tbl)
    For i = 1 To db.TableDefs(tbl).Fields.count
        tdf.Fields.Append tdf.CreateField(db.TableDefs(tbl).Fields(i - 1).Name, db.TableDefs(tbl).Fields(i - 1).Type, db.TableDefs(tbl).Fields(i - 1).Size)
        tdf.Fields(i - 1).DefaultValue = db.TableDefs(tbl).Fields(i - 1).DefaultValue
        tdf.Fields(i - 1).Attributes = db.TableDefs(tbl).Fields(i - 1).Attributes
    Next
        
    For i = 1 To db.TableDefs(tbl).Indexes.count
        tdf.Indexes.Append tdf.CreateIndex(db.TableDefs(tbl).Indexes(i - 1).Name)
        tdf.Indexes(db.TableDefs(tbl).Indexes(i - 1).Name).Fields.Append tdf.Indexes(db.TableDefs(tbl).Indexes(i - 1).Name).CreateField(db.TableDefs(tbl).Indexes(i - 1).Fields(0).Name)
        tdf.Indexes(db.TableDefs(tbl).Indexes(i - 1).Name).Primary = db.TableDefs(tbl).Indexes(i - 1).Primary
        tdf.Indexes(db.TableDefs(tbl).Indexes(i - 1).Name).Unique = db.TableDefs(tbl).Indexes(i - 1).Unique
    Next
    DBKetoan.TableDefs.Append tdf
    
    db.Close
    Set db = Nothing
    
    ExecuteSQL5 "INSERT INTO " + tbl + " SELECT * FROM " + tbl + " IN '" + fname + "'"
End Sub


Public Sub XoaBang(tbl As String)
    If Not BangDaCo(tbl) Then Exit Sub
    DBKetoan.TableDefs.Delete tbl
End Sub

Public Function RoundMoney(tien As Double) As Double
    Dim X As Double
    X = IIf(tien >= 0, 0.5, -0.5)
    RoundMoney = IIf(pTien = 0, Fix(X + tien), Fix(X + Mask_N * tien) / Mask_N)
End Function

Public Function RoundMoneySQL(tien As String) As String
    Dim X As String
    X = "IIf(" + tien + " >= 0, 0.5, -0.5)"
    RoundMoneySQL = IIf(pTien = 0, "Fix(" + X + " + " + tien + ")", "Fix(" + X + " + " + CStr(Mask_N) + " * (" + tien + ") / " + CStr(Mask_N) + ")")
End Function

Public Function DoiRaNT(tien As Double, tygia As Double) As Double
    If tygia <> 0 Then
        If pTien = 0 Then
            DoiRaNT = Fix(0.5 + Mask_N * tien / tygia) / Mask_N
        Else
            DoiRaNT = Fix(0.5 + Mask_N * tien * tygia) / Mask_N
        End If
    End If
End Function

Public Function DoiRaTien(nt As Double, tygia As Double) As Double
    If tygia <> 0 Then
        If pTien = 0 Then
            DoiRaTien = Fix(0.5 + nt * tygia)
        Else
            DoiRaTien = Fix(0.5 + Mask_N * nt / tygia) / Mask_N
        End If
    End If
End Function

Public Function LoaiFont(tenfont As String) As Integer
    If UCase(Left(tenfont, 3)) = "VNI" Then LoaiFont = 2
    If UCase(Left(tenfont, 2)) = "MS" Or UCase(Left(tenfont, 2)) = "VK" Or UCase(Left(tenfont, 1)) = "." Then LoaiFont = 1
End Function

Public Sub AppIdle(k As Long)
    Dim i As Long
    For i = 1 To k
        DoEvents
    Next
End Sub

Public Function WNgay(f As String, ndau As Date, ncuoi As Date) As String
    WNgay = "(" + f + ">=#" + Format(ndau, Mask_DB) + "# AND " + f + "<=#" + Format(ncuoi, Mask_DB) + "#)"
End Function

Public Function WNgay2(f As String, ndau As Date, ncuoi As Date) As String
    WNgay2 = "(" + f + ">#" + Format(ndau, Mask_DB) + "# AND " + f + "<#" + Format(ncuoi, Mask_DB) + "#)"
End Function

Public Sub CopyTable2(tbl As String, tbl2 As String, Optional CP As Integer)
    Dim tdf As TableDef, i As Integer
    
    If Not BangDaCo(tbl2) Then
        Set tdf = DBKetoan.CreateTableDef(tbl2)
           
           For i = 1 To DBKetoan.TableDefs(tbl).Fields.count
            tdf.Fields.Append tdf.CreateField(DBKetoan.TableDefs(tbl).Fields(i - 1).Name, DBKetoan.TableDefs(tbl).Fields(i - 1).Type, DBKetoan.TableDefs(tbl).Fields(i - 1).Size)
            tdf.Fields(i - 1).DefaultValue = DBKetoan.TableDefs(tbl).Fields(i - 1).DefaultValue
            tdf.Fields(i - 1).Attributes = DBKetoan.TableDefs(tbl).Fields(i - 1).Attributes
        Next
        For i = 1 To DBKetoan.TableDefs(tbl).Indexes.count
            tdf.Indexes.Append tdf.CreateIndex(DBKetoan.TableDefs(tbl).Indexes(i - 1).Name)
            tdf.Indexes(DBKetoan.TableDefs(tbl).Indexes(i - 1).Name).Fields.Append tdf.Indexes(DBKetoan.TableDefs(tbl).Indexes(i - 1).Name).CreateField(DBKetoan.TableDefs(tbl).Indexes(i - 1).Fields(0).Name)
            tdf.Indexes(DBKetoan.TableDefs(tbl).Indexes(i - 1).Name).Primary = DBKetoan.TableDefs(tbl).Indexes(i - 1).Primary
            tdf.Indexes(DBKetoan.TableDefs(tbl).Indexes(i - 1).Name).Unique = DBKetoan.TableDefs(tbl).Indexes(i - 1).Unique
        Next
       
        DBKetoan.TableDefs.Append tdf
        If CP > 0 Then ExecuteSQL5 "INSERT INTO " + tbl2 + " SELECT * FROM " + tbl
    End If
End Sub

Public Function GetDiskSpace() As Long
Dim Sectors As Long
Dim Bytes As Long
Dim freeClusters As Long
Dim totalClusters As Long
Dim retValue As Long

    retValue = GetDiskFreeSpace(Left(pDataPath, 2) & "\", Sectors, Bytes, freeClusters, totalClusters)
    If retValue > 0 Then GetDiskSpace = Fix(Sectors * Bytes * (freeClusters / 1048576))
End Function

Public Function ChungTu2TKNC(loai As Integer, Optional p As Integer = 0) As String
    Dim sh As String
    sh = IIf(p > 0, "P", "")
    Select Case loai
        Case 0:
            ChungTu2TKNC = " (ChungTu" + sh + " INNER JOIN HethongTK ON ChungTu" + sh + ".MaTKNo=HethongTK.MaSo) INNER JOIN HethongTK AS TK ON ChungTu" + sh + ".MaTKCo=TK.MaSo "
        Case 10:
            ChungTu2TKNC = " (ChungTu" + sh + " INNER JOIN HethongTK ON ChungTu" + sh + ".MaTKTCNo=HethongTK.MaSo) INNER JOIN HethongTK AS TK ON ChungTu" + sh + ".MaTKTCCo=TK.MaSo "
        Case -1:
            ChungTu2TKNC = " ChungTu" + sh + " INNER JOIN HethongTK ON ChungTu" + sh + ".MaTKNo=HethongTK.MaSo "
        Case 1:
            ChungTu2TKNC = " ChungTu" + sh + " INNER JOIN HethongTK ON ChungTu" + sh + ".MaTKCo=HethongTK.MaSo "
        Case -2:
            ChungTu2TKNC = " ChungTu" + sh + " INNER JOIN HethongTK ON ChungTu" + sh + ".MaTKTCNo=HethongTK.MaSo "
        Case 2:
            ChungTu2TKNC = " ChungTu" + sh + " INNER JOIN HethongTK ON ChungTu" + sh + ".MaTKTCCo=HethongTK.MaSo "
    End Select
End Function

Public Function ChungTu2TKHD(loai As Integer) As String
    Select Case loai
        Case 0:
            ChungTu2TKHD = " HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo "
        Case 10:
            ChungTu2TKHD = " (HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) LEFT JOIN KhachHang ON HoaDon.MaKhachHang=KhachHang.MaSo "
        Case -1:
            ChungTu2TKHD = " (HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo "
        Case 1:
            ChungTu2TKHD = " (HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo "
        Case -2:
            ChungTu2TKHD = " ((HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN KhachHang ON HoaDon.MaKhachHang=KhachHang.MaSo "
        Case 2:
            ChungTu2TKHD = " ((HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) LEFT JOIN KhachHang ON HoaDon.MaKhachHang=KhachHang.MaSo "
    End Select
End Function

Public Sub DeleteRel()
    Dim i As Integer, j As Integer, L As Integer
    
    On Error GoTo n
    For i = 0 To DBKetoan.TableDefs.count - 1
        If Left(DBKetoan.TableDefs(i).Name, 4) <> "MSys" Then
            j = 0
            Do While j < DBKetoan.TableDefs(i).Indexes.count
                L = Len(DBKetoan.TableDefs(i).Name)
                If Right(DBKetoan.TableDefs(i).Indexes(j).Name, L) = DBKetoan.TableDefs(i).Name Then

                    DBKetoan.TableDefs(i).Indexes.Delete DBKetoan.TableDefs(i).Indexes(j).Name
                    
                Else
n:
                    j = j + 1
                End If
            Loop
        End If
    Next
    On Error GoTo 0
End Sub

Public Sub FixCode(tbl As String, fld As String)
    Dim rs As Recordset, m As Long
    
    If Not BangDaCo(tbl) Then Exit Sub
    Set rs = DBKetoan.OpenRecordset("SELECT " + fld + " FROM " + tbl + " ORDER BY " + fld, dbOpenDynaset)
    Do While Not rs.EOF
        If rs.Fields(fld).Value = m Then
            rs.Delete
        Else
            m = rs.Fields(fld).Value
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Function GetComputerName1() As String
' This routine will obtain the Computers name from the system. The
' second time it is called it returns the static 'sName' variable
Dim lBuffLen As Long
Dim sBuffer As String
Dim lRet As Long

lBuffLen = 128
sBuffer = String$(lBuffLen, vbNullChar)
lRet = GetComputerName(sBuffer, lBuffLen)
If lRet < 0 Then
' Handle API error Here
Exit Function
End If
GetComputerName1 = Left$(sBuffer, lBuffLen)
End Function

Public Function NgayCuoiNam() As Date
    If pThangDauKy = 1 Then
        NgayCuoiNam = NgayCuoiThang(pNamTC, 12)
    Else
        NgayCuoiNam = NgayCuoiThang(pNamTC + 1, pThangDauKy - 1)
    End If
End Function

Public Sub XLSFooter(xls As Worksheet, Row As Integer, w As Integer, n As Date)
    xls.Range("A" + CStr(Row), XLSCol(w) + CStr(Row)).MergeCells = True
    xls.Range("A" + CStr(Row), XLSCol(w) + CStr(Row)).HorizontalAlignment = xlHAlignRight
    xls.Range("A" + CStr(Row), XLSCol(w) + CStr(Row)).Font.Italic = True
    xls.Cells(Row, 1) = "Ngµy " + CStr(Day(n)) + " th¸ng " + CStr(Month(n)) + " n¨m " + CStr(Year(n))
    xls.Range("A" + CStr(Row + 1), XLSCol(w) + CStr(Row + 1)).MergeCells = True
    xls.Range("A" + CStr(Row + 1), XLSCol(w) + CStr(Row + 1)).HorizontalAlignment = xlHAlignCenter
    xls.Range("A" + CStr(Row + 1), XLSCol(w) + CStr(Row + 1)).Font.Bold = True
    xls.Cells(Row + 1, 1) = "Ng­êi lËp biÓu                                                        KÕ to¸n tr­ëng                                                        Gi¸m ®èc"
End Sub

Public Sub GridSelAll(g As Grid)
    Dim i As Integer, r As Integer
    
    With g
        .col = 0
        For i = 0 To .Rows - 1
            .Row = i
            If .Text = "" Then
                Exit For
            Else
                r = r + 1
            End If
        Next
        If r = 0 Then r = 1
        .SelStartCol = 0
        .SelEndCol = .Cols - 1
        .SelStartRow = 0
        .SelEndRow = r - 1
    End With
End Sub

Public Function SHCtuMoi(shct As String) As String
    Dim i As Integer, tail As String
        
    If Len(shct) = 0 Then Exit Function
    For i = 1 To Len(shct)
        If Not IsNumeric(Right(shct, i)) Then Exit For
    Next
    i = i - 1
    If i > 0 And i < 20 Then
        tail = CStr(Fix(Cdbl5(Right(shct, i))) + 1)
        Do While Len(tail) < i
            tail = "0" + tail
        Loop
    End If
    SHCtuMoi = Left(shct, Len(shct) - i) + tail
    If SHCtuMoi = shct Then SHCtuMoi = ""
End Function

Public Sub WCenter(frm As Form)
    With frm
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
    End With
End Sub

Public Function SoLanXuatHien(st As String, c As Integer) As Integer
    Dim i As Integer, L As Integer, k As Integer
    
    L = Len(st)
    For i = 1 To L
        If Asc(Mid(st, i, 1)) = c Then k = k + 1
    Next
    SoLanXuatHien = k
End Function

Public Function SetNumericStr(s As String) As String
    Dim i As Integer, s1 As String, a As Integer, k As Integer
    
    For i = 1 To Len(s)
        a = Asc(Mid(s, i, 1))
        If (a > 47 And a < 58) Or (a = 45 And i > 10 And k < 2) Then s1 = s1 + Chr(a)
        If a = 45 Then k = k + 1
    Next
    SetNumericStr = s1
End Function

Public Function NewRowIndex(Grd As Grid, col As Integer) As Integer
    Dim i As Integer
    
    With Grd
        .col = col
        For i = 0 To .Rows - 1
            .Row = i
            If Len(.Text) = 0 Then Exit For
        Next
        NewRowIndex = i
    End With
End Function

'Public Sub SetServer(act As Integer)
'Dim hregkey As Long
'Dim subkey As String
'Dim stringbuffer As String

 '   subkey = "Software\Microsoft\Windows\CurrentVersion\Run"
    
  '  retval = RegOpenKeyEx(HKEY_CURRENT_USER, subkey, 0, KEY_WRITE, hregkey)
   ' If retval <> 0 Then Exit Sub
    
    'stringbuffer = IIf(act > 0, pCurDir & App.EXEName & ".EXE" & vbNullChar, vbNullChar)
    'retval = RegSetValueEx(hregkey, App.ProductName, 0, REG_SZ, ByVal stringbuffer, Len(stringbuffer))
    
    'RegCloseKey hregkey
'End Sub

Public Function GetShortDateFormat() As String
   Dim sReturn As String, dwLocaleID As Long, r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, LOCALE_SSHORTDATE, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, LOCALE_SSHORTDATE, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetShortDateFormat = Left$(sReturn, r - 1)
      
      End If
   End If
End Function

Public Sub SearchObj(id As Integer, Optional o1 As ListBox, Optional o2 As Grid, Optional col As Integer = 0)
    Dim st As String, i As Integer, k As Integer
    
    st = FrmGetStr.GetString("Tõ kho¸ cÇn t×m", "T×m kiÕm")
    If Len(st) > 0 Then
        Select Case id
            Case 0:
                k = o1.ListIndex
                For i = 0 To o1.ListCount - 1
                    If InStr(o1.List(i), st) > 0 Then
                        o1.ListIndex = i
                        GoTo KT1
                    End If
                Next
                o1.ListIndex = k
KT1:
                RFocus o1
            Case 1:
                With o2
                    .col = col
                    k = .Row
                    For i = 0 To .Rows - 1
                        .Row = i
                        If InStr(.Text, st) > 0 Then
                            If Not .RowIsVisible(i) Then .TopRow = i - 1
                            GoTo KT
                        End If
                    Next
                    .Row = k
                End With
KT:
                RFocus o2
        End Select
    End If
End Sub

Public Sub SetListIndex2(combo_box As Object, sh As String)
Dim n As Integer
      If combo_box.ListCount = 0 Or sh = "" Then Exit Sub
      For n = 0 To combo_box.ListCount - 1
        If Left(combo_box.List(n), Len(sh)) = sh Then
            combo_box.ListIndex = n
            Exit For
        End If
      Next
End Sub

Private Function Register(fname$, Value%) As Integer
    Dim regLib&, process&, succeed&
    Dim h1&, xc&, id&
    Dim p$
    
    Select Case Value
        Case 0: p = "DllUnregisterServer"
        Case 1: p = "DllRegisterServer"
        Case Else: Register = 0
                    Exit Function
    End Select

    regLib = LoadLibraryRegister(fname)
    If regLib = 0 Then
        Register = 1
        Exit Function
    End If
        
    process = GetProcAddressRegister(regLib, p)
    
    If process = 0 Then
        Register = 2
    Else
        h1 = CreateThreadForRegister(ByVal 0&, 0&, _
            ByVal process, ByVal 0&, 0&, id)
        If h1 = 0 Then
            Register = 3
        Else
            succeed = (WaitForSingleObject(h1, 10000) = 0)
            If succeed Then
                CloseHandle h1
                Register = 4
            Else
                GetExitCodeThread h1, xc
                ExitThread xc
                Register = 5
            End If
        End If
    End If

    FreeLibraryRegister regLib
End Function

Public Function findwindowpartial(ByVal titlepart$) As Long
    Dim hwndtmp As Long
    Dim nRet As Long
    Dim titletmp As String

    titlepart = UCase(titlepart)
    hwndtmp = FindWindow(0&, 0&)
    Do Until hwndtmp = 0
        If GetParent(hwndtmp) = 0 Then
            titletmp = Space(256)
            nRet = GetWindowText(hwndtmp, titletmp, Len(titletmp))
            If nRet Then
                titletmp = UCase(Left(titletmp, nRet))
                        If InStr(titletmp, titlepart) > 0 Then
                            findwindowpartial = hwndtmp
                            Exit Do
                        End If
            End If
        End If
        hwndtmp = GetWindow(hwndtmp, gw_hwndnext)
    Loop
End Function

Public Sub setMDSettings()
    
    blnMDSettingsChanged = False
' read short date
    LCID = GetUserDefaultLCID()
    iRet = GetLocaleInfo(LCID, LOCALE_SSHORTDATE, lpLCDataVar, 0)
    Symbol = String$(iRet, 0)
    iRet2 = GetLocaleInfo(LCID, LOCALE_SSHORTDATE, Symbol, iRet)
    pos = InStr(Symbol, Chr$(0))
    If pos > 0 Then
        Symbol = Left$(Symbol, pos - 1)
    End If
    
    If Symbol <> "dd/MM/yy" Then
        'change thousand separator
        blnMDSettingsChanged = True
        old_LOCALE_SSHORTDATE = Symbol
        LCID = GetUserDefaultLCID()
        Call SetLocaleInfo(LCID, LOCALE_SSHORTDATE, "dd/MM/yy")
        Call PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, ByVal 0&)
    End If
End Sub

Public Sub restoreSettings()
'restore the original settings in cp
    If old_LOCALE_SSHORTDATE <> vbNullString Then
    ' restore short date
        LCID = GetUserDefaultLCID()
        Call SetLocaleInfo(LCID, LOCALE_SSHORTDATE, old_LOCALE_SSHORTDATE)
        Call PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, ByVal 0&)
    End If
End Sub

Public Sub SetRunAtStartup(ByVal app_name As String, ByVal app_path As String, Optional ByVal run_at_startup As Boolean = True)
Dim hKey As Long
Dim key_value As String
Dim Status As Long

    On Error GoTo SetStartupError

    ' Open the key, creating it if it doesn't exist.
    If RegCreateKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Run", _
        ByVal 0&, ByVal 0&, ByVal 0&, _
        KEY_WRITE, ByVal 0&, hKey, _
        ByVal 0&) <> ERROR_SUCCESS _
    Then
        MsgBox "Error " & Err.number & " opening key" & _
            vbCrLf & Err.Description
        Exit Sub
    End If

    ' See if we should run at startup.
    If run_at_startup Then
        ' Create the key.
        key_value = app_path & "\" & app_name & ".exe" & vbNullChar
        Status = RegSetValueEx(hKey, App.EXEName, 0, REG_SZ, _
            ByVal key_value, Len(key_value))

        If Status <> ERROR_SUCCESS Then
            MsgBox "Error " & Err.number & " setting key" & _
                vbCrLf & Err.Description
        End If
    Else
        ' Delete the value.
        RegDeleteValue hKey, app_name
    End If

    ' Close the key.
    RegCloseKey hKey
    Exit Sub

SetStartupError:
    MsgBox Err.number & " " & Err.Description
    Exit Sub
End Sub
' Return True if the program is set to run at startup.
Public Function WillRunAtStartup(ByVal app_name As String) As Boolean
Dim hKey As Long
Dim value_type As Long

    ' See if the key exists.
    If RegOpenKeyEx(HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Run", _
        0, KEY_READ, hKey) = ERROR_SUCCESS _
    Then
        ' Look for the subkey named after the application.
        WillRunAtStartup = _
            (RegQueryValueEx(hKey, app_name, _
                ByVal 0&, value_type, ByVal 0&, ByVal 0&) = _
            ERROR_SUCCESS)

        ' Close the registry key handle.
        RegCloseKey hKey
    Else
        ' Can't find the key.
        WillRunAtStartup = False
    End If
End Function

Public Sub NenTep(f1 As String, F2 As String)
    On Error Resume Next
    FrmZip.Show 0
    On Error GoTo 0
    FrmZip.zip.InputFile = f1
    FrmZip.zip.OutputFile = F2
    On Error Resume Next
    FrmZip.zip.Compress
    On Error GoTo 0
    Unload FrmZip
    Set FrmZip = Nothing
End Sub

Public Sub GianTepNen(f1 As String, F2 As String)
    On Error Resume Next
    FrmZip.Show 0
    On Error GoTo 0
    FrmZip.zip.InputFile = f1
    FrmZip.zip.OutputFile = F2
    On Error Resume Next
    FrmZip.zip.Decompress
    On Error GoTo 0
    Unload FrmZip
    Set FrmZip = Nothing
End Sub

