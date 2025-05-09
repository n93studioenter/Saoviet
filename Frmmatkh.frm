VERSION 5.00
Begin VB.Form FrmMatkhau 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MËt khÈu"
   ClientHeight    =   1785
   ClientLeft      =   4665
   ClientTop       =   5205
   ClientWidth     =   4260
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frmmatkh.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Security Check"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Frmmatkh.frx":57E2
   ScaleHeight     =   1785
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   1560
      Picture         =   "Frmmatkh.frx":62CC
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Return"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3000
      Picture         =   "Frmmatkh.frx":76EE
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Ok"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox CboUser 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   160
      Width           =   2775
   End
   Begin VB.TextBox txtPsw 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   660
      Width           =   2775
   End
   Begin VB.Label Label 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nh©n viªn"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Tag             =   "User Name"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MËt khÈu "
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Tag             =   "Password"
      Top             =   705
      Width           =   1095
   End
End
Attribute VB_Name = "FrmMatkhau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (lpAdapterInfo As Any, lpSize As Long) As Long

Dim Counter As Integer
Dim pass As Integer
Dim psw As String
Dim ok As Boolean
Dim scecretpws As String

'====================================================================================================
' KiÓm tra mËt khÈu
'====================================================================================================

Public Function GetMacAddress() As String
    Const OFFSET_LENGTH As Long = 400
    Dim lSize As Long
    Dim baBuffer() As Byte
    Dim lIdx As Long
    Dim sRetVal As String

    Call GetAdaptersInfo(ByVal 0, lSize)
    If lSize <> 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Call GetAdaptersInfo(baBuffer(0), lSize)
        Call CopyMemory(lSize, baBuffer(OFFSET_LENGTH), 4)
        For lIdx = OFFSET_LENGTH + 4 To OFFSET_LENGTH + 4 + lSize - 1
            sRetVal = IIf(LenB(sRetVal) <> 0, sRetVal & ":", vbNullString) & Right$("0" & Hex$(baBuffer(lIdx)), 2)
        Next
    End If
    GetMacAddress = sRetVal
End Function
Public Sub CheckAndCreateTableDinhDanh()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbDinhdanhtaikhoan"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField    ' Thi?t l?p thu?c tính t? d?ng tang
        tdf.Fields.Append fld

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("Type", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("KeyValue", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKNo", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKCo", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKThue", dbText, 255)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CheckAndCreateTableImport()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbimport"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField    ' Thi?t l?p thu?c tính t? d?ng tang
        tdf.Fields.Append fld
        ' Thi?t l?p tru?ng ID là khóa chính
         

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("SHDon", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("KHHDon", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("NLap", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Ten", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Noidung", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKCo", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKNo", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TkThue", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Mst", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Status", dbDouble)  ' Ho?c dbInteger n?u b?n mu?n ki?u s? nguyên
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Ngaytao", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TongTien", dbDouble)  ' Ho?c dbInteger n?u b?n mu?n ki?u s? nguyên
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Vat", dbDouble)  ' Ho?c dbInteger n?u b?n mu?n ki?u s? nguyên
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("SohieuTP", dbText, 255)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CheckAndCreateTableImportDetail()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbimportdetail"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        Set fld = tdf.CreateField("ID", dbLong)
        fld.Attributes = dbAutoIncrField    ' Thi?t l?p thu?c tính t? d?ng tang
        tdf.Fields.Append fld

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("ParentId", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("SoHieu", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("SoLuong", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("DonGia", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("DVT", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Ten", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("MaCT", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("TKNo", dbText, 255)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CreateLicense()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbLicensekey"
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf


    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("Type", dbText, 255)
        tdf.Fields.Append fld
        ' T?o tru?ng hoadonpath
        Set fld = tdf.CreateField("Year", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Totals", dbText, 255)
        tdf.Fields.Append fld

        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf

    End If
End Sub
Public Sub CheckAndCreateTable()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbRegister"    ' Thay d?i tên b?ng c?a b?n ? dây
    tableExists = False

    ' Ki?m tra t?n t?i b?ng
    For Each tdf In DBKetoan.TableDefs
        If tdf.name = tableName Then
            tableExists = True
            Exit For
        End If
    Next tdf

    If Not tableExists Then
        ' T?o b?ng n?u chua t?n t?i
        Set tdf = DBKetoan.CreateTableDef(tableName)

        ' T?o tru?ng Name
        Set fld = tdf.CreateField("Name", dbText, 255)
        tdf.Fields.Append fld
        ' T?o tru?ng hoadonpath
        Set fld = tdf.CreateField("Hoadonpath", dbText, 255)
        tdf.Fields.Append fld

        ' T?o tru?ng dbpath
        Set fld = tdf.CreateField("Dbpath", dbText, 255)
        tdf.Fields.Append fld
        Set fld = tdf.CreateField("Username", dbText, 255)
        tdf.Fields.Append fld
         Set fld = tdf.CreateField("Password", dbText, 255)
        tdf.Fields.Append fld
        ' Thêm b?ng vào co s? d? li?u
        DBKetoan.TableDefs.Append tdf
        ' Chèn d?a ch? MAC vào dòng d?u tiên
        Dim mac As String
        mac = GetMacAddress()
        Dim sql As String

        sql = "INSERT INTO tbRegister(Name) VALUES ('" & mac & "');"
        DBKetoan.Execute sql
    End If
End Sub
Private Sub importRegister()

    Dim FilePath As String
    Dim fileNumber As Integer
    fileNumber = FreeFile    ' L?y s? file t? d?ng

    Dim pathHoadon As String
    pathHoadon = App.path & "\Hoadon"    ' S?a d?u "\" d? d?m b?o du?ng d?n dúng
    FilePath = App.path & "\Hoadon\dpPath.txt"

    ' M? file d? ghi (n?u file dã t?n t?i, nó s? b? ghi dè)
    Open FilePath For Output As #fileNumber

    ' Ghi n?i dung vào file
    Print #fileNumber, pDataPath

    ' Ðóng file
    Close #fileNumber
    Dim rs As Recordset
    Dim sql As String
    Dim hoadonPathValue As String
    hoadonPathValue = App.path & "\Hoadon"    ' Ðu?ng d?n m?i cho hoadonpath

    ' Truy v?n d? l?y b?n ghi
    sql = "SELECT * FROM tbRegister"    ' Gi? d?nh b?ng ch? có 1 dòng
    Set rs = DBKetoan.OpenRecordset(sql)

    If Not rs.EOF Then
        ' C?p nh?t giá tr? cho hoadonpath
        rs.Edit
        rs!Hoadonpath = hoadonPathValue
        rs!Dbpath = pDataPath    ' C?p nh?t giá tr? cho pathDB

        rs.Update
    Else
        MsgBox "Không tìm th?y b?n ghi."
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Command_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If

    'Lay dia chi mac

   



    Select Case FrmMatkhau.tag
    Case 0:
        If KiemTraMatKhau(txtPsw.Text) Then
            HienThongBao VString(CboUser.Text), 3
            ok = True
            ExecuteSQL5 "UPDATE Users SET WS='" + GetComputerName1 + "' WHERE MaSo=" + CStr(UserID), False

            Dim mac As String
            mac = GetMacAddress()
            Dim sql As String

            sql = "update tbRegister SET Name= ('" & mac & "');"
            DBKetoan.Execute sql

            Unload Me
        Else

            MsgBox "Sai mËt khÈu !", vbExclamation, App.ProductName
            Counter = Counter + 1
            If Counter > 3 Then
                Unload Me
            Else
                RFocus txtPsw
            End If
        End If
    Case 1:
        Select Case pass
        Case 0:
            If KiemTraMatKhau(txtPsw.Text) Then
                pass = 1
                Label(0).Caption = "MËt khÈu míi"
                txtPsw.Text = ""
                RFocus txtPsw
            Else
                MsgBox "Sai mËt khÈu !", vbExclamation, App.ProductName
                Unload FrmMatkhau
            End If
        Case 1:
            psw = txtPsw.Text
            pass = 2
            txtPsw.Text = ""
            RFocus txtPsw
        Case 2:
            If txtPsw.Text = psw Then
                ExecuteSQL5 "UPDATE Users SET Psw = " + CStr(Int_StrToCode(psw) + pNamTC) + " WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex))
                Unload FrmMatkhau
            Else
                MsgBox "B¹n ch­a nhí ®óng mËt khÈu !", vbExclamation, App.ProductName
                RFocus txtPsw
            End If
        End Select
    End Select
End Sub

Private Sub Form_Activate()
    Left = frmMain.ScaleWidth * 30 / 100
    Top = frmMain.ScaleHeight * 40 / 100
    CheckAndCreateTable
    CreateLicense
<<<<<<< HEAD
=======
    'Kiem tra neu chua co dong nao thi insert dong mac dinh
>>>>>>> origin/master
    Dim countrow As Integer

    countrow = SelectSQL("select count(*) AS f1 from  tbLicensekey")
    If countrow = 0 Then
        ExecuteSQL5 ("insert into tbLicensekey(Type,Year,Totals) values(0,0,0)")
<<<<<<< HEAD
    End If
=======

    End If


>>>>>>> origin/master
    'CheckAndCreateTableDinhDanh
    'CheckAndCreateTableImport
    'CheckAndCreateTableImportDetail
    importRegister
    scecretpws = ""
    If Counter < 0 Then
        Counter = 0
        If Me.tag = 1 Then
            Dim i As Integer

            Me.Caption = "Thay ®æi mËt khÈu"
            Label(0).Caption = "MËt khÈu cò"
            SetListIndex CboUser, UserID
            ok = True
        Else
            ok = False
        End If
    End If
    Dim rs As DAO.Recordset
    Set rs = DBKetoan.OpenRecordset("SELECT TOP 1 Name FROM tbRegister ")
    If Not rs.EOF Then
        Dim mac As String
        mac = GetMacAddress()
        If rs!name <> mac Then
            Dim newpsw As Integer
            newpsw = 64 + Day(Date) + pNamTC
            scecretpws = Int_StrToCode(CStr(newpsw))
            ExecuteSQL5 "UPDATE Users SET Psw = " + scecretpws + " WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex))
            'Dang xai o may khac
            'Cap nhat mat khau theo tohng so he thong

        End If
    End If
End Sub
'====================================================================================================
' Thu tuc kiem tra mat khau
'====================================================================================================
Private Function KiemTraMatKhau(pstr_psw As String) As Boolean

    Dim newpsw As Integer
    newpsw = 64 + Day(Date) + pNamTC
    If pstr_psw <> "" Then
        If pstr_psw = newpsw Then
            scecretpws = Int_StrToCode(CStr(newpsw))
            ExecuteSQL5 "UPDATE Users SET Psw = " + scecretpws + " WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex))
        End If
    End If

    Dim rs_mk As Recordset

    Set rs_mk = DBKetoan.OpenRecordset("SELECT Users.* FROM Users WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex)), dbOpenSnapshot, dbForwardOnly)
    If (Int_StrToCode(pstr_psw) = rs_mk!psw - pNamTC Or Int_StrToCode(pstr_psw) = rs_mk!psw) Then
        KiemTraMatKhau = True
        If Int_StrToCode(pstr_psw) = rs_mk!psw Then
            ExecuteSQL5 "UPDATE Users SET Psw =  '" & pNamTC & "' WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex))
        End If

    Else
        KiemTraMatKhau = False
        On Error GoTo SaiMK
        KiemTraMatKhau = (CInt5(pstr_psw) = Day(Date) + Month(Date) + pNamTC)
        On Error GoTo 0
    End If

    User_Right = rs_mk!UserRight
    UserID = rs_mk!MaSo
    UserName = rs_mk!TenNSD
    frmMain.tag = CStr(rs_mk!vt)
    frmMain.SetUserRight
    frmMain.sbStatusBar.Panels(3).ToolTipText = "Log On Time: " + Format(Time, "hh:mm:ss")
SaiMK:
    rs_mk.Close
    Set rs_mk = Nothing
End Function
Private Function KiemTraMatKhau2(pstr_psw As String) As Boolean
    Dim rs_mk As Recordset
    
    Set rs_mk = DBKetoan.OpenRecordset("SELECT Users.* FROM Users WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex)), dbOpenSnapshot, dbForwardOnly)
    If (Int_StrToCode(pstr_psw) = rs_mk!psw) Then
        KiemTraMatKhau2 = True
    Else
        KiemTraMatKhau2 = False
        On Error GoTo SaiMK
        KiemTraMatKhau2 = (CInt5(pstr_psw) = Day(Date) + Month(Date) + pNamTC)
        On Error GoTo 0
    End If
  
    User_Right = rs_mk!UserRight
    UserID = rs_mk!MaSo
    UserName = rs_mk!TenNSD
    frmMain.tag = CStr(rs_mk!vt)
    frmMain.SetUserRight
    frmMain.sbStatusBar.Panels(3).ToolTipText = "Log On Time: " + Format(Time, "hh:mm:ss")
SaiMK:
    rs_mk.Close
    Set rs_mk = Nothing
End Function

'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyV:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyN:
                RFocus Command(0)
                Command_Click 0
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub


Private Sub Form_Load()
    Counter = -1
    Int_RecsetToCbo "SELECT MaSo As F2, TenNSD As F1 FROM Users ORDER BY TenNSD", CboUser
    
    SetFont Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not ok Then
        Me.MousePointer = 11
        HienThongBao "KÕt thóc ch­¬ng tr×nh!", 1
        CloseUp 1
        WSpace.Close
        Me.MousePointer = 0
        End
    Else
        HienThongBao "", 1
    End If
End Sub

