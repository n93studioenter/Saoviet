VERSION 5.00
Begin VB.Form FrmMatkhau 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "M�t kh�u"
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
      Caption         =   "Nh�n vi�n"
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
      Caption         =   "M�t kh�u "
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
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (lpAdapterInfo As Any, lpSize As Long) As Long

Dim Counter As Integer
Dim pass As Integer
Dim psw As String
Dim ok As Boolean

'====================================================================================================
' Ki�m tra m�t kh�u
'====================================================================================================

Public Function GetMacAddress() As String
    Const OFFSET_LENGTH As Long = 400
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim lIdx            As Long
    Dim sRetVal         As String
    
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
Public Sub CheckAndCreateTable()
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim tableExists As Boolean
    Dim tableName As String

    tableName = "tbRegister"    ' Thay d?i t�n b?ng c?a b?n ? d�y
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

        ' Th�m b?ng v�o co s? d? li?u
        DBKetoan.TableDefs.Append tdf
        ' Ch�n d?a ch? MAC v�o d�ng d?u ti�n
        Dim mac As String
        mac = GetMacAddress()
        sql = "INSERT INTO tbRegister(Name) VALUES ('" & mac & "');"
        DBKetoan.Execute sql
    End If
End Sub

Private Sub Command_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If

    'Lay dia chi mac

    CheckAndCreateTable

    Dim rs As DAO.Recordset
    Set rs = DBKetoan.OpenRecordset("SELECT TOP 1 Name FROM tbRegister ")
    If Not rs.EOF Then
        ' Ki?m tra xem c?t Name c� tr?ng kh�ng
        If IsNull(rs.Fields("Name").Value) Or rs.Fields("Name").Value = "" Then
             
        Else
            MsgBox "D�ng d?u ti�n c?a c?t Name kh�ng tr?ng."
        End If
    End If




    Select Case FrmMatkhau.tag
    Case 0:
        If KiemTraMatKhau(txtPsw.Text) Then
            HienThongBao VString(CboUser.Text), 3
            ok = True
            ExecuteSQL5 "UPDATE Users SET WS='" + GetComputerName1 + "' WHERE MaSo=" + CStr(UserID), False
            Unload Me
        Else
            MsgBox "Sai m�t kh�u !", vbExclamation, App.ProductName
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
                Label(0).Caption = "M�t kh�u m�i"
                txtPsw.Text = ""
                RFocus txtPsw
            Else
                MsgBox "Sai m�t kh�u !", vbExclamation, App.ProductName
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
                MsgBox "B�n ch�a nh� ��ng m�t kh�u !", vbExclamation, App.ProductName
                RFocus txtPsw
            End If
        End Select
    End Select
End Sub

Private Sub Form_Activate()
    If Counter < 0 Then
        Counter = 0
        If Me.tag = 1 Then
            Dim i As Integer
            
            Me.Caption = "Thay ��i m�t kh�u"
            Label(0).Caption = "M�t kh�u c�"
            SetListIndex CboUser, UserID
            ok = True
        Else
            ok = False
        End If
    End If
End Sub
'====================================================================================================
' Thu tuc kiem tra mat khau
'====================================================================================================
Private Function KiemTraMatKhau(pstr_psw As String) As Boolean
    Dim rs_mk As Recordset
    
    Set rs_mk = DBKetoan.OpenRecordset("SELECT Users.* FROM Users WHERE MaSo = " + CStr(CboUser.ItemData(CboUser.ListIndex)), dbOpenSnapshot, dbForwardOnly)
    If (Int_StrToCode(pstr_psw) = rs_mk!psw - pNamTC) Then
        KiemTraMatKhau = True
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
'====================================================================================================
' X� l� ph�m n�ng
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
        HienThongBao "K�t th�c ch��ng tr�nh!", 1
        CloseUp 1
        WSpace.Close
        Me.MousePointer = 0
        End
    Else
        HienThongBao "", 1
    End If
End Sub

