VERSION 5.00
Begin VB.Form FrmMK 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UA Code"
   ClientHeight    =   2190
   ClientLeft      =   4065
   ClientTop       =   3405
   ClientWidth     =   3855
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
   Icon            =   "FrmMK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   3
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Version 2005"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&KÕt thóc"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Version 2004"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   2
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Tax code"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Application Code"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmMK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pcname As String
Dim WSpace As Workspace
Dim db As Database

Private Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As Long, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, ByVal lpMaximumComponentLength As Long, ByVal lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As Long, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function VolumeSerial(DriveLetter) As Long
    Dim Serial As Long
    Call GetVolumeSerialNumber(UCase(DriveLetter) & ":\", 0&, 0&, Serial, 0&, 0&, 0&, 0&)
    VolumeSerial = Serial
End Function

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0:
            Text1(0).Text = UCase(Text1(0).Text)
            Text1(2).Text = CStr(Int_StrToCode(Text1(0).Text) + Int_StrToCode(Text1(1).Text))
            RFocus Text1(2)
        Case 1:
            Text1(0).Text = UCase(Text1(0).Text)
            Text1(3).Text = CStr(Int_StrToCode2(Text1(0).Text) + Int_StrToCode2(Text1(1).Text))
            RFocus Text1(3)
        Case 2:
            Unload Me
            End
    End Select
End Sub

Private Function Int_StrToCode(Str As String) As Long
    Dim i As Long, Length As Integer, K As Long, kq As Long
    
    Length = Len(Str)
    If Length > 0 Then
        If Length > 12 Then
            For i = Length To 1 Step -1
                K = Asc(Right(Str, i))
                kq = kq + 2 * i * (K ^ 2)
            Next
        Else
            For i = Length To 1 Step -1
                K = Asc(Right(Str, i))
                kq = kq + 8 * i * (K ^ 3)
            Next
        End If
    End If
    Int_StrToCode = kq
End Function

Private Function Int_StrToCode2(Str As String) As Long
    Dim i As Long, Length As Integer, K As Long, kq As Long
    
    Length = Len(Str)
    If Length > 0 Then
        If Length > 12 Then
            For i = Length To 1 Step -1
                K = Asc(Right(Str, i))
                kq = 1 + kq + 2 * i * (K ^ 2)
            Next
        Else
            For i = Length To 1 Step -1
                K = Asc(Right(Str, i))
                kq = 1 + kq + 8 * i * (K ^ 3)
            Next
        End If
    End If
    Int_StrToCode2 = kq
End Function

Private Sub Form_Load()
    Dim st As String, c As String, sr As String, s As Long
    
    c = CurrentDrive
    sr = VolumeSerial(c)
    
    s = GetSetting("UA_0", "0", c, 0)
    If s <> sr Then
        st = FPsw.GetPswX()
        If st <> CStr(Year(Date)) + CStr(Month(Date)) + CStr(Day(Date)) Then End
        SaveSetting "UA_0", "0", c, sr
    End If
    
    st = FPsw.GetPswX()
    pcname = GetComputerName1
    If UCase(st) <> UCase(pcname) Then End
    
    Set WSpace = DBEngine.CreateWorkspace(CStr(Time), "Admin", "", dbUseJet)
End Sub

Private Function CurrentDrive() As String
    Dim retValue As Long, Buffer As String * 255
    
    retValue = GetWindowsDirectory(Buffer, 255)
    CurrentDrive = Left(Buffer, 1)
End Function

Private Function GetComputerName1() As String
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

Private Function CLng5(st As String) As Long
    If IsNumeric(st) Then CLng5 = CLng(st) Else CLng5 = 0
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set db = Nothing
    WSpace.Close
    Set WSpace = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    AutoSelect Text1(Index)
End Sub

Private Sub AutoSelect(text_box As Object)
      text_box.SelStart = 0
      text_box.SelLength = Len(text_box.Text)
End Sub

Private Sub RFocus(obj As Object)
    On Error Resume Next
    obj.SetFocus
    On Error GoTo 0
End Sub

Private Function SelectSQL(sql As String, Optional SoPhu As Variant = 0) As Variant
    Dim rs As Recordset
    
    On Error Resume Next
    SoPhu = 0
      Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
      If rs.RecordCount > 0 Then
            If IsNull(rs!F1) Then
                SelectSQL = 0
            Else
                SelectSQL = rs!F1
                SoPhu = rs!F2
            End If
      Else
            SelectSQL = 0
      End If
      rs.Close
      Set rs = Nothing
      On Error GoTo 0
End Function

Private Sub GetSerial2005()
    Dim scode As Long, i As Integer, r As String
    
    On Error GoTo Er
    Set db = WSpace.OpenDatabase("\\SERVER\Common\Serial.MDB", False, False)
    
    db.Execute "INSERT INTO List (SDate,STime,PCName,AppCode,TaxCode) VALUES (#" + Format(Date, "mm/dd/yy") + "#,#" + Format(Time, "hh:mm:ss") + "#,'" + pcname + "','" + Text1(0).Text + "','" + Text1(1).Text + "')"
    If db.RecordsAffected > 0 Then
        scode = SelectSQL("SELECT Max(SCode) AS F1 FROM List")
        DoEvents
        Do While SelectSQL("SELECT Flag AS F1,Seri2005 AS F2 FROM List WHERE SCode=" + CStr(scode), r) = 0
            For i = 1 To 30
                DoEvents
            Next
        Loop
        Text1(3).Text = r
        db.Execute "DELETE * FROM List WHERE SCode=" + CStr(scode)
    End If
    db.Close
    On Error GoTo 0
Er:
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub
