VERSION 5.00
Begin VB.Form FrmMK 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Installation Code"
   ClientHeight    =   660
   ClientLeft      =   4065
   ClientTop       =   3405
   ClientWidth     =   2595
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
   Icon            =   "FrmMKSV.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   660
   ScaleWidth      =   2595
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer STimer 
      Interval        =   3000
      Left            =   600
      Top             =   120
   End
   Begin VB.CommandButton Command 
      Caption         =   "&§ãNG"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   120
      Picture         =   "FrmMKSV.frx":0442
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FrmMK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Wspace As Workspace
Dim db As Database
Dim db1 As Database
Dim rs As Recordset

Dim pcname As String

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub Command_Click()
    Unload Me
End Sub

Private Function Int_StrToCode(Str As String, off As Integer) As Long
    Dim i As Long, Length As Integer, K As Long, kq As Long
    
    Length = Len(Str)
    If Length > 0 Then
        If Length > 12 Then
            For i = Length To 1 Step -1
                K = Asc(Right(Str, i)) + off
                kq = kq + 2 * i * (K ^ 2)
            Next
        Else
            For i = Length To 1 Step -1
                K = Asc(Right(Str, i)) + off
                kq = kq + 8 * i * (K ^ 3)
            Next
        End If
    End If
    Int_StrToCode = kq
End Function

Private Sub Form_Load()
    pcname = GetComputerName1
    If UCase(pcname) <> "SERVER" Then End
    
    Set Wspace = DBEngine.CreateWorkspace(CStr(Time), "Admin", "", dbUseJet)
    Set db = Wspace.OpenDatabase("\\SERVER\Common\Serial.MDB", False, False)
    Set db1 = Wspace.OpenDatabase("\\SERVER\Documents$\VANETData\Serial.MDB", False, False)
    Set rs = db.OpenRecordset("SELECT List.* FROM List INNER JOIN PCName ON List.PCName=PCName.PCName WHERE List.Flag=0", dbOpenSnapshot)
End Sub

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

Private Sub Form_Unload(Cancel As Integer)
    rs.Close
    db.Close
    db1.Close
    Wspace.Close
    Set rs = Nothing
    Set db = Nothing
    Set db1 = Nothing
    Set Wspace = Nothing
End Sub

Private Sub STimer_Timer()
    Dim r As String, off As Integer
    
    rs.Requery
    If Not rs.BOF Then rs.MoveFirst
    Do While Not rs.EOF
        off = CInt5(Left(rs!TaxCode, 2))
        off = IIf(off Mod 2 = 0, 2, 1)
        
        r = CStr(Int_StrToCode(rs!AppCode, off) + Int_StrToCode(rs!TaxCode, off))
        db.Execute "UPDATE List SET Seri2005='" + r + "',Flag=1 WHERE SCode=" + CStr(rs!scode)
        
        db1.Execute "INSERT INTO List (SDate,STime,PCName,AppCode,TaxCode,Seri2005) VALUES (#" + Format(rs!SDate, "mm/dd/yy") + "#,#" + Format(rs!STime, "hh:mm:ss") + "#,'" + rs!pcname + "','" + rs!AppCode + "','" + rs!TaxCode + "','" + r + "')"
        rs.MoveNext
    Loop
End Sub

Private Function CInt5(st As String) As Double
    Dim x As Double
    If IsNumeric(st) Then
        x = CDbl(st)
        If x >= -32768 And x <= 32767 Then
            CInt5 = CInt(x)
        Else
            CInt5 = 0
        End If
    Else
        CInt5 = 0
    End If
End Function
