VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FU1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch ng­êi sö dông"
   ClientHeight    =   2910
   ClientLeft      =   1650
   ClientTop       =   3015
   ClientWidth     =   3390
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
   Icon            =   "FU1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "User Rights"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2910
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin MSGrid.Grid GrdNT 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Tag             =   "10"
      Top             =   360
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   4260
      _StockProps     =   77
      BackColor       =   16777215
      Rows            =   20
      Cols            =   3
      FixedRows       =   0
      ScrollBars      =   2
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QuyÒn"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Tag             =   "Rights"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ng­êi sö dông"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Tag             =   "User Name"
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "FU1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim M_TK As Long
Dim M_BC As Integer
Dim f1 As Integer
'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
'====================================================================================================
' Khëi t¹o cöa sæ
'====================================================================================================
Private Sub Form_Load()
    ColumnSetUp GrdNT, 0, 1780, 0
    ColumnSetUp GrdNT, 1, 1060, 2
    ColumnSetUp GrdNT, 2, 1, 0
    SetFont Me
End Sub

Public Sub QuyenBCChitiet(bc As Integer, dg As String)
    Dim rs_ngte As Recordset, st As String
    
    M_BC = bc
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT Users.* FROM Users WHERE UserRight>0 AND UserRight<>1 ORDER BY TenNSD DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        If Not TruongDaCo("Users", "C" + CStr(bc)) Then
            st = ""
        Else
            st = IIf(SelectSQL("SELECT C" + CStr(bc) + " AS F1 FROM Users WHERE MaSo=" + CStr(rs_ngte!MaSo)) > 0, "x", "")
        End If
        GrdNT.AddItem rs_ngte!TenNSD + Chr(9) + st + Chr(9) + CStr(rs_ngte!MaSo), 0
        rs_ngte.MoveNext
    Loop
    GrdNT.Rows = IIf(rs_ngte.RecordCount > GrdNT.tag, rs_ngte.RecordCount, GrdNT.tag)
    GrdNT.Row = 0
    rs_ngte.Close
    Set rs_ngte = Nothing
    Me.Caption = dg
    f1 = 2
    Me.Show 1
End Sub

Public Sub QuyenBCTongHop(bc As Integer, dg As String)
    Dim rs_ngte As Recordset, st As String
    
    M_BC = bc
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT Users.* FROM Users WHERE UserRight>0 AND UserRight<>1 ORDER BY TenNSD DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        If Not TruongDaCo("Users", "T" + CStr(bc)) Then
            st = ""
        Else
            st = IIf(SelectSQL("SELECT T" + CStr(bc) + " AS F1 FROM Users WHERE MaSo=" + CStr(rs_ngte!MaSo)) > 0, "x", "")
        End If
        GrdNT.AddItem rs_ngte!TenNSD + Chr(9) + st + Chr(9) + CStr(rs_ngte!MaSo), 0
        rs_ngte.MoveNext
    Loop
    GrdNT.Rows = IIf(rs_ngte.RecordCount > GrdNT.tag, rs_ngte.RecordCount, GrdNT.tag)
    GrdNT.Row = 0
    rs_ngte.Close
    Set rs_ngte = Nothing
    Me.Caption = dg
    f1 = 3
    Me.Show 1
End Sub

Private Sub GrdNT_DblClick()
    Dim m As Long
    With GrdNT
        .col = 2
        m = CLng5(.Text)
        If m = 0 Then Exit Sub
        .col = 1
        Select Case f1
            Case 1:
                If .Text = "" Then
                    .Text = "x"
                    If Not TruongDaCo("HethongTK", "U" + CStr(m)) Then ThemTruong "HethongTK", "U" + CStr(m), dbInteger
                    ExecuteSQL5 "UPDATE HethongTK SET U" + CStr(m) + "=1 WHERE MaSo=" + CStr(M_TK)
                Else
                    .Text = ""
                    ExecuteSQL5 "UPDATE HethongTK SET U" + CStr(m) + "=0 WHERE MaSo=" + CStr(m)
                End If
            Case 2:
                If Len(.Text) = 0 Then
                    .Text = "x"
                    If Not TruongDaCo("Users", "C" + CStr(M_BC)) Then ThemTruong "Users", "C" + CStr(M_BC), dbInteger
                    ExecuteSQL5 "UPDATE Users SET C" + CStr(M_BC) + "=1 WHERE MaSo=" + CStr(m)
                Else
                    .Text = ""
                    ExecuteSQL5 "UPDATE Users SET C" + CStr(M_BC) + "=0 WHERE MaSo=" + CStr(m)
                End If
            Case 3:
                If Len(.Text) = 0 Then
                    .Text = "x"
                    If Not TruongDaCo("Users", "T" + CStr(M_BC)) Then ThemTruong "Users", "T" + CStr(M_BC), dbInteger
                    ExecuteSQL5 "UPDATE Users SET T" + CStr(M_BC) + "=1 WHERE MaSo=" + CStr(m)
                Else
                    .Text = ""
                    ExecuteSQL5 "UPDATE Users SET T" + CStr(M_BC) + "=0 WHERE MaSo=" + CStr(m)
                End If
        End Select
    End With
End Sub

Public Sub QuyenTaiKhoan(mtk As Long)
    Dim rs_ngte As Recordset, st As String
    M_TK = mtk
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT Users.* FROM Users WHERE UserRight>0 AND UserRight<>2 ORDER BY TenNSD DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        If Not TruongDaCo("HethongTK", "U" + CStr(rs_ngte!MaSo)) Then
            st = ""
        Else
            st = IIf(SelectSQL("SELECT U" + CStr(rs_ngte!MaSo) + " AS F1 FROM HethongTK WHERE MaSo=" + CStr(mtk)) > 0, "x", "")
        End If
        GrdNT.AddItem rs_ngte!TenNSD + Chr(9) + st + Chr(9) + CStr(rs_ngte!MaSo), 0
        rs_ngte.MoveNext
    Loop
    GrdNT.Rows = IIf(rs_ngte.RecordCount > GrdNT.tag, rs_ngte.RecordCount, GrdNT.tag)
    GrdNT.Row = 0
    rs_ngte.Close
    Set rs_ngte = Nothing
    Me.Caption = "TK: " + MaSo2SoHieu(mtk, "HethongTK")
    f1 = 1
    Me.Show 1
End Sub

