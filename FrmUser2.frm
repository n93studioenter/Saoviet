VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmUser2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch truy nhËp chøng tõ ghi sæ"
   ClientHeight    =   3750
   ClientLeft      =   2670
   ClientTop       =   3240
   ClientWidth     =   5655
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
   Icon            =   "FrmUser2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "User Rights"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3750
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin MSGrid.Grid GrdNT 
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Tag             =   "20"
      Top             =   360
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   5741
      _StockProps     =   77
      Rows            =   20
      Cols            =   3
      FixedRows       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   3
      Left            =   4440
      Picture         =   "FrmUser2.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   1
      Left            =   4440
      Picture         =   "FrmUser2.frx":6C04
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NhËp sè liÖu"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Tag             =   "Rights"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chøng tõ ghi sæ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "Book"
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "FrmUser2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ms As Long
'====================================================================================================
' Thªm, Ghi, Xãa nguyªn tÖ
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 1:
            ExecuteSQL5 "DELETE * FROM User2 WHERE User=" + CStr(ms)
            With GrdNT
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 0
                    If Len(.Text) = 0 Then Exit Sub
                    .col = 1
                    If .Text = "X" Then
                        .col = 2
                        ExecuteSQL5 "INSERT INTO User2 (MaSo,User,CTGS) VALUES (" + CStr(Lng_MaxValue("MaSo", "User2") + 1) + "," + CStr(ms) + "," + .Text + ")"
                    End If
                Next
            End With
        Case 3:
            Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    If Me.tag > 0 Then
        ms = Me.tag
        LietKeNgte
        Me.tag = 0
    End If
End Sub

'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyG:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyV:
                RFocus Command(3)
                Command_Click 3
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
'====================================================================================================
' Khëi t¹o cöa sæ
'====================================================================================================
Private Sub Form_Load()
    ColumnSetUp GrdNT, 0, 2860, 0
    ColumnSetUp GrdNT, 1, 1060, 2
    ColumnSetUp GrdNT, 2, 1, 0
    
    SetFont Me
End Sub

Private Sub GrdNt_click()
    With GrdNT
        .col = 1
        If Len(.Text) = 0 Then .Text = "X" Else .Text = ""
    End With
End Sub

Private Sub GrdNt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then GrdNt_click
End Sub

'====================================================================================================
' HiÓn thÞ danh s¸ch nguyªn tÖ
'====================================================================================================
Private Sub LietKeNgte()
    Dim rs_ngte As Recordset, s As String
    
    SetSQL "MienTru", "SELECT * FROM User2 WHERE User=" + CStr(ms)
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT SoHieu+ ' - '+DienGiai AS CTGS, CTGhiSo.MaSo,MienTru.User FROM CTGhiSo LEFT JOIN MienTru ON CTGhiSo.MaSo=MienTru.CTGS WHERE CTGhiSo.MaSo>1 ORDER BY SoHieu DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        If IsNull(rs_ngte!USER) Then s = "" Else s = "X"
        GrdNT.AddItem rs_ngte!CTGS + Chr(9) + s + Chr(9) + CStr(rs_ngte!MaSo), 0
        rs_ngte.MoveNext
    Loop
    If rs_ngte.RecordCount > 0 Then GrdNT.Rows = rs_ngte.RecordCount
    rs_ngte.Close
    Set rs_ngte = Nothing
End Sub

