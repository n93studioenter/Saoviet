VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmCTGS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "§¨ng ký chøng tõ ghi sæ"
   ClientHeight    =   3750
   ClientLeft      =   690
   ClientTop       =   1740
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
   Icon            =   "FrmCTGS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3750
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Voucher Book"
   Begin MSGrid.Grid GrdNT 
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Tag             =   "20"
      Top             =   360
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   5318
      _StockProps     =   77
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   20
      FixedRows       =   0
      ScrollBars      =   2
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   4440
      Picture         =   "FrmCTGS.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&Return"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   4440
      Picture         =   "FrmCTGS.frx":6C04
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "&Delete"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   4440
      Picture         =   "FrmCTGS.frx":80E6
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "&Save"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   4440
      Picture         =   "FrmCTGS.frx":9514
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "&Add"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   80
      TabIndex        =   5
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "..."
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DiÔn gi¶i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Tag             =   "Description"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè hiÖu"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Tag             =   "Code"
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FrmCTGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer
'====================================================================================================
' Thªm, Ghi, Xãa nguyªn tÖ
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0:
            ThemMoi = 1
            For i = 0 To 1
                txtNhap(i).Text = "..."
            Next
            RFocus txtNhap(0)
        Case 1:
            Select Case ThemMoi
                Case 0:
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then Exit Sub
                    If ExecuteSQL5("UPDATE CTGhiSo SET SoHieu='" + txtNhap(0).Text + "',DienGiai='" + txtNhap(1).Text + _
                        "' WHERE SoHieu='" + GrdNT.Text + "'") <> 0 Then Exit Sub
                    For i = 0 To 1
                        GrdNT.col = i
                        GrdNT.Text = txtNhap(i).Text
                    Next
                Case 1:
                    If ExecuteSQL5("INSERT INTO CTGhiSo (MaSo,SoHieu,DienGiai) VALUES (" + CStr(Lng_MaxValue("MaSo", "CTGhiSo") + 1) + ",'" + txtNhap(0).Text + "','" _
                        + txtNhap(1).Text + "')") <> 0 Then Exit Sub
                    GrdNT.AddItem txtNhap(0).Text + Chr(9) + txtNhap(1).Text, InsertGridRow(GrdNT, 0, txtNhap(0).Text)
                    ThemMoi = 0
                    GrdNT.Row = GrdNT.Rows - 1
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then GrdNT.RemoveItem GrdNT.Row
                    GrdNT.Row = 0
            End Select
        Case 2:
            GrdNT.col = 0
            If Len(GrdNT.Text) = 0 Then Exit Sub
            If ExecuteSQL5("DELETE FROM CTGhiSo WHERE SoHieu='" + GrdNT.Text + "'") <> 0 Then Exit Sub
            GrdNT.RemoveItem GrdNT.Row
            If GrdNT.Rows < GrdNT.tag Then GrdNT.Rows = GrdNT.tag
        Case 3:
            Unload Me
    End Select
End Sub

'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyT:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyG:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyX:
                RFocus Command(2)
                Command_Click 2
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
    ColumnSetUp GrdNT, 0, 1300, 2
    ColumnSetUp GrdNT, 1, 2620, 0
    
    Caption = Caption + " - " + CStr(pNamTC)
    LietKeNgte
    
    SetFont Me
End Sub

Private Sub GrdNt_click()
    Dim i As Integer
    
    SendKeys "{Home}", True
    SetGridIndex GrdNT, GrdNT.Row
    With GrdNT
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        For i = 0 To 1
            .col = i
            txtNhap(i).Text = .Text
        Next
        ThemMoi = 0
    End With
End Sub

Private Sub GrdNt_KeyPress(KeyAscii As Integer)
    SendKeys "{Home}", True
    SetGridIndex GrdNT, GrdNT.Row
    
    If KeyAscii = 13 Then GrdNt_click
End Sub

Private Sub GrdNT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdNT, GrdNT.col
    End If
End Sub

Private Sub txtNhap_GotFocus(Index As Integer)
    AutoSelect txtNhap(Index)
End Sub
'====================================================================================================
' HiÓn thÞ danh s¸ch nguyªn tÖ
'====================================================================================================
Private Sub LietKeNgte()
    Dim rs_ngte As Recordset
    
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT * FROM CTGhiSo WHERE MaSo>1 ORDER BY SoHieu DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        GrdNT.AddItem rs_ngte!sohieu + Chr(9) + rs_ngte!diengiai, 0
        rs_ngte.MoveNext
    Loop
    GrdNT.Rows = IIf(rs_ngte.RecordCount > GrdNT.tag, rs_ngte.RecordCount, GrdNT.tag)
    GrdNT.Row = 0
    GrdNT.col = 0
    rs_ngte.Close
    Set rs_ngte = Nothing
    GrdNt_click
End Sub

Private Sub txtNhap_LostFocus(Index As Integer)
    Select Case Index
        Case 0, 1:
            If Len(txtNhap(Index).Text) = 0 Then txtNhap(Index).Text = "..."
    End Select
End Sub

