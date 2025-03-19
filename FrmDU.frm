VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmDU 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sè hiÖu chøng tõ mÆc ®Þnh"
   ClientHeight    =   4455
   ClientLeft      =   2460
   ClientTop       =   3000
   ClientWidth     =   6360
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
   Icon            =   "FrmDU.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Tag             =   "Default Voucher Code"
   Begin MSGrid.Grid GrdNT 
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Tag             =   "20"
      Top             =   360
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   6588
      _StockProps     =   77
      BackColor       =   16777215
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
      Cols            =   3
      FixedRows       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   3
      Left            =   5160
      Picture         =   "FrmDU.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "&Return"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   2
      Left            =   5160
      Picture         =   "FrmDU.frx":6C04
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "&Delete"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   5160
      Picture         =   "FrmDU.frx":80E6
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Save"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   5160
      Picture         =   "FrmDU.frx":9514
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Add"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   720
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "0"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   10
      Tag             =   "1"
      Top             =   4125
      Width           =   3015
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DiÔn gi¶i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   9
      Tag             =   "Description"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tµi kho¶n"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Tag             =   "Account"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè hiÖu"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Tag             =   "V. Code"
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmDU"
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
    Select Case Index
        Case 0:
            ThemMoi = 1
            txtNhap(0).Text = ""
            txtNhap(1).Text = ""
            Label(3).Caption = ""
            RFocus txtNhap(0)
        Case 1:
            If Len(txtNhap(0).Text) = 0 Then
                RFocus txtNhap(0)
                Exit Sub
            End If
            If txtNhap(1).tag = 0 Then
                RFocus txtNhap(1)
                Exit Sub
            End If
            Select Case ThemMoi
                Case 0:
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then Exit Sub
                    If ExecuteSQL5("UPDATE SHChungTu SET SoHieu='" + txtNhap(0).Text + "',SHTK='" _
                        + txtNhap(1).Text + " WHERE SoHieu='" + GrdNT.Text + "'") <> 0 Then Exit Sub
                    GrdNT.Text = txtNhap(0).Text
                    GrdNT.col = 1
                    GrdNT.Text = txtNhap(1).Text
                    GrdNT.col = 2
                    GrdNT.Text = Label(3).Caption
                Case 1:
                    If ExecuteSQL5("INSERT INTO SHChungTu (MaSo,SoHieu,SHTK) VALUES (" + CStr(Lng_MaxValue("MaSo", "SHChungTu") + 1) + ",'" + txtNhap(0).Text + "','" + txtNhap(1).Text + "')") <> 0 Then Exit Sub
                    GrdNT.AddItem txtNhap(0).Text + Chr(9) + txtNhap(1).Text + Chr(9) + Label(3).Caption, InsertGridRow(GrdNT, 0, txtNhap(0).Text)
                    ThemMoi = 0
                    GrdNT.Row = GrdNT.Rows - 1
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then GrdNT.RemoveItem GrdNT.Row
                    GrdNT.Row = 0
            End Select
        Case 2:
            If ExecuteSQL5("DELETE FROM SHChungTu WHERE Sohieu='" + GrdNT.Text + "'") <> 0 Then Exit Sub
            GrdNT.RemoveItem GrdNT.Row
            If GrdNT.Rows <= GrdNT.tag Then GrdNT.Rows = GrdNT.tag
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
    ColumnSetUp GrdNT, 0, 580, 2
    ColumnSetUp GrdNT, 1, 1300, 0
    ColumnSetUp GrdNT, 2, 2740, 0
    txtNhap(0).MaxLength = SHCT_Len
    Caption = Caption + " - " + CStr(pNamTC)
    LietKeNgte
    
    SetFont Me
End Sub

Private Sub GrdNt_click()
    SendKeys "{Home}", True
    SetGridIndex GrdNT, GrdNT.Row
    With GrdNT
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        txtNhap(0).Text = .Text
        .col = 1
        txtNhap(1).Text = .Text
        .col = 2
        Label(3).Caption = .Text
        .col = 0
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
    
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT SHChungTu.SoHieu,SHTK,Ten FROM SHChungTu INNER JOIN HethongTK ON SHChungTu.SHTK=HethongTK.SoHieu ORDER BY SHChungTu.SoHieu DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        GrdNT.AddItem rs_ngte!sohieu + Chr(9) + rs_ngte!shtk + Chr(9) + rs_ngte!Ten, 0
        rs_ngte.MoveNext
    Loop
    GrdNT.Rows = IIf(rs_ngte.RecordCount > GrdNT.tag, rs_ngte.RecordCount, GrdNT.tag)
    GrdNT.Row = 0
    GrdNT.col = 0
    rs_ngte.Close
    Set rs_ngte = Nothing
    GrdNt_click
End Sub

Private Sub txtNhap_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then KeyProcess txtNhap(Index), KeyAscii
End Sub

Private Sub txtNhap_LostFocus(Index As Integer)
Dim mtk As Long
    If Index = 1 Then
        Label(3).Caption = tentk(txtNhap(1).Text, mtk)
        txtNhap(1).tag = mtk
    End If
End Sub
