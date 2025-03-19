VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmNguyente 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch nguyªn tÖ vµ tû gi¸"
   ClientHeight    =   3135
   ClientLeft      =   4845
   ClientTop       =   3240
   ClientWidth     =   3930
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
   Icon            =   "Frmngte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   Tag             =   "Exchange Rate List"
   Begin MSGrid.Grid GrdNT 
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Tag             =   "20"
      Top             =   360
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   4260
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
      FixedRows       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   3
      Left            =   2760
      Picture         =   "Frmngte.frx":57E2
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
      Left            =   2760
      Picture         =   "Frmngte.frx":6C04
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
      Left            =   2760
      Picture         =   "Frmngte.frx":80E6
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
      Left            =   2760
      Picture         =   "Frmngte.frx":9514
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Add"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Frmngte.frx":AA6E
      Top             =   2760
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
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tû gi¸"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Tag             =   "Ex. Rate"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nguyªn tÖ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Tag             =   "Currency"
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmNguyente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer
Dim SoNT As Integer

'====================================================================================================
' Thªm, Ghi, Xãa nguyªn tÖ
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim tygia As Double
    
    tygia = RoundMoney(Cdbl5(txtNhap(1).Text))
    Select Case Index
        Case 0:
            ThemMoi = 1
            txtNhap(0).Text = ""
            RFocus txtNhap(0)
        Case 1:
            If Len(txtNhap(0).Text) = 0 Then
                RFocus txtNhap(0)
                Exit Sub
            End If
            Select Case ThemMoi
                Case 0:
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then Exit Sub
                    If ExecuteSQL5("UPDATE Nguyente SET Kyhieu='" + txtNhap(0).Text + "',TyGia=" _
                        + DoiDau(tygia) + " WHERE KyHieu='" + GrdNT.Text + "'") <> 0 Then Exit Sub
                    GrdNT.Text = txtNhap(0).Text
                    GrdNT.col = 1
                    GrdNT.Text = txtNhap(1).Text
                Case 1:
                    If ExecuteSQL5("INSERT INTO Nguyente (MaSo,KyHieu,TyGia) VALUES (" + CStr(Lng_MaxValue("MaSo", "NguyenTe") + 1) + ",'" + txtNhap(0).Text + "'," _
                    + DoiDau(tygia) + ")") <> 0 Then Exit Sub
                    GrdNT.AddItem txtNhap(0).Text + Chr(9) + txtNhap(1).Text, InsertGridRow(GrdNT, 0, txtNhap(0).Text)
                    ThemMoi = 0
                    SoNT = SoNT + 1
                    Command(0).Enabled = SoNT < 12
                    GrdNT.Row = GrdNT.Rows - 1
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then GrdNT.RemoveItem GrdNT.Row
                    GrdNT.Row = 0
            End Select
        Case 2:
            Dim sql As String
            
            GrdNT.col = 0
            If Len(GrdNT.Text) = 0 Then Exit Sub
            sql = "SELECT DISTINCTROW Count(HethongTK.MaNT) AS F1 FROM HethongTK INNER JOIN NguyenTe ON" _
                        & " HethongTK.MaNT = NguyenTe.MaSo WHERE NguyenTe.KyHieu = '" + GrdNT.Text + "'"
            If SelectSQL(sql) = 0 Then
                If ExecuteSQL5("DELETE FROM Nguyente WHERE Kyhieu='" + GrdNT.Text + "'") <> 0 Then Exit Sub
                GrdNT.RemoveItem GrdNT.Row
                If GrdNT.Rows <= GrdNT.tag Then GrdNT.Rows = GrdNT.tag
                SoNT = SoNT - 1
                Command(0).Enabled = SoNT < 12
            Else
                ErrMsg er_CoPS
            End If
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
    ColumnSetUp GrdNT, 0, 940, 2
    ColumnSetUp GrdNT, 1, 1300, 1
    Caption = Caption + " - " + CStr(pNamTC)
    LietKeNgte
    
    SetFont Me
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    CloseItemList
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
        .col = 0
        ThemMoi = 0
        RFocus txtNhap(1)
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
    
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT Kyhieu, TyGia FROM Nguyente WHERE KyHieu<>'" + pTienStr + "' ORDER BY Kyhieu DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        GrdNT.AddItem rs_ngte!KyHieu + Chr(9) + Format(rs_ngte!tygia, Mask_0), 0
        rs_ngte.MoveNext
    Loop
    SoNT = rs_ngte.RecordCount
    Command(0).Enabled = SoNT < 12
    GrdNT.Rows = IIf(SoNT > GrdNT.tag, SoNT, GrdNT.tag)
    GrdNT.Row = 0
    GrdNT.col = 0
    rs_ngte.Close
    Set rs_ngte = Nothing
    GrdNt_click
End Sub

Private Sub txtNhap_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0:
            If KeyAscii = 32 Or KeyAscii = 35 Or KeyAscii = 39 Or KeyAscii = 42 Then KeyAscii = 0
        Case 1:
            KeyProcess txtNhap(Index), KeyAscii
    End Select
End Sub

Private Sub txtNhap_LostFocus(Index As Integer)
    If Index = 1 Then
        txtNhap(1).Text = Format(txtNhap(1).Text, Mask_0)
    End If
End Sub
