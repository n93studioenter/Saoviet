VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmHD 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   Caption         =   "Danh s¸ch Hîp ®ång kinh tÕ"
   ClientHeight    =   7080
   ClientLeft      =   75
   ClientTop       =   1065
   ClientWidth     =   10710
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
   Icon            =   "FrmHD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Contract List"
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   8400
      MaxLength       =   20
      TabIndex        =   6
      Tag             =   "0"
      Text            =   "..."
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   7320
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   5
      Tag             =   "0"
      Text            =   "FrmHD.frx":57E2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "0"
      Text            =   "..."
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   4320
      MaxLength       =   20
      TabIndex        =   3
      Tag             =   "0"
      Text            =   "..."
      Top             =   6240
      Width           =   1095
   End
   Begin MSGrid.Grid GrdNT 
      Height          =   5895
      Left            =   120
      TabIndex        =   14
      Tag             =   "30"
      Top             =   360
      Width           =   10455
      _Version        =   65536
      _ExtentX        =   18441
      _ExtentY        =   10398
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
      Rows            =   30
      Cols            =   8
      FixedRows       =   0
      ScrollBars      =   2
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   9240
      Picture         =   "FrmHD.frx":57E4
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   8040
      Picture         =   "FrmHD.frx":6C06
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "&Delete"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   6840
      Picture         =   "FrmHD.frx":80E8
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "&Save"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   5640
      Picture         =   "FrmHD.frx":9516
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "&Add"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "..."
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Text            =   "..."
      Top             =   6240
      Width           =   1335
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   6240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   1
      Left            =   9480
      TabIndex        =   7
      Top             =   6240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ngµy ký"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   9480
      TabIndex        =   20
      Tag             =   "Date"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tªn ®¬n vÞ ký H§"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   19
      Tag             =   "Name"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ghi chó"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   8400
      TabIndex        =   18
      Tag             =   "Notes"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gi¸ trÞ H§"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7320
      TabIndex        =   17
      Tag             =   "Value"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§¬n vÞ ký H§"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   16
      Tag             =   "Company"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Néi dung"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   15
      Tag             =   "Description"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ngµy ký"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   13
      Tag             =   "Date"
      Top             =   120
      Width           =   855
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
      TabIndex        =   12
      Tag             =   "Contract Code"
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "FrmHD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ngay(0 To 1)   As Date
Dim ThemMoi As Integer
Dim okh As New ClsKhachHang
'====================================================================================================
' Thªm, Ghi, Xãa nguyªn tÖ
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0:
            ThemMoi = 1
            For i = 1 To 3
                txtNhap(i).Text = "..."
            Next
            txtNhap(5).Text = "..."
            okh.InitKhachHangMaSo 0
            txtNhap(0).Text = SoHieuVTMoi(0, 4)
            RFocus txtNhap(0)
        Case 1:
            If okh.MaSo = 0 Then
                RFocus txtNhap(2)
                Exit Sub
            End If
            Select Case ThemMoi
                Case 0:
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then Exit Sub
                    If ExecuteSQL5("UPDATE DoituongCT SET SoHieu='" + txtNhap(0).Text + "',DienGiai='" + txtNhap(1).Text + "',NgayKy=#" + Format(ngay(0), Mask_DB) _
                        + "#,MaKhachHang=" + CStr(okh.MaSo) + ",TriGia=" + DoiDau(txtNhap(4).Text) + ",GhiChu='" + txtNhap(5).Text + "',NgayTH=#" + Format(ngay(1), Mask_DB) + "# WHERE SoHieu='" + GrdNT.Text + "'") <> 0 Then Exit Sub
                    GrdNT.Text = txtNhap(0).Text
                    GrdNT.col = 1
                    GrdNT.Text = MedNgay(0).Text
                    GrdNT.col = 7
                    GrdNT.Text = MedNgay(1).Text
                    For i = 2 To 6
                        GrdNT.col = i
                        GrdNT.Text = txtNhap(i - 1).Text
                    Next
                Case 1:
                    If ExecuteSQL5("INSERT INTO DoituongCT (MaSo,SoHieu,DienGiai,NgayKy,MaKhachHang,TriGia,GhiChu,NgayTH) VALUES (" + CStr(Lng_MaxValue("MaSo", "DoituongCT") + 1) + ",'" + txtNhap(0).Text + "','" + txtNhap(1).Text _
                        + "',#" + Format(ngay(0), Mask_DB) + "#," + CStr(okh.MaSo) + "," + DoiDau(Cdbl5(txtNhap(4).Text)) + ",'" + txtNhap(5).Text + "',#" + Format(ngay(1), Mask_DB) + "#)") <> 0 Then Exit Sub
                    GrdNT.AddItem txtNhap(0).Text + Chr(9) + Format(ngay(0), Mask_D) + Chr(9) + txtNhap(1).Text + Chr(9) + txtNhap(2).Text + Chr(9) + txtNhap(3).Text + Chr(9) + txtNhap(4).Text + Chr(9) + txtNhap(5).Text + Chr(9) + Format(ngay(1), Mask_D), InsertGridRow(GrdNT, 0, txtNhap(0).Text)
                    ThemMoi = 0
                    GrdNT.Row = GrdNT.Rows - 1
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then GrdNT.RemoveItem GrdNT.Row
                    GrdNT.Row = 0
            End Select
            Command_Click 0
        Case 2:
            GrdNT.col = 0
            If Len(GrdNT.Text) = 0 Then Exit Sub
            If pNhapKhau > 0 Then
                If SelectSQL("SELECT CPGVHD.MaSo AS F1 FROM CPGVHD INNER JOIN DoiTuongCT ON CPGVHD.MaDT=DoituongCT.MaSo WHERE DoituongCT.SoHieu='" + GrdNT.Text + "'") > 0 Then
                    If MsgBox("§· cã ph¸t sinh, tiÕp tôc ?", vbCritical + vbYesNo, App.ProductName) = vbNo Then Exit Sub
                    ExecuteSQL5 "DELETE CPGVHD.* FROM CPGVHD INNER JOIN DoiTuongCT ON CPGVHD.MaDT=DoituongCT.MaSo WHERE DoituongCT.SoHieu='" + GrdNT.Text + "'"
                End If
            End If
            If ExecuteSQL5("DELETE FROM DoituongCT WHERE SoHieu='" + GrdNT.Text + "'") <> 0 Then Exit Sub
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
    Dim i As Integer
    
    ColumnSetUp GrdNT, 0, 1300, 0
    ColumnSetUp GrdNT, 1, 820, 2
    ColumnSetUp GrdNT, 2, 2020, 0
    ColumnSetUp GrdNT, 3, 1060, 0
    ColumnSetUp GrdNT, 4, 1900, 0
    ColumnSetUp GrdNT, 5, 1060, 1
    ColumnSetUp GrdNT, 6, 1060, 0
    ColumnSetUp GrdNT, 7, 820, 2
    
    For i = 0 To 1
        InitDateVars MedNgay(i), ngay(i)
    Next

    Caption = Caption + " - " + CStr(pNamTC)
    LietKeNgte
    
    SetFont Me
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set okh = Nothing
End Sub

Private Sub GrdNt_click()
    Dim i As Integer
    
    SendKeys "{Home}", True
    SetGridIndex GrdNT, GrdNT.Row
    With GrdNT
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        txtNhap(0).Text = .Text
        For i = 2 To 6
            .col = i
            txtNhap(i - 1).Text = .Text
        Next
        okh.InitKhachHangSohieu txtNhap(2).Text
        .col = 1
        On Error Resume Next
        MedNgay(0).Text = .Text
        .col = 7
        MedNgay(1).Text = .Text
        For i = 0 To 1
            MedNgay_LostFocus i
        Next
        On Error GoTo 0
        ThemMoi = 0
    End With
    RFocus txtNhap(0)
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
    Dim rs_ngte As Recordset, sh As String, Ten As String, nt As Long
    
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT DoituongCT.*,KhachHang.SoHieu AS SHKH,KhachHang.Ten AS TenKH,KhachHang.MaNT AS NT FROM DoituongCT LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo WHERE DoituongCT.MaKhachHang>0 AND DoituongCT.MaSo>1 ORDER BY KhachHang.Ten DESC,DoituongCT.SoHieu DESC,DienGiai DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        If IsNull(rs_ngte!shkh) Then sh = "..." Else sh = rs_ngte!shkh
        If IsNull(rs_ngte!TenKH) Then Ten = "..." Else Ten = rs_ngte!TenKH
        If IsNull(rs_ngte!nt) Then nt = 0 Else nt = rs_ngte!nt
        GrdNT.AddItem rs_ngte!sohieu + Chr(9) + Format(rs_ngte!NgayKy, Mask_D) + Chr(9) + rs_ngte!diengiai + Chr(9) + sh + Chr(9) + Ten + Chr(9) + Format(rs_ngte!TriGia, IIf(nt > 0, Mask_2, Mask_0)) + Chr(9) + rs_ngte!GhiChu + Chr(9) + Format(rs_ngte!NgayTH, Mask_D), 0
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
    Select Case Index
        Case 0:
            If KeyAscii = 32 Or KeyAscii = 35 Or KeyAscii = 39 Or KeyAscii = 42 Then KeyAscii = 0
        Case 4:            KeyProcess txtNhap(Index), KeyAscii
        Case 2:
            If KeyAscii = vbKeyReturn Then
                Me.MousePointer = 11
                txtNhap(Index).Text = FrmKhachHang.ChonKhachHang(txtNhap(Index).Text)
                Me.MousePointer = 0
                txtNhap_LostFocus Index
            End If
    End Select
End Sub

Private Sub txtNhap_LostFocus(Index As Integer)
    Select Case Index
        Case 4:
            If okh.MaNT > 0 Then
                txtNhap(4).Text = Format(txtNhap(4).Text, Mask_2)
            Else
                txtNhap(4).Text = Format(txtNhap(4).Text, Mask_0)
            End If
        Case 1, 3, 5:
            If Len(txtNhap(Index).Text) = 0 Then txtNhap(Index).Text = "..."
        Case 2:
            If Len(txtNhap(Index).Text) = 0 Then
                txtNhap(Index).Text = "..."
                okh.InitKhachHangMaSo 0
            Else
                okh.InitKhachHangSohieu txtNhap(Index).Text
                txtNhap(3).Text = okh.Ten
            End If
    End Select
End Sub

Private Sub MedNgay_GotFocus(Index As Integer)
    AutoSelect MedNgay(Index)
End Sub

Private Sub MedNgay_LostFocus(Index As Integer)
    If IsDate(MedNgay(Index).Text) Then
        ngay(Index) = CDate(MedNgay(Index).Text)
    Else
        RFocus MedNgay(Index)
    End If
End Sub

