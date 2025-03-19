VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmKC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch chøng tõ kÕt chuyÓn"
   ClientHeight    =   7080
   ClientLeft      =   750
   ClientTop       =   930
   ClientWidth     =   9855
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
   Icon            =   "FrmKC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Conversion Voucher List"
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   4
      Left            =   8640
      Picture         =   "FrmKC.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&Detail"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   6960
      MaxLength       =   80
      TabIndex        =   8
      Tag             =   "0"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   5640
      MaxLength       =   80
      TabIndex        =   7
      Tag             =   "0"
      Top             =   6720
      Width           =   1335
   End
   Begin MSGrid.Grid GrdNT 
      Height          =   6375
      Left            =   120
      TabIndex        =   11
      Tag             =   "30"
      Top             =   360
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
      _ExtentY        =   11245
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
      Cols            =   5
      FixedRows       =   0
      ScrollBars      =   2
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   8640
      Picture         =   "FrmKC.frx":6C44
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "&Return"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   8640
      Picture         =   "FrmKC.frx":8066
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Delete"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   8640
      Picture         =   "FrmKC.frx":9548
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "&Save"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   8640
      Picture         =   "FrmKC.frx":A976
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
      Left            =   600
      MaxLength       =   80
      TabIndex        =   6
      Top             =   6720
      Width           =   5055
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "0"
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§Õn tµi kho¶n"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6960
      TabIndex        =   13
      Tag             =   "To Account"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tõ tµi kho¶n"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   12
      Tag             =   "From Account"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DiÔn gi¶i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   10
      Tag             =   "Desciption"
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Tag             =   "Ord."
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FrmKC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer
Dim oTaikhoan As New ClsTaikhoan
'====================================================================================================
' Thªm, Ghi, Xãa nguyªn tÖ
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0:
            ThemMoi = 1
            txtNhap(0).Text = "0"
            For i = 1 To 3
                txtNhap(i).Text = "..."
            Next
            RFocus txtNhap(0)
        Case 1:
            For i = 2 To 3
                If txtNhap(i).tag = 0 Then
                    RFocus txtNhap(i)
                    Exit Sub
                End If
            Next
            Select Case ThemMoi
                Case 0:
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then Exit Sub
                    If ExecuteSQL5("UPDATE CTKetChuyen SET STT=" + txtNhap(0).Text + ",DienGiai='" + txtNhap(1).Text + _
                        "',TK1='" + txtNhap(2).Text + "',TK2='" + txtNhap(3).Text + "' WHERE STT=" + GrdNT.Text) <> 0 Then Exit Sub
                    For i = 0 To 3
                        GrdNT.col = i
                        GrdNT.Text = txtNhap(i).Text
                    Next
                Case 1:
                    If ExecuteSQL5("INSERT INTO CTKetChuyen (MaSo,STT,DienGiai,TK1,TK2) VALUES (" + CStr(Lng_MaxValue("MaSo", "CTKetChuyen") + 1) + "," + txtNhap(0).Text + ",'" _
                        + txtNhap(1).Text + "','" + txtNhap(2).Text + "','" + txtNhap(3).Text + "')") <> 0 Then Exit Sub
                    GrdNT.AddItem txtNhap(0).Text + Chr(9) + txtNhap(1).Text + Chr(9) + txtNhap(2).Text + Chr(9) + txtNhap(3).Text + Chr(9) + CStr(Lng_MaxValue("MaSo", "CTKetChuyen")), InsertGridRow(GrdNT, 0, txtNhap(0).Text)
                    ThemMoi = 0
                    GrdNT.Row = GrdNT.Rows - 1
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then GrdNT.RemoveItem GrdNT.Row
                    GrdNT.Row = 0
                    KiemTra
            End Select
        Case 2:
            GrdNT.col = 0
            If Len(GrdNT.Text) = 0 Then Exit Sub
            If ExecuteSQL5("DELETE FROM CTKetChuyen WHERE STT=" + GrdNT.Text) <> 0 Then Exit Sub
            GrdNT.RemoveItem GrdNT.Row
            If GrdNT.Rows < GrdNT.tag Then GrdNT.Rows = GrdNT.tag
        Case 3:
            Unload Me
        Case 4:
            GrdNT.col = 0
            If Len(GrdNT.Text) = 0 Then Exit Sub
            'Load FrmMauKC
            With FrmMauKC
                GrdNT.col = 4
                .tag = CLng5(GrdNT.Text)
                GrdNT.col = 1
                .TK(2).Caption = GrdNT.Text
                GrdNT.col = 2
                oTaikhoan.InitTaikhoanSohieu GrdNT.Text
                .TK(0).Caption = oTaikhoan.sohieu + " - " + oTaikhoan.Ten
                .TK(0).tag = oTaikhoan.MaSo
                GrdNT.col = 3
                oTaikhoan.InitTaikhoanSohieu GrdNT.Text
                .TK(1).Caption = oTaikhoan.sohieu + " - " + oTaikhoan.Ten
                .TK(1).tag = oTaikhoan.MaSo
                .Show 1
            End With
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
    ColumnSetUp GrdNT, 0, 465, 2
    ColumnSetUp GrdNT, 1, 5020, 0
    ColumnSetUp GrdNT, 2, 1300, 0
    ColumnSetUp GrdNT, 3, 1300, 0
    ColumnSetUp GrdNT, 4, 1, 0
    
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
        For i = 0 To 3
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
    
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT * FROM CTKetChuyen ORDER BY STT DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        GrdNT.AddItem IIf(rs_ngte!stt < 10, "0", "") + CStr(rs_ngte!stt) + Chr(9) + rs_ngte!diengiai + Chr(9) + rs_ngte!TK1 + Chr(9) + rs_ngte!tk2 + Chr(9) + CStr(rs_ngte!MaSo), 0
        rs_ngte.MoveNext
    Loop
    GrdNT.Rows = IIf(rs_ngte.RecordCount > GrdNT.tag, rs_ngte.RecordCount, GrdNT.tag)
    GrdNT.Row = 0
    GrdNT.col = 0
    rs_ngte.Close
    Set rs_ngte = Nothing
    KiemTra
    GrdNt_click
End Sub

Private Sub txtNhap_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0:            KeyProcess txtNhap(Index), KeyAscii
        Case 2, 3:
            If KeyAscii = vbKeyReturn Then
                Me.MousePointer = 11
                txtNhap(Index).Text = FrmTaikhoan.ChonTk(txtNhap(Index).Text)
                Me.MousePointer = 0
                txtNhap_LostFocus Index
            End If
    End Select
End Sub

Private Sub txtNhap_LostFocus(Index As Integer)
    Select Case Index
        Case 0:
            txtNhap(0).Text = Format(txtNhap(0).Text, Mask_0)
            If Len(txtNhap(0).Text) < 2 Then txtNhap(0).Text = "0" + txtNhap(0).Text
        Case 1:
            If Len(txtNhap(Index).Text) = 0 Then txtNhap(Index).Text = "..."
        Case 2, 3:
            If Len(txtNhap(Index).Text) = 0 Then
                txtNhap(Index).Text = "..."
            Else
                oTaikhoan.InitTaikhoanSohieu txtNhap(Index).Text
                If oTaikhoan.MaSo > 0 Then
                    txtNhap(Index).tag = IIf(oTaikhoan.MaTC = 0 Or oTaikhoan.MaTC = oTaikhoan.MaSo, CLng5(oTaikhoan.sohieu), 0)
                Else
                    txtNhap(Index).tag = 0
                End If
            End If
    End Select
End Sub

Private Sub KiemTra()
    Command(0).Enabled = SelectSQL("SELECT Count(MaSo) AS F1 FROM CTKetChuyen") < MaxKC
End Sub
