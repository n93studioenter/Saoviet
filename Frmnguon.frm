VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmNguon 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch nguyån nhËp xuÊt vËt t­"
   ClientHeight    =   5535
   ClientLeft      =   1905
   ClientTop       =   930
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
   Icon            =   "Frmnguon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1"
   Begin MSGrid.Grid GrdNT 
      Height          =   4815
      Left            =   120
      TabIndex        =   8
      Tag             =   "20"
      Top             =   360
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   8493
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
      Left            =   4440
      Picture         =   "Frmnguon.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Return"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   2
      Left            =   4440
      Picture         =   "Frmnguon.frx":6C04
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Delete"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   4440
      Picture         =   "Frmnguon.frx":80E6
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "&Save"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   4440
      Picture         =   "Frmnguon.frx":9514
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&Add"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   960
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "..."
      Top             =   5160
      Width           =   3135
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Text            =   "..."
      Top             =   5160
      Width           =   855
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
      Left            =   960
      TabIndex        =   7
      Tag             =   "Description"
      Top             =   120
      Width           =   3135
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
      TabIndex        =   6
      Tag             =   "Code"
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmNguon"
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
            RFocus txtNhap(0)
        Case 1:
            Me.MousePointer = 11
            Select Case ThemMoi
                Case 0:
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then GoTo XongNguon
                    If ExecuteSQL5("UPDATE NguonNhapXuat SET SoHieu= '" + txtNhap(0).Text + "', DienGiai='" _
                        + txtNhap(1).Text + "' WHERE SoHieu='" + GrdNT.Text + "'") <> 0 Then GoTo XongNguon
                    GrdNT.Text = txtNhap(0).Text
                    GrdNT.col = 1
                    GrdNT.Text = txtNhap(1).Text
                    GrdNT.col = 0
                Case 1:
                    If ExecuteSQL5("INSERT INTO NguonNhapXuat (MaSo,SoHieu,DienGiai) VALUES (" + CStr(Lng_MaxValue("MaSo", "NguonNhapXuat") + 1) + ",'" + txtNhap(0).Text + "','" + txtNhap(1).Text + "')") <> 0 Then GoTo XongNguon
                    GrdNT.AddItem txtNhap(0).Text + Chr(9) + txtNhap(1).Text, InsertGridRow(GrdNT, 0, txtNhap(0).Text)
                    ThemMoi = 0
                    GrdNT.Row = GrdNT.Rows - 1
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then GrdNT.RemoveItem GrdNT.Row
                    GrdNT.Row = 0
            End Select
        Case 2:
            Dim sql As String
        
            GrdNT.col = 0
            If Len(GrdNT.Text) = 0 Then Exit Sub
            
            Me.MousePointer = 11
                  
            sql = "SELECT DISTINCTROW Count(Chungtu.MaSo) AS F1 FROM Chungtu INNER JOIN NguonNhapXuat ON Chungtu.MaNguon = NguonNhapXuat.MaSo" _
                & " WHERE (Chungtu.MaLoai = 1 Or Chungtu.MaLoai=2) AND Chungtu.MaVattu>0 And NguonNhapXuat.SoHieu = '" + GrdNT.Text + "'"
            If SelectSQL(sql) = 0 Then
                If ExecuteSQL5("DELETE FROM NguonNhapXuat WHERE Sohieu='" + GrdNT.Text + "'") <> 0 Then GoTo XongNguon
                GrdNT.RemoveItem GrdNT.Row
                If GrdNT.Rows <= 20 Then GrdNT.Rows = 20
            Else
                ErrMsg er_CoPS
            End If
        
        Case 3:
            Unload Me
    End Select
XongNguon:
    Me.MousePointer = 0
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
    
    ColumnSetUp GrdNT, 0, 820, 2
    ColumnSetUp GrdNT, 1, 3340, 0
                
    ThemMoi = 0
    Caption = Caption + " - " + CStr(pNamTC)
    LietKeNguon
    
    SetFont Me
End Sub

Private Sub GrdNt_click()
    With GrdNT
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        txtNhap(0).Text = .Text
        .col = 1
        txtNhap(1).Text = .Text
        .col = 0
    End With
End Sub

Private Sub GrdNT_DblClick()
    RFocus txtNhap(0)
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
' HiÓn thÞ danh s¸ch nguån nhËp xuÊt
'====================================================================================================
Private Sub LietKeNguon()
    Dim rs_nguon As Recordset
    
    Set rs_nguon = DBKetoan.OpenRecordset("SELECT SoHieu, DienGiai FROM NguonNhapXuat ORDER BY SoHieu DESC", dbOpenSnapshot)
    Do While Not rs_nguon.EOF
        GrdNT.AddItem rs_nguon!sohieu + Chr(9) + rs_nguon!diengiai, 0
        rs_nguon.MoveNext
    Loop
    GrdNT.Rows = IIf(rs_nguon.RecordCount > GrdNT.tag, rs_nguon.RecordCount, GrdNT.tag)
    GrdNT.Row = 0
    GrdNt_click
    rs_nguon.Close
    Set rs_nguon = Nothing
End Sub

Private Sub txtNhap_LostFocus(Index As Integer)
    If Len(txtNhap(Index).Text) = 0 Then txtNhap(Index).Text = "..."
End Sub
