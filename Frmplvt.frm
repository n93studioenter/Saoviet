VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmPLVT 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ph©n lo¹i vËt t­, hµng hãa"
   ClientHeight    =   5535
   ClientLeft      =   1755
   ClientTop       =   810
   ClientWidth     =   6255
   ClipControls    =   0   'False
   Icon            =   "Frmplvt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "0"
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   4
      Left            =   5040
      Picture         =   "Frmplvt.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin MSGrid.Grid GrdNT 
      Height          =   4815
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   4815
      _Version        =   65536
      _ExtentX        =   8493
      _ExtentY        =   8493
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNI-Times"
         Size            =   9.75
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
      Height          =   375
      Index           =   0
      Left            =   5040
      Picture         =   "Frmplvt.frx":064C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   5040
      Picture         =   "Frmplvt.frx":098E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   5040
      Picture         =   "Frmplvt.frx":0CD0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   5040
      Picture         =   "Frmplvt.frx":1012
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   960
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "..."
      Top             =   5160
      Width           =   3375
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
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VAT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   11
      Top             =   120
      Width           =   375
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
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   3375
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
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmPLVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer
'====================================================================================================
' LiÖt kª danh s¸ch ph©n lo¹i vËt t­ thuéc tµi kho¶n ®­îc lùa chän
'====================================================================================================
Private Sub LietKePhanLoai()
    Dim rs_plvt As Recordset
    
    ClearGrid GrdNT, 20
    Set rs_plvt = DBKetoan.OpenRecordset("SELECT SoHieu, TenPhanLoai,VAT FROM PhanLoaiVattu ORDER BY SoHieu DESC", dbOpenSnapshot)
    
    Do While Not rs_plvt.EOF
        GrdNT.AddItem rs_plvt!SoHieu + Chr(9) + rs_plvt!TenPhanLoai + Chr(9) + CStr(rs_plvt!vat), 0
        rs_plvt.MoveNext
    Loop
    
    GrdNT.Rows = IIf(rs_plvt.RecordCount > 20, rs_plvt.RecordCount, 20)
    
    rs_plvt.Close
    Set rs_plvt = Nothing
End Sub
'====================================================================================================
' Thªm, Ghi, Xãa ph©n lo¹i vËt t­
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Me.MousePointer = 11
    Select Case Index
        Case 0:
            ThemMoi = 1
            txtNhap(0).Text = "..."
            txtNhap(1).Text = "..."
            txtNhap(0).SetFocus
        Case 1:
            If txtNhap(0).Text = "" Then
                txtNhap(0).SetFocus
                GoTo XongPL
            End If
            If txtNhap(1).Text = "" Then
                txtNhap(1).SetFocus
                GoTo XongPL
            End If
            If SoHieuOK(txtNhap(0).Text) Then
                  Select Case ThemMoi
                      Case 0:
                          GrdNT.col = 0
                          If GrdNT.Text = "" Then GoTo XongPL
                          If ExecuteSQL5("UPDATE PhanLoaiVattu SET SoHieu = '" + txtNhap(0).Text + "',TenPhanLoai='" _
                              + txtNhap(1).Text + "',VAT=" + txtNhap(2).Text + " WHERE SoHieu = '" + GrdNT.Text + "'") <> 0 Then GoTo XongPL
                          GrdNT.Text = txtNhap(0).Text
                          GrdNT.col = 1
                          GrdNT.Text = txtNhap(1).Text
                          GrdNT.col = 2
                          GrdNT.Text = txtNhap(2).Text
                      Case 1:
                          If ExecuteSQL5("INSERT INTO PhanLoaiVattu (SoHieu,TenPhanLoai,VAT) VALUES ('" _
                              + txtNhap(0).Text + "','" + txtNhap(1).Text + "'," + CStr(txtNhap(2).Text) + ")") <> 0 Then GoTo XongPL
                          GrdNT.AddItem txtNhap(0).Text + Chr(9) + txtNhap(1).Text + Chr(9) + txtNhap(2).Text, InsertGridRow(GrdNT, 0, txtNhap(0).Text)
                          ThemMoi = 0
                          GrdNT.Row = GrdNT.Rows - 1
                          GrdNT.col = 0
                          If GrdNT.Text = "" Then GrdNT.RemoveItem GrdNT.Row
                          GrdNT.Row = 0
                  End Select
            Else
                  MsgBox "Ph©n lo¹i ®· ®¨ng ký !", vbExclamation, App.ProductName
            End If
        Case 2:
            GrdNT.col = 0
            If GrdNT.Text = "" Then GoTo XongPL
            If ExecuteSQL5("DELETE FROM PhanLoaiVattu WHERE Sohieu='" + GrdNT.Text + "'") <> 0 Then GoTo XongPL
            GrdNT.RemoveItem GrdNT.Row
            If GrdNT.Rows < 20 Then GrdNT.Rows = 20
        Case 3:
            Unload Me
        Case 4:
            Dim MaPL As Long
            
            GrdNT.col = 0
            If GrdNT.Text <> "" Then
                        On Error Resume Next
                        Unload FrmVattu
                        On Error GoTo 0
                        
                        TenPLVT GrdNT.Text, MaPL
                        Load FrmVattu
                        FrmVattu.Tag = -MaPL
                        FrmVattu.Show 1
            End If
    End Select
XongPL:
    Me.MousePointer = 0
End Sub

Private Sub Form_Activate()
    HienThongBao Caption, 1
End Sub

'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask Then
        Select Case KeyCode
            Case vbKeyT:
                Command(0).SetFocus
                Command_Click 0
            Case vbKeyG:
                Command(1).SetFocus
                Command_Click 1
            Case vbKeyX:
                Command(2).SetFocus
                Command_Click 2
            Case vbKeyV:
                Command(3).SetFocus
                Command_Click 3
            Case vbKeyC:
                Command(4).SetFocus
                Command_Click 4
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
'====================================================================================================
' Khëi t¹o cöa sæ
'====================================================================================================
Private Sub Form_Load()
    ColumnSetUp GrdNT, 0, 820, 2
    ColumnSetUp GrdNT, 1, 3340, 0
    ColumnSetUp GrdNT, 2, 345, 0
    Caption = Caption + " - " + CStr(pNamTC)
    LietKePhanLoai
End Sub

Private Sub Form_Unload(Cancel As Integer)
        On Error Resume Next
        Unload FrmVattu
        On Error GoTo 0
End Sub

Private Sub GrdNt_click()
    With GrdNT
        .col = 0
        If .Text = "" Then Exit Sub
        txtNhap(0).Text = .Text
        .col = 1
        txtNhap(1).Text = .Text
        .col = 2
        txtNhap(2).Text = .Text
        .col = 0
        ThemMoi = 0
    End With
End Sub

Private Sub GrdNT_DblClick()
    txtNhap(0).SetFocus
End Sub

Private Sub txtNhap_GotFocus(Index As Integer)
    AutoSelect txtNhap(Index)
End Sub

Private Function SoHieuOK(sh As String) As Boolean
      Dim i As Integer
      
      SoHieuOK = False
      If ThemMoi = 1 Then
            With GrdNT
                  .col = 0
                  For i = 0 To .Rows - 1
                        If .Text = "" Then Exit For
                        If .Text = sh Then
                              Exit Function
                        End If
                  Next
            End With
      End If
      SoHieuOK = True
End Function

Private Sub txtNhap_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then KeyProcess KeyAscii
End Sub

Private Sub txtNhap_LostFocus(Index As Integer)
    Select Case Index
        Case 0, 1:
            If txtNhap(Index).Text = "" Then txtNhap(Index).Text = "..."
        Case 2:
            txtNhap(2).Text = Format(txtNhap(2).Text, Mask_0)
            If txtNhap(2).Text = "" Then txtNhap(2).Text = "0"
    End Select
End Sub
