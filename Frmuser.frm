VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmUser 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch ng­êi sö dông"
   ClientHeight    =   3135
   ClientLeft      =   1650
   ClientTop       =   3015
   ClientWidth     =   7950
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
   Icon            =   "Frmuser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "User List"
   Begin MSGrid.Grid GrdNT 
      Height          =   2415
      Left            =   120
      TabIndex        =   23
      Tag             =   "0"
      Top             =   360
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   4260
      _StockProps     =   77
      BackColor       =   16777215
      Rows            =   1
      Cols            =   10
      FixedRows       =   0
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   6000
      TabIndex        =   10
      Top             =   2800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   9
      Top             =   2800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   8
      Top             =   2800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   7
      Top             =   2800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   4
      Left            =   6720
      Picture         =   "Frmuser.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "&Detail"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   6
      Top             =   2800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   2800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00FFFFC0&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   2800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   3
      Left            =   6720
      Picture         =   "Frmuser.frx":6C44
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "&Return"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   2
      Left            =   6720
      Picture         =   "Frmuser.frx":8066
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "&Delete"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   1
      Left            =   6720
      Picture         =   "Frmuser.frx":9548
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "&Save"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   0
      Left            =   6720
      Picture         =   "Frmuser.frx":A976
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "&Add"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox CboRight 
      Height          =   315
      ItemData        =   "Frmuser.frx":BED0
      Left            =   1920
      List            =   "Frmuser.frx":BEE0
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   120
      MaxLength       =   50
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XuÊt"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   3960
      TabIndex        =   22
      Tag             =   "Export"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Söa"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   21
      Tag             =   "Update"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TH"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   20
      Tag             =   "Common"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TC"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   19
      Tag             =   "Money"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BH"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   17
      Tag             =   "Sale"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TSC§"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   16
      Tag             =   "Assets"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NhËp"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   15
      Tag             =   "Import"
      Top             =   120
      Width           =   495
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
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer

Private Sub Chk_Click(Index As Integer)
    Dim i As Integer
    
    If Chk(0).Value = 1 Then
        For i = 1 To 6
            Chk(i).Value = 1
        Next
    End If
    For i = 1 To 6
        Chk(i).Enabled = (Chk(0).Value <> 1)
    Next
End Sub

'====================================================================================================
' Thªm, Ghi, Xãa nguyªn tÖ
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim i As Integer, sql As String
    
    If (Index = 1 And ThemMoi = 0) Or Index = 2 Then
        GrdNT.col = 9
        If CInt5(GrdNT.Text) = 0 Then
            MsgBox "Kh«ng thay ®æi quyÒn Admin !", vbExclamation, App.ProductName
            Exit Sub
        End If
    End If

    Select Case Index
        Case 0:
            ThemMoi = 1
            txtNhap.Text = ""
            RFocus txtNhap
        Case 1:
            If Len(txtNhap.Text) = 0 Then
                RFocus txtNhap
                Exit Sub
            End If
            If CboRight.ItemData(CboRight.ListIndex) = 0 Then
                MsgBox "Kh«ng sö dông quyÒn Admin !", vbExclamation, App.ProductName
                Exit Sub
            End If
            Select Case ThemMoi
                Case 0:
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then Exit Sub
                    If ExecuteSQL5("UPDATE Users SET TenNSD = '" + txtNhap.Text + "', UserRight = " + CStr(CboRight.ListIndex) + ", VT = " + CStr(MakeR) _
                        + " WHERE TenNSD = '" + GrdNT.Text + "'") <> 0 Then Exit Sub
                    GrdNT.Text = txtNhap.Text
                    GrdNT.col = 1
                    GrdNT.Text = CboRight.Text
                    For i = 2 To 8
                        GrdNT.col = i
                        GrdNT.Text = IIf(Chk(i - 2).Value = 1, "X", "")
                    Next
                    GrdNT.col = 9
                    GrdNT.Text = CboRight.ListIndex
                Case 1:
                    If ExecuteSQL5("INSERT INTO Users (MaSo,TenNSD, UserRight, Psw, VT) VALUES (" + CStr(Lng_MaxValue("MaSo", "Users") + 1) + ",'" + txtNhap.Text + "'," _
                        + CStr(CboRight.ListIndex) + "," + CStr(pNamTC) + "," + CStr(MakeR) + ")") <> 0 Then Exit Sub
                    For i = 0 To 6
                        sql = sql + Chr(9) + IIf(Chk(i).Value = 1, "X", "")
                    Next
                    GrdNT.AddItem txtNhap.Text + Chr(9) + CboRight.Text + sql + Chr(9) + CStr(CboRight.ListIndex) ', InsertGridRow(GrdNT, 0, txtNhap.Text)
                    ThemMoi = 0
                    GrdNT.Row = GrdNT.Rows - 1
                    GrdNT.col = 0
                    If Len(GrdNT.Text) = 0 Then GrdNT.RemoveItem GrdNT.Row
                    GrdNT.Row = 0
                    ThemMoi = 0
            End Select
        Case 2:
            GrdNT.col = 0
            If Len(GrdNT.Text) = 0 Then Exit Sub
            If SelectSQL("SELECT ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN Users ON ChungTu.User_ID=Users.MaSo WHERE TenNSD = '" + GrdNT.Text + "'") = 0 Then
                If ExecuteSQL5("DELETE FROM Users WHERE TenNSD = '" + GrdNT.Text + "'") <> 0 Then Exit Sub
                GrdNT.RemoveItem GrdNT.Row
                If GrdNT.Rows <= 10 Then GrdNT.Rows = 10
            Else
                MsgBox "Ng­êi sö dông ®· nhËp chøng tõ!", vbCritical, App.ProductName
            End If
        Case 3:
            Unload Me
        Case 4:
            With GrdNT
                .col = 0
                If Len(.Text) = 0 Then Exit Sub
                'Load FrmUser2
                FrmUser2.tag = SelectSQL("SELECT MaSo AS F1 FROM Users WHERE TenNSD='" + .Text + "'")
                FrmUser2.Show vbModal
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
    ColumnSetUp GrdNT, 0, 1780, 0
    ColumnSetUp GrdNT, 1, 1060, 0
    ColumnSetUp GrdNT, 2, 460, 2
    ColumnSetUp GrdNT, 3, 460, 2
    ColumnSetUp GrdNT, 4, 460, 2
    ColumnSetUp GrdNT, 5, 460, 2
    ColumnSetUp GrdNT, 6, 460, 2
    ColumnSetUp GrdNT, 7, 460, 2
    ColumnSetUp GrdNT, 8, 460, 2
    ColumnSetUp GrdNT, 9, 1, 0
    CboRight.ListIndex = 0
    LietKeUser
    
    Command(4).Visible = (pSoKT Mod 100 >= 10)
    
    SetFont Me
End Sub

Private Sub GrdNt_click()
    Dim i As Integer
    
    SendKeys "{Home}", True
    SetGridIndex GrdNT, GrdNT.Row
    
    With GrdNT
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        txtNhap.Text = .Text
        .col = 9
        CboRight.ListIndex = CInt5(.Text)
        For i = 2 To 8
            .col = i
            If Len(.Text) > 0 Then Chk(i - 2).Value = 1 Else Chk(i - 2).Value = 0
        Next
        .col = 0
        RFocus txtNhap
    End With
End Sub

Private Sub GrdNt_KeyPress(KeyAscii As Integer)
    SendKeys "{Home}", True
    SetGridIndex GrdNT, GrdNT.Row
    If KeyAscii = 13 Then GrdNt_click
End Sub

Private Sub txtNhap_GotFocus()
    AutoSelect txtNhap
End Sub
'====================================================================================================
' HiÓn thÞ danh s¸ch nguyªn tÖ
'====================================================================================================
Private Sub LietKeUser()
    Dim rs_ngte As Recordset, i As Integer, sql As String, k As Long
    
    Set rs_ngte = DBKetoan.OpenRecordset("SELECT Users.* FROM Users ORDER BY TenNSD DESC", dbOpenSnapshot)
    Do While Not rs_ngte.EOF
        sql = ""
        For i = 0 To 6
            k = 10 ^ i
            sql = sql + Chr(9) + IIf(rs_ngte!vt Mod 10 * k >= k, "X", "")
        Next

        GrdNT.AddItem rs_ngte!TenNSD + Chr(9) + CboRight.List(rs_ngte!UserRight) + sql + Chr(9) + CStr(rs_ngte!UserRight), 0
        rs_ngte.MoveNext
    Loop
   GrdNT.Rows = IIf(rs_ngte.RecordCount > GrdNT.tag, rs_ngte.RecordCount, GrdNT.tag)
    GrdNT.Row = 0
    GrdNt_click
    rs_ngte.Close
    Set rs_ngte = Nothing
End Sub

Private Function MakeR() As Long
    Dim i As Integer, r As Long
    
    For i = 0 To 6
        If Chk(i).Value = 1 Then r = r + (10 ^ i)
    Next
    MakeR = r
End Function
