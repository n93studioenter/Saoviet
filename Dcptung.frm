VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmDCPTung 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dông cô kÌm theo tµi s¶n"
   ClientHeight    =   3375
   ClientLeft      =   975
   ClientTop       =   2055
   ClientWidth     =   6975
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
   Icon            =   "Dcptung.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3375
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Equipment List"
   Begin MSGrid.Grid Grid 
      Height          =   2490
      Left            =   105
      TabIndex        =   12
      Top             =   420
      Width           =   5565
      _Version        =   65536
      _ExtentX        =   9816
      _ExtentY        =   4392
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
      Cols            =   5
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   5760
      Picture         =   "Dcptung.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "&Return"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   5760
      Picture         =   "Dcptung.frx":6C04
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "&Delete"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   5760
      Picture         =   "Dcptung.frx":80E6
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Save"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   5760
      Picture         =   "Dcptung.frx":9514
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Add"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tªn"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Tag             =   "Description"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§¬n vÞ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   10
      Tag             =   "Unit"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè l­îng"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   9
      Tag             =   "Quantity"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TrÞ gi¸"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   8
      Tag             =   "Amount"
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmDCPTung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GRID_ROWS = 11
Private Const GRID_COLS = 5
Private Type tpDCPTung
      MaSo As Long
      MaTS As Long
      Ten As String
      DonVi As String
      SoLuong As Double
      GiaThanh As Double
End Type
Dim DCPTung As tpDCPTung
' Key Down
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      If (Shift And vbAltMask) > 0 Then
            Select Case KeyCode
                  Case vbKeyT: RFocus Command(0): DoEvents: Command_Click (0)
                  Case vbKeyG:  RFocus Command(1): DoEvents: Command_Click (1)
                  Case vbKeyX: RFocus Command(2): DoEvents: Command_Click (2)
                  Case vbKeyV:  RFocus Command(3): DoEvents: Command_Click (3)
            End Select
      End If
      If KeyCode = vbKeyEscape Then
            Unload frmDCPTung
            Set frmDCPTung = Nothing
      End If
End Sub
' Load
Private Sub Form_Load()
Dim rs_dcptung As Recordset, sql As String
      ' L­íi Grid
      InitGrid Grid, GRID_ROWS, GRID_COLS
      ColumnSetUp Grid, 0, 1, 0
      ColumnSetUp Grid, 1, Label(0).Width - 60 * 0.75, 0
      ColumnSetUp Grid, 2, Label(1).Width - 40 * 0.75, 2
      ColumnSetUp Grid, 3, Label(2).Width - 40 * 0.75, 1
      ColumnSetUp Grid, 4, Label(3).Width - 360 * 0.75, 1
      ' LÊy danh s¸ch ®èi t­îng
      sql = "SELECT * FROM DCPTung WHERE MaTS = " + CStr(pMaTaiSan) + " ORDER BY Ten"
      Set rs_dcptung = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      Do Until rs_dcptung.EOF
            Grid.AddItem CStr(rs_dcptung!MaSo) _
                                           + Chr(9) + rs_dcptung!Ten _
                                           + Chr(9) + rs_dcptung!DonVi _
                                           + Chr(9) + Format(rs_dcptung!SoLuong, Mask_2) _
                                           + Chr(9) + Format(rs_dcptung!GiaThanh, Mask_0) ', 0
            rs_dcptung.MoveNext
      Loop
      SetGridIndex Grid, 0 ', GRID_ROWS, GRID_COLS
      rs_dcptung.Close
      Set rs_dcptung = Nothing
      DCPTung.MaTS = pMaTaiSan
      
      SetFont Me
End Sub
'======================================================================================
' GRID
'======================================================================================
' Click
Private Sub Grid_Click()
      SendKeys "{Home}", True
      SetGridIndex Grid, Grid.Row ', GRID_ROWS, GRID_COLS
End Sub
' Double Click
Private Sub Grid_DblClick()
      Grid.col = 0
      If Len(Grid.Text) > 0 Then
            DCPTung.MaSo = CLng5(Grid.Text)
            Grid.col = 1: Text(0).Text = Grid.Text
            Grid.col = 2: Text(1).Text = Grid.Text
            Grid.col = 3: Text(2).Text = Grid.Text
            Grid.col = 4: Text(3).Text = Grid.Text
            RFocus Text(0)
      Else
            DCPTung.MaSo = 0
      End If
End Sub
' Got Focus
Private Sub Grid_GotFocus()
      KhoiTao False
End Sub
' Key Down
Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
            Case vbKeyHome, vbKeyEnd, vbKeyRight, vbKeyLeft
            Case vbKeyUp:                 SetGridIndex Grid, Grid.Row - 1 ', GRID_ROWS, GRID_COLS
            Case vbKeyDown:            SetGridIndex Grid, Grid.Row + 1 ', GRID_ROWS, GRID_COLS
            Case vbKeyPageUp:      SetGridIndex Grid, Grid.Row - GRID_ROWS ', GRID_ROWS, GRID_COLS
            Case vbKeyPageDown: SetGridIndex Grid, Grid.Row + GRID_ROWS ', GRID_ROWS, GRID_COLS
            Case vbKeyReturn: Grid_DblClick
            Case vbKeyDelete: Command_Click (2)
            Case Else: Exit Sub
      End Select
      KeyCode = 0
End Sub
'======================================================================================
' command
'======================================================================================
Private Sub Command_Click(Index As Integer)
      Me.MousePointer = 11
      Select Case Index
            Case 0      ' Míi
                  KhoiTao True
            Case 1      ' Ghi
                  If HopLe = 0 Then
                        If DCPTung.MaSo = 0 Then
                              If ThemMoi = 0 Then KhoiTao True
                        Else
                              If SuaDoi = 0 Then KhoiTao False
                        End If
                  End If
            Case 2      ' Xo¸
                  Grid.col = 0: If Len(Grid.Text) = 0 Then Exit Sub
                  If vbYes = MsgBox("Xo¸ ®èi t­îng hiÖn t¹i", vbQuestion + vbYesNo) Then
                        If xoa = 0 Then KhoiTao False
                  End If
            Case 3      ' Trë vÒ
                  SendKeys "{Escape}", False
      End Select
      Me.MousePointer = 0
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , Grid, Grid.col
    End If
End Sub

'======================================================================================
' TEXT
'======================================================================================
' Got Focus
Private Sub Text_GotFocus(Index As Integer)
      AutoSelect Text(Index)
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index > 1 Then KeyProcess Text(Index), KeyAscii
End Sub

' Lost Focus
Private Sub Text_LostFocus(Index As Integer)
      If Len(Text(Index).Text) = 0 Then
            If Index < 2 Then
                  Text(Index).Text = "(...)"
            Else
                  Text(Index).Text = "0"
            End If
      End If
      On Error GoTo Err_DataTypeConvertion
      Select Case Index
            Case 0
                  DCPTung.Ten = Text(0).Text
            Case 1
                  DCPTung.DonVi = Text(1).Text
            Case 2
                  DCPTung.SoLuong = Cdbl5(Text(2).Text)
            Case 3
                  DCPTung.GiaThanh = Cdbl5(Text(3).Text)
      End Select
      On Error GoTo 0
      Exit Sub
Err_DataTypeConvertion:
      RFocus Text(Index)
End Sub
'======================================================================================
' SUB HopLe
'======================================================================================
Private Function HopLe()
Dim thong_bao As String
      If Len(DCPTung.Ten) = 0 Then Text_LostFocus (0)
      If Len(DCPTung.DonVi) = 0 Then Text_LostFocus (1)
      If DCPTung.SoLuong = 0 Then Text_LostFocus (2)
      If DCPTung.GiaThanh = 0 Then Text_LostFocus (3)
      If DCPTung.MaTS = 0 Then thong_bao = "Tµi s¶n ch­a ®­îc ®¨ng ký": GoTo Err_InValidate
      If Len(DCPTung.Ten) = 0 Or DCPTung.Ten = "(...)" Then thong_bao = "ThiÕu tªn cña dông cô hoÆc phô tïng kÌm theo tµi s¶n": GoTo Err_InValidate
      If Len(DCPTung.DonVi) = 0 Then thong_bao = "Ch­a cã ®¬n vÞ tÝnh cña dông cô hoÆc phô tïng kÌm theo tµi s¶n": GoTo Err_InValidate
      If DCPTung.SoLuong = 0 Then thong_bao = "ThiÕu sè l­îng cña dông cô hoÆc phô tïng kÌm theo tµi s¶n": GoTo Err_InValidate
      If DCPTung.GiaThanh = 0 Then thong_bao = "ThiÕu gi¸ thµnh cña dông cô hoÆc phô tïng kÌm theo tµi s¶n": GoTo Err_InValidate
      HopLe = 0
      Exit Function
Err_InValidate:
      Beep
      MsgBox thong_bao, vbCritical
      HopLe = -1
End Function
'======================================================================================
' FUNCTION ThemMoi
'======================================================================================
Private Function ThemMoi() As Integer
      If ExecuteSQL5("INSERT INTO DCPTung (MaSo, MaTS, Ten, DonVi, SoLuong, GiaThanh) VALUES (" _
        + CStr(Lng_MaxValue("MaSo", "DCPTung") + 1) + "," + CStr(DCPTung.MaTS) + ",'" + DCPTung.Ten + "','" + DCPTung.DonVi + "'," _
        + DoiDau(DCPTung.SoLuong) + "," + DoiDau(DCPTung.GiaThanh) + ")") = 0 Then
'      pSelectSQL = "SELECT Max(MaSo) AS F1 FROM DCPTung"
            DCPTung.MaSo = Lng_MaxValue("MaSo", "DCPTung")
            Grid.AddItem CStr(DCPTung.MaSo) + Chr(9) + DCPTung.Ten _
                + Chr(9) + DCPTung.DonVi _
                + Chr(9) + Format(DCPTung.SoLuong, Mask_2) _
                + Chr(9) + Format(DCPTung.GiaThanh, Mask_0) ', NewRowIndex(Grid, 0)
            SetGridIndex Grid, 0 ', GRID_ROWS, GRID_COLS
            ThemMoi = 0
      Else
            DCPTung.MaSo = 0
            ThemMoi = -1
      End If
End Function
'======================================================================================
' FUNCTION SuaDoi
'======================================================================================
Private Function SuaDoi() As Integer
      If ExecuteSQL5("UPDATE DCPTung SET Ten = '" + DCPTung.Ten + "'" _
        + ", DonVi = '" + DCPTung.DonVi + "', SoLuong = " + DoiDau(DCPTung.SoLuong) _
        + ", GiaThanh = " + DoiDau(DCPTung.GiaThanh) _
        + " WHERE MaSo = " + CStr(DCPTung.MaSo)) = 0 Then
            Grid.col = 1: Grid.Text = Text(0).Text
            Grid.col = 2: Grid.Text = Text(1).Text
            Grid.col = 3: Grid.Text = Format(Text(2).Text, Mask_2)
            Grid.col = 4: Grid.Text = Format(Text(3).Text, Mask_0)
            SuaDoi = 0
      Else
            SuaDoi = -1
      End If
End Function
'======================================================================================
' FUNCTION Xoa
'======================================================================================
Private Function xoa() As Integer
      If ExecuteSQL5("DELETE * FROM DCPTung WHERE MaSo = " + Grid.Text) = 0 Then
            Grid.RemoveItem Grid.Row
            SetGridIndex Grid, Grid.Row - 1 ', GRID_ROWS, GRID_COLS
            xoa = 0
      Else
            xoa = -1
      End If
End Function
'======================================================================================
' SUB KhoiTao
'======================================================================================
Private Sub KhoiTao(tiep_tuc As Boolean)
      DCPTung.MaSo = 0
      DCPTung.Ten = ""
      DCPTung.DonVi = ""
      DCPTung.SoLuong = 0
      DCPTung.GiaThanh = 0
      Text(0).Text = ""
      Text(1).Text = ""
      Text(2).Text = ""
      Text(3).Text = ""
      If tiep_tuc = True Then RFocus Text(0)
End Sub



