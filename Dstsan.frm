VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmDSTaiSan 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch tµi s¶n"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   555
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
   Icon            =   "Dstsan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Fixed Assets List"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   2
      Left            =   6000
      Picture         =   "Dstsan.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "&Delete"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   8400
      Picture         =   "Dstsan.frx":6CC4
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   7200
      Picture         =   "Dstsan.frx":80E6
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Select"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   2
      ItemData        =   "Dstsan.frx":9548
      Left            =   5280
      List            =   "Dstsan.frx":954A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   4215
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   1
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   4215
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   0
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   3
      ItemData        =   "Dstsan.frx":954C
      Left            =   840
      List            =   "Dstsan.frx":954E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   1095
   End
   Begin MSGrid.Grid Grid 
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   9615
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   8705
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
      Rows            =   1
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      MouseIcon       =   "Dstsan.frx":9550
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KH"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   7560
      TabIndex        =   18
      Tag             =   "Year"
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Møc KH"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   8160
      TabIndex        =   17
      Tag             =   "Depreciation"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cßn l¹i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   6240
      TabIndex        =   16
      Tag             =   "Rest Value"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nguyªn gi¸"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   15
      Tag             =   "Original Value"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ngµy t¨ng"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   14
      Tag             =   "Inc. Date"
      Top             =   1320
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
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Tag             =   "Code"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tªn"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   12
      Tag             =   "Description"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Nhãm :"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   11
      Tag             =   "Group"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Lo¹i :"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   10
      Tag             =   "Class"
      Top             =   510
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tµi kho¶n :"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   9
      Tag             =   "Account"
      Top             =   150
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Th¸ng :"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Tag             =   "Month"
      Top             =   75
      Width           =   615
   End
End
Attribute VB_Name = "frmDSTaiSan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GRID_ROWS = 25
Private Const GRID_COLS = 8

Dim KhoiTao As Boolean
Dim TSChon As String
Dim SoLieu(1 To 12) As Boolean
' KEYDOWN
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      If (Shift And vbAltMask) > 0 Then
            Select Case KeyCode
                  Case vbKeyC And pNghiepVu = NV_KHONG: RFocus Command(0): Command_Click (0)
                  Case vbKeyC And Not pNghiepVu = NV_KHONG: RFocus Command(0): Command_Click (0)
                  Case vbKeyV:  RFocus Command(1): Command_Click (1)
                  Case vbKeyX:  RFocus Command(2): Command_Click (2)
            End Select
      End If
      If KeyCode = vbKeyEscape Then
            Unload frmDSTaiSan
            Set frmDSTaiSan = Nothing
      End If
End Sub
' LOAD
Private Sub Form_Load()
Dim chi_so As Integer
      ' NÕu nhËp ®Çu kú th× cho phÐp xo¸ tµi s¶n
      
      ' Khëi t¹o l­íi Grid
      InitGrid Grid, GRID_ROWS, GRID_COLS
      ColumnSetUp Grid, 0, 1, 0
      ColumnSetUp Grid, 1, 1300, 0
      ColumnSetUp Grid, 2, 2620, 0
      ColumnSetUp Grid, 3, 820, 2
      ColumnSetUp Grid, 4, 1300, 1
      ColumnSetUp Grid, 5, 1300, 1
      ColumnSetUp Grid, 6, 580, 2
      ColumnSetUp Grid, 7, 1300, 1
      
      ' LÊy danh s¸ch ph©n lo¹i
      KhoiTao = True
      Int_RecsetToCbo "SELECT SoHieu + '  ' + Ten AS F1, MaSo as F2 FROM LoaiTaiSan" _
                                                                                                                                        & " WHERE CapTren = 0", Combo(0)
      KhoiTao = False
      ' §Æt th¸ng ngÇm ®Þnh (dÉn ®Õn Events_Click t­¬ng øng - LÊy danh s¸ch tµi s¶n)
      AddMonthToCbo Combo(3)
      Select Case pNghiepVu
            Case NV_KHONG:                        Me.Caption = " Danh s¸ch tµi s¶n theo ph©n lo¹i vµ thêi gian"
            Case NV_GIAM, NV_DGLAI:     Me.Caption = " ChØ ®Þnh tµi s¶n bÞ t¸c ®éng"
      End Select

      pGhichungtu = 0
      Caption = Caption + " - " + CStr(pNamTC)
      
      For chi_so = 1 To 12
        SoLieu(chi_so) = False
      Next
    
      SetFont Me
End Sub
'======================================================================================
' COMBO
'======================================================================================
Private Sub Combo_Click(Index As Integer)
      Select Case Index
            Case 0             ' Tµi kho¶n
                  Int_RecsetToCbo "SELECT SoHieu + '  ' + Ten AS F1, MaSo as F2 FROM LoaiTaiSan " _
                        & "WHERE CapTren = " + CStr(Combo(0).ItemData(Combo(0).ListIndex)) + " ORDER BY SoHieu", Combo(1)
                  If Combo(1).ListCount = 0 Then
                        Combo(2).Clear
                        ClearGrid Grid, GRID_ROWS
                  End If
            Case 1            ' Ph©n lo¹i
                  Int_RecsetToCbo "SELECT SoHieu + '  ' + Ten AS F1, MaSo as F2 FROM LoaiTaiSan " _
                        & "WHERE CapTren = " + CStr(Combo(1).ItemData(Combo(1).ListIndex)) + " ORDER BY SoHieu", Combo(2)
                  If (Not KhoiTao) And Combo(2).ListCount = 0 Then LayDanhSachTaiSan
            Case 2, 3       ' Ph©n nhãm / Th¸ng
                  If KhoiTao Or Combo(1).ListCount = 0 Then Exit Sub Else LayDanhSachTaiSan
                  Command(2).Visible = (Combo(3).ItemData(Combo(3).ListIndex) = pThangDauKy)
      End Select
End Sub
'======================================================================================
' GRID
'======================================================================================
' CLICK
Private Sub Grid_Click()
      SendKeys "{Home}", True
      SetGridIndex Grid, Grid.Row
      If pNghiepVu = NV_KHONG And Me.tag = 1 Then
          Grid.col = 1
          TSChon = Grid.Text
      End If
End Sub
' DOUBLECLICK
Private Sub Grid_DblClick()
      Command_Click (0)
End Sub
' KEYDOWN
Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
            Case vbKeyHome, vbKeyEnd, vbKeyRight, vbKeyLeft
            Case vbKeyUp:                 SetGridIndex Grid, Grid.Row - 1
            Case vbKeyDown:            SetGridIndex Grid, Grid.Row + 1
            Case vbKeyPageUp:      SetGridIndex Grid, Grid.Row - GRID_ROWS
            Case vbKeyPageDown: SetGridIndex Grid, Grid.Row + GRID_ROWS
            Case vbKeyReturn: Command_Click (0)
            Case Else: Exit Sub
      End Select
      KeyCode = 0
End Sub
'======================================================================================
' command
'     1. Söa ®æi néi dung tµi s¶n
'     2. Gi¶m tµi s¶n
'     3. §¸nh gi¸ l¹i
'======================================================================================
Private Sub Command_Click(Index As Integer)
      Me.MousePointer = 11
      Select Case Index
            Case 0      ' Söa ®æi ........................................................................................................................................................
                  If pNghiepVu = NV_KHONG And Me.tag = 1 Then
                        Grid.col = 1
                        TSChon = Grid.Text
                        Unload Me
                        Exit Sub
                  End If
                  Grid.col = 0
                  If Len(Grid.Text) > 0 Then
                        ' Tµi s¶n ®· gi¶m trong n¨m sÏ kh«ng thÓ ®­îc ®¸nh gi¸ l¹i hay gi¶m lÇn thø 2
                        If Not pNghiepVu = NV_KHONG And TaiSanGiamTrongNam(CLng5(Grid.Text)) = True Then
                              Beep
                              MsgBox "Tµi s¶n ®· bÞ gi¶m trong n¨m. Ph¶i xo¸ chøng tõ gi¶m ®i nÕu muèn tiÕp tôc", vbExclamation
                              GoTo XongDSTS
                        End If
                        ' ChuÈn bÞ c¸c biÕn trao ®æi d÷ liÖu víi frmChungTu
                        pThangTacDong = Combo(3).ItemData(Combo(3).ListIndex)
                        pMaTaiSan = CLng5(Grid.Text)                     ' M· sè tµi s¶n hiÖn thêi
                        pMaChungTu = 0                                              ' Cho phÐp ghi chøng tõ míi
                        Select Case pNghiepVu                                ' Lo¹i nghiÖp vô ®­îc x¸c ®Þnh theo chøc
                              Case NV_KHONG:                                     ' n¨ng trªn menu ®· gäi ra frmDSTaiSan
                                    frmTaiSan.Show 1
                              Case NV_DGLAI:
                                    frmTangGiam.Show 1
                                    If pGhichungtu = 1 Then
                                        SetListIndex FrmChungtu.CboThang, CLng(pThangTacDong)
                                        Unload Me
                                        Exit Sub
                                    End If
                              Case NV_GIAM:
                                    GiamTaiSan pMaTaiSan, pThangTacDong
                                    ' Ghi chøng tõ
                                    pGhichungtu = 1
                                    SetListIndex FrmChungtu.CboThang, CLng(pThangTacDong)
                                    Unload Me
                                    Exit Sub
                        End Select
                        pMaTaiSan = 0
                        pMaChungTu = 0
                  Else
                        Beep
                  End If
                  Combo_Click (3)
                  DoEvents
            Case 1      ' Trë vÒ ...........................................................................................................................................................
                  TSChon = ""
                  SendKeys "{Escape}", False
            Case 2      ' Xo¸ ................................................................................................................................................................
                  Grid.col = 0
                  If Len(Grid.Text) > 0 Then
                        If vbYes = MsgBox("Xo¸ tµi s¶n hiÖn t¹i ?", vbQuestion + vbYesNo) Then
                              XoaTaiSan CLng5(Grid.Text)
                              Grid.RemoveItem Grid.Row
                              SetGridIndex Grid, 0
                              SoDuTKTS
                        End If
                  End If
      End Select
XongDSTS:
      Me.MousePointer = 0
End Sub
'======================================================================================
' SUB LayDanhSachTaiSan
'======================================================================================
Private Sub LayDanhSachTaiSan()
    Dim mnhom As Long, sql As String, i As Integer, chi_so As Integer
    Dim rs_danhsach As Recordset
            
      If Combo(2).ListCount > 0 Then
            mnhom = Combo(2).ItemData(Combo(2).ListIndex)
      Else
            mnhom = 0
      End If
      
      Me.MousePointer = 11
      ClearGrid Grid, GRID_ROWS
      i = Combo(3).ItemData(Combo(3).ListIndex)
      
      For chi_so = pThangDauKy To CThangDB(i)
            If SoLieu(chi_so) = False Then
                  CapNhatGiaTriTaiSan chi_so, FBcKt.GauGe
                  SoLieu(chi_so) = True
            End If
      Next
      
      sql = "SELECT TaiSan.MaSo, TaiSan.SoHieu, TaiSan.Ten, NCT, (ThongSo.NG_NS + ThongSo.NG_TBS + ThongSo.NG_CNK + ThongSo.NG_TD) AS NguyenGia, (ThongSo.CL_NS + ThongSo.CL_TBS + ThongSo.CL_CNK + ThongSo.CL_TD) AS ConLai, (ThongSo.KH_NS + ThongSo.KH_TBS + ThongSo.KH_CNK + ThongSo.KH_TD) AS KhauHao, NamKH" _
            & " FROM TaiSan LEFT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS" _
            & " WHERE (TaiSan.MaLoai = " + CStr(Combo(1).ItemData(Combo(1).ListIndex)) + ") AND ((TaiSan.MaNhom = " + CStr(mnhom) + ") OR (TaiSan.MaNhom=0)) " _
            + " AND " + WThang("ThangTang", 0, i) + " AND " + WThang("ThangGiam", i, 0) + " AND ThongSo.Thang= " + CStr(CThangDB(i)) _
            + " ORDER BY TaiSan.SoHieu DESC"
      Set rs_danhsach = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      On Error GoTo Err_NoCurrentRecord
      rs_danhsach.MoveFirst
      On Error GoTo 0
      Do Until rs_danhsach.EOF
            Grid.AddItem CStr(rs_danhsach!MaSo) + Chr(9) + rs_danhsach!sohieu + Chr(9) + rs_danhsach!Ten _
                + Chr(9) + Format(rs_danhsach!NCT, Mask_D) _
                + Chr(9) + Format(rs_danhsach!NguyenGia, Mask_0) + Chr(9) + Format(rs_danhsach!ConLai, Mask_0) _
                + Chr(9) + CStr(rs_danhsach!NamKH) + Chr(9) + Format(rs_danhsach!KhauHao, Mask_0), 0
            rs_danhsach.MoveNext
      Loop
Err_NoCurrentRecord:
      SetGridIndex Grid, 0
      rs_danhsach.Close
      Set rs_danhsach = Nothing
      Me.MousePointer = 1
End Sub
'======================================================================================
' FUNCTION TaiSanGiamTrongNam
'======================================================================================
Private Function TaiSanGiamTrongNam(ma_ts As Long) As Boolean
    Dim sql As String
    
      sql = "SELECT ThangGiam AS F1 FROM TaiSan WHERE MaSo = " + CStr(ma_ts)
      If CInt(SelectSQL(sql)) = 13 Then TaiSanGiamTrongNam = False Else TaiSanGiamTrongNam = True
End Function

Public Function ChonTaiSan(sh As String, tdau As Integer, tcuoi As Integer) As String
    Dim rs As Recordset
    
    Set rs = DBKetoan.OpenRecordset("SELECT LoaiTaiSan.MaSo AS MaTK, LoaiTaiSan_1.MaSo AS MaLoai, LoaiTaiSan_2.MaSo AS MaNhom, TaiSan.SoHieu, TaiSan.ThangTang, TaiSan.ThangGiam" _
        & " FROM ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) LEFT JOIN LoaiTaiSan AS LoaiTaiSan_1 ON TaiSan.MaLoai = LoaiTaiSan_1.MaSo) LEFT JOIN LoaiTaiSan AS LoaiTaiSan_2 ON TaiSan.MaNhom = LoaiTaiSan_2.MaSo" _
        & " WHERE TaiSan.SoHieu LIKE '" + sh + "*' AND " + WThang("ThangTang", 0, tcuoi) + " AND " + WThang("ThangGiam", tdau, 0) + " ORDER BY TaiSan.SoHieu", dbOpenSnapshot)
    If rs.RecordCount > 0 Then
        SetListIndex Combo(3), IIf(rs!ThangGiam < tcuoi, rs!ThangGiam, tcuoi)
        SetListIndex Combo(0), rs!MaTK
        SetListIndex Combo(1), rs!maloai
        If Not IsNull(rs!MaNhom) Then SetListIndex Combo(2), rs!MaNhom
        With Grid
            .col = 1
            .Row = 0
            Do While InStr(1, .Text, sh, vbTextCompare) = 0
                .Row = .Row + 1
            Loop
        End With
    End If
    rs.Close
    Set rs = Nothing
    pNghiepVu = NV_KHONG
    Me.tag = 1
        
    Grid_Click

    Me.Show 1
    ChonTaiSan = TSChon
End Function

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , Grid, Grid.col
    End If
End Sub
