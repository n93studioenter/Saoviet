VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmDuphong 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dù phßng gi¶m gi¸ hµng tån kho"
   ClientHeight    =   7080
   ClientLeft      =   195
   ClientTop       =   705
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
   Icon            =   "Frmdupho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Provision for decreasing stock "
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ph©n lo¹i vËt t­"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   3
      Tag             =   "Inventory Class"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LËp cho kho"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   1
      Tag             =   "Store"
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      Index           =   1
      ItemData        =   "Frmdupho.frx":57E2
      Left            =   6240
      List            =   "Frmdupho.frx":580A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   3255
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      Index           =   0
      ItemData        =   "Frmdupho.frx":5835
      Left            =   6240
      List            =   "Frmdupho.frx":585D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.ComboBox CboThang 
      Height          =   315
      ItemData        =   "Frmdupho.frx":5888
      Left            =   840
      List            =   "Frmdupho.frx":588A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtGia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Frmdupho.frx":588C
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdghi 
      Caption         =   "­"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8160
      TabIndex        =   7
      Top             =   6240
      Width           =   255
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   3
      Left            =   4800
      Picture         =   "Frmdupho.frx":588E
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "&List"
      Top             =   6600
      Width           =   1095
   End
   Begin MSGrid.Grid GrdGia 
      Height          =   5175
      Left            =   120
      TabIndex        =   5
      Tag             =   "20"
      Top             =   1080
      Width           =   9615
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   9128
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
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   4
      ScrollBars      =   2
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   2
      Left            =   8400
      Picture         =   "Frmdupho.frx":6C58
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   7200
      Picture         =   "Frmdupho.frx":807A
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "&Print"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   6000
      Picture         =   "Frmdupho.frx":94DC
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "&View"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label LbDuphong 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label LbDuphong 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   13
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label LbDuphong 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   14
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label LbDuphong 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   4560
      TabIndex        =   15
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label LbDuphong 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   5760
      TabIndex        =   16
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th¸ng"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Tag             =   "Month"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tªn vËt t­"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   23
      Tag             =   "Description"
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§.v.t"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   22
      Tag             =   "Unit"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§¬n gi¸"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   21
      Tag             =   "Unit price"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§¬n gi¸ Thùc tÕ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6960
      TabIndex        =   20
      Tag             =   "Real Price"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè hiÖu vËt t­"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Tag             =   "Code"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chªnh lÖch"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   8160
      TabIndex        =   18
      Tag             =   "Difference"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè l­îng tån"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   17
      Tag             =   "Quantity"
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "FrmDuphong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 0, 1:
            If Cbo(0).ListCount = 0 Or Cbo(1).ListCount = 0 Then
                ErrMsg er_VTKoTon
                Exit Sub
            End If
            InDuPhong CboThang.ItemData(CboThang.ListIndex), IIf(Chk(0).Value = 1, Cbo(0).ItemData(Cbo(0).ListIndex), 0), IIf(Chk(1).Value = 1, Cbo(1).ItemData(Cbo(1).ListIndex), 0), Index
        Case 2:
            Unload Me
        Case 3:
            If Cbo(0).ListCount = 0 Or Cbo(1).ListCount = 0 Then Exit Sub
            ListGia CboThang.ItemData(CboThang.ListIndex), IIf(Chk(0).Value = 1, Cbo(0).ItemData(Cbo(0).ListIndex), 0), IIf(Chk(1).Value = 1, Cbo(1).ItemData(Cbo(1).ListIndex), 0)
    End Select
End Sub

Private Sub cmdghi_Click()
    Dim gia As Double, giacu As Double, SoLuong As Double, sh As String
    
    gia = Cdbl5(txtGia.Text)
    With GrdGia
            .col = 0
            sh = .Text
            If Len(sh) = 0 Then Exit Sub
            .col = 5
            .Text = Format(gia, Mask_2)
            .col = 4
            giacu = Cdbl5(.Text)
            .col = 3
            SoLuong = Cdbl5(.Text)
            .col = 6
            .Text = Format(SoLuong * (gia - giacu), Mask_0)
            DBKetoan.Execute "UPDATE Vattu SET GiaDuPhong = " + DoiDau(gia) + " WHERE SoHieu = '" + sh + "'"
            .col = 0
            .Row = .Row + 1
    End With
    GrdGia_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyV:
                RFocus Command(2)
                Command_Click 2
            Case vbKeyI:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyX:
                RFocus Command(0)
                Command_Click 0
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    ColumnSetUp GrdGia, 0, 1300, 2
    ColumnSetUp GrdGia, 1, 2260, 0
    ColumnSetUp GrdGia, 2, 820, 2
    ColumnSetUp GrdGia, 3, 1180, 1
    ColumnSetUp GrdGia, 4, 1180, 1
    ColumnSetUp GrdGia, 5, 1180, 1
    ColumnSetUp GrdGia, 6, 1300, 1
    ColumnSetUp GrdGia, 7, 1, 0
    
    AddMonthToCbo CboThang
    
    Int_RecsetToCbo "SELECT MaSo As F2,TenKho As F1 FROM KhoHang ORDER BY TenKho", Cbo(0)
    Int_RecsetToCbo "SELECT MaSo As F2, SoHieu + ' - ' + TenPhanLoai As F1 FROM PhanLoaiVattu ORDER BY SoHieu", Cbo(1)
        
    Caption = Caption + " - " + CStr(pNamTC)
    
    SetFont Me
End Sub

Private Sub GrdGia_Click()
    Dim i As Integer
    
    SendKeys "{Home}", True
    SetGridIndex GrdGia, GrdGia.Row
    
    With GrdGia
        .col = 0
        If Len(.Text) > 0 Then
            For i = 0 To 4
                .col = i
                LbDuphong(i).Caption = .Text
            Next
            .col = 4
            txtGia.Text = .Text
            RFocus txtGia
        End If
    End With
End Sub

Private Sub GrdGia_KeyPress(KeyAscii As Integer)
    SendKeys "{Home}", True
    SetGridIndex GrdGia, GrdGia.Row
    
    If KeyAscii = 13 Then GrdGia_Click
End Sub

Private Sub GrdGia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdGia, GrdGia.col
    End If
End Sub

Private Sub txtGia_GotFocus()
    AutoSelect txtGia
End Sub

Private Sub txtGia_KeyPress(KeyAscii As Integer)
    KeyProcess txtGia, KeyAscii
End Sub

'====================================================================================================
' Thu tuc liet ke vat tu va gia du phong
'====================================================================================================
Private Sub ListGia(thang As Integer, mkho As Long, maloai As Long)
    Dim rs_vattu As Recordset, tien As Double, luong As Double, i As Integer, dongia As Double, chenhlech As Double, tien2 As Double
        
    Me.MousePointer = 11
    ClearGrid GrdGia, GrdGia.tag
    Set rs_vattu = DBKetoan.OpenRecordset("SELECT Vattu.MaSo, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, Vattu.GiaDuPhong " _
        & " FROM (Vattu INNER JOIN TonKho ON Vattu.MaSo=TonKho.MaVattu) INNER JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo " _
        & " WHERE Luong_" + CStr(thang) + "<>0 " + IIf(mkho > 0, " AND MaSoKho = " + CStr(mkho), "") + IIf(maloai <> 0, "PhanLoaiVattu.SoHieu LIKE '" + MaSo2SoHieu(maloai, "PhanLoaiVattu") + "*'", "") _
        + " ORDER BY PhanLoaiVattu.SoHieu DESC, Vattu.SoHieu DESC", dbOpenSnapshot, dbForwardOnly)
    i = 0
    Do While Not rs_vattu.EOF
        luong = SoTonKho(thang, mkho, 0, rs_vattu!MaSo, tien, tien2)
        If luong <> 0 Then
            i = i + 1
            If rs_vattu!GiaDuPhong > 0 Then
                dongia = rs_vattu!GiaDuPhong
            Else
                dongia = Fix(0.5 + Mask_N * tien / luong) / Mask_N
                ExecuteSQL5 "UPDATE Vattu SET GiaDuPhong = " + DoiDau(dongia) + " WHERE MaSo = " + CStr(rs_vattu!MaSo)
            End If
            chenhlech = IIf(rs_vattu!GiaDuPhong > 0, luong * (dongia - tien / luong), 0)
            GrdGia.AddItem rs_vattu!sohieu + Chr(9) + rs_vattu!TenVattu + Chr(9) + rs_vattu!DonVi + Chr(9) + Format(luong, Mask_2) + Chr(9) + Format(tien / luong, Mask_2) + Chr(9) + Format(dongia, Mask_2) + Chr(9) + Format(chenhlech, Mask_0) + Chr(9) + CStr(rs_vattu!MaSo), 0
        End If
        rs_vattu.MoveNext
    Loop
    rs_vattu.Close
    Set rs_vattu = Nothing
    GrdGia.Rows = IIf(i > GrdGia.tag, i, GrdGia.tag)
    Me.MousePointer = 0
    GrdGia.Row = 0
End Sub
'======================================================================================
' Thu tuc in du phong giam gia hang ton kho
'======================================================================================
Private Sub InDuPhong(thang As Integer, mkho As Long, mloai As Long, dest As Integer)
    Dim rs_luyke As Recordset
    Dim luong As Double, tien As Double, tien2 As Double
    
    Me.MousePointer = 11
    ExecuteSQL5 "DELETE FROM BaoCaoCP"
    
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu) SELECT DISTINCTROW Vattu.MaSo,Cstr(Vattu.MaSo) " _
        & " FROM (Vattu INNER JOIN TonKho ON Vattu.MaSo=TonKho.MaVattu) INNER JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo " _
        & " WHERE Luong_" + CStr(thang) + "<>0 " + IIf(mkho > 0, " AND MaSoKho = " + CStr(mkho), "") + IIf(mloai <> 0, "PhanLoaiVattu.SoHieu LIKE '" + MaSo2SoHieu(mloai, "PhanLoaiVattu") + "*'", "") _
        + " GROUP BY Vattu.MaSo"
    
    If DBKetoan.RecordsAffected = 0 Then
        Me.MousePointer = 0
        ErrMsg er_VTKoTon
        Exit Sub
    End If
    
    Set rs_luyke = DBKetoan.OpenRecordset("SELECT * FROM BaoCaoCP", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_luyke.EOF
        luong = SoTonKho(thang, 0, 0, rs_luyke!MaSo, tien, tien2)
        ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1 = " + DoiDau(luong) + ", Kq2 = " + DoiDau(tien) + " WHERE MaSo = " + CStr(rs_luyke!MaSo)
        rs_luyke.MoveNext
    Loop
    rs_luyke.Close
    Set rs_luyke = Nothing
    
    SetSQL "QDuPhong", "SELECT DISTINCTROW PhanLoaiVattu.SoHieu, PhanLoaiVattu.TenPhanLoai, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, Vattu.GiaDuPhong, BaoCaoCP.Kq1, BaoCaoCP.Kq2" _
        & " FROM PhanLoaiVattu INNER JOIN (Vattu INNER JOIN BaoCaoCP ON Vattu.MaSo = BaoCaoCP.BC_ID) ON PhanLoaiVattu.MaSo = Vattu.MaPhanLoai ORDER BY PhanLoaiVattu.SoHieu,Vattu.SoHieu"
    
    SetRptInfo
    frmMain.Rpt.Formulas(3) = "Thang=" + CStr(thang)
    If mkho > 0 Then
        frmMain.Rpt.Formulas(4) = "TenKho = 'Kho: " + Cbo(0).Text + "'"
    Else
        frmMain.Rpt.Formulas(4) = "TenKho = '" + ABCtoVNI("Tæng hîp c¸c kho") + "'"
    End If
    frmMain.Rpt.ReportFileName = "DUPHONG.RPT"
    RptSetDate NgayCuoiThang(pNamTC, thang)
    frmMain.Rpt.Destination = dest
    InBaoCaoRPT
    Me.MousePointer = 0
End Sub
