VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FVTDauKy 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tån kho ®Çu kú"
   ClientHeight    =   7080
   ClientLeft      =   870
   ClientTop       =   735
   ClientWidth     =   9945
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
   Icon            =   "Fvtdauky.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.CommandButton loadexel 
      BackColor       =   &H8000000E&
      Caption         =   "CËp nhËt tõ exel"
      Height          =   375
      Left            =   3000
      MaskColor       =   &H8000000E&
      TabIndex        =   32
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton laydatabase 
      BackColor       =   &H8000000E&
      Caption         =   "CËp nhËt tõ data"
      Height          =   375
      Left            =   1320
      MaskColor       =   &H8000000E&
      TabIndex        =   31
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton xoa 
      BackColor       =   &H8000000E&
      Caption         =   "Xãa tån ®Çu kú"
      Height          =   375
      Left            =   4560
      MaskColor       =   &H8000000E&
      TabIndex        =   30
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   9480
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Fvtdauky.frx":57E2
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H8000000E&
      Caption         =   "&Tån kho"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   26
      Tag             =   "&Done"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox CboDV 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Fvtdauky.frx":57E4
      Left            =   4560
      List            =   "Fvtdauky.frx":57E6
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Danh s¸ch ®¬n vÞ tÝnh"
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H8000000E&
      Caption         =   "&Xö lý"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "&Done"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      ItemData        =   "Fvtdauky.frx":57E8
      Left            =   5400
      List            =   "Fvtdauky.frx":57EA
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   8160
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "Fvtdauky.frx":57EC
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   6840
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Fvtdauky.frx":57EE
      Top             =   6240
      Width           =   1335
   End
   Begin MSGrid.Grid GrdVT 
      Height          =   5415
      Left            =   120
      TabIndex        =   20
      Tag             =   "30"
      Top             =   840
      Width           =   9615
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   9551
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
      Cols            =   11
      FixedRows       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin VB.CommandButton cmdct 
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      Picture         =   "Fvtdauky.frx":57F0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   255
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   8400
      Picture         =   "Fvtdauky.frx":5B92
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   5520
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Fvtdauky.frx":6FB4
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2415
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   2
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   20
      TabIndex        =   1
      Top             =   6240
      Width           =   975
   End
   Begin VB.ComboBox CboKho 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   9560
      TabIndex        =   29
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tû gi¸ ®Çu kú:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   28
      Tag             =   "Store"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thµnh tiÒn USD"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   9480
      TabIndex        =   27
      Tag             =   "Amount"
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th¸ng"
      Height          =   255
      Index           =   9
      Left            =   4680
      TabIndex        =   24
      Tag             =   "Month"
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§¬n gi¸"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   22
      Tag             =   "Unit Price"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè hiÖu TK"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Tag             =   "Account"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tæng tiÒn"
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   19
      Tag             =   "Total"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   8120
      TabIndex        =   18
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thµnh tiÒn"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   8160
      TabIndex        =   17
      Tag             =   "Amount"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè l­îng"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   16
      Tag             =   "Quantity"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§.vÞ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   15
      Tag             =   "Unit"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tªn vËt t­"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   14
      Tag             =   "Description"
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè hiÖu VT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Tag             =   "Inventory Code"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Kho"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Tag             =   "Store"
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FVTDauKy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim taikhoan As New ClsTaikhoan
Dim vattu As New ClsVattu
Dim phatsinh As Boolean
Dim f1 As Integer, tbl As String
Dim tygia As Double

Private Sub Cbo_Click()
    If f1 > 0 And CboKho.ListIndex >= 0 Then LietKeTonKho CboKho.ItemData(CboKho.ListIndex), IIf(f1 = 0, 0, Cbo.ItemData(Cbo.ListIndex)), 0
End Sub

Private Sub CboKho_Click()
    If f1 >= 0 And CboKho.ListIndex >= 0 Then LietKeTonKho CboKho.ItemData(CboKho.ListIndex), IIf(f1 = 0, 0, Cbo.ItemData(Cbo.ListIndex)), 0
End Sub

Public Sub CmdCt_Click()
    Dim luong As Double, tien As Double, i As Integer, dgia As Double, dvt As Long, tien2 As Double
    
    If CboKho.ListIndex < 0 Then
        ErrMsg er_KhoHang
        Exit Sub
    End If
    
    If taikhoan.MaSo = 0 Or taikhoan.tkcon > 0 Or taikhoan.tk_id <> TKVT_ID Then
        ErrMsg er_SHTaiKhoan1
        RFocus txtTon(0)
        Exit Sub
    End If
        
    If vattu.MaSo = 0 Then
        ErrMsg er_SHVattu
        RFocus txtTon(1)
        Exit Sub
    End If
    
    luong = Cdbl5(txtTon(4).Text)
    tien = Cdbl5(txtTon(6).Text)
    tien2 = Cdbl5(txtTon(7).Text)
    If luong <> 0 Then dgia = Fix(0.5 + Mask_N * tien / luong) / Mask_N Else dgia = 0
    
    If (tien <> 0) And (luong = 0) Then
        MsgBox "H·y nhËp sè l­îng tån kho !", vbExclamation, App.ProductName
        RFocus txtTon(4)
        Exit Sub
    End If
    
    If vattu.Dvt2 > 0 Then
        dvt = CboDV.ItemData(CboDV.ListIndex)
    Else
        dvt = 0
    End If
    
    Me.MousePointer = 0
    With GrdVT
        If OutCost = 0 Or f1 > 0 Then
            For i = 0 To .Rows - 1
                .col = 7
                .Row = i
                If Len(.Text) = 0 Then Exit For
                If CLng5(.Text) = taikhoan.MaSo Then
                    .col = 8
                    If CLng5(.Text) = vattu.MaSo Then
                        GrdVT.RemoveItem i
                        GrdVT.AddItem taikhoan.sohieu + Chr(9) + vattu.sohieu + Chr(9) + vattu.TenVattu + Chr(9) + IIf(dvt = 0, vattu.DonVi, CboDV.Text) + Chr(9) _
                            + Format(luong, Mask_2) + Chr(9) + Format(dgia, Mask_2) + Chr(9) + Format(tien, Mask_0) + Chr(9) + CStr(taikhoan.MaSo) _
                            + Chr(9) + CStr(vattu.MaSo) + Chr(9) + CStr(dvt) + IIf(pGiaUSD > 0, Chr(9) + Format(tien2, Mask_2), ""), i
                        If dvt > 0 Then luong = QuyDoiTheoDVT1(vattu.MaSo, dvt, luong)
                        GhiDauKy CboKho.ItemData(CboKho.ListIndex), taikhoan.MaSo, vattu.MaSo, luong, tien, tien2
                        phatsinh = True
                        TongTien
                        RFocus txtTon(0)
                        GoTo XongDK
                    End If
                End If
            Next
        Else
            ExecuteSQL5 "INSERT INTO VTDauNam (MaSo,MaSoKho,MaTaiKhoan,MaVattu,Luong_0,Tien_0) VALUES (" + CStr(Lng_MaxValue("MaSo", "VTDauNam") + 1) + "," + CStr(CboKho.ItemData(CboKho.ListIndex)) + "," + CStr(taikhoan.MaSo) + "," + CStr(vattu.MaSo) + "," + DoiDau(luong) + "," + DoiDau(tien) + ")"
        End If
        phatsinh = True
        .AddItem taikhoan.sohieu + Chr(9) + vattu.sohieu + Chr(9) + vattu.TenVattu + Chr(9) + IIf(dvt = 0, vattu.DonVi, CboDV.Text) + Chr(9) _
            + Format(luong, Mask_2) + Chr(9) + Format(dgia, Mask_2) + Chr(9) + Format(tien, Mask_0) + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) _
            + CStr(vattu.MaSo) + Chr(9) + CStr(dvt) + IIf(pGiaUSD > 0, Chr(9) + Format(tien2, Mask_2), ""), NewRowIndex(GrdVT, 0)
        If dvt > 0 Then luong = QuyDoiTheoDVT1(vattu.MaSo, dvt, luong)
        GhiDauKy CboKho.ItemData(CboKho.ListIndex), taikhoan.MaSo, vattu.MaSo, luong, tien, tien2
        .Row = .Rows - 1
        .col = 0
        If Len(.Text) = 0 Then .RemoveItem .Row
        .Row = 0
        RFocus txtTon(0)
    End With
    TongTien
XongDK:
    Me.MousePointer = 0
End Sub

Private Sub Command_Click(Index As Integer)
    Dim thang As Integer
    
    thang = Cbo.ItemData(Cbo.ListIndex)
    Select Case Index
        Case 0:   If CboKho.ListIndex >= 0 Then XuLyChenhLech thang, CboKho.ItemData(CboKho.ListIndex)
        Case 1:   Unload Me
        Case 2:
            If f1 >= 0 Then
                ExecuteSQL5 "UPDATE KiemKe SET Luong_" + CStr(thang) + "=0,Tien_" + CStr(thang) + "=0"
                LietKeTonKho CboKho.ItemData(CboKho.ListIndex), IIf(f1 = 0, 0, thang), 1
            End If
    End Select
End Sub

Public Sub Form_Activate()
    f1 = Me.tag
    If f1 > 0 Then
        Me.Caption = IIf(pNN = 0, "KiÓm kª cuèi kú", "Inventory")
        Label(9).Visible = True
        Cbo.Visible = True
        Command(0).Visible = True
        Command(2).Visible = True
        txtTon(7).Enabled = False
        tbl = "KiemKe"
    Else
        tbl = "TonKho"
        CboKho_Click
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbAltMask) > 0 And KeyCode = vbKeyV) Or KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    ColumnSetUp GrdVT, 0, 940, 2
    ColumnSetUp GrdVT, 1, 1300, 2
    ColumnSetUp GrdVT, 2, 2380, 0
    ColumnSetUp GrdVT, 3, 700, 2
    ColumnSetUp GrdVT, 4, 1300, 1
    ColumnSetUp GrdVT, 5, 1300, 1
    ColumnSetUp GrdVT, 6, 1300, 1
    ColumnSetUp GrdVT, 7, 1, 0
    ColumnSetUp GrdVT, 8, 1, 0
    ColumnSetUp GrdVT, 9, 1, 0
    ColumnSetUp GrdVT, 10, 1260, 1
            
    AddMonthToCbo Cbo
    
    Caption = Caption + " - " + CStr(pNamTC)
    f1 = -1
    Int_RecsetToCbo "SELECT MaSo As F2,TenKho As F1 FROM KhoHang ORDER BY TenKho", CboKho
    f1 = 0
    phatsinh = False
    
    If pGiaUSD > 0 Then
        tygia = TyGiaDK
        Label(11).Caption = Label(11).Caption + "  " + Format(tygia, Mask_0)
        Me.Width = Me.Width + 1300
        Label(10).Visible = True
        Label(11).Visible = True
        GrdVT.Width = GrdVT.Width + 1300
        cmdct.Left = cmdct.Left + 1300
        txtTon(7).Visible = True
        FCenter Me
    End If
    
    SetFont Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If phatsinh Then
        Me.MousePointer = 11
        SoDuTKVT
        Me.MousePointer = 0
    End If
    Set taikhoan = Nothing
    Set vattu = Nothing
End Sub

Private Sub GrdVT_DblClick()
    Dim i As Integer, ms As Long
    
    With GrdVT
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        For i = 0 To 1
            .col = i
            txtTon(i).Text = .Text
        Next
        txtTon_LostFocus 0
        txtTon_LostFocus 1
        .col = 9
        If vattu.Dvt2 > 0 Then SetListIndex CboDV, CLng5(.Text)
        For i = 2 To 6
            .col = i
            txtTon(i).Text = .Text
        Next
        If pGiaUSD > 0 Then
            .col = 10
            txtTon(7).Text = .Text
        End If
        If OutCost <> 0 Then
            ms = SelectSQL("SELECT TOP 1 MaSo AS F1 FROM VTDauNam WHERE MaSoKho=" + CStr(CboKho.ItemData(CboKho.ListIndex)) + " AND MaTaiKhoan=" + CStr(taikhoan.MaSo) _
                + " AND MaVattu=" + CStr(vattu.MaSo) + " AND Luong_0=" + DoiDau(Cdbl5(txtTon(4).Text)) + " AND Tien_0=" + DoiDau(Cdbl5(txtTon(6).Text)))
            If ms > 0 Then ExecuteSQL5 "DELETE * FROM VTDauNam WHERE MaSo=" + CStr(ms)
            GhiDauKy CboKho.ItemData(CboKho.ListIndex), taikhoan.MaSo, vattu.MaSo, 0, 0, 0
            .RemoveItem .Row
            If .Rows < .tag Then .Rows = .tag
        End If
        
        RFocus txtTon(0)
    End With
End Sub

Private Sub GrdVT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then GrdVT_DblClick
End Sub

Private Sub GrdVT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdVT, GrdVT.col
    End If
End Sub

Private Sub laydatabase_Click()
Dim psw As String, fn As String
  Dim rs_chungtu As Recordset
  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
  Dim sql, st
  Dim i As Integer
  If Len(psw) > 0 Then
  
  sql = "insert into phanloaivattu select * from [" + psw + ";PWD=" + pPSW + "].phanloaivattu where maso  not in (select maso from phanloaivattu) "
  DBKetoan.Execute sql
   sql = "insert into dvtvattu select * from [" + psw + ";PWD=" + pPSW + "].dvtvattu where maso  not in (select maso from dvtvattu) "
  DBKetoan.Execute sql
  sql = "insert into vattu select * from [" + psw + ";PWD=" + pPSW + "].vattu where maso  not in (select maso from vattu) "
  DBKetoan.Execute sql
 ' sql = " insert into [Tonkho] select * from  [" + psw + ";PWD=" + pPSW + "].TonKho where maso not in (select maso from tonkho) "
  sql = " insert into [Tonkho] select * from  [" + psw + ";PWD=" + pPSW + "].TonKho where mavattu not in (select maso from vattu)"
  DBKetoan.Execute sql

  'sql = " Update [Tonkho] b INNER JOIN [" + psw + ";PWD=" + pPSW + "].TonKho a ON b.maso = a.maso "
   sql = " Update [Tonkho] b INNER JOIN [" + psw + ";PWD=" + pPSW + "].TonKho a ON b.mavattu = a.mavattu and b.MaSoKho = a.MaSoKho and   b.MaTaiKhoan = a.MaTaiKhoan"
        sql = sql + " SET b.tien_0 = a.tien_12"
        sql = sql + " ,b.Luong_0 = a.Luong_12"
        For i = 1 To 12
        st = CStr(i)
        sql = sql + ",b.Luong_" + st + " = a.Luong_12, b.Tien_" + st + " = a.Tien_12"
      Next
  DBKetoan.Execute sql
  CboKho_Click
  MsgBox "B¹n ®· chuyÓn d÷ liÖu ®Çu kú thµnh c«ng."
  Dim sttt As String
 ' ' sttt = SelectSQL("select mataikhoan as f1 from [" + psw + ";PWD=" + pPSW + "].Tonkho where maso not in (select maso from tonkho)")
  ' sttt = SelectSQL("select sohieu as f1 from [" + psw + ";PWD=" + pPSW + "].vattu where maso = 1510")
  ' sttt = SelectSQL("select sohieu as f1 from vattu where maso = 1510")
 '  sttt = SelectSQL("select ten as f1 from hethongtk where maso = 0")
  '  MsgBox sttt
  End If
End Sub

Private Sub loadexel_Click()
 frmTonDauSanPham.Show 1
 CboKho_Click
End Sub

Private Sub txtTon_GotFocus(Index As Integer)
    AutoSelect txtTon(Index)
End Sub

Private Sub txtTon_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0:
            If KeyAscii = 13 Then
                txtTon(0).Text = FrmTaikhoan.ChonTk(txtTon(0).Text)
            End If
        Case 1:
            If KeyAscii = 13 Then
                txtTon(1).Text = FrmVattu.ChonVattu(txtTon(1).Text)
            End If
        Case 2, 3:
            KeyAscii = 0
        Case 4, 5, 6:
            If Index = 6 And KeyAscii = 13 Then
                CmdCt_Click
            Else
                KeyProcess txtTon(Index), KeyAscii, True
            End If
    End Select
End Sub

Public Sub txtTon_LostFocus(Index As Integer)
    Dim luong As Double
    
    Select Case Index
        Case 0:
            If Len(txtTon(0).Text) > 0 Then
                taikhoan.InitTaikhoanSohieu txtTon(0).Text
            Else
                taikhoan.InitTaikhoanMaSo 0
            End If
        Case 1:
            If Len(txtTon(1).Text) > 0 Then
                vattu.InitVattuSohieu txtTon(1).Text
            Else
                vattu.InitVattuMaSo 0
            End If
            txtTon(2).Text = vattu.TenVattu
            txtTon(3).Text = vattu.DonVi
            If vattu.GiaHT > 0 Then txtTon(5).Text = Format(vattu.GiaHT, Mask_2)
            
            
            If vattu.Dvt2 > 0 Then
                Int_RecsetToCbo "SELECT MaSo AS F2, DonVi AS F1 FROM DVTVattu WHERE MaVattu=" + CStr(vattu.MaSo) + " ORDER BY DonVi", CboDV
                CboDV.AddItem vattu.DonVi, 0
                CboDV.ListIndex = 0
                CboDV.Visible = True
                RFocus CboDV
            Else
                CboDV.Visible = False
            End If
        Case 4:
            txtTon(4).Text = Format(txtTon(4).Text, Mask_2)
            txtTon(6).Text = Format(RoundMoney(Cdbl5(txtTon(4).Text) * Cdbl5(txtTon(5).Text)), Mask_0)
        Case 5:
            txtTon(5).Text = Format(txtTon(5).Text, Mask_2)
            txtTon(6).Text = Format(RoundMoney(Cdbl5(txtTon(4).Text) * Cdbl5(txtTon(5).Text)), Mask_0)
        Case 6:
            luong = Cdbl5(txtTon(4).Text)
            If luong <> 0 Then
                txtTon(5).Text = Format(Cdbl5(txtTon(6).Text) / luong, Mask_2)
            Else
                txtTon(5).Text = "0"
            End If
        Case 7:
            txtTon(7).Text = Format(txtTon(7).Text, Mask_2)
    End Select
End Sub
'======================================================================================
' Thñ tôc liÖt kª tån kho
'======================================================================================
Private Sub LietKeTonKho(mkho As Long, thang As Integer, lk As Integer)
    Dim rs_ton As Recordset, dgia As Double, tb As String, t2 As Double
        
    Me.MousePointer = 11
    If lk = 1 Then tb = "TonKho" Else tb = IIf(f1 = 0, IIf(OutCost <> 0, "VTDauNam", "TonKho"), "KiemKe")
    ClearGrid GrdVT, GrdVT.tag
    Set rs_ton = DBKetoan.OpenRecordset("SELECT HethongTK.MaSo,HethongTK.SoHieu AS SHTK,MaVattu,Vattu.SoHieu,Vattu.TenVattu,Vattu.DonVi," _
        & " Luong_" + CStr(CThangDB(thang)) + " AS Luong, Tien_" + CStr(CThangDB(thang)) + " AS Tien FROM (" + tb + " INNER JOIN Vattu ON " + tb + ".MaVattu = Vattu.MaSo) INNER JOIN HethongTK ON " + tb + ".MaTaiKhoan=HethongTK.MaSo" _
        & " WHERE MaSoKho = " + CStr(mkho) + " AND (Luong_" + CStr(CThangDB(thang)) + "<> 0 OR Tien_" + CStr(CThangDB(thang)) + "<>0) ORDER BY HethongTK.SoHieu DESC, Vattu.SoHieu DESC", dbOpenSnapshot)
    Do While Not rs_ton.EOF
        If rs_ton!luong <> 0 Then dgia = Fix(0.5 + Mask_N * rs_ton!tien / rs_ton!luong) / Mask_N Else dgia = 0
        If pGiaUSD > 0 Then t2 = SelectSQL("SELECT USDTien_" + CStr(CThangDB(thang)) + " AS F1 FROM TonKho WHERE MaTaiKhoan=" + CStr(rs_ton!MaSo) + " AND MaSoKho=" + CStr(mkho) + " AND MaVattu=" + CStr(rs_ton!MaVattu))
        GrdVT.AddItem rs_ton!shtk + Chr(9) + rs_ton!sohieu + Chr(9) + rs_ton!TenVattu + Chr(9) + rs_ton!DonVi + Chr(9) _
            + Format(rs_ton!luong, Mask_2) + Chr(9) + Format(dgia, Mask_2) + Chr(9) + Format(rs_ton!tien, Mask_0) + Chr(9) + CStr(rs_ton!MaSo) + Chr(9) + CStr(rs_ton!MaVattu) + Chr(9) + "" + Chr(9) + Format(t2, Mask_2), 0
        If lk = 1 Then GhiDauKy mkho, rs_ton!MaSo, rs_ton!MaVattu, rs_ton!luong, rs_ton!tien, t2
        rs_ton.MoveNext
    Loop
    GrdVT.Rows = IIf(rs_ton.RecordCount > GrdVT.tag, rs_ton.RecordCount, GrdVT.tag)
    rs_ton.Close
    Set rs_ton = Nothing
    GrdVT.Row = 0
    TongTien
    Me.MousePointer = 0
End Sub
'======================================================================================
' Thñ tôc ghi sè tån kho ®Çu kú
'======================================================================================
Public Sub GhiDauKy(mkho As Long, mtk As Long, mvt As Long, luong As Double, tien As Double, tien2 As Double)
    Dim rs_ton As Recordset, sql As String, i As Integer, sql2 As String, st1 As String, st2 As String, st3 As String, thang As Integer, sql3 As String, sql4 As String
    
    If OutCost <> 0 And f1 = 0 Then
        tien = SelectSQL("SELECT Sum(Tien_0) AS F1, Sum(Luong_0) AS F2" + IIf(pGiaUSD > 0, ", Sum(USDTien_0) AS F3", "") + " FROM VTDauNam WHERE MaSoKho=" + CStr(mkho) + " AND MaTaiKhoan=" + CStr(mtk) + " AND MaVattu=" + CStr(mvt), luong, tien2)
    End If
    thang = IIf(f1 = 0, 0, Cbo.ItemData(Cbo.ListIndex))
    
    Set rs_ton = DBKetoan.OpenRecordset("SELECT Luong_" + CStr(CThangDB(thang)) + " AS Luong, Tien_" + CStr(CThangDB(thang)) + " AS Tien" + IIf(pGiaUSD > 0 And f1 = 0, ",USDTien_" + CStr(CThangDB(thang)) + " AS Tien2", "") + " FROM " + tbl + " INNER JOIN Vattu ON " + tbl + ".MaVattu = Vattu.MaSo WHERE MaSoKho=" + CStr(mkho) + " AND MaTaiKhoan=" + CStr(mtk) + " AND MaVattu=" + CStr(mvt), dbOpenSnapshot)
    WSpace.BeginTrans
    If rs_ton.RecordCount = 0 Then
        st1 = DoiDau(luong)
        st2 = DoiDau(tien)
        st3 = DoiDau(tien2)
        
        For i = CThangDB(thang) To 12
            sql = sql + ", Luong_" + CStr(i) + ", Tien_" + CStr(i)
            sql2 = sql2 + "," + st1 + "," + st2
            sql3 = sql3 + ", USDTien_" + CStr(i)
            sql4 = sql4 + "," + st3
        Next
        
        ExecuteSQL5 "INSERT INTO " + tbl + " (MaSo,MaSoKho,MaTaiKhoan,MaVattu" + sql + IIf(pGiaUSD > 0 And f1 = 0, sql3, "") + ") VALUES (" + CStr(Lng_MaxValue("MaSo", tbl) + 1) + "," + CStr(mkho) _
            + "," + CStr(mtk) + "," + CStr(mvt) + sql2 + IIf(pGiaUSD > 0 And f1 = 0, sql4, "") + ")"
    Else
        st1 = DoiDau(luong - rs_ton!luong)
        st2 = DoiDau(tien - rs_ton!tien)
        If pGiaUSD > 0 And f1 = 0 Then st3 = DoiDau(tien2 - rs_ton!tien2)
        sql = "Luong_" + CStr(CThangDB(thang)) + "  = " + DoiDau(luong) + ", Tien_" + CStr(CThangDB(thang)) + "  = " + DoiDau(tien)
        If pGiaUSD > 0 And f1 = 0 Then sql = sql + ", USDTien_" + CStr(CThangDB(thang)) + "=" + DoiDau(tien2)
        For i = CThangDB(thang) + 1 To 12
            sql2 = CStr(i)
            sql = sql + ", Luong_" + sql2 + " = Luong_" + sql2 + " + (" + st1 + "), Tien_" + sql2 + " = Tien_" + sql2 + " + " + st2
            If pGiaUSD > 0 And f1 = 0 Then sql = sql + ", USDTien_" + sql2 + "=USDTien_" + sql2 + " + " + st3
        Next
        ExecuteSQL5 "UPDATE " + tbl + " SET " + sql + " WHERE MaSoKho=" + CStr(mkho) + " AND MaTaiKhoan=" + CStr(mtk) + " AND MaVattu=" + CStr(mvt)
    End If
    WSpace.CommitTrans
    rs_ton.Close
    Set rs_ton = Nothing
End Sub

Private Sub TongTien()
    Dim thang As Integer, T As Double, tb As String
    
    tb = IIf(f1 = 0, IIf(OutCost <> 0, "VTDauNam", "TonKho"), "KiemKe")
    thang = IIf(f1 = 0, 0, Cbo.ItemData(Cbo.ListIndex))

    T = SelectSQL("SELECT Sum(Tien_" + CStr(CThangDB(thang)) + ") As F1 FROM " + tb + " WHERE MaSoKho = " + CStr(CboKho.ItemData(CboKho.ListIndex)))
    LbTien(0).Caption = Format(T, Mask_0)
    If pGiaUSD > 0 Then
        T = SelectSQL("SELECT Sum(USDTien_" + CStr(CThangDB(thang)) + ") As F1 FROM TonKho WHERE MaSoKho = " + CStr(CboKho.ItemData(CboKho.ListIndex)))
        LbTien(1).Caption = Format(T, Mask_2)
    End If
End Sub

Private Sub XuLyChenhLech(thang As Integer, mkho As Long)
    Dim rs As Recordset, luong As Double, tien As Double, i As Integer
    
    Load FrmChungtu
    FrmChungtu.OptLoai(2).Value = True
    SetListIndex FrmChungtu.CboNguon(1), mkho
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaTaiKhoan, MaVattu, HethongTK.SoHieu AS SHTK, Vattu.SoHieu AS SHVT, Luong_" + CStr(CThangDB(thang)) + " AS Luong, Tien_" + CStr(CThangDB(thang)) + " AS Tien FROM " _
        & " (TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo) INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo WHERE MaSoKho=" + CStr(mkho) + " ORDER BY HethongTK.SoHieu DESC, Vattu.SoHieu DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        luong = SelectSQL("SELECT Luong_" + CStr(CThangDB(thang)) + " AS F1, Tien_" + CStr(CThangDB(thang)) + " AS F2 FROM KiemKe WHERE MaSoKho=" + CStr(mkho) + " AND MaTaiKhoan=" + CStr(rs!MaTaiKhoan) + " AND MaVattu=" + CStr(rs!MaVattu), tien)
        If luong <> rs!luong Or tien <> rs!tien Then
            FrmChungtu.txtchungtu(0).Text = rs!shtk
            FrmChungtu.txtchungtu(2).Text = rs!shvt
            FrmChungtu.txtChungtu_LostFocus 0
            FrmChungtu.txtChungtu_LostFocus 2
            FrmChungtu.txtchungtu(3).Text = Format(rs!luong - luong, Mask_2)
            FrmChungtu.txtchungtu(5).Text = ""
            FrmChungtu.txtchungtu(6).Text = Format(rs!tien - tien, Mask_2)
        
            FrmChungtu.CmdChitiet_Click
            i = i + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If i > 0 Then
        pFunction = 10
        FrmChungtu.Show 1
    End If
End Sub

Public Sub xoa_Click()
Dim sql As String
 If MsgBox("B¹n cã muèn xãa toµn bé tån ®Çu kh«ng ?", vbYesNo + vbCritical, App.ProductName) = vbYes Then
 sql = " Update [Tonkho] b  "
        sql = sql + " SET b.tien_0 = 0"
        sql = sql + " ,b.Luong_0 = 0"
        DBKetoan.Execute sql
        CboKho_Click
            DBKetoan.Execute "update hethongtk set duno_0 = 0,duco_0 = 0 where sohieu like '156*'"
       DBKetoan.Execute "update hethongtk set duno_0 = 0,duco_0 = 0 where sohieu like '154*'"
       DBKetoan.Execute "update hethongtk set duno_0 = 0,duco_0 = 0 where sohieu like '152*'"

   '     MsgBox "Xãa thµnh c«ng"
End If
End Sub
