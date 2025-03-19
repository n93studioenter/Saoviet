VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmLuuChuyen 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "L≠u chuy”n nÈi bÈ"
   ClientHeight    =   7080
   ClientLeft      =   180
   ClientTop       =   780
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
   Icon            =   "Flchuyen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Inventory Internal Flow"
   Begin VB.CommandButton Command 
      Caption         =   "&Barcode"
      Height          =   375
      Index           =   5
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   48
      Tag             =   "&Print"
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox CboVV 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   7320
      Style           =   2  'Dropdown List
      TabIndex        =   42
      ToolTipText     =   "Nh n chuÈt ph∂i Æ” Æ®ng k˝"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox CboVV 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   44
      ToolTipText     =   "Nh n chuÈt ph∂i Æ” Æ®ng k˝"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox CboVV 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   43
      ToolTipText     =   "Nh n chuÈt ph∂i Æ” Æ®ng k˝"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtChuyen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   9480
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox CboDV 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Flchuyen.frx":57E2
      Left            =   3840
      List            =   "Flchuyen.frx":57E4
      Style           =   2  'Dropdown List
      TabIndex        =   9
      ToolTipText     =   "Danh s∏ch Æ¨n vﬁ t›nh"
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtChuyen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   6120
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      ItemData        =   "Flchuyen.frx":57E6
      Left            =   960
      List            =   "Flchuyen.frx":57E8
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txtChuyen 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   6120
      MaxLength       =   150
      TabIndex        =   4
      Text            =   "..."
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtChuyen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   8520
      MaxLength       =   20
      TabIndex        =   14
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   4
      Left            =   3600
      Picture         =   "Flchuyen.frx":57EA
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "&Print"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtChuyen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   5760
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   12
      Tag             =   "0"
      Top             =   6240
      Width           =   1215
   End
   Begin MSGrid.Grid GrdChuyen 
      Height          =   4095
      Left            =   120
      TabIndex        =   22
      Tag             =   "60"
      Top             =   2160
      Width           =   9615
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   7223
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
      Cols            =   15
      FixedRows       =   0
      ScrollBars      =   2
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   8400
      Picture         =   "Flchuyen.frx":6C4C
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   7200
      Picture         =   "Flchuyen.frx":806E
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "&Delete"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   6000
      Picture         =   "Flchuyen.frx":9550
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "&Save"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   4800
      Picture         =   "Flchuyen.frx":A97E
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "&Add"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdct 
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
      Left            =   9480
      TabIndex        =   16
      Top             =   6240
      Width           =   255
   End
   Begin VB.ComboBox CboSohieu 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtChuyen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   4800
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   11
      Tag             =   "0"
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtChuyen 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4080
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox txtChuyen 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox txtChuyen 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   7
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtChuyen 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   20
      TabIndex        =   6
      Top             =   6240
      Width           =   975
   End
   Begin VB.ComboBox CboKho 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   6240
      Width           =   1575
   End
   Begin VB.ComboBox CboKho 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VÀt t≠"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   34
      Tag             =   "Inventory"
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SË l≠Óng"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   6
      Left            =   4800
      TabIndex        =   28
      Tag             =   "Quantity"
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thµnh ti“n"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   7
      Left            =   5760
      TabIndex        =   29
      Tag             =   "Amount"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label LbTT 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Th´ng tin 1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   47
      Tag             =   "Index"
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LbTT 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Th´ng tin 2"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   46
      Tag             =   "Index"
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LbTT 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Th´ng tin 3"
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   45
      Tag             =   "Index"
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ng≠Íi nhÀp"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   41
      Tag             =   "Input by"
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label LbUser 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   1200
      TabIndex        =   40
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "T˚ gi∏"
      Height          =   255
      Index           =   15
      Left            =   4800
      TabIndex        =   39
      Tag             =   "Ex. Rate"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Th∏ng"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   38
      Tag             =   "Month"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Di‘n gi∂i"
      Height          =   255
      Index           =   13
      Left            =   5040
      TabIndex        =   37
      Tag             =   "Description"
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "TÊng sË ti“n"
      Height          =   255
      Index           =   11
      Left            =   4800
      TabIndex        =   32
      Tag             =   "Total"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kho nhÀp"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   12
      Left            =   6960
      TabIndex        =   36
      Tag             =   "Import to Store"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tµi kho∂n ghi nÓ"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   10
      Left            =   8520
      TabIndex        =   35
      Tag             =   "Debit Account"
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5715
      TabIndex        =   33
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   SË hi÷u   Tµi kho∂n"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   31
      Tag             =   "Account"
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SË hi÷u"
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   30
      Tag             =   "V. Code"
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ß.vﬁ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   27
      Tag             =   "Unit"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Di‘n gi∂i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   26
      Tag             =   "Description"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SË hi÷u"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   25
      Tag             =   "Code"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Kho xu t"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Tag             =   "Store"
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ngµy"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Tag             =   "Date"
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "FrmLuuChuyen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ngay As Date
Dim taikhoan As New ClsTaikhoan
Dim taikhoan1 As New ClsTaikhoan
Dim vattu As New ClsVattu
Dim MaNhap As Long
Dim flag As Integer

Dim bcstop As Integer

Private Sub Cbo_Click()
    Int_RecsetToCbo "SELECT DISTINCTROW MaCT As F2,SoHieu As F1 FROM ChungTu WHERE MaLoai=4 AND ThangCT=" + CStr(Cbo.ItemData(Cbo.ListIndex)) + " GROUP BY MaCT, SoHieu, NgayCT ORDER BY NgayCT", CboSohieu
    XoaPhieuTrenManHinh
End Sub

Private Sub CboDV_Click()
    Dim dvt As Long, luong As Double, tien As Double, tien2 As Double
    
    If OutCost <> 0 Or vattu.MaSo = 0 Or vattu.Dvt2 = 0 Then Exit Sub
    dvt = CboDV.ItemData(CboDV.ListIndex)
    luong = SoTonKho(Month(ngay), CboKho(0).ItemData(CboKho(0).ListIndex), taikhoan.MaSo, vattu.MaSo, tien, tien2)
    If dvt = 0 Then
        txtChuyen(4).tag = luong
    Else
        txtChuyen(4).tag = QuyDoiTheoDVT2(vattu.MaSo, dvt, luong)
    End If
    txtChuyen(4).Text = Format(txtChuyen(4).tag, Mask_2)
End Sub

'====================================================================================================
' Chu»n bﬁ phi’u l≠u chuy”n
'====================================================================================================
Private Sub CboKho_Click(Index As Integer)
    If Index = 0 Then
        ClearGrid GrdChuyen, GrdChuyen.tag
        If CboSohieu.ListIndex >= 0 Then
            CboSohieu.ListIndex = -1
            CboSohieu.Text = ""
        End If
    End If
End Sub
'====================================================================================================
' Hi”n thﬁ phi’u Æ≠Óc ch‰n
'====================================================================================================
Private Sub CboSohieu_Click()
    Dim rs_chungtu As Recordset, luong As Double, tien As Double
    
    If CboSohieu.ListIndex < 0 Then Exit Sub
    ClearGrid GrdChuyen, GrdChuyen.tag
    Set rs_chungtu = DBKetoan.OpenRecordset("SELECT NgayCT,DienGiai,MaKho,MaNguon, MaTKNo, MaTKCo,TK.SoHieu AS SHTK,MaVattu,Vattu.SoHieu AS SHVT,TenVattu,DonVi,SoPS,SoPS2Co,CT_ID,HethongTK.SoHieu AS SHNo,User_ID,DVT,TyGia" + IIf(pGiaUSD > 0, ",PSUSD", "") + ",MaDT1,MaDT2,MaDT3 FROM (" + ChungTu2TKNC(0) + ") INNER JOIN Vattu ON Chungtu.MaVattu=Vattu.MaSo" _
        & " WHERE MaCT=" + CStr(CboSohieu.ItemData(CboSohieu.ListIndex)), dbOpenSnapshot)
    ngay = rs_chungtu!NgayCT
    MedNgay.Text = Format(ngay, Mask_D)
    txtChuyen(7).Text = rs_chungtu!diengiai
    SetListIndex CboKho(0), rs_chungtu!MaKho
    LbUser.Caption = TenUser(rs_chungtu!User_ID)
    If pTygia > 0 Then txtChuyen(8).Text = Format(rs_chungtu!tygia, Mask_0)
    
    If pSoVV > 0 And rs_chungtu!MaDT1 > 0 Then SetListIndex CboVV(0), rs_chungtu!MaDT1
    If pSoVV > 1 And rs_chungtu!MaDT2 > 0 Then SetListIndex CboVV(1), rs_chungtu!MaDT2
    If pSoVV > 2 And rs_chungtu!MaDT3 > 0 Then SetListIndex CboVV(2), rs_chungtu!MaDT3
    
    Do While Not rs_chungtu.EOF
        SetListIndex CboKho(1), rs_chungtu!MaNguon
        
        luong = IIf(rs_chungtu!MaVattu > 0 And rs_chungtu!dvt > 0, QuyDoiTheoDVT2(rs_chungtu!MaVattu, rs_chungtu!dvt, rs_chungtu!SoPS2Co), rs_chungtu!SoPS2Co)
        If pGiaUSD > 0 Then tien = rs_chungtu!PSUSD Else tien = 0
        
        GrdChuyen.AddItem rs_chungtu!shtk + Chr(9) + rs_chungtu!shvt + Chr(9) + rs_chungtu!TenVattu + Chr(9) + IIf(rs_chungtu!dvt = 0, rs_chungtu!DonVi, TenDVT(rs_chungtu!dvt)) + Chr(9) + Format(luong, Mask_2) _
            + Chr(9) + Format(rs_chungtu!sops, Mask_0) + Chr(9) + CboKho(1).Text + Chr(9) + rs_chungtu!shno + Chr(9) + CStr(rs_chungtu!MaTkCo) + Chr(9) + CStr(rs_chungtu!MaVattu) _
            + Chr(9) + CStr(rs_chungtu!CT_ID) + Chr(9) + CStr(rs_chungtu!MaNguon) + Chr(9) + CStr(rs_chungtu!MaTkNo) + Chr(9) + CStr(rs_chungtu!dvt) + Chr(9) + Format(tien, Mask_2), 0
        rs_chungtu.MoveNext
    Loop
    GrdChuyen.Rows = IIf(rs_chungtu.RecordCount > GrdChuyen.tag, rs_chungtu.RecordCount, GrdChuyen.tag)
    TongTien
    rs_chungtu.Close
    Set rs_chungtu = Nothing
    vattu.InitVattuMaSo 0
    taikhoan.InitTaikhoanMaSo 0
End Sub
'====================================================================================================
' Th™m, sˆa, x„a phi’u l≠u chuy”n
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim chungtu As New ClsChungtu, sh As String, shvt As String, tkno As String, ttien As Double, dvt As Long, tien2 As Double, ttien2 As Double
    
    Me.MousePointer = 11
    Select Case Index
        Case 0:
            XoaPhieuTrenManHinh
            If CboSohieu.Visible Then RFocus CboSohieu
            GrdChuyen.Row = 0
        Case 1:
            Dim MaCT As Long, i As Integer, mvt As Long, luong As Double, tien As Double, mtk As Long, mk As Long, mtk1 As Long
            
            If Not KiemTraPhieu Then GoTo XongLC
            
            MaCT = Lng_MaxValue("MaCT", "ChungTu") + 1
            With GrdChuyen
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 8
                    If Len(.Text) = 0 Then Exit For
                    mtk = CLng5(.Text)
                    .col = 9
                    mvt = CLng5(.Text)
                    .col = 11
                    mk = CLng5(.Text)
                    .col = 12
                    mtk1 = CLng5(.Text)
                    .col = 13
                    dvt = CLng5(.Text)
                    .col = 4
                    luong = Cdbl5(.Text)
                    .col = 5
                    tien = Cdbl5(.Text)
                    .col = 14
                    tien2 = Cdbl5(.Text)
                    If dvt > 0 Then luong = QuyDoiTheoDVT1(mvt, dvt, luong)
                    chungtu.InitChungtu 0, 4, CboSohieu.Text, Cbo.ItemData(Cbo.ListIndex), ngay, ngay, mk, _
                            CboKho(0).ItemData(CboKho(0).ListIndex), txtChuyen(7).Text, mtk1, mtk, tien, luong, luong, mvt, "...", 1, "", "", "", ""
                    If pGiaUSD > 0 Then chungtu.PSUSD = tien2
                    chungtu.MaCT = MaCT
                    .col = 10
                    chungtu.CT_ID = -CLng5(.Text)
                    If pTygia > 0 Then chungtu.tygia = Cdbl5(txtChuyen(8).Text)
                    chungtu.User_ID = UserID
                    chungtu.dvt = dvt
                    If pSoVV > 0 And CboVV(0).ListIndex >= 0 Then chungtu.MaDT1 = CboVV(0).ItemData(CboVV(0).ListIndex)
                    If pSoVV > 1 And CboVV(1).ListIndex >= 0 Then chungtu.MaDT2 = CboVV(1).ItemData(CboVV(1).ListIndex)
                    If pSoVV > 2 And CboVV(2).ListIndex >= 0 Then chungtu.MaDT3 = CboVV(2).ItemData(CboVV(2).ListIndex)
                    chungtu.GhiChungtu
                Next
            End With
            sh = CboSohieu.Text
            Command_Click 2
            CboSohieu.AddItem sh
            CboSohieu.ItemData(CboSohieu.NewIndex) = MaCT
        Case 2:
            If CboSohieu.ListIndex >= 0 Then
                Dim rs_chungtu As Recordset
                Set rs_chungtu = DBKetoan.OpenRecordset("SELECT MaSo FROM ChungTu WHERE MaCT=" + CStr(CboSohieu.ItemData(CboSohieu.ListIndex)), dbOpenSnapshot)
                Do While Not rs_chungtu.EOF
                    chungtu.InitChungtu rs_chungtu!MaSo, 0, "", 0, ngay, ngay, 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
                    chungtu.XoaChungtu
                    rs_chungtu.MoveNext
                Loop
                rs_chungtu.Close
                Set rs_chungtu = Nothing
                CboSohieu.RemoveItem CboSohieu.ListIndex
            End If
            XoaPhieuTrenManHinh
            GrdChuyen.Row = 0
        Case 3:
            Unload Me
        Case 4:
            SetRptInfo
            ExecuteSQL5 "DELETE * FROM PhieuNX"
            With GrdChuyen
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 0
                    sh = .Text
                    If Len(sh) = 0 Then Exit For
                    .col = 1
                    shvt = .Text
                    .col = 7
                    tkno = .Text
                    .col = 4
                    luong = Cdbl5(.Text)
                    .col = 5
                    tien = Cdbl5(.Text)
                    ttien = ttien + tien
                    .col = 14
                    tien2 = Cdbl5(.Text)
                    ttien2 = ttien2 + tien2
                    .col = 6
                    ExecuteSQL5 "INSERT INTO PhieuNX (MaSo,SoCT,DienGiaiCT,SoHieu,DienGiai,SoLuong,ThanhTien,ThanhTien2) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhieuNX") + 1) + ",'" + sh + "','" + .Text _
                        + "','" + shvt + "','" + tkno + "'," + DoiDau(luong) + "," + DoiDau(tien) + "," + DoiDau(tien2) + ")"
                Next
            End With

            frmMain.Rpt.ReportFileName = "PHIEULC.RPT"
            frmMain.Rpt.Formulas(4) = "Kho='" + CboKho(0).Text + "'"
            frmMain.Rpt.Formulas(5) = "SHPhieu='" + CboSohieu.Text + "'"
            frmMain.Rpt.Formulas(6) = "Sotien='" + ToVNText(ttien) + " VND'"
            frmMain.Rpt.Formulas(7) = "Ngay='" + Format(ngay, Mask_DR) + "'"
            frmMain.Rpt.Formulas(8) = "LyDo='" + txtChuyen(7).Text + "'"
            InBaoCaoRPT
        Case 5:
            If CboKho(0).ListCount = 1 Then GoTo XongLC
            
            If CboKho(1).ListIndex = CboKho(0).ListIndex Then
                MsgBox "Ch‰n kho nhÀp", vbCritical, App.ProductName
                RFocus CboKho(1)
                GoTo XongLC
            End If
    
            LayXuatKho
    End Select
XongLC:
    Set chungtu = Nothing
    Me.MousePointer = 0
End Sub
'====================================================================================================
' Ghi chi ti’t l≠u chuy”n
'====================================================================================================
Private Sub CmdCt_Click()
    Dim luong As Double, tien As Double, i As Integer, dvt As Long, tien2 As Double
    
    If (taikhoan.MaSo = 0 Or taikhoan.tk_id <> TKVT_ID) Then                ' Or (taikhoan.MaSo <> mtkc And mtkc > 0)
        RFocus txtChuyen(0)
        Exit Sub
    End If
    
    If (taikhoan1.MaSo = 0 Or (Left(taikhoan1.sohieu, 1) <> "6") And taikhoan1.tk_id <> TKVT_ID) Then              ' Or (taikhoan.MaSo <> mtkc And mtkc > 0)
        RFocus txtChuyen(6)
        Exit Sub
    End If
    
    If CboKho(0).ListIndex < 0 Then Exit Sub
    
    If vattu.MaSo = 0 Then
        RFocus txtChuyen(1)
        Exit Sub
    End If
    
    If CboKho(1).ListIndex = CboKho(0).ListIndex And taikhoan1.tk_id = TKVT_ID Then
        RFocus CboKho(1)
        Exit Sub
    End If
    
    luong = Cdbl5(txtChuyen(4).Text)
    tien2 = Cdbl5(txtChuyen(9).Text)
    tien = RoundMoney(Cdbl5(txtChuyen(5).Text))
    If luong = 0 Then
        MsgBox "H∑y nhÀp sË l≠Óng c«n l≠u chuy”n !", vbExclamation, App.ProductName
        RFocus txtChuyen(4)
        Exit Sub
    End If
    
    If luong > txtChuyen(4).tag Then
        If MsgBox("Xu t qu∏ l≠Óng tÂn kho, ti’p tÙc ?", vbCritical + vbYesNo, App.ProductName) = vbNo Then
            RFocus txtChuyen(4)
            Exit Sub
        End If
    End If
    
    If OutCost = 0 And vattu.Dvt2 > 0 Then
        dvt = CboDV.ItemData(CboDV.ListIndex)
    Else
        dvt = 0
    End If
    
    If Not (taikhoan.loai = 0 Or OutCost > 0) Then
        With GrdChuyen
            For i = 0 To .Rows - 1
                .col = 8
                .Row = i
                If Len(.Text) = 0 Then Exit For
                If CLng5(.Text) = taikhoan.MaSo Then
                    .col = 9
                    If CLng5(.Text) = vattu.MaSo Then
                        .col = 11
                        If CLng5(.Text) = CboKho(1).ItemData(CboKho(1).ListIndex) Then
                            .col = 12
                            If CLng5(.Text) = taikhoan1.MaSo Then
                                .RemoveItem i
                                .AddItem taikhoan.sohieu + Chr(9) + vattu.sohieu + Chr(9) + vattu.TenVattu + Chr(9) + IIf(dvt = 0, vattu.DonVi, CboDV.Text) + Chr(9) + Format(luong, Mask_2) + Chr(9) + Format(tien, Mask_0) + Chr(9) + CboKho(1).Text + Chr(9) + taikhoan1.sohieu _
                                    + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) + CStr(vattu.MaSo) + Chr(9) + CStr(MaNhap) + Chr(9) + CStr(CboKho(1).ItemData(CboKho(1).ListIndex)) + Chr(9) + CStr(taikhoan1.MaSo) + Chr(9) + CStr(dvt) + IIf(pGiaUSD > 0, Chr(9) + Format(tien2, Mask_2), ""), i
                                GoTo ABC
                            End If
                        End If
                    End If
                End If
            Next
        End With
    End If
    
    GrdChuyen.AddItem taikhoan.sohieu + Chr(9) + vattu.sohieu + Chr(9) + vattu.TenVattu + Chr(9) + IIf(dvt = 0, vattu.DonVi, CboDV.Text) + Chr(9) + Format(luong, Mask_2) + Chr(9) + Format(tien, Mask_0) + Chr(9) + CboKho(1).Text + Chr(9) + taikhoan1.sohieu _
        + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) + CStr(vattu.MaSo) + Chr(9) + CStr(MaNhap) + Chr(9) + CStr(CboKho(1).ItemData(CboKho(1).ListIndex)) + Chr(9) + CStr(taikhoan1.MaSo) + Chr(9) + CStr(dvt) + IIf(pGiaUSD > 0, Chr(9) + Format(tien2, Mask_2), ""), NewRowIndex(GrdChuyen, 0)
        
    GrdChuyen.Row = GrdChuyen.Rows - 1
    GrdChuyen.col = 5
    If Len(GrdChuyen.Text) = 0 Then GrdChuyen.RemoveItem GrdChuyen.Row
    GrdChuyen.Row = 0
    GrdChuyen.col = 0
    TongTien
    'mtkc = taikhoan.MaSo
ABC:
    For i = 0 To 3
        txtChuyen(i).Text = ""
    Next
    vattu.InitVattuMaSo 0
    taikhoan.InitTaikhoanMaSo 0
    MaNhap = 0
    RFocus txtChuyen(0)
End Sub
'====================================================================================================
' Xˆ l˝ c∏c ph›m n„ng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (Shift And vbAltMask) > 0 Then
        i = -1
        Select Case KeyCode
            Case vbKeyT: i = 0
            Case vbKeyG: i = 1
            Case vbKeyX:  i = 2
            Case vbKeyV:  i = 3
            Case vbKeyI:  i = 4
        End Select
        If i >= 0 Then
            If Command(i).Enabled Then
                RFocus Command(i)
                Command_Click i
            End If
        End If
    End If
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyP Then frmMain.mnHT_Click 8
    If KeyCode = vbKeyEscape Then
        bcstop = 1
        Unload Me
    End If
End Sub
'====================================================================================================
' KhÎi tπo cˆa sÊ nhÀp
'====================================================================================================
Private Sub Form_Load()
    Dim chi_so As Integer
    
    ColumnSetUp GrdChuyen, 0, 940, 0
    ColumnSetUp GrdChuyen, 1, 1180, 0
    ColumnSetUp GrdChuyen, 2, 1780, 0
    ColumnSetUp GrdChuyen, 3, 700, 0
    ColumnSetUp GrdChuyen, 4, 940, 1
    ColumnSetUp GrdChuyen, 5, 1180, 1
    ColumnSetUp GrdChuyen, 6, 1540, 0
    ColumnSetUp GrdChuyen, 7, 940, 0
    ColumnSetUp GrdChuyen, 8, 1, 0
    ColumnSetUp GrdChuyen, 9, 1, 0
    ColumnSetUp GrdChuyen, 10, 1, 0
    ColumnSetUp GrdChuyen, 11, 1, 0
    ColumnSetUp GrdChuyen, 12, 1, 0
    ColumnSetUp GrdChuyen, 13, 1, 0
    ColumnSetUp GrdChuyen, 14, 1180, 1
    InitDateVars MedNgay, ngay
        
    Int_RecsetToCbo "SELECT MaSo As F2,TenKho As F1 FROM KhoHang ORDER BY TenKho", CboKho(0)
    Int_RecsetToCbo "SELECT MaSo As F2,TenKho As F1 FROM KhoHang ORDER BY TenKho", CboKho(1)
    
    'Int_RecsetToCbo "SELECT MaSo As F2,SoHieu As F1 FROM HethongTK WHERE TKCon=0 AND TK_ID=" + CStr(TKVT_ID) + " ORDER BY SoHieu", CboTK
    Command_Click 0
    
    Caption = Caption + " - " + CStr(pNamTC)
    'mtkc = 0
    
    'txtChuyen(5).TabStop = Not FCost
    
    AddMonthToCbo Cbo
    
    Label(15).Visible = (pTygia > 0)
    txtChuyen(8).Visible = (pTygia > 0)
    If pTygia > 0 Then txtChuyen(8).Text = Format(TyGiaCuoi, Mask_0)
    
    LbUser.Caption = UserName
        
    For chi_so = 1 To pSoVV
        LbTT(chi_so - 1).Visible = True
        CboVV(chi_so - 1).Visible = True
        Int_RecsetToCbo "SELECT MaSo As F2,DienGiai As F1 FROM DoituongCT" + CStr(chi_so) + " ORDER BY DoituongCT" + CStr(chi_so) + ".DienGiai", CboVV(chi_so - 1)
    Next
    
    Command(5).Visible = (pBarCode > 0 And CboKho(0).ListCount > 1)
    
    SetFont Me
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set taikhoan = Nothing
    Set taikhoan1 = Nothing
    Set vattu = Nothing
End Sub

Private Sub GrdChuyen_Click()
    Dim i As Integer
    
    SendKeys "{Home}", True
    SetGridIndex GrdChuyen, GrdChuyen.Row
    With GrdChuyen
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        txtChuyen(0).Text = .Text
        .col = 1
        txtChuyen(1).Text = .Text
        txtChuyen_LostFocus 0
        txtChuyen_LostFocus 1
        .col = 13
        If vattu.Dvt2 > 0 Then SetListIndex CboDV, CLng5(.Text)
        For i = 2 To 5
            .col = i
            txtChuyen(i).Text = .Text
        Next
        .col = 14
        txtChuyen(9).Text = .Text
        .col = 10
        MaNhap = CLng5(.Text)
        .col = 11
        SetListIndex CboKho(1), CLng5(.Text)
                
        .col = 7
        taikhoan1.InitTaikhoanSohieu .Text
        txtChuyen(6).Text = .Text
        .RemoveItem .Row
        If .Rows < .tag Then .Rows = .tag
        .col = 0
        TongTien
        RFocus txtChuyen(0)
    End With
End Sub

Private Sub GrdChuyen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then GrdChuyen_Click
End Sub

Private Sub GrdChuyen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdChuyen, GrdChuyen.col
    End If
End Sub

Private Sub MedNgay_GotFocus()
    AutoSelect MedNgay
End Sub
'====================================================================================================
' Ki”m tra ngµy l≠u chuy”n
'====================================================================================================
Private Sub MedNgay_LostFocus()
    Dim m As Integer
    m = Month(ngay)
    On Error GoTo LoiNgayChuyen
    ngay = CDate(MedNgay.Text)
    On Error GoTo 0
    If Month(ngay) <> m Then
        m = Month(ngay)
        ClearGrid GrdChuyen, GrdChuyen.tag
        If CboSohieu.ListIndex >= 0 Then
            CboSohieu.ListIndex = -1
            CboSohieu.Text = ""
        End If
    End If
    Exit Sub
LoiNgayChuyen:
    RFocus MedNgay
End Sub

Private Sub txtChuyen_GotFocus(Index As Integer)
    AutoSelect txtChuyen(Index)
End Sub

Private Sub txtChuyen_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 6:
            If KeyAscii = 13 Then
                flag = 1
                Me.MousePointer = 11
                txtChuyen(Index).Text = FrmTaikhoan.ChonTk(txtChuyen(Index).Text)
                Me.MousePointer = 0
                txtChuyen_LostFocus Index
                If Index = 0 Then
                    RFocus txtChuyen(1)
                Else
                    RFocus cmdct
                End If
                flag = 0
            End If
        Case 1:
            If KeyAscii = 13 Then
                Me.MousePointer = 11
                txtChuyen(1).Text = FrmVattu.ChonVattu(txtChuyen(1).Text)
                Me.MousePointer = 0
                RFocus txtChuyen(1)
            End If
        Case 2, 3:
            KeyAscii = 0
        Case 4:
            If KeyAscii = 13 Then
                CmdCt_Click
            Else
                KeyProcess txtChuyen(Index), KeyAscii
            End If
        Case 5:
            If FCost And txtChuyen(4).tag <> 0 Then KeyAscii = 0 Else KeyProcess txtChuyen(Index), KeyAscii
        Case 8, 9:
            KeyProcess txtChuyen(Index), KeyAscii
    End Select
End Sub

Private Sub txtChuyen_LostFocus(Index As Integer)
    Select Case Index
        Case 0:
            taikhoan.InitTaikhoanSohieu txtChuyen(0).Text
            If Len(txtChuyen(6).Text) = 0 Then
                txtChuyen(6).Text = txtChuyen(0).Text
                taikhoan.DanXuat taikhoan1
            End If
        Case 6:
            taikhoan1.InitTaikhoanSohieu txtChuyen(6).Text
        Case 1:
            Dim ThanhTien As Double, luong As Double, tien2 As Double
            
            CboDV.Visible = False
            
            If FDsNhap.tag > 0 Or FrmVattu.Visible Or flag = 1 Then Exit Sub
            flag = 1
            vattu.InitVattuSohieu txtChuyen(1).Text
            txtChuyen(2).Text = vattu.TenVattu
            txtChuyen(3).Text = vattu.DonVi
                        
            If vattu.MaSo = 0 Then GoTo KT
                        
            If (taikhoan.loai = 0 Or OutCost > 0) Then
                FDsNhap.tag = vattu.MaSo
                MaNhap = FDsNhap.XuatDichDanh(Month(ngay), vattu.sohieu + " - " + vattu.TenVattu + ABCtoVNI(" - ß.v.t: ") + vattu.DonVi, CboKho(0).ItemData(CboKho(0).ListIndex), luong, ThanhTien, tien2)
                FDsNhap.tag = 0
                If luong = 0 Then
                    luong = SoTonKho(Month(ngay), CboKho(0).ItemData(CboKho(0).ListIndex), taikhoan.MaSo, vattu.MaSo, ThanhTien, tien2)
                End If
                txtChuyen(4).tag = luong
                txtChuyen(5).tag = ThanhTien
                txtChuyen(9).tag = tien2
                txtChuyen(4).Text = Format(luong, Mask_2)
                txtChuyen(5).Text = Format(ThanhTien, Mask_0)
                txtChuyen(9).Text = Format(tien2, Mask_2)
            Else
                txtChuyen(4).tag = SoTonKho(Month(ngay), CboKho(0).ItemData(CboKho(0).ListIndex), taikhoan.MaSo, vattu.MaSo, ThanhTien, tien2)
                txtChuyen(5).tag = ThanhTien
                txtChuyen(4).Text = Format(txtChuyen(4).tag, Mask_2)
                txtChuyen(5).Text = Format(ThanhTien, Mask_0)
                txtChuyen(9).tag = tien2
                txtChuyen(9).Text = Format(tien2, Mask_2)
            End If
            If OutCost = 0 And vattu.Dvt2 > 0 Then
                Int_RecsetToCbo "SELECT MaSo AS F2, DonVi AS F1 FROM DVTVattu WHERE MaVattu=" + CStr(vattu.MaSo) + " ORDER BY DonVi", CboDV
                CboDV.AddItem vattu.DonVi, 0
                CboDV.ListIndex = 0
                CboDV.Visible = True
                RFocus CboDV
            Else
                RFocus txtChuyen(4)
            End If
KT:
            flag = 0
        Case 4:
            If txtChuyen(4).tag <= 0 Then
            '    txtChuyen(4).Text = "0"
                Exit Sub
            End If
            If pGiaUSD > 0 Then
                luong = (Cdbl5(txtChuyen(4).Text) * txtChuyen(9).tag) / txtChuyen(4).tag
                txtChuyen(9).Text = Format(luong, Mask_2)
                txtChuyen(5).Text = Format(luong * Cdbl5(txtChuyen(8).Text), Mask_0)
            Else
                txtChuyen(5).Text = Format((Cdbl5(txtChuyen(4).Text) * txtChuyen(5).tag) / txtChuyen(4).tag, Mask_0)
            End If
        Case 7:
            If Len(txtChuyen(7).Text) = 0 Then txtChuyen(7).Text = "..."
    End Select
    HienThongBao "", 1
End Sub
'====================================================================================================
' ThÒ tÙc ki”m tra phi’u l≠u chuy”n c„ hÓp l÷ ?
'====================================================================================================
Private Function KiemTraPhieu() As Boolean
    KiemTraPhieu = False
    If Len(txtChuyen(7).Text) = 0 Then txtChuyen(7).Text = "..."
    If Len(CboSohieu.Text) = 0 Then
        MsgBox "Thi’u sË hi÷u phi’u l≠u chuy”n !", vbExclamation, App.ProductName
        RFocus CboSohieu
        Exit Function
    End If
    If CboSohieu.ListIndex >= 0 Then
        If SelectSQL("SELECT DISTINCTROW Count(MaSo) AS F1 FROM ChungTu WHERE SoHieu = '" + CboSohieu.Text + "' AND MaCT<>" + CStr(CboSohieu.ItemData(CboSohieu.ListIndex)), dbOpenSnapshot) > 0 Then
            ErrMsg er_SHChTu
            RFocus CboSohieu
        End If
    End If
    If CboKho(0).ListIndex < 0 Then
        ErrMsg er_KhoHang
        RFocus CboKho(1)
        Exit Function
    End If
    GrdChuyen.Row = 0
    GrdChuyen.col = 0
    If Len(GrdChuyen.Text) = 0 Then
        MsgBox "H∑y nhÀp chi ti’t l≠u chuy”n !", vbExclamation, App.ProductName
        RFocus txtChuyen(0)
        Exit Function
    End If
    If User_Right <> 0 Then
        If SelectSQL("SELECT Lock" + CStr(Cbo.ItemData(Cbo.ListIndex)) + " Mod 10 AS F1 FROM License") > 0 Then
            MsgBox "Th∏ng Æ∑ bﬁ kho∏ kh´ng cho nhÀp sË li÷u!", vbCritical, App.ProductName
            Exit Function
        End If
    End If
    KiemTraPhieu = True
End Function
'======================================================================================
' ThÒ tÙc x„a phi’u tr™n mµn h◊nh
'======================================================================================
Private Sub XoaPhieuTrenManHinh()
    CboSohieu.ListIndex = -1
    CboSohieu.Text = ""
    ClearGrid GrdChuyen, GrdChuyen.tag
    'GrdChuyen_Click
    txtChuyen(0).Text = ""
    txtChuyen(1).Text = ""
    txtChuyen(2).Text = ""
    txtChuyen(3).Text = "0"
    txtChuyen(4).Text = "0"
    txtChuyen(5).Text = "0"
    txtChuyen(6).Text = ""
    txtChuyen(9).Text = ""
    txtChuyen(7).Text = ABCtoVNI("L≠u chuy”n nÈi bÈ")
    LbTien.Caption = ""
    vattu.InitVattuMaSo 0
    taikhoan.InitTaikhoanMaSo 0
    taikhoan1.InitTaikhoanMaSo 0
    MaNhap = 0
End Sub

Private Function TongTien() As Double
    Dim i As Integer
    
    With GrdChuyen
        For i = 0 To .Rows - 1
            .col = 0
            .Row = i
            If Len(.Text) = 0 Then Exit For
            .col = 5
            TongTien = TongTien + Cdbl5(.Text)
            
        Next
    End With
    LbTien.Caption = Format(TongTien, Mask_0)
End Function

Private Sub LayXuatKho()
    Dim id As Double
    Dim FileNum As Integer
    Dim BytesNeeded As Long
    Dim Buffers As Long
    Dim i As Long, st As String, j As Integer, shtk As String, st2 As String, luong As Double, tien As Double, ms As Long, mtk As Long, mv As Long, sl As Double, T As Double
    Dim Buffer(32)  As Byte
                    
    If Len(Dir(pCurDir + "DOWNLOAD.EXE")) = 0 Then Exit Sub
    Me.MousePointer = 11
    ChDir Left(pCurDir, Len(pCurDir) - 1)
    Recycle pCurDir + "BARCODE.FIL"
    On Error GoTo E
    id = Shell(pCurDir + "DOWNLOAD.EXE", vbMaximizedFocus)
    On Error GoTo 0
    Do While bcstop = 0
        AppIdle 1000
        i = i + 1
        If Len(Dir(pCurDir + "BARCODE.FIL")) > 0 Or i > 10000 Then Exit Do
        'If OpenProcess(PROCESS_ALL_ACCESS, 0&, id) = 0 Then Exit Do
    Loop
    bcstop = 0
    
    If Len(Dir(pCurDir + "BARCODE.FIL")) = 0 Then GoTo KT
        
    XoaPhieuTrenManHinh
    
    FileNum = FreeFile
    Open pCurDir + "BARCODE.FIL" For Binary As #FileNum
    BytesNeeded = LOF(FileNum)
    
    Buffers = BytesNeeded \ 32
    For i = 0 To Buffers - 1
        Get #FileNum, , Buffer
        If i = 0 Then
            st = ""
            For j = 0 To 9
                st = st + Chr(Buffer(j))
            Next
            If IsDate(st) Then
                ngay = CVDate(st)
                MedNgay.Text = Format(ngay, Mask_D)
            End If
        End If
            
        st = ""
        j = 10
        Do While Buffer(j) <> 32 And j < 32
            'If Buffer(j) <> 42 Or pBCode <> 39 Then st = st + Chr(Buffer(j))
            If Buffer(j) <> 42 Then st = st + Chr(Buffer(j))
            j = j + 1
        Loop
        'If pBCode <> 39 Then
        '    If Not IsNumeric(st) Then GoTo E1
        'End If
        
        st2 = ""
        j = 26
        Do While Buffer(j) <> 32 And j < 32
            st2 = st2 + Chr(Buffer(j))
            j = j + 1
        Loop
        luong = Cdbl5(st2)
        
        If pBarCode = 1 Then
            vattu.InitVattuSohieu st
            If vattu.MaSo > 0 Then
                If Len(shtk) = 0 Then
                    shtk = SelectSQL("SELECT HethongTK.Sohieu AS F1 FROM TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo WHERE MaVattu=" + CStr(vattu.MaSo) + " AND MaSoKho=" + CStr(CboKho(0).ItemData(CboKho(0).ListIndex)))
                    
                    Do While SoHieu2MaSo(shtk, "HethongTK") = 0
                        shtk = FrmGetStr.GetString("SË hi÷u tµi kho∂n ghi c„:", "Phi’u xu t kho")
                        If SelectSQL("SELECT TK_ID AS F1 FROM HethongTK WHERE TKCon=0 AND Sohieu='" + shtk + "'") <> TKVT_ID Then shtk = "0"
                    Loop
                End If
                txtChuyen(0).Text = shtk
                txtChuyen(6).Text = shtk
                txtChuyen(1).Text = vattu.sohieu
                txtChuyen_LostFocus 0
                txtChuyen_LostFocus 6
                txtChuyen_LostFocus 1
                txtChuyen(4).Text = Format(luong, Mask_2)
                txtChuyen_LostFocus 4
                CmdCt_Click
            End If
        End If
E1:
    Next

    Close #FileNum
    Me.MousePointer = 0
    Exit Sub
E:
    MsgBox Err.Description
KT:
    Me.MousePointer = 0
End Sub


