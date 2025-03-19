VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmThanhPham2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "B∂ng k™ thµnh ph»m hoµn thµnh trong k˙"
   ClientHeight    =   7080
   ClientLeft      =   825
   ClientTop       =   465
   ClientWidth     =   10380
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
   Icon            =   "FTPham2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Finished Good Voucher"
   Begin MSGrid.Grid Grd 
      Height          =   3495
      Left            =   120
      TabIndex        =   34
      Tag             =   "20"
      Top             =   2760
      Width           =   10215
      _Version        =   65536
      _ExtentX        =   18018
      _ExtentY        =   6165
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
      Cols            =   14
      FixedRows       =   0
      ScrollBars      =   2
   End
   Begin VB.TextBox txtChuyen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   6960
      LinkItem        =   "Gi∏ trﬁ hµng chuy”n Æi"
      Locked          =   -1  'True
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   0
      Left            =   1560
      LinkItem        =   "SË hi÷u chi ti’t c«n xem"
      MaxLength       =   12
      TabIndex        =   7
      Tag             =   "0"
      Top             =   1485
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "&NVL TT"
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
      Index           =   5
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   53
      Tag             =   "Material"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox txtChuyen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   7800
      LinkItem        =   "Gi∏ trﬁ hµng chuy”n Æi"
      Locked          =   -1  'True
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   19
      Tag             =   "0"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox TxtCT 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      LinkItem        =   "SË hi÷u vÀt t≠ c«n xem"
      MaxLength       =   20
      TabIndex        =   20
      Tag             =   "0"
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdtp 
      Height          =   375
      Left            =   5520
      Picture         =   "FTPham2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   4
      Left            =   3600
      Picture         =   "FTPham2.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   47
      Tag             =   "&Print"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CheckBox Chk 
      Caption         =   "T˘ ÆÈng xu t NVL tı kho"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Tag             =   "Export Material from"
      Top             =   720
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox txtChuyen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   6480
      LinkItem        =   "Gi∏ trﬁ hµng chuy”n Æi"
      Locked          =   -1  'True
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdtk 
      Height          =   375
      Index           =   2
      Left            =   2925
      Picture         =   "FTPham2.frx":0822
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   2
      Left            =   1560
      LinkItem        =   "SË hi÷u chi ti’t c«n xem"
      MaxLength       =   12
      TabIndex        =   11
      Tag             =   "0"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   1
      Left            =   1560
      LinkItem        =   "SË hi÷u chi ti’t c«n xem"
      MaxLength       =   12
      TabIndex        =   9
      Tag             =   "0"
      Top             =   1815
      Width           =   1335
   End
   Begin VB.CommandButton cmdtk 
      Height          =   375
      Index           =   1
      Left            =   2925
      Picture         =   "FTPham2.frx":09A0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdtk 
      Height          =   375
      Index           =   0
      Left            =   2925
      Picture         =   "FTPham2.frx":0B1E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   375
   End
   Begin VB.ComboBox CboThang 
      Height          =   315
      ItemData        =   "FTPham2.frx":0C9C
      Left            =   960
      List            =   "FTPham2.frx":0C9E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   8400
      Picture         =   "FTPham2.frx":0CA0
      Style           =   1  'Graphical
      TabIndex        =   33
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   7200
      Picture         =   "FTPham2.frx":0FE2
      Style           =   1  'Graphical
      TabIndex        =   32
      Tag             =   "&Delete"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   6000
      Picture         =   "FTPham2.frx":1324
      Style           =   1  'Graphical
      TabIndex        =   31
      Tag             =   "&Save"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   4800
      Picture         =   "FTPham2.frx":1666
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "&Add"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdct 
      Caption         =   "≠"
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
      Left            =   10080
      TabIndex        =   21
      Tag             =   "-1"
      Top             =   6240
      Width           =   255
   End
   Begin VB.ComboBox CboSohieu 
      Height          =   315
      Left            =   6960
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtChuyen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   5160
      LinkItem        =   "Gi∏ trﬁ hµng chuy”n Æi"
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   17
      Tag             =   "0"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtChuyen 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4440
      LinkItem        =   "SË l≠Óng chuy”n Æi"
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox txtChuyen 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2400
      LinkItem        =   "ß¨n vﬁ t›nh"
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtChuyen 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1080
      LinkItem        =   "T™n vÀt t≠, thµnh ph»m"
      MaxLength       =   20
      TabIndex        =   14
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtChuyen 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      LinkItem        =   "SË hi÷u vÀt t≠, thµnh ph»m (nh n ENTER Æ” xem danh s∏ch)"
      MaxLength       =   20
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin VB.ComboBox CboKho 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Width           =   2655
   End
   Begin VB.ComboBox CboKho 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   327680
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
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L´ hµng"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   9120
      TabIndex        =   56
      Tag             =   "Account"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "T˚ gi∏"
      Height          =   255
      Index           =   18
      Left            =   5760
      TabIndex        =   54
      Tag             =   "Ex. Rate"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ß¨n gi∏"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   6480
      TabIndex        =   52
      Tag             =   "Unit price"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Di‘n gi∂i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   26
      Tag             =   "Description"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label 
      Caption         =   "L´ hµng"
      Height          =   255
      Index           =   13
      Left            =   3240
      TabIndex        =   51
      Tag             =   "Portion"
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label LbCT 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6000
      TabIndex        =   50
      Tag             =   "1"
      Top             =   3525
      Width           =   3975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thµnh ph»m"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   25
      Tag             =   "Product Code"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tµi kho∂n"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   35
      Tag             =   "Account"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label LbGia 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   49
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label 
      Caption         =   "CP SX chung"
      Height          =   255
      Index           =   11
      Left            =   1800
      TabIndex        =   48
      Tag             =   "General Expenses"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label LbGia 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   46
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label LbGia 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   45
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label 
      Caption         =   "CP Nh©n c´ng"
      Height          =   255
      Index           =   16
      Left            =   1800
      TabIndex        =   44
      Tag             =   "Labour Expenses"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gi∏ thµnh"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7800
      TabIndex        =   43
      Tag             =   "Amount"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "CP NVL"
      Height          =   255
      Index           =   14
      Left            =   1800
      TabIndex        =   42
      Tag             =   "Material Expenses"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label LbTenTk 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   41
      Tag             =   "1"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "TK chi ph› NC"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   40
      Tag             =   "Labour Acc."
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label LbTenTk 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   39
      Tag             =   "1"
      Top             =   1800
      Width           =   5655
   End
   Begin VB.Label LbTenTk 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   38
      Tag             =   "1"
      Top             =   1485
      Width           =   5655
   End
   Begin VB.Label Label 
      Caption         =   "TK chi ph› NVL"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   37
      Tag             =   "Material Acc."
      Top             =   1845
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "TK chi ph› SXKD"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   36
      Tag             =   "P. Cost Acc."
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "SË hi÷u"
      Height          =   255
      Index           =   8
      Left            =   5760
      TabIndex        =   29
      Tag             =   "Code"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SË l≠Óng"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   28
      Tag             =   "Quantity"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ß.v.t"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   27
      Tag             =   "Unit"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Th∏ng"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Tag             =   "Month"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Kho TP"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Tag             =   "Pro. Store"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Ngµy"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   22
      Tag             =   "Date"
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FrmThanhPham2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ngay As Date
Dim taikhoan As New ClsTaikhoan
Dim vattu As New ClsVattu
Dim m As New Cls154
Dim F1 As Integer
Dim shct As String
'====================================================================================================
' Chu»n bﬁ phi’u l≠u chuy”n
'====================================================================================================
Private Sub CboKho_Click(Index As Integer)
    If Index = 0 Then
        ClearGrid Grd, Grd.tag
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
    Dim rs As Recordset, i As Integer, tt As Double, dgia As Double
    Dim nvl As Double, nC As Double, sxc As Double
    
    If CboSohieu.ListIndex < 0 Then Exit Sub
    ClearGrid Grd, Grd.tag
    
    Set rs = DBKetoan.OpenRecordset("SELECT ThanhPham.*,Vattu.SoHieu AS SHVT,TenVattu,DonVi,HethongTK.SoHieu AS SHTK FROM (ThanhPham INNER JOIN Vattu ON ThanhPham.MaTP=Vattu.MaSo) INNER JOIN HethongTK ON ThanhPham.MaTK=HethongTK.MaSo WHERE ThanhPham.SoHieu='" + CboSohieu.Text + "' AND Thang=" + CStr(CboThang.ItemData(CboThang.ListIndex)), dbOpenSnapshot)
    If rs.RecordCount = 0 Then GoTo KT
    SetListIndex CboThang, rs!thang
    SetListIndex CboKho(0), rs!MaKhoTP
    SetListIndex CboKho(1), rs!MaKhoNVL
    txtShTk(0).tag = rs!MaTKSX
    txtShTk(1).tag = rs!MaTKCP
    txtShTk(2).tag = rs!MaTKNC
    
    For i = 0 To 2
        taikhoan.InitTaikhoanMaSo txtShTk(i).tag
        txtShTk(i).Text = taikhoan.SoHieu
        LbTenTk(i).Caption = taikhoan.Ten
    Next
    ngay = rs!ngay
    MedNgay.Text = Format(ngay, Mask_D)
    Chk.Value = rs!xk
    Do While Not rs.EOF
        If rs!Ma154 > 0 Then m.InitTPMaSo rs!Ma154 Else m.InitTPMaSo 0
        
        tt = IIf(pDTTP = 0 And pGiaHT = 0, rs!CPNVL + rs!cpnc + rs!cpsxc, rs!GiaThanh)
        If rs!SoLuong <> 0 Then dgia = Fix(0.5 + Mask_N * tt / rs!SoLuong) / Mask_N Else dgia = 0
        Grd.AddItem rs!shtk + Chr(9) + rs!shvt + Chr(9) + rs!TenVattu + Chr(9) + rs!DonVi + Chr(9) + Format(rs!SoLuong, Mask_2) + Chr(9) + Format(dgia, Mask_2) + Chr(9) + Format(tt, Mask_0) _
            + Chr(9) + CStr(rs!MaTK) + Chr(9) + CStr(rs!MaTP) + Chr(9) + DoiDau(rs!CPNVL) + Chr(9) + DoiDau(rs!cpnc) + Chr(9) + DoiDau(rs!cpsxc) + Chr(9) + m.SoHieu + Chr(9) + CStr(m.MaSo), 0
        nvl = nvl + rs!CPNVL
        nC = nC + rs!cpnc
        sxc = sxc + rs!cpsxc
        rs.MoveNext
    Loop
    
    vattu.InitVattuMaSo 0
    taikhoan.InitTaikhoanMaSo 0
    
    rs.MoveFirst
KT:
    LbGia(0).Caption = Format(nvl, Mask_0)
    LbGia(1).Caption = Format(nC, Mask_0)
    LbGia(2).Caption = Format(sxc, Mask_0)
End Sub

Private Sub CboSohieu_LostFocus()
    Dim L As Integer
    
    If CboSohieu.ListIndex < 0 Then
        CboSohieu.Text = UCase(CboSohieu.Text)
        If Len(CboSohieu.Text) > 0 Then
            L = Len(CboSohieu.Text)
            If L > 0 Then
                If Not IsNumeric(CboSohieu.Text) Then
                    shct = SelectSQL("SELECT TOP 1 SoHieu AS F1 FROM ThanhPham WHERE Len(SoHieu)>" + CStr(L) + " AND IsNumeric(Right(SoHieu,Len(SoHieu)-" + CStr(L) + ")) AND SoHieu LIKE'" + CboSohieu.Text + "*' AND Thang=" + CStr(CboThang.ItemData(CboThang.ListIndex)) + " ORDER BY SoHieu DESC")
                    If shct <> "0" Then CboSohieu.Text = SHCtuMoi(shct)
                End If
            End If
                        
            
        End If
    End If
End Sub

Private Sub cboThang_Click()
    If m.MaSo = 0 Then
        If Chk.Value = 0 Then LbGia(0).Caption = Format(SoDuTKSH(txtShTk(1).Text, CboThang.ItemData(CboThang.ListIndex)), Mask_0)
        LbGia(1).Caption = Format(SoDuTKSH(IIf(Len(txtShTk(2).Text) = 0, "622", txtShTk(2).Text), CboThang.ItemData(CboThang.ListIndex)), Mask_0)
        LbGia(2).Caption = Format(SoDuTKSH("627", CboThang.ItemData(CboThang.ListIndex)), Mask_0)
    Else
        LbGia(0).Caption = Format(SoPSCPTP(txtShTk(1).Text, m, CboThang.ItemData(CboThang.ListIndex)) + m.SoCPNVLPB(CboThang.ItemData(CboThang.ListIndex), CboThang.ItemData(CboThang.ListIndex)), Mask_0)
        LbGia(1).Caption = Format(SoPSCPTP(txtShTk(2).Text, m, CboThang.ItemData(CboThang.ListIndex)) + m.SoCPNCPB(CboThang.ItemData(CboThang.ListIndex), CboThang.ItemData(CboThang.ListIndex)), Mask_0)
        LbGia(2).Caption = Format(SoPSCPTP("627", m, CboThang.ItemData(CboThang.ListIndex)) + m.SoCPSXC(CboThang.ItemData(CboThang.ListIndex), CboThang.ItemData(CboThang.ListIndex)), Mask_0)
    End If
    If F1 = 1 Then Exit Sub
    F1 = 1
    Int_RecsetToCbo "SELECT First(MaCT) As F2,SoHieu As F1 FROM ThanhPham WHERE Thang=" + CStr(CboThang.ItemData(CboThang.ListIndex)) + " GROUP BY SoHieu", CboSohieu, -1
    XoaPhieuTrenManHinh
    F1 = 0
End Sub

Private Sub Chk_Click()
    CboKho(1).Enabled = (Chk.Value = 1)
End Sub
'====================================================================================================
' Th™m, sˆa, x„a phi’u l≠u chuy”n
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim chungtu As New ClsChungtu, sh As String, DaNhap As Integer, m1 As New Cls154
    Dim MaCT As Long, i As Integer, mvt As Long, luong As Double, tien As Double, mtk As Long
    Dim nvl As Double, nC As Double, sxc As Double, n As Date, dgia As Double, thang As Integer
            
    Me.MousePointer = 11
    thang = CboThang.ItemData(CboThang.ListIndex)
    Select Case Index
        Case 0:
            XoaPhieuTrenManHinh
            Grd.Row = 0
            cboThang_Click
            If pGiaHT > 0 Then
                mtk = SelectSQL("SELECT MaSo AS F1 FROM ThanhPham WHERE Thang=" + CStr(thang) + " AND Ma154=" + CStr(m.MaSo))
                If mtk > 0 Then SetListIndex CboSohieu, mtk Else LietKeTP
            Else
                If Len(shct) > 0 Then CboSohieu.Text = SHCtuMoi(shct)
                RFocus CboSohieu
            End If
        Case 1:
            Me.Refresh
            If Not KiemTraPhieu Then GoTo XongLC
            sh = CboSohieu.Text
            If CboSohieu.ListIndex >= 0 Then
                i = SelectSQL("SELECT TOP 1 Thang AS F1,Ma154 AS F2 FROM ThanhPham WHERE MaCT=" + CStr(CboSohieu.ItemData(CboSohieu.ListIndex)), mtk)
                n = SelectSQL("SELECT TOP 1 Ngay AS F1 FROM ThanhPham WHERE MaCT=" + CStr(CboSohieu.ItemData(CboSohieu.ListIndex)))
                ExecuteSQL5 "DELETE * FROM ThanhPham WHERE MaCT=" + CStr(CboSohieu.ItemData(CboSohieu.ListIndex))
                XoaCTTheoID i, 610000000 + CboSohieu.ItemData(CboSohieu.ListIndex), 0, n, n
                CboSohieu.RemoveItem CboSohieu.ListIndex
            Else
                n = ngay
            End If
            'F1 = 1
            'cboThang_Click
            'F1 = 0
            'GiaBQ
            MaCT = Lng_MaxValue("MaCT", "ThanhPham") + 1
            With Grd
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 7
                    If Len(.Text) = 0 Then Exit For
                    mtk = CLng5(.Text)
                    .col = 8
                    mvt = CLng5(.Text)
                    .col = 4
                    luong = Cdbl5(.Text)
                    .col = 6
                    tien = Cdbl5(.Text)
                    .col = 9
                    nvl = Cdbl5(.Text)
                    .col = 10
                    nC = Cdbl5(.Text)
                    .col = 11
                    sxc = Cdbl5(.Text)
                    .col = 13
                    If luong <> 0 Then dgia = tien / luong Else dgia = 0
                    ExecuteSQL5 "INSERT INTO ThanhPham (MaCT,Thang,SoHieu,Ngay,MaKhoTP,MaKhoNVL,MaTK,MaTP,SoLuong,MaTKCP,MaTKNC,MaTKSX,Ma154,GiaThanh,CPNVL,CPNC,CPSXC,XK) VALUES (" + CStr(MaCT) + "," + CStr(thang) _
                        + ",'" + sh + "',#" + Format(ngay, Mask_DB) + "#," + CStr(CboKho(0).ItemData(CboKho(0).ListIndex)) + "," + CStr(CboKho(1).ItemData(CboKho(1).ListIndex)) _
                        + "," + CStr(mtk) + "," + CStr(mvt) + "," + DoiDau(luong) + "," + CStr(txtShTk(1).tag) + "," + CStr(txtShTk(2).tag) + "," + CStr(txtShTk(0).tag) + "," + CStr(CLng5(.Text)) + "," + DoiDau(tien) + "," + DoiDau(nvl) + "," + DoiDau(nC) + "," + DoiDau(sxc) + "," + CStr(Chk.Value) + ")"
                    If pGiaHT > 0 Then DieuChinhGiaTP thang, m.MaSo, mtk, mvt, dgia, tien
                Next
            End With
            
            CboSohieu.AddItem sh
            CboSohieu.ItemData(CboSohieu.NewIndex) = MaCT
            m.DanXuat m1
            shct = chungtu.SoHieu
            Command_Click 0
            DaNhap = 1
        Case 2:
            If CboSohieu.ListIndex >= 0 Then
                MaCT = CboSohieu.ItemData(CboSohieu.ListIndex)
                If pGiaHT > 0 Then DatLaiGiaHT MaCT
                ExecuteSQL5 "DELETE * FROM ThanhPham WHERE MaCT=" + CStr(MaCT)
                CboSohieu.RemoveItem CboSohieu.ListIndex
                DaNhap = 1
            End If
            m.DanXuat m1
            XoaPhieuTrenManHinh
            Grd.Row = 0
        Case 3:
            Unload Me
        Case 4:
            InGiaThanh
        Case 5:
            InNVL
    End Select
    If DaNhap = 1 And (pGiaHT = 0 Or pDTTP = 0) Then
        GhiXuatNVL MaCT, ngay, thang, Chk.Value, m1, , Cdbl5(txtChuyen(7).Text)
        If m1.MaSo > 0 Then
            If m.GiaThanhCK(ThangTruoc(thang)) = 0 Then
                nC = m1.SoCPNC(thang, thang) + m1.SoCPNCPB(thang, thang)
            Else
                nC = m1.SoCPNC(pThangDauKy, thang) + m1.SoCPNCPB(pThangDauKy, thang)
            End If
            sxc = SelectSQL("SELECT Sum(CPNC) AS F1 FROM ThanhPham WHERE Thang=" + CStr(thang) + " AND Ma154=" + CStr(m1.MaSo))
            If nC <> sxc Then DieuChinhGiaThanh m1, thang, 1            'TinhLaiGiaThanhPham i, i
            'DieuChinhGiaThanh m1, i, 1
        Else
            TinhLaiGiaVonTP m1, thang
        End If
        DaNhap = 0
    End If
    If DaNhap = 1 And pGiaHT > 0 Then
        KiemTraVatTu
        KiemTraTaiKhoan
    End If
XongLC:
    Set m1 = Nothing
    Set chungtu = Nothing
    Me.MousePointer = 0
    If pGiaHT > 0 And Index = 1 Then MsgBox "ß∑ t›nh xong gi∏ thµnh!", vbInformation, App.ProductName
End Sub
'====================================================================================================
' Ghi chi ti’t l≠u chuy”n
'====================================================================================================
Private Sub CmdCt_Click()
    Dim luong As Double, i As Integer, dgia As Double, tien As Double
    
    If (taikhoan.MaSo = 0 Or taikhoan.tk_id <> TKVT_ID) Then
        RFocus txtChuyen(0)
        Exit Sub
    End If
    
    If vattu.MaSo = 0 Then
        RFocus txtChuyen(1)
        Exit Sub
    End If
    
    luong = Cdbl5(txtChuyen(4).Text)
    If luong = 0 Then
        MsgBox "H∑y nhÀp sË l≠Óng thµnh ph»m !", vbExclamation, App.ProductName
        RFocus txtChuyen(4)
        Exit Sub
    End If
    
    If pGiaHT = 0 Then
        With Grd
            For i = 0 To .Rows - 1
                .col = 7
                .Row = i
                If Len(.Text) = 0 Then Exit For
                If CLng5(.Text) = taikhoan.MaSo Then
                    .col = 8
                     If CLng5(.Text) = vattu.MaSo Then
                        .RemoveItem i
                        .AddItem taikhoan.SoHieu + Chr(9) + vattu.SoHieu + Chr(9) + vattu.TenVattu + Chr(9) + vattu.DonVi + Chr(9) _
                            + Format(luong, Mask_2) + Chr(9) + "0" + Chr(9) + "0" + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) + CStr(vattu.MaSo) + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + CStr(m.SoHieu) + Chr(9) + CStr(m.MaSo), i
                        GoTo ABC
                    End If
                End If
            Next
        End With
        i = NewRowIndex(Grd, 0)
        Grd.AddItem taikhoan.SoHieu + Chr(9) + vattu.SoHieu + Chr(9) + vattu.TenVattu + Chr(9) + vattu.DonVi + Chr(9) _
            + Format(luong, Mask_2) + Chr(9) + "0" + Chr(9) + "0" + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) + CStr(vattu.MaSo) + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + CStr(m.SoHieu) + Chr(9) + CStr(m.MaSo), i
        Grd.Row = Grd.Rows - 1
        Grd.col = 7
        If Len(Grd.Text) = 0 Then Grd.RemoveItem Grd.Row
        Grd.Row = 0
        Grd.col = 0
        
        GiaBQ i
    
ABC:
        For i = 0 To 3
            txtChuyen(i).Text = ""
        Next
        For i = 4 To 6
            txtChuyen(i).Text = "0"
        Next
        vattu.InitVattuMaSo 0
        taikhoan.InitTaikhoanMaSo 0
        RFocus txtChuyen(0)
    Else
        If cmdct.tag < 0 Then Exit Sub
        tien = Cdbl5(txtChuyen(6).Text)
        If luong <> 0 Then dgia = tien / luong Else dgia = 0
        With Grd
            .Row = cmdct.tag
            .col = 5
            .Text = Format(dgia, Mask_2)
            .col = 6
            .Text = Format(tien, Mask_0)
            If .Row < .Rows - 1 Then .Row = cmdct.tag + 1
            Grd_DblClick
        End With
    End If
End Sub

'====================================================================================================
' Xˆ l˝ c∏c ph›m n„ng
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
            Case vbKeyI:
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
' KhÎi tπo cˆa sÊ nhÀp
'====================================================================================================
Private Sub Form_Load()
    Dim i As Integer
    
    ColumnSetUp Grd, 0, 940, 0
    ColumnSetUp Grd, 1, 1300, 0
    ColumnSetUp Grd, 2, 2020, 0
    ColumnSetUp Grd, 3, 700, 0
    ColumnSetUp Grd, 4, 1300, 1
    ColumnSetUp Grd, 5, 1300, 1
    ColumnSetUp Grd, 6, 1300, 1
    ColumnSetUp Grd, 7, 1, 0                ' MaTK
    ColumnSetUp Grd, 8, 1, 0                ' MaTP
    ColumnSetUp Grd, 9, 1, 1
    ColumnSetUp Grd, 10, 1, 1
    ColumnSetUp Grd, 11, 1, 1
    ColumnSetUp Grd, 12, 1180, 0
    ColumnSetUp Grd, 13, 1, 0
    
    AddMonthToCbo CboThang
    InitDateVars MedNgay, ngay

    Int_RecsetToCbo "SELECT MaSo As F2,TenKho As F1 FROM KhoHang ORDER BY TenKho", CboKho(0)
    Int_RecsetToCbo "SELECT MaSo As F2,TenKho As F1 FROM KhoHang ORDER BY TenKho", CboKho(1)
    
    CboSohieu.ListIndex = -1
    'Command_Click 0
    
    Caption = Caption + " - " + CStr(pNamTC)
            
    txtShTk(0).Text = "154"
    txtShTk(1).Text = "621"
    txtShTk(2).Text = "622"
    For i = 0 To 2
        txtShTk_LostFocus i
    Next
    CboKho(1).Enabled = (Chk.Value = 1)
    Chk.Value = SelectSQL("SELECT TOP 1 XK AS F1 FROM ThanhPham ORDER BY Ngay DESC")
    
    Label(13).Visible = (pDTTP <> 0)
    TxtCT.Visible = (pDTTP <> 0)
    cmdtp.Visible = (pDTTP <> 0)
    
    Label(18).Visible = (pTygia > 0)
    txtChuyen(7).Visible = (pTygia > 0)
    If pTygia > 0 Then txtChuyen(7).Text = Format(TyGiaCuoi, Mask_0)
    If pGiaHT > 0 Then
        txtChuyen(4).Locked = True
        txtChuyen(5).Locked = False
        txtChuyen(6).Locked = False
    End If
    SetFont Me
LoiNgay:
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    SaveSetting IniPath, "Stock", "AutoOut", Chk.Value
    Set taikhoan = Nothing
    Set vattu = Nothing
    Set m = Nothing
End Sub

Private Sub Grd_DblClick()
    Dim i As Integer
    
    SendKeys "{Home}", True
    SetGridIndex Grd, Grd.Row
    With Grd
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        txtChuyen(0).Text = .Text
        .col = 7
        taikhoan.InitTaikhoanMaSo CLng5(.Text)
        .col = 8
        vattu.InitVattuMaSo CLng5(.Text)
        For i = 1 To 4
            .col = i
            txtChuyen(i).Text = .Text
        Next
        If pGiaHT = 0 Then
            .RemoveItem .Row
            If .Rows < .tag Then .Rows = .tag
            'GiaBQ i
            RFocus txtChuyen(0)
        Else
            cmdct.tag = .Row
            RFocus txtChuyen(5)
        End If
        .col = 0
    End With
End Sub

Private Sub Grd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Grd_DblClick
End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , Grd, Grd.col
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
        ClearGrid Grd, Grd.tag
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
        Case 0:
            If KeyAscii = 13 Then
                Me.MousePointer = 11
                txtChuyen(0).Text = FrmTaikhoan.ChonTk(txtChuyen(0).Text)
                Me.MousePointer = 0
                RFocus txtChuyen(0)
            End If
        Case 1:
            If KeyAscii = 13 Then
                Me.MousePointer = 11
                txtChuyen(1).Text = FrmVattu.ChonVattu(txtChuyen(1).Text)
                Me.MousePointer = 0
                RFocus txtChuyen(1)
            End If
        Case 4, 7:
            If KeyAscii = 13 Then
                CmdCt_Click
            Else
                KeyProcess txtChuyen(Index), KeyAscii
            End If
        Case 5, 6:
            KeyProcess txtChuyen(Index), KeyAscii
    End Select
End Sub

Private Sub txtChuyen_LostFocus(Index As Integer)
    Dim luong As Double, dgia As Double
    Select Case Index
        Case 0:
            taikhoan.InitTaikhoanSohieu txtChuyen(0).Text
        Case 1:
            vattu.InitVattuSohieu txtChuyen(1).Text
            txtChuyen(2).Text = vattu.TenVattu
            txtChuyen(3).Text = vattu.DonVi
        Case 4:
            txtChuyen(4).Text = Format(txtChuyen(4).Text, Mask_2)
        Case 5:
            txtChuyen(6).Text = Format(Cdbl5(txtChuyen(4).Text) * Cdbl5(txtChuyen(5).Text), Mask_0)
        Case 6:
            luong = Cdbl5(txtChuyen(4).Text)
            If luong > 0 Then dgia = Cdbl5(txtChuyen(6).Text) / luong Else dgia = 0
            txtChuyen(5).Text = Format(dgia, Mask_2)
    End Select
End Sub
'====================================================================================================
' ThÒ tÙc ki”m tra phi’u l≠u chuy”n c„ hÓp l÷ ?
'====================================================================================================
Private Function KiemTraPhieu() As Boolean
    KiemTraPhieu = False
    
    If Len(CboSohieu.Text) = 0 Then
        ErrMsg er_SoHieu
        RFocus CboSohieu
        Exit Function
    End If
        
    If CboKho(0).ListIndex < 0 Then
        ErrMsg er_KhoHang
        RFocus CboKho(1)
        Exit Function
    End If
    
    
    Grd.Row = 0
    Grd.col = 0
    If Len(Grd.Text) = 0 Then
        MsgBox "H∑y nhÀp chi ti’t b∂ng k™ !", vbExclamation, App.ProductName
        RFocus txtChuyen(0)
        Exit Function
    End If
    
    If Left(txtShTk(0).Text, 3) <> "154" Or txtShTk(0).tag = 0 Then
        RFocus txtShTk(0)
        Exit Function
    End If
    
    If Left(txtShTk(1).Text, 3) <> "621" Or txtShTk(1).tag = 0 Then
        RFocus txtShTk(1)
        Exit Function
    End If
    
    If Left(txtShTk(2).Text, 3) <> "622" Or txtShTk(2).tag = 0 Then
        RFocus txtShTk(2)
        Exit Function
    End If
        
    If pDTTP <> 0 And m.MaSo = 0 Then
        RFocus TxtCT
        Exit Function
    End If
    
    If User_Right <> 0 Then
        If SelectSQL("SELECT Lock" + CStr(CboThang.ItemData(CboThang.ListIndex)) + " Mod 10 AS F1 FROM License") > 0 Then
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
    Dim i As Integer
    
    CboSohieu.ListIndex = -1
    CboSohieu.Text = ""
    ClearGrid Grd, Grd.tag
    For i = 0 To 3
        txtChuyen(i).Text = ""
    Next
    For i = 4 To 6
        txtChuyen(i).Text = "0"
    Next
    If pGiaHT = 0 Then
        TxtCT.Text = ""
        TxtCT_LostFocus
    End If
    If F1 = 0 Then
        For i = 0 To 2
            LbGia(i).Caption = ""
        Next
    End If
    vattu.InitVattuMaSo 0
    taikhoan.InitTaikhoanMaSo 0
    cmdct.tag = -1
End Sub

Private Sub txtShTk_GotFocus(Index As Integer)
    AutoSelect txtShTk(Index)
End Sub

Private Sub txtShTk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdtk_Click Index
End Sub
'====================================================================================================
' Ki”m tra sË hi÷u tµi kho∂n Æ∑ nhÀp
'====================================================================================================
Private Sub txtShTk_LostFocus(Index As Integer)
    Dim mtk As Long
    
    mtk = 0
    LbTenTk(Index).Caption = tentk(txtShTk(Index).Text, mtk)
    txtShTk(Index).tag = mtk
    If SelectSQL("SELECT TKCon AS F1 FROM HethongTK WHERE MaSo=" + CStr(mtk)) > 0 Then txtShTk(Index).tag = 0
End Sub

Private Sub cmdtk_Click(Index As Integer)
    Me.MousePointer = 11
    txtShTk(Index).Text = FrmTaikhoan.ChonTk(txtShTk(Index).Text)
    RFocus txtShTk(Index)
    Me.MousePointer = 0
End Sub

Private Sub GiaBQ(r2 As Integer)
    Dim i As Integer, luong As Double, luong1 As Double, mvt As Long, k As Integer
    Dim nvl As Double, nC As Double, nvl1 As Double, nc1 As Double, CP As Double, cp1 As Double, tcp As Double
    Dim n2 As Double, v2 As Double
    
    'gia = tp.GiaThanhCK(CboThang.ListIndex + 1)
    If pGiaHT > 0 And Chk.Value = 0 Then Exit Sub
    With Grd
            .Row = r2
            .col = 0
            .col = 4
            luong1 = Cdbl5(.Text)
            luong = luong + luong1
            .col = 8
            mvt = CLng5(.Text)
            If Chk.Value = 1 Then
                nvl1 = GiaNVLTheoDM(mvt, luong1, CboKho(1).ItemData(CboKho(1).ListIndex), ngay)
            End If
            nc1 = GiaNCTheoDM(mvt, luong1, CboThang.ItemData(CboThang.ListIndex))
            If nc1 > 0 Then k = 1
            nvl = nvl + nvl1
            nC = nC + nc1
            .col = 9
            .Text = Format(nvl1, Mask_0)
            .col = 10
            .Text = Format(nc1, Mask_0)
            .col = 11
            .Text = "0"
            .col = 6
            .Text = Format(nvl1 + nc1, Mask_0)

        CP = 0
        
        If m.MaSo > 0 Then
            v2 = SoPSCPTP_DM("621", m, CboThang.ItemData(CboThang.ListIndex)) + m.SoCPNVLPB(CboThang.ItemData(CboThang.ListIndex), CboThang.ItemData(CboThang.ListIndex))
            n2 = SoPSCPTP_DM("622", m, CboThang.ItemData(CboThang.ListIndex)) + m.SoCPNCPB(CboThang.ItemData(CboThang.ListIndex), CboThang.ItemData(CboThang.ListIndex))
        Else
            If k = 0 Then n2 = Cdbl5(LbGia(1).Caption) Else n2 = 0
        End If
    
        tcp = v2
        If tcp <> 0 Then
            .Row = r2
            .col = 6
            .Text = Format(Cdbl5(.Text) + tcp, Mask_0)
            .col = 9
            .Text = Format(Cdbl5(.Text) + tcp, Mask_0)
        End If
            
            ' Phan bo NC
            CP = 0
            tcp = n2
            If tcp <> 0 Then
                .Row = r2
                .col = 6
                .Text = Format(Cdbl5(.Text) + tcp, Mask_0)
                .col = 10
                .Text = Format(Cdbl5(.Text) + tcp, Mask_0)
            End If
        
        
        nvl = nvl + v2
        nC = nC + n2
        LbGia(0).Caption = Format(nvl, Mask_0)
        LbGia(1).Caption = Format(nC, Mask_0)
        
        ' Phan bo CP chung
        tcp = Cdbl5(LbGia(2).Caption)
        CP = 0
        If nvl <> 0 And tcp <> 0 Then
            .Row = r2
            .col = 6
            .Text = Format(Cdbl5(.Text) + tcp, Mask_0)
            .col = 11
            .Text = Format(Cdbl5(.Text) + tcp, Mask_0)
        End If
        
            .Row = r2
            .col = 0

            .col = 4
            luong1 = Cdbl5(.Text)
            If luong1 <> 0 Then
                .col = 6
                CP = Cdbl5(.Text)
                .col = 5
                .Text = Format(CP / luong1, Mask_2)
            End If

    End With
End Sub

Private Sub InGiaThanh()
    Dim i As Integer, sh As String, Ten As String, sl As Double, nvl As Double, nC As Double, sxc As Double, tnvl As Double
    
    Me.MousePointer = 11
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    With Grd
        For i = 0 To .Rows - 1
            .Row = i
            .col = 1
            If Len(.Text) = 0 Then Exit For
            sh = .Text
            .col = 2
            Ten = .Text
            .col = 4
            sl = Cdbl5(.Text)
            .col = 9
            nvl = Cdbl5(.Text)
            tnvl = tnvl + nvl
            .col = 10
            nC = Cdbl5(.Text)
            .col = 11
            sxc = Cdbl5(.Text)
            ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten, Kq1,Kq2,Kq3,Kq4) VALUES (" + CStr(i) + ",'" + sh + "','" + Ten + "'," + DoiDau(sl) + "," + DoiDau(nvl) + "," + DoiDau(nC) + "," + DoiDau(sxc) + ")"
        Next
    End With
    
    SetSQL "QNhatky", "SELECT * FROM BaoCaoCP2 ORDER BY SoHieu"
    SetRptInfo
    frmMain.Rpt.ReportFileName = "GTCT.RPT"
    i = CboThang.ItemData(CboThang.ListIndex)
    RptSetDate NgayCuoiThang(pNamTC, i)
    frmMain.Rpt.Formulas(4) = "Thang = " + CStr(i)
    frmMain.Rpt.Formulas(13) = "ThangCuoi = " + CStr(i)
    frmMain.Rpt.Formulas(5) = "CPNVL = " + DoiDau(tnvl)
    frmMain.Rpt.Formulas(6) = "6271 = " + DoiDau(PSTKCP("6271", i, i, m.MaSo))
    frmMain.Rpt.Formulas(7) = "6272 = " + DoiDau(PSTKCP("6272", i, i, m.MaSo))
    frmMain.Rpt.Formulas(8) = "6273 = " + DoiDau(PSTKCP("6273", i, i, m.MaSo))
    frmMain.Rpt.Formulas(9) = "6274 = " + DoiDau(PSTKCP("6274", i, i, m.MaSo))
    frmMain.Rpt.Formulas(10) = "6277 = " + DoiDau(PSTKCP("6277", i, i, m.MaSo))
    frmMain.Rpt.Formulas(11) = "6278 = " + DoiDau(PSTKCP("6278", i, i, m.MaSo))
    
    'If m.MaSo > 0 Then frmMain.Rpt.Formulas(12) = "DG='" + m.SoHieu + " - " + m.TenVattu + "'"
    
    frmMain.Rpt.WindowTitle = "Gi∏ thµnh thµnh ph»m hoµn thµnh trong k˙"
    frmMain.Rpt.Destination = crptToWindow
    Me.MousePointer = 0
    InBaoCaoRPT
End Sub

Private Sub TxtCT_GotFocus()
    AutoSelect TxtCT
End Sub

Private Sub TxtCT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtCT.Text = FrmTP.ChonTP(TxtCT.Text)
        RFocus TxtCT
    End If
End Sub

Private Sub TxtCT_LostFocus()
    If pDTTP = 0 Then Exit Sub
    m.InitTPSohieu TxtCT.Text
    LbCT.Caption = m.TenVattu
    TxtCT.Text = m.SoHieu
    F1 = 1
    cboThang_Click
    F1 = 0
End Sub

Private Sub CmdTP_Click()
    TxtCT.Text = FrmTP.ChonTP(TxtCT.Text)
    TxtCT_LostFocus
End Sub

Private Sub InNVL()
    Dim sql As String, rs As Recordset, i As Integer, sl As Double, mvt As Long, thangdm As String

    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,Kq1,Kq2) " _
        & " SELECT MaVattu, Vattu.SoHieu, Vattu.TenVattu, Sum(SoPS2Co), Sum(SoPS) " _
        & " FROM (" + ChungTu2TKNC(-1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo" _
        & " WHERE HethongTK.SoHieu LIKE '621*' AND ChungTu.MaLoai=2 AND ThangCT=" + CStr(CboThang.ItemData(CboThang.ListIndex)) + IIf(m.MaSo > 0, " AND ChungTu.MaTP=" + CStr(m.MaSo), "") + " GROUP BY MaVattu, Vattu.SoHieu, Vattu.TenVattu"
    sql = "SELECT MaVattu, Sum(SoPS2Co) AS SL, Sum(SoPS) AS TT" _
        & " FROM " + ChungTu2TKNC(1) _
        & " WHERE HethongTK.SoHieu LIKE '621*' AND ChungTu.MaLoai=1 AND ThangCT=" + CStr(CboThang.ItemData(CboThang.ListIndex)) + IIf(m.MaSo > 0, " AND ChungTu.MaTP=" + CStr(m.MaSo), "") + " GROUP BY MaVattu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq1=Kq1-" + DoiDau(rs!sl) + ",Kq2=Kq2-" + DoiDau(rs!tt) + " WHERE MaSo=" + CStr(rs!MaVattu)
        rs.MoveNext
    Loop
    rs.Close
        
    With Grd
        For i = 0 To .Rows - 1
            .Row = i
            .col = 8
            If Not IsNumeric(.Text) Then Exit For
            mvt = CLng(.Text)
            thangdm = SelectSQL("SELECT TOP 1 Thang AS F1 FROM DinhMuc WHERE MaTP=" + CStr(mvt) + " AND " + WThang("Thang", 0, CboThang.ItemData(CboThang.ListIndex)) + " ORDER BY " + SetMonthOrder("Thang") + " DESC")
            .col = 4
            sl = Cdbl5(.Text)
            Set rs = DBKetoan.OpenRecordset("SELECT MaNVL, SoLuong, SoHieu,TenVattu FROM DinhMuc INNER JOIN Vattu ON DinhMuc.MaNVL=Vattu.MaSo WHERE MaTP=" + CStr(mvt) + " AND Thang=" + CStr(thangdm))
            Do While Not rs.EOF
                ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq3=Kq3+" + DoiDau(sl * rs!SoLuong) + " WHERE MaSo=" + CStr(rs!MaNVL)
                If DBKetoan.RecordsAffected = 0 Then
                    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,Kq3) VALUES (" + CStr(rs!MaNVL) + ",'" + rs!SoHieu + "','" + rs!TenVattu + "'," + DoiDau(sl * rs!SoLuong) + ")"
                End If
                rs.MoveNext
            Loop
            rs.Close
        Next
    End With
        
    Set rs = Nothing
    
    SetSQL "QNhatky", "SELECT * FROM BaoCaoCP2 ORDER BY SoHieu"
    SetRptInfo
    frmMain.Rpt.ReportFileName = "NVLTT.RPT"
    i = CboThang.ItemData(CboThang.ListIndex)
    RptSetDate NgayCuoiThang(pNamTC, i)
    frmMain.Rpt.Formulas(4) = "Thang = " + CStr(i)
    frmMain.Rpt.Formulas(5) = "ThangCuoi = " + CStr(i)
    If m.MaSo > 0 Then frmMain.Rpt.Formulas(12) = "DG='" + m.SoHieu + " - " + m.TenVattu + "'"
    frmMain.Rpt.WindowTitle = "Chi ti’t chi ph› nguy™n vÀt li÷u tr˘c ti’p"
    frmMain.Rpt.Destination = crptToWindow
    Me.MousePointer = 0
    InBaoCaoRPT
End Sub

Private Sub LietKeTP()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT MaTKNo, MaVattu, Sum(SoPS2No) AS SL, Vattu.SoHieu, Vattu.TenVattu, DonVi, HethongTK.SoHieu AS SHTK " _
        & " FROM (" + ChungTu2TKNC(-1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo " _
        & " WHERE ThangCT=" + CStr(CboThang.ItemData(CboThang.ListIndex)) + " AND HethongTK.SoHieu LIKE '" + ShTkTP + "*' AND MaTP=" + CStr(m.MaSo) _
        & " GROUP BY MaTKNo, MaVattu, Vattu.SoHieu, Vattu.TenVattu, DonVi, HethongTK.SoHieu ORDER BY HethongTK.SoHieu DESC, Vattu.SoHieu DESC"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        Grd.AddItem rs!shtk + Chr(9) + rs!SoHieu + Chr(9) + rs!TenVattu + Chr(9) + rs!DonVi + Chr(9) + Format(rs!sl, Mask_2) + Chr(9) + "" + Chr(9) + "" + Chr(9) + CStr(rs!MaTkNo) + Chr(9) + CStr(rs!MaVattu), 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Grd.Row = 0
    'If Chk.Value > 0 Then GiaBQ
    'If pGiaHT = 0 Then Grd_DblClick
End Sub

