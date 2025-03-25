VERSION 5.00
Begin VB.Form FrmTP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HÖ thèng danh ®iÓm c«ng tr×nh, s¶n phÈm"
   ClientHeight    =   7080
   ClientLeft      =   510
   ClientTop       =   870
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
   Icon            =   "FrmTP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Products and Contructions"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   0
      Left            =   6240
      LinkItem        =   "Sè hiÖu chi tiÕt cÇn xem"
      MaxLength       =   12
      TabIndex        =   6
      Tag             =   "0"
      Text            =   "154"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdtk 
      Height          =   375
      Index           =   0
      Left            =   7680
      Picture         =   "FrmTP.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1860
      Width           =   375
   End
   Begin VB.ComboBox CboThang 
      Height          =   315
      ItemData        =   "FrmTP.frx":5960
      Left            =   8280
      List            =   "FrmTP.frx":5962
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox TxtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   6240
      MaxLength       =   20
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox TxtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   6240
      MaxLength       =   50
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox TxtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   6240
      MaxLength       =   12
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox TxtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   6240
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox TxtVT 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   8
      Text            =   "0"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   3
      Left            =   8640
      Picture         =   "FrmTP.frx":5964
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   2
      Left            =   7440
      Picture         =   "FrmTP.frx":6D86
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "&Delete"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   1
      Left            =   6240
      Picture         =   "FrmTP.frx":8268
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "&Save"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   0
      Left            =   5040
      Picture         =   "FrmTP.frx":9696
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "&Add"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton SSCmdF 
      BackColor       =   &H80000013&
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
      Left            =   4560
      TabIndex        =   15
      Top             =   6600
      Width           =   255
   End
   Begin VB.OptionButton SSOpt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tªn TP"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   14
      Tag             =   "Description"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.OptionButton SSOpt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè hiÖu"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Tag             =   "Code"
      Top             =   6600
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Left            =   2760
      TabIndex        =   12
      Top             =   6600
      Width           =   1695
   End
   Begin VB.ComboBox CboLoai 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.ListBox LstVt 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   41
      Tag             =   "1"
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tµi kho¶n"
      Height          =   255
      Index           =   16
      Left            =   5280
      TabIndex        =   40
      Tag             =   "Notes"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Th¸ng"
      Height          =   255
      Index           =   15
      Left            =   7560
      TabIndex        =   38
      Tag             =   "Notes"
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè hiÖu"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   37
      Tag             =   "Code"
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tªn"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   36
      Tag             =   "Desc."
      Top             =   720
      Width           =   615
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   6240
      X2              =   8160
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line 
      Index           =   1
      X1              =   6240
      X2              =   9360
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "§¬n vÞ"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   35
      Tag             =   "Unit"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Line Line 
      Index           =   2
      X1              =   6240
      X2              =   8160
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ghi chó"
      Height          =   255
      Index           =   13
      Left            =   5280
      TabIndex        =   34
      Tag             =   "Notes"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Line Line 
      Index           =   8
      X1              =   6240
      X2              =   9360
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ nguyªn vËt liÖu trùc tiÕp"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   33
      Tag             =   "Material Expenses"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ Nh©n c«ng trùc tiÕp"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   32
      Tag             =   "Labour Expenses"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ M¸y thi c«ng"
      Height          =   255
      Index           =   7
      Left            =   5280
      TabIndex        =   31
      Tag             =   "Machine Expenses"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ KhÊu hao TSC§"
      Height          =   255
      Index           =   8
      Left            =   5280
      TabIndex        =   30
      Tag             =   "Depreciation"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ s¶n xuÊt chung"
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   29
      Tag             =   "General Expenses"
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gi¸ thµnh"
      Height          =   255
      Index           =   11
      Left            =   5280
      TabIndex        =   28
      Tag             =   "Product Cost"
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label LbCP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   27
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label LbCP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   8040
      TabIndex        =   26
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label LbCP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   8040
      TabIndex        =   25
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label LbCP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   8040
      TabIndex        =   24
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label LbCP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   23
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label LbCP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   8040
      TabIndex        =   22
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dë dang ®Çu kú"
      Height          =   255
      Index           =   12
      Left            =   5280
      TabIndex        =   21
      Tag             =   "Opening Balance"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Line Line 
      Index           =   3
      X1              =   7920
      X2              =   9360
      Y1              =   3405
      Y2              =   3405
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6375
      Index           =   14
      Left            =   4920
      TabIndex        =   20
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   5
      Left            =   4980
      TabIndex        =   11
      Top             =   180
      Width           =   4695
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   4
      Left            =   5040
      TabIndex        =   10
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label2"
      Height          =   5955
      Index           =   9
      Left            =   165
      TabIndex        =   9
      Top             =   540
      Width           =   4695
   End
End
Attribute VB_Name = "FrmTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer          ' =1 neu them moi, -1 neu sua cu
Dim tp As New Cls154      ' vat tu duoc tham chieu
Dim doiloai As Integer               ' =1 neu co thay doi loai vat tu dang sua doi
Dim MaDaTim As Long
Dim TK As New ClsTaikhoan
'======================================================================================
' Liet ke cac vat tu trong loai vat tu duoc chon
'======================================================================================
Private Sub CboLoai_Click()
    If ThemMoi <> -1 Then
        Me.MousePointer = 11
        Int_RecsetToCbo "SELECT MaSo As F2, SoHieu + Chr(9) + TenVattu As F1 FROM TP154 WHERE MaPhanLoai=" + CStr(CboLoai.ItemData(CboLoai.ListIndex)) + " ORDER BY SoHieu", LstVt
        ThemMoi = 0
        doiloai = 0
        If LstVt.ListIndex < 0 Then tp.InitTPMaSo 0
        Me.MousePointer = 0
    Else
        doiloai = 1
    End If
End Sub

Private Sub cboThang_Click()
    Dim thang As Integer, dk As Double, nvl As Double, nC As Double, m As Double, kh As Double, cp2 As Double, tcuoi As Integer
    
    If CboThang.ListIndex < 0 Then Exit Sub
    thang = CboThang.ItemData(CboThang.ListIndex)
    
    dk = tp.GiaThanhCK(IIf(thang > 0, ThangTruoc(thang), 0))
    
    If thang > 0 Then
        tcuoi = thang
    Else
        thang = pThangDauKy
        tcuoi = IIf(pThangDauKy > 1, pThangDauKy - 1, 12)
    End If
    
    txtVT(4).Text = Format(dk, Mask_0)
    nvl = tp.SoCPNVL(thang, tcuoi) + tp.SoCPNVLPB(thang, tcuoi)
    LbCP(0).Caption = Format(nvl, Mask_0)
    nC = tp.SoCPNC(thang, tcuoi) + tp.SoCPNCPB(thang, tcuoi)
    LbCP(1).Caption = Format(nC, Mask_0)
    m = tp.SoCPM(thang, tcuoi) + tp.SoCPMPB(thang, tcuoi)
    LbCP(2).Caption = Format(m, Mask_0)
    kh = tp.SoCPKH(thang, tcuoi)
    LbCP(3).Caption = Format(kh, Mask_0)
    cp2 = tp.SoCPSXC(thang, tcuoi) + tp.SoCPSXCTT(thang, tcuoi)
    LbCP(4).Caption = Format(cp2, Mask_0)
    LbCP(5).Caption = Format(dk + nvl + nC + m + cp2, Mask_0)
End Sub

Private Sub cmdtk_Click(Index As Integer)
    Me.MousePointer = 11
    txtShTk(Index).Text = FrmTaikhoan.ChonTk(txtShTk(Index).Text)
    RFocus txtShTk(Index)
    Me.MousePointer = 0
End Sub

Private Sub Command_Click(Index As Integer)
    Dim vt As New Cls154, i As Integer
    
    If (User_Right = 2) And (Index < 3) Then
        HienThongBao "Kh«ng cã quyÒn truy cËp!", 1
        Exit Sub
    End If
    
    Me.MousePointer = 11
    If Index < 3 Then
        If CboLoai.ListIndex < 0 Then
                ErrMsg er_PhanLoai
                GoTo XongVT
        End If
    End If
    
    Select Case Index
        Case 0:
            txtVT(0).Text = SoHieuVTMoi(CboLoai.ItemData(CboLoai.ListIndex), 1)
            txtVT(1).Text = ""
            
            RFocus txtVT(0)
            ThemMoi = 1
        Case 1:
            Select Case ThemMoi
                Case 1:
                    If Not KiemTraSoLieu Then GoTo XongVT
                    If tp.GhiTP = 0 Then
                        tp.GhiDK Cdbl5(txtVT(4).Text)
                        LstVt.AddItem tp.sohieu + Chr(9) + tp.TenVattu
                        LstVt.ItemData(LstVt.NewIndex) = tp.MaSo
                        LstVt.ListIndex = LstVt.NewIndex
                    Else
                        ErrMsg er_SoHieu
                        vt.InitTPSohieu txtVT(0).Text
                        If vt.MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex) Then
                            SetListIndex LstVt, vt.MaSo
                        End If
                    End If
                    ThemMoi = 0
                Case 0:
                    If LstVt.ListIndex < 0 Then GoTo XongVT
                    If Not KiemTraSoLieu Then GoTo XongVT
                    
                    If tp.SuaTP = 0 Then
                        tp.GhiDK Cdbl5(txtVT(4).Text)
                        If doiloai = 1 Then
                            CboLoai_Click
                            doiloai = 0
                        Else
                            LstVt.List(LstVt.ListIndex) = tp.sohieu + Chr(9) + tp.TenVattu
                        End If
                    Else
                        vt.InitTPSohieu txtVT(0).Text
                        ErrMsg er_SoHieu
                        If vt.MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex) Then SetListIndex LstVt, vt.MaSo
                    End If
                    ThemMoi = 0
            End Select
            RFocus LstVt
        Case 2:
            i = LstVt.ListIndex
            If i < 0 Then GoTo XongVT
            If tp.XoaTP = 0 Then
                LstVt.RemoveItem i
                If LstVt.ListCount > 0 Then LstVt.ListIndex = i - 1
            Else
                ErrMsg er_CoPS
            End If
            RFocus LstVt
        Case 3:
            Hide
    End Select
XongVT:
    Me.MousePointer = 0
End Sub

Private Sub Form_Activate()
    If Me.tag < 0 Then
        SetListIndex CboLoai, -Me.tag
        Me.tag = 0
    End If
    If ThemMoi = 0 And Me.tag = 1 Then RFocus LstVt
End Sub
'======================================================================================
' Xu ly cac phim nong
'======================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyV:
                RFocus Command(3)
                Command_Click 3
            Case vbKeyT:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyX:
                RFocus Command(2)
                Command_Click 2
            Case vbKeyG:
                RFocus Command(1)
                Command_Click 1
        End Select
    End If
    If KeyCode = vbKeyEscape Then Hide
End Sub
'======================================================================================
' Khoi tao form
'======================================================================================
Private Sub Form_Load()
    ThemMoi = 0
    doiloai = 0
    Caption = Caption + " - " + CStr(pNamTC)
    Int_RecsetToCbo "SELECT DISTINCTROW MaSo As F2, SoHieu + ' - '  + TenPhanLoai As F1 FROM PhanLoai154 WHERE PLCon=0 ORDER BY SoHieu", CboLoai
    
    AddMonthToCbo CboThang
    CboThang.AddItem "TB", 0
    CboThang.ListIndex = 0
    
    SetFont Me
    
    TK.InitTaikhoanSohieu txtShTk(0).Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tp = Nothing
End Sub

'======================================================================================
' Khoi tao vat tu duoc chon
'======================================================================================
Private Sub LstVt_Click()
    tp.InitTPMaSo LstVt.ItemData(LstVt.ListIndex)
    ShowChitiet tp
End Sub
'======================================================================================
' Thu tuc hien thong tin chi tiet
'======================================================================================
Private Sub ShowChitiet(tp As Cls154)
    
    txtVT(0).Text = tp.sohieu
    txtVT(1).Text = tp.TenVattu
    txtVT(2).Text = tp.DonVi
    txtVT(3).Text = tp.GhiChu
    TK.InitTaikhoanMaSo tp.MaTK
    txtShTk(0).Text = TK.sohieu
    LbTenTk(0).Caption = TK.Ten
    
    cboThang_Click
End Sub
'======================================================================================
' Thu tuc chon vat tu
' sh: so hieu vat tu can chon
' Tra ve so hieu vat tu duoc chon
'======================================================================================
Public Function ChonTP(sh As String) As String
    Dim mpl As Long, shtk As String
    Dim j As Integer, i As Integer, pos As Integer, Length As Integer
    
    If Len(sh) > 0 Then
        shtk = "SELECT DISTINCTROW TOP 1 MaPhanLoai AS F1 FROM TP154 WHERE SoHieu LIKE '" + sh + "*' ORDER BY SoHieu"
        mpl = SelectSQL(shtk)
        If mpl > 0 Then
            If CboLoai.ItemData(CboLoai.ListIndex) <> mpl Then SetListIndex CboLoai, mpl
        End If
         i = 0
         j = LstVt.ListCount - 1
         pos = 0
         Length = Len(sh)
         Do While i <= j - 1
                pos = Fix(0.5 + (i + j) / 2)
                shtk = Left(LstVt.List(pos), Length)
                If sh = shtk Then
                    i = pos - 1
                    Do While (sh = Left(LstVt.List(i), Length)) And (i > 0)
                        i = i - 1
                    Loop
                    pos = i + 1
                    Exit Do
                End If
                If sh > shtk Then
                    i = pos
                Else
                    If j = 1 Then
                        pos = 0
                        Exit Do
                    Else
                        If j = pos Then Exit Do
                        j = pos
                    End If
                End If
        Loop
        If LstVt.ListCount > 0 Then LstVt.ListIndex = pos
    End If
    Me.tag = 1
    On Error Resume Next
    Me.Show 1
    On Error GoTo 0
    If tp.MaSo > 0 Then
        ChonTP = tp.sohieu
    Else
        ChonTP = ""
    End If
End Function

Private Sub LstVt_DblClick()
    If Me.tag = 1 Then Hide
End Sub

Private Sub LstVt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then LstVt_DblClick
End Sub
'======================================================================================
' Thu tuc kiem tra va dua so lieu vao object
'======================================================================================
Private Function KiemTraSoLieu() As Boolean
    KiemTraSoLieu = False
    
    If Len(txtVT(0).Text) = 0 Then
        ErrMsg er_SoHieu
        RFocus txtVT(0)
        Exit Function
    End If
    
    If Len(txtVT(1).Text) = 0 Then
        ErrMsg er_Ten
        RFocus txtVT(1)
        Exit Function
    End If
    
    If Len(txtVT(2).Text) = 0 Then
        MsgBox "ThiÕu ®¬n vÞ tÝnh!", vbExclamation, App.ProductName
        RFocus txtVT(2)
        Exit Function
    End If
    
    If TK.MaSo = 0 Or TK.tkcon > 0 Or (Not TK.sohieu Like ShTkSPDo + "*") Then
        RFocus txtShTk(0)
        Exit Function
    End If
With tp
    If ThemMoi = 1 Then .MaSo = 0
    .MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex)
    .sohieu = txtVT(0).Text
    .TenVattu = txtVT(1).Text
    .DonVi = txtVT(2).Text
    .GhiChu = IIf(Len(txtVT(3).Text) > 0, txtVT(3).Text, "...")
    .MaTK = TK.MaSo
End With
    KiemTraSoLieu = True
End Function

Private Sub SSCmdF_Click()
    Dim sql As String
    
    If Len(txtF.Text) = 0 Then
        RFocus txtF
        Exit Sub
    End If
    
    Me.MousePointer = 11
    sql = "SELECT DISTINCTROW Top 1 SoHieu AS F1 FROM TP154 WHERE MaSo>" + CStr(MaDaTim) + IIf(SSOpt(0).Value, " AND SoHieu LIKE '" + txtF.Text + "'", " AND InStr(TenVattu,'" + txtF.Text + "')>0")
    sql = CStr(SelectSQL(sql))
    If sql <> "0" Then
        ChonTP sql
        MaDaTim = tp.MaSo
    Else
        MaDaTim = 0
    End If
    Me.MousePointer = 0
End Sub

Private Sub txtF_GotFocus()
    AutoSelect txtF
    MaDaTim = 0
End Sub

Private Sub txtShTk_GotFocus(Index As Integer)
    AutoSelect txtShTk(Index)
End Sub

Private Sub txtShTk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdtk_Click Index
End Sub

Private Sub txtShTk_LostFocus(Index As Integer)
    txtShTk(Index).Text = UCase(txtShTk(Index).Text)
    TK.InitTaikhoanSohieu txtShTk(Index).Text
    LbTenTk(Index).Caption = TK.Ten
End Sub

Private Sub Txtvt_GotFocus(Index As Integer)
    AutoSelect txtVT(Index)
End Sub

Private Sub TxtVT_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0: If KeyAscii = 32 Or KeyAscii = 35 Or KeyAscii = 39 Or KeyAscii = 42 Then KeyAscii = 0
        Case 4: KeyProcess txtVT(Index), KeyAscii
    End Select
End Sub

Private Sub TxtVT_LostFocus(Index As Integer)
    Select Case Index
        Case 0:
            txtVT(Index).Text = UCase(txtVT(Index).Text)
        Case 1, 2, 3:
            If Len(txtVT(Index).Text) = 0 Then txtVT(Index).Text = "..."
        Case 4:
            txtVT(Index).Text = Format(txtVT(Index).Text, Mask_0)
    End Select
End Sub
