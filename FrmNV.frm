VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmNhanVien 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch Nh©n viªn b¸n hµng"
   ClientHeight    =   7080
   ClientLeft      =   1080
   ClientTop       =   1005
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
   Icon            =   "FrmNV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Salesman List"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   360
      Left            =   5850
      TabIndex        =   34
      Top             =   1860
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Left            =   3120
      TabIndex        =   17
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "..."
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   5760
      MaxLength       =   100
      TabIndex        =   3
      Text            =   "..."
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   5
      Text            =   "..."
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   5760
      MaxLength       =   100
      TabIndex        =   4
      Text            =   "..."
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   6
      Text            =   "..."
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   8040
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "..."
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   8
      Text            =   "..."
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   5760
      MaxLength       =   100
      TabIndex        =   9
      Text            =   "..."
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox txtVT 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   5760
      MaxLength       =   100
      TabIndex        =   10
      Text            =   "..."
      Top             =   5880
      Width           =   3615
   End
   Begin VB.OptionButton SSOpt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CMND"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   20
      Tag             =   "Id"
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   3
      Left            =   8640
      Picture         =   "FrmNV.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   2
      Left            =   7440
      Picture         =   "FrmNV.frx":6C04
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "&Delete"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   1
      Left            =   6240
      Picture         =   "FrmNV.frx":80E6
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "&Save"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   0
      Left            =   5040
      Picture         =   "FrmNV.frx":9514
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "&Add"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton SSCmdF 
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   21
      Top             =   6720
      Width           =   255
   End
   Begin VB.OptionButton SSOpt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tªn "
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   19
      Tag             =   "Name"
      Top             =   6720
      Width           =   855
   End
   Begin VB.OptionButton SSOpt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sè hiÖu"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Tag             =   "Code"
      Top             =   6720
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.ComboBox CboLoai 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.ListBox LstVt 
      Height          =   5910
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè hiÖu"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   33
      Tag             =   "Code"
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tªn"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   32
      Tag             =   "Name"
      Top             =   960
      Width           =   735
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   5760
      X2              =   7800
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line 
      Index           =   1
      X1              =   5760
      X2              =   9360
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "§Þa chØ"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   31
      Tag             =   "Address"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Line Line 
      Index           =   2
      X1              =   5760
      X2              =   9360
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè CMND"
      Height          =   255
      Index           =   10
      Left            =   4920
      TabIndex        =   30
      Tag             =   "ID"
      Top             =   2400
      Width           =   735
   End
   Begin VB.Line Line 
      Index           =   3
      X1              =   5760
      X2              =   7800
      Y1              =   2685
      Y2              =   2685
   End
   Begin VB.Line Line 
      Index           =   4
      X1              =   5760
      X2              =   7080
      Y1              =   3165
      Y2              =   3165
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tel 1"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   29
      Top             =   2880
      Width           =   735
   End
   Begin VB.Line Line 
      Index           =   5
      X1              =   8040
      X2              =   9360
      Y1              =   3165
      Y2              =   3165
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tel 2"
      Height          =   255
      Index           =   11
      Left            =   7440
      TabIndex        =   28
      Top             =   2880
      Width           =   375
   End
   Begin VB.Line Line 
      Index           =   7
      X1              =   5760
      X2              =   7080
      Y1              =   3645
      Y2              =   3645
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email"
      Height          =   255
      Index           =   13
      Left            =   4920
      TabIndex        =   27
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ngµy sinh"
      Height          =   255
      Index           =   14
      Left            =   4920
      TabIndex        =   26
      Tag             =   "Birth Date"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Line Line 
      Index           =   9
      X1              =   5760
      X2              =   9360
      Y1              =   4125
      Y2              =   4125
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  T.kho¶n"
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   25
      Tag             =   "Bank Acc."
      Top             =   3840
      Width           =   855
   End
   Begin VB.Line Line 
      Index           =   10
      X1              =   5760
      X2              =   9360
      Y1              =   6165
      Y2              =   6165
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ghi chó"
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   24
      Tag             =   "Notes"
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   17
      Left            =   4680
      TabIndex        =   23
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   5
      Left            =   4740
      TabIndex        =   16
      Top             =   180
      Width           =   4935
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   4
      Left            =   4800
      TabIndex        =   15
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label 
      BackColor       =   &H00808080&
      Height          =   5910
      Index           =   9
      Left            =   165
      TabIndex        =   14
      Top             =   540
      Width           =   4455
   End
End
Attribute VB_Name = "FrmNhanVien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer          ' =1 neu them moi, -1 neu sua cu
Dim onv As New ClsNhanVien       ' vat tu duoc tham chieu
Dim doiloai As Integer               ' =1 neu co thay doi loai vat tu dang sua doi
Dim MaDaTim As Long
Dim ngay As Date
'======================================================================================
' Liet ke cac vat tu trong loai vat tu duoc chon
'======================================================================================
Private Sub CboLoai_Click()
    If ThemMoi <> -1 Then
        Me.MousePointer = 11
        Int_RecsetToCbo "SELECT MaSo As F2, SoHieu + ' - ' + Ten As F1 FROM NhanVien WHERE MaPhanLoai=" + CStr(CboLoai.ItemData(CboLoai.ListIndex)) + " ORDER BY SoHieu", LstVt
        ThemMoi = 0
        doiloai = 0
        Me.MousePointer = 0
    Else
        doiloai = 1
    End If
End Sub

Private Sub Command_Click(Index As Integer)
    Dim vt As New ClsNhanVien, i As Integer
    
    If (User_Right = 2) And (Index < 3) Then
        HienThongBao "Kh«ng cã quyÒn truy cËp!", 1
        GoTo XongVT
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
            txtVT(0).Text = SoHieuVTMoi(CboLoai.ItemData(CboLoai.ListIndex), 3)
            For i = 1 To 8
                txtVT(i).Text = "..."
            Next

            RFocus txtVT(0)
            ThemMoi = 1
        Case 1:
            Select Case ThemMoi
                Case 1:
                    If Not KiemTraSoLieu Then GoTo XongVT
                    If onv.GhiNV = 0 Then
                        LstVt.AddItem onv.sohieu + " - " + onv.Ten
                        LstVt.ItemData(LstVt.NewIndex) = onv.MaSo
                        LstVt.ListIndex = LstVt.NewIndex
                    Else
                        ErrMsg er_SoHieu
                        vt.InitNVSoHieu txtVT(0).Text
                        If vt.MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex) Then
                            SetListIndex LstVt, vt.MaSo
                        End If
                    End If
                    ThemMoi = 0
                Case 0:
                    If LstVt.ListIndex < 0 Then GoTo XongVT
                    If Not KiemTraSoLieu Then GoTo XongVT
                    
                    If onv.SuaNV = 0 Then
                        If doiloai = 1 Then
                            CboLoai_Click
                            doiloai = 0
                        Else
                            LstVt.List(LstVt.ListIndex) = onv.sohieu + " - " + onv.Ten
                        End If
                    Else
                        vt.InitNVSoHieu txtVT(0).Text
                        ErrMsg er_SoHieu
                        If vt.MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex) Then SetListIndex LstVt, vt.MaSo
                    End If
                    ThemMoi = 0
            End Select
            RFocus LstVt
        Case 2:
            i = LstVt.ListIndex
            If i < 0 Then GoTo XongVT
            If onv.XoaNV = 0 Then
                LstVt.RemoveItem i
                If LstVt.ListCount > 0 Then LstVt.ListIndex = i - 1
            Else
                ErrMsg er_CoPS1
            End If
            RFocus LstVt
        Case 3:
            Hide
    End Select
XongVT:
    Set vt = Nothing
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
    Int_RecsetToCbo "SELECT DISTINCTROW MaSo As F2,SoHieu + ' - '  + TenPhanLoai As F1 FROM PhanLoaiNhanVien WHERE PLCon=0 ORDER BY SoHieu", CboLoai
    InitDateVars MedNgay, ngay
    SetFont Me
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set onv = Nothing
End Sub

'======================================================================================
' Khoi tao vat tu duoc chon
'======================================================================================
Private Sub LstVt_Click()
    onv.InitNVMaSo LstVt.ItemData(LstVt.ListIndex)
    ShowChitiet onv
End Sub
'======================================================================================
' Thu tuc hien thong tin chi tiet
'======================================================================================
Private Sub ShowChitiet(otk As ClsNhanVien)
    txtVT(0).Text = onv.sohieu
    txtVT(1).Text = onv.Ten
    txtVT(2).Text = onv.DiaChi
    txtVT(3).Text = onv.CMND
    txtVT(4).Text = onv.Tel1
    txtVT(5).Text = onv.Tel2
    txtVT(6).Text = onv.email
    txtVT(7).Text = onv.taikhoan
    txtVT(8).Text = onv.GhiChu
    ngay = onv.NgaySinh
    MedNgay.Text = Format(ngay, Mask_D)
End Sub
'======================================================================================
' Thu tuc chon vat tu
' sh: so hieu vat tu can chon
' Tra ve so hieu vat tu duoc chon
'======================================================================================
Public Function ChonNV(sh As String) As String
    Dim mpl As Long, shtk As String
    Dim j As Integer, i As Integer, pos As Integer, Length As Integer
    
    If Len(sh) > 0 Then
        shtk = "SELECT DISTINCTROW TOP 1 MaPhanLoai AS F1 FROM NhanVien WHERE SoHieu LIKE '" + sh + "*' ORDER BY SoHieu"
        mpl = SelectSQL(shtk)
        If mpl > 0 And CboLoai.ListIndex >= 0 Then
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
    If onv.MaSo > 0 Then
        ChonNV = onv.sohieu
    Else
        ChonNV = ""
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
    
With onv
    If ThemMoi = 1 Then .MaSo = 0
    .MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex)
    .sohieu = txtVT(0).Text
    .Ten = txtVT(1).Text
    .DiaChi = txtVT(2).Text
    .NgaySinh = ngay
    .CMND = txtVT(3).Text
    .Tel1 = txtVT(4).Text
    .Tel2 = txtVT(5).Text
    .email = txtVT(6).Text
    .taikhoan = txtVT(7).Text
    .GhiChu = txtVT(8).Text
    If .CMND <> "..." And SelectSQL("SELECT MaSo AS F1 FROM NhanVien WHERE CMND='" + .CMND + "' AND MaSo<>" + CStr(.MaSo)) > 0 Then
        If MsgBox("Sè CMND ®· cã, cho phÐp nhËp?", vbYesNo + vbCritical, App.ProductName) = vbNo Then Exit Function
    End If
End With
    KiemTraSoLieu = True
End Function

Private Sub LstVt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sh As String, m As Long
    
    If Button = 2 And LstVt.ListIndex >= 0 And ThemMoi = 0 Then
        sh = FrmGetStr.GetString("ChuyÓn " + VString(onv.sohieu + " - " + onv.Ten) + " sang ph©n lo¹i cã sè hiÖu:", App.ProductName)
        If Len(sh) > 0 Then
            m = SelectSQL("SELECT MaSo AS F1 FROM PhanLoaiNhanVien WHERE PLCon=0 AND SoHieu='" + sh + "'")
            If m > 0 And m <> onv.MaPhanLoai Then
                ExecuteSQL5 "UPDATE NhanVien SET MaPhanLoai=" + CStr(m) + " WHERE MaSo = " + CStr(onv.MaSo)
                CboLoai_Click
            End If
        End If
    End If
End Sub

Private Sub MedNgay_LostFocus()
    If IsDate(MedNgay.Text) Then
        ngay = CDate(MedNgay.Text)
    Else
        RFocus MedNgay
    End If
End Sub

Private Sub SSCmdF_Click()
    Dim sql As String
    
    If Len(txtF.Text) = 0 Then
        RFocus txtF
        Exit Sub
    End If
    
    Me.MousePointer = 11
    sql = "SELECT DISTINCTROW Top 1 SoHieu AS F1 FROM NhanVien WHERE MaSo>" + CStr(MaDaTim)
    If SSOpt(0).Value Then sql = sql + " AND SoHieu LIKE '" + txtF.Text + "*'"
    If SSOpt(1).Value Then sql = sql + " AND InStr(Ten,'" + txtF.Text + "')>0"
    If SSOpt(2).Value Then sql = sql + " AND CMND LIKE '" + txtF.Text + "*'"
    sql = CStr(SelectSQL(sql))
    If sql <> "0" Then
        ChonNV sql
        MaDaTim = onv.MaSo
    Else
        MaDaTim = 0
    End If
    Me.MousePointer = 0
End Sub

Private Sub txtF_GotFocus()
    AutoSelect txtF
    MaDaTim = 0
End Sub

Private Sub Txtvt_GotFocus(Index As Integer)
    AutoSelect txtVT(Index)
End Sub

Private Sub TxtVT_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0: If KeyAscii = 32 Or KeyAscii = 35 Or KeyAscii = 39 Or KeyAscii = 42 Then KeyAscii = 0
    End Select
End Sub

Private Sub TxtVT_LostFocus(Index As Integer)
    If Len(txtVT(Index).Text) = 0 Then txtVT(Index).Text = "..."
End Sub

Private Sub MedNgay_GotFocus()
    AutoSelect MedNgay
End Sub
