VERSION 5.00
Begin VB.Form FrmKho 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kho vËt t­, thµnh phÈm, ®¹i lý..."
   ClientHeight    =   3870
   ClientLeft      =   4170
   ClientTop       =   2445
   ClientWidth     =   4230
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
   Icon            =   "Frmkho.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3870
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "List of Items"
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   4
      Left            =   3000
      Picture         =   "Frmkho.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "&Select"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   3000
      Picture         =   "Frmkho.frx":6C44
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Return"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   3000
      Picture         =   "Frmkho.frx":8066
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Delete"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   3000
      Picture         =   "Frmkho.frx":9548
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "&Save"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   3000
      Picture         =   "Frmkho.frx":A976
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&Add"
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox LstKho 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000012&
      Height          =   3150
      ItemData        =   "Frmkho.frx":ACB8
      Left            =   120
      List            =   "Frmkho.frx":ACBA
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtTenkho 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   0
      Top             =   3440
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "BÊm Ctrl-F ®Ó t×m kiÕm"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "FrmKho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim flag As Integer       '=1 nÏu nhÄp mêi
                                             '=0 nÏu Åang sôa
Dim TenTB As String
Dim TenFL As String
Dim f1 As Integer
Dim MaChon As Long

'====================================================================================================
' Thªm, Ghi, Xãa kho
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim i As Integer, tenkho As String, mk As Long, luong As Double, tien As Double
    
    Select Case Index
        Case 0:
            flag = 1
            txtTenkho = ""
            RFocus txtTenkho
        Case 1:
            If Len(txtTenkho.Text) = 0 Then
                RFocus txtTenkho
                Exit Sub
            End If
            Me.MousePointer = 11
            Select Case flag
                Case 1:
                     If ExecuteSQL5("INSERT INTO " + TenTB + " (MaSo, " + TenFL + ") VALUES (" + CStr(Lng_MaxValue("MaSo", TenTB) + 1) + ",'" + txtTenkho + "')") <> 0 Then GoTo XongKho
                     LstKho.AddItem txtTenkho.Text
                     LstKho.ItemData(LstKho.NewIndex) = Lng_MaxValue("MaSo", TenTB)
                     flag = 0
                     LstKho.ListIndex = LstKho.NewIndex
                Case 0:
                     If LstKho.ListIndex < 0 Then GoTo XongKho
                     If txtTenkho.Text = LstKho.List(LstKho.ListIndex) Then GoTo XongKho
                     If ExecuteSQL5("UPDATE " + TenTB + " SET " + TenFL + "='" + txtTenkho + "' WHERE MaSo=" + CStr(LstKho.ItemData(LstKho.ListIndex))) <> 0 Then GoTo XongKho
                     LstKho.List(LstKho.ListIndex) = txtTenkho.Text
            End Select
            RFocus txtTenkho
        Case 2:
            Select Case flag
                Case 0:
                     i = LstKho.ListIndex
                    If i < 0 Then GoTo XongKho
                    If f1 = 1 And SelectSQL("SELECT MaSo AS F1 FROM TonKho WHERE MaSoKho=" + CStr(LstKho.ItemData(i))) > 0 Then
                        tenkho = FrmGetStr.GetString("ChuyÓn c¸c ph¸t sinh ®· cã sang kho", App.ProductName)
                        mk = SelectSQL("SELECT MaSo AS F1 FROM KhoHang WHERE TenKho='" + tenkho + "'")
                        If mk = 0 Or mk = LstKho.ItemData(i) Then GoTo XongKho
                        Me.MousePointer = 11
                        ChuyenKho LstKho.ItemData(i), mk
                    End If
                    If f1 = 5 And pNhapKhau > 0 Then
                        ExecuteSQL5 "DELETE * FROM CPGVHD WHERE MaSo=" + CStr(LstKho.ItemData(i))
                    End If
                    If ExecuteSQL5("DELETE * FROM " + TenTB + " WHERE MaSo=" + CStr(LstKho.ItemData(i))) <> 0 Then GoTo XongKho
                    LstKho.RemoveItem i
                    If LstKho.ListCount > 0 Then LstKho.ListIndex = i - 1
                Case 1:
                    flag = 0
                    LstKho_Click
            End Select
        Case 3:
            Unload Me
        Case 4:
            If LstKho.ListIndex < 0 Then Exit Sub
            MaChon = LstKho.ItemData(LstKho.ListIndex)
            Unload Me
    End Select
XongKho:
    Me.MousePointer = 0
End Sub

Private Sub Form_Activate()
    If Me.tag > 0 Then
        f1 = Me.tag
        Me.tag = 0
        Select Case f1
            Case 1:
                If pNN = 0 Then Caption = "Danh s¸ch kho vËt t­, hµng ho¸"
                TenTB = "KhoHang"
                TenFL = "TenKho"
            Case 2:
                If pNN = 0 Then Caption = "Danh s¸ch n­íc s¶n xuÊt TSC§"
                TenTB = "QuocGia"
                TenFL = "Ten"
            Case 3:
                If pNN = 0 Then Caption = "T×nh tr¹ng sö dông TSC§"
                TenTB = "TinhTrang"
                TenFL = "Ten"
            Case 4:
                If pNN = 0 Then Caption = "Danh s¸ch ®¬n vÞ qu¶n lý TSC§"
                TenTB = "DTQLy"
                TenFL = "Ten"
            Case 5:
                If pNN = 0 Then Caption = "C¸c vô viÖc liªn quan cña chøng tõ"
                TenTB = "DoituongCT"
                TenFL = "DienGiai"
            Case 6:
                If pNN = 0 Then Caption = "§¨ng ký chøng tõ ghi sæ"
                TenTB = "CTGhiSo"
                TenFL = "SoHieu"
                txtTenkho.MaxLength = 20
            Case 10, 11, 12:
                If pNN = 0 Then Caption = "Th«ng tin chøng tõ " + CStr(f1 - 9)
                TenTB = "DoituongCT" + CStr(f1 - 9)
                TenFL = "DienGiai"
        End Select
        If pNN = 0 And pKhongDau > 0 Then Me.Caption = ABCtoKDau(Me.Caption)
        Int_RecsetToCbo "SELECT MaSo As F2," + TenFL + " As F1 FROM " + TenTB + IIf(f1 = 5, " WHERE MaSo > 1 AND MaKhachHang=0", "") + IIf(f1 = 6, " WHERE MaSo > 1", "") + " ORDER BY " + TenFL, LstKho
        If f1 = 10 And MaChon > 0 Then
            Caption = Caption + " liªn quan"
            Me.Top = Screen.Height / 2
            Me.Left = Screen.Width / 2
            SetListIndex LstKho, MaChon
        End If
    End If
End Sub

'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
            Select Case KeyCode
                Case vbKeyV:
                    RFocus Command(3)
                    Command_Click 3
                Case vbKeyG:
                    RFocus Command(1)
                    Command_Click 1
                Case vbKeyT:
                    RFocus Command(0)
                    Command_Click 0
                Case vbKeyX:
                    RFocus Command(2)
                    Command_Click 2
            End Select
    End If
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyF And LstKho.ListCount > 0 Then
        SearchObj 0, LstKho
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
'====================================================================================================
' Khëi t¹o cöa sæ
'====================================================================================================
Private Sub Form_Load()
    flag = 0
    SetFont Me
End Sub

Private Sub LstKho_Click()
    flag = 0
    txtTenkho.Text = LstKho.List(LstKho.ListIndex)
End Sub

Private Sub LstKho_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim m As Long, shtk As String, TK As ClsTaikhoan
    
    If Button = 2 And LstKho.ListIndex >= 0 And flag = 0 And f1 = 1 And pGiaVon > 0 Then
        Set TK = New ClsTaikhoan
        'ThemTruong "KhoHang", "MaTK", dbLong
        'ThemTruong "KhoHang", "MaTKGV", dbLong
        m = LstKho.ItemData(LstKho.ListIndex)
        shtk = SelectSQL("SELECT HethongTK.SoHieu AS F1 FROM KhoHang INNER JOIN HethongTK ON KhoHang.MaTKGV=HethongTK.MaSo WHERE KhoHang.MaSo=" + CStr(m))
        If shtk = "0" Then shtk = ""
        shtk = FrmGetStr.GetString("Sè hiÖu tµi kho¶n tÝnh gi¸ vèn tù ®éng: ", App.ProductName, shtk)

        TK.InitTaikhoanSohieu shtk
        If (TK.MaSo > 0 And TK.tkcon = 0 And TK.tk_id <> TKVT_ID And TK.tk_id <> TKDT_ID And TK.tk_id <> GTGTKT_ID And TK.tk_id <> GTGTPN_ID) Or (UCase(shtk) = "0") Then
            ExecuteSQL5 "UPDATE KhoHang SET MaTKGV=" + CStr(TK.MaSo) + " WHERE MaSo=" + CStr(m)
        Else
            If Len(shtk) > 0 Then ErrMsg er_SHTaiKhoan1
        End If
        
        shtk = SelectSQL("SELECT HethongTK.SoHieu AS F1 FROM KhoHang INNER JOIN HethongTK ON KhoHang.MaTK=HethongTK.MaSo WHERE KhoHang.MaSo=" + CStr(m))
        If shtk = "0" Then shtk = ""
        shtk = FrmGetStr.GetString("Sè hiÖu tµi kho¶n tÝnh gi¸ vèn tù ®éng hµng khuyÕn m¹i: ", App.ProductName, shtk)

        TK.InitTaikhoanSohieu shtk
        If (TK.MaSo > 0 And TK.tkcon = 0 And TK.tk_id <> TKVT_ID And TK.tk_id <> TKDT_ID And TK.tk_id <> GTGTKT_ID And TK.tk_id <> GTGTPN_ID) Or (UCase(shtk) = "0") Then
            ExecuteSQL5 "UPDATE KhoHang SET MaTK=" + CStr(TK.MaSo) + " WHERE MaSo=" + CStr(m)
        Else
            If Len(shtk) > 0 Then ErrMsg er_SHTaiKhoan1
        End If

        Set TK = Nothing
    End If
    
    If Button = 2 And LstKho.ListIndex >= 0 And flag = 0 And f1 = 11 Then
        Dim f As New FrmKho
        m = SelectSQL("SELECT MaKhachHang AS F1 FROM DoituongCT2 WHERE MaSo=" + CStr(LstKho.ItemData(LstKho.ListIndex)))
        m = f.ChonKho(10, m)
        If m > 0 Then ExecuteSQL5 "UPDATE DoiTuongCT2 SET MaKhachHang=" + CStr(m) + " WHERE MaSo=" + CStr(LstKho.ItemData(LstKho.ListIndex))
        Set f = Nothing
    End If
End Sub

Private Sub txtTenkho_GotFocus()
    AutoSelect txtTenkho
End Sub

Public Function ChonKho(id As Long, mc As Long) As Long
    Dim i As Integer
    
    Me.tag = id
    Command(4).Visible = True
    For i = 0 To 2
        Command(i).Visible = False
    Next
    MaChon = mc
    'Me.StartUpPosition = 0
    Me.Show 1
    ChonKho = MaChon
End Function
