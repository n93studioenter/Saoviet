VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FThuChi 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Th«ng tin vÒ phiÕu thu - chi"
   ClientHeight    =   2895
   ClientLeft      =   3780
   ClientTop       =   4050
   ClientWidth     =   7740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FThuChi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Additional Voucher Information"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.CheckBox Checkinbangkevahoadon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "In hãa ®¬n kÌm b¶ng kª"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Tag             =   "Direct Export"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CheckBox checkinbangke 
      BackColor       =   &H00FFFFFF&
      Caption         =   "In  b¶ng kª"
      Height          =   255
      Left            =   6000
      TabIndex        =   22
      Tag             =   "Direct Export"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckBox3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lien 3"
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   3000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox CheckBox2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lien 2"
      Height          =   375
      Left            =   4560
      TabIndex        =   20
      Top             =   3000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox CheckBox1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lien 1"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3240
      MaskColor       =   &H8000000A&
      TabIndex        =   19
      Top             =   2880
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "In mÉu hoa ®¬n"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Tag             =   "Direct Export"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Néi dung theo diÔn gi¶i"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Tag             =   "Direct Export"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   4
      Left            =   2040
      MaxLength       =   150
      TabIndex        =   13
      Text            =   "..."
      Top             =   120
      Width           =   5415
   End
   Begin VB.TextBox T 
      Height          =   345
      Index           =   3
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "..."
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdkh 
      Height          =   375
      Left            =   3360
      Picture         =   "FThuChi.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtshkh 
      Height          =   345
      Left            =   2040
      LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
      MaxLength       =   20
      TabIndex        =   5
      Tag             =   "0"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox T 
      Height          =   315
      Index           =   2
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "..."
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   1
      Text            =   "..."
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "..."
      Top             =   480
      Width           =   5415
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   6480
      Picture         =   "FThuChi.frx":5C5C
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "&Save"
      Top             =   2400
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tªn C«ng ty:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Tag             =   "Name of receiver,payer:"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "H×nh thøc thanh to¸n"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Tag             =   "Object code"
      Top             =   2160
      Width           =   1575
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   330
      Left            =   2040
      TabIndex        =   16
      Top             =   2160
      Width           =   1695
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2990;591"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "VNI-Times"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lbkh 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Tag             =   "1"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè hiÖu ®èi t­îng"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Tag             =   "Object code"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè hãa ®¬n:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Tag             =   "Number of Voucher"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "§Þa chØ :"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Tag             =   "Address:"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tªn ng­êi nép tiÒn:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Tag             =   "Name of receiver,payer:"
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FThuChi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FThuChiForm As Integer


Dim s(0 To 3) As String
Dim kh As New ClsKhachHang
Dim ngay As Date
Dim f1 As Integer
Public Sub Test()
 
    Unload Me
End Sub
Private Sub Timer1_Timer()
    Unload Me ' Ðóng form sau khi Timer h?t th?i gian
End Sub

Public Sub Command_Click()
    Dim i As Integer
    '  ExecuteSQL5 "update chungtu set nguoimuahang = '" + T(4).Text + "'  where sohieu = '" + FrmChungtu.txt(0).Text + "'"
    ExecuteSQL5 "update chungtu set nguoimuahang = '" + T(4).Text + "',hinhthucthanhtoan = '" + ComboBox1.Text + "',sophieudathang = '" + T(3).Text + "' ,chondiengiai = '" + str(Check1.Value) + "'  where sohieu = '" + FrmChungtu.txt(0).Text + "'"
    For i = 0 To 3
        s(i) = T(i).Text
    Next
    FrmChungtu.Check1.Value = Check1.Value
    FrmChungtu.Check2.Value = Check2.Value
    FrmChungtu.hinhthucthanhtoan.Text = ComboBox1.Text + "  "
    FrmChungtu.thoihanthanhtoan.Text = MedNgay.Text
    FrmChungtu.sochungtu = T(3).Text
    FrmChungtu.CheckBox1 = CheckBox1.Value
    FrmChungtu.CheckBox2 = CheckBox2.Value
    FrmChungtu.CheckBox3 = CheckBox3.Value
    FrmChungtu.checkinbangke.Value = checkinbangke.Value
    FrmChungtu.Checkinbangkevahoadon.Value = Checkinbangkevahoadon.Value
    Unload Me
    If FThuChiForm = 1 Then
        FrmChungtu.DoneSetup
    End If

    If FThuChiForm = 2 Then
        FrmChungtu.Timer5.Enabled = True
    End If
End Sub

Private Sub Form_Activate()


    If Me.tag > 0 Then
        f1 = Me.tag
        Select Case f1
            Case 1:
                Label1(0).Caption = "Tªn ng­êi nhËn tiÒn:"
                Label1(1).Caption = "§Þa chØ ng­êi nhËn tiÒn:"
                T(2).Text = FrmChungtu.txt(0).Text
            Case 2:
                Me.Caption = "GiÊy Uû nhiÖm chi"
                Label1(0).Caption = "Tªn ®¬n vÞ nhËn tiÒn:"
                Label1(1).Caption = "Sè tµi kho¶n:"
                Label1(2).Caption = "T¹i Ng©n hµng:"
                T(0).MaxLength = 50
                T(2).MaxLength = 50
            Case 3:
                Me.Caption = "Ho¸ ®¬n b¸n hµng"
                Label1(0).Caption = "Tªn ng­êi mua hµng:"
                Label1(1).Caption = "§Þa chØ:"
                Label1(2).Caption = "H¹n thanh to¸n:"
                Label1(3).Caption = "Sè phiÕu ®Æt hµng:"
                T(2).Visible = False
                T(3).Visible = True
                txtshkh.Visible = False
                cmdkh.Visible = False
                lbkh.Visible = False
                MedNgay.Visible = True
            Case 10:
                Me.Caption = "Th«ng tin b¸o c¸o"
                Label1(0).Caption = "Ng­êi lËp biÓu:"
                Label1(1).Caption = "KÕ to¸n tr­ëng:"
                Label1(2).Caption = "Gi¸m ®èc:"
                Label1(3).Visible = False
                txtshkh.Visible = False
                cmdkh.Visible = False
                lbkh.Visible = False
                MedNgay.Visible = False
        End Select
        Me.tag = 0
    End If
    RFocus Command
    
    ComboBox1.AddItem ("Tieàn maët")
    ComboBox1.AddItem ("Chuyeån khoaûn")
    ComboBox1.AddItem ("Coâng nôï")
    ComboBox1.AddItem ("TM/CK")
  
    
        Dim sql As String
    Dim rs_chungtu As Recordset
    sql = "SELECT iif(Nguoimuahang is null ,'...',Nguoimuahang) as aa1,"
    sql = sql + "iif(hinhthucthanhtoan is null ,'...',hinhthucthanhtoan) as bb,"
    sql = sql + "iif(sophieudathang is null ,'...',sophieudathang) as cc,"
    sql = sql + "iif(chondiengiai is null ,'0',chondiengiai) as chon1  from chungtu where sohieu = '" + FrmChungtu.txt(0).Text + "'"
        Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        If rs_chungtu.recordCount > 0 Then
           T(4).Text = rs_chungtu!AA1
           T(3).Text = rs_chungtu!cc
           If rs_chungtu!chon1 = "2" Then
           Check1.Value = 1
           End If
           
           ComboBox1.Text = rs_chungtu!bb
        End If
        If FThuChiForm = 1 Then
        Command_Click
        End If
        If FThuChiForm = 2 Then
        Command_Click
        End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 0 To 3
        s(i) = "..."
    Next
    ngay = CVDate("01/01/1900")
    
    SetFont Me
End Sub

Private Sub T_GotFocus(Index As Integer)
    AutoSelect T(Index)
End Sub

Private Sub T_LostFocus(Index As Integer)
    If Len(T(Index).Text) = 0 Then T(Index).Text = "..."
End Sub

Public Sub GetPhieu(s1 As String, s2 As String, s3 As String, makh As Long, Optional d As Date, Optional s4 As String)
    kh.InitKhachHangMaSo makh
    T(0).Text = s1
    T(1).Text = s2
    'T(2).Text = s3
    T(2).Text = FrmChungtu.txt(0).Text
    T(3).Text = s4
    txtshkh.Text = kh.sohieu
    lbkh.Caption = kh.Ten
    ngay = d
    If Year(d) > 1900 Then MedNgay.Text = Format(d, Mask_D)
    If Not Me.Visible Then
        Me.Show vbModal
    End If
    s1 = s(0)
    s2 = s(1)
    s3 = s(2)
    s4 = s(3)
    makh = kh.MaSo
    d = ngay
    Set kh = Nothing
    If FThuChiForm = 1 Then
        Command_Click
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyG:
                RFocus Command
                Command_Click
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtshkh_GotFocus()
    AutoSelect txtshkh
End Sub

Private Sub txtshkh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdkh_Click
End Sub

Private Sub txtshkh_LostFocus()
    Dim xxx As String, i As Integer
    
    kh.InitKhachHangSohieu txtshkh
    lbkh.Caption = kh.Ten
    If Len(T(0).Text) = 0 Or T(0).Text = "..." Then T(0).Text = kh.Ten
    Select Case f1
        Case 0, 1:
            If (Len(T(1).Text) = 0 Or T(1).Text = "...") And kh.DiaChi <> "..." Then T(1).Text = kh.DiaChi
        Case 2:
            xxx = Trim(LaySH(kh.taikhoan, 1, "-"))
            If (Len(T(1).Text) = 0 Or T(1).Text = "...") And IsNumeric(Left(xxx, 2)) Then T(1).Text = xxx
            i = Len(kh.taikhoan) - Len(xxx)
            If i > 0 Then
                xxx = Right(kh.taikhoan, i - 1)
                If (Len(T(2).Text) = 0 Or T(2).Text = "...") And Len(xxx) > 0 Then T(2).Text = xxx
            End If
    End Select
End Sub

Private Sub cmdkh_Click()
    Me.MousePointer = 11
    txtshkh.Text = FrmKhachHang.ChonKhachHang(txtshkh.Text)
    Me.MousePointer = 0
    RFocus txtshkh
End Sub

Private Sub MedNgay_GotFocus()
    AutoSelect MedNgay
End Sub

Private Sub MedNgay_LostFocus()
    If MedNgay.Text <> "__/__/__" Then
        If IsDate(MedNgay.Text) Then
            ngay = CDate(MedNgay.Text)
        Else
            RFocus MedNgay
        End If
    Else
        ngay = CVDate("01/01/1900")
    End If
End Sub

