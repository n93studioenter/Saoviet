VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCDT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KÕ to¸n chñ ®Çu t­"
   ClientHeight    =   7905
   ClientLeft      =   450
   ClientTop       =   2025
   ClientWidth     =   6630
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
   Icon            =   "FrmCDT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B¸o c¸o quyÕt to¸n vèn ®Çu t­ hoµn thµnh - MÉu sè 09/QTDA"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   25
      Top             =   6720
      Width           =   5775
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "T×nh h×nh c«ng nî vµ B¶ng ®èi chiÕu vèn - MÉu sè 07 vµ 08/QTDA"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   24
      Top             =   6360
      Width           =   5775
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TSC§ míi t¨ng vµ tµi s¶n l­u ®éng bµn giao - MÉu sè 05 vµ 06/QTDA"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   23
      Top             =   6000
      Width           =   5775
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Thùc hiÖn ®Çu t­ theo dù ¸n, c«ng tr×nh, h¹ng môc c«ng tr×nh"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   5055
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ThuyÕt minh b¸o c¸o Tµi chÝnh"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   4095
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Thùc hiÖn ®Çu t­"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   4095
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nguån vèn ®Çu t­"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   4095
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DuyÖt quyÕt to¸n vèn ®Çu t­ hoµn thµnh - MÉu sè 10/QTDA"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   26
      Top             =   7080
      Width           =   5055
   End
   Begin VB.TextBox txtShVT 
      Height          =   285
      Index           =   0
      Left            =   3720
      LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
      MaxLength       =   20
      TabIndex        =   14
      Tag             =   "0"
      Top             =   2820
      Width           =   1335
   End
   Begin VB.CommandButton cmdvt 
      Height          =   375
      Index           =   0
      Left            =   5160
      Picture         =   "FrmCDT.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Danh ®iÓm vËt t­"
      Top             =   2760
      Width           =   375
   End
   Begin VB.ComboBox CboNK 
      Height          =   315
      ItemData        =   "FrmCDT.frx":5960
      Left            =   5880
      List            =   "FrmCDT.frx":5979
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   0
      Left            =   1080
      LinkItem        =   "Sè hiÖu chi tiÕt cÇn xem"
      MaxLength       =   12
      TabIndex        =   6
      Tag             =   "0"
      Top             =   1005
      Width           =   1335
   End
   Begin VB.CommandButton cmdtk 
      Height          =   375
      Index           =   0
      Left            =   2520
      Picture         =   "FrmCDT.frx":598F
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "HÖ thèng tµi kho¶n"
      Top             =   960
      Width           =   375
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "QuyÕt to¸n chi phÝ ®Çu t­ theo c«ng tr×nh, h¹ng môc hoµn thµnh - MÉu sè 04/QTDA"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   5640
      Width           =   6375
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "T×nh h×nh thùc hiÖn ®Çu t­ qua c¸c n¨m - MÉu sè 03/QTDA"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   5280
      Width           =   5775
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B¸o c¸o tæng hîp quyÕt to¸n vèn ®Çu t­ hoµn thµnh - MÉu sè 01vµ 02/QTDA"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   4920
      Width           =   6015
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   3000
      Picture         =   "FrmCDT.frx":5B0D
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   4200
      Picture         =   "FrmCDT.frx":6C7F
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   5400
      Picture         =   "FrmCDT.frx":80E1
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7440
      Width           =   1095
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sæ chi tiÕt doanh thu s¶n phÈm s¶n xuÊt thö"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   3495
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sæ chi tiÕt nguån vèn ®Çu t­"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   2775
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sæ chi phÝ ®Çu t­ x©y dùng"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   2775
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sæ chi phÝ kh¸c"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   2775
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sæ chi phÝ ban qu¶n lý dù ¸n"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   4
      Left            =   2280
      TabIndex        =   30
      Top             =   0
      Width           =   4335
      Begin VB.OptionButton OptTG 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tõ ngµy"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptTG 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tõ th¸ng"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.ComboBox CboThang 
         Height          =   315
         Index           =   1
         ItemData        =   "FrmCDT.frx":9503
         Left            =   3120
         List            =   "FrmCDT.frx":9505
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.ComboBox CboThang 
         Height          =   315
         Index           =   0
         ItemData        =   "FrmCDT.frx":9507
         Left            =   1080
         List            =   "FrmCDT.frx":9509
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin MSMask.MaskEdBox MedNgay 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
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
      Begin MSMask.MaskEdBox MedNgay 
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
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
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "®Õn ngµy"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "®Õn th¸ng"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   31
         Top             =   120
         Width           =   735
      End
   End
   Begin ComctlLib.ProgressBar GauGe 
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   7560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4320
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label LbTenVT 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   37
      Tag             =   "1"
      Top             =   2820
      Width           =   2895
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "CÊp TK ®èi øng"
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   36
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   35
      Tag             =   "1"
      Top             =   1005
      Width           =   3495
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tµi kho¶n"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   960
      Width           =   855
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   4320
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "FrmCDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim baocao As Integer
Dim ngay(0 To 1) As Date
Dim taikhoan As New ClsTaikhoan
Dim xlapp As Excel.Application, xlsheet As Worksheet

Private Sub txtShVT_GotFocus(Index As Integer)
    AutoSelect txtShVT(Index)
End Sub

Private Sub txtShVT_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdvt_Click Index
End Sub

Private Sub txtShVT_LostFocus(Index As Integer)
    Dim mvt As Long, dv As String
    
    LbTenVT(Index).Caption = TenVT(txtShVT(Index).Text, mvt, dv)
    txtShVT(Index).tag = mvt
    HienThongBao Me.Caption, 1
End Sub

Private Sub cmdvt_Click(Index As Integer)
    Me.MousePointer = 11
    txtShVT(Index).Text = FrmVattu.ChonVattu(txtShVT(Index).Text)
    Me.MousePointer = 0
    RFocus txtShVT(Index)
End Sub

Private Sub Command_Click(Index As Integer)
    Dim i As Integer, tdau As Integer, tcuoi As Integer
    Dim TK As New ClsTaikhoan, ndau As Date, ncuoi As Date
    
    Select Case Index
        Case 0, 1:
            Me.MousePointer = 11
            GauGe.Value = 0
            GauGe.Max = 3
            If CboThang(1).ListIndex < CboThang(0).ListIndex Then CboThang(1).ListIndex = CboThang(0).ListIndex
            tdau = CboThang(0).ItemData(CboThang(0).ListIndex)
            tcuoi = CboThang(1).ItemData(CboThang(1).ListIndex)
            If OptTG(0).Value Then
                ndau = NgayDauThang(pNamTC, tdau)
                ncuoi = NgayCuoiThang(pNamTC, tcuoi)
            Else
                ndau = ngay(0)
                ncuoi = ngay(1)
            End If
            SetRptInfo
            Select Case baocao
                Case 0:
                    BKChiTiet2 xlapp, xlsheet, "642", ndau, ncuoi, CboNK.ListIndex + 1, "Sæ chi phÝ Ban qu¶n lý dù ¸n", 1
                    GoTo KhongInBC
                Case 1:
                    If Not (KiemTraTK("2412")) Then GoTo KhongInBC
                    InChiPhiDT tdau, tcuoi, taikhoan
                    'BKChiTiet2 xlapp, xlsheet, "811", ndau, ncuoi, CboNK.ListIndex + 1, "Sæ chi phÝ kh¸c", 2
                    'GoTo KhongInBC
                Case 2:
                    If Not (KiemTraTK("2412")) Then GoTo KhongInBC
                    SoCPDauTuXD taikhoan, ndau, ncuoi
                    GoTo KhongInBC
                Case 3:
                    If Not (KiemTraTK("441") Or KiemTraTK("3")) Then GoTo KhongInBC
                    SoCTNguonVonDT taikhoan, ndau, ncuoi
                Case 4:
                    If Not KiemTraTK("511") Then GoTo KhongInBC
                    InCTDoanhThu2 ndau, ncuoi, txtShVT(0).tag, txtShTk(0).Text
                Case 5:
                    If Not (KiemTraTK("2412")) Then GoTo KhongInBC
                    THQuyetToan taikhoan
                    GoTo KhongInBC
                Case 6:
                    If Not (KiemTraTK("2412")) Then GoTo KhongInBC
                    ThucHienDauTu taikhoan
                    GoTo KhongInBC
                Case 7:
                    If Not (KiemTraTK("2412")) Then GoTo KhongInBC
                    QuyetToanChiPhi taikhoan
                    GoTo KhongInBC
                Case 8:
                    If Not (KiemTraTK("2412")) Then GoTo KhongInBC
                    DuyetQuyetToan taikhoan
                    GoTo KhongInBC
                Case 9:
                    NguonVonDauTu tdau, tcuoi
                Case 10:
                    If Not (KiemTraTK("2412", 3) Or KiemTraTK("2412", 5) Or KiemTraTK("2412", 7) Or KiemTraTK("2412", 9)) Then GoTo KhongInBC
                    ThucHienDauTu2 tdau, tcuoi, taikhoan
                Case 11:
                    TMCDT tdau, tcuoi
                    GoTo KhongInBC
                Case 12:
                    If Not (KiemTraTK("2412")) Then GoTo KhongInBC
                    ThucHienDauTu3 tdau, tcuoi, taikhoan
                    GoTo KhongInBC
                Case 13:
                    If Not (KiemTraTK("2412")) Then GoTo KhongInBC
                    TSBanGiao tdau, tcuoi, taikhoan
                    GoTo KhongInBC
                Case 14:
                    If Not (KiemTraTK("2412")) Then GoTo KhongInBC
                    CongNo_Von tdau, tcuoi, taikhoan
                    GoTo KhongInBC
                Case 15:
                    If Not (KiemTraTK("2412")) Then GoTo KhongInBC
                    QTDauTu tdau, tcuoi, taikhoan
                    GoTo KhongInBC
            End Select
            If Not RptOK(frmMain.Rpt.ReportFileName, 0) Then
                MsgBox "MÉu b¸o c¸o ®· bÞ thay ®æi!", vbCritical, App.ProductName
                GoTo KhongInBC
            End If
            GauGe.Value = GauGe.Max
            frmMain.Rpt.WindowTitle = OptBC(baocao).Caption
            InBaoCaoRPT
            GoTo KhongInBC
        Case 2:
            Unload Me
    End Select
KhongInBC:
    Set TK = Nothing
    GauGe.Value = GauGe.Max
    HienThongBao Me.Caption, 1
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim chi_so As Integer
    
    For chi_so = 0 To 1
        AddMonthToCbo CboThang(chi_so)
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next
    
    CboNK.ListIndex = 1
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set taikhoan = Nothing
    HienThongBao "", 1
End Sub

Private Sub OptBc_Click(Index As Integer)
    baocao = Index
    Select Case baocao
        Case 0, 2, 3, 4:
            OptTG(0).Enabled = True
            OptTG(1).Enabled = True
        Case 5, 6, 7, 8:
            OptTG(0).Enabled = False
            OptTG(1).Enabled = False
        Case 1, 9, 10, 11, 12, 13:
            OptTG(0).Enabled = True
            OptTG(0).Value = True
            OptTG(1).Enabled = False
    End Select
End Sub

Private Sub txtShTk_GotFocus(Index As Integer)
    AutoSelect txtShTk(Index)
End Sub

Private Sub txtShTk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdtk_Click Index
End Sub

'====================================================================================================
' KiÓm tra sè hiÖu tµi kho¶n ®· nhËp
'====================================================================================================
Private Sub txtShTk_LostFocus(Index As Integer)
    taikhoan.InitTaikhoanSohieu txtShTk(Index).Text
    LbTenTk(Index).Caption = taikhoan.Ten
End Sub

Private Sub cmdtk_Click(Index As Integer)
    Me.MousePointer = 11
    txtShTk(Index).Text = FrmTaikhoan.ChonTk(txtShTk(Index).Text)
    RFocus txtShTk(Index)
    Me.MousePointer = 0
End Sub

Private Sub MedNgay_GotFocus(Index As Integer)
    AutoSelect MedNgay(Index)
End Sub

Private Sub MedNgay_LostFocus(Index As Integer)
    If IsDate(MedNgay(Index).Text) Then
        ngay(Index) = CDate(MedNgay(Index).Text)
    Else
        MsgBox "NhËp sai ngµy !", vbExclamation, App.ProductName
        RFocus MedNgay(Index)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyX:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyI:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyV:
                RFocus Command(2)
                Command_Click 2
        End Select
    End If
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyP Then frmMain.mnHT_Click 8
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub SoCPDauTuXD(ctk As ClsTaikhoan, ndau As Date, ncuoi As Date)
    Dim rs As Recordset, sql As String, i As Integer
    Dim psn As Double, psc As Double, psnt1 As Double, psnt2 As Double, j As Integer
                
    Recycle pCurDir + "CTPS3.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\CTPS3.XLS", pCurDir + "CTPS3.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "CTPS3.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(4, 1) = "Tõ ngµy " + Format(ndau, Mask_DR) + " ®Õn ngµy " + Format(ncuoi, Mask_DR)
    xlsheet.Cells(5, 1) = "Dù ¸n: " + ctk.TenDA
    xlsheet.Cells(6, 1) = "C«ng tr×nh: " + ctk.Ten
    xlsheet.Cells(7, 1) = "Ngµy khëi c«ng: " + Format(ctk.NgayKC, Mask_DR) + " - Ngµy hoµn thµnh: " + Format(ctk.NgayHT, Mask_DR)
    xlsheet.Cells(8, 1) = "Dù to¸n: " + Format(ctk.DuToan, Mask_0)
    
    ctk.SoDuNgay ndau - 1, psn, psc, psnt1
    xlsheet.Cells(13, 6) = psn
    GauGe.Value = 1
    sql = "SELECT DISTINCTROW NgayGS,NgayCT,HethongTK.SoHieu AS SHN,TK.SoHieu AS SHC, ChungTu.SoHieu,DienGiai,SoPS FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) INNER JOIN HethongTK AS TK ON ChungTu.MaTKTCCo=TK.MaSo" _
        & " WHERE (HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*') AND " + WNgay("NgayGS", ndau, ncuoi) + " AND SoPS<>0 ORDER BY NgayGS"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 0
    Do While Not rs.EOF
        j = j + 1
        xlsheet.Cells(13 + j, 1) = Format(rs!NgayGS, Mask_D)
        xlsheet.Cells(13 + j, 2) = rs!sohieu
        xlsheet.Cells(13 + j, 3) = Format(rs!NgayCT, Mask_D)
        xlsheet.Cells(13 + j, 4) = rs!diengiai
        xlsheet.Cells(13 + j, 5) = rs!shc
        xlsheet.Cells(13 + j, 6) = rs!sops
        sql = ctk.ShCapDuoi(rs!SHN)
        If sql <> "" Then sql = Right(sql, 1)
        If IsNumeric(sql) Then
            i = CInt(sql)
            Select Case i
                Case 1, 2:
                    xlsheet.Cells(13 + j, 6 + i) = rs!sops
                Case 3:
                    xlsheet.Cells(13 + j, 12) = rs!sops
            End Select
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    GauGe.Value = 2
    
    If j > 0 Then
        For i = 6 To 12
            xlsheet.Range(XLSCol(i) + CStr(14 + j)).Formula = "=SUM(" + XLSCol(i) + "14:" + XLSCol(i) + CStr(13 + j) + ")"
        Next
    End If
    
    xlsheet.Cells(14 + j, 4) = "Céng ph¸t sinh trong kú"

    xlsheet.Cells(15 + j, 4) = "Luü kÕ ph¸t sinh tõ ®Çu n¨m"
    xlsheet.Cells(16 + j, 4) = "Luü kÕ ph¸t sinh tõ khëi c«ng"
    ctk.SoPhatSinhN NgayDauThang(pNamTC, pThangDauKy), ncuoi, psn, psnt1, psc, psnt2
    xlsheet.Cells(15 + j, 6) = psn
    xlsheet.Cells(16 + j, 6) = ctk.PSTuKhoiCong(ncuoi, -1)
        
    With xlsheet.Range("A10", "M" + CStr(16 + j))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 8
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
        
    xlapp.Workbooks.Close
    
    CallExcel "CTPS3.XLS"
    
KetThuc:
    Set ctk = Nothing
End Sub

Private Sub SoCTNguonVonDT(taikhoan As ClsTaikhoan, ndau As Date, ncuoi As Date)
    InSocaiTk taikhoan, 0, 0, ndau, ncuoi, False, "", 0, 0
    If taikhoan.MaTC <> taikhoan.MaSo And taikhoan.MaTC > 0 Then DBKetoan.QueryDefs("QSoCai").sql = DBKetoan.QueryDefs("QChiTiet").sql
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\VONDT.RPT"
    GauGe.Value = 1
    frmMain.Rpt.Formulas(60) = "LKN=" + CStr(taikhoan.PSTuKhoiCong(ncuoi, -1))
    frmMain.Rpt.Formulas(61) = "LKC=" + CStr(taikhoan.PSTuKhoiCong(ncuoi, 1))
End Sub

Private Function KiemTraTK(Optional shtk As String = "241", Optional cap As Integer = 0) As Boolean
    If Left(taikhoan.sohieu, Len(shtk)) <> shtk Then
        RFocus txtShTk(0)
        KiemTraTK = False
    Else
        KiemTraTK = IIf(cap > 0, taikhoan.cap = cap, True)
    End If
End Function

Private Sub THQuyetToan(taikhoan As ClsTaikhoan)
    Dim TK As New ClsTaikhoan, i As Integer
    
    Recycle pCurDir + "THQTDT.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\THQTDT.XLS", pCurDir + "THQTDT.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "THQTDT.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(4, 4) = taikhoan.TenDA
    xlsheet.Cells(5, 4) = taikhoan.Ten
    xlsheet.Cells(6, 4) = pTenCty
    xlsheet.Cells(9, 4) = taikhoan.DiaDiem
    xlsheet.Cells(11, 4) = taikhoan.DuToan
    xlsheet.Cells(12, 4) = taikhoan.NgayKC
    xlsheet.Cells(13, 4) = taikhoan.NgayHT
    
    xlsheet.Cells(24, 3) = taikhoan.Von1
    xlsheet.Cells(25, 3) = taikhoan.Von2
    xlsheet.Cells(28, 3) = taikhoan.Von3
    GauGe.Value = 1
    
    For i = 1 To 4
        TK.InitTaikhoanSohieuCT taikhoan.sohieu, CStr(i), taikhoan.cap + 1
        If TK.MaSo > 0 Then
            xlsheet.Cells(35 + i, 4) = TK.DuToan
            xlsheet.Cells(35 + i, 5) = TK.PSTuKhoiCong(NgayCuoiNam, -1)
        End If
    Next
    
    For i = 1 To 4
        TK.InitTaikhoanSohieuCT taikhoan.sohieu, CStr(i), taikhoan.cap + 1
        If TK.MaSo > 0 Then
            xlsheet.Cells(46 + i, 3) = TK.PSTuKhoiCong(NgayCuoiNam, -1)
        End If
    Next
    GauGe.Value = 2
    xlsheet.Cells(65, 3) = taikhoan.PSTuKhoiCong(NgayCuoiNam, -1)
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "THQTDT.XLS"
KetThuc:
    Set TK = Nothing
End Sub

Private Sub ThucHienDauTu(taikhoan As ClsTaikhoan)
    Dim TK As New ClsTaikhoan, i As Integer, psn As Double, psc As Double, namdau As Integer, j As Integer, k As Integer
    
    Recycle pCurDir + "THDT.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\THDT.XLS", pCurDir + "THDT.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "THDT.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(4, 4) = taikhoan.TenDA
    xlsheet.Cells(5, 4) = taikhoan.Ten
    xlsheet.Cells(6, 4) = pTenCty
    xlsheet.Cells(9, 4) = taikhoan.DiaDiem
    xlsheet.Cells(11, 4) = taikhoan.DuToan
    xlsheet.Cells(12, 4) = taikhoan.NgayKC
    xlsheet.Cells(13, 4) = taikhoan.NgayHT
    
    i = pNamTC - 1
    Do While TruongDaCo("HethongTK", "PSNLK" + CStr(i))
        namdau = i
        i = i - 1
    Loop
    GauGe.Value = 1
    If namdau = 0 Then namdau = pNamTC
    If namdau > 0 Then
        For i = namdau To pNamTC - 1
            j = j + 1
            xlsheet.Cells(17 + j, 1) = j
            xlsheet.Cells(17 + j, 2) = i
            xlsheet.Cells(17 + j, 4) = taikhoan.PSLKNam(i, -1)
            For k = 1 To 3
                TK.InitTaikhoanSohieuCT taikhoan.sohieu, CStr(k), taikhoan.cap + 1
                If TK.MaSo > 0 Then xlsheet.Cells(17 + j, 4 + k) = taikhoan.PSLKNam(i, -1)
            Next
        Next
        If namdau = pNamTC Then
            j = j + 1
            xlsheet.Cells(17 + j, 1) = j
            xlsheet.Cells(17 + j, 2) = "Tr­íc " + CStr(pNamTC)
            xlsheet.Cells(17 + j, 4) = taikhoan.PSNLK
            For k = 1 To 3
                TK.InitTaikhoanSohieuCT taikhoan.sohieu, CStr(k), taikhoan.cap + 1
                If TK.MaSo > 0 Then
                    xlsheet.Cells(17 + j, 4 + k) = TK.PSNLK
                End If
            Next
        End If
        j = j + 1
        xlsheet.Cells(17 + j, 1) = j
        xlsheet.Cells(17 + j, 2) = pNamTC
        taikhoan.SoPhatSinh 0, 0, psn, psc
        xlsheet.Cells(17 + j, 4) = psn
        For k = 1 To 3
            TK.InitTaikhoanSohieuCT taikhoan.sohieu, CStr(k), taikhoan.cap + 1
            If TK.MaSo > 0 Then
                TK.SoPhatSinh 0, 0, psn, psc
                xlsheet.Cells(17 + j, 4 + k) = psn
            End If
        Next
    End If
    GauGe.Value = 2
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "THDT.XLS"
KetThuc:
    Set TK = Nothing
End Sub

Private Sub QuyetToanChiPhi(taikhoan As ClsTaikhoan)
    Dim TK As New ClsTaikhoan, i As Integer, k As Integer, rs As Recordset, shtk As String, cap As Integer
    
    Recycle pCurDir + "QTCP.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\QTCP.XLS", pCurDir + "QTCP.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "QTCP.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(4, 4) = taikhoan.TenDA
    xlsheet.Cells(5, 4) = taikhoan.Ten
    xlsheet.Cells(6, 4) = pTenCty
    xlsheet.Cells(9, 4) = taikhoan.DiaDiem
    xlsheet.Cells(11, 4) = taikhoan.DuToan
    xlsheet.Cells(12, 4) = taikhoan.NgayKC
    xlsheet.Cells(13, 4) = taikhoan.NgayHT
    GauGe.Value = 1
    'Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo FROM HethongTK WHERE Sohieu LIKE '" + taikhoan.SoHieu + "*' AND (Cap=3 OR Cap=5 OR Cap=7 OR Cap=9) ORDER BY SoHieu", dbOpenSnapshot)
    'Do While Not rs.EOF
    '    i = i + 1
    '    TK.InitTaikhoanMaSo rs!MaSo
    '    xlsheet.Cells(17 + i, 1) = TK.Ten
    '    xlsheet.Cells(17 + i, 2) = taikhoan.DuToan
    '    xlsheet.Cells(17 + i, 3) = TK.PSTuKhoiCong(NgayCuoiNam, -1)
    '    shtk = TK.SoHieu
    '    cap = TK.cap + 1
    '    For k = 1 To 3
    '        TK.InitTaikhoanSohieuCT shtk, CStr(k), cap
    '        If TK.MaSo > 0 Then
    '            xlsheet.Cells(17 + i, 3 + k) = TK.PSTuKhoiCong(NgayCuoiNam, -1)
    '        End If
    '    Next
    '    rs.MoveNext
    'Loop
    'rs.Close
        i = i + 1
        xlsheet.Cells(17 + i, 1) = taikhoan.Ten
        xlsheet.Cells(17 + i, 2) = taikhoan.DuToan
        xlsheet.Cells(17 + i, 3) = taikhoan.PSTuKhoiCong(NgayCuoiNam, -1)
        shtk = taikhoan.sohieu
        cap = taikhoan.cap + 1
        For k = 1 To 3
            TK.InitTaikhoanSohieuCT shtk, CStr(k), cap
            If TK.MaSo > 0 Then
                xlsheet.Cells(17 + i, 3 + k) = TK.PSTuKhoiCong(NgayCuoiNam, -1)
            End If
        Next
    
    Set rs = Nothing
    GauGe.Value = 2
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "QTCP.XLS"
KetThuc:
    Set TK = Nothing
End Sub

Private Sub DuyetQuyetToan(taikhoan As ClsTaikhoan)
    Dim TK As New ClsTaikhoan, i As Integer, k As Integer
    
    Recycle pCurDir + "DuyetQT.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\DuyetQT.XLS", pCurDir + "DuyetQT.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "DuyetQT.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(7, 4) = taikhoan.TenDA
    xlsheet.Cells(8, 4) = taikhoan.Ten
    xlsheet.Cells(9, 4) = pTenCty
    xlsheet.Cells(12, 4) = taikhoan.DiaDiem
    xlsheet.Cells(14, 4) = taikhoan.DuToan
    xlsheet.Cells(15, 4) = taikhoan.NgayKC
    xlsheet.Cells(16, 4) = taikhoan.NgayHT
    
    xlsheet.Cells(23, 2) = taikhoan.Von1
    xlsheet.Cells(25, 2) = taikhoan.Von2
    xlsheet.Cells(27, 2) = taikhoan.Von3
    
    xlsheet.Cells(52, 2) = taikhoan.Von1
    xlsheet.Cells(54, 2) = taikhoan.Von2
    xlsheet.Cells(56, 2) = taikhoan.Von3
    
    GauGe.Value = 1
    For k = 1 To 3
        TK.InitTaikhoanSohieuCT taikhoan.sohieu, CStr(k), taikhoan.cap + 1
        If TK.MaSo > 0 Then
            xlsheet.Cells(32 + k, 2) = TK.DuToan
            xlsheet.Cells(32 + k, 3) = TK.PSTuKhoiCong(NgayCuoiNam, -1)
        End If
    Next
    
    xlsheet.Cells(44, 2) = taikhoan.PSTuKhoiCong(NgayCuoiNam, -1)
    
    xlsheet.Cells(59, 2) = SoDuTKSH(pSHPT, ThangCuoiNamTC, -1)
    xlsheet.Cells(60, 2) = SoDuTKSH("331", ThangCuoiNamTC, 1)
    
    GauGe.Value = 2
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "DuyetQT.XLS"
KetThuc:
    Set TK = Nothing
End Sub

Public Sub NguonVonDauTu(tdau As Integer, tcuoi As Integer)
    Dim st As String, sqln As String, sqlc As String, i As Integer
    Dim sqlnx As String, sqlcx As String
    
    GauGe.Max = 3
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + No_" + st
        sqlc = sqlc + " + Co_" + st
    Next
    For i = 1 To CThangDB(tcuoi)
        st = CStr(i)
        sqlnx = sqlnx + " + No_" + st
        sqlcx = sqlcx + " + Co_" + st
    Next
    GauGe.Value = 1
    st = CStr(CThangDB(tcuoi))

    SetSQL "QTongHopCT", "SELECT DISTINCTROW SoHieu, First(Cap) AS FirstOfCap, First(Ten) AS FirstOfTen, First(Kieu) AS FirstOfKieu, First(Loai) AS FirstOfLoai, First(MaTC = MaSo OR MaTC = 0) AS FirstOfChuanLa, First(TKCon) AS FirstOfTKCon, First(CapDuoi) AS CD," _
        & " Sum(DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkNo, Sum(DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkCo, Sum(" + sqln + ") AS PsNo, Sum(" + sqlc + ") AS PsCo, Sum(" + sqlnx + ") AS PsNoX, Sum(" + sqlcx + ") AS PsCoX, Sum(HeThongTK.DuNo_" + st + ") AS CkNo, Sum(HeThongTK.DuCo_" + st + ") AS CkCo, SUM(PSNLK+" + sqlnx + ") AS LKN, SUM(PSCLK+" + sqlcx + ") AS LKC" _
        & " FROM HeThongTK WHERE Cap>1 AND (SoHieu LIKE '341*' OR SoHieu LIKE '441*') GROUP BY SoHieu HAVING (Sum(DuNo_" + st + ")<>0 OR Sum(DuCo_" + st + ")<>0 OR Sum(" + sqln + ")<>0 OR Sum(" + sqlc + ")<>0 OR Sum(" + sqlnx + ")<>0 OR Sum(" + sqlcx + ")<>0) ORDER BY SoHieu"

    GauGe.Value = 2
    XDCapTK "341"
    XDCapTK "441"
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\NVONDT.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi)
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi) + "'"
End Sub

Public Sub ThucHienDauTu2(tdau As Integer, tcuoi As Integer, taikhoan As ClsTaikhoan)
    Dim st As String, sqln As String, sqlc As String, i As Integer
    Dim sqlnx As String, sqlcx As String
    
    GauGe.Max = 3
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + No_" + st
        sqlc = sqlc + " + Co_" + st
    Next
    For i = 1 To CThangDB(tcuoi)
        st = CStr(i)
        sqlnx = sqlnx + " + No_" + st
        sqlcx = sqlcx + " + Co_" + st
    Next
    GauGe.Value = 1
    st = CStr(CThangDB(tcuoi))

    SetSQL "QTongHopCT", "SELECT DISTINCTROW SoHieu, First(Cap) AS FirstOfCap, First(Ten) AS FirstOfTen, First(Kieu) AS FirstOfKieu, First(Loai) AS FirstOfLoai, First(MaTC = MaSo OR MaTC = 0) AS FirstOfChuanLa, First(TKCon) AS FirstOfTKCon, First(CapDuoi) AS CD," _
        & " Sum(DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkNo, Sum(DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkCo, Sum(" + sqln + ") AS PsNo, Sum(" + sqlc + ") AS PsCo, Sum(" + sqlnx + ") AS PsNoX, Sum(" + sqlcx + ") AS PsCoX, Sum(HeThongTK.DuNo_" + st + ") AS CkNo, Sum(HeThongTK.DuCo_" + st + ") AS CkCo, SUM(PSNLK+" + sqlnx + ") AS LKN, SUM(PSCLK+" + sqlcx + ") AS LKC" _
        & " FROM HeThongTK WHERE SoHieu LIKE '" + taikhoan.sohieu + "*' AND Cap>" + CStr(taikhoan.cap) + " GROUP BY SoHieu HAVING (Sum(DuNo_" + st + ")<>0 OR Sum(DuCo_" + st + ")<>0 OR Sum(" + sqln + ")<>0 OR Sum(" + sqlc + ")<>0 OR Sum(" + sqlnx + ")<>0 OR Sum(" + sqlcx + ")<>0) ORDER BY SoHieu"

    GauGe.Value = 2
    XDCapTK "2412"
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\THDAUTU.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi)
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi) + "'"
    frmMain.Rpt.Formulas(4) = "DienGiai='" + taikhoan.Ten + "'"
    frmMain.Rpt.Formulas(5) = "Cap=" + CStr(taikhoan.cap + 1)
End Sub

Private Sub TMCDT(tdau As Integer, tcuoi As Integer)
    Dim i As Integer, rs As Recordset, sql As String, sqlx As String
    
    Recycle pCurDir + "TMCDT.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\TMCDT.XLS", pCurDir + "TMCDT.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "TMCDT.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi)
    sql = ""
    sqlx = ""
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sql = sql + "+Co_" + CStr(i)
    Next
    For i = 1 To CThangDB(tcuoi)
        sqlx = sqlx + "+Co_" + CStr(i)
    Next
    i = 0
    Set rs = DBKetoan.OpenRecordset("SELECT Ten,Sum(" + sql + ") AS KN, Sum(" + sqlx + ") AS LK FROM HethongTK WHERE SoHieu LIKE '511*' AND TKCon=0 GROUP BY SoHieu, Ten HAVING Sum(" + sqlx + ")<>0", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(23 + i, 1) = rs!Ten
        xlsheet.Cells(23 + i, 2) = rs!kn
        xlsheet.Cells(23 + i, 3) = rs!lk
        rs.MoveNext
    Loop
    sql = ""
    sqlx = ""
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sql = sql + "+No_" + CStr(i)
    Next
    For i = 1 To CThangDB(tcuoi)
        sqlx = sqlx + "+No_" + CStr(i)
    Next
    i = 0
    Set rs = DBKetoan.OpenRecordset("SELECT Ten,Sum(" + sql + ") AS KN, Sum(" + sqlx + ") AS LK FROM HethongTK WHERE SoHieu LIKE '6*' AND TKCon=0 GROUP BY SoHieu, Ten HAVING Sum(" + sqlx + ")<>0", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(34 + i, 1) = rs!Ten
        xlsheet.Cells(34 + i, 2) = rs!kn
        xlsheet.Cells(34 + i, 3) = rs!lk
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    xlsheet.Cells(74, 2) = SoDuTKSH("2112", ThangTruoc(tdau))
    xlsheet.Cells(74, 3) = SoDuTKSH("2113", ThangTruoc(tdau))
    xlsheet.Cells(74, 4) = SoDuTKSH("2114", ThangTruoc(tdau))
    xlsheet.Cells(74, 5) = SoDuTKSH("2115", ThangTruoc(tdau))
    xlsheet.Cells(74, 6) = SoDuTKSH("2116", ThangTruoc(tdau))
    xlsheet.Cells(74, 7) = SoDuTKSH("2118", ThangTruoc(tdau))
    xlsheet.Cells(75, 2) = SoPSTK("2112", tdau, tcuoi, -1)
    xlsheet.Cells(75, 3) = SoPSTK("2113", tdau, tcuoi, -1)
    xlsheet.Cells(75, 4) = SoPSTK("2114", tdau, tcuoi, -1)
    xlsheet.Cells(75, 5) = SoPSTK("2115", tdau, tcuoi, -1)
    xlsheet.Cells(75, 6) = SoPSTK("2116", tdau, tcuoi, -1)
    xlsheet.Cells(75, 7) = SoPSTK("2118", tdau, tcuoi, -1)
    xlsheet.Cells(76, 2) = SoPSTK("2112", tdau, tcuoi, 1)
    xlsheet.Cells(76, 3) = SoPSTK("2113", tdau, tcuoi, 1)
    xlsheet.Cells(76, 4) = SoPSTK("2114", tdau, tcuoi, 1)
    xlsheet.Cells(76, 5) = SoPSTK("2115", tdau, tcuoi, 1)
    xlsheet.Cells(76, 6) = SoPSTK("2116", tdau, tcuoi, 1)
    xlsheet.Cells(76, 7) = SoPSTK("2118", tdau, tcuoi, 1)
    xlsheet.Cells(79, 2) = GTHaoMon("2112", ThangTruoc(tdau))
    xlsheet.Cells(79, 3) = GTHaoMon("2113", ThangTruoc(tdau))
    xlsheet.Cells(79, 4) = GTHaoMon("2114", ThangTruoc(tdau))
    xlsheet.Cells(79, 5) = GTHaoMon("2115", ThangTruoc(tdau))
    xlsheet.Cells(79, 6) = GTHaoMon("2116", ThangTruoc(tdau))
    xlsheet.Cells(79, 7) = GTHaoMon("2118", ThangTruoc(tdau))
    xlsheet.Cells(80, 2) = SoKHTS("2112", tdau, tcuoi)
    xlsheet.Cells(80, 3) = SoKHTS("2113", tdau, tcuoi)
    xlsheet.Cells(80, 4) = SoKHTS("2114", tdau, tcuoi)
    xlsheet.Cells(80, 5) = SoKHTS("2115", tdau, tcuoi)
    xlsheet.Cells(80, 6) = SoKHTS("2116", tdau, tcuoi)
    xlsheet.Cells(80, 7) = SoKHTS("2118", tdau, tcuoi)
    xlsheet.Cells(82, 2) = GTHaoMon("2112", tcuoi)
    xlsheet.Cells(82, 3) = GTHaoMon("2113", tcuoi)
    xlsheet.Cells(82, 4) = GTHaoMon("2114", tcuoi)
    xlsheet.Cells(82, 5) = GTHaoMon("2115", tcuoi)
    xlsheet.Cells(82, 6) = GTHaoMon("2116", tcuoi)
    xlsheet.Cells(82, 7) = GTHaoMon("2118", tcuoi)
    GauGe.Value = 2
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "TMCDT.XLS"
KetThuc:
End Sub

Private Sub ThucHienDauTu3(tdau As Integer, tcuoi As Integer, taikhoan As ClsTaikhoan)
    Dim i As Integer, k As Integer, shct As String
    Dim rs As Recordset, sql As String, sqlx As String
    Dim s(2 To 15) As Double
    
    Recycle pCurDir + "THDTCT.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\THDTCT.XLS", pCurDir + "THDTCT.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "THDTCT.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi)
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sql = sql + "+No_" + CStr(i)
    Next
    For i = 1 To CThangDB(tcuoi)
        sqlx = sqlx + "+No_" + CStr(i)
    Next
    i = 0
    Set rs = DBKetoan.OpenRecordset("SELECT Sohieu,Cap,Ten,DuToan,(" + sql + ") AS KN, (" + sqlx + ") AS LK,(PSNLK+" + sqlx + ") AS LK2 FROM HethongTK WHERE SoHieu LIKE '" + taikhoan.sohieu + "*' AND Cap>" + CStr(taikhoan.cap) + " AND (PSNLK+" + sqlx + "<>0 OR DuToan<>0) ORDER BY SoHieu", dbOpenSnapshot)             'AND (Cap=3 OR Cap=5 OR Cap=7 OR Cap=9)
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(9 + i, 1) = Space(2 * (rs!cap - taikhoan.cap)) + rs!Ten
        xlsheet.Cells(9 + i, 2) = rs!DuToan
        shct = taikhoan.ShCapDuoi(rs!sohieu)
        If shct <> "" Then shct = Right(shct, 1)
        If IsNumeric(shct) Then
            k = CInt5(shct)
            xlsheet.Cells(9 + i, 3 + k) = rs!kn
            xlsheet.Cells(9 + i, 7 + k) = rs!lk
            xlsheet.Cells(9 + i, 11 + k) = rs!lk2
            If rs!cap = taikhoan.cap + 1 Then
                s(2) = s(2) + rs!DuToan
                s(3 + k) = s(3 + k) + rs!kn
                s(7 + k) = s(7 + k) + rs!lk
                s(11 + k) = s(11 + k) + rs!lk2
                s(7) = s(4) + s(5) + s(6)
                s(11) = s(8) + s(9) + s(10)
                s(15) = s(12) + s(13) + s(14)
            End If
        End If
        xlsheet.Range("G" + CStr(9 + i)).Formula = "=SUM(D" + CStr(9 + i) + ":F" + CStr(9 + i) + ")"
        xlsheet.Range("K" + CStr(9 + i)).Formula = "=SUM(H" + CStr(9 + i) + ":J" + CStr(9 + i) + ")"
        xlsheet.Range("O" + CStr(9 + i)).Formula = "=SUM(L" + CStr(9 + i) + ":N" + CStr(9 + i) + ")"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    xlsheet.Cells(10 + i, 1) = ABCtoVNI("Tæng céng")
    If i > 0 Then
        For k = 2 To 15
            xlsheet.Cells(10 + i, k) = s(k)
        Next
    End If
    GauGe.Value = 2
    With xlsheet.Range("A8", "O" + CStr(10 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 8
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "THDTCT.XLS"
KetThuc:
End Sub

Public Sub InChiPhiDT(tdau As Integer, tcuoi As Integer, taikhoan As ClsTaikhoan)
    Dim st As String, sqln As String, sqlc As String, i As Integer
    Dim sqlnx As String, sqlcx As String, TK As New ClsTaikhoan
    
    TK.InitTaikhoanSohieuCT taikhoan.sohieu, "3", taikhoan.cap + 1
    
    GauGe.Max = 3
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + No_" + st
        sqlc = sqlc + " + Co_" + st
    Next
    For i = 1 To CThangDB(tcuoi)
        st = CStr(i)
        sqlnx = sqlnx + " + No_" + st
        sqlcx = sqlcx + " + Co_" + st
    Next
    GauGe.Value = 1
    st = CStr(CThangDB(tcuoi))

    SetSQL "QTongHopCT", "SELECT DISTINCTROW SoHieu, First(Cap) AS FirstOfCap, First(Ten) AS FirstOfTen, First(Kieu) AS FirstOfKieu, First(Loai) AS FirstOfLoai, First(MaTC = MaSo OR MaTC = 0) AS FirstOfChuanLa, First(TKCon) AS FirstOfTKCon, First(CapDuoi) AS CD," _
        & " Sum(DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkNo, Sum(DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkCo, Sum(" + sqln + ") AS PsNo, Sum(" + sqlc + ") AS PsCo, Sum(" + sqlnx + ") AS PsNoX, Sum(" + sqlcx + ") AS PsCoX, Sum(HeThongTK.DuNo_" + st + ") AS CkNo, Sum(HeThongTK.DuCo_" + st + ") AS CkCo, SUM(PSNLK+" + sqlnx + ") AS LKN, SUM(PSCLK+" + sqlcx + ") AS LKC" _
        & " FROM HeThongTK WHERE SoHieu LIKE '" + TK.sohieu + "*' AND Cap>" + CStr(TK.cap) + " GROUP BY SoHieu HAVING (Sum(DuNo_" + st + ")<>0 OR Sum(DuCo_" + st + ")<>0 OR Sum(" + sqln + ")<>0 OR Sum(" + sqlc + ")<>0 OR Sum(" + sqlnx + ")<>0 OR Sum(" + sqlcx + ")<>0) ORDER BY SoHieu"

    GauGe.Value = 2
    XDCapTK TK.sohieu
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\CPK_CDT.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi)
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi) + "'"
    frmMain.Rpt.Formulas(4) = "DienGiai='" + taikhoan.Ten + "'"
    frmMain.Rpt.Formulas(5) = "Cap=" + CStr(TK.cap + 1)
    
    Set TK = Nothing
End Sub

Private Sub TSBanGiao(tdau As Integer, tcuoi As Integer, taikhoan As ClsTaikhoan)
    Recycle pCurDir + "THDTCT.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\TSCD_LD.XLS", pCurDir + "THDTCT.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "THDTCT.XLS"
    Set xlsheet = xlapp.Worksheets(1)
        
    On Error GoTo 0
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "TSCD_LD.XLS"
KetThuc:
End Sub

Private Sub CongNo_Von(tdau As Integer, tcuoi As Integer, taikhoan As ClsTaikhoan)
    Recycle pCurDir + "CN_DA.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\CN_DA.XLS", pCurDir + "CN_DA.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "CN_DA.XLS"
    Set xlsheet = xlapp.Worksheets(2)
    
    xlsheet.Cells(4, 4) = taikhoan.TenDA

    xlsheet.Cells(6, 4) = pTenCty
    
    On Error GoTo 0
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "CN_DA.XLS"
KetThuc:
End Sub

Private Sub QTDauTu(tdau As Integer, tcuoi As Integer, taikhoan As ClsTaikhoan)
    Recycle pCurDir + "QTDT.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\QTDT.XLS", pCurDir + "QTDT.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "QTDT.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    
    xlsheet.Cells(4, 4) = taikhoan.TenDA
        
    On Error GoTo 0
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "QTDT.XLS"
KetThuc:
End Sub

Private Sub InCTDoanhThu2(ndau As Date, ncuoi As Date, mvt As Long, Optional shtk As String = "")
    Dim sql As String, st As String, dv As String
        
    sql = "SELECT DISTINCTROW ChungTu.MaCT,ChungTu.SoHieu, NgayCT, NgayGS, ChungTu.DienGiai, ChungTu.SoPS, ChungTu.SoPS2Co, HethongTK.SoHieu AS TKDU " _
        & " FROM (ChungTu INNER JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo = HethongTK.MaSo" _
        + " WHERE " + WNgay("NgayGS", ndau, ncuoi) + " AND TK.TK_ID=" + CStr(TKDT_ID) + " AND MaVattu=" + CStr(mvt) + IIf(shtk <> "", " AND TK.SoHieu LIKE '" + shtk + "*'", "")
    sql = sql + " UNION SELECT DISTINCTROW ChungTu.MaCT,ChungTu.SoHieu, NgayCT, NgayGS, ChungTu.DienGiai, ChungTu.SoPS, ChungTu.SoPS2No, TK.SoHieu AS TKDU " _
        & " FROM (ChungTu INNER JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo = HethongTK.MaSo" _
        + " WHERE " + WNgay("NgayGS", ndau, ncuoi) + " AND HethongTK.TK_ID=" + CStr(TKDT_ID) + " AND TK.TK_ID=" + CStr(TKGT_ID) + " AND MaVattu=" + CStr(mvt) + IIf(shtk <> "", " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "")
    SetSQL "QChitiet", sql
    
    frmMain.Rpt.WindowTitle = "Sæ chi tiÕt doanh thu"
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\CTDT3.RPT"
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN 'Tõ ngµy " + Format(ndau, "dd/mm/yy") + " ®Õn " + Format(ncuoi, "dd/mm/yy") + "'"
    sql = "SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-2) + " WHERE MaLoai=2 AND HethongTK.SoHieu LIKE '632*' AND MaVattu=" + CStr(mvt) + " AND " + WNgay("NgayGS", ndau, ncuoi)
    frmMain.Rpt.Formulas(4) = "GV = " + CStr(SelectSQL(sql))
    sql = TenVT(st, mvt, dv)
    frmMain.Rpt.Formulas(5) = "TenVt = IF PageNumber() = 1 THEN '" + st + " - " + sql + ABCtoVNI(" - §¬n vÞ tÝnh: ") + dv + "'"
    RptSetDate ncuoi
End Sub

Private Function BKChiTiet2(xlapp As Excel.Application, xlsheet As Worksheet, TK As String, ndau As Date, ncuoi As Date, cap As Integer, Optional dg As String = "", Optional loaibc As Integer = 0) As Boolean
    Dim rs As Recordset, sql As String, i As Integer, ctk As New ClsTaikhoan
    Dim psn As Double, psc As Double, psnt1 As Double, psnt2 As Double, j As Integer, k As Integer
    
    XDSoHieuCap cap
    ' Lay danh sach tai khoan tai chinh
    sql = "SELECT DISTINCTROW BaoCaoCP.SoHieu AS SHN FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo = HethongTK.MaSo) LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE (HethongTK.SoHieu LIKE '" + TK + "*') AND " + WNgay("NgayGS", ndau, ncuoi) + " AND SoPS<>0 GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        
    Recycle pCurDir + "CTPS2.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\CTPS2.XLS", pCurDir + "CTPS2.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "CTPS2.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    If dg <> "" Then
        xlsheet.Cells(3, 1) = dg
    Else
        xlsheet.Cells(3, 1) = xlsheet.Cells(3, 1)
    End If
    xlsheet.Cells(4, 1) = "Tõ ngµy " + Format(ndau, Mask_DR) + " ®Õn ngµy " + Format(ncuoi, Mask_DR)
    If rs.RecordCount = 0 Then GoTo AA
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, 5 + i) = rs!SHN
        rs.MoveNext
    Loop
    xlsheet.Cells(6, 6 + i) = "Ghi chó"
    
    xlsheet.Range("F6", XLSCol(5 + i) + "6").MergeCells = True
    xlsheet.Range(XLSCol(6 + i) + "6", XLSCol(6 + i) + "7").MergeCells = True
    
    sql = "SELECT DISTINCTROW NgayGS,NgayCT,BaoCaoCP.SoHieu AS SHN,LEFT(TK.SoHieu,4) AS SHC, ChungTu.SoHieu,DienGiai,SoPS FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE (HethongTK.SoHieu LIKE '" + TK + "*') AND " + WNgay("NgayGS", ndau, ncuoi) + " AND SoPS<>0 ORDER BY NgayGS"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 0
    Do While Not rs.EOF
        j = j + 1
        xlsheet.Cells(7 + j, 1) = Format(rs!NgayGS, Mask_D)
        xlsheet.Cells(7 + j, 2) = rs!sohieu
        xlsheet.Cells(7 + j, 3) = Format(rs!NgayCT, Mask_D)
        xlsheet.Cells(7 + j, 4) = rs!diengiai
        xlsheet.Cells(7 + j, 5) = rs!sops
        For k = 1 To i
            If xlsheet.Cells(7, 5 + k) = rs!SHN Then
                xlsheet.Cells(7 + j, 5 + k) = rs!sops
                Exit For
            End If
        Next
        rs.MoveNext
    Loop
    
    For k = 0 To i
        xlsheet.Range(XLSCol(5 + k) + CStr(8 + j)).Formula = "=SUM(" + XLSCol(5 + k) + "8:" + XLSCol(5 + k) + CStr(7 + j) + ")"
    Next
    xlsheet.Cells(8 + j, 4) = "Céng ph¸t sinh trong kú"
AA:
    ctk.InitTaikhoanSohieu TK
    Select Case loaibc
        Case 1:
            xlsheet.Cells(9 + j, 4) = "Luü kÕ ph¸t sinh tõ ®Çu n¨m"
            xlsheet.Cells(10 + j, 4) = "Luü kÕ ph¸t sinh tõ khëi c«ng"
            ctk.SoPhatSinhN NgayDauThang(pNamTC, pThangDauKy), ncuoi, psn, psnt1, psc, psnt2
            xlsheet.Cells(9 + j, 5) = psn
            xlsheet.Cells(10 + j, 5) = ctk.PSTuKhoiCong(ncuoi, -1)
        Case 2:
            xlsheet.Cells(9 + j, 4) = "Chi phÝ ®· ph©n bæ"
            xlsheet.Cells(10 + j, 4) = "Chi phÝ ch­a ph©n bæ"
            xlsheet.Cells(9 + j, 5) = ctk.CPDaPhanBo(ndau, ncuoi, "241")
            xlsheet.Cells(10 + j, 5) = xlsheet.Cells(8 + j, 5) - xlsheet.Cells(9 + j, 5)
    End Select
        
    With xlsheet.Range("A6", XLSCol(6 + i) + CStr(8 + j))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 8
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
        
    xlapp.Workbooks.Close
    
    CallExcel "CTPS2.XLS"
    
KetThuc:
    rs.Close
    Set rs = Nothing
    Set ctk = Nothing
End Function
