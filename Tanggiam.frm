VERSION 5.00
Begin VB.Form frmTangGiam 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "T¨ng gi¶m gi¸ trÞ tµi s¶n"
   ClientHeight    =   6015
   ClientLeft      =   1590
   ClientTop       =   705
   ClientWidth     =   6255
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
   Icon            =   "Tanggiam.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6015
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Fixed Assets Increasing - Decreasing"
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   3240
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   3240
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   3240
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   3240
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   3240
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   3240
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   3240
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   3240
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   4920
      Picture         =   "Tanggiam.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "&Save"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   4920
      Picture         =   "Tanggiam.frx":6C10
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "&Return"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   4800
      X2              =   4800
      Y1              =   120
      Y2              =   5880
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   4800
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   4800
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ng©n s¸ch :"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   46
      Tag             =   "Budget:"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tù bæ sung :"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   45
      Tag             =   "Capital:"
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "C¸c nguån kh¸c :"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   44
      Tag             =   "Others:"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ng©n s¸ch :"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   43
      Tag             =   "Budget:"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tù bæ sung :"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   42
      Tag             =   "Capital:"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "C¸c nguån kh¸c :"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   41
      Tag             =   "Others:"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   1680
      TabIndex        =   40
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   39
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   10
      Left            =   1800
      TabIndex        =   38
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   11
      Left            =   1800
      TabIndex        =   37
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   12
      Left            =   1680
      TabIndex        =   36
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   13
      Left            =   1800
      TabIndex        =   35
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   14
      Left            =   1800
      TabIndex        =   34
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   15
      Left            =   1800
      TabIndex        =   33
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   16
      Left            =   3240
      TabIndex        =   32
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   17
      Left            =   3240
      TabIndex        =   31
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "TÝn dông :"
      Height          =   255
      Index           =   22
      Left            =   240
      TabIndex        =   30
      Tag             =   "Credit:"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "TÝn dông :"
      Height          =   255
      Index           =   23
      Left            =   240
      TabIndex        =   29
      Tag             =   "Credit:"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   30
      Left            =   1800
      TabIndex        =   28
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   31
      Left            =   1800
      TabIndex        =   27
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   32
      Left            =   1800
      TabIndex        =   26
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   33
      Left            =   1680
      TabIndex        =   25
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gi¸ trÞ cßn l¹i"
      Height          =   255
      Index           =   37
      Left            =   240
      TabIndex        =   24
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   38
      Left            =   3360
      TabIndex        =   23
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   39
      Left            =   3360
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   40
      Left            =   3360
      TabIndex        =   21
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   41
      Left            =   3240
      TabIndex        =   20
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nguyªn gi¸"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Tag             =   "Original Cost:"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hao mßn"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   24
      Left            =   1800
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   35
      Left            =   3360
      TabIndex        =   16
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   42
      Left            =   1800
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   43
      Left            =   1800
      TabIndex        =   14
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "TÝn dông :"
      Height          =   255
      Index           =   18
      Left            =   240
      TabIndex        =   13
      Tag             =   "Credit:"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "C¸c nguån kh¸c :"
      Height          =   255
      Index           =   19
      Left            =   240
      TabIndex        =   12
      Tag             =   "Others:"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tù bæ sung :"
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   11
      Tag             =   "Capital:"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ng©n s¸ch :"
      Height          =   255
      Index           =   21
      Left            =   240
      TabIndex        =   10
      Tag             =   "Budget:"
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "frmTangGiam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NG_NS As Double
Dim NG_TBS As Double
Dim NG_CNK As Double
Dim NG_TD As Double

Dim NS As Double
Dim TBS As Double
Dim CNK As Double
Dim TD As Double

'======================================================================================
' FORM
'======================================================================================
' KeyDown
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      If (Shift And vbAltMask) > 0 Then
            Select Case KeyCode
                  Case vbKeyG: RFocus Command(0): DoEvents: Command_Click (0)
                  Case vbKeyV:  RFocus Command(1): DoEvents: Command_Click (1)
            End Select
      End If
      If KeyCode = vbKeyEscape Then Unload frmTangGiam
End Sub
' Load
Private Sub Form_Load()
      'pTiepTucHoatDong = False
      ' TÝnh gi¸ trÞ cßn l¹i cña tµi s¶n theo th¸ng t¸c ®éng (ch­a trÝch khÊu hao)
      TinhGiaTriTaiSan pMaTaiSan, pThangTacDong, KH_KHONG
      ' HiÓn thÞ c¸c th«ng tin
      With GiaTri
            ' Nguyªn gi¸
            NG_NS = .NG_NS
            NG_TBS = .NG_TBS
            NG_TD = .NG_TD
            NG_CNK = .NG_CNK
            Label(9).Caption = Format(.NG_NS, Mask_0)
            Label(10).Caption = Format(.NG_TBS, Mask_0)
            Label(11).Caption = Format(.NG_CNK, Mask_0)
            Label(42).Caption = Format(.NG_TD, Mask_0)
            Label(8).Caption = Format(.NG_NS + .NG_TBS + .NG_CNK + .NG_TD, Mask_0)
            ' Gi¸ trÞ cßn l¹i
            Label(32).Caption = Format(.CL_NS, Mask_0)
            Label(31).Caption = Format(.CL_TBS, Mask_0)
            Label(30).Caption = Format(.CL_CNK, Mask_0)
            Label(24).Caption = Format(.CL_TD, Mask_0)
            Label(33).Caption = Format(.CL_NS + .CL_TBS + .CL_CNK + .CL_TD, Mask_0)
            ' Tæng hao mßn
            Label(12).Caption = Format((.NG_NS + .NG_TBS + .NG_CNK + .NG_TD) - (.CL_NS + .CL_TBS + .CL_CNK + .CL_TD), Mask_0)
            Label(13).Caption = Format(.NG_NS - .CL_NS, Mask_0)
            Label(14).Caption = Format(.NG_TBS - .CL_TBS, Mask_0)
            Label(43).Caption = Format(.NG_TD - .CL_TD, Mask_0)
            Label(15).Caption = Format(.NG_CNK - .CL_CNK, Mask_0)
      End With
      NS = 0
      TBS = 0
      TD = 0
      CNK = 0
'      NG_NS = 0
'      NG_TBS = 0
'      NG_TD = 0
'      NG_CNK = 0
      XoaGiaTri         ' BiÕn chung GiaTri ®­îc khëi t¹o l¹i ®Ó nhËn th«ng tin
End Sub
'======================================================================================
' command
'     - NhËn l­îng t¨ng gi¶m gi¸ trÞ
'     - Thµnh lËp c¸c dßng ph¸t sinh thÓ hiÖn nghiÖp vô
'     - Ghi chøng tõ t¨ng gi¶m
'     - CËp nhËt l¹i l­îng khÊu hao míi
'======================================================================================
Private Sub Command_Click(Index As Integer)
    Dim sql As String
    
      Me.MousePointer = 11
      Select Case Index
            Case 0      ' Ghi
                  If HopLe = -1 Then GoTo XongTG
                  ' Thµnh lËp c¸c dßng ph¸t sinh thÓ hiÖn ho¹t ®éng t¨ng gi¶m
                  sql = "SELECT MaTaiKhoan AS F1 FROM TaiSan WHERE MaSo = " + CStr(pMaTaiSan)
                  ThanhLapPhatSinh NV_DGLAI, CLng5(SelectSQL(sql))
                  ' Ghi chøng tõ míi
                  pGhichungtu = 1
                  Unload Me
                  Exit Sub
            Case 1      ' Trë vÒ
                  SendKeys "{Escape}", False
      End Select
XongTG:
      Me.MousePointer = 0
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    If pGhichungtu = 0 Then XoaGiaTri
End Sub

'======================================================================================
' TEXT
'======================================================================================
' GotFocus
Private Sub Text_GotFocus(Index As Integer)
      AutoSelect Text(Index)
End Sub
' KeyPress
Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
      KeyProcess Text(Index), KeyAscii, True
End Sub
' LostFocus
Private Sub Text_LostFocus(Index As Integer)
      Text(Index).Text = Format(Text(Index).Text, Mask_0)
      On Error GoTo Err_DataTypeConvertion
      Select Case Index
            Case 0: GiaTri.NG_NS = Cdbl5(Text(0).Text)
            Case 1: GiaTri.NG_TBS = Cdbl5(Text(1).Text)
            Case 2: GiaTri.NG_CNK = Cdbl5(Text(2).Text)
            Case 3: GiaTri.NG_TD = Cdbl5(Text(3).Text)
            Case 4: NS = Cdbl5(Text(4).Text)
            Case 5: TBS = Cdbl5(Text(5).Text)
            Case 6: CNK = Cdbl5(Text(6).Text)
            Case 7: TD = Cdbl5(Text(7).Text)
      End Select
      On Error GoTo 0
      With GiaTri
            GiaTri.CL_NS = GiaTri.NG_NS - NS
            GiaTri.CL_TBS = GiaTri.NG_TBS - TBS
            GiaTri.CL_CNK = GiaTri.NG_CNK - CNK
            GiaTri.CL_TD = GiaTri.NG_TD - TD
            ' Tæng biÕn ®éng nguyªn gi¸ vµ khÊu hao
            Label(16).Caption = Format(.NG_NS + .NG_TBS + .NG_CNK + .NG_TD, Mask_0)
            Label(17).Caption = Format(NS + TBS + CNK + TD, Mask_0)
            ' L­îng biÕn ®éng t¸c ®éng lªn gi¸ trÞ cßn l¹i
            Label(40).Caption = Format(.CL_NS, Mask_0)
            Label(39).Caption = Format(.CL_TBS, Mask_0)
            Label(38).Caption = Format(.CL_CNK, Mask_0)
            Label(35).Caption = Format(.CL_TD, Mask_0)
            Label(41).Caption = Format(.CL_NS + .CL_TBS + .CL_CNK + .CL_TD, Mask_0)
      End With
      Exit Sub
Err_DataTypeConvertion:
      RFocus Text(Index)
End Sub
'======================================================================================
' FUNCTION HopLe
'======================================================================================
Private Function HopLe() As Integer
Dim thong_bao  As String
Dim ng1 As Double
      With GiaTri
            If (.NG_NS = 0 And .NG_TBS = 0 And .NG_CNK = 0 And .NG_TD = 0) And (NS = 0 And TBS = 0 And CNK = 0 And TD = 0) Then _
                  thong_bao = "Kh«ng cã biÕn ®éng g× vÒ gi¸ trÞ tµi s¶n": GoTo Err_InValidate
            If (NG_NS = 0 And .NG_NS = 0 And NS <> 0) Or (NG_TBS = 0 And .NG_TBS = 0 And TBS <> 0) Or (NG_CNK = 0 And .NG_CNK = 0 And CNK <> 0) Or (NG_TD = 0 And .NG_TD = 0 And TD <> 0) Then _
                  thong_bao = "L­îng hao mßn chØ cã nÕu tµi s¶n cã thµnh phÇn nguyªn gi¸ t­¬ng øng": GoTo Err_InValidate
'            If (Not .NG_NS = 0 And Label(9).Caption = "0") Or (Not .NG_TBS = 0 And Label(10).Caption = "0") Or (Not .NG_CNK = 0 And Label(11).Caption = "0") Or (Not .NG_TD = 0 And Label(42).Caption = "0") Then _
'                  thong_bao = "L­îng t¨ng gi¶m nguyªn gi¸ chØ cã nÕu tµi s¶n cã thµnh phÇn nguyªn gi¸ t­¬ng øng": GoTo Err_InValidate
            HopLe = 0
      End With
      Exit Function
Err_InValidate:
      Beep
      MsgBox thong_bao, vbCritical, App.ProductName
      HopLe = -1
End Function
