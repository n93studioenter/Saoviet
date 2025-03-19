VERSION 5.00
Begin VB.Form FrmPBCP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ph©n bæ chi phÝ"
   ClientHeight    =   3615
   ClientLeft      =   1500
   ClientTop       =   2415
   ClientWidth     =   6510
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
   Icon            =   "FrmPBCP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Expenses Allocation"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3615
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.OptionButton ChkKC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ tµi chÝnh"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   24
      Tag             =   "627"
      ToolTipText     =   "Pending Expenses "
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Xo¸ PB"
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
      Index           =   3
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "&Detail"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.OptionButton ChkKC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ m¸y thi c«ng"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   4
      Tag             =   "641"
      ToolTipText     =   "Machine expenses"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.OptionButton ChkKC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ nguyªn vËt liÖu trùc tiÕp"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   2
      Tag             =   "641"
      ToolTipText     =   "Raw material Expenses"
      Top             =   720
      Width           =   2655
   End
   Begin VB.OptionButton ChkKC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ nh©n c«ng trùc tiÕp"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Tag             =   "641"
      ToolTipText     =   "Labour Expenses"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Chi tiÕt"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "&Detail"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.OptionButton ChkKC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ chê ph©n bæ"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Tag             =   "627"
      ToolTipText     =   "Pending Expenses "
      Top             =   3240
      Width           =   2415
   End
   Begin VB.OptionButton ChkKC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ qu¶n lý doanh nghiÖp"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Tag             =   "642"
      ToolTipText     =   "Administrative Cost"
      Top             =   2520
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tiªu thøc ph©n bæ"
      Height          =   1935
      Left            =   2880
      TabIndex        =   19
      Tag             =   "Criteria"
      Top             =   720
      Width           =   2295
      Begin VB.CheckBox Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Chi phÝ s¶n xuÊt chung"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Tag             =   "6"
         ToolTipText     =   "General Expenses"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox T 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "100"
         Top             =   1560
         Width           =   495
      End
      Begin VB.CheckBox Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Chi phÝ m¸y thi c«ng"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Tag             =   "7"
         ToolTipText     =   "Machine expenses"
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nh©n c«ng trùc tiÕp"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Tag             =   "2"
         ToolTipText     =   "Labour Expenses"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nguyªn vËt liÖu trùc tiÕp"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Tag             =   "1"
         ToolTipText     =   "Raw material Expenses"
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Doanh thu"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Tag             =   "4"
         ToolTipText     =   "Turnover"
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label LbKC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tû lÖ ph©n bæ"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Tag             =   "Allocation Rate"
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label LbKC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "%"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   20
         Top             =   1590
         Width           =   255
      End
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      Index           =   1
      ItemData        =   "FrmPBCP.frx":57E2
      Left            =   4440
      List            =   "FrmPBCP.frx":57E4
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      Index           =   0
      ItemData        =   "FrmPBCP.frx":57E6
      Left            =   2280
      List            =   "FrmPBCP.frx":57E8
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton ChkKC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ b¸n hµng"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Tag             =   "641"
      ToolTipText     =   "Sale expense"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton ChkKC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chi phÝ s¶n xuÊt chung"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Tag             =   "627"
      ToolTipText     =   "General Expense"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   5280
      Picture         =   "FrmPBCP.frx":57EA
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "&Return"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ph©n bæ"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "&Allocate"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label LbKC 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "®Õn th¸ng"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   18
      Tag             =   "to"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LbKC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ph©n bæ chi phÝ tõ th¸ng"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Tag             =   "Expenses distributed from"
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmPBCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim thang(1 To 12) As Integer

Private Sub ChkKC_Click(Index As Integer)
    Select Case Index
        Case 1:
            Opt(0).Enabled = False
            If Opt(0).Value Then Opt(1).Value = 1
            Opt(1).Enabled = True
            Opt(2).Enabled = True
            Opt(3).Enabled = True
            Opt(4).Enabled = False
        Case 0, 2:
            Opt(0).Enabled = True
            If Not Opt(0).Value Then Opt(0).Value = 1
            Opt(1).Enabled = True
            Opt(2).Enabled = True
            Opt(3).Enabled = True
            Opt(4).Enabled = True
        Case 4:
            Opt(0).Enabled = False
            If Opt(0).Value Then Opt(1).Value = 1
            Opt(1).Enabled = True
            Opt(2).Enabled = False
            If Opt(2).Value Then Opt(2).Value = 0
            Opt(3).Enabled = True
            Opt(4).Enabled = False
        Case 5:
            Opt(0).Enabled = False
            If Opt(0).Value Then Opt(2).Value = 1
            Opt(1).Enabled = False
            If Opt(1).Value Then Opt(1).Value = 0
            Opt(2).Enabled = True
            Opt(3).Enabled = True
            Opt(4).Enabled = False
        Case 6:
            Opt(0).Enabled = False
            If Opt(0).Value Then Opt(1).Value = 1
            Opt(1).Enabled = True
            Opt(2).Enabled = True
            Opt(3).Enabled = False
            If Opt(3).Value Then Opt(3).Value = 0
            Opt(4).Enabled = False
    End Select
End Sub
'====================================================================================================
' Thªm, Ghi, Xãa nguyªn tÖ
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim tc As Integer, i As Integer, sql As String, tdau As Integer, tcuoi As Integer
    
    If Cbo(1).ListIndex < Cbo(0).ListIndex Then Cbo(1).ListIndex = Cbo(0).ListIndex
    tdau = Cbo(0).ItemData(Cbo(0).ListIndex)
    tcuoi = Cbo(1).ItemData(Cbo(1).ListIndex)
    Select Case Index
        Case 0:
            tc = 0
            For i = 0 To 4
                If Opt(i).Value = 1 Then tc = tc + Opt(i).tag
            Next
            Me.MousePointer = 11
            If ChkKC(0).Value Then
                PhanBoCP tdau, tcuoi, "641", ABCtoVNI("Chi phÝ b¸n hµng"), 4, CInt5(T.Text)
                PhanBoCP64 tdau, tcuoi, tc, "641"
            End If
            If ChkKC(1).Value Then
                PhanBoCP tdau, tcuoi, "627", ABCtoVNI("Chi phÝ s¶n xuÊt chung"), tc, CInt5(T.Text)
                If pDTTP <> 0 Then PhanBoCP2 tdau, tcuoi, tc
            End If
            If ChkKC(2).Value Then
                PhanBoCP tdau, tcuoi, "642", ABCtoVNI("Chi phÝ qu¶n lý doanh nghiÖp"), 4, CInt5(T.Text)
                PhanBoCP64 tdau, tcuoi, tc, "642"
            End If
            If ChkKC(3).Value Then PhanBoCP tdau, tcuoi, "142", ABCtoVNI("Chi phÝ tr¶ tr­íc"), 4, CInt5(T.Text)
            If ChkKC(4).Value Then
                PhanBoCP tdau, tcuoi, "622", ABCtoVNI("Chi phÝ nh©n c«ng trùc tiÕp"), 4, CInt5(T.Text)
                If pDTTP <> 0 Then PhanBoCP3 tdau, tcuoi, tc, "622"
            End If
            If ChkKC(5).Value Then
                PhanBoCP tdau, tcuoi, "621", ABCtoVNI("Chi phÝ nguyªn vËt liÖu trùc tiÕp"), 4, CInt5(T.Text)
                If pDTTP <> 0 Then PhanBoCP3 tdau, tcuoi, tc, "621"
            End If
            If ChkKC(6).Value Then
                PhanBoCP tdau, tcuoi, "623", ABCtoVNI("Chi phÝ m¸y thi c«ng"), 4, CInt5(T.Text)
                If pDTTP <> 0 Then PhanBoCP3 tdau, tcuoi, tc, "623"
            End If
            If ChkKC(7).Value Then
                PhanBoCP tdau, tcuoi, "635", ABCtoVNI("Chi phÝ tµi chÝnh"), 4, CInt5(T.Text)
                PhanBoCP64 tdau, tcuoi, tc, "635"
            End If
            SetDBMonth
        Case 1:
            Unload Me
        Case 2:
            If pDTTP = 0 Then GoTo KT
            If Cbo(0).ListIndex <> Cbo(1).ListIndex Then
                FrmPBCT.Caption = "Chi tiÕt tËp hîp chi phÝ tõ th¸ng " + Cbo(0).Text + " ®Õn th¸ng " + Cbo(1).Text
            Else
                FrmPBCT.Caption = "Chi tiÕt tËp hîp chi phÝ th¸ng " + Cbo(0).Text
            End If
            FrmPBCT.tag = 100 * (tdau) + tcuoi
            FrmPBCT.Show 1
            If Me.tag = 1 Then SetDBMonth
        Case 3:
            If pDTTP = 0 Then GoTo KT
            Me.MousePointer = 11
            sql = ""
            For i = CThangDB(tdau) To CThangDB(tcuoi)
                If ChkKC(0).Value Then sql = sql + ", CPBH" + CStr(i) + "=0"
                If ChkKC(1).Value Then sql = sql + ", CPSXC_" + CStr(i) + "=0"
                If ChkKC(2).Value Then sql = sql + ", CPQL" + CStr(i) + "=0"
                If ChkKC(3).Value Then GoTo KT
                If ChkKC(4).Value Then sql = sql + ", CPNCPB" + CStr(i) + "=0"
                If ChkKC(5).Value Then sql = sql + ", CPNVLPB" + CStr(i) + "=0"
                If ChkKC(6).Value Then sql = sql + ", CPMPB" + CStr(i) + "=0"
            Next
            ExecuteSQL5 "UPDATE TP154 SET DonVi=DonVi" + sql
            sql = ""
            If ChkKC(0).Value Then sql = sql + ", CPBH=0"
            If ChkKC(1).Value Then sql = sql + ", CPQL=0"
            If ChkKC(2).Value Then sql = sql + ", CPTC=0"
            ExecuteSQL5 "UPDATE ThanhPham SET Thang=Thang" + sql + " WHERE " + WThang("Thang", tdau, tcuoi)
    End Select
KT:
    Me.MousePointer = 0
End Sub
'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyK:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyV:
                RFocus Command(1)
                Command_Click 1
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
'====================================================================================================
' Khëi t¹o cöa sæ
'====================================================================================================
Private Sub Form_Load()
    AddMonthToCbo Cbo(0)
    AddMonthToCbo Cbo(1)
    ChkKC_Click 2
    Command(2).Enabled = (pDTTP <> 0)
    Command(3).Enabled = (pDTTP <> 0)
    SetFont Me
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Dim i As Integer, tdau As Integer, tcuoi As Integer
    
    Me.MousePointer = 11
    For i = 1 To 12
        If thang(i) = 1 Then
            tdau = CThangFR(i)
            Exit For
        End If
    Next
    
    For i = 12 To 1 Step -1
        If thang(i) = 1 Then
            tcuoi = CThangFR(i)
            Exit For
        End If
    Next
        
    If tdau > 0 And tcuoi > 0 And pDTTP <> 0 Then TinhLaiGiaThanhPham tdau, tcuoi
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

Private Sub Opt_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 0:
            If Opt(0).Value = 1 Then
                For i = 1 To 4
                    Opt(i).Value = 0
                Next
            End If
        Case 1, 2, 3:
            If Opt(Index).Value = 1 Then
                Opt(0).Value = 0
                Opt(4).Value = 0
            End If
        Case 4:
            If Opt(4).Value = 1 Then
                For i = 0 To 3
                    Opt(i).Value = 0
                Next
            End If
    End Select
End Sub

Private Sub T_GotFocus()
    AutoSelect T
End Sub

Private Sub T_KeyPress(KeyAscii As Integer)
    KeyProcess T, KeyAscii
End Sub

Private Sub SetDBMonth()
    Dim i As Integer
    
    For i = Cbo(0).ListIndex + 1 To Cbo(1).ListIndex + 1
        thang(i) = 1
    Next
End Sub

