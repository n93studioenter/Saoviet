VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmTaikhoan 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HÖ thèng tµi kho¶n"
   ClientHeight    =   7200
   ClientLeft      =   5250
   ClientTop       =   3225
   ClientWidth     =   11535
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
   Icon            =   "Frmtkn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Chart of Account"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7200
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Th«ng tin c«ng tr×nh, h¹ng môc"
      ForeColor       =   &H80000008&
      Height          =   3975
      Index           =   1
      Left            =   -1680
      TabIndex        =   46
      Top             =   2640
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   960
         MaxLength       =   50
         TabIndex        =   47
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   960
         MaxLength       =   2
         TabIndex        =   48
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   960
         MaxLength       =   50
         TabIndex        =   50
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtDu 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   2640
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   49
         Text            =   "Frmtkn.frx":57E2
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtDu 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   960
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   51
         Text            =   "Frmtkn.frx":57E6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtDu 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   960
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   52
         Text            =   "Frmtkn.frx":57EA
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtDu 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   960
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   53
         Text            =   "Frmtkn.frx":57EE
         Top             =   2160
         Width           =   1455
      End
      Begin MSMask.MaskEdBox MedNgay 
         Height          =   315
         Index           =   0
         Left            =   3360
         TabIndex        =   54
         Top             =   1800
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
         Left            =   3360
         TabIndex        =   55
         Top             =   2160
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
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tªn DA"
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   615
      End
      Begin VB.Line Line 
         Index           =   5
         X1              =   960
         X2              =   4320
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nhãm DA"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   63
         Top             =   600
         Width           =   735
      End
      Begin VB.Line Line 
         Index           =   6
         X1              =   960
         X2              =   1320
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "§Þa ®iÓm"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   62
         Top             =   960
         Width           =   735
      End
      Begin VB.Line Line 
         Index           =   7
         X1              =   960
         X2              =   4320
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dù to¸n"
         Height          =   255
         Index           =   19
         Left            =   1920
         TabIndex        =   61
         Top             =   600
         Width           =   735
      End
      Begin VB.Line Line 
         Index           =   8
         X1              =   2640
         X2              =   4320
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vèn NS"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   60
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vèn vay"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   59
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vèn kh¸c"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   58
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ngµy KC"
         Height          =   255
         Index           =   23
         Left            =   2520
         TabIndex        =   57
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ngµy HT"
         Height          =   255
         Index           =   24
         Left            =   2520
         TabIndex        =   56
         Top             =   2160
         Width           =   735
      End
      Begin VB.Line Line 
         Index           =   9
         X1              =   960
         X2              =   2400
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Line Line 
         Index           =   10
         X1              =   960
         X2              =   2400
         Y1              =   2085
         Y2              =   2085
      End
      Begin VB.Line Line 
         Index           =   11
         X1              =   960
         X2              =   2400
         Y1              =   2445
         Y2              =   2445
      End
   End
   Begin VB.TextBox txtDu 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   6
      Left            =   6840
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "Frmtkn.frx":57F2
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Chi tiÕt CT"
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
      Index           =   4
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   5760
      MaxLength       =   60
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ComboBox CboNT 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   6720
      Width           =   2655
      Begin VB.OptionButton SSOpt 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tªn TK"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   27
         Tag             =   "Desciption"
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton SSOpt 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sè hiÖu"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Tag             =   "Code"
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   5760
      MaxLength       =   60
      TabIndex        =   4
      Top             =   720
      Width           =   3735
   End
   Begin VB.OptionButton OptNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nî"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Tag             =   "Debit"
      Top             =   1560
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton OptCo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cã"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   7
      Tag             =   "Credit"
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox ChkBtc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bé tµi chÝnh quy ®Þnh"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7560
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "Regulated by MF"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtDu 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   6360
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Frmtkn.frx":57F6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtDu 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   6360
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "Frmtkn.frx":57FA
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtDuNT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "Frmtkn.frx":57FE
      Top             =   3960
      Width           =   1635
   End
   Begin VB.TextBox txtDuNT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "Frmtkn.frx":5800
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox txtDuNT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "Frmtkn.frx":5802
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton CmdNT 
      DragIcon        =   "Frmtkn.frx":5804
      BeginProperty Font 
         Name            =   "Vni 14 AlexBrush"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11160
      Picture         =   "Frmtkn.frx":AFE6
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox ChkDT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Theo dâi chi tiÕt"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Tag             =   "Detailed"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   5760
      MaxLength       =   50
      TabIndex        =   17
      Top             =   6075
      Width           =   3735
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   0
      Left            =   6720
      Picture         =   "Frmtkn.frx":B388
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "&Add"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   1
      Left            =   7920
      Picture         =   "Frmtkn.frx":C8E2
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "&Save"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   2
      Left            =   9120
      Picture         =   "Frmtkn.frx":DD10
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "&Delete"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   3
      Left            =   10320
      Picture         =   "Frmtkn.frx":F1F2
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "&Return"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txtF 
      Height          =   285
      Left            =   2880
      TabIndex        =   28
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton SSCmdF 
      BeginProperty Font 
         Name            =   "Vni 14 AlexBrush"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      TabIndex        =   25
      ToolTipText     =   "T×m kiÕm"
      Top             =   6840
      Width           =   255
   End
   Begin VB.ComboBox CboLoai 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin MSOutl.Outline OtlTk 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   10821
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Frmtkn.frx":10614
      PicturePlus     =   "Frmtkn.frx":10630
      PictureMinus    =   "Frmtkn.frx":1072A
      PictureLeaf     =   "Frmtkn.frx":10824
      PictureOpen     =   "Frmtkn.frx":1091E
      PictureClosed   =   "Frmtkn.frx":10A18
   End
   Begin MSGrid.Grid GrdNT 
      Height          =   735
      Index           =   0
      Left            =   5040
      TabIndex        =   30
      Tag             =   "10"
      Top             =   3240
      Width           =   6375
      _Version        =   65536
      _ExtentX        =   11245
      _ExtentY        =   1296
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
      Rows            =   10
      Cols            =   5
      FixedRows       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin MSGrid.Grid GrdNT 
      Height          =   735
      Index           =   1
      Left            =   5040
      TabIndex        =   31
      Tag             =   "10"
      Top             =   4800
      Width           =   6375
      _Version        =   65536
      _ExtentX        =   11245
      _ExtentY        =   1296
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
      Rows            =   10
      Cols            =   4
      FixedRows       =   0
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CP luü kÕ ®Õn ®Çu n¨m"
      Height          =   255
      Index           =   26
      Left            =   5040
      TabIndex        =   66
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Line Line 
      Index           =   12
      X1              =   6840
      X2              =   8400
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ng.tÖ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   35
      Tag             =   "Curr."
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nî"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   34
      Tag             =   "Debit"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Line Line1 
      Index           =   5
      Visible         =   0   'False
      X1              =   5760
      X2              =   9480
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Desc."
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   45
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè hiÖu"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   43
      Tag             =   "Code"
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tªn"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   42
      Tag             =   "Desc (V)"
      Top             =   720
      Width           =   615
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   5760
      X2              =   7080
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   5760
      X2              =   9480
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "KiÓu"
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   41
      Tag             =   "Type"
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cã"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7560
      TabIndex        =   40
      Tag             =   "Credit"
      Top             =   3000
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L­îng NT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   9240
      TabIndex        =   39
      Tag             =   "F. Currency"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè d­ tèi thiÓu"
      Height          =   255
      Index           =   9
      Left            =   5040
      TabIndex        =   38
      Tag             =   "Min Balance"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè d­ tèi ®a"
      Height          =   255
      Index           =   10
      Left            =   5040
      TabIndex        =   37
      Tag             =   "Max Balance"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   6360
      X2              =   7680
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   6360
      X2              =   7680
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè d­ ®Çu kú"
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   36
      Tag             =   "Opening Balance"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè d­ hiÖn thêi"
      Height          =   255
      Index           =   16
      Left            =   5040
      TabIndex        =   33
      Tag             =   "Closing Balance"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ghi chó"
      Height          =   195
      Index           =   14
      Left            =   5040
      TabIndex        =   32
      Tag             =   "Notes"
      Top             =   6120
      Width           =   555
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   5760
      X2              =   9480
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6495
      Index           =   15
      Left            =   4800
      TabIndex        =   29
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   12
      Left            =   4800
      TabIndex        =   24
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   11
      Left            =   4800
      TabIndex        =   23
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFC0C0&
      Height          =   6015
      Index           =   13
      Left            =   180
      TabIndex        =   22
      Top             =   540
      Width           =   4695
   End
End
Attribute VB_Name = "FrmTaikhoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer          ' = 1 neu them moi                                                      ' =-1 neu sua tai khoan                                                      ' = 0 trong che do xem
Dim TkNhap As New ClsTaikhoan   ' tai khoan tham chieu
Dim enable As Boolean
Dim ngay(0 To 1) As Date

'====================================================================================================
' HiÓn thÞ danh s¸ch tµi kho¶n trong lo¹i
'====================================================================================================
Private Sub CboLoai_Click()
    If enable Then ListTk
End Sub

Private Sub CboNT_Click()
    If CboNT.ItemData(CboNT.ListIndex) = 0 Then
        txtDuNT(2).Text = "0"
        txtDuNT(2).Enabled = False
    Else
        txtDuNT(2).Enabled = True
    End If
End Sub

Private Sub cmdNt_Click()
    Dim i As Integer, no As Double, co As Double, nt As Double, mnt As Long
    Dim taikhoan As New ClsTaikhoan, capnhatphatsinh As Boolean, st  As String
    
    If Not ChoDieuChinhDauKy Then Exit Sub
            
    If (Left(TkNhap.sohieu, 3)) = "154" And pDTTP <> 0 Then
        MsgBox "Theo dâi chi tiÕt qua danh ®iÓm thµnh phÈm!", vbExclamation, App.ProductName
        Exit Sub
    End If
            
    If ThemMoi = 0 And TkNhap.tkcon > 0 Then
        MsgBox "Tµi kho¶n cã chi tiÕt kh«ng nhËp sè liÖu trùc tiÕp!", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    If TkNhap.tk_id = TKVT_ID And STDetail Then
        MsgBox "Kh«ng nhËp trùc tiÕp sè d­ vµo tµi kho¶n vËt t­, hµng hãa. H·y nhËp tån kho ®Çu kú !", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    If (TkNhap.tk_id = TSCD_ID Or TkNhap.tk_id = KHTSCD_ID) And FADetail Then
        MsgBox "Kh«ng nhËp trùc tiÕp sè d­ vµo tµi kho¶n tµi s¶n cè ®Þnh. H·y nhËp tµi s¶n ®Çu kú !", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    If (TkNhap.tk_id = TKCNKH_ID Or TkNhap.tk_id = TKCNPT_ID) And KHDetail Then
        MsgBox "Kh«ng nhËp trùc tiÕp sè d­ vµo tµi kho¶n c«ng nî. H·y nhËp c«ng nî ®Çu kú !", vbExclamation, App.ProductName
        Exit Sub
    End If
    
    st = txt(1).Text
        
    no = Cdbl5(txtDuNT(0).Text)
    co = Cdbl5(txtDuNT(1).Text)
    nt = Cdbl5(txtDuNT(2).Text)
    On Error GoTo 0
    If no <> 0 And co <> 0 Then
        RFocus txtDuNT(0)
        Exit Sub
    End If
    
    mnt = CboNT.ItemData(CboNT.ListIndex)
    With GrdNT(0)
        .col = 4
        For i = 0 To .Rows - 1
            .Row = i
            If Len(.Text) = 0 Then Exit For
            If CInt5(.Text) = mnt Or (mnt <= 0 And CInt5(.Text) <= 0) Then
                .RemoveItem i
                .AddItem CboNT.Text + Chr(9) + Format(no, IIf(Left(TkNhap.sohieu, 3) <> "007", Mask_0, Mask_2)) + Chr(9) + Format(co, IIf(Left(TkNhap.sohieu, 3) <> "007", Mask_0, Mask_2)) _
                + Chr(9) + Format(nt, Mask_2) + Chr(9) + CStr(mnt), i
                If ThemMoi = 0 Then
                    taikhoan.InitTaikhoanMaSo MaTKNguyenTe(TkNhap.sohieu, mnt)
                    taikhoan.NoDauKy = no
                    taikhoan.CoDauKy = co
                    taikhoan.NTDauKy = nt
                    taikhoan.CapNhatTk
                    OtlTk_Click
                    txt(1).Text = st
                End If
                RFocus CboNT
                Exit Sub
            End If
        Next
        If (ThemMoi = 0) And (TkNhap.MaTC = 0 Or TkNhap.MaTC = TkNhap.MaSo Or TkNhap.tkcon > 0) Then Exit Sub
        .AddItem CboNT.Text + Chr(9) + Format(no, Mask_0) + Chr(9) + Format(co, Mask_0) _
                + Chr(9) + Format(nt, Mask_2) + Chr(9) + CStr(mnt), 0
        .Row = .Rows - 1
        .col = 4
        If Len(.Text) = 0 Then .RemoveItem .Row
        .Row = 0
        If ThemMoi = 0 Then
            TkNhap.DanXuat taikhoan
            taikhoan.MaSo = 0
            taikhoan.MaNT = CboNT.ItemData(CboNT.ListIndex)
            taikhoan.NoDauKy = no
            taikhoan.CoDauKy = co
            taikhoan.NTDauKy = nt
            taikhoan.ThemTk
            txt(1).Text = st
            OtlTk_Click
        End If
        RFocus CboNT
    End With
End Sub
'======================================================================================
' Thªm, Ghi, Xãa tµi kho¶n
'======================================================================================
Private Sub Command_Click(Index As Integer)
    Dim i As Integer, cap As Integer, Length As Integer, j As Integer
    Dim capnhatsodu As Boolean, taikhoan As New ClsTaikhoan
    Dim rs_tk As Recordset, tail As String
    
    If (User_Right = 2) And (Index < 3) Then
        HienThongBao "Kh«ng cã quyÒn truy cËp!", 1
        Exit Sub
    End If
    
    If Index < 3 Then
        If (Not KtraMKAdmin) Then Exit Sub
    End If
    
    Me.MousePointer = 11
    Select Case Index
        Case 0:
            If (TkNhap.tk_id = TSCD_ID And FADetail) Then
                MsgBox "Theo dâi chi tiÕt qua danh ®iÓm tµi s¶n!", vbExclamation, App.ProductName
                GoTo XongTK
            End If
            If TkNhap.MaTC = 0 Then
                ChkBtc.Value = 1
            Else
                ChkBtc.Value = 0
            End If
            ThemMoi = 1
            txt(0).Text = ""
            txt(1).Text = ""
            txt(2).Text = "..."
            txt(3).Text = "..."
            ClearGrid GrdNT(0), 10
            ClearGrid GrdNT(1), 10
            CboNT.ListIndex = 0
            GrdNT(0).AddItem pTienStr + Chr(9) + "0" + Chr(9) + "0" + Chr(9) + "0" + Chr(9) + "0", 0
            If TkNhap.tkcon > 0 Then
                Set rs_tk = DBKetoan.OpenRecordset("SELECT Top 1 HeThongTK.SoHieu FROM HethongTK" _
                    & " WHERE TkCha0=" + CStr(TkNhap.MaSo) + " AND MaNT<=0 ORDER BY SoHieu DESC", dbOpenSnapshot)
                If rs_tk.RecordCount > 0 Then
                    Length = Len(rs_tk!sohieu)
                    On Error Resume Next
                    tail = CStr(CInt5(Right(rs_tk!sohieu, Length - Len(TkNhap.sohieu))) + 1)
                    On Error GoTo 0
                Else
                    Length = Len(TkNhap.sohieu) + 2
                End If
                rs_tk.Close
                Set rs_tk = Nothing
                Do While Len(TkNhap.sohieu + tail) < Length
                    tail = "0" + tail
                Loop
                txt(0).Text = TkNhap.sohieu + tail
            Else
                txt(0).Text = TkNhap.sohieu + "01"
            End If
            RFocus txt(0)
        Case 1:
            If KiemTraDuLieu Then
                Select Case ThemMoi
                    Case 0:
                        If TkNhap.CapNhatTk() = 0 Then
                            OtlTk.List(OtlTk.ListIndex) = TkNhap.sohieu + " " + TkNhap.Ten
                        Else
                            ErrMsg er_SoHieu
                            RFocus txt(0)
                            GoTo XongTK
                        End If
                    Case 1
                        ' Them tai khoan tren man hinh
                        TkNhap.PSNLK = 0
                        TkNhap.PSCLK = 0
                        If TkNhap.ThemTk() = -1 Then
                            ErrMsg er_SoHieu
                            RFocus txt(0)
                            GoTo XongTK
                        End If
                        TkNhap.DanXuat taikhoan
                        taikhoan.MaSo = 0
                        With GrdNT(0)
                            For i = 0 To .Rows - 1
                                .Row = i
                                .col = 4
                                If Len(.Text) = 0 Then Exit For
                                taikhoan.MaNT = CInt5(.Text)
                                If taikhoan.MaNT <> 0 Then
                                    .col = 1
                                    taikhoan.NoDauKy = Cdbl5(.Text)
                                    .col = 2
                                    taikhoan.CoDauKy = Cdbl5(.Text)
                                    .col = 3
                                    taikhoan.NTDauKy = Cdbl5(.Text)
                                    taikhoan.ThemTk
                                End If
                            Next
                        End With
                        i = OtlTk.ListIndex
                        j = i
                        If OtlTk.HasSubItems(i) Then
                             i = i + 1
                             cap = OtlTk.indent(i)
                             Length = Len(TkNhap.sohieu)
                             Do While (OtlTk.indent(i) >= cap Or TkNhap.cap = 1) And (TkNhap.sohieu > Left(OtlTk.List(i), Length) And (i < OtlTk.ListCount - 1))
                                 i = i + 1
                             Loop
                             If i = OtlTk.ListCount - 1 And (TkNhap.cap = OtlTk.indent(i) Or TkNhap.cap = 1) Then i = i + 1
                        Else
                            i = i + 1
                        End If
                        OtlTk.AddItem TkNhap.sohieu + " " + TkNhap.Ten, i
                        OtlTk.indent(i) = TkNhap.cap
                        OtlTk.ItemData(i) = TkNhap.MaSo
                        OtlTk.Expand(j) = True
                        ThemMoi = 0
                        OtlTk.Refresh
                End Select
                RFocus OtlTk
            End If
        Case 2:
            Select Case ThemMoi
                Case 0:
                Dim so
                so = SelectSQL("SELECT Count(MaSo) As F1 FROM HethongTk WHERE THEMMOI = '1' and MaSo=" + CStr(TkNhap.MaSo))
'                If so = 0 Then
'                     i = 1000
'                  Else
                    i = TkNhap.XoaTk
 '               End If
                    If TkNhap.MaNT = 0 Then
                        Select Case i
                            Case -2:
                                MsgBox "H·y xo¸ hÕt chi tiÕt cña tµi kho¶n !", vbExclamation, App.ProductName
                            Case -3:
                                ErrMsg er_CoPS
                            Case -4:
                                MsgBox "Tµi kho¶n cã chi tiÕt nguyªn tÖ, h·y xo¸ c¸c chi tiÕt nguyªn tÖ !", vbExclamation, App.ProductName
                            Case -5:
                            Case 0:
                                i = OtlTk.ListIndex
                                If i > 0 Then
                                    i = i - 1
                                    Do While Not OtlTk.IsItemVisible(i) And i > 0
                                        i = i - 1
                                    Loop
                                End If
                                OtlTk.RemoveItem OtlTk.ListIndex
                                If OtlTk.ListCount > 0 Then
                                    OtlTk.ListIndex = i
                                    OtlTk_Click
                                End If
                        End Select
                    Else
                        Dim TK As New ClsTaikhoan
                        Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo FROM HethongTK WHERE SoHieu = '" + TkNhap.sohieu + "' ORDER BY MaNT DESC", dbOpenSnapshot)
                        Do While Not rs_tk.EOF
                            TK.InitTaikhoanMaSo rs_tk!MaSo
                            TK.XoaTk
                            rs_tk.MoveNext
                        Loop
                        rs_tk.Close
                        Set rs_tk = Nothing
                        i = OtlTk.ListIndex
                        OtlTk.RemoveItem OtlTk.ListIndex
                        OtlTk.ListIndex = i - 1
                        OtlTk_Click
                        OtlTk.Refresh
                    End If
                Case Else
                    ThemMoi = 0
                    OtlTk_Click
            End Select
            RFocus OtlTk
        Case 3:
            Hide
        Case 4:
            Frame(1).Visible = Not Frame(1).Visible
            If Frame(1).Visible Then RFocus txt(6)
    End Select
XongTK:
    Me.MousePointer = 0
End Sub

Private Sub Form_Activate()
    If ThemMoi = 0 And FrmTaikhoan.tag = 0 Then RFocus OtlTk
End Sub

'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      Dim i As Integer
      
      If (Shift And vbAltMask) > 0 Then
            i = -1
            Select Case KeyCode
                  Case vbKeyV: i = 3
                  Case vbKeyT: i = 0
                  Case vbKeyG: i = 1
                  Case vbKeyX: i = 2
            End Select
            If i >= 0 Then
                If Command(i).Enabled Then
                        RFocus Command(i)
                        DoEvents
                        Command_Click (i)
                End If
            End If
      End If
          
    If KeyCode = vbKeyAdd Then
        If OtlTk.HasSubItems(OtlTk.ListIndex) Then OtlTk.Expand(OtlTk.ListIndex) = True
        KeyCode = 0
    End If
    If KeyCode = vbKeySubtract Then
        If OtlTk.HasSubItems(OtlTk.ListIndex) Then OtlTk.Expand(OtlTk.ListIndex) = False
        KeyCode = 0
    End If
    If KeyCode = vbKeyEscape Then Hide
End Sub
'====================================================================================================
' Khëi t¹o cöa sæ
'====================================================================================================
Private Sub Form_Load()
    Dim chi_so As Integer
    
    ColumnSetUp GrdNT(0), 0, 700, 2
    ColumnSetUp GrdNT(0), 1, 1180 + 600, 1
    ColumnSetUp GrdNT(0), 2, 1180 + 600, 1
    ColumnSetUp GrdNT(0), 3, 1060 + 600, 1
    ColumnSetUp GrdNT(0), 4, 1, 0
    
    ColumnSetUp GrdNT(1), 0, 700, 2
    ColumnSetUp GrdNT(1), 1, 1180 + 600, 1
    ColumnSetUp GrdNT(1), 2, 1180 + 600, 1
    ColumnSetUp GrdNT(1), 3, 1060 + 600, 1
    
    For chi_so = 0 To 1
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next
    
    enable = False
    ' Liet ke danh sach loai tai khoan
    Int_RecsetToCbo "SELECT Loai As F2,Ten" + IIf(pSongNgu And pVersion <> 3, "+' - '+TenE", "") + " As F1 FROM HethongTk WHERE Cap=0 AND Len(SoHieu)>1 ORDER BY SoHieu", CboLoai
    Int_RecsetToCbo "SELECT MaSo As F2,KyHieu As F1 FROM NguyenTe WHERE KyHieu<>'" + pTienStr + "' ORDER BY KyHieu", CboNT
    CboNT.AddItem pTienStr, 0
    CboNT.ItemData(0) = 0
    enable = True
    CboLoai.ListIndex = 1
    ThemMoi = 0
'    xPhatsinh = False
    CmdNT.Enabled = (User_Right <> 2)
    
    Caption = Caption + " - " + CStr(pNamTC)
    
    Label(8).Visible = pSongNgu
    txt(3).Visible = pSongNgu
    Line1(5).Visible = pSongNgu
    
    SetFont Me
     ExecuteSQL5_Themmoi ("ALTER TABLE HethongTk ADD THEMMOI text")
End Sub
'====================================================================================================
' HiÓn thÞ danh s¸ch tµi kho¶n thuéc mét lo¹i
'====================================================================================================
Private Sub ListTk()
    Dim RsTk As Recordset
    Dim i As Integer
    
    OtlTk.Clear
    Set RsTk = DBKetoan.OpenRecordset("SELECT DISTINCTROW First(HeThongTK.MaSo) AS FirstOfMaSo, HeThongTK.SoHieu, First(HeThongTK.Cap) AS FirstOfCap, First(HeThongTK.Ten" + IIf(pNN = 1, "E", "") + ") AS FirstOfTen" _
        & " From HeThongTK WHERE LEFT(SoHieu,1)<>'#' AND Loai=" + str(CboLoai.ItemData(CboLoai.ListIndex)) + " And Cap>0 And MaNT <=0 GROUP BY HeThongTK.SoHieu, HethongTK.MaNT ORDER BY HeThongTK.SoHieu, HeThongTK.MaNT", dbOpenSnapshot)
    
    i = 0
    On Error Resume Next
    Do While Not RsTk.EOF
        OtlTk.AddItem RsTk!sohieu + " " + RsTk!FirstOfTen
        OtlTk.indent(i) = RsTk!FirstOfCap
        OtlTk.ItemData(i) = RsTk!FirstOfMaSo
        i = i + 1
        RsTk.MoveNext
    Loop
    On Error GoTo 0
    
    RsTk.Close
    Set RsTk = Nothing
    If OtlTk.ListCount > 0 Then
        OtlTk.ListIndex = 0
        Command(0).Enabled = True
        Command(1).Enabled = True
        Command(2).Enabled = True
        OtlTk_Click
    Else
        Command(0).Enabled = False
        Command(1).Enabled = False
        Command(2).Enabled = False
        ThemMoi = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set TkNhap = Nothing
End Sub

Private Sub GrdNT_DblClick(Index As Integer)
If Index = 0 Then
    With GrdNT(0)
        .col = 4
        If Len(.Text) = 0 Then Exit Sub
        If CInt5(.Text) <= 0 Then Exit Sub
        If ThemMoi = 0 Then
            Dim taikhoan As New ClsTaikhoan
            
            taikhoan.InitTaikhoanMaSo MaTKNguyenTe(TkNhap.sohieu, CInt5(.Text))
            If taikhoan.TkCoPS(0, 0) Then
                ErrMsg er_CoPS
                Exit Sub
            Else
                taikhoan.XoaTk
                OtlTk_Click
            End If
        Else
            .RemoveItem .Row
            If .Rows < 10 Then .Rows = 10
        End If
    End With
End If
End Sub
'====================================================================================================
' HiÓn thÞ th«ng tin vÒ tµi kho¶n ®­îc chän
'====================================================================================================
Private Sub OtlTk_Click()
    Dim dkn As Double, dkc As Double, dknt As Double, ckn As Double, ckc As Double, cknt As Double
    Dim rs_tk As Recordset, sh As String, sql As String
    
    txtDuNT(0).Text = "0"
    txtDuNT(1).Text = "0"
    txtDuNT(2).Text = "0"
    ' Khoi tao tham chieu
With TkNhap
    .InitTaikhoanMaSo OtlTk.ItemData(OtlTk.ListIndex)
    Frame(1).Visible = False
    Command(4).Visible = (Left(TkNhap.sohieu, 4) = "2412" And (TkNhap.cap >= 3) And frmMain.Chk(2).Value <> 0)
    ChkDT.Visible = False
    If .tk_id = TKVT_ID Then
        ChkDT.Value = IIf(STDetail, 1, 0)
        ChkDT.Visible = True
    End If
    If (.tk_id = TSCD_ID Or .tk_id = KHTSCD_ID) Then
        ChkDT.Value = IIf(FADetail, 1, 0)
        ChkDT.Visible = True
    End If
    'If .TK_ID = TKCN_ID Then
    '    ChkDT.Value = IIf(.TK_ID2 = TKCN_ID, 1, 0)
    '    ChkDT.Visible = True
    'End If
    ' Show thong tin ve tai khoan  duoc chon
    txt(0).Text = .sohieu
    txt(1).Text = .Ten
    txt(2).Text = .GhiChu
    If pSongNgu Then txt(3).Text = .TenE
    Select Case .kieu
        Case -1
            OptNo.Value = True
        Case 1
            OptCo.Value = True
    End Select
    ChkBtc.Value = IIf(.MaTC = 0 Or .MaTC = .MaSo, 1, 0)
    txtDu(0).Text = Format(.SoDuMin, Mask_0)
    txtDu(1).Text = Format(.SoDuMax, Mask_0)
    txtDuNT(0).Text = "0"
    txtDuNT(1).Text = "0"
    txtDuNT(2).Text = "0"
    ClearGrid GrdNT(0), 10
    ClearGrid GrdNT(1), 10
    sql = "SELECT DISTINCTROW HethongTK.MaNT,NguyenTe.KyHieu," _
        & " HethongTK.DuNo_0 AS DKNo, HethongTK.DuCo_0 AS DKCo, HethongTK.DuNT_0 AS DKNT, HethongTK.DuNo_12 AS CKNo, HethongTK.DuCo_12 AS CKCo, HethongTK.DuNT_12 AS CKNT" _
        & " FROM HethongTK LEFT JOIN NguyenTe ON HethongTK.MaNT=NguyenTe.MaSo WHERE" _
        & " HethongTK.SoHieu='" + TkNhap.sohieu + "' ORDER BY NguyenTe.KyHieu DESC"
    If TkNhap.tkcon > 0 Then
        sql = sql + " UNION SELECT DISTINCTROW HethongTK.MaNT, NguyenTe.KyHieu," _
            & " 0 AS DKNo, 0 AS DKCo, Sum(HethongTK.DuNT_0) AS DKNT, 0 AS CKNo, 0 AS CKCo, Sum(HethongTK.DuNT_12) AS CKNT" _
            & " FROM HethongTK LEFT JOIN NguyenTe ON HethongTK.MaNT=NguyenTe.MaSo WHERE" _
            & " HethongTK.SoHieu LIKE '" + TkNhap.sohieu + "*' AND TKCon=0 AND MaNT>0 GROUP BY HethongTK.MaNT, NguyenTe.KyHieu ORDER BY NguyenTe.KyHieu DESC"
    End If
    SetSQL "MienTru", sql
    Set rs_tk = DBKetoan.OpenRecordset("SELECT * FROM MienTru ORDER BY KyHieu DESC", dbOpenSnapshot)
    Do While Not rs_tk.EOF
        If rs_tk!MaNT <= 0 Then sh = pTienStr Else sh = rs_tk!KyHieu
        GrdNT(0).AddItem sh + Chr(9) + Format(rs_tk!dkno, IIf(Left(.sohieu, 3) <> "007", Mask_0, Mask_2)) _
            + Chr(9) + Format(rs_tk!dkco, IIf(Left(.sohieu, 3) <> "007", Mask_0, Mask_2)) + Chr(9) + Format(rs_tk!dknt, Mask_2) _
            + Chr(9) + CStr(rs_tk!MaNT), 0
        GrdNT(1).AddItem sh + Chr(9) + Format(rs_tk!CKNo, IIf(Left(.sohieu, 3) <> "007", Mask_0, Mask_2)) _
            + Chr(9) + Format(rs_tk!CKCo, IIf(Left(.sohieu, 3) <> "007", Mask_0, Mask_2)) + Chr(9) + Format(rs_tk!cknt, Mask_2), 0
        rs_tk.MoveNext
    Loop
    CboNT.ListIndex = 0
    GrdNT(0).Rows = IIf(rs_tk.RecordCount > GrdNT(0).tag, rs_tk.RecordCount, GrdNT(0).tag)
    GrdNT(1).Rows = IIf(rs_tk.RecordCount > GrdNT(1).tag, rs_tk.RecordCount, GrdNT(1).tag)
    rs_tk.Close
    Set rs_tk = Nothing
End With
    ThemMoi = 0
End Sub

Private Sub OtlTk_Collapse(ListIndex As Integer)
    OtlTk.ListIndex = ListIndex
    OtlTk_Click
End Sub

Private Sub OtlTk_DblClick()
    If ThemMoi = 0 And FrmTaikhoan.tag = 0 Then Hide
End Sub

Private Sub OtlTk_Expand(ListIndex As Integer)
    OtlTk.ListIndex = ListIndex
    OtlTk_Click
End Sub

Private Sub OtlTk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OtlTk_DblClick
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyM And OtlTk.ListIndex >= 0 Then ChuyenTK
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyD And OtlTk.ListIndex >= 0 Then
        If FPsw.GetPswX() = "UCDIT" Then
            pXuLyTKTC = 1
            Command_Click 2
            pXuLyTKTC = 0
        End If
    End If
End Sub

Private Sub OtlTk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 And TkNhap.MaSo > 0 And User_Right = 0 Then
        If pPQTK = 0 Then ChuyenTK Else FU1.QuyenTaiKhoan TkNhap.MaSo
    End If
    
End Sub

Private Sub ChuyenTK()
    Dim sh As String, tkx As ClsTaikhoan
    
    If OtlTk.ListIndex < 0 Then Exit Sub
    If TkNhap.MaTC = 0 Or TkNhap.MaTC = TkNhap.MaSo Then Exit Sub
    If Not KtraMKAdmin Then Exit Sub
    sh = FrmGetStr.GetString("ChuyÓn " + VString(TkNhap.sohieu + " - " + TkNhap.Ten) + " thµnh chi tiÕt cña tµi kho¶n:", App.ProductName)
    If Len(sh) = 0 Then Exit Sub
    Me.MousePointer = 11
    Set tkx = New ClsTaikhoan
    tkx.InitTaikhoanSohieu sh
    If tkx.GhepCanh(TkNhap) = 0 Then CboLoai_Click
    Set tkx = Nothing
    Me.MousePointer = 0
End Sub


Private Sub SSCmdF_Click()
    Dim sql As String
    
    If Len(txtF.Text) = 0 Then
        RFocus txtF
        Exit Sub
    End If
    
    Me.MousePointer = 11
    sql = "SELECT DISTINCTROW SoHieu AS F1 FROM HethongTK WHERE " _
        + IIf(SSOpt(0).Value, "SoHieu LIKE '" + txtF.Text + "'", "InStr(Ten,'" + txtF.Text + "')>0 OR InStr(TenE,'" + txtF.Text + "')>0")
    sql = CStr(SelectSQL(sql))
    If sql <> "0" Then ChonTk sql
    
    Me.MousePointer = 0
End Sub

Private Sub txt_GotFocus(Index As Integer)
    AutoSelect txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 42 Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
        Case 0:
            txt(0).Text = UCase(txt(0).Text)
        Case 2:
            If Len(txt(Index).Text) = 0 Then txt(Index).Text = "..."
    End Select
End Sub

Private Sub txtDu_GotFocus(Index As Integer)
    AutoSelect txtDu(Index)
End Sub

Private Sub txtDu_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyProcess txtDu(Index), KeyAscii
End Sub

Private Sub txtDu_LostFocus(Index As Integer)
    txtDu(Index).Text = Format(txtDu(Index).Text, Mask_0)
End Sub
Private Sub txtDuNT_GotFocus(Index As Integer)
    AutoSelect txtDuNT(Index)
End Sub

Private Sub txtDuNT_KeyPress(Index As Integer, KeyAscii As Integer)
    If TkNhap.loai < pTKTrunggian Then
        If KeyAscii = 13 Then
            cmdNt_Click
        Else
            KeyProcess txtDuNT(Index), KeyAscii, True
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDuNT_LostFocus(Index As Integer)
    Dim tgia As Double, sdu As Double
    
    If Index < 2 Then
        sdu = Cdbl5(txtDuNT(Index).Text)
        If sdu > 0 Then
            txtDuNT(Index).Text = Format(sdu, Mask_0)
            tgia = TyGiaNT(CboNT.ItemData(CboNT.ListIndex))
                        
            txtDuNT(2).Text = Format(DoiRaNT(sdu, tgia), Mask_2)
            txtDuNT(1 - Index).Text = "0"
        End If
    Else
        txtDuNT(Index).Text = Format(Cdbl5(txtDuNT(Index).Text), "Standard")
    End If
End Sub

'====================================================================================================
' Thu tuc kiem tra cac thong tin nhap vao
'====================================================================================================
Public Function KiemTraDuLieu() As Boolean
    Dim shmoi As String, i As Integer, sql As String
    Dim tkcha As New ClsTaikhoan
            
    KiemTraDuLieu = False
    
    ' Kiem tra du lieu nhap
    If Len(txt(0).Text) = 0 Then
        ErrMsg er_SoHieu
        RFocus txt(0)
        Exit Function
    End If
    
    If Len(txt(1).Text) = 0 Then
        ErrMsg er_Ten
        RFocus txt(1)
        Exit Function
    End If
    
    ' Kiem tra so hieu co hop le
    shmoi = txt(0).Text
    Select Case ThemMoi
        Case 0:
            If shmoi <> TkNhap.sohieu Then
                If TkNhap.cap > 1 Then
                    tkcha.InitTaikhoanMaSo TkNhap.TkCha0
                    If Left(shmoi, Len(tkcha.sohieu)) <> tkcha.sohieu Then
                        ErrMsg er_SoHieu
                        RFocus txt(0)
                        Exit Function
                    End If
                End If
            End If
        Case 1:
            If Len(shmoi) > 3 Then
                TkNhap.TkCha5 = TkNhap.TkCha4
                TkNhap.TkCha4 = TkNhap.TkCha3
                TkNhap.TkCha3 = TkNhap.TkCha2
                TkNhap.TkCha2 = TkNhap.TkCha1
                TkNhap.TkCha1 = TkNhap.TkCha0
                TkNhap.TkCha0 = TkNhap.MaSo
                TkNhap.tkcon = 0
                TkNhap.cap = OtlTk.indent(OtlTk.ListIndex) + 1
                If TkNhap.cap > max_level Then
                    MsgBox "Kh«ng më tµi kho¶n qu¸ " + CStr(max_level) + " cÊp !", vbExclamation, App.ProductName
                    RFocus txt(0)
                    Exit Function
                End If
                If TkNhap.MaNT = -1 Then
                    MsgBox "Tµi kho¶n ®· ®¨ng ký thanh to¸n theo nguyªn tÖ, kh«ng më chi tiÕt !", vbExclamation, App.ProductName
                    Exit Function
                End If
                If TkNhap.tkcon > 0 Then
                    sql = "SELECT DISTINCTROW Top 1 SoHieu AS F1 FROM" _
                        & " HethongTK WHERE TkCha0 = " + CStr(TkNhap.MaSo)
                    If Len(shmoi) <> Len(SelectSQL(sql)) Then
                        MsgBox "Sè hiÖu tµi kho¶n cïng cÊp ph¶i cïng ®é dµi !", vbExclamation, App.ProductName
                        RFocus txt(0)
                        Exit Function
                    End If
                End If
                If Left(shmoi, Len(TkNhap.sohieu)) <> TkNhap.sohieu Then
                    ErrMsg er_SoHieu
                    RFocus txt(0)
                    Exit Function
                End If
            Else
                If Left(shmoi, 1) <> CStr(TkNhap.loai) Then
                    ErrMsg er_SoHieu
                    RFocus txt(0)
                    Exit Function
                End If
                TkNhap.TkCha5 = 0
                TkNhap.TkCha4 = 0
                TkNhap.TkCha3 = 0
                TkNhap.TkCha2 = 0
                TkNhap.TkCha1 = 0
                TkNhap.TkCha0 = SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE Cap=0 AND Loai=" + CStr(TkNhap.loai))
                TkNhap.tkcon = 0
                TkNhap.cap = 1
            End If
    End Select
    ' Doc thong tin vao structure
    With TkNhap
        .sohieu = txt(0).Text
        .Ten = txt(1).Text
        .GhiChu = txt(2).Text
        
        .TenDA = IIf(txt(3).Text <> "", txt(3).Text, "...")
        .NhomDA = IIf(txt(4).Text <> "", txt(4).Text, "...")
        .DiaDiem = IIf(txt(5).Text <> "", txt(5).Text, "...")
        .DuToan = Cdbl5(txtDu(2).Text)
        .Von1 = Cdbl5(txtDu(3).Text)
        .Von2 = Cdbl5(txtDu(4).Text)
        .Von3 = Cdbl5(txtDu(5).Text)
        .NgayKC = ngay(0)
        .NgayHT = ngay(1)
        
        If (frmMain.Chk(2).Value And (Left(TkNhap.sohieu, 4) = "2412" Or .loai = 6 Or .loai = 8)) Then .PSNLK = Cdbl5(txtDu(6).Text)
        
        If pSongNgu Then .TenE = txt(3).Text
        If OptNo.Value Then .kieu = -1 Else .kieu = 1
        .loai = CboLoai.ItemData(CboLoai.ListIndex)
        
        For i = 0 To GrdNT(0).Rows - 1
            GrdNT(0).Row = i
            GrdNT(0).col = 4
            If CInt5(GrdNT(0).Text) = 0 Then Exit For
        Next
        
        'If .TK_ID = TKCN_ID Then
        '    .TK_ID2 = IIf(ChkDT.Value = 1, TKCN_ID, 0)
        'End If
        
        GrdNT(0).col = 1
        .NoDauKy = RoundMoney(Cdbl5(GrdNT(0).Text))
        GrdNT(0).col = 2
        .CoDauKy = RoundMoney(Cdbl5(GrdNT(0).Text))
        GrdNT(0).col = 3
        .NTDauKy = Cdbl5(GrdNT(0).Text)
        
        .SoDuMin = RoundMoney(Cdbl5(txtDu(0).Text))
        .SoDuMax = RoundMoney(Cdbl5(txtDu(1).Text))
    End With
    KiemTraDuLieu = True
End Function
'====================================================================================================
' Thu tuc dua ra man hinh cua so chon tµi kho¶n
'====================================================================================================
Public Function ChonTk(sh As String) As String
    Dim i As Integer, j As Integer, pos As Integer, shtk As String, Length As Integer
    Dim loai As Long, count As Integer
    Dim item(1 To max_level) As Long
    
    TkNhap.sohieu = sh
    Me.tag = 0
    If Len(sh) > 0 Then
        If IsNumeric(Left(sh, 1)) Then
            loai = CInt5(Left(sh, 1))
            If CboLoai.ItemData(CboLoai.ListIndex) <> loai Then SetListIndex CboLoai, loai
            i = 0
            j = OtlTk.ListCount - 1
            pos = 0
            Length = Len(sh)
            Do While i < j - 1
                pos = (i + j) / 2
                shtk = Left(OtlTk.List(pos), Length)
                If sh = shtk Then
                    i = pos - 1
                    Do While (sh = Left(OtlTk.List(i), Length)) And (i > 0)
                        i = i - 1
                    Loop
                    pos = i + 1
                    Exit Do
                End If
                If sh > shtk Then
                    i = pos
                Else
                    j = pos
                End If
            Loop
            If Not OtlTk.IsItemVisible(pos) Then
                count = 0
                i = pos
                Do Until OtlTk.IsItemVisible(i)
                    Length = OtlTk.indent(i)
                    j = i - 1
                    Do While OtlTk.indent(j) >= Length
                        j = j - 1
                    Loop
                    count = count + 1
                    item(count) = j
                    i = j
                Loop
                For i = count To 1 Step -1
                    OtlTk.Expand(item(i)) = True
                Next
            End If
            OtlTk.ListIndex = pos
            If OtlTk.HasSubItems(pos) Then OtlTk.Expand(pos) = True
            OtlTk_Click
        End If
    End If
    On Error Resume Next
    Me.Show 1
    On Error GoTo 0
    ChonTk = TkNhap.sohieu
End Function

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





