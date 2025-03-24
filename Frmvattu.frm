VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmVattu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HÖ thèng danh ®iÓm vËt t­, hµng ho¸"
   ClientHeight    =   8130
   ClientLeft      =   345
   ClientTop       =   375
   ClientWidth     =   10830
   ClipControls    =   0   'False
   Icon            =   "Frmvattu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Inventory Items"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8130
   ScaleWidth      =   10830
   Tag             =   "0"
   Begin VB.CommandButton Command 
      Caption         =   "X.Lô"
      Height          =   375
      Index           =   5
      Left            =   5280
      TabIndex        =   90
      Top             =   6600
      Width           =   610
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1815
      Left            =   5310
      TabIndex        =   79
      Top             =   5640
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3201
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   5
      ShowFocusRect   =   0   'False
      OLEDropMode     =   1
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Frmvattu.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label(33)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(34)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(35)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label(36)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "GrdNT(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txthandung"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtsolo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtsoluong"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton Command2 
         Caption         =   "Xãa"
         Height          =   375
         Left            =   0
         TabIndex        =   84
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ghi"
         Height          =   375
         Left            =   600
         TabIndex        =   83
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtsoluong 
         Height          =   375
         Left            =   1200
         TabIndex        =   80
         Text            =   "0"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtsolo 
         Height          =   375
         Left            =   2280
         TabIndex        =   81
         Text            =   "..."
         Top             =   1440
         Width           =   1150
      End
      Begin MSMask.MaskEdBox txthandung 
         Height          =   375
         Left            =   3435
         TabIndex        =   82
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "99/99/99"
         PromptChar      =   "_"
      End
      Begin MSGrid.Grid GrdNT 
         Height          =   1215
         Index           =   4
         Left            =   0
         TabIndex        =   88
         Tag             =   "10"
         Top             =   240
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   2143
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
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tån ®Çu kú"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   36
         Left            =   0
         TabIndex        =   89
         Tag             =   "U. Price"
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sè l­îng"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   35
         Left            =   1200
         TabIndex        =   87
         Tag             =   "U. Price"
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sè l«"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   34
         Left            =   2160
         TabIndex        =   86
         Tag             =   "U. Price"
         Top             =   0
         Width           =   1270
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H¹n dïng"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   33
         Left            =   3370
         TabIndex        =   85
         Tag             =   "U. Price"
         Top             =   0
         Width           =   1600
      End
   End
   Begin VB.PictureBox Panel 
      BackColor       =   &H00FFFFFF&
      Height          =   7335
      Index           =   0
      Left            =   5040
      ScaleHeight     =   7275
      ScaleWidth      =   5595
      TabIndex        =   25
      Tag             =   "0"
      Top             =   120
      Width           =   5655
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   1665
         TabIndex        =   78
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox TxtVT 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "0"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox TxtVT 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "0"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox TxtVT 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   3600
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "Frmvattu.frx":57FE
         Top             =   6600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton CmdChitiet 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5040
         TabIndex        =   16
         ToolTipText     =   "Ghi ph¸t sinh"
         Top             =   6600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "C¸c ®¬n vÞ tÝnh chuyÓn ®æi vµ tû lÖ quy ®æi so víi ®.v.t c¬ b¶n"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   0
         TabIndex        =   12
         Tag             =   "Conversion Units"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtVT 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   13
         Top             =   6600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtVT 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   2400
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "Frmvattu.frx":5800
         Top             =   6600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtVT 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   2880
         MaxLength       =   20
         TabIndex        =   8
         Text            =   "0"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxtVT 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "0"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxtVT 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   360
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "0"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxtVT 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "0"
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox TxtVT 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1450
         Width           =   4095
      End
      Begin VB.TextBox TxtVT 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1560
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Frmvattu.frx":5802
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtTon 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3960
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "Frmvattu.frx":5804
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtTon 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1560
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "Frmvattu.frx":5808
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox TxtVT 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TxtVT 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox TxtVT 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin MSGrid.Grid GrdNT 
         Height          =   2895
         Index           =   2
         Left            =   240
         TabIndex        =   67
         Tag             =   "10"
         Top             =   4200
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   5106
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
      Begin MSGrid.Grid GrdNT 
         Height          =   975
         Index           =   3
         Left            =   1320
         TabIndex        =   74
         Tag             =   "10"
         Top             =   5160
         Visible         =   0   'False
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   1720
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
         Enabled         =   0   'False
         Rows            =   10
         Cols            =   4
         FixedRows       =   0
         ScrollBars      =   2
         HighLight       =   0   'False
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tû lÖ thuÕ NK (%)"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   2520
         TabIndex        =   77
         Tag             =   "VAT Rate (%)"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Line Line 
         Index           =   7
         X1              =   4080
         X2              =   4920
         Y1              =   3045
         Y2              =   3045
      End
      Begin VB.Line Line 
         Index           =   4
         X1              =   4560
         X2              =   4920
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "CK (%)"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   4440
         TabIndex        =   76
         Tag             =   "VAT Rate (%)"
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "§¬n gi¸"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   27
         Left            =   3120
         TabIndex        =   75
         Tag             =   "U. Price"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "§¬n vÞ tÝnh"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   26
         Left            =   1320
         TabIndex        =   73
         Tag             =   "Unit"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tû lÖ QD"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   2400
         TabIndex        =   72
         Tag             =   "Rate"
         Top             =   4920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Thµnh tiÒn"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   30
         Left            =   3640
         TabIndex        =   71
         Tag             =   "Amount"
         Top             =   3960
         Width           =   1560
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "§¬n gi¸"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   29
         Left            =   2400
         TabIndex        =   70
         Tag             =   "Unit Price"
         Top             =   3960
         Width           =   1305
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sè l­îng"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   1430
         TabIndex        =   69
         Tag             =   "Quantity"
         Top             =   3960
         Width           =   985
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kho"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   68
         Tag             =   "Store"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Line Line 
         Index           =   13
         X1              =   2880
         X2              =   3840
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Line Line 
         Index           =   12
         X1              =   1560
         X2              =   2520
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Line Line 
         Index           =   10
         X1              =   360
         X2              =   1320
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C¸c ®¬n gi¸ b¸n"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   65
         Tag             =   "Sale Price 1"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tû lÖ thuÕ VAT (%)"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   2880
         TabIndex        =   64
         Tag             =   "VAT Rate (%)"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Line Line 
         Index           =   9
         X1              =   4080
         X2              =   4320
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Line Line 
         Index           =   8
         X1              =   1200
         X2              =   4440
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ghi chó"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   40
         Tag             =   "Notes"
         Top             =   1560
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line 
         Index           =   3
         X1              =   1560
         X2              =   2520
         Y1              =   3045
         Y2              =   3045
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gi¸ HT"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   39
         Tag             =   "P. Price"
         Top             =   2880
         Width           =   615
      End
      Begin VB.Line Line 
         Index           =   6
         X1              =   4080
         X2              =   4920
         Y1              =   3405
         Y2              =   3405
      End
      Begin VB.Line Line 
         Index           =   5
         X1              =   1560
         X2              =   2520
         Y1              =   3405
         Y2              =   3405
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tèi ®a"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   2880
         TabIndex        =   31
         Tag             =   "Max"
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tån kho tèi thiÓu"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Tag             =   "Minimum Stock"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Line Line 
         Index           =   2
         X1              =   1200
         X2              =   2640
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tån kho hiÖn thêi"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   29
         Tag             =   "Current Stock"
         Top             =   3650
         Width           =   1455
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "§¬n vÞ tÝnh"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Tag             =   "Unit"
         Top             =   1200
         Width           =   975
      End
      Begin VB.Line Line 
         Index           =   1
         X1              =   1200
         X2              =   4920
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line 
         Index           =   0
         X1              =   1200
         X2              =   3120
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tªn"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Tag             =   "Desc."
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "M· sè"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Tag             =   "Code"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox CboThang 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Frmvattu.frx":580C
      Left            =   7320
      List            =   "Frmvattu.frx":580E
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton CmdChitiet 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   9240
      TabIndex        =   50
      ToolTipText     =   "Ghi ph¸t sinh"
      Top             =   5400
      Width           =   255
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   8400
      MaxLength       =   20
      TabIndex        =   49
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   6360
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   5160
      MaxLength       =   20
      TabIndex        =   47
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   6840
      MaxLength       =   12
      MultiLine       =   -1  'True
      TabIndex        =   51
      Text            =   "Frmvattu.frx":5810
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H8000000E&
      Caption         =   "§Þnh &møc"
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
      TabIndex        =   21
      Tag             =   "&Norm"
      Top             =   7680
      Width           =   1095
   End
   Begin MSGrid.Grid GrdNT 
      Height          =   1575
      Index           =   0
      Left            =   5160
      TabIndex        =   52
      Tag             =   "10"
      Top             =   1440
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   2778
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
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   5160
      MaxLength       =   20
      TabIndex        =   42
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   8400
      MaxLength       =   20
      TabIndex        =   45
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton CmdChitiet 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   9240
      TabIndex        =   46
      ToolTipText     =   "Ghi ph¸t sinh"
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   9600
      Picture         =   "Frmvattu.frx":5814
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "&Return"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7200
      Picture         =   "Frmvattu.frx":6C36
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "&Delete"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6000
      Picture         =   "Frmvattu.frx":8118
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "&Save"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4800
      Picture         =   "Frmvattu.frx":9546
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "&Add"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton SSCmdF 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   38
      Top             =   7800
      Width           =   255
   End
   Begin VB.OptionButton SSOpt 
      BackColor       =   &H8000000E&
      Caption         =   "Tªn VT"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   37
      Tag             =   "Desc."
      Top             =   7800
      Width           =   855
   End
   Begin VB.OptionButton SSOpt 
      BackColor       =   &H8000000E&
      Caption         =   "Sè hiÖu"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Tag             =   "Code"
      Top             =   7800
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txtF 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   35
      Top             =   7800
      Width           =   1335
   End
   Begin VB.ComboBox CboLoai 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin MSGrid.Grid GrdNT 
      Height          =   1335
      Index           =   1
      Left            =   5160
      TabIndex        =   62
      Tag             =   "10"
      Top             =   4080
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   2355
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
      Height          =   7020
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   4815
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "§Þnh møc ¸p dông tõ th¸ng"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   5160
      TabIndex        =   66
      Tag             =   "Norm applied from month"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè hiÖu"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   5160
      TabIndex        =   56
      Tag             =   "Code"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tªn Nguyªn vËt liÖu"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   6000
      TabIndex        =   55
      Tag             =   "Description"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tªn TSC§"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   6360
      TabIndex        =   60
      Tag             =   "Description"
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè hiÖu"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   5160
      TabIndex        =   59
      Tag             =   "Code"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HÖ sè ph©n bæ chi phÝ khÊu hao TSC§"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   5400
      TabIndex        =   63
      Tag             =   "Rate of Depriciation"
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HÖ sè"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   8400
      TabIndex        =   61
      Tag             =   "Rate"
      Top             =   3840
      Width           =   855
   End
   Begin VB.Line Line 
      Index           =   11
      X1              =   6840
      X2              =   7920
      Y1              =   6165
      Y2              =   6165
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "§Þnh møc Nh©n c«ng"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   5160
      TabIndex        =   58
      Tag             =   "Norm of Labour"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "§Þnh møc Nguyªn vËt liÖu"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   5160
      TabIndex        =   57
      Tag             =   "Norm of Material"
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§.v.t"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7680
      TabIndex        =   54
      Tag             =   "Unit"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè l­îng"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   8400
      TabIndex        =   53
      Tag             =   "Quantity"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   7310
      Index           =   5
      Left            =   5040
      TabIndex        =   34
      Top             =   165
      Width           =   5655
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6255
      Index           =   4
      Left            =   5040
      TabIndex        =   33
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6900
      Index           =   9
      Left            =   120
      TabIndex        =   32
      Top             =   540
      Width           =   4815
   End
End
Attribute VB_Name = "FrmVattu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer          ' =1 neu them moi, -1 neu sua cu
Dim vattu As New ClsVattu      ' vat tu duoc tham chieu
Dim vt As New ClsVattu
Dim ts As New clsTaiSan
Dim doiloai As Integer               ' =1 neu co thay doi loai vat tu dang sua doi
Dim ChucNang As Integer
Dim MaDaTim As Long
Dim xT As Integer
Dim xSH As String
Dim so_dong_hien_tai As Integer
'======================================================================================
' Liet ke cac vat tu trong loai vat tu duoc chon
'======================================================================================
Private Sub CboLoai_Click()
    If ThemMoi <> -1 Then
        Me.MousePointer = 11
        Int_RecsetToCbo "SELECT MaSo As F2, SoHieu + Chr(9) + TenVattu As F1 FROM Vattu WHERE MaPhanLoai=" + CStr(CboLoai.ItemData(CboLoai.ListIndex)) + " ORDER BY SoHieu", LstVt
        ThemMoi = 0
        doiloai = 0
        If LstVt.ListIndex < 0 Then
            vattu.InitVattuMaSo 0
            ClearGrid GrdNT(0), GrdNT(0).tag
        End If
        Me.MousePointer = 0
    Else
        doiloai = 1
    End If
End Sub

Private Sub cboThang_Click()
    Dim thang1 As Integer, TM As Integer
    
    If pDinhmuc <> 0 Then
        thang1 = SelectSQL("SELECT TOP 1 Thang AS F1 FROM DinhMuc WHERE MaTP=" + CStr(vattu.MaSo) + " AND " + WThang("Thang", 0, CboThang.ItemData(CboThang.ListIndex)) + " ORDER BY " + SetMonthOrder("Thang") + " DESC")
        If thang1 > 0 And thang1 <> CboThang.ItemData(CboThang.ListIndex) And Me.Visible Then
            TM = IIf(MsgBox("Thªm ®Þnh møc míi ¸p dông tõ th¸ng " + CboThang.Text + " ?", vbInformation + vbYesNo, App.ProductName) = vbYes, 1, 0)
        End If
    End If
    LietKeDinhMuc TM
End Sub

Private Sub Chk_Click()
    GrdNT(3).Enabled = Chk.Value > 0
    txtVT(4).Enabled = Chk.Value > 0
    txtVT(5).Enabled = Chk.Value > 0
    txtVT(11).Enabled = Chk.Value > 0
    CmdChitiet(2).Enabled = Chk.Value > 0
    If ThemMoi = 1 And Chk.Value > 0 Then RFocus txtVT(4)
End Sub

Public Sub Command_Click(Index As Integer)
    Dim vt1 As New ClsVattu, i As Integer, dv As String, qd As Double, gb As Double
    
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
            txtVT(0).Text = SoHieuVTMoi(CboLoai.ItemData(CboLoai.ListIndex))
            txtVT(1).Text = ""
            txtVT(13).Text = ""
            
            txtTon(0).Text = "0"
            txtTon(1).Text = "0"
            ClearGrid GrdNT(2), GrdNT(2).tag
            
            Chk.Value = 0
            ClearGrid GrdNT(3), GrdNT(3).tag
            
            RFocus txtVT(0)
            ThemMoi = 1
        Case 1:
            Select Case ThemMoi
                Case 1:
                    If Not KiemTraSoLieu Then GoTo XongVT
                    If vattu.GhiVattu = 0 Then
                        If Chk.Value = 1 Then
                            With GrdNT(3)
                                For i = 0 To .Rows - 1
                                    .Row = i
                                    .col = 0
                                    dv = .Text
                                    If Len(dv) = 0 Then Exit For
                                    .col = 1
                                    qd = Cdbl5(.Text)
                                    .col = 2
                                    gb = Cdbl5(.Text)
                                    ExecuteSQL5 "INSERT INTO DVTVattu (MaSo,MaVattu,DonVi,TyleQD,GiaBan) VALUES (" + CStr(Lng_MaxValue("MaSo", "DVTVattu") + 1) + "," + CStr(vattu.MaSo) + ",'" + dv + "'," + DoiDau(qd) + "," + DoiDau(gb) + ")"
                                Next
                            End With
                        End If
                        
                        LstVt.AddItem vattu.sohieu + Chr(9) + vattu.TenVattu
                        LstVt.ItemData(LstVt.NewIndex) = vattu.MaSo
                        LstVt.ListIndex = LstVt.NewIndex
                    Else
                        ErrMsg er_PhanLoai
                        vt1.InitVattuSohieu txtVT(0).Text
                        If vt1.MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex) Then
                            SetListIndex LstVt, vt1.MaSo
                        End If
                    End If
                    ThemMoi = 0
                Case 0:
                    If LstVt.ListIndex < 0 Then GoTo XongVT
                    If Not KiemTraSoLieu Then GoTo XongVT
'                    vt.InitVattuMaSo vattu.MaSo
                    
                    If vattu.SuaVT = 0 Then
                        If doiloai = 1 Then
                            CboLoai_Click
                            doiloai = 0
                        Else
                            LstVt.List(LstVt.ListIndex) = vattu.sohieu + Chr(9) + vattu.TenVattu
                        End If
                    Else
                        vt1.InitVattuSohieu txtVT(0).Text
                        ErrMsg er_SoHieu
                        If vt1.MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex) Then SetListIndex LstVt, vt1.MaSo
                    End If
                    ThemMoi = 0
            End Select
            RFocus LstVt
        Case 2:
            i = LstVt.ListIndex
            If i < 0 Then GoTo XongVT
            If vattu.XoaVT = 0 Then
                LstVt.RemoveItem i
                If LstVt.ListCount > 0 Then LstVt.ListIndex = i - 1
            Else
                ErrMsg er_CoPS
            End If
            RFocus LstVt
        Case 3:
        LO_XXXX = ""
            Hide
          Case 5:
          GrdNT(4).col = 2
          LO_XXXX = GrdNT(4).Text
                GrdNT(4).col = 1
    SL_XXXX = GrdNT(4).Text
            Hide
        Case 4:
            ChucNang = 1 - ChucNang
            Panel(0).Visible = (ChucNang = 0)
            For i = 0 To 2
                Command(i).Enabled = (ChucNang = 0)
            Next
            If LstVt.ListIndex >= 0 Then
                LstVt_Click
            Else
                vattu.InitVattuMaSo 0
                ClearGrid GrdNT(0), GrdNT(0).tag
                ClearGrid GrdNT(1), GrdNT(1).tag
            End If
    End Select
XongVT:
    Set vt1 = Nothing
    Me.MousePointer = 0
End Sub

Private Sub Command1_Click()
Dim st As String
Dim solo As String, handung As String, SoLuong As String
If Len(txtsoluong.Text) > 0 Then
ExecuteSQL5 "insert into SoLoThuoc(mavattu,Solo,handung,soluong) values(" + CStr(vattu.MaSo) + ",'" + txtsolo.Text + "','" + txthandung.Text + "'," + txtsoluong.Text + ")"
End If
ShowChitiet vattu
End Sub

Private Sub Command2_Click()
Dim st As String
Dim solo As String, handung As String, SoLuong As String
With GrdNT(4)
'.Row = so_dong_hien_tai
.col = 1
SoLuong = .Text
.col = 2
solo = .Text
.col = 3
handung = .Text

End With
st = "delete from SoLoThuoc where mavattu  = " + CStr(vattu.MaSo) + " and solo = '" + CStr(solo) + "' " 'and handung = #" + handung + "#"
If (Len(Trim(handung)) > 0) Then
If Len(handung) < 2 Then
st = "delete from SoLoThuoc where mavattu  = " + CStr(vattu.MaSo) + " and solo = '" + CStr(solo) + "' and soluong = " + SoLuong + ""
End If
ExecuteSQL5 st
ShowChitiet vattu
End If
End Sub

Public Sub Form_Activate()
    If Me.tag < 0 Then
        SetListIndex CboLoai, -Me.tag
        Me.tag = 0
    End If
    
    If ThemMoi = 0 And Me.tag = 1 And ChucNang = 0 Then RFocus LstVt
    
    If xT = 1 Then
        If xSH <> "" Then SetListIndex CboLoai, LayMaPhanLoai(xSH, "Vattu")
        Command_Click 0
        txtVT(0).Text = xSH
    End If
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



' ColumnSetUp GrdDanhSachLo, 0, 820, 0
  '  ColumnSetUp GrdDanhSachLo, 1, 1660, 0
  '  ColumnSetUp GrdDanhSachLo, 2, 700, 0
    
    ColumnSetUp GrdNT(0), 0, 820, 0
    ColumnSetUp GrdNT(0), 1, 1660, 0
    ColumnSetUp GrdNT(0), 2, 700, 0
    ColumnSetUp GrdNT(0), 3, 820, 1
    
    ColumnSetUp GrdNT(1), 0, 1180, 0
    ColumnSetUp GrdNT(1), 1, 2020, 0
    ColumnSetUp GrdNT(1), 2, 820, 1
    
'    ColumnSetUp GrdNT(2), 0, 1180, 0
'    ColumnSetUp GrdNT(2), 1, 940, 1
'    ColumnSetUp GrdNT(2), 2, 940, 1
'    ColumnSetUp GrdNT(2), 3, 940, 1
    
    ColumnSetUp GrdNT(2), 0, 1180, 0
    ColumnSetUp GrdNT(2), 1, 940, 1
    ColumnSetUp GrdNT(2), 2, 940 + 300, 1
    ColumnSetUp GrdNT(2), 3, 940 + 445 + 150, 1
    
    
    ColumnSetUp GrdNT(3), 0, 1060, 2
    ColumnSetUp GrdNT(3), 1, 700, 1
    ColumnSetUp GrdNT(3), 2, 1180 + 0, 1
    ColumnSetUp GrdNT(3), 3, 1, 0
    
    ColumnSetUp GrdNT(4), 0, 1180, 0
    ColumnSetUp GrdNT(4), 1, 940, 1
    ColumnSetUp GrdNT(4), 2, 940 + 300, 1
    ColumnSetUp GrdNT(4), 3, 940 + 445 + 150, 1
    
    
    Pic.Visible = (pBarCode > 0)
    txtVT(3).Enabled = pGiaHT > 0
    ThemMoi = 0
    doiloai = 0
    Caption = Caption + " - " + CStr(pNamTC)
    Int_RecsetToCbo "SELECT DISTINCTROW PhanLoaiVattu.MaSo As F2, PhanLoaiVattu.SoHieu + ' - '  + PhanLoaiVattu.TenPhanLoai As F1 FROM PhanLoaiVattu WHERE PLCon=0 ORDER BY PhanLoaiVattu.SoHieu", CboLoai
        
    Label(28).Visible = (pDinhmuc <> 0)
    CboThang.Visible = (pDinhmuc <> 0)
    AddMonthToCbo CboThang
    If pDinhmuc = 0 Then CboThang.ListIndex = 0
    
    Label(31).Visible = pChietKhau > 0
    txtVT(12).Visible = pChietKhau > 0
        
    'Lines(7).Visible = (pNhapKhau > 0)
    Label(32).Visible = (pNhapKhau > 0)
    txtVT(13).Visible = (pNhapKhau > 0)
        
    FCenter Me

    SetFont Me
    If SelectSQL("SELECT banthuoc as f1 from license ") = 0 Then
    GrdNT(4).Visible = False
    SSTab1.Visible = False
    Else
    GrdNT(4).Visible = True
    SSTab1.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set vattu = Nothing
    Set vt = Nothing
    Set ts = Nothing
End Sub

Private Sub GrdNT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 And Button = 2 Then GrdNt_KeyPress 0, vbKeyH
End Sub

'======================================================================================
' Khoi tao vat tu duoc chon
'======================================================================================
Private Sub LstVt_Click()
    vattu.InitVattuMaSo LstVt.ItemData(LstVt.ListIndex)
    Select Case ChucNang
        Case 0:
            ShowChitiet vattu
        Case 1:
            cboThang_Click
    End Select
End Sub
'======================================================================================
' Thu tuc hien thong tin chi tiet
'======================================================================================
Private Sub ShowChitiet(vattu As ClsVattu)
    Dim rs As Recordset, dgia As Double, st As String
 
    txtVT(0).Text = vattu.sohieu
    txtVT(1).Text = vattu.TenVattu
    txtVT(2).Text = vattu.DonVi
    txtVT(3).Text = Format(vattu.GiaHT, Mask_0)
    'Chk.Value = vattu.Dvt2
    'txtVT(4).Text = vattu.DonVi2
    'txtVT(5).Text = Format(vattu.TyleQD, Mask_2)
    txtVT(6).Text = vattu.GhiChu
    txtVT(7).Text = CStr(vattu.vat)
    txtVT(8).Text = Format(vattu.GiaBan1, Mask_2)
    txtVT(9).Text = Format(vattu.GiaBan2, Mask_2)
    txtVT(10).Text = Format(vattu.GiaBan3, Mask_2)
    txtVT(12).Text = Format(vattu.CK, Mask_2)
    txtVT(13).Text = Format(vattu.ThueNK, Mask_2)
    
    txtTon(0).Text = Format(vattu.TonMin, Mask_2)
    txtTon(1).Text = Format(vattu.TonMax, Mask_2)
    
    If pBarCode > 0 Then TxtVT_Change 0
    
    ClearGrid GrdNT(3), GrdNT(3).tag
   
    Set rs = DBKetoan.OpenRecordset("SELECT * FROM DVTVattu WHERE MaVattu=" + CStr(vattu.MaSo) + " ORDER BY DonVi DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        GrdNT(3).AddItem rs!DonVi + Chr(9) + Format(rs!TyLeQD, Mask_2) + Chr(9) + Format(rs!GiaBan, Mask_2) + Chr(9) + CStr(rs!MaSo), 0
        rs.MoveNext
    Loop
    Chk.Value = IIf(rs.RecordCount > 0, 1, 0)
    'Chk_Click
    rs.Close
    
    ClearGrid GrdNT(2), GrdNT(2).tag
    st = CStr(CThangDB(ThangCuoiNamTC))
    Set rs = DBKetoan.OpenRecordset("SELECT TenKho,Sum(Luong_" + st + ") AS Luong, Sum(Tien_" + st + ") AS Tien FROM TonKho INNER JOIN KhoHang ON TonKho.MaSoKho=KhoHang.MaSo WHERE MaVattu=" + CStr(vattu.MaSo) + " GROUP BY TenKho HAVING Sum(Luong_" + st + ")<>0 OR Sum(Tien_" + st + ")<>0 ORDER BY TenKho DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        If rs!luong <> 0 Then dgia = rs!tien / rs!luong Else dgia = 0
        GrdNT(2).AddItem rs!tenkho + Chr(9) + Format(rs!luong, Mask_2) + Chr(9) + Format(dgia, Mask_2) + Chr(9) + Format(rs!tien, Mask_0), 0
        rs.MoveNext
    Loop
    rs.Close
    
    
     ClearGrid GrdNT(4), GrdNT(4).tag
    Set rs = DBKetoan.OpenRecordset("select * from DanhSachVatTu where mavattu = " + CStr(vattu.MaSo) + "", dbOpenSnapshot)
    Do While Not rs.EOF
        GrdNT(4).AddItem "" + Chr(9) + CStr(rs!conlai) + Chr(9) + rs!solo + Chr(9) + rs!handung, 0
        rs.MoveNext
    Loop
    rs.Close


    Set rs = Nothing
End Sub
'======================================================================================
' Thu tuc chon vat tu
' sh: so hieu vat tu can chon
' Tra ve so hieu vat tu duoc chon
'======================================================================================
Public Function ChonVattu(sh As String, Optional c As Integer = 0) As String
    Dim mpl As Long, shtk As String
    Dim j As Integer, i As Integer, pos As Integer, Length As Integer
    
    If Len(sh) > 0 Then
        shtk = "SELECT DISTINCTROW TOP 1 Vattu.MaPhanLoai AS F1 FROM Vattu WHERE SoHieu LIKE '" + sh + "*' ORDER BY SoHieu"
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
                If UCase(sh) = UCase(shtk) Then
                    i = pos - 1
                    Do While (UCase(sh) = UCase(Left(LstVt.List(i), Length))) And (i > 0)
                        i = i - 1
                    Loop
                    pos = i + 1
                    Exit Do
                End If
                If UCase(sh) > UCase(shtk) Then
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
    FCenter Me
    
    If c <> 0 Then
        Me.Top = Me.Top + 240
        Me.Left = Me.Left + 580
        Command(4).Enabled = False
    End If
    
    On Error Resume Next
    Me.Show 1
    On Error GoTo 0
    If vattu.MaSo > 0 Then
        ChonVattu = vattu.sohieu
    Else
        ChonVattu = ""
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
        MsgBox "ThiÕu ®¬n vÞ tÝnh vËt t­!", vbExclamation, App.ProductName
        RFocus txtVT(2)
        Exit Function
    End If

With vattu
    If ThemMoi = 1 Then .MaSo = 0
    .MaPhanLoai = CboLoai.ItemData(CboLoai.ListIndex)
    .sohieu = txtVT(0).Text
    .TenVattu = txtVT(1).Text
    .DonVi = txtVT(2).Text
    .GiaHT = Cdbl5(txtVT(3).Text)
    .TonMin = Cdbl5(txtTon(0).Text)
    .TonMax = Cdbl5(txtTon(1).Text)
    .TyLeQD = Cdbl5(txtVT(5).Text)
    .vat = CInt5(txtVT(7).Text)
    .GiaBan1 = Cdbl5(txtVT(8).Text)
    .GiaBan2 = Cdbl5(txtVT(9).Text)
    .GiaBan3 = Cdbl5(txtVT(10).Text)
    .CK = Cdbl5(txtVT(12).Text)
    .Dvt2 = Chk.Value
    .DonVi2 = IIf(Len(txtVT(4).Text) > 0, txtVT(4).Text, "...")
    .GhiChu = IIf(Len(txtVT(6).Text) > 0, txtVT(6).Text, "...")
    If pNhapKhau > 0 Then .ThueNK = Cdbl5(txtVT(13).Text)
End With
    KiemTraSoLieu = True
End Function

Private Sub LstVt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sh As String, m As Long, Ten As String
    Dim luong As Double, tien As Double, flag As Integer, rs As Recordset
    
    If Button = 2 And LstVt.ListIndex >= 0 And ThemMoi = 0 Then
        If OutCost = 0 And SelectSQL("SELECT MaSo AS F1 FROM DVTVattu WHERE MaVattu=" + CStr(vattu.MaSo)) = 0 Then
            sh = FrmGetStr.GetString("ChuyÓn " + VString(vattu.sohieu + " - " + vattu.TenVattu) + " sang ph©n lo¹i kh¸c, hoÆc gép vµo vËt t­ cã sè hiÖu:", App.ProductName)
            flag = 1
        Else
            sh = FrmGetStr.GetString("ChuyÓn " + VString(vattu.sohieu + " - " + vattu.TenVattu) + " sang ph©n lo¹i cã sè hiÖu:", App.ProductName)
        End If
        If Len(sh) > 0 Then
            m = SelectSQL("SELECT MaSo AS F1 FROM PhanLoaiVattu WHERE PLCon=0 AND SoHieu='" + sh + "'")
            If m > 0 And m <> vattu.MaPhanLoai Then
                ExecuteSQL5 "UPDATE Vattu SET MaPhanLoai=" + CStr(m) + " WHERE MaSo = " + CStr(vattu.MaSo)
                CboLoai_Click
            End If
            If m > 0 Or flag = 0 Then GoTo KT
            m = SelectSQL("SELECT MaSo AS F1, TenVattu AS F2 FROM Vattu WHERE SoHieu='" + sh + "'", Ten)
            If m > 0 And m <> vattu.MaSo Then
                If MsgBox("B¹n ®· chøc ch¾n chuyÓn gép vËt t­ " + vattu.sohieu + " - " + vattu.TenVattu + " vµo vËt t­ " + sh + " - " + Ten + " ?", vbCritical + vbYesNo, App.ProductName) = vbYes Then
                    Me.MousePointer = 11
                    ExecuteSQL5 "UPDATE ChungTu SET MaVattu=" + CStr(m) + " WHERE MaSo=" + CStr(vattu.MaSo)
                    ExecuteSQL5 "UPDATE ChungTu2 SET MaVattu=" + CStr(m) + " WHERE MaSo=" + CStr(vattu.MaSo)
                    ExecuteSQL5 "UPDATE ChungTuP SET MaVattu=" + CStr(m) + " WHERE MaSo=" + CStr(vattu.MaSo)
                    
                    Set rs = DBKetoan.OpenRecordset("SELECT MaSo, MaSoKho, MaTaiKhoan, Luong_0, Tien_0 FROM TonKho WHERE MaVattu=" + CStr(vattu.MaSo) + " AND (Luong_0<>0 OR Tien_0<>0)", dbOpenSnapshot)
                    Do While Not rs.EOF
                        ExecuteSQL5 "UPDATE TonKho SET Luong_0=Luong_0+" + DoiDau(rs!Luong_0) + ",Tien_0=Tien_0+" + DoiDau(rs!Tien_0) + " WHERE MaVattu=" + CStr(m) + " AND MaSoKho=" + CStr(rs!MaSoKho) + " AND MaTaiKhoan=" + CStr(rs!MaTaiKhoan)
                        If DBKetoan.RecordsAffected = 0 Then ExecuteSQL5 "UPDATE TonKho SET MaVattu=" + CStr(m) + " WHERE MaVattu=" + CStr(vattu.MaSo) + " AND MaSoKho=" + CStr(rs!MaSoKho) + " AND MaTaiKhoan=" + CStr(rs!MaTaiKhoan)
                        rs.MoveNext
                    Loop
                    rs.Close
                    Set rs = Nothing
                    
                    ExecuteSQL5 "DELETE * FROM TonKho WHERE MaVattu=" + CStr(vattu.MaSo)
                    vattu.XoaVT
                    
                    KiemTraVatTu
                    CboLoai_Click
                    
                    Me.MousePointer = 0
                End If
            End If
        End If
    End If
KT:
End Sub

Private Sub Pic_DblClick()
    Dim i As Integer, st As String
    
    If pBarCode > 0 And vattu.MaSo > 0 Then
        st = FrmGetStr.GetString("Sè nh·n cÇn in", "In m· v¹ch")
        i = CInt5(st)
        If i > 0 Then PrintBarCode vattu, i
    End If
End Sub

Private Sub SSCmdF_Click()
    Dim sql As String
    
    If Len(txtF.Text) = 0 Then
        RFocus txtF
        Exit Sub
    End If
    
    Me.MousePointer = 11
    sql = "SELECT DISTINCTROW Top 1 SoHieu AS F1 FROM Vattu WHERE MaSo>" + CStr(MaDaTim) + IIf(SSOpt(0).Value, " AND SoHieu LIKE '" + txtF.Text + "'", " AND InStr(TenVattu,'" + txtF.Text + "')>0")
    sql = CStr(SelectSQL(sql))
    If sql <> "0" Then
        ChonVattu sql
        MaDaTim = vattu.MaSo
    Else
        MaDaTim = 0
    End If
    Me.MousePointer = 0
End Sub

Private Sub txtF_GotFocus()
    AutoSelect txtF
    MaDaTim = 0
End Sub

Private Sub txthandung_LostFocus()
txthandung.SelStart = 0
txthandung.SelLength = Len(txthandung.Text)
End Sub

Private Sub txtsolo_Click()
txtsolo.SelStart = 0
txtsolo.SelLength = Len(txtsolo.Text)
End Sub

Private Sub txtsolo_LostFocus()
txtsolo.SelStart = 0
txtsolo.SelLength = Len(txtsolo.Text)

End Sub

Private Sub txtsoluong_Click()
txtsoluong.SelStart = 0
txtsoluong.SelLength = Len(txthandung.Text)
End Sub

Private Sub txtsoluong_LostFocus()
txtsoluong.SelStart = 0
txtsoluong.SelLength = Len(txthandung.Text)
End Sub

Private Sub TxtVT_Change(Index As Integer)
    If Index = 0 And pBarCode = 1 Then
        Pic.Cls
        BarCode txtVT(0).Text, Pic, 8, 400, 0, 0
    End If
End Sub

Private Sub Txtvt_GotFocus(Index As Integer)
    AutoSelect txtVT(Index)
End Sub

Private Sub txtTon_GotFocus(Index As Integer)
    AutoSelect txtTon(Index)
End Sub

Private Sub txtTon_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyProcess txtTon(Index), KeyAscii
End Sub

Private Sub txtTon_LostFocus(Index As Integer)
    txtTon(Index).Text = Format(txtTon(Index).Text, Mask_2)
    
    If (Index = 2) And vattu.MaSo > 0 And ChucNang = 1 Then
        ExecuteSQL5 "UPDATE DinhMuc SET SoLuong=" + DoiDau(txtTon(Index).Text) + " WHERE MaNVL=0 AND MaTP=" + CStr(vattu.MaSo) + " AND Thang=" + CStr(CboThang.ItemData(CboThang.ListIndex))
        If DBKetoan.RecordsAffected = 0 Then ExecuteSQL5 "INSERT INTO DinhMuc (MaSo,MaTP,MaNVL,SoLuong,Thang) VALUES (" + CStr(Lng_MaxValue("MaSo", "DinhMuc") + 1) + "," + CStr(vattu.MaSo) + ",0," + DoiDau(txtTon(Index).Text) + "," + CStr(CboThang.ItemData(CboThang.ListIndex)) + ")"
    End If
End Sub

Private Sub TxtVT_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0:
            'If pBCode = 39 And pBarCode > 0 Then
            '    If KeyAscii < 48 Or (KeyAscii > 57 And KeyAscii < 65) Or (KeyAscii > 90 And KeyAscii < 97) Or KeyAscii > 122 Then KeyAscii = 0
            'Else
           If KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 42 Then KeyAscii = 0
            'End If
        Case 3, 5, 7, 8, 9, 10, 13: KeyProcess txtVT(Index), KeyAscii
    End Select
End Sub

Private Sub TxtVT_LostFocus(Index As Integer)
    Select Case Index
        Case 0:
            txtVT(0).Text = UCase(txtVT(0).Text)
        Case 3:
            txtVT(3).Text = Format(txtVT(3).Text, Mask_0)
        Case 5, 7, 8, 9, 10, 13:
            txtVT(Index).Text = Format(txtVT(Index).Text, Mask_2)
    End Select
End Sub

Private Sub CmdChitiet_Click(Index As Integer)
    Dim luong As Double, i As Integer, gia As Double
    
    If vattu.MaSo = 0 Then Exit Sub
    i = IIf(pDinhmuc <> 0, CboThang.ItemData(CboThang.ListIndex), 1)
    Select Case Index
        Case 0:
            If vt.MaSo = 0 Then GoTo KT

            luong = Cdbl5(txtNhap(3).Text)
            If luong <= 0 Then
                RFocus txtNhap(3)
                Exit Sub
            End If
            
            ExecuteSQL5 ("UPDATE DinhMuc SET SoLuong=" + DoiDau(luong) + "  WHERE MaTP=" + CStr(vattu.MaSo) + " AND MaNVL=" + CStr(vt.MaSo) + " AND Thang=" + CStr(i))
            If DBKetoan.RecordsAffected > 0 Then
                cboThang_Click
            Else
                If ExecuteSQL5("INSERT INTO DinhMuc (MaSo,MaTP,MaNVL,SoLuong,Thang) VALUES (" + CStr(Lng_MaxValue("MaSo", "DinhMuc") + 1) + "," + CStr(vattu.MaSo) + "," + CStr(vt.MaSo) + "," + DoiDau(luong) + "," + CStr(i) + ")") = 0 Then
                    GrdNT(0).AddItem vt.sohieu + Chr(9) + vt.TenVattu + Chr(9) + vt.DonVi + Chr(9) + Format(luong, Mask_2), 0
                End If
            End If
            
            vt.InitVattuMaSo 0
            For i = 0 To 3
                txtNhap(i).Text = ""
            Next
KT:
            RFocus txtNhap(0)
        Case 1:
            If ts.MaSo = 0 Then GoTo KT1
        
            luong = Cdbl5(txtNhap(6).Text)
            If luong <= 0 Then
                RFocus txtNhap(6)
                Exit Sub
            End If
            ExecuteSQL5 ("UPDATE DinhMuc SET SoLuong=" + DoiDau(luong) + "  WHERE MaTP=" + CStr(vattu.MaSo) + " AND MaNVL=" + CStr(-ts.MaSo) + " AND Thang=" + CStr(i))
            If DBKetoan.RecordsAffected > 0 Then
                cboThang_Click
            Else
                If ExecuteSQL5("INSERT INTO DinhMuc (MaSo,MaTP,MaNVL,SoLuong,Thang) VALUES (" + CStr(Lng_MaxValue("MaSo", "DinhMuc") + 1) + "," + CStr(vattu.MaSo) + "," + CStr(-ts.MaSo) + "," + DoiDau(luong) + "," + CStr(i) + ")") = 0 Then
                    GrdNT(1).AddItem ts.sohieu + Chr(9) + ts.Ten + Chr(9) + Format(luong, Mask_2), 0
                End If
            End If
            ts.KhoiTao
            For i = 4 To 6
                txtNhap(i).Text = ""
            Next
KT1:
            RFocus txtNhap(4)
        Case 2:
            luong = Cdbl5(txtVT(5).Text)
            If vattu.MaSo = 0 Or luong = 0 Then Exit Sub
            gia = Cdbl5(txtVT(11).Text)
            With GrdNT(3)
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 0
                    If Len(.Text) = 0 Then Exit For
                    If .Text = txtVT(4).Text Then
                        If ThemMoi = 0 Then
                            .col = 3
                            ExecuteSQL5 "UPDATE DVTVattu SET TyleQD=" + DoiDau(luong) + ",GiaBan=" + DoiDau(gia) + " WHERE MaSo=" + .Text
                        End If
                        .col = 1
                        .Text = Format(luong, Mask_2)
                        .col = 2
                        .Text = Format(gia, Mask_2)
                        Exit Sub
                    End If
                Next
                If ThemMoi = 0 Then
                    ExecuteSQL5 "INSERT INTO DVTVattu (MaSo,MaVattu,DonVi,TyleQD,GiaBan) VALUES (" + CStr(Lng_MaxValue("MaSo", "DVTVattu") + 1) + "," + CStr(vattu.MaSo) + ",'" + txtVT(4).Text + "'," + DoiDau(luong) + "," + DoiDau(gia) + ")"
                    GrdNT(3).AddItem txtVT(4).Text + Chr(9) + Format(luong, Mask_2) + Chr(9) + Format(gia, Mask_2) + Chr(9) + CStr(Lng_MaxValue("MaSo", "DVTVattu")), 0
                    vattu.KTraDVT2
                Else
                    GrdNT(3).AddItem txtVT(4).Text + Chr(9) + Format(luong, Mask_2) + Chr(9) + Format(gia, Mask_2), 0
                End If
            End With
KT2:
            txtVT(4).Text = ""
            txtVT(5).Text = ""
            txtVT(11).Text = ""
            RFocus txtVT(4)
    End Select
End Sub

Private Sub GrdNT_DblClick(Index As Integer)
    Dim i As Integer, ms As Long
    
    Select Case Index
        Case 0:
            With GrdNT(0)
                .col = 0
                If Len(.Text) = 0 Then Exit Sub
                vt.InitVattuSohieu .Text
                ms = SelectSQL("SELECT MaSo AS F1 FROM DinhMuc WHERE MaTP=" + CStr(vattu.MaSo) + " AND MaNVL=" + CStr(vt.MaSo))
                If ExecuteSQL5("DELETE * FROM DinhMuc WHERE MaTP=" + CStr(vattu.MaSo) + " AND MaNVL=" + CStr(vt.MaSo)) <> 0 Then
                    vt.InitVattuMaSo 0
                    Exit Sub
                End If
                For i = 0 To 3
                    .col = i
                    txtNhap(i).Text = .Text
                Next
                .RemoveItem .Row
                If .Rows < .tag Then .Rows = .tag
                vt.InitVattuSohieu txtNhap(0).Text
                RFocus txtNhap(0)
            End With
        Case 1:
            With GrdNT(1)
                .col = 0
                If Len(.Text) = 0 Then Exit Sub
                ts.ChiDinhSH .Text
                If ExecuteSQL5("DELETE * FROM DinhMuc WHERE MaTP=" + CStr(vattu.MaSo) + " AND MaNVL=" + CStr(-ts.MaSo)) <> 0 Then
                    ts.KhoiTao
                    Exit Sub
                End If
                For i = 0 To 2
                    .col = i
                    txtNhap(i + 4).Text = .Text
                Next
                .RemoveItem .Row
                If .Rows < .tag Then .Rows = .tag
                ts.ChiDinhSH txtNhap(4).Text
                RFocus txtNhap(4)
            End With
        Case 3:
            With GrdNT(3)
                .col = 0
                If Len(.Text) = 0 Then Exit Sub
                If ThemMoi = 0 Then
                    txtVT(4).Text = .Text
                    .col = 1
                    txtVT(5).Text = .Text
                    .col = 2
                    txtVT(11).Text = .Text
                    .col = 3
                    If vattu.XoaDVT(CLng5(.Text)) Then
xoa:
                        .RemoveItem .Row
                        If GrdNT(3).Rows < GrdNT(3).tag Then GrdNT(3).Rows = GrdNT(3).tag
                    End If
                Else
                    GoTo xoa
                End If
                RFocus txtVT(4)
            End With
    Case 4:
     With GrdNT(4)
      .col = 1
     txtsoluong.Text = .Text
     .col = 2
     txtsolo.Text = .Text
     .col = 3
     txthandung.Text = IIf(Len(.Text) <= 2, "01/01/11", .Text)
    
      so_dong_hien_tai = .Row
     End With
    End Select
End Sub

Private Sub GrdNt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ms As Long
    
    Select Case KeyAscii
        Case 13:                GrdNT_DblClick (Index)
    End Select
End Sub

Private Sub txtNhap_GotFocus(Index As Integer)
    AutoSelect txtNhap(Index)
End Sub
'====================================================================================================
' HiÓn thÞ danh s¸ch nguyªn tÖ
'====================================================================================================
Private Sub LietKeDinhMuc(Optional ThemMoi As Integer = 0)
    Dim rs As Recordset, thang1 As Integer
    
    ClearGrid GrdNT(0), GrdNT(0).tag
    ClearGrid GrdNT(1), GrdNT(1).tag
    txtTon(2).Text = "0"
    If pDinhmuc <> 0 Then
        thang1 = SelectSQL("SELECT TOP 1 Thang AS F1 FROM DinhMuc WHERE MaTP=" + CStr(vattu.MaSo) + " AND " + WThang("Thang", 0, CboThang.ItemData(CboThang.ListIndex)) + " ORDER BY " + SetMonthOrder("Thang") + " DESC")
        If thang1 = 0 Then Exit Sub
        If ThemMoi = 0 Then CboThang.ListIndex = thang1 - 1
    Else
        thang1 = 1
    End If
    ClearGrid GrdNT(0), GrdNT(0).tag
    ClearGrid GrdNT(1), GrdNT(1).tag
    Set rs = DBKetoan.OpenRecordset("SELECT SoHieu,TenVattu,DonVi,SoLuong,MaNVL FROM DinhMuc INNER JOIN Vattu ON DinhMuc.MaNVL=Vattu.MaSo WHERE MaTP=" + CStr(vattu.MaSo) + " AND Thang=" + CStr(thang1) + " ORDER BY SoHieu DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        GrdNT(0).AddItem rs!sohieu + Chr(9) + rs!TenVattu + Chr(9) + rs!DonVi + Chr(9) + Format(rs!SoLuong, Mask_2), 0
        rs.MoveNext
    Loop
    GrdNT(0).Rows = IIf(rs.RecordCount > GrdNT(0).tag, rs.RecordCount, GrdNT(0).tag)
    GrdNT(0).Row = 0
    GrdNT(0).col = 0
    txtTon(2).Text = Format(SelectSQL("SELECT SoLuong AS F1 FROM DinhMuc WHERE MaNVL=0 AND MaTP=" + CStr(vattu.MaSo) + " And Thang = " + CStr(thang1)), Mask_2)
    
    Set rs = DBKetoan.OpenRecordset("SELECT SoHieu,Ten,SoLuong FROM DinhMuc INNER JOIN TaiSan ON DinhMuc.MaNVL=-TaiSan.MaSo WHERE MaTP=" + CStr(vattu.MaSo) + " AND Thang=" + CStr(thang1) + " ORDER BY SoHieu DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        GrdNT(1).AddItem rs!sohieu + Chr(9) + rs!Ten + Chr(9) + Format(rs!SoLuong, Mask_2), 0
        rs.MoveNext
    Loop
    GrdNT(1).Rows = IIf(rs.RecordCount > GrdNT(1).tag, rs.RecordCount, GrdNT(1).tag)
    GrdNT(1).Row = 0
    GrdNT(1).col = 0
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub txtNhap_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim f As FrmVattu
    Select Case Index
        Case 0:
            If KeyAscii = vbKeyReturn Then
                Set f = New FrmVattu
                'Load f
                txtNhap(0).Text = f.ChonVattu(txtNhap(0).Text, 1)
                RFocus txtNhap(0)
                Set f = Nothing
            End If
        Case 3:
            If KeyAscii = 13 Then CmdChitiet_Click 0 Else KeyProcess txtNhap(Index), KeyAscii
        Case 6:         KeyProcess txtNhap(Index), KeyAscii
        Case 4:
            If KeyAscii = 13 Then
                txtNhap(4).Text = frmDSTaiSan.ChonTaiSan(txtNhap(4).Text, 1, 12)
                RFocus txtNhap(4)
            End If
    End Select
End Sub

Private Sub txtNhap_LostFocus(Index As Integer)
    Select Case Index
        Case 0:
            vt.InitVattuSohieu txtNhap(0).Text
            txtNhap(1).Text = vt.TenVattu
            txtNhap(2).Text = vt.DonVi
        Case 3, 6:
            txtNhap(Index).Text = Format(txtNhap(Index).Text, Mask_2)
        Case 4:
            ts.ChiDinhSH txtNhap(Index).Text
            txtNhap(5).Text = ts.Ten
    End Select
End Sub

Public Function ThemVattu(sh As String) As String
    If xT = 1 Then Exit Function
    Me.tag = 1
    xT = 1
    xSH = sh
    Me.Show 1
    xT = 0
    xSH = ""
    ThemVattu = vattu.sohieu
End Function

