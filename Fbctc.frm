VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FBcTC 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "B¸o c¸o thuÕ - Tµi chÝnh"
   ClientHeight    =   6195
   ClientLeft      =   1620
   ClientTop       =   2610
   ClientWidth     =   9855
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Fbctc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6195
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Financial Report"
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   495
      Left            =   360
      TabIndex        =   116
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
      Begin VB.CheckBox ChkDu 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nhãm theo ngµy"
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
         Left            =   120
         TabIndex        =   131
         Tag             =   "Group by Date"
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox ChkDu 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nhãm theo ®èi øng"
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
         Left            =   120
         TabIndex        =   130
         Tag             =   "Group by Reference Account"
         Top             =   2880
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00E0E0E0&
         Caption         =   ".1"
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
         Index           =   14
         Left            =   1560
         TabIndex        =   129
         Tag             =   "Voucher Book"
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sæ c¸i tµi kho¶n"
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
         Left            =   120
         TabIndex        =   128
         Tag             =   "Ledger"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sè d­ tµi kho¶n"
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
         TabIndex        =   127
         Tag             =   "Closing balance of Account"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox CTGS 
         Height          =   330
         ItemData        =   "Fbctc.frx":57E2
         Left            =   240
         List            =   "Fbctc.frx":57E9
         Style           =   2  'Dropdown List
         TabIndex        =   126
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sæ c¸i tµi kho¶n (chøng tõ ghi sæ)"
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
         Index           =   15
         Left            =   120
         TabIndex        =   125
         Tag             =   "Ledger 2 (group by V. Book)"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng tæng hîp chøng tõ gèc cïng lo¹i"
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
         Left            =   -120
         TabIndex        =   124
         Top             =   1080
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Chøng tõ ghi sæ 2"
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
         Index           =   21
         Left            =   1320
         TabIndex        =   123
         Top             =   2400
         Width           =   1695
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sæ c¸i tµi kho¶n"
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
         Left            =   240
         TabIndex        =   122
         Tag             =   "Ledger"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sæ tæng hîp chøng tõ ghi sæ"
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
         Index           =   16
         Left            =   0
         TabIndex        =   121
         Tag             =   "Summary Voucher Book"
         Top             =   720
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ListBox lstVNI 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "Fbctc.frx":57F3
         Left            =   600
         List            =   "Fbctc.frx":57F5
         TabIndex        =   120
         Top             =   0
         Width           =   2625
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00FFC0C0&
         Caption         =   "B¶ng c©n ®èi ph¸t sinh bµn cê"
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   119
         Tag             =   "Group Balance Sheet"
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00FFC0C0&
         Caption         =   "In toµn bé c¸c chøng tõ ghi sæ"
         Height          =   255
         Index           =   19
         Left            =   0
         TabIndex        =   118
         Tag             =   "Print All Voucher Book"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00FFC0C0&
         Caption         =   "In toµn bé sæ c¸i"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   117
         Tag             =   "Print All Ledgers"
         Top             =   1320
         Width           =   3015
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2160
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tæng"
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
      Index           =   22
      Left            =   1440
      TabIndex        =   115
      Tag             =   "Ledger"
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   3
      Left            =   840
      Picture         =   "Fbctc.frx":57F7
      Style           =   1  'Graphical
      TabIndex        =   114
      Tag             =   "&View"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdtk 
      Height          =   375
      Index           =   3
      Left            =   1920
      Picture         =   "Fbctc.frx":5B39
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdtk 
      Height          =   375
      Index           =   1
      Left            =   1920
      Picture         =   "Fbctc.frx":5FB3
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdtk 
      Height          =   375
      Index           =   0
      Left            =   1680
      Picture         =   "Fbctc.frx":642D
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   120
      Picture         =   "Fbctc.frx":68A7
      Style           =   1  'Graphical
      TabIndex        =   106
      Tag             =   "&Return"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   120
      Picture         =   "Fbctc.frx":7CC9
      Style           =   1  'Graphical
      TabIndex        =   105
      Tag             =   "&Print"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   1440
      Picture         =   "Fbctc.frx":912B
      Style           =   1  'Graphical
      TabIndex        =   104
      Tag             =   "&View"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Index           =   1
      Left            =   2760
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ThuyÕt minh b¸o c¸o tµi chÝnh (excel)"
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
         Index           =   7
         Left            =   120
         TabIndex        =   133
         Tag             =   "Yearly financial report"
         Top             =   1680
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton OptBC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¸o c¸o l­u chuyÓn tiÒn tÖ (trùc tiÕp)"
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
         Index           =   11
         Left            =   120
         TabIndex        =   132
         Tag             =   "Cash flow"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Theo mÉu Q§ sè 15/2006/Q§-BTC"
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
         Left            =   3120
         TabIndex        =   103
         Tag             =   "Print Bar Code"
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.OptionButton OptCD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng c©n ®èi kÕ to¸n"
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
         TabIndex        =   31
         Tag             =   "Balance Sheet"
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton OptCD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "C¸c chØ tiªu ngoµi b¶ng"
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
         Left            =   120
         TabIndex        =   32
         Tag             =   "Off Balance Sheet Items"
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00E0E0E0&
      Height          =   6975
      Index           =   2
      Left            =   2760
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   6
         Left            =   5280
         Picture         =   "Fbctc.frx":A29D
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   5
         Left            =   5280
         Picture         =   "Fbctc.frx":A717
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   4
         Left            =   5280
         Picture         =   "Fbctc.frx":AB91
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng kª ®Çu ra theo tµi kho¶n"
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
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   102
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "§èi øng"
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
         Left            =   4200
         TabIndex        =   101
         Tag             =   "Print Bar Code"
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   6
         Left            =   4200
         LinkItem        =   "Sè hiÖu tµi kho¶n cÇn xem"
         MaxLength       =   12
         TabIndex        =   99
         Tag             =   "0"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   5
         Left            =   4200
         LinkItem        =   "Sè hiÖu tµi kho¶n cÇn xem"
         MaxLength       =   12
         TabIndex        =   86
         Tag             =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng kª ®Çu vµo theo tµi kho¶n"
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
         Index           =   9
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   85
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "M· v¹ch"
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
         Left            =   4200
         TabIndex        =   55
         Tag             =   "Print Bar Code"
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tê khai thuÕ tiªu thô ®Æc biÖt"
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
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   68
         Tag             =   "Monthly Excise Tax Declaration Form"
         Top             =   4080
         Width           =   3015
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   4
         Left            =   4680
         LinkItem        =   "Sè hiÖu tµi kho¶n cÇn xem"
         MaxLength       =   12
         TabIndex        =   51
         Tag             =   "0"
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox ChkDu 
         BackColor       =   &H00E0E0E0&
         Caption         =   "TK"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   50
         Tag             =   "Acc"
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng kª hµng ho¸ dÞch vô b¸n ra chÞu thuÕ TT§B"
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
         Index           =   7
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   66
         Tag             =   "Sales Product List of Excise Tax"
         Top             =   3600
         Width           =   4215
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng kª ho¸ ®¬n, chøng tõ ®Çu ra cã thuÕ TT§B"
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
         Index           =   6
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   65
         Tag             =   "Invoice List of Excise tax"
         Top             =   3120
         Width           =   4335
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng kª ho¸ ®¬n, chøng tõ ®Çu vµo kh«ng ho¸ ®¬n"
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
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   44
         Tag             =   "VAT Inputs without Invoice"
         Top             =   2640
         Width           =   4215
      End
      Begin VB.ComboBox CboTL 
         Height          =   330
         ItemData        =   "Fbctc.frx":AD0F
         Left            =   4680
         List            =   "Fbctc.frx":AD37
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tê khai thuÕ gi¸ trÞ gia t¨ng"
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
         Index           =   5
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   45
         Tag             =   "Monthly VAT Declaration Form"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng kª ho¸ ®¬n ®Çu vµo"
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
         Index           =   4
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   43
         Tag             =   "VAT Input Table"
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng kª ho¸ ®¬n  ®Çu ra"
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
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   36
         Tag             =   "VAT Ouput Table"
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sæ chi tiÕt thuÕ GTGT ®­îc miÔn gi¶m"
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
         Left            =   960
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   35
         Tag             =   "VAT Abatement Detail Report"
         Top             =   4920
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.OptionButton OptVAT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sæ chi tiÕt thuÕ GTGT ®­îc hoµn l¹i"
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
         Left            =   2520
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   34
         Tag             =   "Returned VAT Detail Report"
         Top             =   4800
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   100
         Tag             =   "1"
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LbTenTk 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   87
         Tag             =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LbTenTk 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   73
         Tag             =   "1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LbTenTk 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   67
         Tag             =   "1"
         Top             =   3360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tû lÖ %"
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
         Left            =   4080
         TabIndex        =   48
         Tag             =   "Rate (%)"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00E0E0E0&
      Height          =   6975
      Index           =   3
      Left            =   2760
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      Begin VB.OptionButton OptQT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Chi tiÕt thuÕ GTGT"
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
         Index           =   6
         Left            =   120
         TabIndex        =   64
         Tag             =   "Detail Output VAT"
         Top             =   3120
         Width           =   4215
      End
      Begin VB.OptionButton OptQT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng ph©n tÝch thuÕ GTGT ®Çu vµo ®­îc khÊu trõ"
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
         Index           =   4
         Left            =   120
         TabIndex        =   63
         Tag             =   "Detail Deducted Input VAT"
         Top             =   2160
         Width           =   4215
      End
      Begin VB.OptionButton OptQT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Doanh sè vµ thuÕ GTGT ®Çu ra"
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
         Index           =   5
         Left            =   120
         TabIndex        =   62
         Tag             =   "Summary Turnover and Output VAT Report"
         Top             =   2640
         Width           =   4215
      End
      Begin VB.OptionButton OptQT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng kª t×nh h×nh nép thuÕ"
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
         Left            =   120
         TabIndex        =   61
         Tag             =   "Detail Tax Payment Report"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.OptionButton OptQT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "QuyÕt to¸n thuÕ thu nhËp doanh nghiÖp"
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
         Left            =   120
         TabIndex        =   60
         Tag             =   "Yearly Income Tax Report"
         Top             =   1200
         Width           =   4215
      End
      Begin VB.OptionButton OptQT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng tæng hîp thanh to¸n thuÕ"
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
         TabIndex        =   59
         Tag             =   "Tax Payment Report"
         Top             =   240
         Value           =   -1  'True
         Width           =   4215
      End
      Begin VB.OptionButton OptQT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "QuyÕt to¸n thuÕ GTGT"
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
         Left            =   120
         TabIndex        =   58
         Tag             =   "Yearly VAT Report"
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00E0E0E0&
      Height          =   7095
      Index           =   7
      Left            =   2760
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      Begin VB.OptionButton OptKP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¸o c¸o t×nh h×nh thùc hiÖn c¸c chØ tiªu dù to¸n"
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
         Index           =   9
         Left            =   120
         TabIndex        =   98
         Top             =   4560
         Width           =   5655
      End
      Begin VB.OptionButton OptKP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¸o c¸o t×nh h×nh sö dông sè KP quyÕt to¸n n¨m tr­íc chuyÓn sang"
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
         Left            =   120
         TabIndex        =   97
         Top             =   4080
         Width           =   5655
      End
      Begin VB.OptionButton OptKP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¸o c¸o t×nh h×nh cho vay, thu nî, thu l·i"
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
         Index           =   7
         Left            =   120
         TabIndex        =   96
         Top             =   3600
         Width           =   5655
      End
      Begin VB.OptionButton OptKP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¸o c¸o tiÕp nhËn vµ sö dông viÖn trî kh«ng hoµn l¹i"
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
         Index           =   6
         Left            =   120
         TabIndex        =   95
         Top             =   3120
         Width           =   5655
      End
      Begin VB.OptionButton OptKP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng ®èi chiÕu h¹n møc kinh phÝ"
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
         Index           =   5
         Left            =   120
         TabIndex        =   94
         Top             =   2640
         Width           =   5655
      End
      Begin VB.OptionButton OptKP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Chi tiÕt kinh phÝ ho¹t ®éng ®Ò nghÞ quyÕt to¸n"
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
         Index           =   4
         Left            =   120
         TabIndex        =   93
         Top             =   2160
         Width           =   5655
      End
      Begin VB.OptionButton OptKP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tæng hîp t×nh h×nh kinh phÝ vµ quyÕt to¸n kinh phÝ ®· sö dông"
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
         Left            =   120
         TabIndex        =   92
         Top             =   1680
         Width           =   5655
      End
      Begin VB.OptionButton OptKP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sæ theo dâi h¹n møc kinh phÝ"
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
         Left            =   120
         TabIndex        =   91
         Top             =   1200
         Width           =   5655
      End
      Begin VB.OptionButton OptKP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tæng hîp chi tiÕt nguån kinh phÝ"
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
         Left            =   120
         TabIndex        =   90
         Top             =   720
         Width           =   5655
      End
      Begin VB.OptionButton OptKP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sæ theo dâi nguån kinh phÝ"
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
         TabIndex        =   89
         Top             =   240
         Value           =   -1  'True
         Width           =   5655
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00E0E0E0&
      Height          =   6975
      Index           =   0
      Left            =   2760
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   2
         Left            =   5400
         Picture         =   "Fbctc.frx":AD65
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   1499
         Width           =   375
      End
      Begin VB.OptionButton OptKqkd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng ph©n tÝch doanh thu chi phÝ"
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
         Index           =   9
         Left            =   120
         TabIndex        =   77
         Tag             =   "Summary Production Expenses"
         Top             =   4560
         Width           =   2775
      End
      Begin VB.ComboBox CboNK 
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
         Index           =   2
         ItemData        =   "Fbctc.frx":AEE3
         Left            =   5280
         List            =   "Fbctc.frx":B076
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox ChkDu 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Chi tiÕt theo tµi kho¶n 911 ®Õn cÊp"
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
         Index           =   5
         Left            =   2280
         TabIndex        =   21
         Tag             =   "Group by Date"
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton OptKqkd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ThuÕ GTGT ®­îc khÊu trõ, ®­îc hoµn l¹i, miÔn gi¶m"
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
         Left            =   120
         TabIndex        =   24
         Tag             =   "Deducted, Refunded, Exempted VAT"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.OptionButton OptKqkd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¸o c¸o l·i lç"
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
         TabIndex        =   20
         Tag             =   "Profit and Lost"
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton OptKqkd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¸o c¸o thuÕ"
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
         Left            =   120
         TabIndex        =   23
         Tag             =   "Summary Tax Report"
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton OptKqkd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng kª chi tiÕt doanh thu, chi phÝ, thu nhËp"
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
         Left            =   120
         TabIndex        =   29
         Tag             =   "Detail Profit and Loss"
         Top             =   4080
         Width           =   5655
      End
      Begin VB.OptionButton OptKqkd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Doanh thu chi tiÕt"
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
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Tag             =   "Detail Turnover Report"
         Top             =   3600
         Width           =   3615
      End
      Begin VB.OptionButton OptKqkd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sæ chi phÝ s¶n xuÊt kinh doanh"
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
         Index           =   6
         Left            =   120
         TabIndex        =   69
         Tag             =   "Summary Production Expenses"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.OptionButton OptKqkd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tæng hîp phÝ"
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
         Index           =   5
         Left            =   120
         TabIndex        =   53
         Tag             =   "Summary Expenses Report"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.OptionButton OptKqkd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¶ng kª tæng hîp hîp ®ång kinh tÕ"
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
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Tag             =   "Contract List"
         Top             =   3120
         Width           =   4215
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   2
         Left            =   4200
         LinkItem        =   "Sè hiÖu tµi kho¶n ®èi øng"
         MaxLength       =   12
         TabIndex        =   26
         Tag             =   "0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton OptKqkd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B¸o c¸o chi phÝ"
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
         Left            =   120
         TabIndex        =   25
         Tag             =   "Detail Expenses Report"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label LbTenTk 
         BackColor       =   &H00E0E0E0&
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
         Height          =   615
         Index           =   2
         Left            =   4440
         TabIndex        =   46
         Tag             =   "1"
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00E0E0E0&
      Height          =   6975
      Index           =   6
      Left            =   2760
      TabIndex        =   78
      Top             =   -120
      Width           =   7095
      Begin VB.CheckBox ChkTT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Th«ng tin 3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   84
         Tag             =   "Index"
         ToolTipText     =   "by Object"
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox CboTT 
         Height          =   330
         Index           =   2
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   960
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox ChkTT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Th«ng tin 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Tag             =   "Index"
         ToolTipText     =   "by Object"
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox CboTT 
         Height          =   330
         Index           =   1
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   600
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox ChkTT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Th«ng tin 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   80
         Tag             =   "Index"
         ToolTipText     =   "by Object"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox CboTT 
         Height          =   330
         Index           =   0
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Index           =   5
      Left            =   3000
      TabIndex        =   74
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
      Begin VB.OptionButton OptNN 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Anh"
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
         Left            =   1320
         TabIndex        =   76
         Tag             =   "English"
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton OptNN 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ViÖt"
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
         TabIndex        =   75
         Tag             =   "Vietnamese"
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   4
      Left            =   240
      TabIndex        =   70
      Top             =   2520
      Width           =   1575
      Begin VB.ComboBox CboThang 
         Height          =   330
         Index           =   0
         ItemData        =   "Fbctc.frx":B186
         Left            =   240
         List            =   "Fbctc.frx":B18D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox CboThang 
         Height          =   330
         Index           =   1
         ItemData        =   "Fbctc.frx":B19B
         Left            =   240
         List            =   "Fbctc.frx":B1A2
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton OptTG 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tõ th¸ng"
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
         TabIndex        =   0
         Tag             =   "From"
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptTG 
         BackColor       =   &H00FFFF80&
         Caption         =   "Tõ ngµy"
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
         TabIndex        =   3
         Tag             =   "From"
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSMask.MaskEdBox MedNgay 
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         BackColor       =   &H00FFFFC0&
         Caption         =   "®Õn th¸ng"
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
         Index           =   6
         Left            =   240
         TabIndex        =   72
         Tag             =   "to"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "®Õn ngµy"
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
         TabIndex        =   71
         Tag             =   "to"
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H0080FF80&
      Caption         =   "C¸c quyÕt to¸n thuÕ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   17
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   56
      Tag             =   "Yearly final tax reports"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ComboBox CboNK 
      Height          =   330
      Index           =   1
      ItemData        =   "Fbctc.frx":B1B0
      Left            =   1920
      List            =   "Fbctc.frx":B1B7
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox CboNK 
      Height          =   315
      Index           =   0
      ItemData        =   "Fbctc.frx":B1C2
      Left            =   0
      List            =   "Fbctc.frx":B1C9
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   6480
      Width           =   615
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00FFFFC0&
      Caption         =   "B¶ng kª"
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
      Left            =   480
      TabIndex        =   16
      Tag             =   "Debit Journal"
      Top             =   6240
      Width           =   1935
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NhËt ký chøng tõ"
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
      Left            =   0
      TabIndex        =   15
      Tag             =   "Credit Journal"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B¶ng kª ph¸t sinh cã"
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
      Index           =   9
      Left            =   0
      TabIndex        =   13
      Tag             =   "Credit Table of Account"
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00FFFFC0&
      Caption         =   "B¶ng kª ph¸t sinh nî"
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
      Index           =   6
      Left            =   1440
      TabIndex        =   12
      Tag             =   "Debit Table of Account"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   3
      Left            =   7320
      LinkItem        =   "Sè hiÖu tµi kho¶n cÇn xem"
      MaxLength       =   12
      TabIndex        =   14
      Tag             =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   0
      Left            =   3600
      LinkItem        =   "Sè hiÖu tµi kho¶n cÇn xem"
      MaxLength       =   12
      TabIndex        =   7
      Tag             =   "0"
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox CboVV 
      Height          =   330
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CheckBox ChkDu 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Theo vô viÖc"
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
      Left            =   4560
      TabIndex        =   8
      Tag             =   "Index"
      ToolTipText     =   "by Object"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ComctlLib.ProgressBar GauGe 
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   6600
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H0080FF80&
      Caption         =   "B¸o c¸o thuÕ GTGT"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   33
      Tag             =   "VAT and Excise Tax Reports"
      Top             =   240
      Width           =   2295
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H0080FF80&
      Caption         =   "B¶ng c©n ®èi kÕ to¸n"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   30
      Tag             =   "Balance sheet"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.OptionButton OptBC 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "KÕt qu¶ kinh doanh"
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
      Height          =   405
      Index           =   3
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "Profit and Lost"
      Top             =   720
      Width           =   2295
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sæ nhËt ký chung"
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
      TabIndex        =   6
      Tag             =   "Journal Ledger"
      Top             =   6240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   1
      Left            =   6000
      LinkItem        =   "Sè hiÖu tµi kho¶n ®èi øng"
      MaxLength       =   12
      TabIndex        =   11
      Tag             =   "0"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox ChkDu 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Theo ®èi øng"
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
      Left            =   1080
      TabIndex        =   10
      Tag             =   "by Ref. Acc."
      ToolTipText     =   "by Ref. Acc."
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "CÊp TK ®èi øng"
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
      Index           =   4
      Left            =   0
      TabIndex        =   54
      Tag             =   "Level of Account"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sè"
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
      Left            =   5160
      TabIndex        =   52
      Tag             =   "No."
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   -1080
      TabIndex        =   47
      Tag             =   "1"
      Top             =   5520
      Width           =   4575
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   7710
      TabIndex        =   38
      Tag             =   "1"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   37
      Tag             =   "1"
      Top             =   4680
      Width           =   4695
   End
End
Attribute VB_Name = "FBcTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const w1 = 100

Dim baocao As Integer
Dim StopPrint As Boolean
Dim Pdelay As Integer
Dim tkdt1 As String, tkdt2 As String, tkdt3 As String, tkdt4 As String, vatr As String
Dim vatv1 As String, vatv2 As String, vatv3 As String, vatv4 As String
Dim xlapp As Excel.Application, xlsheet As Worksheet
Dim sodcthue As Double
Dim NLB As String, KTT As String, GD As String
Dim ngay(0 To 1) As Date
Dim nn As Integer

Private Sub CboTT_Click(Index As Integer)
    If Index = 0 Then CboVVClick CboTT(0), CboTT(1)
End Sub
Sub PopulateListBoxFromFile(ByVal AListBox, ByVal AFileName As String, ByVal HasFolder As Boolean)
' Display a text file in a listbox
   Dim theFileName As String
   Dim FileNum, aLine
   ' If Filename is not fullpath then use LocalDirectory
   If HasFolder Then
      theFileName = AFileName
   Else
      theFileName = GetLocalDirectory & AFileName
   End If
   ' Get out if the input file does not exist
   If Dir(theFileName) = "" Then Exit Sub
   FileNum = FreeFile
   AListBox.Clear  ' Clear the listbox
   ' Open input file
   Open theFileName For Input As FileNum
   ' Read till End-Of-File
   Do While Not EOF(FileNum)
      Line Input #FileNum, aLine  '  Read a Text line
      AListBox.AddItem aLine      ' Add the line to the listbox
   Loop
   Close FileNum  ' Close the input file
End Sub
Sub chutieng_viet()
Dim MyUnicodeText
Dim UVowels As String
Dim mDOMVowels As Class1
 'Set MyUnicodeText = New Class1
  Set mDOMVowels = New Class1

        ' Read Unicode Text from file txtFileName and display in TextBox1(0)
       'TextBox1(0).Text = MyUnicodeText.ReadUnicode(txtFileName)
         UVowels = mDOMVowels.ReadUnicode(pCurDir & "UnicodeVowels.xml")
           PopulateListBoxFromFile lstVNI, "VNIVowelMap.txt", False
      ' LbCongty.Caption = MyUnicodeText.ReadUnicode("D:\soft\sv\Accounting\config.xml")
End Sub
Function MulticharToUnicode(UVowels, chuoi) As String
' Convert a multi-character vowel in VNI or VIQR to Unicode
Dim letter As String
Dim Text1 As String
Dim Text2 As String
Dim MapNum As String
Dim i, pos, item, TLen
    ' Assign content of input Textbox to Text1
    'Text1 = TextBox1(0).Text
    If IsNull(chuoi) Then
        Text1 = ""
    Else
        Text1 = chuoi
    End If
    ' Replace every multi-character vowel in Text1 with a string like |067 that represents
    ' the 67th Unicode vowel
    ' Iterate through every multi-character vowel
    For i = 0 To lstVNI.ListCount - 1
       ' Get an item from the Vowel listbox
       item = lstVNI.List(i)  ' like a^~016
       letter = RTrim(Left(item, 3))  ' isolate the multi-character vowel eg: a^~
       MapNum = "|" & Right(item, 3)  ' Prefix the | character to the digit string eg: &016
      
       Text1 = Replace(Text1, letter, MapNum) ' replace all occurences of the vowel
    Next
    ' Now map the position strings like &016, &114 to 16th and 114th Unicode vowels
    i = 1
    TLen = Len(Text1)
    Do While i <= TLen
       ' get a character
       letter = Mid(Text1, i, 1)
        ' if it's a Carriage return or a LineFeed then just copy across
        If (letter = vbCr) Then
           Text2 = Text2 & vbCr
               i = i + 1
        ElseIf (letter = vbLf) Then
           Text2 = Text2 & vbLf
               i = i + 1
        ElseIf letter <> "|" Then
            ' merely copy across everything else
               Text2 = Text2 & letter
               i = i + 1
        Else
           ' get here if encounter a "&", obtain the position of the Unicode vowel
           ' Note that there'll be a bug if the text string contains genuine "|" character
           pos = val(Mid(Text1, i + 1, 3))
           If pos = 0 Then
               Text2 = Text2 & letter
               i = i + 1
           Else
             ' get the Unicode vowel for output
               Text2 = Text2 & Mid(UVowels, pos, 1)
               i = i + 4
           End If
        End If
    Loop
    ' Return the result
    MulticharToUnicode = Text2
End Function
Private Sub insert_rows(UVowels As String, xlsheet, rs)

End Sub

'====================================================================================================
' Chän sè hiÖu tµi kho¶n
'====================================================================================================
Private Sub cmdtk_Click(Index As Integer)
    Me.MousePointer = 11
    txtShTk(Index).Text = FrmTaikhoan.ChonTk(txtShTk(Index).Text)
    Me.MousePointer = 0
    RFocus txtShTk(Index)
End Sub
'====================================================================================================
' Xem, In hoÆc kÕt thóc
'====================================================================================================
Public Sub Command_Click(Index As Integer)
If OptTG(0).Value = False Then
  If IsDate(MedNgay(1).Text) And IsDate(MedNgay(0).Text) Then
        If CDate(MedNgay(1).Text) < CDate(MedNgay(0).Text) Then
            MedNgay(1).Text = MedNgay(0).Text
          
        End If
  End If
  Else
   If IsDate(CboThang(0).Text) And IsDate(CboThang(1).Text) Then
        If CDate(CboThang(1).Text) < CDate(CboThang(0).Text) Then
            CboThang(1).Text = CboThang(0).Text
          
        End If
  End If
  End If
   If IsDate(MedNgay(1).Text) And IsDate(MedNgay(0).Text) Then
        If CDate(MedNgay(1).Text) < CDate(MedNgay(0).Text) Then
            MedNgay(1).Text = MedNgay(0).Text
           ' MsgBox "Thong bao"
       '     Exit Sub
        End If
  End If
  'ket thuc kiem tra ngay
    
    Dim taikhoan As New ClsTaikhoan
    Dim doiung As New ClsTaikhoan, n As Date
    Dim i As Integer, tdau As Integer, tcuoi As Integer, kq1 As Boolean
    Dim mdt1 As Long, mdt2 As Long, mdt3 As Long
    
    Select Case Index
        Case 0, 1:
            If pRpt = 1 And (Not ChoXemBC(baocao, "T")) Then
                ErrMsg er_KoSD
                Exit Sub
            End If
            Me.MousePointer = 11
            GauGe.Value = 0
            If CboThang(1).ListIndex < CboThang(0).ListIndex Then CboThang(1).ListIndex = CboThang(0).ListIndex
            tdau = CboThang(0).ItemData(CboThang(0).ListIndex)
            tcuoi = CboThang(1).ItemData(CboThang(1).ListIndex)
            
            If pSoVV > 0 And ChkTT(0).Value = 1 And CboTT(0).ListIndex >= 0 Then mdt1 = CboTT(0).ItemData(CboTT(0).ListIndex) Else mdt1 = 0
            If pSoVV > 1 And ChkTT(1).Value = 1 And CboTT(1).ListIndex >= 0 Then mdt2 = CboTT(1).ItemData(CboTT(1).ListIndex) Else mdt2 = 0
            If pSoVV > 2 And ChkTT(2).Value = 1 And CboTT(2).ListIndex >= 0 Then mdt3 = CboTT(2).ItemData(CboTT(2).ListIndex) Else mdt3 = 0
            
            SetRptInfo
            Select Case baocao
                Case 20:
                    If txtShTk(0).tag = 0 Then
                        RFocus txtShTk(0)
                        GoTo KhongInBC
                    End If
                    taikhoan.InitTaikhoanMaSo txtShTk(0).tag
                    If OptTG(0).Value And tdau = tcuoi Then
                        ngay(0) = NgayDauThang(pNamTC, tdau)
                        OptTG(1).Value = True
                        For i = 1 To 5
                            ngay(1) = ngay(0) + 6
                            If Month(ngay(1)) <> tcuoi Then ngay(1) = NgayCuoiThang(pNamTC, tcuoi)
                            
                            'Command_Click Index
                            If taikhoan.loai < 5 Or taikhoan.kieu < 0 Then
                                If BangTHCTuGoc(CTGS.ItemData(CTGS.ListIndex), taikhoan.sohieu, ngay(0), ngay(1), -1) Then
                                    InBaoCaoRPT nn
                                    SetRptInfo
                                End If
                            End If
                            If taikhoan.loai < 5 Or taikhoan.kieu > 0 Then
                                If BangTHCTuGoc(CTGS.ItemData(CTGS.ListIndex), taikhoan.sohieu, ngay(0), ngay(1), 1) Then
                                    InBaoCaoRPT nn
                                    SetRptInfo
                                End If
                            End If
                            If Month(ngay(0) + 7) = tdau Then
                                ngay(0) = ngay(0) + 7
                            Else
                                Exit For
                            End If
                        Next
                        MedNgay(0).Text = Format(ngay(0), Mask_D)
                        MedNgay(1).Text = Format(ngay(1), Mask_D)
                        GoTo KhongInBC
                    Else
                        If BangTHCTuGoc(CTGS.ItemData(CTGS.ListIndex), taikhoan.sohieu, IIf(OptTG(0).Value, NgayDauThang(pNamTC, tdau), ngay(0)), IIf(OptTG(0).Value, NgayCuoiThang(pNamTC, tcuoi), ngay(1)), -1) Then
                            InBaoCaoRPT nn
                            SetRptInfo
                        End If
                        If BangTHCTuGoc(CTGS.ItemData(CTGS.ListIndex), taikhoan.sohieu, IIf(OptTG(0).Value, NgayDauThang(pNamTC, tdau), ngay(0)), IIf(OptTG(0).Value, NgayCuoiThang(pNamTC, tcuoi), ngay(1)), 1) Then InBaoCaoRPT nn
                        GoTo KhongInBC
                    End If
                Case 18:
                    SoDangKyCT tdau, tcuoi, IIf(OptTG(1).Value, 1, 0), ngay(0), ngay(1)
                Case 0:
                    If Not InNhatKy(tdau, tcuoi, IIf(ChkDu(1).Value, CboVV.ItemData(CboVV.ListIndex), 0), IIf(OptTG(1).Value, 1, 0), ngay(0), ngay(1), nn, , mdt1, mdt2, mdt3) Then GoTo KhongInBC
                Case 1, 15:
                    If txtShTk(0).tag = 0 Then
                        RFocus txtShTk(0)
                        GoTo KhongInBC
                    End If
                    taikhoan.InitTaikhoanMaSo txtShTk(0).tag
                    If ChkDu(0).Value = 1 And txtShTk(1).tag = 0 Then
                        RFocus txtShTk(1)
                        GoTo KhongInBC
                    End If
                    If ChkDu(0).Value = 1 Then
                        doiung.InitTaikhoanMaSo txtShTk(1).tag
                    Else
                        doiung.InitTaikhoanMaSo 0
                    End If
                    If (taikhoan.MaTC <> taikhoan.MaSo) And (taikhoan.MaTC > 0) Then
                        If OptTG(0).Value Then
                            If Not InSoChitiet(taikhoan, tdau, tcuoi, ngay(0), ngay(1), True, doiung.sohieu, IIf(ChkDu(1).Value, CboVV.ItemData(CboVV.ListIndex), 0), ChkDu(3).Value, nn) Then GoTo KhongInBC
                        Else
                            If Not InSoChitiet(taikhoan, 0, 0, ngay(0), ngay(1), True, doiung.sohieu, IIf(ChkDu(1).Value, CboVV.ItemData(CboVV.ListIndex), 0), ChkDu(3).Value, nn) Then GoTo KhongInBC
                        End If
                    Else
                        If baocao = 1 Then
                            If OptTG(0).Value Then
                                If Not InSocaiTk(taikhoan, tdau, tcuoi, ngay(0), ngay(1), True, doiung.sohieu, IIf(ChkDu(1).Value, CboVV.ItemData(CboVV.ListIndex), 0), ChkDu(3).Value, nn, mdt1, mdt2, mdt3) Then GoTo KhongInBC
                            Else
                                If Not InSocaiTk(taikhoan, tdau, tcuoi, ngay(0), ngay(1), True, doiung.sohieu, IIf(ChkDu(1).Value, CboVV.ItemData(CboVV.ListIndex), 0), ChkDu(3).Value, nn, mdt1, mdt2, mdt3) Then GoTo KhongInBC
                            End If
                        Else
                            If Not InSocaiTk2(taikhoan, tdau, tcuoi, True, doiung.sohieu, IIf(ChkDu(1).Value, CboVV.ItemData(CboVV.ListIndex), 0), nn) Then GoTo KhongInBC
                        End If
                    End If
                Case 2:
                    Dim rs_tk As Recordset
                    If Index = 0 Then
                        ErrMsg er_KoXem
                        GoTo KhongInBC
                    End If
                    Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo, SoHieu, Ten FROM HethongTK WHERE MaTC = MaSo AND Loai > 0 ORDER BY SoHieu", dbOpenSnapshot, dbForwardOnly)
                    StopPrint = False
                    frmMain.Rpt.Destination = 1
                    GauGe.Max = 200
                    
                    Dim nunmberloop As Integer
                    Select Case baocao
                     Case 2, 19:
                      nunmberloop = 1
                    End Select
                    Do While Not rs_tk.EOF And (Not StopPrint)
                        taikhoan.InitTaikhoanMaSo rs_tk!MaSo
                        SetRptInfo
                        If (pSoKT Mod 10 >= 1) Or (pSoKT Mod 1000 >= 100) Then
                            kq1 = InSocaiTk(taikhoan, tdau, tcuoi, ngay(0), ngay(1), False, "", 0, ChkDu(3).Value, nn)
                        Else
                            kq1 = InSocaiTk2(taikhoan, tdau, tcuoi, False, "", 0, nn)
                        End If
                        If kq1 Then
                            HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
                            InBaoCaoRPT
                            AppIdle Pdelay * 100
                        End If
                        If GauGe.Value < GauGe.Max Then GauGe.Value = GauGe.Value + 1
                        rs_tk.MoveNext
                        If nunmberloop = 1 Then Exit Do
                    
                    Loop
                    rs_tk.Close
                    Set rs_tk = Nothing
                    GoTo KhongInBC
                Case 3:
                frmMain.Rpt.Formulas(30) = "diachi='" + frmMain.LbCty(2).Caption + ", " + frmMain.LbCty(10).Caption + ", " + frmMain.LbCty(11).Caption + "'"
                    If OptKqkd(0).Value Then InKqkd tdau, tcuoi, IIf(ChkDu(5).Value = 1, CboNK(2).ListIndex + 2, 0), nn
                    If OptKqkd(1).Value Then InThue tdau, tcuoi
                    If OptKqkd(2).Value Then InVAT tdau, tcuoi
                    If OptKqkd(4).Value Then
                        If Not InTongHopHD(tdau, tcuoi) Then GoTo KhongInBC
                    End If
                    If OptKqkd(3).Value Or OptKqkd(5).Value Then
                        If txtShTk(2).tag = 0 Then
                            RFocus txtShTk(2)
                            GoTo KhongInBC
                        End If
                    End If
                    If OptKqkd(3).Value Then InCTChiphi tdau, tcuoi, txtShTk(2).Text, LbTenTk(2).Caption
                    If OptKqkd(5).Value Then
                        If Not InTongHopPhi(txtShTk(2).Text, tdau, tcuoi) Then GoTo KhongInBC
                    End If
                    If OptKqkd(6).Value Then
                        If Len(txtShTk(2).Text) = 0 Then
                            RFocus txtShTk(2)
                            GoTo KhongInBC
                        End If
                        If Not BKChiTietRPT(txtShTk(2).Text, NgayDauThang(pNamTC, tdau), NgayCuoiThang(pNamTC, tcuoi)) Then GoTo KhongInBC
                    End If
                    If OptKqkd(7).Value Then ChiTietDoanhThu tdau, tcuoi
                    If OptKqkd(8).Value Then InCTKQKD pThangDauKy, ThangCuoiNamTC
                    If OptKqkd(9).Value Then
                        PTDTCP tdau, tcuoi
                        GoTo KhongInBC
                    End If
                Case 4:
                frmMain.Rpt.Formulas(30) = "diachi='" + frmMain.LbCty(2).Caption + ", " + frmMain.LbCty(10).Caption + ", " + frmMain.LbCty(11).Caption + "'"
                    ' them tai khoan 244
                If SelectSQL("SELECT count(maso) AS F1 from HethongTK where SoHieu LIKE '621*'") <= 0 Then
                 
                
                ' If SelectSQL("SELECT count(maso) as f1 from cdts2005 where maso = 320") > 0 Or SelectSQL("SELECT count(maso) as f1 from cdts2005") <= 0 Then
                    ExecuteSQL5 "drop table cdts2005 "
                    ExecuteSQL5 "SELECT * INTO cdts2005 FROM [MS Access;PWD=1234;DATABASE=" + pCurDir + "\REPORTS\bc.rpt].cdts2005"
                     ExecuteSQL5 "UPDATE CDTS2005 SET shtk1 = '244' WHERE MaSo= 241"
               ' End If
                End If
                    If OptCD(0).Value Then
                        If pVersion <> 3 Then
                            InCdts tdau, tcuoi, Chk(2).Value, nn
                        Else
                            InCdts_HCSN tdau, tcuoi, CboNK(1).ListIndex + 1
                        End If
                    Else
                        InTsNb tcuoi
                    End If
                Case 5:
                    If ChkDu(4).Value = 1 And Len(txtShTk(4).Text) > 0 Then
                        taikhoan.InitTaikhoanSohieu txtShTk(4).Text
                    Else
                        taikhoan.InitTaikhoanMaSo 0
                    End If
                    If OptVAT(0).Value Then
                        If Not InVATHoanLai(tdau, tcuoi) Then GoTo KhongInBC
                    End If
                    If OptVAT(1).Value Then
                        If Not InVATMienGiam(tdau, tcuoi) Then GoTo KhongInBC
                    End If
                    
                    'In bang ke
                    If OptVAT(3).Value And Chk(1).Value = 0 Then ' bang ke ra
                     '   If Not InVATDauRa(tdau, tcuoi, CInt(CboTL.ItemData(CboTL.ListIndex)), taikhoan) Then GoTo KhongInBC
                      ' If Not InVATDauRa_exel(tdau, tcuoi, CInt(CboTL.ItemData(CboTL.ListIndex)), taikhoan) Then GoTo KhongInBC
                      ' MsgBox " thanh cong"
                    If MsgBox("b¹n cã mèn chuyÓn ra excel ®Ó cËp nhËt vµo m· v¹ch?", vbYesNo + vbQuestion, App.ProductName) = vbYes Then
                       If Not InVATDauRa_exel(tdau, tcuoi, CInt(CboTL.ItemData(CboTL.ListIndex)), taikhoan) Then GoTo KhongInBC
                      Else
                        If Not InVATDauRa(tdau, tcuoi, CInt(CboTL.ItemData(CboTL.ListIndex)), taikhoan) Then GoTo KhongInBC
                     End If
                    End If
                    'In bang ke
                    If OptVAT(3).Value And Chk(1).Value = 1 Then
                        InVATDauRaMV tcuoi, CInt5(CboTL.ItemData(CboTL.ListIndex)), taikhoan
                        GoTo KhongInBC
                    End If
                    If OptVAT(4).Value And Chk(1).Value = 0 Then
                      '  If Not InVATDauVao2(tdau, tcuoi, CInt(CboTL.ItemData(CboTL.ListIndex)), 1, taikhoan) Then GoTo KhongInBC
                      'bang ke vao ///////////////////////////////////////////////////////
                       If MsgBox("b¹n cã mèn chuyÓn ra excel ®Ó cËp nhËt vµo m· v¹ch?", vbYesNo + vbQuestion, App.ProductName) = vbYes Then
                       If Not InVATDauvao_exel(tdau, tcuoi, CInt(CboTL.ItemData(CboTL.ListIndex)), 1, taikhoan) Then GoTo KhongInBC
                      Else
                        If Not InVATDauVao2(tdau, tcuoi, CInt(CboTL.ItemData(CboTL.ListIndex)), 1, taikhoan) Then GoTo KhongInBC
                    End If
                    End If
                    If OptVAT(4).Value And Chk(1).Value = 1 Then
                        InVATDauVaoMV tcuoi, CInt(CboTL.ItemData(CboTL.ListIndex)), 1, taikhoan
                        GoTo KhongInBC
                    End If
                    If OptVAT(2).Value And Chk(1).Value = 0 Then
                        If Not InVATDauVao2(tdau, tcuoi, -1, 0, taikhoan) Then GoTo KhongInBC
                    End If
                    If OptVAT(2).Value And Chk(1).Value = 1 Then
                        InVATDauVaoMV tcuoi, -1, 0, taikhoan
                        GoTo KhongInBC
                    End If
                    If OptVAT(5).Value And Chk(1).Value = 0 Then
                        ToKhaiVAT tdau, tcuoi, taikhoan
                    End If
                    If OptVAT(5).Value And Chk(1).Value = 1 Then
                        ToKhaiVAT2 tcuoi, taikhoan
                        GoTo KhongInBC
                    End If
                    If OptVAT(6).Value Then InThueTTDB tdau, tcuoi, CInt(CboTL.ItemData(CboTL.ListIndex)), taikhoan
                    If OptVAT(7).Value Then BangKeBanRa tdau, tcuoi, taikhoan
                    If OptVAT(8).Value Then ToKhaiTTDB tdau, tcuoi, taikhoan
                    If OptVAT(9).Value Or OptVAT(10).Value Then
                        InBKTheoTK IIf(OptVAT(9).Value, -1, 1), tdau, tcuoi, txtShTk(5).Text, IIf(Chk(0).Value = 1, txtShTk(6).Text, "")
                    End If
                Case 11:
                    InLCTT2 tdau, tcuoi
                Case 7:
                    InTMTC tdau, tcuoi
                    GoTo KhongInBC
                Case 6:
                    If txtShTk(3).tag = 0 Then
                        RFocus txtShTk(3)
                        GoTo KhongInBC
                    End If
                   BK1 tdau, tcuoi, txtShTk(3).Text, ChkDu(2).Value, CboNK(1).ListIndex + 1
                    GoTo KhongInBC
                Case 9:
                    If txtShTk(3).tag = 0 Then
                        RFocus txtShTk(3)
                        GoTo KhongInBC
                    End If
                     NK1 tdau, tcuoi, txtShTk(3).Text, ChkDu(2).Value, CboNK(1).ListIndex + 1
                    GoTo KhongInBC
                Case 10:
                    InCDBanCo tdau, tcuoi
                    GoTo KhongInBC
                Case 8:
                    If txtShTk(3).tag = 0 Then
                        ErrMsg er_SHTaiKhoan
                        RFocus txtShTk(3)
                        GoTo KhongInBC
                    End If
                    taikhoan.InitTaikhoanMaSo txtShTk(3).tag
                    InSoDuTK taikhoan, tcuoi
                Case 12:
                    If CboNK(0).ListIndex = 3 Or CboNK(0).ListIndex = 9 Then
                        If txtShTk(3).tag = 0 Then
                            ErrMsg er_SHTaiKhoan
                            RFocus txtShTk(3)
                            GoTo KhongInBC
                        End If
                    End If
                    Select Case CboNK(0).ListIndex
                        Case 0:   NK1 tdau, tcuoi, "111", 1, CboNK(1).ListIndex + 1, 1
                        Case 1:   NK1 tdau, tcuoi, "112", 0, CboNK(1).ListIndex + 1, 2
                        Case 2:   NK1 tdau, tcuoi, "113", 0, CboNK(1).ListIndex + 1, 3
                        Case 3:   NK4 tdau, tcuoi, txtShTk(3).Text, CboNK(1).ListIndex + 1
                        Case 4:   NK5 tdau, tcuoi, "331", CboNK(1).ListIndex + 1
                        Case 5:   NK6 tdau, tcuoi, "151", CboNK(1).ListIndex + 1
                        Case 6:   NK7 tdau, tcuoi
                        Case 7:   NK8 tdau, tcuoi
                        Case 8:   NK9 tdau, tcuoi, CboNK(1).ListIndex + 1
                        Case 9:   NK10 tdau, tcuoi, txtShTk(3).Text, 0, CboNK(1).ListIndex + 1
                    End Select
                    GoTo KhongInBC
                Case 13:
                    If CboNK(0).ListIndex = 5 Then
                        If txtShTk(3).tag = 0 Then
                            ErrMsg er_SHTaiKhoan
                            RFocus txtShTk(3)
                            GoTo KhongInBC
                        End If
                    End If
                    Select Case CboNK(0).ListIndex
                        Case 0:   BK1 tdau, tcuoi, "111", 1, CboNK(1).ListIndex + 1, 1
                        Case 1:   BK1 tdau, tcuoi, "112", 0, CboNK(1).ListIndex + 1, 2
                        Case 2:   GiaNhapKho tdau, tcuoi
                                          GoTo InBC
                        Case 3:   BK4 tdau, tcuoi, "SoHieu LIKE '154*' OR SoHieu LIKE '621*' OR SoHieu LIKE '622*' OR SoHieu LIKE '623*' OR SoHieu LIKE '627*'", 4
                        Case 4:   BK4 tdau, tcuoi, "SoHieu LIKE '241*' OR SoHieu LIKE '641*' OR SoHieu LIKE '642*'", 5
                        Case 5:   NK10 tdau, tcuoi, txtShTk(3).Text, 1, CboNK(1).ListIndex + 1
                        Case 10: BK11 tdau, tcuoi, pSHPT, CboNK(1).ListIndex + 1
                    End Select
                    GoTo KhongInBC
                Case 14:
                    CTGhiSo CTGS.ItemData(CTGS.ListIndex), tdau, tcuoi, IIf(OptTG(1).Value, 1, 0), ngay(0), ngay(1), IIf(ChkDu(1).Value, CboVV.ItemData(CboVV.ListIndex), 0)
                Case 21:
                    CTGhiSo2 CTGS.ItemData(CTGS.ListIndex), tdau, tcuoi
                Case 19:
                    For i = 0 To CTGS.ListCount - 1
                        CTGhiSo CTGS.ItemData(i), tdau, tcuoi, IIf(OptTG(1).Value, 1, 0), ngay(0), ngay(1), 0
                    Next
                Case 16:
                    taikhoan.InitTaikhoanMaSo txtShTk(0).tag
                    CTGhiSoTH taikhoan, tdau, tcuoi, IIf(OptTG(1).Value, 1, 0), ngay(0), ngay(1)
                Case 17:
                    If OptQT(0).Value Then InQTThue tdau, tcuoi
                    If OptQT(1).Value Then InQTVAT tdau, tcuoi
                    If OptQT(2).Value Then
                        InQTTTN tdau, tcuoi
                        GoTo KhongInBC
                    End If
                    If OptQT(3).Value Then BKNopThue tdau, tcuoi
                    If OptQT(4).Value Then QTDauVao tdau, tcuoi
                    If OptQT(5).Value Then QTDauRa
                    If OptQT(6).Value Then GTGTCT
            End Select
InBC:
            frmMain.Rpt.Destination = Index
            If Len(NLB) > 0 And NLB <> "..." Then frmMain.Rpt.Formulas(52) = "NLB='" + NLB + "'"
            If Len(KTT) > 0 And KTT <> "..." Then frmMain.Rpt.Formulas(53) = "KTT='" + KTT + "'"
            If Len(GD) > 0 And GD <> "..." Then frmMain.Rpt.Formulas(54) = "GD='" + GD + "'"
a:
            If Not RptOK(frmMain.Rpt.ReportFileName, nn) Then
                MsgBox "MÉu b¸o c¸o ®· bÞ thay ®æi!", vbCritical, App.ProductName
                GoTo KhongInBC
            End If
            frmMain.Rpt.WindowTitle = OptBC(baocao).Caption
         '   frmMain.Rpt.WindowTitle = OptBC(104).Caption
           On Error GoTo LoiIn
            frmMain.Rpt.Action = 1
            On Error GoTo 0
            GoTo KhongInBC
        Case 2:
            Unload Me
    End Select
    GoTo KhongInBC
LoiIn:
    If Err.Number = 20510 Then
        frmMain.Rpt.Formulas(52) = ""
        frmMain.Rpt.Formulas(53) = ""
        frmMain.Rpt.Formulas(54) = ""
        GoTo a
    End If
    MsgBox "Error " + CStr(Err.Number) + ": " + Err.Description, vbExclamation, App.ProductName
KhongInBC:
    Set taikhoan = Nothing
    Set doiung = Nothing
    GauGe.Value = GauGe.Max
    HienThongBao Me.Caption, 1
    Me.MousePointer = 0
End Sub

Private Sub Form_DblClick()
    FThuChi.tag = 10
    FThuChi.GetPhieu NLB, KTT, GD, 0
    SaveSetting IniPath, "Environment", "SReport1", NLB
    SaveSetting IniPath, "Environment", "SReport2", KTT
    SaveSetting IniPath, "Environment", "SReport3", GD
   
End Sub
'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
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
'====================================================================================================
' Khëi t¹o cöa sæ
'====================================================================================================
Public Sub Form_Load()
    Dim chi_so As Integer
    
    baocao = 0
    Pdelay = GetSetting(IniPath, "Environment", "PrinterDelay", 30)
    Caption = Caption + " - " + CStr(pNamTC)
    Int_RecsetToCbo "SELECT DoituongCT.MaSo As F2,(IIF(DoituongCT.MaKhachHang>0,KhachHang.Ten+' - '+DoituongCT.Sohieu+' - ','')+DienGiai+IIF(DoituongCT.MaKhachHang>0,' - '+ Format(NgayKy,'dd/mm/yy'),'')) As F1 FROM DoituongCT LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo ORDER BY  KhachHang.Ten,DoituongCT.SoHieu,DienGiai", CboVV
    
    tkdt1 = GetSetting(IniPath, "VAT", "TK_DTHU1", "51")
    tkdt2 = GetSetting(IniPath, "VAT", "TK_DTHU2", "3387")
    tkdt3 = GetSetting(IniPath, "VAT", "TK_DTHU3", "711")
    tkdt4 = GetSetting(IniPath, "VAT", "TK_DTHU4", "721")
    
    vatr = GetSetting(IniPath, "VAT", "TK_VATR", "3331")
    
    vatv1 = GetSetting(IniPath, "VAT", "TK_VATV1", pVATV + "111")
    vatv2 = GetSetting(IniPath, "VAT", "TK_VATV2", pVATV + "211")
    vatv3 = GetSetting(IniPath, "VAT", "TK_VATV3", pVATV + "12")
    vatv4 = GetSetting(IniPath, "VAT", "TK_VATV4", pVATV + "22")
    
    sodcthue = GetSetting(IniPath, "LastYear", "IncTax" + CStr(pNamTC), 0)
    
    For chi_so = 0 To 1
        AddMonthToCbo CboThang(chi_so)
       CboNK(chi_so).ListIndex = 0
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next
    
    Int_RecsetToCbo "SELECT MaSo As F2,SoHieu+ ' - '+DienGiai As F1 FROM CTGhiSo ORDER BY SoHieu", CTGS
    
    NLB = GetSetting(IniPath, "Environment", "SReport1", "")
    KTT = GetSetting(IniPath, "Environment", "SReport2", "")
    GD = GetSetting(IniPath, "Environment", "SReport3", "")
       
    CTGS.Visible = (pSoKT Mod 100 >= 10)
    OptBC(14).Visible = (pSoKT Mod 100 >= 10)
    OptBC(15).Visible = (pSoKT Mod 100 >= 10)
    OptBC(16).Visible = (pSoKT Mod 100 >= 10)
    OptBC(18).Visible = (pSoKT Mod 100 >= 10)
    OptBC(20).Visible = (pSoKT Mod 100 >= 10)
    OptBC(21).Visible = (pSoKT Mod 100 >= 10)
    
    OptBC(12).Visible = (pSoKT Mod 1000 >= 100)
    OptBC(13).Visible = (pSoKT Mod 1000 >= 100)
    'OptBC(13).Visible
    CboNK(0).Visible = (pSoKT Mod 1000 >= 100)
    Label(3).Visible = (pSoKT Mod 1000 >= 100)
    
    OptBC(1).Visible = (pSoKT Mod 1000 >= 100) Or (pSoKT Mod 10 >= 1)
    OptKqkd(9).Visible = (Len(Dir(pCurDir + "REPORTS\DTCP.XLS")) > 0)
            
    Frame(5).Visible = pSongNgu
    
    If pNN = 1 Then OptNN(1).Value = True
    CboTL.ListIndex = 1
    CboNK(2).ListIndex = 0
    
    If pNoiBo > 0 Then
        OptBC(5).Enabled = False
        OptBC(7).Enabled = False
        OptBC(11).Enabled = False
        OptBC(17).Enabled = False
    End If
        
    Frame(6).Visible = (pSoVV > 0)
    For chi_so = 1 To pSoVV
        ChkTT(chi_so - 1).Visible = True
        CboTT(chi_so - 1).Visible = True
        Int_RecsetToCbo "SELECT MaSo As F2,DienGiai As F1 FROM DoituongCT" + CStr(chi_so) + " ORDER BY DoituongCT" + CStr(chi_so) + ".DienGiai", CboTT(chi_so - 1)
    Next
    
    If pVersion = 3 Then
        OptBC(11).Caption = "C¸c b¸o c¸o vÒ h¹n møc kinh phÝ"
        OptQT(2).Visible = False
    End If
    
    SetFont Me
    
    Frame(2).Visible = True
    OptBc_Click 5
    
    OptBC(1).Visible = False
    
      Dim ngaytt As Date
 ngaytt = "01/01/" + Mid(str(pNamTC), 4, 5)
 MedNgay(0).Text = ngaytt
 MedNgay(1).Text = ngaytt
OptBC(5).Value = True
OptBC(5).BackColor = 8438015

ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD MauSoHD text")
   ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD MauSoHD text")
   ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD LoaiHoaDon text")
   ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD LoaiHoaDon text")
   
    If SelectSQL("SELECT count(maso) AS F1 from HethongTK where SoHieu LIKE '621*'") <= 0 Then
              '   If SelectSQL("SELECT count(maso) as f1 from cdts2005 where maso = 320") > 0 Or SelectSQL("SELECT count(maso) as f1 from cdts2005") <= 0 Then
                    ExecuteSQL5 "drop table cdts2005 "
                    ExecuteSQL5 "drop table KQKD "
                    ExecuteSQL5 "SELECT * INTO cdts2005 FROM [MS Access;PWD=1234;DATABASE=" + pCurDir + "\REPORTS\bc.rpt].cdts2005"
                    ExecuteSQL5 "SELECT * INTO KQKD FROM [MS Access;PWD=1234;DATABASE=" + pCurDir + "\REPORTS\bc.rpt].KQKD"
               ' End If
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopPrint = True
    
    Set xlapp = Nothing
    Set xlsheet = Nothing
End Sub
'======================================================================================
' Thñ tôc in b¸o c¸o thuÕ
'======================================================================================
Private Sub InThue(tdau As Integer, tcuoi As Integer)
    Dim rs_thue As Recordset, rs As Recordset, i As Integer, th As String, mtk As Long
    Dim dkc As Double, NOP As Double, danop As Double, noplk As Double, danoplk As Double, dn As Double
    
    ExecuteSQL5 "DELETE * FROM Thue WHERE MaSo>20 AND MaSo<30"
    Set rs_thue = DBKetoan.OpenRecordset("SELECT MaSo,Ten FROM HethongTK WHERE SoHieu LIKE '3338*' AND Cap=3", dbOpenSnapshot)
    Do While Not rs_thue.EOF
        i = i + 1
        ExecuteSQL5 "INSERT INTO Thue (MaSo,DienGiai,MaTK,TongHop) VALUES (" + CStr(20 + i) + ",'  " + rs_thue!Ten + "'," + CStr(rs_thue!MaSo) + ",1)"
        rs_thue.MoveNext
    Loop
    rs_thue.Close
    
    Set rs_thue = DBKetoan.OpenRecordset("SELECT DISTINCTROW Thue.MaSo As MaThue, Thue.MaTk, Thue.KyTruocSang, " _
        & " Thue.KyNay, Thue.DaNop, Thue.TongHop, HeThongTK.* FROM Thue LEFT JOIN HeThongTK ON Thue.MaTk" _
        & " = HeThongTK.MaSo WHERE (Thue.MaTk>0)", dbOpenSnapshot)
    If rs_thue.RecordCount > 0 Then
        rs_thue.MoveLast
        GauGe.Max = Fix(0.5 + 1.5 * rs_thue.RecordCount)
    End If
    Do While Not rs_thue.BOF
        If Not IsNull(rs_thue!sohieu) Then
                dn = dn + rs_thue!DuCo_0 - rs_thue!DuNo_0
                dkc = rs_thue!DuCo_0 - rs_thue!DuNo_0
                noplk = 0
                danoplk = 0
                For i = 1 To CThangDB(ThangTruoc(tdau))
                    th = CStr(i)
                    dkc = dkc + rs_thue.Fields("Co_" + th) - rs_thue.Fields("No_" + th)
                    noplk = noplk + rs_thue.Fields("Co_" + th)
                    danoplk = danoplk + rs_thue.Fields("No_" + th)
                Next
                NOP = 0
                danop = 0
                For i = CThangDB(tdau) To CThangDB(tcuoi)
                    th = CStr(i)
                    NOP = NOP + rs_thue.Fields("Co_" + th)
                    danop = danop + rs_thue.Fields("No_" + th)
                    noplk = noplk + rs_thue.Fields("Co_" + th)
                    danoplk = danoplk + rs_thue.Fields("No_" + th)
                Next
                ExecuteSQL5 "UPDATE Thue SET KyTruocSang = " + DoiDau(dkc) + ", KyNay = " + DoiDau(NOP) + ", DaNop = " + DoiDau(danop) _
                    + ", LuyKePN = " + DoiDau(noplk) + ", LuyKeDN = " + DoiDau(danoplk) + " WHERE MaSo=" + CStr(rs_thue!MaThue)
                If Left(rs_thue!sohieu, 5) = "33311" Then
                    NOP = PSDu("11", pVATV, tdau, tcuoi, , 1)
                    danop = SoPSTK("333112", tdau, tcuoi, 1)
                    ExecuteSQL5 "UPDATE Thue SET DaHoan = " + DoiDau(NOP) + ", MienGiam = " + DoiDau(danop) + " WHERE MaSo=" + CStr(rs_thue!MaThue)
                End If
                If Left(rs_thue!sohieu, 5) = "33312" Then
                    danop = SoPSTK("333122", tdau, tcuoi, 1)
                    ExecuteSQL5 "UPDATE Thue SET MienGiam = " + DoiDau(danop) + " WHERE MaSo=" + CStr(rs_thue!MaThue)
                End If
        Else
                ExecuteSQL5 "UPDATE Thue SET KyTruocSang = 0, KyNay = 0, DaNop = 0, LuyKePN = 0, LuyKeDN = 0 WHERE MaSo=" + CStr(rs_thue!MaThue)
        End If
        GauGe.Value = GauGe.Value + 1
        rs_thue.MovePrevious
    Loop
            
    ExecuteSQL5 "UPDATE Thue SET KyTruocSang=KyTruocSang+" + DoiDau(sodcthue) + " WHERE MaSo=15"
        
    If pNamTC > 2000 Then
        NOP = -PSDu("33311", pVATV, tdau, tcuoi)
        noplk = -PSDu("33311", pVATV, pThangDauKy, tcuoi)
                
        dkc = SelectSQL("SELECT DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + "-DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + " AS F1 FROM HethongTK WHERE SoHieu='33311'")
        NOP = NOP + SelectSQL("SELECT SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F1 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi))
        noplk = noplk + SelectSQL("SELECT SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F1 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", pThangDauKy, tcuoi))
        
        Set rs_thue = DBKetoan.OpenRecordset("SELECT * FROM HethongTK WHERE SoHieu='" + pVATV + "'", dbOpenSnapshot)
        dkc = dkc - rs_thue.Fields("DuNo_" + CStr(CThangDB(ThangTruoc(tdau))))
        If dkc < 0 Then dkc = 0
        danop = SelectSQL("Select SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",SoPS,0)) AS F1, SUM(SoPS) AS F2 FROM " + ChungTu2TKNC(0) + " WHERE " _
            + WThang("ThangCT", 0, tcuoi) + " AND HethongTK.SoHieu LIKE '33311*' AND (TK.SoHieu LIKE '11*') AND MaLoai<>8", danoplk)
               
        If NOP < 0 Then
            NOP = 0
            danop = dkc
        End If
        If noplk < 0 Then
            noplk = 0
            danoplk = dkc
        End If
        
        ExecuteSQL5 "UPDATE Thue SET KyTruocSang = " + DoiDau(dkc) + ", KyNay = " _
            + DoiDau(NOP) + ", DaNop = " + DoiDau(danop) + ", LuyKePN = " + DoiDau(noplk) + ", LuyKeDN = " + DoiDau(danoplk) + " WHERE Ma=2"
    End If
    rs_thue.Close
    
    ' Céng dån
    Set rs_thue = DBKetoan.OpenRecordset("SELECT Ma FROM Thue WHERE Tonghop=0", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_thue.EOF
        Set rs = DBKetoan.OpenRecordset("SELECT SUM(KyTruocSang) as SoKyTruoc,SUM(KyNay) as SoKyNay,SUM(DaNop) as SoDaNop,SUM(LuyKePN) as SoPhaiNop,SUM(LuyKeDN) as SoLKDN FROM Thue WHERE Tonghop = " + CStr(rs_thue!ma), dbOpenSnapshot)
        ExecuteSQL5 "UPDATE Thue SET KyTruocSang = " + DoiDau(rs!SoKyTruoc) + ", KyNay = " _
            + DoiDau(rs!SoKyNay) + ", DaNop = " + DoiDau(rs!SoDaNop) + ", LuyKePN = " + DoiDau(rs!SoPhaiNop) + ", LuyKeDN = " + DoiDau(rs!SoLKDN) + " WHERE Ma=" + CStr(rs_thue!ma)
        rs.Close
        GauGe.Value = GauGe.Value + 1
        rs_thue.MoveNext
    Loop
    rs_thue.Close
    
    tentk GetSetting(IniPath, "TAX", "TNDN", "3334"), mtk
    SoDuTK mtk, 0, danop, NOP
    
    frmMain.Rpt.ReportFileName = "THUE.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThueNT = " + DoiDau(dn)
    frmMain.Rpt.Formulas(4) = "ThueTNNT = " + DoiDau(NOP - danop)
    
    Set rs_thue = Nothing
    Set rs = Nothing
    GauGe.Value = GauGe.Max
End Sub
'======================================================================================
' Thñ tôc in Tµi s¶n ngoµi b¶ng
'======================================================================================
Public Sub InTsNb(thang As Integer)
    Dim st As String
    
    GauGe.Max = 2
    st = CStr(CThangDB(thang))
    SetSQL "QTaisanNB", "SELECT DISTINCTROW Ten" + IIf(nn > 0, "E", "") + ", SoHieu, DuNo_0, DuCo_0, DuNo_" + st + " AS DuNo, DuCo_" + st + " AS DuCo" _
    & " From HeThongTK Where ((HeThongTK.cap = 1) And (HeThongTK.Loai = 0)) ORDER BY HethongTK.SoHieu"
    GauGe.Value = 1
    frmMain.Rpt.ReportFileName = "TSNB.RPT"
    RptSetDate NgayCuoiThang(pNamTC, thang), nn
    'frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(thang, thang, nn) + "'"
    GauGe.Value = 2
End Sub

'====================================================================================================
' Thñ tôc in nhËt ký chung
'====================================================================================================
Private Sub OptBc_Click(Index As Integer)
    baocao = Index
    Frame(0).Visible = (Index = 3 And pVersion <> 3)
    Frame(1).Visible = ((Index = 4 Or Index = 7 Or Index = 11) And pVersion <> 3)
    Frame(2).Visible = (Index = 5)
    Frame(3).Visible = (Index = 17)
    Frame(6).Visible = ((Index = 0 Or Index = 1) And pSoVV > 0)
    Frame(7).Visible = (Index = 11 And pVersion = 3)
    Select Case baocao
        Case 0, 1, 14, 16, 18, 20:
            OptTG(1).Enabled = True
            MedNgay(0).Enabled = True
            MedNgay(1).Enabled = True
        Case Else
            OptTG(0).Value = True
            OptTG(1).Enabled = False
            MedNgay(0).Enabled = False
            MedNgay(1).Enabled = False
    End Select
End Sub

Private Sub OptBC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If pRpt = 1 And Button = 2 And User_Right = 0 Then
        FU1.QuyenBCTongHop Index, OptBC(Index).Caption
    End If
    

      
End Sub

Private Sub OptBC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i

        OptBC(5).BackColor = &H80FF80 '&HC0FFC0    '&H80000003
        OptBC(3).BackColor = &H80FF80 '&HC0FFC0    '&H80000003
        OptBC(4).BackColor = &H80FF80 '&HC0FFC0    '&H80000003
        OptBC(17).BackColor = &H80FF80 '&HC0FFC0    '&H80000003

If (Index = 5 Or Index = 3 Or Index = 4 Or Index = 17) Then
 OptBC(Index).BackColor = 8438015
 End If
End Sub

Private Sub OptNN_Click(Index As Integer)
    nn = Index
    If CboTL.ListIndex < 0 Then Exit Sub
    SetFont Me, 1
End Sub

Public Sub txtShTk_GotFocus(Index As Integer)
    AutoSelect txtShTk(Index)
End Sub

Private Sub txtShTk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdtk_Click Index
End Sub

Public Sub txtShTk_LostFocus(Index As Integer)
    Dim mtk As Long
    
    txtShTk(Index).Text = UCase(txtShTk(Index).Text)
    If Len(txtShTk(Index).Text) > 0 Then
        mtk = 0
        LbTenTk(Index).Caption = tentk(txtShTk(Index).Text, mtk)
        txtShTk(Index).tag = mtk
    Else
        LbTenTk(Index).Caption = ""
        txtShTk(Index).tag = 0
    End If
    HienThongBao Me.Caption, 1
End Sub
'======================================================================================
' Hµm tr¶ vÒ sè d­ tµi kho¶n cuèi th¸ng
'======================================================================================
Private Sub SoDuTKTC(mtk As Long, thang As Integer, dnno As Double, dnco As Double, duno As Double, duco As Double)
    Dim rs_tk As Recordset, st As String
    
    st = CStr(CThangDB(thang))
    Set rs_tk = DBKetoan.OpenRecordset("SELECT DuNo_0, DuCo_0, DuNo_" + st + " As DuNo, DuCo_" + st _
        + " As DuCo FROM HethongTK WHERE MaSo=" + CStr(mtk), dbOpenSnapshot)
    dnno = rs_tk!DuNo_0
    dnco = rs_tk!DuCo_0
    duno = rs_tk!duno
    duco = rs_tk!duco
    rs_tk.Close
    Set rs_tk = Nothing
End Sub
'======================================================================================
' Thñ tôc in b¸o c¸o VAT phan III
'======================================================================================
Private Sub InVAT(tdau As Integer, tcuoi As Integer)
    Dim sql As String, kn As Double, lk As Double, CK As Double, cklk As Double, TK As New ClsTaikhoan
    Dim dk As Double, duno As Double, lk1 As Double, duco As Double, X As Double, Y As Double, lk2 As Double
    
    GauGe.Max = 10
    
    ExecuteSQL5 "UPDATE VAT SET KyNay = 0,LuyKe=0"
    
    TK.InitTaikhoanSohieu pVATV
    TK.SoDuTK ThangTruoc(tdau), duno, duco, X
    kn = duno - duco
    TK.SoDuTK 0, duno, duco, X
    lk = duno - duco
    TK.InitTaikhoanSohieu vatr
    TK.SoDuTK ThangTruoc(tdau), duno, duco, X
    kn = kn + duno - duco
    TK.SoDuTK 0, duno, duco, X
    lk = lk + duno - duco
    lk2 = -kn
    
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(IIf(kn > 0, kn, 0)) + ",LuyKe = " + DoiDau(IIf(lk > 0, lk, 0)) + " WHERE MaSo = '10'"
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(IIf(kn < 0, -kn, 0)) + ",LuyKe = " + DoiDau(IIf(lk < 0, -lk, 0)) + " WHERE MaSo = '40'"
    
    VATHoanLai tdau, tcuoi, dk, duno, lk1, duco, X
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(dk) + " WHERE MaSo = '20'"
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(duno) + ", LuyKe = " + DoiDau(lk1) + " WHERE MaSo = '21'"
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(duco) + ", LuyKe = " + DoiDau(X) + " WHERE MaSo = '22'"
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(dk + duno - duco) + " WHERE MaSo = '23'"
    
    VATMienGiam tdau, tcuoi, dk, duno, lk1, duco, X
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(dk) + " WHERE MaSo = '30'"
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(duno) + ", LuyKe = " + DoiDau(lk1) + " WHERE MaSo = '31'"
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(duco) + ", LuyKe = " + DoiDau(X) + " WHERE MaSo = '32'"
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(dk + duno - duco) + " WHERE MaSo = '33'"
    
    GauGe.Value = 1
    SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKNC(-1) + " WHERE (HethongTK.SoHieu LIKE '521*' OR HethongTK.SoHieu LIKE '531*' OR HethongTK.SoHieu LIKE '532*') AND " + WThang("ThangCT", 0, tcuoi)
    
    sql = "SELECT SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS),0)) AS F1, SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM (" + ChungTu2TKHD(1) + ") LEFT JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT WHERE IsNull(MienTru.MaCT) AND HoaDon.Loai=1 AND " + WThang("ThangCT", 0, tcuoi) + " AND KCT=0 AND DC=0"
    kn = SelectSQL(sql, lk)
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(kn) + ",LuyKe = " + DoiDau(lk) + " WHERE MaSo = '41'"
    lk2 = lk2 + kn
    
    
    If TK.tk_id = GTGTKT_ID Then
        sql = "SELECT SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",SoPS,0)) AS F1,SUM(SoPS) AS F2 FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT " _
            & " WHERE IsNull(MienTru.MaCT) AND (HethongTK.SoHieu LIKE '33312*') AND " + WThang("ThangCT", 0, tcuoi) + " AND (TK.SoHieu LIKE '11*' OR TK.SoHieu LIKE '136*')"
        X = SelectSQL(sql, Y) - PSDu(pVATV, "33312", tdau, tcuoi)
        Y = Y - PSDu(pVATV, "33312", pThangDauKy, tcuoi)
    Else
        X = 0
        Y = 0
    End If
    
    sql = "SELECT SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",SoPS,0)) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND " + WThang("ThangCT", 0, tcuoi)
    kn = SelectSQL(sql, lk)
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(kn - X) + ",LuyKe = " + DoiDau(lk - Y) + " WHERE MaSo = '11'"
    lk2 = lk2 - (kn - X)
    
    sql = "SELECT SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",SoPS,0)) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKNC(0) _
        & " WHERE (HethongTK.SoHieu LIKE '142*' OR HethongTK.SoHieu LIKE '242*' OR HethongTK.SoHieu LIKE '6*' OR HethongTK.SoHieu LIKE '421*') AND (TK.SoHieu LIKE '" + pVATV + "*') AND " + WThang("ThangCT", 0, tcuoi)
    kn = SelectSQL(sql, lk)
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(kn) + ",LuyKe = " + DoiDau(lk) + " WHERE MaSo = '16'"
    CK = kn
    cklk = lk
    sql = "SELECT SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",SoPS,0)) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKNC(0) _
        & " WHERE (HethongTK.SoHieu LIKE '3331*') AND (TK.SoHieu LIKE '" + pVATV + "*') AND " + WThang("ThangCT", 0, tcuoi)
    kn = SelectSQL(sql, lk)
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(kn) + ",LuyKe = " + DoiDau(lk) + " WHERE MaSo = '13' OR MaSo = '42'"
    CK = CK + kn
    cklk = cklk + lk
    sql = "SELECT SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",SoPS,0)) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKNC(0) _
        & " WHERE (HethongTK.SoHieu LIKE '11*' OR HethongTK.SoHieu LIKE '138*') AND (TK.SoHieu LIKE '" + pVATV + "112*' OR TK.SoHieu LIKE '" + pVATV + "212*') AND " + WThang("ThangCT", 0, tcuoi)
    kn = SelectSQL(sql, lk)
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(kn) + ",LuyKe = " + DoiDau(lk) + " WHERE MaSo = '14'"
    CK = CK + kn
    cklk = cklk + lk
    sql = "SELECT SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",SoPS,0)) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKNC(0) _
        & " WHERE (TK.SoHieu LIKE '133111*') AND (HethongTK.SoHieu LIKE '331*' OR HethongTK.SoHieu LIKE '111*') AND " + WThang("ThangCT", 0, tcuoi)
    'sql = "SELECT SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",SoPS,0)) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKNC(0) _
        & " WHERE (HethongTK.SoHieu LIKE '3331*') AND (TK.SoHieu LIKE '" + pSHPT + "*' OR TK.SoHieu LIKE '111*') AND " + WThang("ThangCT", 0, tcuoi)
    kn = SelectSQL(sql, lk)
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(kn) + ",LuyKe = " + DoiDau(lk) + " WHERE MaSo = '15'"
    CK = CK + kn
    cklk = cklk + lk
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(CK) + ",LuyKe = " + DoiDau(cklk) + " WHERE MaSo = '12'"
    'lk2 = lk2 + CK
    
    sql = "SELECT SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",SoPS,0)) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKNC(0) _
        & " WHERE (HethongTK.SoHieu LIKE '3331*') AND (TK.SoHieu LIKE '" + pSHPT + "*' OR TK.SoHieu LIKE '111*') AND " + WThang("ThangCT", 0, tcuoi)
    kn = SelectSQL(sql, lk)
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(kn) + ",LuyKe = " + DoiDau(lk) + " WHERE MaSo = '43'"
    
    lk2 = lk2 - lk
    
    sql = "SELECT SUM(IIF(" + WThang("ThangCT", tdau, 0) + ",SoPS,0)) AS F1,SUM(SoPS) AS F2 FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT" _
        & " WHERE IsNull(MienTru.MaCT) AND (HethongTK.SoHieu LIKE '33311*') AND (TK.SoHieu LIKE '11*') AND " + WThang("ThangCT", 0, tcuoi)
    kn = SelectSQL(sql, lk)
    ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(kn) + ",LuyKe = " + DoiDau(lk) + " WHERE MaSo = '45'"
    lk2 = lk2 - kn
    
    If lk2 >= 0 Then
        ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(lk2) + ",LuyKe = " + DoiDau(lk2) + " WHERE MaSo = '46'"
    Else
        ExecuteSQL5 "UPDATE VAT SET KyNay = " + DoiDau(-lk2) + ",LuyKe = " + DoiDau(-lk2) + " WHERE MaSo = '17'"
    End If
    
    sql = "SELECT VAT.MaSo, First(VAT.DienGiai" + IIf(nn > 0, "E", "") + ") AS FirstOfDienGiai, First(VAT.TongHop) AS FirstOfTongHop, Sum(VAT.KyNay) AS SumOfKyNay, Sum(VAT.LuyKe) AS SumOfLuyKe From VAT GROUP BY VAT.MaSo"
    SetSQL "QVAT", sql
    
    Set TK = Nothing
    frmMain.Rpt.ReportFileName = "VAT.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
End Sub

Private Function InVATHoanLai(tdau As Integer, tcuoi As Integer) As Boolean
Dim shtkhl1 As String, shtkhl2 As String
Dim dkno As Double, psno As Double, psco As Double
Dim taikhoan As New ClsTaikhoan
Dim dkn As Double, dkc As Double, dknt As Double, psn As Double, psc As Double
    GauGe.Max = 5
    shtkhl1 = GetSetting(IniPath, "VAT", "TK_HOANLAI", pVATV + "112")
    shtkhl2 = GetSetting(IniPath, "VAT", "TK_HOANLAI", pVATV + "212")
    
    taikhoan.InitTaikhoanSohieu shtkhl1
    taikhoan.SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
    dkno = dkn - dkc
    GauGe.Value = 1
    
    taikhoan.SoPhatSinh tdau, tcuoi, psn, psc
    psno = psn
    psco = psc
    GauGe.Value = 2
    
    taikhoan.InitTaikhoanSohieu shtkhl2
    taikhoan.SoDuTK tdau, dkn, dkc, dknt
    dkno = dkno + dkn - dkc
    GauGe.Value = 3
    
    taikhoan.SoPhatSinh tdau, tcuoi, psn, psc
    psno = psno + psn
    psco = psco + psc
    GauGe.Value = 4
    
    If dkno = 0 And psno = 0 And psco = 0 Then GoTo KhongIn
    
    If psno <> 0 Or psco <> 0 Then
        SetSQL "QSocai", "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.GhiChu, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.MaTKTCNo, ChungTu.MaTKTCCo" _
            & " FROM HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo" _
            & " Where ((HethongTK.SoHieu LIKE '" + shtkhl1 + "*' OR HethongTK.SoHieu LIKE '" + shtkhl2 + "*'" _
            + ") Or (HethongTK_1.SoHieu LIKE '" + shtkhl1 + "*' OR HethongTK_1.SoHieu LIKE '" + shtkhl2 + "*'" _
            + ")) And " + WThang("ThangCT", tdau, tcuoi) + " ORDER BY ChungTu.NgayGS, ChungTu.MaCT"
    Else
        SetSQL "QSocai", "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.GhiChu, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.MaTKTCNo, ChungTu.MaTKTCCo" _
            & " FROM HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo" _
            & " Where (chungtu.MaCT = 0)"
            frmMain.Rpt.Formulas(8) = "DuCk=" + DoiDau(dkno)
    End If
    GauGe.Value = 5
    frmMain.Rpt.Formulas(3) = "SoHieuTK1 = '" + shtkhl1 + "'"
    frmMain.Rpt.Formulas(4) = "SoHieuTK2 = '" + shtkhl2 + "'"
    frmMain.Rpt.Formulas(5) = "NoDk=" + DoiDau(dkno)
    frmMain.Rpt.Formulas(6) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.ReportFileName = "VATHL.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    InVATHoanLai = True
    GoTo KT
KhongIn:
    InVATHoanLai = False
    ErrMsg er_KoPS1
KT:
    Set taikhoan = Nothing
End Function

Private Function InVATMienGiam(tdau As Integer, tcuoi As Integer) As Boolean
Dim shtkmg1 As String, shtkmg2 As String, shtkdg1 As String, shtkdg2 As String, shtktntc As String, dkco As Double, sql As String
Dim taikhoan As New ClsTaikhoan, dkn As Double, dkc As Double, dknt As Double, psn As Double, psc As Double
    
    GauGe.Max = 3
    shtkmg1 = GetSetting(IniPath, "VAT", "TK_DAURA", "333111")
    shtkmg2 = GetSetting(IniPath, "VAT", "TK_NK", "333121")
    shtkdg1 = GetSetting(IniPath, "VAT", "TK_GIAM", "333112")
    shtkdg2 = GetSetting(IniPath, "VAT", "TK_NKGIAM", "333122")
    shtktntc = GetSetting(IniPath, "VAT", "TK_TNTC", "721")
    
    taikhoan.InitTaikhoanSohieu shtkdg1
    taikhoan.SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
    dkco = dkc - dkn
    GauGe.Value = 1
    
    taikhoan.InitTaikhoanSohieu shtkdg2
    taikhoan.SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
    dkco = dkco + dkc - dkn
    GauGe.Value = 2
    
    sql = "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.GhiChu, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.MaTKTCNo, ChungTu.MaTKTCCo" _
            & " FROM HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo" _
            & " Where ((HethongTK.SoHieu LIKE '" + shtkmg1 + "*' AND HethongTK_1.SoHieu LIKE '" + shtkdg1 + "*'" _
            + ") Or (HethongTK.SoHieu LIKE '" + shtkmg2 + "*' AND HethongTK_1.SoHieu LIKE '" + shtkdg2 + "*'" _
            + ") Or ((HethongTK.SoHieu LIKE '" + shtkdg1 + "*' OR HethongTK.SoHieu LIKE '" + shtkdg2 + "*') AND HethongTK_1.SoHieu LIKE '" + shtktntc + "*'" _
            + ")) And " + WThang("ThangCT", tdau, tcuoi) + " ORDER BY ChungTu.NgayGS, ChungTu.MaCT"
    
    If GetRowNumber(sql) Then
        SetSQL "QSocai", sql
    Else
        If dkco = 0 Then GoTo KhongIn
        SetSQL "QSocai", "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.GhiChu, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.MaTKTCNo, ChungTu.MaTKTCCo" _
            & " FROM HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo" _
            & " Where (chungtu.MaCT = 0)"
            frmMain.Rpt.Formulas(8) = "DuCk=" + DoiDau(dkco)
    End If
    GauGe.Value = 3
    
    frmMain.Rpt.Formulas(3) = "SoHieuTK1 = '" + shtkmg1 + "'"
    frmMain.Rpt.Formulas(4) = "SoHieuTK2 = '" + shtkmg2 + "'"
    frmMain.Rpt.Formulas(5) = "NoDk=" + DoiDau(dkco)
    frmMain.Rpt.Formulas(6) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.ReportFileName = "VATMG.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    InVATMienGiam = True
    GoTo KT
KhongIn:
    InVATMienGiam = False
    ErrMsg er_KoPS1
KT:
    Set taikhoan = Nothing
End Function
'====================================================================================================
' Thñ tôc in bang ke hoa don vao
'====================================================================================================
Private Function InVATDauVao2(tdau As Integer, tcuoi As Integer, tl As Integer, HD As Integer, TK As ClsTaikhoan) As Boolean
    Dim sql As String, tk2 As New ClsTaikhoan, Fx As Integer
    
    Fx = IIf(Left(frmMain.LbCty(8).Caption, 2) = "35" Or Left(frmMain.LbCty(3).Caption, 2) = "64", 35, 0)
    
    tk2.InitTaikhoanSohieu "33312"
    If tk2.tk_id = GTGTKT_ID Then
        sql = "DELETE HoaDon.*  FROM ((HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo " _
            & " WHERE (HethongTK.SoHieu LIKE '" + pVATV + "*') AND (TK.SoHieu LIKE '33312*'  OR InStr(ChungTu.GhiChu,'33312')>0) AND HoaDon.Loai=-1"
        ExecuteSQL5 sql
    End If
    
    GauGe.Max = 1
    'ChungTu.MauSoHD as
    If TK.MaSo = 0 Then
        Select Case tl
            Case -3:
                If Fx = 0 Then
                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HDBL=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                Else
                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HDBL=1 AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                End If
            Case -2:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
            Case -1:
                If Fx = 0 Then
                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                Else
                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HDBL=0 OR KCT=0) AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                End If
            Case 30:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND (TyLe=5 OR TyLe=10) AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
            Case Else
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND TyLe=" + CStr(tl) + " AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HDBL=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
        End Select
        frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    Else
        Select Case tl
            Case -3:
                If Fx = 0 Then
                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HDBL=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                Else
                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HDBL=1 AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                End If
            Case -2:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
            Case -1:
                If Fx = 0 Then
                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                Else
                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HDBL=0 OR KCT=0) AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                End If
            Case 30:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND (TyLe=5 OR TyLe=10) AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
            Case Else
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND TyLe=" + CStr(tl) + " AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HDBL=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
        End Select
'    If TK.MaSo = 0 Then
'        Select Case tl
'            Case -3:
'                If Fx = 0 Then
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HDBL=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                Else
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HDBL=1 AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                End If
'            Case -2:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'            Case -1:
'                If Fx = 0 Then
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                Else
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HDBL=0 OR KCT=0) AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                End If
'            Case 30:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND (TyLe=5 OR TyLe=10) AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'            Case Else
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND TyLe=" + CStr(tl) + " AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HDBL=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'        End Select
'        frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
'    Else
'        Select Case tl
'            Case -3:
'                If Fx = 0 Then
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND HDBL=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                Else
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND HDBL=1 AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                End If
'            Case -2:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'            Case -1:
'                If Fx = 0 Then
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                Else
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HDBL=0 OR KCT=0) AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                End If
'            Case 30:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND (TyLe=5 OR TyLe=10) AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'            Case Else
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND TyLe=" + CStr(tl) + " AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HDBL=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'        End Select
        frmMain.Rpt.Formulas(1) = "TenCn='" + TK.Ten + "'"
        frmMain.Rpt.Formulas(5) = "DiaChi = '" + TK.GhiChu + "'"
    End If

    GauGe.Value = 1
    If HD = 0 Then
        frmMain.Rpt.ReportFileName = "BANGKEV3.RPT"
    Else
        If (tl > 0 And tl < 5) Or tl = -3 Then
            frmMain.Rpt.ReportFileName = "BANGKEV4.RPT"
        Else
            If GetSetting(IniPath, "Invoice", "ListDetail", 0) = 1 Then
            
                SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKHD(0) + " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HoaDon.Loai=-1 GROUP BY MaCT"
                SetSQL "QCdt", "SELECT MaCT FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaVattu>0 AND MaLoai=1 GROUP BY MaCT HAVING Count(MaVattu)>1"
                SetSQL "QDuPhong", "SELECT ChungTu.MaCT,TenVattu,SoPS2No AS SL,SoPS AS TT FROM (((ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN QCdt ON ChungTu.MaCT=QCdt.MaCT WHERE (TK_ID=" + CStr(TKVT_ID) + " OR HethongTK.Loai=6)"
                SetSQL "QChitiet", "SELECT QNhatKy.*,TenVattu,SL,TT FROM QNhatKy LEFT JOIN QDuPhong ON QNhatKy.MaCT=QDuPhong.MaCT ORDER BY TyLe,NgayPH,SoHD"
            
                'Dim rs As Recordset
                '
               ' SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKHD(0) + " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HoaDon.Loai=-1 GROUP BY MaCT"
                'ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
               ' SetSQL "QCdt", "SELECT MaCT FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaVattu>0 AND MaLoai=1 GROUP BY MaCT HAVING Count(MaVattu)>1"
               ' ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,MaCha,Ten,Kq1,Kq2,MK) SELECT ChungTu.MaSo,Cstr(ChungTu.MaSo),ChungTu.MaCT,TenVattu,SoPS2No,SoPS,MaVattu FROM (((ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN QCdt ON ChungTu.MaCT=QCdt.MaCT WHERE (TK_ID=" + CStr(TKVT_ID) + " OR HethongTK.Loai=6) ORDER BY ChungTu.MaCT,ChungTu.MaSo"
               ' Set rs = DBKetoan.OpenRecordset("SELECT MaCT,ChungTu.MaSo FROM " + ChungTu2TKHD(0) + " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HoaDon.Loai=-1 ORDER BY ChungTu.MaSo", dbOpenSnapshot)
               ' Do While Not rs.EOF
               '     ExecuteSQL5 "UPDATE BaoCaoCP2 SET BC_ID=" + CStr(rs!MaSo) + " WHERE BC_ID=0 AND MaCha=" + CStr(rs!MaCT) '+ " AND MaSo>=" + CStr(rs!MaSo)
               '     rs.MoveNext
               ' Loop
               ' rs.Close
               ' Set rs = Nothing
              '
              '  SetSQL "QDuPhong", "SELECT BC_ID,First(Ten) AS TenVattu,Sum(Kq1) AS SL,Sum(Kq2) AS TT FROM BaoCaoCP2 GROUP BY BC_ID,MK"
              '  SetSQL "QChitiet", "SELECT QNhatKy.*,TenVattu,SL,TT FROM QNhatKy LEFT JOIN QDuPhong ON QNhatKy.MaSo=QDuPhong.BC_ID ORDER BY TyLe,NgayPH,MaSo"
                frmMain.Rpt.ReportFileName = "BANGKEV6.RPT"
            Else
                frmMain.Rpt.ReportFileName = "BANGKEV2.RPT"
            End If
        End If
    End If
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(6) = "MSThue = '" + frmMain.LbCty(8).Caption + "'"
    frmMain.Rpt.Formulas(7) = "TLKT = " + CStr(tl)
    InVATDauVao2 = True
    Set tk2 = Nothing
End Function
'====================================================================================================
' Thñ tôc in bang ke hoa don ra
'====================================================================================================
Private Function InVATDauRa(tdau As Integer, tcuoi As Integer, tl As Integer, TK As ClsTaikhoan) As Boolean
    Dim sql As String

    GauGe.Max = 1
'    If TK.MaSo = 0 Then
'        Select Case tl
'            Case -3:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -2:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -1:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case 30:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case Else
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'        End Select
'        frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
'    Else
'        Select Case tl
'            Case -3:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -2:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -1:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case 30:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case Else
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'        End Select
'        frmMain.Rpt.Formulas(1) = "TenCn='" + TK.Ten + "'"
'        frmMain.Rpt.Formulas(5) = "DiaChi = '" + TK.GhiChu + "'"
'    End If
  If TK.MaSo = 0 Then
        Select Case tl
            Case -3:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,ChungTu.MauSoHD as Diachi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case -2:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,ChungTu.MauSoHD as DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case -1:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,ChungTu.MauSoHD as DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case 30:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,ChungTu.MauSoHD as DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case Else
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,ChungTu.MauSoHD as DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
        End Select
        frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    Else
        Select Case tl
            Case -3:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case -2:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case -1:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case 30:
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case Else
                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
        End Select
        frmMain.Rpt.Formulas(1) = "TenCn='" + TK.Ten + "'"
        frmMain.Rpt.Formulas(5) = "DiaChi = '" + TK.GhiChu + "'"
    End If
    GauGe.Value = 1
    If GetSetting(IniPath, "Invoice", "ListDetail", 0) = 1 Then
        SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKHD(0) + " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HoaDon.Loai=1 GROUP BY MaCT"
        SetSQL "QCdt", "SELECT MaCT FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaVattu>0 AND MaLoai=8 GROUP BY MaCT HAVING Count(MaVattu)>1"
        SetSQL "QDuPhong", "SELECT ChungTu.MaCT,TenVattu AS MH,SoPS2Co AS SL,SoPS AS TT FROM (((ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN QCdt ON ChungTu.MaCT=QCdt.MaCT WHERE TK_ID=" + CStr(TKDT_ID)
        SetSQL "QChitiet", "SELECT QNhatKy.*,QDuPhong.* FROM QNhatKy LEFT JOIN QDuPhong ON QNhatKy.MaCT=QDuPhong.MaCT ORDER BY TyLe,NgayPH,SoHD"
        frmMain.Rpt.ReportFileName = "BANGKER6.RPT"
    Else
        frmMain.Rpt.ReportFileName = "BANGKER.RPT"
    End If
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(6) = "MSThue = '" + frmMain.LbCty(8).Caption + "'"
    frmMain.Rpt.Formulas(8) = "TLKT = " + CStr(tl)
    InVATDauRa = True
End Function
'==================================================
'In bang ke dau ra bang exel
'==================================================
'Private Function InVATDauRa_exel(tdau As Integer, tcuoi As Integer, tl As Integer, TK As ClsTaikhoan) As Boolean
'    Dim sql As String
'
'    GauGe.Max = 1
'    If TK.MaSo = 0 Then
'        Select Case tl
'            Case -3:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -2:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -1:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case 30:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case Else
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'        End Select
'        frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
'    Else
'        Select Case tl
'            Case -3:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -2:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -1:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case 30:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case Else
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'        End Select
'        frmMain.Rpt.Formulas(1) = "TenCn='" + TK.Ten + "'"
'        frmMain.Rpt.Formulas(5) = "DiaChi = '" + TK.GhiChu + "'"
'    End If
'    GauGe.Value = 1
'    If GetSetting(IniPath, "Invoice", "ListDetail", 0) = 1 Then
'        SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKHD(0) + " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HoaDon.Loai=1 GROUP BY MaCT"
'        SetSQL "QCdt", "SELECT MaCT FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaVattu>0 AND MaLoai=8 GROUP BY MaCT HAVING Count(MaVattu)>1"
'        SetSQL "QDuPhong", "SELECT ChungTu.MaCT,TenVattu AS MH,SoPS2Co AS SL,SoPS AS TT FROM (((ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN QCdt ON ChungTu.MaCT=QCdt.MaCT WHERE TK_ID=" + CStr(TKDT_ID)
'        SetSQL "QChitiet", "SELECT QNhatKy.*,QDuPhong.* FROM QNhatKy LEFT JOIN QDuPhong ON QNhatKy.MaCT=QDuPhong.MaCT ORDER BY TyLe,NgayPH,SoHD"
'        frmMain.Rpt.ReportFileName = "BANGKER6.RPT"
'    Else
'        frmMain.Rpt.ReportFileName = "BANGKER.RPT"
'    End If
'    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
'    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
'    frmMain.Rpt.Formulas(6) = "MSThue = '" + frmMain.LbCty(8).Caption + "'"
'    frmMain.Rpt.Formulas(8) = "TLKT = " + CStr(tl)
'    InVATDauRa = True
'End Function
'
'================================================================
'================================================================
Private Sub VATHoanLai(tdau As Integer, tcuoi As Integer, dk As Double, hoanlai As Double, lkhl As Double, dahoanlai As Double, lkdhl As Double)
Dim shtkhl1 As String, shtkhl2 As String
Dim taikhoan As New ClsTaikhoan
Dim dkn As Double, dkc As Double, dknt As Double, psn As Double, psc As Double
    shtkhl1 = GetSetting(IniPath, "VAT", "TK_HOANLAI", pVATV + "112")
    shtkhl2 = GetSetting(IniPath, "VAT", "TK_HOANLAI", pVATV + "212")
    
    taikhoan.InitTaikhoanSohieu shtkhl1
    taikhoan.SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
    dk = dkn - dkc
    
    taikhoan.SoPhatSinh tdau, tcuoi, psn, psc
    hoanlai = psn
    dahoanlai = psc
    
    taikhoan.SoPhatSinh pThangDauKy, tcuoi, psn, psc
    lkhl = psn
    lkdhl = psc
    
    taikhoan.InitTaikhoanSohieu shtkhl2
    taikhoan.SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
    dk = dk + dkn - dkc
    
    taikhoan.SoPhatSinh tdau, tcuoi, psn, psc
    hoanlai = hoanlai + psn
    dahoanlai = dahoanlai + psc
    taikhoan.SoPhatSinh pThangDauKy, tcuoi, psn, psc
    lkhl = lkhl + psn
    lkdhl = lkdhl + psc
    
    Set taikhoan = Nothing
End Sub

Private Sub VATMienGiam(tdau As Integer, tcuoi As Integer, dk As Double, GIAM As Double, lkg As Double, dagiam As Double, lkdg As Double)
Dim shtkmg1 As String, shtkmg2 As String, shtkdg1 As String, shtkdg2 As String, shtktntc As String, sql As String
Dim taikhoan As New ClsTaikhoan, dkn As Double, dkc As Double, dknt As Double

    shtkmg1 = GetSetting(IniPath, "VAT", "TK_DAURA", "333111")
    shtkmg2 = GetSetting(IniPath, "VAT", "TK_NK", "333121")
    shtkdg1 = GetSetting(IniPath, "VAT", "TK_GIAM", "333112")
    shtkdg2 = GetSetting(IniPath, "VAT", "TK_NKGIAM", "333122")
    shtktntc = GetSetting(IniPath, "VAT", "TK_TNTC", "721")
    
    taikhoan.InitTaikhoanSohieu shtkdg1
    taikhoan.SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
    dk = dkc - dkn
    taikhoan.InitTaikhoanSohieu shtkdg2
    taikhoan.SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
    dk = dk + dkc - dkn
    
    sql = "SELECT DISTINCTROW Sum(IIF(" + WThang("ThangCT", tdau, 0) + ",ChungTu.SoPS,0)) As F1, Sum(ChungTu.SoPS) As F2" _
            & " FROM HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo" _
            & " Where ((HethongTK.SoHieu LIKE '" + shtkmg1 + "*' AND HethongTK_1.SoHieu LIKE '" + shtkdg1 + "*'" _
            + ") Or (HethongTK.SoHieu LIKE '" + shtkmg2 + "*' AND HethongTK_1.SoHieu LIKE '" + shtkdg2 + "*'" _
            + ")) And " + WThang("ThangCT", 0, tcuoi)
    GIAM = SelectSQL(sql, lkg)
    
    sql = "SELECT DISTINCTROW Sum(IIF(" + WThang("ThangCT", tdau, 0) + ",ChungTu.SoPS,0)) As F1, Sum(ChungTu.SoPS) AS F2" _
            & " FROM HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo" _
            & " Where ((HethongTK.SoHieu LIKE '" + shtkdg1 + "*' OR HethongTK.SoHieu LIKE '" + shtkdg2 + "*') AND (HethongTK_1.SoHieu LIKE '" + shtktntc + "*'" _
            + ")) And " + WThang("ThangCT", 0, tcuoi)
    dagiam = SelectSQL(sql, lkdg)
    Set taikhoan = Nothing
End Sub
'====================================================================================================
' Thu tuc in chi tiet tai khoan chi phi
'====================================================================================================
Private Sub InCTChiphi(tdau As Integer, tcuoi As Integer, shtkcp As String, dg As String)
    Dim sqln As String, sqlnx As String, i As Integer
    Dim RsTk As Recordset
    
    GauGe.Max = 3
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqln = sqln + " + No_" + CStr(i)
    Next
    For i = 1 To CThangDB(tcuoi)
        sqlnx = sqlnx + " + No_" + CStr(i)
    Next
    GauGe.Value = 1
    ExecuteSQL5 "UPDATE HethongTK SET KC_N = 0, KC_C = 0"
    Set RsTk = DBKetoan.OpenRecordset("SELECT DISTINCTROW HethongTK.MaSo, SUM(IIF(" + WThang("ThangCT", tdau, tcuoi) + ",ChungTu.SoPS,0)) AS GTru1, SUM(IIF(" + WThang("ThangCT", 0, tcuoi) + ",ChungTu.SoPS,0)) AS GTru2 " _
        & " FROM HethongTK INNER JOIN ChungTu ON HethongTK.MaSo = ChungTu.MaTkCo WHERE HethongTK.SoHieu LIKE '" + shtkcp + "*' AND ChungTu.MaLoai < 3 GROUP BY HethongTK.MaSo")
    Do While Not RsTk.EOF
        ExecuteSQL5 "UPDATE HethongTK SET KC_N = " + DoiDau(RsTk!GTru1) + ", KC_C = " + DoiDau(RsTk!GTru2) + " WHERE MaSo = " + CStr(RsTk!MaSo)
        RsTk.MoveNext
    Loop
    Set RsTk = DBKetoan.OpenRecordset("SELECT DISTINCTROW HethongTK.MaSo, SUM(IIF(" + WThang("ThangCT", tdau, tcuoi) + ",ChungTu.SoPS,0)) AS GTru1, SUM(IIF(" + WThang("ThangCT", 0, tcuoi) + ",ChungTu.SoPS,0)) AS GTru2 " _
        & " FROM (HethongTK INNER JOIN ChungTu ON HethongTK.MaSo = ChungTu.MaTkNo) INNER JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo WHERE HethongTK.SoHieu LIKE '" + shtkcp + "*' AND ChungTu.MaLoai = 3 AND LEFT(TK.SoHieu,1)='9' GROUP BY HethongTK.MaSo")
    Do While Not RsTk.EOF
        ExecuteSQL5 "UPDATE HethongTK SET KC_N = KC_N + " + DoiDau(RsTk!GTru1) + ", KC_C = KC_C + " + DoiDau(RsTk!GTru2) + " WHERE MaSo = " + CStr(RsTk!MaSo)
        RsTk.MoveNext
    Loop
    RsTk.Close
    Set RsTk = Nothing
    
    CongDonKC False
    SetSQL "QTongHopCT", "SELECT DISTINCTROW SoHieu, First(Cap) AS FirstOfCap, First(Ten" + IIf(nn > 0, "E", "") + ") AS FirstOfTen, First(Kieu) AS FirstOfKieu, First(Loai) AS FirstOfLoai, First(MaTC = MaSo) AS FirstOfChuanLa, First(CapDuoi) AS CD," _
        & " Sum(" + sqln + " - KC_N) AS PsNo, Sum(" + sqlnx + " - KC_C) AS PsNoX" _
        & " From HeThongTK GROUP BY SoHieu Having (SoHieu LIKE '" + shtkcp + "*' AND SoHieu <> '" + shtkcp + "')"
    
    GauGe.Value = 2
    XDCapTK shtkcp
    frmMain.Rpt.ReportFileName = "THCPQL.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(5) = "TenBaoCao = IF PageNumber()=1 THEN '" + dg + "'"
    GauGe.Value = 3
End Sub
'====================================================================================================
' Xac dinh cac phat sinh do ket chuyen
'====================================================================================================
Public Sub XacdinhKCCP(tdau As Integer, tcuoi As Integer)
      Dim rs_taikhoan As Recordset, sql As String
     
      ExecuteSQL5 "UPDATE HethongTK SET KC_N = 0, KC_C = 0"
               
      SetSQL "QCdt", "SELECT DISTINCTROW ChungTu.MaTkNo, SUM(SoPS) AS KC_N FROM ChungTu WHERE (ChungTu.MaLoai = 3) AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY ChungTu.MaTkNo"
            
      SetSQL "QChitiet", "SELECT DISTINCTROW ChungTu.MaTkCo, SUM(SoPS) AS KC_C FROM ChungTu WHERE (ChungTu.MaLoai = 3) AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY ChungTu.MaTkCo"
            
      Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT DISTINCTROW  QCdt.* FROM QCdt", dbOpenSnapshot, dbForwardOnly)
      
      Do While Not rs_taikhoan.EOF
            sql = "UPDATE HethongTK SET KC_N = " + DoiDau(rs_taikhoan!KC_N) + " WHERE MaSo = " + CStr(rs_taikhoan!MaTkNo)
            ExecuteSQL5 sql
            rs_taikhoan.MoveNext
      Loop
            
      Set rs_taikhoan = DBKetoan.OpenRecordset("SELECT DISTINCTROW  QChitiet.* FROM QChitiet", dbOpenSnapshot, dbForwardOnly)
      
      Do While Not rs_taikhoan.EOF
            sql = "UPDATE HethongTK SET KC_C = " + DoiDau(rs_taikhoan!KC_C) + " WHERE MaSo = " + CStr(rs_taikhoan!MaTkCo)
            ExecuteSQL5 sql
            rs_taikhoan.MoveNext
      Loop
      
    rs_taikhoan.Close
    Set rs_taikhoan = Nothing
End Sub
'====================================================================================================
' Thu tuc in so du cac chi tiet cua mot tai khoan tai khoan
'====================================================================================================
Private Sub InSoDuTK(taikhoan As ClsTaikhoan, thang As Integer)
    Dim st As String, sqln As String, sqlc As String, i As Integer
    Dim sqlnx As String, sqlcx As String
    Dim ckn As Double, ckc As Double, cknt As Double
    
    GauGe.Max = 2
    st = CStr(CThangDB(thang))
    SetSQL "QTongHopCT", "SELECT DISTINCTROW SoHieu, First(Cap) AS FirstOfCap, First(Ten" + IIf(nn > 0, "E", "") + ") AS FirstOfTen, First(Kieu) AS FirstOfKieu, First(Loai) AS FirstOfLoai, First(MaTC = MaSo OR MaTC = 0) AS FirstOfChuanLa," _
        & " Sum(DuNo_" + st + ") AS CkNo, Sum(DuCo_" + st + ") AS CkCo, First(CapDuoi) AS CD" _
        & " From HeThongTK GROUP BY SoHieu Having (SoHieu LIKE '" + taikhoan.sohieu + "*' AND SoHieu <> '" + taikhoan.sohieu + "')"
    
    GauGe.Value = 1
    XDCapTK taikhoan.sohieu
    frmMain.Rpt.ReportFileName = "SoDuTK.RPT"
    RptSetDate NgayCuoiThang(pNamTC, thang), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(thang, thang, nn) + "'"
    If taikhoan.MaSo > 0 Then
        frmMain.Rpt.Formulas(5) = "TaiKhoan = IF PageNumber() = 1 THEN '" + IIf(nn > 0, ABCtoVNI("Account: ") + taikhoan.sohieu + " - " + taikhoan.TenE, ABCtoVNI("Tµi kho¶n: ") + taikhoan.sohieu + " - " + taikhoan.Ten) + "'"
    End If
    taikhoan.SoDuTK thang, ckn, ckc, cknt
    frmMain.Rpt.Formulas(6) = "DuNoCK = " + DoiDau(ckn)
    frmMain.Rpt.Formulas(7) = "DuCoCK = " + DoiDau(ckc)
    GauGe.Value = 2
End Sub

Private Sub InTMTC(tdau As Integer, tcuoi As Integer)
    Dim i As Integer, j As Integer, ps As Double, ps2 As Double, Verx As Integer, fn As String
      fn = pCurDir + "TMBCTC" + IIf(tdau < 10, "0", "") + CStr(tdau) + IIf(tcuoi < 10, "0", "") + CStr(tcuoi) + CStr(pNamTC Mod 100) + ".XLS"
   
   ' fn = pCurDir + IIf(tdau < 10, "0", "") + CStr(tdau) + IIf(tcuoi < 10, "0", "") + CStr(tcuoi) + CStr(pNamTC Mod 100) + ".XLS"
    'fn = "D:\TMBCTC.XLS"
    If Len(Dir(fn)) > 0 Then
        If MsgBox("ThuyÕt minh tµi chÝnh ®· ®­îc lËp lËp t¹i " + fn + ", cã cÇn lËp l¹i ?", vbCritical + vbYesNo, App.ProductName) = vbNo Then
            GoTo MoTep
        Else
            Recycle fn
        End If
    End If
    
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\TMTC.XLS", fn
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open fn
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(5, 1) = ThoiGian(tdau, tcuoi, nn)
    
    InCTKQKD tdau, tcuoi
    ps = PSDu("642", "334", tdau, tcuoi) + PSDu("642", "3382", tdau, tcuoi) + PSDu("642", "3383", tdau, tcuoi)
    xlsheet.Cells(39, 2) = SelectSQL("SELECT Sum(KyNay) AS F1 FROM KQKDCT WHERE MaSo=14")
    xlsheet.Cells(40, 2) = SelectSQL("SELECT Sum(KyNay) AS F1 FROM KQKDCT WHERE MaSo=15") + ps
    xlsheet.Cells(41, 2) = SelectSQL("SELECT Sum(KyNay) AS F1 FROM KQKDCT WHERE MaSo=13")
    ps = 0
    j = 9
    Do While Len(xlsheet.Cells(42, j)) > 0
        ps = ps + SoPSTK(xlsheet.Cells(42, j), tdau, tcuoi, -1)
        j = j + 1
    Loop
    xlsheet.Cells(42, 2) = ps
    xlsheet.Cells(44, 2) = SelectSQL("SELECT Sum(KyNay) AS F1 FROM KQKDCT WHERE MaSo=12")
    
    ' Hang ton kho
    For i = 58 To 60
        j = 9
        ps = 0
        Do While Len(xlsheet.Cells(i, j)) > 0
            ps = ps + SoDuTKSH(xlsheet.Cells(i, j), tcuoi)
            j = j + 1
        Loop
        xlsheet.Cells(i, 2) = ps
    Next
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        CapNhatGiaTriTaiSan i, GauGe
    Next
    
    ' Tinh hinh tang giam TSCD
    xlsheet.Cells(67, 2) = SoDuTKSH("2112", ThangTruoc(tdau))
    xlsheet.Cells(67, 3) = SoDuTKSH("2113", ThangTruoc(tdau))
    xlsheet.Cells(67, 4) = SoDuTKSH("2114", ThangTruoc(tdau))
    xlsheet.Cells(67, 5) = SoDuTKSH("2115", ThangTruoc(tdau))
    xlsheet.Cells(67, 6) = SoDuTKSH("2116", ThangTruoc(tdau))
    xlsheet.Cells(67, 7) = SoDuTKSH("2118", ThangTruoc(tdau))
    
    xlsheet.Cells(68, 2) = SoPSTK("2112", tdau, tcuoi, -1)
    xlsheet.Cells(68, 3) = SoPSTK("2113", tdau, tcuoi, -1)
    xlsheet.Cells(68, 4) = SoPSTK("2114", tdau, tcuoi, -1)
    xlsheet.Cells(68, 5) = SoPSTK("2115", tdau, tcuoi, -1)
    xlsheet.Cells(68, 6) = SoPSTK("2116", tdau, tcuoi, -1)
    xlsheet.Cells(68, 7) = SoPSTK("2118", tdau, tcuoi, -1)
    
    xlsheet.Cells(70, 2) = SoTangGiamTS("2112", tdau, tcuoi, 41)
    xlsheet.Cells(70, 3) = SoTangGiamTS("2113", tdau, tcuoi, 41)
    xlsheet.Cells(70, 4) = SoTangGiamTS("2114", tdau, tcuoi, 41)
    xlsheet.Cells(70, 5) = SoTangGiamTS("2115", tdau, tcuoi, 41)
    xlsheet.Cells(70, 6) = SoTangGiamTS("2116", tdau, tcuoi, 41)
    xlsheet.Cells(70, 7) = SoTangGiamTS("2118", tdau, tcuoi, 41)
    
    xlsheet.Cells(72, 2) = SoPSTK("2112", tdau, tcuoi, 1)
    xlsheet.Cells(72, 3) = SoPSTK("2113", tdau, tcuoi, 1)
    xlsheet.Cells(72, 4) = SoPSTK("2114", tdau, tcuoi, 1)
    xlsheet.Cells(72, 5) = SoPSTK("2115", tdau, tcuoi, 1)
    xlsheet.Cells(72, 6) = SoPSTK("2116", tdau, tcuoi, 1)
    xlsheet.Cells(72, 7) = SoPSTK("2118", tdau, tcuoi, 1)
    
    xlsheet.Cells(74, 2) = SoTangGiamTS("2112", tdau, tcuoi, 48)
    xlsheet.Cells(74, 3) = SoTangGiamTS("2113", tdau, tcuoi, 48)
    xlsheet.Cells(74, 4) = SoTangGiamTS("2114", tdau, tcuoi, 48)
    xlsheet.Cells(74, 5) = SoTangGiamTS("2115", tdau, tcuoi, 48)
    xlsheet.Cells(74, 6) = SoTangGiamTS("2116", tdau, tcuoi, 48)
    xlsheet.Cells(74, 7) = SoTangGiamTS("2118", tdau, tcuoi, 48)
    
    xlsheet.Cells(79, 2) = NGHetKH("2112", tcuoi)
    xlsheet.Cells(79, 3) = NGHetKH("2113", tcuoi)
    xlsheet.Cells(79, 4) = NGHetKH("2114", tcuoi)
    xlsheet.Cells(79, 5) = NGHetKH("2115", tcuoi)
    xlsheet.Cells(79, 6) = NGHetKH("2116", tcuoi)
    xlsheet.Cells(79, 7) = NGHetKH("2118", tcuoi)
    
    xlsheet.Cells(82, 2) = GTHaoMon("2112", ThangTruoc(tdau))
    xlsheet.Cells(82, 3) = GTHaoMon("2113", ThangTruoc(tdau))
    xlsheet.Cells(82, 4) = GTHaoMon("2114", ThangTruoc(tdau))
    xlsheet.Cells(82, 5) = GTHaoMon("2115", ThangTruoc(tdau))
    xlsheet.Cells(82, 6) = GTHaoMon("2116", ThangTruoc(tdau))
    xlsheet.Cells(82, 7) = GTHaoMon("2118", ThangTruoc(tdau))
    
    xlsheet.Cells(83, 2) = SoKHTS("2112", tdau, tcuoi)
    xlsheet.Cells(83, 3) = SoKHTS("2113", tdau, tcuoi)
    xlsheet.Cells(83, 4) = SoKHTS("2114", tdau, tcuoi)
    xlsheet.Cells(83, 5) = SoKHTS("2115", tdau, tcuoi)
    xlsheet.Cells(83, 6) = SoKHTS("2116", tdau, tcuoi)
    xlsheet.Cells(83, 7) = SoKHTS("2118", tdau, tcuoi)
    
    xlsheet.Cells(85, 2) = GTHaoMon("2112", tcuoi)
    xlsheet.Cells(85, 3) = GTHaoMon("2113", tcuoi)
    xlsheet.Cells(85, 4) = GTHaoMon("2114", tcuoi)
    xlsheet.Cells(85, 5) = GTHaoMon("2115", tcuoi)
    xlsheet.Cells(85, 6) = GTHaoMon("2116", tcuoi)
    xlsheet.Cells(85, 7) = GTHaoMon("2118", tcuoi)
    
    xlsheet.Cells(99, 3) = SoPSTK("334", tdau, tcuoi, 1)
    
    'Tinh hinh tang giam nguon von CSH
    For i = 112 To 139
        j = 9
        If Len(xlsheet.Cells(i, j)) > 0 Then
            xlsheet.Cells(i, 2) = SoDuTKSH(xlsheet.Cells(i, j), ThangTruoc(tdau))
            xlsheet.Cells(i, 3) = SoPSTK(xlsheet.Cells(i, j), tdau, tcuoi, 1)
            xlsheet.Cells(i, 4) = SoPSTK(xlsheet.Cells(i, j), tdau, tcuoi, -1)
        End If
    Next
    
    ' Tinh hinh tang giam dau tu
    For i = 148 To 154
        j = 9
        If Len(xlsheet.Cells(i, j)) > 0 Then
            xlsheet.Cells(i, 2) = SoDuTKSH(xlsheet.Cells(i, j), ThangTruoc(tdau))
            xlsheet.Cells(i, 3) = SoPSTK(xlsheet.Cells(i, j), tdau, tcuoi, -1)
            xlsheet.Cells(i, 4) = SoPSTK(xlsheet.Cells(i, j), tdau, tcuoi, 1)
        End If
    Next
    InCdts pThangDauKy, tcuoi, 0
    ' Cac khoan phai thu va phai tra
    For i = 165 To 174
        j = 9
        Do While Len(xlsheet.Cells(i, j)) > 0
            xlsheet.Cells(i, 2) = xlsheet.Cells(i, 2) + SoDuTKSH(xlsheet.Cells(i, j), ThangTruoc(tdau), -1) - SoDuTKSH(xlsheet.Cells(i, j), ThangTruoc(tdau), 1)
            xlsheet.Cells(i, 4) = xlsheet.Cells(i, 4) + SoPSTK(xlsheet.Cells(i, j), tdau, tcuoi, -1)
            xlsheet.Cells(i, 5) = xlsheet.Cells(i, 5) + SoPSTK(xlsheet.Cells(i, j), tdau, tcuoi, 1)
            j = j + 1
        Loop
    Next
    
    For i = 176 To 191
        j = 9
        Do While Len(xlsheet.Cells(i, j)) > 0
            xlsheet.Cells(i, 2) = xlsheet.Cells(i, 2) + SoDuTKSH(xlsheet.Cells(i, j), ThangTruoc(tdau), 1) - SoDuTKSH(xlsheet.Cells(i, j), ThangTruoc(tdau), -1)
            xlsheet.Cells(i, 5) = xlsheet.Cells(i, 5) + SoPSTK(xlsheet.Cells(i, j), tdau, tcuoi, -1)
            xlsheet.Cells(i, 4) = xlsheet.Cells(i, 4) + SoPSTK(xlsheet.Cells(i, j), tdau, tcuoi, 1)
            j = j + 1
        Loop
    Next
        
    GauGe.Value = 0
    
    ps = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=100 OR MaSo=200")
    If ps > 0 Then xlsheet.Cells(223, 4) = 100 * SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=200") / ps
    ps = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=300 OR MaSo=400")
    ps2 = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=300")
    If ps > 0 Then xlsheet.Cells(226, 4) = 100 * ps2 / ps
    
    If ps2 > 0 Then xlsheet.Cells(229, 4) = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=100 OR MaSo=200") / ps2
    ps2 = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=310")
    If ps2 > 0 Then
        xlsheet.Cells(230, 4) = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=100") / ps2
        xlsheet.Cells(231, 4) = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=110 OR MaSo=120") / ps2
    End If
    ps2 = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=320")
    If ps2 > 0 Then xlsheet.Cells(232, 4) = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=100") / ps2
        
    InKqkd tdau, tcuoi, 0, 0
    ps = SoPSTK("511", tdau, tcuoi, 1)
    ps = ps + SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=21") + SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=31")
    If ps = 0 Then GoTo KT
    
    If SelectSQL("SELECT Tag AS F1 FROM License") = "Z" Then Verx = 1 Else Verx = 0
    
    xlsheet.Cells(235, 4) = 100 * SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=50") / ps
    xlsheet.Cells(236, 4) = 100 * SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=" + IIf(Verx = 0, "60", "70")) / ps

    ps2 = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=100 OR MaSo=200")
    If ps2 > 0 Then
        xlsheet.Cells(238, 4) = 100 * SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=50") / ps2
        xlsheet.Cells(239, 4) = 100 * SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=" + IIf(Verx = 0, "60", "70")) / ps2
    End If
    
    ps2 = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=410")
    If ps2 > 0 Then xlsheet.Cells(240, 4) = 100 * SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=" + IIf(Verx = 0, "60", "70")) / ps2
KT:
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
MoTep:
    CallExcel IIf(tdau < 10, "0", "") + CStr(tdau) + IIf(tcuoi < 10, "0", "") + CStr(tcuoi) + CStr(pNamTC Mod 100) + ".XLS"
KetThuc:
 MsgBox ("ThuyÕt minh tµi chÝnh ®· lËp t¹i:" + pCurDir + "TMBCTC3" + IIf(tdau < 10, "0", "") + CStr(tdau) + IIf(tcuoi < 10, "0", "") + CStr(tcuoi) + CStr(pNamTC Mod 100) + ".XLS")

End Sub

Private Sub InLCTT2(tdau As Integer, tcuoi As Integer)
    Dim sql As String, KT As Double, kn As Double
    Dim rs_lailo  As Recordset
    
    GauGe.Max = 4
    ExecuteSQL5 "UPDATE LCTT SET KyTruoc = 0, KyNay = 0"
    Set rs_lailo = DBKetoan.OpenRecordset("SELECT DISTINCTROW * FROM LCTT WHERE TKNo<>'0' AND TKCo<>'0' ORDER BY MaSo", dbOpenSnapshot)
    Do While Not rs_lailo.EOF
        sql = "SELECT DISTINCTROW Sum(IIF(" + WThang2("ThangCT", 0, tdau) + ",ChungTu.SoPS,0)) AS F1, Sum(IIF(" + WThang("ThangCT", tdau, 0) + ",ChungTu.SoPS,0)) AS F2 " _
            & "FROM (HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) INNER JOIN HeThongTK AS HeThongTK_1 ON ChungTu.MaTKCo = HeThongTK_1.MaSo " _
            & "WHERE HethongTK.SoHieu LIKE '" + rs_lailo!tkno + "*' AND HethongTK_1.SoHieu LIKE '" + rs_lailo!TkCo + "*' AND " + WThang("ThangCT", 0, tcuoi)
        KT = SelectSQL(sql, kn)

        If rs_lailo!dau > 0 Then
            ExecuteSQL5 "UPDATE LCTT SET KyTruoc=" + DoiDau(KT) + ", KyNay = " + DoiDau(kn) + " WHERE Ma=" + CStr(rs_lailo!ma)
        Else
            ExecuteSQL5 "UPDATE LCTT SET KyTruoc=" + DoiDau(-KT) + ", KyNay = " + DoiDau(-kn) + " WHERE Ma=" + CStr(rs_lailo!ma)
        End If
        rs_lailo.MoveNext
    Loop
    
    GauGe.Value = 2
    KT = SoPSTK("111", pThangDauKy, ThangTruoc(tdau), -1) + SoPSTK("112", pThangDauKy, ThangTruoc(tdau), -1) - SoPSTK("113", pThangDauKy, ThangTruoc(tdau), 1) - PSDu("111", "3364", pThangDauKy, ThangTruoc(tdau)) _
        - SelectSQL("SELECT Sum(KyTruoc) AS F1 FROM LCTT WHERE MaSo=1 OR MaSo=22 OR MaSo=24 OR MaSo=26 OR MaSo=27 OR MaSo=31 OR MaSo=33")
    kn = SoPSTK("111", tdau, tcuoi, -1) + SoPSTK("112", tdau, tcuoi, -1) - SoPSTK("113", tdau, tcuoi, 1) - PSDu("111", "3364", tdau, tcuoi) _
        - SelectSQL("SELECT Sum(KyNay) AS F1 FROM LCTT WHERE MaSo=1 OR MaSo=22 OR MaSo=24 OR MaSo=26 OR MaSo=27 OR MaSo=31 OR MaSo=33")
    ExecuteSQL5 "UPDATE LCTT SET KyTruoc=" + DoiDau(KT) + ", KyNay = " + DoiDau(kn) + " WHERE MaSo=6"
    
    KT = SoPSTK("111", pThangDauKy, ThangTruoc(tdau), 1) + SoPSTK("112", pThangDauKy, ThangTruoc(tdau), 1) - SoPSTK("113", pThangDauKy, ThangTruoc(tdau), -1) - PSDu("111", "3364", pThangDauKy, ThangTruoc(tdau)) _
        + SelectSQL("SELECT Sum(KyTruoc) AS F1 FROM LCTT WHERE MaSo=2 OR MaSo=3 OR MaSo=4 OR MaSo=5 OR MaSo=21 OR MaSo=23 OR MaSo=25 OR MaSo=32 OR MaSo=34 OR MaSo=35 OR MaSo=36")
    kn = SoPSTK("111", tdau, tcuoi, 1) + SoPSTK("112", tdau, tcuoi, 1) - SoPSTK("113", tdau, tcuoi, -1) - PSDu("111", "3364", tdau, tcuoi) _
        + SelectSQL("SELECT Sum(KyNay) AS F1 FROM LCTT WHERE MaSo=2 OR MaSo=3 OR MaSo=4 OR MaSo=5 OR MaSo=21 OR MaSo=23 OR MaSo=25 OR MaSo=32 OR MaSo=34 OR MaSo=35 OR MaSo=36")
    ExecuteSQL5 "UPDATE LCTT SET KyTruoc=" + DoiDau(-KT) + ", KyNay = " + DoiDau(-kn) + " WHERE MaSo=7"
    
    sql = "SELECT SUM(DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + "-DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS F1,SUM(DuNo_0-DuCo_0) AS F2 FROM HethongTK WHERE SoHieu='111' OR SoHieu='112'"
    kn = SelectSQL(sql, KT)
    ExecuteSQL5 "UPDATE LCTT SET KyTruoc = " + DoiDau(IIf(tdau > pThangDauKy, KT, 0)) + ", KyNay = " + DoiDau(kn) + " WHERE MaSo=60"
    'sql = "SELECT SUM(DuNo_" + CStr(CThangDB(tcuoi)) + "-DuCo_" + CStr(CThangDB(tcuoi)) + ") AS F1,SUM(DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + "-DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS F2 FROM HethongTK WHERE SoHieu='1111' OR SoHieu='1121'"
    'kn = SelectSQL(sql, KT)
    'ExecuteSQL5 "UPDATE LCTT SET KyTruoc = " + CStr(IIf(tdau > pThangDauKy, KT, 0)) + ", KyNay = " + CStr(kn) + " WHERE MaSo=70"
    GauGe.Value = 1
    
    ' Cong don
    Set rs_lailo = DBKetoan.OpenRecordset("SELECT KyTruoc,KyNay,MaSoCha FROM LCTT WHERE TongHop=0 And MaSoCha<>0 ORDER BY MaSo", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_lailo.EOF
         If rs_lailo!MaSoCha > 0 Then
                 ExecuteSQL5 "UPDATE LCTT SET KyTruoc=KyTruoc+" + DoiDau(rs_lailo!KyTruoc) + ",KyNay=KyNay+" + DoiDau(rs_lailo!KyNay) + " WHERE MaSo=" + CStr(rs_lailo!MaSoCha)
        Else
                 ExecuteSQL5 "UPDATE LCTT SET KyTruoc=KyTruoc-" + DoiDau(rs_lailo!KyTruoc) + ",KyNay=KyNay-" + DoiDau(rs_lailo!KyNay) + " WHERE MaSo=" + CStr(-rs_lailo!MaSoCha)
        End If
        rs_lailo.MoveNext
    Loop
    
    Set rs_lailo = DBKetoan.OpenRecordset("SELECT KyTruoc,KyNay,MaSoCha FROM LCTT WHERE TongHop>0 And MaSoCha<>0 ORDER BY MaSo", dbOpenDynaset, dbForwardOnly)
    Do While Not rs_lailo.EOF
        If rs_lailo!MaSoCha > 0 Then
            ExecuteSQL5 "UPDATE LCTT SET KyTruoc=KyTruoc+" + str(rs_lailo.Fields("KyTruoc")) + ",KyNay=KyNay+" + DoiDau(rs_lailo!KyNay) + " WHERE MaSo=" + CStr(rs_lailo!MaSoCha)
        Else
            ExecuteSQL5 "UPDATE LCTT SET KyTruoc=KyTruoc-" + str(rs_lailo.Fields("KyTruoc")) + ",KyNay=KyNay-" + DoiDau(rs_lailo!KyNay) + " WHERE MaSo=" + CStr(-rs_lailo!MaSoCha)
        End If
        rs_lailo.MoveNext
    Loop
    rs_lailo.Close
    Set rs_lailo = Nothing
    GauGe.Value = 4
    
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.ReportFileName = "LCTT.RPT"
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
End Sub
'====================================================================================================
 ' Thu tuc in ban can doi ban co tren EXCEL
'====================================================================================================
Private Sub InCDBanCo(tdau As Integer, tcuoi As Integer)
Dim rs As Recordset, rs2 As Recordset, i As Integer, j As Integer, sql As String
Dim sotkno As Integer, sotkco As Integer, rx As String, cx As String
            
    ' Lay danh sach tai khoan tai chinh
    sql = "SELECT LEFT(HeThongTK.SoHieu,4) AS ShCo FROM " + ChungTu2TKNC(2) _
        & " WHERE MaTKTCNo > 0 AND MaTKTCCo > 0 AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY LEFT(HethongTK.SoHieu,4)"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.RecordCount = 0 Then GoTo KetThuc
    
    Recycle pCurDir + "BANCO.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\BANCO.XLS", pCurDir + "BANCO.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "BANCO.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi)
    
' Liet ke theo hang ngang (phat sinh co)
    i = 1
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(5, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco = i - 1
    
    sql = "SELECT First(HethongTK.MaSo) AS MaTKNo, LEFT(HeThongTK.SoHieu,4) AS ShNo, First(Ten) AS TenTK FROM " + ChungTu2TKNC(-2) _
        & " WHERE MaTKTCNo > 0 AND MaTKTCCo > 0 AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY LEFT(HethongTK.SoHieu,4)"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        rs.MoveFirst
        i = 0
        Do While Not rs.EOF
            HienThongBao VString(rs!shno + " - " + rs!tentk), 1
            i = i + 1
            xlsheet.Cells(i + 5, 1) = rs!shno
            sql = "SELECT LEFT(TK.SoHieu,4) AS ShCo, Sum(ChungTu.SoPS) AS SumOfSoPS" _
                & " FROM " + ChungTu2TKNC(10) _
                & " WHERE (ChungTu.MaLoai<>4 OR Chungtu.MaTKNo<>ChungTu.MaTKCo) AND (" + WThang("ThangCT", tdau, tcuoi) + " And (HethongTK.SoHieu LIKE '" + rs!shno + "*')) GROUP BY LEFT(HeThongTK.SoHieu, 4), LEFT(TK.SoHieu,4)"
    
            Set rs2 = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
            j = 1
            Do While Not rs2.EOF
                j = j + 1
                Do While (rs2!shco <> xlsheet.Cells(5, j)) And (j <= sotkco)
                    j = j + 1
                Loop
                xlsheet.Cells(i + 5, j) = rs2!SumOfSoPS
                rs2.MoveNext
            Loop
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        rs2.Close
        
        sotkno = i
                
        xlsheet.Cells(5, sotkco + 2) = ABCtoVNI("Tæng PS")
        xlsheet.Cells(sotkno + 6, 1) = ABCtoVNI("Tæng PS")
        
        rx = CStr(sotkno + 6)
        cx = XLSCol(sotkco + 2)
        
        sql = XLSCol(sotkco + 1)
        For i = 1 To sotkno
            xlsheet.Range(cx + CStr(i + 5)).Formula = "=SUM(B" + CStr(i + 5) + ":" + sql + CStr(i + 5) + ")"
        Next
        
        For i = 1 To sotkco
            sql = XLSCol(i + 1)
            xlsheet.Range(sql + CStr(sotkno + 6)).Formula = "=SUM(" + sql + "6:" + sql + CStr(sotkno + 5) + ")"
        Next
        xlsheet.Range(cx + rx).Formula = "=SUM(" + cx + "6:" + cx + CStr(sotkno + 5) + ")"
        
        xlsheet.Range("B6", cx + rx).Font.Name = "Arial"
        xlsheet.Range("A" + rx).Font.Name = ".VnArial"
        xlsheet.Range("A" + rx).WrapText = True
        xlsheet.Range(cx + "5").Font.Name = ".VnArial"
        xlsheet.Range(cx + "5").WrapText = True
        xlsheet.Range("A5", cx + "5").HorizontalAlignment = xlHAlignCenter
        xlsheet.Range("A5", "A" + rx).HorizontalAlignment = xlHAlignCenter
        
        xlsheet.Range("A3", cx + "3").MergeCells = True
        xlsheet.Range("A4", cx + "4").MergeCells = True
        
        With xlsheet.Range("A5", cx + rx)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(0, 0, 0)
            .RowHeight = 18
            .Font.Size = 6
            .VerticalAlignment = xlVAlignCenter
        End With
        
        xlapp.Workbooks(1).Save
        
        xlapp.Workbooks.Close
        
        Set xlsheet = Nothing
        Set xlapp = Nothing
        CallExcel "BANCO.XLS"
    Else
        ErrMsg er_KoPS
    End If
KetThuc:
    rs.Close
End Sub

Private Sub NK1(tdau As Integer, tcuoi As Integer, TK As String, nhom As Integer, cap As Integer, Optional id As Integer = 0)
Dim rs As Recordset, i As Integer, j As Integer, sql As String, d As Date, m As Long, r As Integer
Dim sotkno As Integer, sotkco As Integer, rx As String, cx As String, lx As Integer, ofs As Integer

    ofs = IIf(id >= 2 Or (id = 0 And nhom = 0), 2, 0)
    r = IIf(id < 2, 1, 2)
    lx = Len(TK)
    XDSoHieuCap cap
    ' Lay danh sach tai khoan tai chinh
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.RecordCount = 0 Then
        ErrMsg er_KoPS
        GoTo KetThuc
    End If
    
    Recycle pCurDir + "NK" + CStr(r) + ".XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\NK" + CStr(r) + ".XLS", pCurDir + "NK" + CStr(r) + ".XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "NK" + CStr(r) + ".XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = IIf(id > 0, "NhËt ký chøng tõ sè " + CStr(id), "B¶ng kª ph¸t sinh cã tµi kho¶n " + TK)
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    If ofs > 0 Then
        xlsheet.Cells(6, 1) = "STT"
        xlsheet.Cells(6, ofs) = "Sè CT"
        xlsheet.Cells(6, 1 + ofs) = "Ngµy"
        xlsheet.Cells(6, 2 + ofs) = "DiÔn gi¶i"
        xlsheet.Range("B5", "D5").MergeCells = True
    End If
    xlsheet.Cells(5, 3 + ofs) = "Ghi cã tµi kho¶n " + TK + ", ghi nî c¸c tµi kho¶n"
    If nhom = 1 Then xlsheet.Cells(6, 1) = "STT"
    ' Liet ke theo hang ngang (phat sinh co)
    i = 2 + ofs
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(6, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco = i - 2 - ofs

    ' xlSheet.Cells(7, sotkco + 3) = IIf(otk.kieu < 0, dn - dc, dc - dn)
    ' xlsheet.Range("C5","").
    xlsheet.Cells(5, sotkco + 3 + ofs) = "Céng cã"
    
    sql = "SELECT First(Chungtu.MaCT) AS M,First(Chungtu.SoHieu) AS SH,NgayGS,BaoCaoCP.SoHieu AS ShCo,Sum(SoPS) AS PS,First(DienGiai" + IIf(nn > 0, "E", "") + ") AS DG FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu + '*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY NgayGS" + IIf(nhom = 0, ",MaCT", "") + ",BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    d = CVDate("01/01/1900")
    j = 6

    ' Liet ke theo hang doc (phat sinh no)
    rs.MoveLast
    GauGe.Max = rs.RecordCount
    rs.MoveFirst
    Do While Not rs.EOF
        If (rs!NgayGS <> d And nhom = 1) Or (rs!m <> m And nhom = 0) Then
            sotkno = sotkno + 1
            i = 3 + ofs
            j = j + 1
            d = rs!NgayGS
            m = rs!m
            HienThongBao Format(rs!NgayGS, Mask_D), 1
            If id < 2 And id > 0 Then
                xlsheet.Cells(j, 1) = IIf(nhom = 0, rs!sh, CStr(j - 6))
                xlsheet.Cells(j, 2) = rs!NgayGS
            Else
                xlsheet.Cells(j, 1) = CStr(j - 6)
                xlsheet.Cells(j, 2) = rs!sh
                xlsheet.Cells(j, 3) = rs!NgayGS
                xlsheet.Cells(j, 4) = rs!dg
            End If
            'xlSheet.Cells(j, 2) = Format(rs!ngaygs, Mask_D)
        End If
        Do While rs!shco <> Left(xlsheet.Cells(6, i), Len(rs!shco))
            i = i + 1
        Loop
        xlsheet.Cells(j, i) = rs!ps
        GauGe.Value = GauGe.Value + 1
        rs.MoveNext
    Loop
    
    sotkno = j - 6

    xlsheet.Cells(sotkno + 7, 2 + ofs) = ABCtoVNI("Tæng PS")
    'xlSheet.Range("B" + CStr(sotkno + 7)).Font.Bold = True
    
    rx = CStr(sotkno + 7)
    cx = XLSCol(sotkco + 3 + ofs)
    
    For i = 1 To sotkno
        xlsheet.Range(XLSCol(sotkco + 3 + ofs) + CStr(i + 6)).Formula = "=SUM(" + XLSCol(3 + ofs) + CStr(i + 6) + ":" + XLSCol(sotkco + 2 + ofs) + CStr(i + 6) + ")"
    Next
    
    For i = 1 To sotkco
        sql = XLSCol(i + 2 + ofs)
        xlsheet.Range(sql + CStr(sotkno + 7)).Formula = "=SUM(" + sql + "7:" + sql + CStr(sotkno + 6) + ")"
    Next
    
    xlsheet.Range(cx + rx).Formula = "=SUM(" + cx + "7:" + cx + CStr(sotkno + 6) + ")"
    xlsheet.Range("B" + CStr(sotkno + 7), cx + rx).Font.Bold = True
    xlsheet.Range(xlsheet.Cells(5, sotkco + 3 + ofs), cx + rx).Font.Bold = True
    
'    xlSheet.Range("B6", cx + rx).Font.Name = "Arial"
    xlsheet.Range("A" + rx).Font.Name = ".VnArial"
    xlsheet.Range("A" + rx).WrapText = True
    xlsheet.Range(cx + "5").Font.Name = ".VnArial"
    xlsheet.Range(cx + "5").WrapText = True
    xlsheet.Range("A5", cx + "5").HorizontalAlignment = xlHAlignCenter
    xlsheet.Range("A5", "A" + rx).HorizontalAlignment = xlHAlignCenter
    
    xlsheet.Range("A3", cx + "3").MergeCells = True
    xlsheet.Range("A4", cx + "4").MergeCells = True
    
    xlsheet.Range(cx + "5", cx + "6").MergeCells = True
    xlsheet.Range(XLSCol(3 + ofs) + "5", XLSCol(sotkco + 2 + ofs) + "5").MergeCells = True
    xlsheet.Range(XLSCol(3 + ofs) + "5").HorizontalAlignment = 3
'    xlSheet.Range(rx + "5", rx + "6").MergeCells = True
    xlsheet.Range(IIf(ofs > 0, "C7", "B7"), IIf(ofs > 0, "C" + rx, "B" + rx)).NumberFormat = "dd/mm/yy" '.FormatConditions = "dd/mm/yy"
    With xlsheet.Range("A5", cx + rx)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    Set xlsheet = xlapp.Worksheets(2)
    BangKeTHPS xlsheet, TK, tdau, tcuoi, cap, 1
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "NK" + CStr(r) + ".XLS"
KetThuc:
    rs.Close
End Sub

Private Sub BK1(tdau As Integer, tcuoi As Integer, TK As String, nhom As Integer, cap As Integer, Optional id As Integer = 0)
Dim rs As Recordset, i As Integer, j As Integer, sql As String, d As Date, m As Long, mtk As Long, duno As Double, duco As Double, kieu As Integer
Dim sotkno As Integer, sotkco As Integer, rx As String, cx As String, lx As Integer, ofs As Integer, r As Integer

    ofs = IIf(id >= 2 Or (id = 0 And nhom = 0), 2, 0)
    
    r = IIf(id < 2, 1, 2)
    mtk = SelectSQL("SELECT MaSo AS F1, Kieu AS F2 FROM HethongTK WHERE SoHieu='" + TK + "'", kieu)
    lx = Len(TK)
    XDSoHieuCap cap
    ' Lay danh sach tai khoan tai chinh
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM ((ChungTu LEFT JOIN HethongTK ON ChungTu.MaTKCo = HethongTK.MaSo) INNER JOIN HethongTK AS TK ON ChungTu.MaTKNo=TK.MaSo) LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.RecordCount = 0 Then
        ErrMsg er_KoPS
        GoTo KetThuc
    End If
    
    Recycle pCurDir + "BK" + CStr(r) + ".XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\BK" + CStr(r) + ".XLS", pCurDir + "BK" + CStr(r) + ".XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "BK" + CStr(r) + ".XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = IIf(id > 0, "B¶ng kª sè " + CStr(id), "B¶ng kª ph¸t sinh nî tµi kho¶n " + TK)
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
'    xlSheet.Cells(5, 1) = "Tµi kho¶n ghi cã"
    If ofs > 0 Then
        xlsheet.Cells(7, 1) = "STT"
        xlsheet.Cells(7, ofs) = "Sè CT"
        xlsheet.Cells(7, 1 + ofs) = "Ngµy"
        xlsheet.Cells(7, 2 + ofs) = "DiÔn gi¶i"
        xlsheet.Range("B6", "D6").MergeCells = True
    End If
    xlsheet.Cells(6, 3 + ofs) = "Ghi nî tµi kho¶n " + TK + ", ghi cã c¸c tµi kho¶n"
    If nhom = 1 Then xlsheet.Cells(7, 1) = "STT"
' Liet ke theo hang ngang (phat sinh co)
    i = 2 + ofs
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco = i - 2 - ofs
    xlsheet.Cells(7, sotkco + 3 + ofs) = "Céng nî"
    
    sql = "SELECT First(MaCT) AS M,First(Chungtu.SoHieu) AS SH,NgayGS,BaoCaoCP.SoHieu AS ShCo,Sum(SoPS) AS PS,First(DienGiai" + IIf(nn > 0, "E", "") + ") AS DG FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON TK.SoHieu LIKE BaoCaoCP.SoHieu + '*'" _
        & " WHERE SoPS<>0 AND LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY NgayGS" + IIf(nhom = 0, ",MaCT", "") + ",BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    d = CVDate("01/01/1900")
    j = 7
' Liet ke theo hang doc (phat sinh no)
    rs.MoveLast
    GauGe.Max = rs.RecordCount
    rs.MoveFirst
    Do While Not rs.EOF
        If (rs!NgayGS <> d And nhom = 1) Or (rs!m <> m And nhom = 0) Then
            sotkno = sotkno + 1
            i = 3 + ofs
            j = j + 1
            d = rs!NgayGS
            m = rs!m
            HienThongBao Format(rs!NgayGS, Mask_D), 1
            If id < 2 And id > 0 Then
                xlsheet.Cells(j, 1) = IIf(nhom = 0, rs!sh, CStr(j - 7))
                xlsheet.Cells(j, 2) = rs!NgayGS
            Else
                xlsheet.Cells(j, 1) = CStr(j - 7)
                xlsheet.Cells(j, 2) = rs!sh
                xlsheet.Cells(j, 3) = rs!NgayGS
                xlsheet.Cells(j, 4) = rs!dg
            End If
            'xlSheet.Cells(j, 2) = Format(rs!ngaygs, Mask_D)
        End If
        'i = 3
        Do While rs!shco <> Left(xlsheet.Cells(7, i), Len(rs!shco))
            i = i + 1
        Loop
        xlsheet.Cells(j, i) = rs!ps
        GauGe.Value = GauGe.Value + 1
        rs.MoveNext
    Loop
    
    sotkno = j - 7

    xlsheet.Cells(sotkno + 8, 2 + ofs) = ABCtoVNI("Tæng PS")
    'xlSheet.Range("B" + CStr(sotkno + 8)).Font.Bold = True
    
    rx = CStr(sotkno + 8)
    cx = XLSCol(sotkco + 3 + ofs)
    
    For i = 1 To sotkno
        xlsheet.Range(XLSCol(sotkco + 3 + ofs) + CStr(i + 7)).Formula = "=SUM(" + XLSCol(3 + ofs) + CStr(i + 7) + ":" + XLSCol(sotkco + 2 + ofs) + CStr(i + 7) + ")"
    Next
    
    For i = 1 To sotkco
        sql = XLSCol(i + 2 + ofs)
        xlsheet.Range(sql + CStr(sotkno + 8)).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno + 7) + ")"
    Next
        
    xlsheet.Range(cx + rx).Formula = "=SUM(" + cx + "8:" + cx + CStr(sotkno + 7) + ")"
    
    xlsheet.Range("B" + rx, cx + rx).Font.Bold = True
    xlsheet.Range(xlsheet.Cells(5, sotkco + 3 + ofs), XLSCol(sotkco + 4 + ofs) + CStr(sotkno + 9)).Font.Bold = True
    
    xlsheet.Range("A" + rx).Font.Name = ".VnArial"
    xlsheet.Range("A" + rx).WrapText = True
    xlsheet.Range(cx + "6").Font.Name = ".VnArial"
    xlsheet.Range(cx + "6").WrapText = True
    xlsheet.Range("A6", cx + "6").HorizontalAlignment = xlHAlignCenter
    xlsheet.Range("A6", "A" + rx).HorizontalAlignment = xlHAlignCenter
    
    xlsheet.Range("A3", cx + "3").MergeCells = True
    xlsheet.Range("A4", cx + "4").MergeCells = True

    xlsheet.Range(cx + "6", cx + "7").MergeCells = True
    xlsheet.Range(XLSCol(sotkco + 4 + ofs) + "6", XLSCol(sotkco + 4 + ofs) + "7").MergeCells = True
    xlsheet.Range(XLSCol(3 + ofs) + "6", XLSCol(sotkco + 2 + ofs) + "6").MergeCells = True
    xlsheet.Range(XLSCol(3 + ofs) + "6").HorizontalAlignment = 3
    xlsheet.Range(IIf(ofs > 0, "C8", "B8"), IIf(ofs > 0, "C" + rx, "B" + rx)).NumberFormat = "dd/mm/yy" '.FormatConditions = "dd/mm/yy"
    With xlsheet.Range("A6", XLSCol(sotkco + 3 + ofs) + rx)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    Set xlsheet = xlapp.Worksheets(2)
    BangKeTHPS xlsheet, TK, tdau, tcuoi, cap, -1
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "BK" + CStr(r) + ".XLS"
KetThuc:
    rs.Close
End Sub

Private Sub GiaNhapKho(tdau As Integer, tcuoi As Integer)
    Dim duno As Double, duco As Double
    Dim rs As Recordset, i As Integer
    Dim ms(1 To 5) As Long
    
    ExecuteSQL5 "DELETE FROM BaoCaoCP"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,Ten,Cap,MaCha) VALUES (1,'DK','" + ABCtoVNI("I. Sè d­ ®Çu kú") + "',0,1)"
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo FROM HethongTK WHERE (SoHieu LIKE '152*') AND ((Cap=2) OR (Cap=1 AND TKCon=0)) ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        ms(i) = rs!MaSo
        SoDuTK ms(i), ThangTruoc(tdau), duno, duco
        ExecuteSQL5 "UPDATE BaoCaoCP SET Kq" + CStr(i) + "=" + DoiDau(duno)
        If i >= 5 Then Exit Do
        rs.MoveNext
    Loop
    rs.Close
    
    SoDuTK 33, ThangTruoc(tdau), duno, duco
    ExecuteSQL5 "UPDATE BaoCaoCP SET Kq6=" + DoiDau(duno)
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaTKNo,Sum(SoPS) AS PS,LEFT(HethongTK.SoHieu,6) AS SHCo,First(HethongTK.Ten) AS TTK FROM " + ChungTu2TKNC(2) _
        & "WHERE MaVattu>0 AND MaLoai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY LEFT(HethongTK.SoHieu,6),MaTKNo", dbOpenSnapshot)
    Do While Not rs.EOF
        Select Case rs!MaTkNo
            Case ms(1):      i = 1
            Case ms(2):      i = 2
            Case ms(3):      i = 3
            Case ms(4):      i = 4
            Case ms(5):      i = 5
            Case 34, 35, 36:          i = 6
            Case Else:
                i = 7
        End Select
        ExecuteSQL5 "UPDATE BaoCaoCP SET Kq" + CStr(i) + "=" + DoiDau(rs!ps) + " WHERE SoHieu='" + rs!shco + "' AND MaCha=2"
        If DBKetoan.RecordsAffected = 0 Then
            ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,Ten,Cap,Kq" + CStr(i) + ",MaCha) VALUES (" + rs!shco + ",'" + rs!shco + "','TK " + rs!shco + " - " + rs!TTK + "',1," + DoiDau(rs!ps) + ",2)"
        End If
        rs.MoveNext
    Loop
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaTKCo,Sum(SoPS) AS PS,LEFT(HethongTK.SoHieu,6) AS SHCo,First(HethongTK.Ten) AS TTK FROM " + ChungTu2TKNC(-1) _
        & "WHERE MaVattu>0 AND MaLoai=2 AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY LEFT(HethongTK.SoHieu,6),MaTKCo", dbOpenSnapshot)
    Do While Not rs.EOF
        Select Case rs!MaTkCo
            Case ms(1):      i = 1
            Case ms(2):      i = 2
            Case ms(3):      i = 3
            Case ms(4):     i = 4
            Case ms(5):      i = 5
            Case 34, 35, 36:          i = 6
            Case Else:
                i = 7
        End Select
        ExecuteSQL5 "UPDATE BaoCaoCP SET Kq" + CStr(i) + "=" + DoiDau(rs!ps) + " WHERE (SoHieu='" + rs!shco + "' OR SoHieu='" + rs!shco + "x') AND MaCha=3"
        If DBKetoan.RecordsAffected = 0 Then
            ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,Ten,Cap,Kq" + CStr(i) + ",MaCha) VALUES (" + rs!shco + "+1000,'" + rs!shco + "'+'x','TK " + rs!shco + " - " + rs!TTK + "',1," + DoiDau(rs!ps) + ",3)"
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
        
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,Ten,Cap,MaCha) VALUES (999999,'CK','" + ABCtoVNI("I. Sè d­ cuèi kú") + "',0,4)"
    
    For i = 1 To 5
        SoDuTK ms(i), tcuoi, duno, duco
        ExecuteSQL5 "UPDATE BaoCaoCP SET Kq" + CStr(i) + "=" + DoiDau(duno) + " WHERE MaSo=999999"
    Next
    
    SoDuTK 33, tcuoi, duno, duco
    ExecuteSQL5 "UPDATE BaoCaoCP SET Kq6=" + DoiDau(duno) + " WHERE MaSo=999999"
    
    frmMain.Rpt.ReportFileName = "GiaNhap.Rpt"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
End Sub

Private Sub NK4(tdau As Integer, tcuoi As Integer, TK As String, cap As Integer)
Dim rs As Recordset, i As Integer, j As Integer, sql As String, m As Long, mtk As Long, kieu As Integer
Dim sotkno1 As Integer, sotkco1 As Integer, sotkno2 As Integer, sotkco2 As Integer, rx As String, cx As String, lx As Integer
Dim duno As Double, duco As Double

    lx = Len(TK)
    mtk = SelectSQL("SELECT MaSo AS F1, Kieu AS F2 FROM HethongTK WHERE SoHieu='" + TK + "'", kieu)
    XDSoHieuCap cap
    ' Lay danh sach tai khoan tai chinh
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    
    Recycle pCurDir + "NK4.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\NK4.XLS", pCurDir + "NK4.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "NK4.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = "NhËt ký chøng tõ sè 4"
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    xlsheet.Cells(6, 5) = "Ghi cã tµi kho¶n " + TK + ", ghi nî c¸c tµi kho¶n"
    
    ' Liet ke theo hang ngang (phat sinh co)
    i = 4
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco1 = i - 4

    ' xlSheet.Cells(7, sotkco + 3) = IIf(otk.kieu < 0, dn - dc, dc - dn)
    ' xlsheet.Range("C5","").
    xlsheet.Cells(7, sotkco1 + 5) = "Céng cã"
    
    sql = "SELECT First(Chungtu.MaCT) AS M,First(Chungtu.SoHieu) AS SH,NgayGS,BaoCaoCP.SoHieu AS ShCo,Sum(SoPS) AS PS,First(DienGiai" + IIf(nn > 0, "E", "") + ") AS DG FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY NgayGS,MaCT,BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 7

    ' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            If rs!m <> m Then
                sotkno1 = sotkno1 + 1
                i = 5
                j = j + 1
                m = rs!m
                HienThongBao Format(rs!NgayGS, Mask_D), 1
                xlsheet.Cells(j, 1) = CStr(j - 7)
                xlsheet.Cells(j, 2) = rs!sh
                xlsheet.Cells(j, 3) = rs!NgayGS
                xlsheet.Cells(j, 4) = rs!dg
            End If
            Do While rs!shco <> Left(xlsheet.Cells(7, i), Len(rs!shco))
                i = i + 1
            Loop
            xlsheet.Cells(j, i) = rs!ps
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        
        sotkno1 = j - 7
    Else
        sotkno1 = 0
    End If
    
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON TK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    xlsheet.Cells(6, sotkco1 + 6) = "STT"
    xlsheet.Cells(6, sotkco1 + 7) = "Chøng tõ"
    xlsheet.Cells(7, sotkco1 + 7) = "SCT"
    xlsheet.Cells(7, sotkco1 + 8) = "Ngµy"
    
    If sotkco1 > 0 Then xlsheet.Range(XLSCol(sotkco1 + 5) + "6", XLSCol(sotkco1 + 5) + "7").MergeCells = True
    xlsheet.Range(XLSCol(sotkco1 + 6) + "6", XLSCol(sotkco1 + 6) + "7").MergeCells = True
    xlsheet.Range(XLSCol(sotkco1 + 7) + "6", XLSCol(sotkco1 + 8) + "6").MergeCells = True
    xlsheet.Cells(6, sotkco1 + 9) = "Ghi nî tµi kho¶n " + TK + ", ghi cã c¸c tµi kho¶n"
    
' Liet ke theo hang ngang (phat sinh co)
    i = sotkco1 + 8
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco2 = i - (sotkco1 + 8)
    
    xlsheet.Cells(6, sotkco1 + 5 + sotkco2 + 4) = "Céng nî"
    
    sql = "SELECT First(MaCT) AS M,First(Chungtu.SoHieu) AS SH,NgayGS,BaoCaoCP.SoHieu AS ShCo,Sum(SoPS) AS PS,First(DienGiai" + IIf(nn > 0, "E", "") + ") AS DG FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON TK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY NgayGS,MaCT,BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 7
' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        GauGe.Value = 0
        rs.MoveFirst
        Do While Not rs.EOF
            If rs!m <> m Then
                sotkno2 = sotkno2 + 1
                i = sotkco1 + 9
                j = j + 1
                m = rs!m
                HienThongBao Format(rs!NgayGS, Mask_D), 1
                xlsheet.Cells(j, sotkco1 + 6) = CStr(j - 7)
                xlsheet.Cells(j, sotkco1 + 7) = rs!sh
                xlsheet.Cells(j, sotkco1 + 8) = rs!NgayGS
            End If
            'i = 3
            Do While rs!shco <> Left(xlsheet.Cells(7, i), Len(rs!shco))
                i = i + 1
            Loop
            xlsheet.Cells(j, i) = rs!ps
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        
        sotkno2 = j - 7
    Else
        sotkno2 = 0
    End If
    
    xlsheet.Cells(IIf(sotkno1 > sotkno2, sotkno1, sotkno2) + 8, 4) = ABCtoVNI("Tæng PS")
    rx = CStr(IIf(sotkno1 > sotkno2, sotkno1, sotkno2) + 8)
    
    cx = XLSCol(sotkco1 + 5)
    
    For i = 1 To sotkno1
        xlsheet.Range(XLSCol(sotkco1 + 5) + CStr(i + 7)).Formula = "=SUM(E" + CStr(i + 7) + ":" + XLSCol(sotkco1 + 4) + CStr(i + 7) + ")"
    Next
    
    If sotkno1 > 0 Then xlsheet.Range(cx + rx).Formula = "=SUM(" + cx + "8:" + cx + CStr(sotkno1 + 7) + ")"
    
    
    cx = XLSCol(sotkco1 + 5 + sotkco2 + 4)
    
    For i = 1 To sotkno2
        xlsheet.Range(cx + CStr(i + 7)).Formula = "=SUM(" + XLSCol(sotkco1 + 9) + CStr(i + 7) + ":" + XLSCol(sotkco1 + 5 + sotkco2 + 3) + CStr(i + 7) + ")"
    Next
    
    For i = 1 To sotkco1
        sql = XLSCol(i + 4)
'        xlSheet.Range(sql + CStr(sotkno1 + 8)).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno1 + 7) + ")"
        xlsheet.Range(sql + rx).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno1 + 7) + ")"
    Next
    
    For i = 1 To sotkco2
        sql = XLSCol(i + sotkco1 + 8)
        'xlSheet.Range(sql + CStr(sotkno1 + 8)).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno1 + 7) + ")"
        xlsheet.Range(sql + rx).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno2 + 7) + ")"
    Next
    
    If sotkno1 > 0 Or sotkno2 > 0 Then xlsheet.Range(cx + rx).Formula = "=SUM(" + cx + "8:" + cx + CStr(IIf(sotkno1 > sotkno2, sotkno1, sotkno2) + 7) + ")"
    
    xlsheet.Cells(5, sotkco1 + 5 + sotkco2 + 3) = "Sè ®Çu kú"
    SoDuTK mtk, ThangTruoc(tdau), duno, duco
    xlsheet.Cells(5, sotkco1 + 6 + sotkco2 + 3) = IIf(kieu < 0, duno - duco, duco - duno)
    
    xlsheet.Cells(rx + 1, sotkco1 + 5 + sotkco2 + 3) = ABCtoVNI("Sè cuèi kú")
    SoDuTK mtk, tcuoi, duno, duco
    xlsheet.Cells(rx + 1, sotkco1 + 6 + sotkco2 + 3) = IIf(kieu < 0, duno - duco, duco - duno)
            
    xlsheet.Range("A6", cx + "6").HorizontalAlignment = xlHAlignCenter
    xlsheet.Range("A7", "A" + rx).HorizontalAlignment = xlHAlignCenter
    xlsheet.Range(XLSCol(sotkco1 + 6) + "7", XLSCol(sotkco1 + 6) + rx).HorizontalAlignment = xlHAlignCenter
    
    xlsheet.Range("A3", cx + "3").MergeCells = True
    xlsheet.Range("A4", cx + "4").MergeCells = True
    
    xlsheet.Range(cx + "6", cx + "7").MergeCells = True
    xlsheet.Range("A5", cx + "6").Font.Bold = True
    xlsheet.Range("A" + rx, cx + CStr(IIf(sotkno1 > sotkno2, sotkno1, sotkno2) + 9)).Font.Bold = True
    
    
    xlsheet.Range("E6").HorizontalAlignment = 3
'    xlSheet.Range(rx + "5", rx + "6").MergeCells = True
    
    With xlsheet.Range("A6", cx + rx)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "NK4.XLS"
KetThuc:
    rs.Close
End Sub

Private Sub NK5(tdau As Integer, tcuoi As Integer, TK As String, cap As Integer)
Dim rs As Recordset, i As Integer, j As Integer, sql As String, kq As Double
Dim sotkco1 As Integer, sotkno2 As Integer, sotkco2 As Integer, rx As String, cx As String, lx As Integer
Dim sqln As String, sqlc As String, st As String

    lx = Len(TK)
    XDSoHieuCap cap
    ' Lay danh sach tai khoan tai chinh
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    
    Recycle pCurDir + "NK5.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\NK5.XLS", pCurDir + "NK5.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "NK5.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = "NhËt ký chøng tõ sè 5"
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    xlsheet.Cells(6, 5) = "Ghi cã tµi kho¶n " + TK + ", ghi nî c¸c tµi kho¶n"
    
    ' Liet ke theo hang ngang (phat sinh co)
    i = 4
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco1 = i - 4

    xlsheet.Cells(6, sotkco1 + 5) = "Céng cã"
    
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON TK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    
    xlsheet.Cells(6, sotkco1 + 6) = "Ghi nî tµi kho¶n " + TK + ", ghi cã c¸c tµi kho¶n"
    
' Liet ke theo hang ngang (phat sinh co)
    i = sotkco1 + 5
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco2 = i - (sotkco1 + 5)
    
    xlsheet.Cells(6, sotkco1 + 5 + sotkco2 + 1) = "Céng nî"
    xlsheet.Cells(6, sotkco1 + 5 + sotkco2 + 2) = "Sè d­ cuèi th¸ng"
    xlsheet.Cells(7, sotkco1 + 5 + sotkco2 + 2) = "Nî"
    xlsheet.Cells(7, sotkco1 + 5 + sotkco2 + 3) = "Cã"
    cx = XLSCol(sotkco1 + 5 + sotkco2 + 3)
    
    xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 2) + "6", cx + "6").MergeCells = True
    xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 1) + "6", XLSCol(sotkco1 + 5 + sotkco2 + 1) + "7").MergeCells = True
    If sotkco2 > 1 Then xlsheet.Range(XLSCol(sotkco1 + 6) + "6", XLSCol(sotkco1 + 5 + sotkco2) + "6").MergeCells = True
    xlsheet.Range(XLSCol(sotkco1 + 5) + "6", XLSCol(sotkco1 + 5) + "7").MergeCells = True
    If sotkco1 > 1 Then xlsheet.Range("E6", XLSCol(sotkco1 + 4) + "6").MergeCells = True
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqln = sqln + " + SoDuKhachHang.No_" + CStr(i)
        sqlc = sqlc + " + SoDuKhachHang.Co_" + CStr(i)
    Next
    st = CStr(CThangDB(ThangTruoc(tdau)))
    sql = "SELECT KhachHang.MaSo,KhachHang.SoHieu,KhachHang.Ten,Sum(SoDuKhachHang.DuNo_" + st + ") AS ndk,Sum(SoDuKhachHang.DuCo_" + st + ") AS cdk,Sum(SoDuKhachHang.DuNo_" + CStr(CThangDB(tcuoi)) + ") AS nck,Sum(SoDuKhachHang.DuCo_" + CStr(CThangDB(tcuoi)) + ") AS ckc,Sum(" + sqln + ") AS N,Sum(" + sqlc + ") AS C FROM (SoDuKhachHang INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo) INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo " _
        & " WHERE LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "' AND (SoDuKhachHang.DuNo_" + st + "<>0 OR SoDuKhachHang.DuCo_" + st + "<>0 OR " + sqln + "<>0 OR " + sqlc + "<>0) GROUP BY KhachHang.MaSo,KhachHang.SoHieu,KhachHang.Ten"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 7
' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        GauGe.Value = 0
        rs.MoveFirst
        Do While Not rs.EOF
            sotkno2 = sotkno2 + 1
            j = j + 1
            xlsheet.Cells(j, 1) = CStr(j - 7)
            xlsheet.Cells(j, 2) = rs!sohieu + " - " + rs!Ten
            xlsheet.Cells(j, 3) = rs!ndk
            xlsheet.Cells(j, 4) = rs!cdk
            xlsheet.Cells(j, sotkco1 + 5 + sotkco2 + 2) = rs!nck
            xlsheet.Cells(j, sotkco1 + 5 + sotkco2 + 3) = rs!ckc
            If rs!c <> 0 Then
                For i = 1 To sotkco1
                    kq = PSDuKH(rs!MaSo, TK, xlsheet.Cells(7, i + 4), tdau, tcuoi, 1)
                    If kq > 0 Then xlsheet.Cells(j, i + 4) = kq
                Next
            End If
            If rs!n <> 0 Then
                For i = 1 To sotkco2
                    kq = PSDuKH(rs!MaSo, TK, xlsheet.Cells(7, sotkco1 + 5 + i), tdau, tcuoi, -1)
                    If kq > 0 Then xlsheet.Cells(j, sotkco1 + 5 + i) = kq
                Next
            End If
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        
        sotkno2 = j - 7
    Else
        sotkno2 = 0
    End If
    xlsheet.Cells(sotkno2 + 8, 2) = ABCtoVNI("Tæng PS")
    rx = CStr(sotkno2 + 8)
    sql = XLSCol(sotkco1 + 5)
    If sotkco1 > 0 Then
        For i = 1 To sotkno2
            xlsheet.Range(sql + CStr(i + 7)).Formula = "=SUM(E" + CStr(i + 7) + ":" + XLSCol(sotkco1 + 4) + CStr(i + 7) + ")"
        Next
    End If
    If sotkco2 > 0 Then
        For i = 1 To sotkno2
            xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 1) + CStr(i + 7)).Formula = "=SUM(" + XLSCol(sotkco1 + 6) + CStr(i + 7) + ":" + XLSCol(sotkco1 + 5 + sotkco2) + CStr(i + 7) + ")"
        Next
    End If
    For i = 1 To sotkco1
        sql = XLSCol(i + 4)
        xlsheet.Range(sql + CStr(sotkno2 + 8)).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno2 + 7) + ")"
    Next
    
    For i = 1 To sotkco2
        sql = XLSCol(i + sotkco1 + 5)
        xlsheet.Range(sql + CStr(sotkno2 + 8)).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno2 + 7) + ")"
    Next
        
    If sotkno2 > 0 Then
        For i = 1 To sotkco1 + sotkco2 + 6
            sql = XLSCol(i + 2)
            xlsheet.Range(sql + rx).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno2 + 7) + ")"
        Next
    End If
        
    xlsheet.Range("A6", cx + "6").HorizontalAlignment = xlHAlignCenter
    xlsheet.Range("A7", "A" + rx).HorizontalAlignment = xlHAlignCenter
    xlsheet.Range(XLSCol(sotkco1 + 6) + "7", XLSCol(sotkco1 + 6) + rx).HorizontalAlignment = xlHAlignCenter
    
    xlsheet.Range("A3", cx + "3").MergeCells = True
    xlsheet.Range("A4", cx + "4").MergeCells = True
        
    xlsheet.Range("A5", cx + "6").Font.Bold = True
    xlsheet.Range("A" + rx, cx + rx).Font.Bold = True
        
    With xlsheet.Range("A6", cx + rx)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "NK5.XLS"
KetThuc:
    rs.Close
End Sub

Private Function PSDuKH(mkh As Long, sh1 As String, shtk As String, tdau As Integer, tcuoi As Integer, loai As Integer) As Double
    If loai < 0 Then
        PSDuKH = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
            & " WHERE ChungTu.MaKH=" + CStr(mkh) + " AND HethongTK.SoHieu LIKE '" + sh1 + "*' AND TK.SoHieu LIKE '" + shtk + "*' AND " + WThang("ThangCT", tdau, tcuoi))
    Else
        PSDuKH = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
            & " WHERE HethongTK.SoHieu LIKE '" + shtk + "*' AND TK.SoHieu LIKE '" + sh1 + "*' AND ChungTu.MaKHC=" + CStr(mkh) + " AND " + WThang("ThangCT", tdau, tcuoi))
    End If
End Function

Private Function PSDuVT(shno As String, shco As String, tdau As Integer, tcuoi As Integer, mvt As Long) As Double
    PSDuVT = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(0) + " WHERE MaVattu=" + CStr(mvt) + " AND HethongTK.SoHieu LIKE '" + shno + "*' AND TK.SoHieu LIKE '" + shco + "*' AND " + WThang("ThangCT", tdau, tcuoi))
End Function

Private Sub NK6(tdau As Integer, tcuoi As Integer, TK As String, cap As Integer)
Dim rs As Recordset, i As Integer, j As Integer, sql As String, sqln As String, sqlc As String, st As String
Dim sotkno As Integer, sotkco As Integer, rx As String, cx As String, lx As Integer, kq As Double

    lx = Len(TK)
    XDSoHieuCap cap
    ' Lay danh sach tai khoan tai chinh
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    
    Recycle pCurDir + "NK6.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\NK6.XLS", pCurDir + "NK6.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "NK6.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = "NhËt ký chøng tõ sè 6"
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    xlsheet.Cells(6, 8) = "Ghi cã tµi kho¶n " + TK + ", ghi nî c¸c tµi kho¶n"

    ' Liet ke theo hang ngang (phat sinh co)
    i = 7
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco = i - 7

    xlsheet.Cells(6, sotkco + 8) = "Céng cã"
    xlsheet.Cells(6, sotkco + 9) = "Sè d­ cuèi th¸ng"
    cx = XLSCol(sotkco + 9)
    
    xlsheet.Range(XLSCol(sotkco + 8) + "6", XLSCol(sotkco + 8) + "7").MergeCells = True
    xlsheet.Range(cx + "6", cx + "7").MergeCells = True
    If sotkco > 1 Then xlsheet.Range("H6", XLSCol(sotkco + 8) + "6").MergeCells = True
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqln = sqln + " + Tien_Nhap_" + CStr(i)
        sqlc = sqlc + " + Tien_Xuat_" + CStr(i)
    Next
    st = CStr(CThangDB(ThangTruoc(tdau)))
    sql = "SELECT TonKho.MaVattu,Vattu.SoHieu,TenVattu,Tien_" + st + " AS ndk,Tien_" + CStr(CThangDB(tcuoi)) + " AS nck,(" + sqln + ") AS N,(" + sqlc + ") AS C FROM (Vattu INNER JOIN TonKho ON Vattu.MaSo=TonKho.MaVattu) INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo " _
        & " WHERE (LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "') AND (Tien_" + st + "<>0 OR " + sqln + "<>0 OR " + sqlc + "<>0) ORDER BY Vattu.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 7
' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        GauGe.Value = 0
        rs.MoveFirst
        Do While Not rs.EOF
            sotkno = sotkno + 1
            j = j + 1
            xlsheet.Cells(j, 1) = CStr(j - 7)
            xlsheet.Cells(j, 2) = rs!sohieu + " - " + rs!TenVattu
            xlsheet.Cells(j, 3) = rs!ndk
            xlsheet.Cells(j, sotkco + 9) = rs!nck
            
            If rs!c > 0 Then
                For i = 1 To sotkco
                    kq = PSDuVT(xlsheet.Cells(7, i + 7), TK, tdau, tcuoi, rs!MaVattu)
                    If kq > 0 Then xlsheet.Cells(j, i + 7) = kq
                Next
            End If
            
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        
        sotkno = j - 7
    Else
        sotkno = 0
    End If
    
    xlsheet.Cells(sotkno + 8, 2) = "Toång PS"
    rx = CStr(sotkno + 8)
    sql = XLSCol(sotkco + 8)
    If sotkco > 0 Then
        For i = 1 To sotkno
            xlsheet.Range(sql + CStr(i + 7)).Formula = "=SUM(H" + CStr(i + 7) + ":" + XLSCol(sotkco + 7) + CStr(i + 7) + ")"
        Next
    End If
        
    If sotkno > 0 Then
        xlsheet.Range("C" + rx).Formula = "=SUM(C8:" + sql + CStr(sotkno + 7) + ")"
        For i = 1 To sotkco + 2
            sql = XLSCol(i + 7)
            xlsheet.Range(sql + rx).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno + 7) + ")"
        Next
    End If
        
    xlsheet.Range("A6", cx + "6").HorizontalAlignment = xlHAlignCenter
    xlsheet.Range("A7", "A" + rx).HorizontalAlignment = xlHAlignCenter
    
    xlsheet.Range("A3", cx + "3").MergeCells = True
    xlsheet.Range("A4", cx + "4").MergeCells = True
        
    xlsheet.Range("A5", cx + "6").Font.Bold = True
    xlsheet.Range("A" + rx, cx + rx).Font.Bold = True
        
    With xlsheet.Range("A6", cx + rx)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "NK6.XLS"
KetThuc:
    rs.Close
End Sub

Private Sub NK7(tdau As Integer, tcuoi As Integer)
Dim i As Integer, j As Integer, kq As Double, tps As Double, sqln As String
    
    Recycle pCurDir + "NK7.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\NK7.XLS", pCurDir + "NK7.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "NK7.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqln = sqln + " + No_" + CStr(i)
    Next
    
    For i = 1 To 11
        j = 1
        Do While InStr(xlsheet.Cells(7, j + 2), "#") = 0 And j < 30
            kq = PSDu(xlsheet.Cells(i + 8, 2), xlsheet.Cells(7, j + 2), tdau, tcuoi)
            tps = tps + kq
            If kq <> 0 Then xlsheet.Cells(i + 8, j + 2) = kq
            kq = SelectSQL("SELECT (" + sqln + ") AS F1 FROM HethongTK WHERE SoHieu='" + CStr(xlsheet.Cells(i + 8, 2)) + "'") - tps
            If kq <> 0 Then xlsheet.Cells(i + 8, 18) = kq
            j = j + 1
        Loop
    Next
    
    For i = 1 To 9
        j = 1
        Do While InStr(xlsheet.Cells(7, j + 2), "#") = 0 And j < 30
            kq = PSDu(xlsheet.Cells(i + 25, 2), xlsheet.Cells(7, j + 2), tdau, tcuoi)
            If kq > 0 Then xlsheet.Cells(i + 25, j + 2) = kq
            j = j + 1
        Loop
    Next
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "NK7.XLS"
KetThuc:
End Sub

Private Sub NK8(tdau As Integer, tcuoi As Integer)
Dim i As Integer, j As Integer, kq As Double
    
    Recycle pCurDir + "NK8.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\NK8.XLS", pCurDir + "NK8.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "NK8.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    
    For i = 1 To 12
        For j = 1 To 18
            kq = PSDu(xlsheet.Cells(i + 7, 2), xlsheet.Cells(6, j + 2), tdau, tcuoi)
            If kq <> 0 Then xlsheet.Cells(i + 7, j + 2) = kq
        Next
    Next
        
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "NK8.XLS"
KetThuc:
End Sub

Private Sub NK9(tdau As Integer, tcuoi As Integer, cap As Integer)
Dim rs As Recordset, i As Integer, j As Integer, sql As String, m As Long
Dim sotkno1 As Integer, sotkco1 As Integer, sotkno2 As Integer, sotkco2 As Integer, sotkno3 As Integer, sotkco3 As Integer, rx As String, cx As String, sotkno As Integer
    
    Recycle pCurDir + "NK9.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\NK4.XLS", pCurDir + "NK9.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "NK9.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = "NhËt ký chøng tõ sè 9"
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    xlsheet.Cells(6, 5) = "Ghi cã tµi kho¶n 211, ghi nî c¸c tµi kho¶n"
    
    ' Liet ke theo hang ngang (phat sinh co) 211
    XDSoHieuCap cap
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu,3)='211' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    i = 4
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco1 = i - 4
    
    xlsheet.Cells(6, sotkco1 + 5) = "Céng cã 211"
    xlsheet.Range(XLSCol(sotkco1 + 5) + "6", XLSCol(sotkco1 + 5) + "7").MergeCells = True
    If sotkco1 > 1 Then xlsheet.Range("E6", XLSCol(sotkco1 + 4) + "6").MergeCells = True
    
    sql = "SELECT First(Chungtu.MaCT) AS M,First(Chungtu.SoHieu) AS SH,NgayGS,BaoCaoCP.SoHieu AS ShCo,Sum(SoPS) AS PS,First(DienGiai" + IIf(nn > 0, "E", "") + ") AS DG FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu,3)='211' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY NgayGS,MaCT,BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 7

    ' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            If rs!m <> m Then
                sotkno1 = sotkno1 + 1
                i = 5
                j = j + 1
                m = rs!m
                HienThongBao Format(rs!NgayGS, Mask_D), 1
                xlsheet.Cells(j, 1) = CStr(j - 7)
                xlsheet.Cells(j, 2) = rs!sh
                xlsheet.Cells(j, 3) = rs!NgayGS
                xlsheet.Cells(j, 4) = rs!dg
            End If
            Do While rs!shco <> Left(xlsheet.Cells(7, i), Len(rs!shco))
                i = i + 1
            Loop
            xlsheet.Cells(j, i) = rs!ps
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        sotkno1 = j - 7
    Else
        sotkno1 = 0
    End If
    
    For i = 1 To sotkno1
        xlsheet.Range(XLSCol(sotkco1 + 5) + CStr(i + 7)).Formula = "=SUM(E" + CStr(i + 7) + ":" + XLSCol(sotkco1 + 4) + CStr(i + 7) + ")"
    Next
    
    xlsheet.Cells(6, sotkco1 + 6) = "Ghi cã tµi kho¶n 212, ghi nî c¸c tµi kho¶n"
    ' Liet ke theo hang ngang (phat sinh co) 212
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu,3)='212' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    i = sotkco1 + 5
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco2 = i - (sotkco1 + 5)
    
    xlsheet.Cells(6, sotkco1 + 5 + sotkco2 + 1) = "Céng cã 212"
    xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 1) + "6", XLSCol(sotkco1 + 5 + sotkco2 + 1) + "7").MergeCells = True
    If sotkco2 > 1 Then xlsheet.Range(XLSCol(sotkco1 + 6) + "6", XLSCol(sotkco1 + 5 + sotkco2) + "6").MergeCells = True
    m = 0
    sql = "SELECT First(Chungtu.MaCT) AS M,First(Chungtu.SoHieu) AS SH,NgayGS,BaoCaoCP.SoHieu AS ShCo,Sum(SoPS) AS PS,First(DienGiai" + IIf(nn > 0, "E", "") + ") AS DG FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu,3)='212' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY NgayGS,MaCT,BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 7
    ' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            If rs!m <> m Then
                sotkno2 = sotkno2 + 1
                i = sotkco1 + 6
                j = j + 1
                m = rs!m
                HienThongBao Format(rs!NgayGS, Mask_D), 1
            End If
            Do While rs!shco <> Left(xlsheet.Cells(7, i), Len(rs!shco))
                i = i + 1
            Loop
            xlsheet.Cells(j, i) = rs!ps
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        sotkno2 = j - 7
    Else
        sotkno2 = 0
    End If
    
    For i = 1 To sotkno2
        xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 1) + CStr(i + 7)).Formula = "=SUM(" + XLSCol(sotkco1 + 6) + CStr(i + 7) + ":" + XLSCol(sotkco1 + 5 + sotkco2) + CStr(i + 7) + ")"
    Next
    
    xlsheet.Cells(6, sotkco1 + 5 + sotkco2 + 2) = "Ghi cã tµi kho¶n 213, ghi nî c¸c tµi kho¶n"
    ' Liet ke theo hang ngang (phat sinh co) 213
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu,3)='213' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    i = sotkco1 + 5 + sotkco2 + 1
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco3 = i - (sotkco1 + 5 + sotkco2 + 1)
    
    xlsheet.Cells(6, sotkco1 + 5 + sotkco2 + 1 + sotkco3 + 1) = "Céng cã 213"
    xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 1 + sotkco3 + 1) + "6", XLSCol(sotkco1 + 5 + sotkco2 + 1 + sotkco3 + 1) + "7").MergeCells = True
    If sotkco3 > 1 Then xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 2) + "6", XLSCol(sotkco1 + 5 + sotkco2 + 1 + sotkco3) + "6").MergeCells = True
    m = 0
    sql = "SELECT First(Chungtu.MaCT) AS M,First(Chungtu.SoHieu) AS SH,NgayGS,BaoCaoCP.SoHieu AS ShCo,Sum(SoPS) AS PS,First(DienGiai" + IIf(nn > 0, "E", "") + ") AS DG FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu,3)='213' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY NgayGS,MaCT,BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 7
    ' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            If rs!m <> m Then
                sotkno3 = sotkno3 + 1
                i = sotkco1 + 5 + sotkco2 + 2
                j = j + 1
                m = rs!m
                HienThongBao Format(rs!NgayGS, Mask_D), 1
            End If
            Do While rs!shco <> Left(xlsheet.Cells(7, i), Len(rs!shco))
                i = i + 1
            Loop
            xlsheet.Cells(j, i) = rs!ps
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        sotkno3 = j - 7
    Else
        sotkno3 = 0
    End If
    
    For i = 1 To sotkno3
        xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 1 + sotkco3 + 1) + CStr(i + 7)).Formula = "=SUM(" + XLSCol(sotkco1 + 5 + sotkco2 + 2) + CStr(i + 7) + ":" + XLSCol(sotkco1 + 5 + sotkco2 + 1 + sotkco3) + CStr(i + 7) + ")"
    Next
        
    sotkno = IIf(sotkno1 > sotkno2, IIf(sotkno1 > sotkno3, sotkno1, sotkno3), IIf(sotkno2 > sotkno3, sotkno2, sotkno3))
    
    xlsheet.Cells(sotkno + 8, 4) = "Toång PS"
    
    rx = CStr(sotkno + 8)
    cx = XLSCol(sotkco1 + 5 + sotkco2 + 1 + sotkco3 + 1)
                    
    If sotkno1 > 0 Or sotkno2 > 0 Or sotkno3 > 0 Then
        For i = 5 To sotkco1 + 5 + sotkco2 + 1 + sotkco3 + 1
            sql = XLSCol(i)
            xlsheet.Range(sql + rx).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno + 7) + ")"
        Next
    End If
    
'    xlSheet.Range("A6", cx + "6").HorizontalAlignment = xlHAlignCenter
'    xlSheet.Range("A7", "A" + rx).HorizontalAlignment = xlHAlignCenter
    
    xlsheet.Range("A3", cx + "3").MergeCells = True
    xlsheet.Range("A4", cx + "4").MergeCells = True
    
    xlsheet.Range("A6", cx + "7").Font.Bold = True
    xlsheet.Range("A" + rx, cx + rx).Font.Bold = True
    
    xlsheet.Range("E6").HorizontalAlignment = 3
'    xlSheet.Range(rx + "5", rx + "6").MergeCells = True
    
    With xlsheet.Range("A6", cx + rx)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "NK9.XLS"
KetThuc:
    rs.Close
End Sub

Private Sub NK10(tdau As Integer, tcuoi As Integer, TK As String, loai As Integer, cap As Integer)
Dim rs As Recordset, i As Integer, j As Integer, sql As String, kq As Double, tkid As Long
Dim sotkco1 As Integer, sotkno2 As Integer, sotkco2 As Integer, rx As String, cx As String, lx As Integer
Dim sqln As String, sqlc As String, st As String
    
    lx = Len(TK)
    Recycle pCurDir + "NK10.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\NK5.XLS", pCurDir + "NK10.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "NK10.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = IIf(loai = 0, "NhËt ký chøng tõ sè 10", "B¶ng kª sè 6")
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    xlsheet.Cells(6, 5) = "Ghi nî tµi kho¶n " + TK + ", ghi cã c¸c tµi kho¶n"
    
    XDSoHieuCap cap
    ' Liet ke theo hang ngang (phat sinh co)
    ' Lay danh sach tai khoan tai chinh
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON TK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    i = 4
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco1 = i - 4

    xlsheet.Cells(6, sotkco1 + 5) = "Céng nî"
    
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON TK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    
    xlsheet.Cells(6, sotkco1 + 6) = "Ghi cã tµi kho¶n " + TK + ", ghi nî c¸c tµi kho¶n"
    
' Liet ke theo hang ngang (phat sinh co)
    i = sotkco1 + 5
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco2 = i - (sotkco1 + 5)
    
    xlsheet.Cells(6, sotkco1 + 5 + sotkco2 + 1) = "Céng cã"
    xlsheet.Cells(6, sotkco1 + 5 + sotkco2 + 2) = "Sè d­ cuèi th¸ng"
    xlsheet.Cells(7, sotkco1 + 5 + sotkco2 + 2) = "Nî"
    xlsheet.Cells(7, sotkco1 + 5 + sotkco2 + 3) = "Cã"
    cx = XLSCol(sotkco1 + 5 + sotkco2 + 3)
    
    xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 2) + "6", cx + "6").MergeCells = True
    xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 1) + "6", XLSCol(sotkco1 + 5 + sotkco2 + 1) + "7").MergeCells = True
    If sotkco2 > 1 Then xlsheet.Range(XLSCol(sotkco1 + 6) + "6", XLSCol(sotkco1 + 5 + sotkco2) + "6").MergeCells = True
    xlsheet.Range(XLSCol(sotkco1 + 5) + "6", XLSCol(sotkco1 + 5) + "7").MergeCells = True
    If sotkco1 > 1 Then xlsheet.Range("E6", XLSCol(sotkco1 + 4) + "6").MergeCells = True
    
    tkid = GetTK_ID(TK, 0)
    If tkid = TKCNKH_ID Or tkid = TKCNPT_ID Then
        For i = CThangDB(tdau) To CThangDB(tcuoi)
            sqln = sqln + " + SoDuKhachHang.No_" + CStr(i)
            sqlc = sqlc + " + SoDuKhachHang.Co_" + CStr(i)
        Next
        st = CStr(CThangDB(ThangTruoc(tdau)))
        sql = "SELECT MaKhachHang, KhachHang.SoHieu,KhachHang.Ten AS TenTK, HethongTK.SoHieu AS SHTK,Sum(SoDuKhachHang.DuNo_" + st + ") AS ndk,Sum(SoDuKhachHang.DuCo_" + st + ") AS cdk,Sum(SoDuKhachHang.DuNo_" + CStr(CThangDB(tcuoi)) + ") AS nck,Sum(SoDuKhachHang.DuCo_" + CStr(CThangDB(tcuoi)) + ") AS ckc,Sum(" + sqln + ") AS N,Sum(" + sqlc + ") AS C " _
            & " FROM (SoDuKhachHang INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo) INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo " _
            & " WHERE LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "' AND (SoDuKhachHang.DuNo_" + st + "<>0 OR SoDuKhachHang.DuCo_" + st + "<>0 OR " + sqln + "<>0 OR " + sqlc + "<>0) GROUP BY HethongTK.SoHieu, KhachHang.SoHieu, KhachHang.Ten, MaKhachHang"
    Else
        For i = CThangDB(tdau) To CThangDB(tcuoi)
            sqln = sqln + " + No_" + CStr(i)
            sqlc = sqlc + " + Co_" + CStr(i)
        Next
        st = CStr(CThangDB(ThangTruoc(tdau)))
        sql = "SELECT SoHieu,Ten" + IIf(nn > 0, "E", "") + " AS TenTK,Sum(DuNo_" + st + ") AS ndk,Sum(DuCo_" + st + ") AS cdk,Sum(DuNo_" + CStr(CThangDB(tcuoi)) + ") AS nck,Sum(DuCo_" + CStr(CThangDB(tcuoi)) + ") AS ckc,Sum(" + sqln + ") AS N,Sum(" + sqlc + ") AS C FROM HethongTK " _
            & " WHERE TKCon=0 AND LEFT(SoHieu," + CStr(lx) + ")='" + TK + "' AND (DuNo_" + st + "<>0 OR DuCo_" + st + "<>0 OR " + sqln + "<>0 OR " + sqlc + "<>0) GROUP BY SoHieu,Ten"
    End If
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 7
' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        GauGe.Value = 0
        rs.MoveFirst
        Do While Not rs.EOF
            sotkno2 = sotkno2 + 1
            j = j + 1
            xlsheet.Cells(j, 1) = CStr(j - 7)
            xlsheet.Cells(j, 2) = rs!sohieu + " - " + rs!tentk
            xlsheet.Cells(j, 3) = rs!ndk
            xlsheet.Cells(j, 4) = rs!cdk
            xlsheet.Cells(j, sotkco1 + 5 + sotkco2 + 2) = rs!nck
            xlsheet.Cells(j, sotkco1 + 5 + sotkco2 + 3) = rs!ckc
                                    
            If rs!n <> 0 Then
                For i = 1 To sotkco1
                    If tkid = TKCNKH_ID Or tkid = TKCNPT_ID Then
                        kq = PSDuKH(rs!MaKhachHang, rs!shtk, xlsheet.Cells(7, i + 4), tdau, tcuoi, -1)
                    Else
                        kq = PSDu(rs!sohieu, xlsheet.Cells(7, i + 4), tdau, tcuoi)
                    End If
                    If kq > 0 Then xlsheet.Cells(j, i + 4) = kq
                Next
            End If
            If rs!c <> 0 Then
                For i = 1 To sotkco2
                    If tkid = TKCNKH_ID Or tkid = TKCNPT_ID Then
                        kq = PSDuKH(rs!MaKhachHang, rs!shtk, xlsheet.Cells(7, i + 4), tdau, tcuoi, 1)
                    Else
                        kq = PSDu(xlsheet.Cells(7, sotkco1 + 5 + i), rs!sohieu, tdau, tcuoi)
                    End If
                    If kq > 0 Then xlsheet.Cells(j, sotkco1 + 5 + i) = kq
                Next
            End If
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        
        sotkno2 = j - 7
    Else
        sotkno2 = 0
    End If
    rs.Close
    Set rs = Nothing
    
    xlsheet.Cells(sotkno2 + 8, 2) = ABCtoVNI("Tæng PS")
    rx = CStr(sotkno2 + 8)
    sql = XLSCol(sotkco1 + 5)
    If sotkco1 > 0 Then
        For i = 1 To sotkno2
            xlsheet.Range(sql + CStr(i + 7)).Formula = "=SUM(E" + CStr(i + 7) + ":" + XLSCol(sotkco1 + 4) + CStr(i + 7) + ")"
        Next
    End If
    If sotkco2 > 0 Then
        For i = 1 To sotkno2
            xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 1) + CStr(i + 7)).Formula = "=SUM(" + XLSCol(sotkco1 + 6) + CStr(i + 7) + ":" + XLSCol(sotkco1 + 5 + sotkco2) + CStr(i + 7) + ")"
        Next
    End If
    For i = 1 To sotkco1
        sql = XLSCol(i + 4)
        xlsheet.Range(sql + CStr(sotkno2 + 8)).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno2 + 7) + ")"
    Next
    
    For i = 1 To sotkco2
        sql = XLSCol(i + sotkco1 + 5)
        xlsheet.Range(sql + CStr(sotkno2 + 8)).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno2 + 7) + ")"
    Next
        
    If sotkno2 > 0 Then
        For i = 1 To sotkco1 + sotkco2 + 6
            sql = XLSCol(i + 2)
            xlsheet.Range(sql + rx).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno2 + 7) + ")"
        Next
    End If
        
    xlsheet.Range("A6", cx + "6").HorizontalAlignment = xlHAlignCenter
    xlsheet.Range("A7", "A" + rx).HorizontalAlignment = xlHAlignCenter
    xlsheet.Range(XLSCol(sotkco1 + 6) + "7", XLSCol(sotkco1 + 6) + rx).HorizontalAlignment = xlHAlignCenter
    
    xlsheet.Range("A3", cx + "3").MergeCells = True
    xlsheet.Range("A4", cx + "4").MergeCells = True
        
    xlsheet.Range("A5", cx + "6").Font.Bold = True
    xlsheet.Range("A" + rx, cx + rx).Font.Bold = True
        
    With xlsheet.Range("A6", cx + rx)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "NK10.XLS"
KetThuc:
End Sub

Private Sub BK4(tdau As Integer, tcuoi As Integer, wsql As String, loai As Integer)
Dim i As Integer, j As Integer, kq As Double, rs As Recordset, sotkno As Integer, stt As Integer
Dim sqln As String, sql As String
    
    Recycle pCurDir + "BK4.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\BK4.XLS", pCurDir + "BK4.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "BK4.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = "B¶ng kª sè " + CStr(loai)
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqln = sqln + " + No_" + CStr(i)
    Next
    
    sql = "SELECT SoHieu,Ten" + IIf(nn > 0, "E", "") + " AS TenTK,Cap,Sum(" + sqln + ") AS N FROM HethongTK WHERE (" + wsql + ") AND (" + sqln + "<>0) GROUP BY SoHieu,Ten,Cap"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 7
' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        GauGe.Value = 0
        rs.MoveFirst
        Do While Not rs.EOF
            sotkno = sotkno + 1
            j = j + 1
            If rs!cap = 1 Then
                stt = stt + 1
                xlsheet.Cells(j, 1) = CStr(stt)
            End If
            
            xlsheet.Cells(j, 2) = rs!sohieu
            xlsheet.Cells(j, 3) = rs!tentk
            
            For i = 1 To 15
                kq = PSDu(rs!sohieu, xlsheet.Cells(6, i + 3), tdau, tcuoi)
                If kq > 0 Then xlsheet.Cells(j, i + 3) = kq
            Next
            
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        
        sotkno = j - 7
    Else
        sotkno = 0
    End If
    
    For i = 1 To sotkno
        xlsheet.Range("S" + CStr(i + 7)).Formula = "=SUM(D" + CStr(i + 7) + ":R" + CStr(i + 7) + ")"
    Next
    
    With xlsheet.Range("A6", "S" + CStr(sotkno + 7))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "BK4.XLS"
KetThuc:
End Sub

Private Sub BK11(tdau As Integer, tcuoi As Integer, TK As String, cap As Integer)
Dim rs As Recordset, i As Integer, j As Integer, sql As String, kq As Double
Dim sotkco1 As Integer, sotkno2 As Integer, sotkco2 As Integer, rx As String, cx As String, lx As Integer
Dim sqln As String, sqlc As String, st As String

    lx = Len(TK)
    Recycle pCurDir + "BK11.XLS"
    On Error GoTo KetThuc
    Set xlapp = CreateObject("Excel.Application")
    FileCopy pCurDir + "REPORTS\NK5.XLS", pCurDir + "BK11.XLS"
    xlapp.Workbooks.Open pCurDir + "BK11.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = "B¶ng kª sè 11"
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    xlsheet.Cells(6, 5) = "Ghi nî tµi kho¶n " + TK + ", ghi cã c¸c tµi kho¶n"
    
    XDSoHieuCap cap
    ' Liet ke theo hang ngang (phat sinh co)
    ' Lay danh sach tai khoan tai chinh
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON TK.SoHieu LIKE BaoCaoCP.SoHieu +'*'" _
        & " WHERE SoPS<>0 AND LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    i = 4
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco1 = i - 4

    xlsheet.Cells(6, sotkco1 + 5) = "Céng nî"
    
    sql = "SELECT BaoCaoCP.SoHieu AS ShCo FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON HethongTK.SoHieu LIKE BaoCaoCP.SoHieu +'*'" _
        & " WHERE SoPS<>0 AND LEFT(TK.SoHieu," + CStr(lx) + ")='" + TK + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    
    xlsheet.Cells(6, sotkco1 + 6) = "Ghi cã tµi kho¶n " + TK + ", ghi nî c¸c tµi kho¶n"
    
' Liet ke theo hang ngang (phat sinh co)
    i = sotkco1 + 5
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7, i) = rs!shco
        rs.MoveNext
    Loop
    sotkco2 = i - (sotkco1 + 5)
    
    xlsheet.Cells(6, sotkco1 + 5 + sotkco2 + 1) = "Céng cã"
    xlsheet.Cells(6, sotkco1 + 5 + sotkco2 + 2) = "Sè d­ cuèi th¸ng"
    xlsheet.Cells(7, sotkco1 + 5 + sotkco2 + 2) = "Nî"
    xlsheet.Cells(7, sotkco1 + 5 + sotkco2 + 3) = "Cã"
    cx = XLSCol(sotkco1 + 5 + sotkco2 + 3)
    
    xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 2) + "6", cx + "6").MergeCells = True
    xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 1) + "6", XLSCol(sotkco1 + 5 + sotkco2 + 1) + "7").MergeCells = True
    If sotkco2 > 1 Then xlsheet.Range(XLSCol(sotkco1 + 6) + "6", XLSCol(sotkco1 + 5 + sotkco2) + "6").MergeCells = True
    xlsheet.Range(XLSCol(sotkco1 + 5) + "6", XLSCol(sotkco1 + 5) + "7").MergeCells = True
    If sotkco1 > 1 Then xlsheet.Range("E6", XLSCol(sotkco1 + 4) + "6").MergeCells = True
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqln = sqln + " + SoDuKhachHang.No_" + CStr(i)
        sqlc = sqlc + " + SoDuKhachHang.Co_" + CStr(i)
    Next
    st = CStr(CThangDB(ThangTruoc(tdau)))
    sql = "SELECT KhachHang.MaSo,KhachHang.SoHieu,KhachHang.Ten,Sum(SoDuKhachHang.DuNo_" + st + ") AS ndk,Sum(SoDuKhachHang.DuCo_" + st + ") AS cdk,Sum(SoDuKhachHang.DuNo_" + CStr(CThangDB(tcuoi)) + ") AS nck,Sum(SoDuKhachHang.DuCo_" + CStr(CThangDB(tcuoi)) + ") AS ckc,Sum(" + sqln + ") AS N,Sum(" + sqlc + ") AS C FROM (SoDuKhachHang INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo) INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo" _
        & " WHERE LEFT(HethongTK.SoHieu," + CStr(lx) + ")='" + TK + "' AND (SoDuKhachHang.DuNo_" + st + "<>0 OR SoDuKhachHang.DuCo_" + st + "<>0 OR " + sqln + "<>0 OR " + sqlc + "<>0) GROUP BY KhachHang.MaSo,KhachHang.SoHieu,KhachHang.Ten"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    j = 7
' Liet ke theo hang doc (phat sinh no)
    If rs.RecordCount > 0 Then
        rs.MoveLast
        GauGe.Max = rs.RecordCount
        GauGe.Value = 0
        rs.MoveFirst
        Do While Not rs.EOF
            sotkno2 = sotkno2 + 1
            j = j + 1
            xlsheet.Cells(j, 1) = CStr(j - 7)
            xlsheet.Cells(j, 2) = rs!sohieu + " - " + rs!Ten
            xlsheet.Cells(j, 3) = rs!ndk
            xlsheet.Cells(j, 4) = rs!cdk
            xlsheet.Cells(j, sotkco1 + 5 + sotkco2 + 2) = rs!nck
            xlsheet.Cells(j, sotkco1 + 5 + sotkco2 + 3) = rs!ckc
            If rs!n <> 0 Then
                For i = 1 To sotkco1
                    kq = PSDuKH(rs!MaSo, TK, xlsheet.Cells(7, i + 4), tdau, tcuoi, -1)
                    If kq > 0 Then xlsheet.Cells(j, i + 4) = kq
                Next
            End If
            If rs!c <> 0 Then
                For i = 1 To sotkco2
                    kq = PSDuKH(rs!MaSo, TK, xlsheet.Cells(7, sotkco1 + 5 + i), tdau, tcuoi, 1)
                    If kq > 0 Then xlsheet.Cells(j, sotkco1 + 5 + i) = kq
                Next
            End If
            GauGe.Value = GauGe.Value + 1
            rs.MoveNext
        Loop
        
        sotkno2 = j - 7
    Else
        sotkno2 = 0
    End If
    xlsheet.Cells(sotkno2 + 8, 2) = ABCtoVNI("Tæng PS")
    rx = CStr(sotkno2 + 8)
    sql = XLSCol(sotkco1 + 5)
    If sotkco1 > 0 Then
        For i = 1 To sotkno2
            xlsheet.Range(sql + CStr(i + 7)).Formula = "=SUM(E" + CStr(i + 7) + ":" + XLSCol(sotkco1 + 4) + CStr(i + 7) + ")"
        Next
    End If
    If sotkco2 > 0 Then
        For i = 1 To sotkno2
            xlsheet.Range(XLSCol(sotkco1 + 5 + sotkco2 + 1) + CStr(i + 7)).Formula = "=SUM(" + XLSCol(sotkco1 + 6) + CStr(i + 7) + ":" + XLSCol(sotkco1 + 5 + sotkco2) + CStr(i + 7) + ")"
        Next
    End If
    For i = 1 To sotkco1
        sql = XLSCol(i + 4)
        xlsheet.Range(sql + CStr(sotkno2 + 8)).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno2 + 7) + ")"
    Next
    
    For i = 1 To sotkco2
        sql = XLSCol(i + sotkco1 + 5)
        xlsheet.Range(sql + CStr(sotkno2 + 8)).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno2 + 7) + ")"
    Next
        
    If sotkno2 > 0 Then
        For i = 1 To sotkco1 + sotkco2 + 6
            sql = XLSCol(i + 2)
            xlsheet.Range(sql + rx).Formula = "=SUM(" + sql + "8:" + sql + CStr(sotkno2 + 7) + ")"
        Next
    End If
        
    xlsheet.Range("A6", cx + "6").HorizontalAlignment = xlHAlignCenter
    xlsheet.Range("A7", "A" + rx).HorizontalAlignment = xlHAlignCenter
    xlsheet.Range(XLSCol(sotkco1 + 6) + "7", XLSCol(sotkco1 + 6) + rx).HorizontalAlignment = xlHAlignCenter
    
    xlsheet.Range("A3", cx + "3").MergeCells = True
    xlsheet.Range("A4", cx + "4").MergeCells = True
        
    xlsheet.Range("A5", cx + "6").Font.Bold = True
    xlsheet.Range("A" + rx, cx + rx).Font.Bold = True
        
    With xlsheet.Range("A6", cx + rx)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "BK11.XLS"
KetThuc:
    rs.Close
End Sub

Private Function InTongHopHD(tdau As Integer, tcuoi As Integer) As Boolean
    Dim sql As String

    InTongHopHD = False

    sql = "SELECT DISTINCTROW ChungTu.*,DoituongCT.* AS HD,HethongTK.SoHieu AS TKNo,TK.Sohieu AS TKCo,-1 AS ID,KhachHang.* FROM (((ChungTu INNER JOIN DoituongCT ON ChungTu.MaDT=DoituongCT.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) INNER JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo" _
        & " WHERE Not((TK.SoHieu LIKE '" + vatr + "*' OR TK.TK_ID=" + CStr(TKDT_ID) + ")) AND (" + WThang("ThangCT", tdau, tcuoi) + ") AND ChungTu.MaDT>1" _
        & " UNION SELECT DISTINCTROW ChungTu.*,DoituongCT.*,HethongTK.SoHieu AS TKNo,TK.Sohieu AS TKCo,1 AS ID,KhachHang.* FROM (((ChungTu INNER JOIN DoituongCT ON ChungTu.MaDT=DoituongCT.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) INNER JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo" _
        & " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND ((TK.SoHieu LIKE '" + vatr + "*' OR TK.TK_ID=" + CStr(TKDT_ID) + " OR ((HethongTK.SoHieu LIKE '11*') AND (Not TK.SoHieu LIKE '331*')))) AND MaDT>1"
        
    SetSQL "QNhatKy", sql

    frmMain.Rpt.ReportFileName = "BKHD.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi)
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    InTongHopHD = True
End Function

Private Function CTGhiSo(mgs As Long, tdau As Integer, tcuoi As Integer, dktg As Integer, ndau As Date, ncuoi As Date, VV As Long) As Boolean
    Dim sql As String, sqlw As String
    If dktg = 0 Then
        sqlw = WThang("ThangCT", tdau, tcuoi)
        frmMain.Rpt.Formulas(4) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
        RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    Else
        sqlw = WNgay("NgayGS", ndau, ncuoi)
        frmMain.Rpt.Formulas(4) = "ThoiGian='" + ThoiGianN(ndau, ncuoi, nn) + "'"
        RptSetDate ncuoi, nn
    End If
    sql = "SELECT ChungTu.MaCT AS M, First(ChungTu.ThangCT) AS T, ChungTu.SoHieu AS SH, ChungTu.NgayCT AS NCT, ChungTu.NgayGS AS NGS, First(ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ") AS DG, HeThongTK.SoHieu AS TKNo,First(HethongTK.Ten" + IIf(nn > 0, "E", "") + ") AS TNo, HeThongTK_1.SoHieu AS TKCo,First(HethongTK_1.Ten" + IIf(nn > 0, "E", "") + ") AS TCo, Sum(ChungTu.SoPS) AS PS" _
        & " FROM (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) LEFT JOIN HeThongTK AS HeThongTK_1 ON ChungTu.MaTKCo = HeThongTK_1.MaSo WHERE SoPS<>0 AND " + sqlw + " AND CTGS=" + CStr(mgs) + " AND (IsNull(HethongTK.MaSo) OR HethongTK.Loai>0) AND (IsNull(HethongTK_1.MaSo) OR HethongTK_1.Loai>0) " + IIf(VV > 1, " AND ChungTu.MaDT=" + CStr(VV), "") _
        & " GROUP BY MaCT,ChungTu.SoHieu, NgayCT,NgayGS,ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ",HeThongTK.SoHieu,HeThongTK_1.SoHieu ORDER BY NgayCT, ChungTu.SoHieu"
        
    SetSQL "QNhatKy", sql
    frmMain.Rpt.ReportFileName = "CTGHISO.RPT"
        
    frmMain.Rpt.Formulas(6) = "DG='" + SelectSQL("SELECT SoHieu+' - '+DienGiai AS F1 FROM CTGhiSo WHERE MaSo=" + CStr(CTGS.ItemData(CTGS.ListIndex))) + "'"
    frmMain.Rpt.Formulas(7) = "DG2='" + TenVV(VV) + "'"
End Function

Private Function InTongHopPhi(shtk As String, tdau As Integer, tcuoi As Integer) As Boolean
    Dim rs As Recordset, sql As String, i As Integer, st As String
    
    sql = "SELECT DISTINCTROW TK.SoHieu AS SHC FROM " + ChungTu2TKNC(10) _
        & " WHERE (HethongTK.SoHieu LIKE '" + shtk + "*') AND " + WThang("ThangCT", tdau, tcuoi) + " AND SoPS<>0  GROUP BY TK.SoHieu HAVING Sum(ChungTu.SoPS)<>0"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.RecordCount = 0 Then
        InTongHopPhi = False
        GoTo KT
    End If
    Do While Not rs.EOF And i < 10
        i = i + 1
        frmMain.Rpt.Formulas(2 + i) = "TK" + CStr(i) + "='" + rs!shc + "'"
        rs.MoveNext
    Loop
    frmMain.Rpt.Formulas(3 + i) = "SoHieuTK='" + shtk + "'"
    frmMain.Rpt.Formulas(4 + i) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    
    st = "Sum("
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = st + "+HethongTK.No_" + CStr(i)
    Next
    st = st + ")<>0"
    
    sql = "SELECT DISTINCTROW HethongTK.SoHieu AS SHN,First(HethongTK.Cap) AS CapCP,First(HethongTK.TKCon) AS TKConCP,First(HethongTK.Ten" + IIf(nn > 0, "E", "") + ") AS TenCP,TK.SoHieu AS SHC,Sum(ChungTu.SoPS) AS TPS FROM ((ChungTu INNER JOIN HethongTK AS TKN ON ChungTu.MaTKNo=TKN.MaSo) INNER JOIN HethongTK ON LEFT(TKN.SoHieu,Len(HethongTK.SoHieu)) = HethongTK.SoHieu) INNER JOIN HethongTK AS TK ON ChungTu.MaTKTCCo=TK.MaSo" _
        & " WHERE (HethongTK.SoHieu LIKE '" + shtk + "*') AND (ISNULL(ChungTu.MaSo) OR (" + WThang("ThangCT", tdau, tcuoi) + " AND SoPS<>0))  GROUP BY HethongTK.SoHieu,TK.SoHieu HAVING Sum(SoPS)<>0 OR First(HethongTK.TKCon>0) AND " + st
    SetSQL "QSoCai", sql
        
    frmMain.Rpt.ReportFileName = "THCP.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    InTongHopPhi = True
KT:
    rs.Close
    Set rs = Nothing
End Function

Private Sub CTGhiSoTH(TK As ClsTaikhoan, tdau As Integer, tcuoi As Integer, dktg As Integer, ndau As Date, ncuoi As Date)
    Dim sql As String, sqlw As String
    
    If dktg = 0 Then
        sqlw = WThang("ThangCT", tdau, tcuoi)
        frmMain.Rpt.Formulas(4) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    Else
        sqlw = WNgay("NgayGS", ndau, ncuoi)
        frmMain.Rpt.Formulas(4) = "ThoiGian='" + ThoiGianN(ndau, ncuoi, nn) + "'"
    End If
    
    sql = "SELECT ChungTu.MaCT AS M, ChungTu.ThangCT AS T, ChungTu.SoHieu AS SH, ChungTu.NgayCT AS NCT, ChungTu.NgayGS AS NGS, ChungTu.DienGiai AS DG, HeThongTK.SoHieu AS TKNo,HethongTK.Ten AS TNo, HeThongTK_1.SoHieu AS TKCo,HethongTK_1.Ten AS TCo, ChungTu.SoPS AS PS, CTGhiSo.SoHieu AS SHGS, CTGhiSo.DienGiai AS DGGS" _
        & " FROM ((HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) INNER JOIN HeThongTK AS HeThongTK_1 ON ChungTu.MaTKCo = HeThongTK_1.MaSo) INNER JOIN CTGhiSo ON ChungTu.CTGS=CTGhiSo.MaSo WHERE " + sqlw + " AND HethongTK_1.SoHieu LIKE '" + TK.sohieu + "*' AND SoPS<>0 ORDER BY NgayCT, ChungTu.SoHieu"
    SetSQL "QNhatKy", sql
    
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    If TK.MaSo > 0 Then
        frmMain.Rpt.Formulas(3) = "TK='" + TK.sohieu + " - " + TK.Ten + "'"
        frmMain.Rpt.ReportFileName = "CTGSTH.RPT"
    Else
        frmMain.Rpt.ReportFileName = "CTGSTH2.RPT"
    End If
End Sub

Private Sub InQTThue(tdau As Integer, tcuoi As Integer)
    InThue tdau, tcuoi
    frmMain.Rpt.Formulas(30) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    frmMain.Rpt.Formulas(31) = "MST = '" + frmMain.LbCty(8).Caption + "'"
    frmMain.Rpt.Formulas(32) = "Tel='" + frmMain.LbCty(3).Caption + "'"
    frmMain.Rpt.Formulas(33) = "Fax='" + frmMain.LbCty(4).Caption + "'"
    frmMain.Rpt.ReportFileName = "XTHUE.RPT"
    RptSetDate NgayCuoiThang(pNamTC, 12), nn
    
    frmMain.Rpt.Formulas(34) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
End Sub

Private Sub InQTVAT(tdau As Integer, tcuoi As Integer)
    Dim ts As Double, v As Double
    Dim TK As New ClsTaikhoan
    
    ToKhaiVAT tdau, tcuoi, TK
    'v = SelectSQL("SELECT Sum(SoPS) AS F1,Sum(ThanhTien) AS F2 FROM HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo WHERE Loai=-1 AND TS=1 AND HD=1 AND " + WThang("ThangCT", tdau, tcuoi), ts)
    'frmMain.Rpt.Formulas(49) = "TS = " + CStr(ts)
    'frmMain.Rpt.Formulas(50) = "TSVAT = " + CStr(v)
    frmMain.Rpt.ReportFileName = "XTOKHAI.RPT"
    
    frmMain.Rpt.Formulas(48) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    Set TK = Nothing
End Sub

Private Sub InQTTTN(tdau As Integer, tcuoi As Integer)
    Dim i As Integer, j As Integer, ps As Double
    
    Recycle pCurDir + "QTTNDN.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\QTTNDN.XLS", pCurDir + "QTTNDN.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "QTTNDN.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(5, 1) = xlsheet.Cells(5, 1) + ThoiGian(tdau, tcuoi, nn)
    xlsheet.Cells(7, 1) = xlsheet.Cells(7, 1) + frmMain.LbCty(8).Caption
    xlsheet.Cells(8, 1) = xlsheet.Cells(8, 1) + pTenCty
    xlsheet.Cells(9, 1) = xlsheet.Cells(9, 1) + frmMain.LbCty(2).Caption
    xlsheet.Cells(10, 1) = xlsheet.Cells(10, 1) + frmMain.LbCty(10).Caption + "                 " + ABCtoVNI("TØnh, thµnh phè: ") + frmMain.LbCty(11).Caption
    xlsheet.Cells(11, 1) = xlsheet.Cells(11, 1) + frmMain.LbCty(3).Caption + "                 Fax: " + frmMain.LbCty(4).Caption + "                 Email: " + frmMain.LbCty(9).Caption
    
    For i = 15 To 120
        j = 7
        ps = 0
        Do While Len(xlsheet.Cells(i, j)) > 0
            If Len(xlsheet.Cells(i, 6)) > 0 Then
                ps = ps + PSDu(xlsheet.Cells(i, 6), xlsheet.Cells(i, j), tdau, tcuoi) - PSDu(xlsheet.Cells(i, j), xlsheet.Cells(i, 6), tdau, tcuoi)
            Else
                ps = ps + SoPSTK(xlsheet.Cells(i, j), tdau, tcuoi, 1)
            End If
            xlsheet.Cells(i, 5) = ps
            j = j + 1
        Loop
    Next
    
    ps = SoDuTKSH("4211", ThangTruoc(tdau), -1)
    If ps > 0 Then xlsheet.Cells(i, 5) = ps
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "QTTNDN.XLS"
KetThuc:
End Sub

Private Sub BKNopThue(tdau As Integer, tcuoi As Integer)
    Dim sql As String
    
    sql = "SELECT ThangCT,LEFT(HethongTK.SoHieu,4) AS Thue,Sum(SoPS) AS NopThue FROM " + ChungTu2TKNC(0) _
        & " WHERE HethongTK.SoHieu LIKE '333*' AND (TK.SoHieu LIKE '11*' OR TK.SoHieu LIKE '336*' OR TK.SoHieu LIKE '331*') AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY ThangCT,LEFT(HethongTK.SoHieu,4)"
    SetSQL "MienTru", sql
    
    frmMain.Rpt.ReportFileName = "NOPTHUE.RPT"
    RptSetDate NgayCuoiThang(pNamTC, 12)
    frmMain.Rpt.Formulas(3) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    frmMain.Rpt.Formulas(4) = "MST = '" + frmMain.LbCty(8).Caption + "'"
    frmMain.Rpt.Formulas(5) = "Tel='" + frmMain.LbCty(3).Caption + "'"
    frmMain.Rpt.Formulas(6) = "Fax='" + frmMain.LbCty(4).Caption + "'"
    
    frmMain.Rpt.Formulas(7) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
End Sub

Private Sub QTDauVao(tdau As Integer, tcuoi As Integer)
    Dim d As Double, v As Double
    
    'frmMain.Rpt.Formulas(3) = "thang = " + CStr(1)
    'frmMain.Rpt.Formulas(4) = "ThangCuoi = " + CStr(tcuoi)
    frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    frmMain.Rpt.Formulas(6) = "MSThue = '" + frmMain.LbCty(8).Caption + "'"
    frmMain.Rpt.ReportFileName = "QTV.RPT"
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND TyLe=3 AND KCT=0 AND " + WThang("ThangCT", tdau, tcuoi), v)
    frmMain.Rpt.Formulas(7) = "D3=" + DoiDau(d)
    frmMain.Rpt.Formulas(8) = "V3=" + DoiDau(v)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND TyLe=5 AND KCT=0 AND " + WThang("ThangCT", tdau, tcuoi), v)
    frmMain.Rpt.Formulas(9) = "D5=" + DoiDau(d)
    frmMain.Rpt.Formulas(10) = "V5=" + DoiDau(v)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND TyLe=10 AND KCT=0 AND " + WThang("ThangCT", tdau, tcuoi), v)
    frmMain.Rpt.Formulas(11) = "D10=" + DoiDau(d)
    frmMain.Rpt.Formulas(12) = "V10=" + DoiDau(v)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND TyLe=2 AND KCT=0 AND " + WThang("ThangCT", tdau, tcuoi), v)
    frmMain.Rpt.Formulas(13) = "D2=" + DoiDau(d)
    frmMain.Rpt.Formulas(14) = "V2=" + DoiDau(v)
    
    frmMain.Rpt.Formulas(15) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
End Sub

Private Sub QTDauRa()
    Dim d As Double, v As Double
    
    'frmMain.Rpt.Formulas(3) = "thang = " + CStr(1)
    'frmMain.Rpt.Formulas(4) = "ThangCuoi = " + CStr(tcuoi)
    frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    frmMain.Rpt.Formulas(6) = "MSThue = '" + frmMain.LbCty(8).Caption + "'"
    frmMain.Rpt.ReportFileName = "QTR.RPT"
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1 FROM " + ChungTu2TKHD(1) + " WHERE SoPS=0 AND ThangCT<7 AND HoaDon.Loai=1 AND KCT=1")
    frmMain.Rpt.Formulas(7) = "Dk=" + DoiDau(d)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1 FROM " + ChungTu2TKHD(1) + " WHERE SoPS=0 AND ThangCT>6 AND HoaDon.Loai=1 AND KCT=1")
    frmMain.Rpt.Formulas(8) = "Dk2=" + DoiDau(d)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE SoPS=0 AND ThangCT<7 AND HoaDon.Loai=1 AND TyLe=0 AND KCT=0")
    frmMain.Rpt.Formulas(9) = "D0=" + DoiDau(d)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE SoPS=0 AND ThangCT>6 AND HoaDon.Loai=1 AND TyLe=0 AND KCT=0")
    frmMain.Rpt.Formulas(10) = "D02=" + DoiDau(d)
    
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE ThangCT<7 AND HoaDon.Loai=1 AND TyLe=5 AND KCT=0", v)
    frmMain.Rpt.Formulas(11) = "D5=" + DoiDau(d)
    frmMain.Rpt.Formulas(12) = "V5=" + DoiDau(v)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE ThangCT>6 AND HoaDon.Loai=1 AND TyLe=5 AND KCT=0", v)
    frmMain.Rpt.Formulas(13) = "D52=" + DoiDau(d)
    frmMain.Rpt.Formulas(14) = "V52=" + DoiDau(v)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE ThangCT<7 AND HoaDon.Loai=1 AND TyLe=10 AND KCT=0", v)
    frmMain.Rpt.Formulas(15) = "D10=" + DoiDau(d)
    frmMain.Rpt.Formulas(16) = "V10=" + DoiDau(v)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE ThangCT>6 AND HoaDon.Loai=1 AND TyLe=10 AND KCT=0", v)
    frmMain.Rpt.Formulas(17) = "D102=" + DoiDau(d)
    frmMain.Rpt.Formulas(18) = "V102=" + DoiDau(v)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE ThangCT<7 AND HoaDon.Loai=1 AND TyLe=20 AND KCT=0", v)
    frmMain.Rpt.Formulas(19) = "D20=" + DoiDau(d)
    frmMain.Rpt.Formulas(20) = "V20=" + DoiDau(v)
    d = SelectSQL("SELECT SUM(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE ThangCT>6 AND HoaDon.Loai=1 AND TyLe=20 AND KCT=0", v)
    frmMain.Rpt.Formulas(21) = "D202=" + DoiDau(d)
    frmMain.Rpt.Formulas(22) = "V202=" + DoiDau(v)
End Sub


Private Function InThueTTDB(tdau As Integer, tcuoi As Integer, tl As Integer, TK As ClsTaikhoan) As Boolean
    Dim sql As String

    GauGe.Max = 1
    
    Select Case tl
        Case -2:
            SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(TTDB_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,GiaTT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=2 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'ORDER BY NgayPH"
        Case -1:
            SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(TTDB_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,GiaTT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=2 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*' ORDER BY NgayPH"
        Case 30:
            SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(TTDB_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,GiaTT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=2 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*' ORDER BY NgayPH"
        Case Else
            SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(TTDB_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,GiaTT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=2 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*' ORDER BY NgayPH"
    End Select
    
    GauGe.Value = 1
    frmMain.Rpt.ReportFileName = "BANGKER2.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    frmMain.Rpt.Formulas(6) = "MSThue = '" + frmMain.LbCty(8).Caption + "'"
    frmMain.Rpt.Formulas(8) = "TLKT = " + CStr(tl)
    InThueTTDB = True
End Function


Private Function SoNKChuaThue(tdau As Integer, tcuoi As Integer, shct As String) As Double
    SetSQL "QNhatKy", "SELECT MaCT FROM " + ChungTu2TKNC(1) + " WHERE (HethongTK.SoHieu LIKE '33312*' OR HethongTK.SoHieu LIKE '" + pVATV + "*' OR HethongTK.SoHieu LIKE '11*') AND " + WThang("ThangCT", tdau, tcuoi)
    SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKNC(-1) + " WHERE (HethongTK.SoHieu LIKE '" + pVATV + "*' OR HethongTK.SoHieu LIKE '11*') AND " + WThang("ThangCT", tdau, tcuoi)
    SetSQL "QChitiet", "SELECT ChungTu.MaCT FROM ((ChungTu LEFT JOIN QNhatKy ON ChungTu.MaCT=QNhatKy.MaCT) LEFT JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '33332*' AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "' AND IsNull(QNhatKy.MaCT) AND IsNull(MienTru.MaCT) AND " + WThang("ThangCT", tdau, tcuoi)
    SoNKChuaThue = SelectSQL("SELECT Sum(SoPS) AS F1 FROM ChungTu INNER JOIN QChitiet ON ChungTu.MaCT=QChitiet.MaCT WHERE MaTKNo>0")
End Function

Private Sub BangKeBanRa(tdau As Integer, tcuoi As Integer, TK As ClsTaikhoan)
    GauGe.Max = 1
    SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKNC(1) + " WHERE TK_ID=" + CStr(TTDB_ID) + " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY MaCT"
   SetSQL "QNhatKy", "SELECT ChungTu.SoHieu,NgayCT,Vattu.TenVattu,SoPS2Co,SoPS FROM ((ChungTu INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT " _
        & " WHERE TK_ID=" + CStr(TKDT_ID)
    frmMain.Rpt.ReportFileName = "BANGKER3.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    frmMain.Rpt.Formulas(6) = "MSThue = '" + frmMain.LbCty(8).Caption + "'"
    GauGe.Value = 1
End Sub

Private Sub ToKhaiTTDB(tdau As Integer, tcuoi As Integer, TK As ClsTaikhoan)
    Dim kn As Double, st As String
    
    GauGe.Max = 2
    SetSQL "MienTru", "SELECT MaCT, First(TyLe) AS TL FROM (ChungTu INNER JOIN HoaDon ON ChungTu.MaSo=HoaDon.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo " _
        & " WHERE HoaDon.Loai=2 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*' GROUP BY MaCT"
    SetSQL "QNhatKy", "SELECT Vattu.SoHieu,TenVattu,DonVi,Sum(SoPS2Co) AS Luong,Sum(SoPS) AS Tien,First(TL) AS TyLe " _
        & " FROM ((ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKDT_ID) + " GROUP BY Vattu.SoHieu,TenVattu,DonVi"
    GauGe.Value = 1
    
    SetSQL "QChitiet", "SELECT MaCT FROM " + ChungTu2TKNC(-1) _
        & " WHERE MaLoai=8 AND " + WThang("ThangCT", tdau, tcuoi) + " AND TK_ID=" + CStr(TKDT_ID) + " GROUP BY MaCT"
    
    If SelectSQL("SELECT QChitiet.MaCT AS F1 FROM MienTru INNER JOIN QChitiet ON MienTru.MaCT=QChitiet.MaCT") > 0 Then
        frmMain.Rpt.ReportFileName = "TOKHAIDB.RPT"
    Else
        frmMain.Rpt.ReportFileName = "TOKHAID2.RPT"
    End If
    
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    frmMain.Rpt.Formulas(6) = "MSThue = '" + frmMain.LbCty(8).Caption + "'"
    st = CStr(CThangDB(ThangTruoc(tdau)))
    frmMain.Rpt.Formulas(7) = "KT = " + DoiDau(SelectSQL("SELECT Sum(DuCo_" + st + "-DuNo_" + st + ") AS F1 FROM HethongTK WHERE TKCon=0 AND TK_ID=" + CStr(TTDB_ID) + " AND SoHieu LIKE '" + TK.sohieu + "*'"))
    kn = SelectSQL("SELECT Sum(DuCo_" + CStr(CThangDB(tcuoi)) + "-DuNo_" + CStr(CThangDB(tcuoi)) + ") AS F1 FROM HethongTK WHERE TKCon=0 AND TK_ID=" + CStr(TTDB_ID) + " AND SoHieu LIKE '" + TK.sohieu + "*'")
    frmMain.Rpt.Formulas(8) = "KN = " + DoiDau(kn)
    frmMain.Rpt.Formulas(9) = "BC = '" + ToVNText(kn) + " ®ång'"
    frmMain.Rpt.Formulas(10) = "TThue = " + DoiDau(SelectSQL("SELECT Sum(SoPS) AS F1 FROM (ChungTu INNER JOIN HoaDon ON ChungTu.MaSo=HoaDon.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo " _
        & " WHERE HoaDon.Loai=2 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'"))
End Sub

Private Sub SoDangKyCT(tdau As Integer, tcuoi As Integer, dktg As Integer, ndau As Date, ncuoi As Date)
    Dim sqlw As String
    
    If dktg = 0 Then
        sqlw = WThang("ThangCT", tdau, tcuoi)
        frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    Else
        sqlw = WNgay("NgayGS", ndau, ncuoi)
        frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGianN(ndau, ncuoi, nn) + "'"
    End If
    SetSQL "QNhatKy", "SELECT CTGhiSo.SoHieu, CTGhiSo.DienGiai" + IIf(nn > 0, "E", "") + ",Sum(SoPS) AS TPS FROM (" + ChungTu2TKNC(-1) + ") INNER JOIN CTGhiSo ON ChungTu.CTGS=CTGhiSo.MaSo WHERE " + sqlw + " AND HethongTK.Loai>0 GROUP BY CTGhiSo.SoHieu, CTGhiSo.DienGiai"
    
    frmMain.Rpt.ReportFileName = "SODKCT.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(5) = "PSLK = " + DoiDau(SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE " + WThang("ThangCT", 0, tcuoi) + " AND HethongTK.Loai>0"))
End Sub

Private Function BKChiTietRPT(TK As String, ndau As Date, ncuoi As Date) As Boolean
    Dim rs As Recordset, sql As String, i As Integer
        
    sql = "SELECT DISTINCTROW LEFT(HethongTK.SoHieu,4) AS SHN FROM " + ChungTu2TKNC(-1) _
        & " WHERE (HethongTK.SoHieu LIKE '" + TK + "*') AND " + WNgay("NgayGS", ndau, ncuoi) + " AND SoPS<>0 GROUP BY LEFT(HethongTK.SoHieu,4)"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.RecordCount = 0 Then GoTo KgIn
    Do While Not rs.EOF And i < 7
        i = i + 1
        frmMain.Rpt.Formulas(2 + i) = "TK" + CStr(i) + "='" + rs!SHN + "'"
X1:
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    frmMain.Rpt.Formulas(3 + i) = "SoHieuTK='" + TK + "'"
    
    frmMain.Rpt.Formulas(4 + i) = "ThoiGian ='" + ThoiGianN(ndau, ncuoi, nn) + "'"
        
    sql = "SELECT DISTINCTROW LEFT(HethongTK.SoHieu,4) AS SHN,LEFT(TK.SoHieu,4) AS SHC, NgayGS,ChungTu.SoHieu,NgayCT,SoPS,IIF(HethongTK.SoHieu LIKE '" + TK + "*',LEFT(TK.SoHieu,3),LEFT(HethongTK.SoHieu,3)) AS TKDU,ChungTu.DienGiai FROM " + ChungTu2TKNC(0) _
        & " WHERE (HethongTK.SoHieu LIKE '" + TK + "*') AND " + WNgay("NgayGS", ndau, ncuoi) + " AND SoPS<>0 ORDER BY NgayCT"
    SetSQL "QSoCai", sql
    frmMain.Rpt.ReportFileName = "BKCT.RPT"
    RptSetDate ncuoi, nn
    
    BKChiTietRPT = True
    Exit Function
KgIn:
    BKChiTietRPT = False
    rs.Close
End Function

Private Sub ChiTietDoanhThu(tdau As Integer, tcuoi As Integer)
    Dim st As String, sqlc As String, i As Integer
        
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqlc = sqlc + "+Co_" + CStr(i)
    Next
        
    GauGe.Max = 2
    SetSQL "QTongHopCT", "SELECT DISTINCTROW SoHieu, First(Cap) AS FirstOfCap, First(Ten" + IIf(nn > 0, "E", "") + ") AS FirstOfTen, First(Kieu) AS FirstOfKieu, First(Loai) AS FirstOfLoai, First(MaTC = MaSo OR MaTC = 0) AS FirstOfChuanLa," _
        & " Sum(" + sqlc + ") AS PsCo, First(CapDuoi) AS CD From HeThongTK WHERE Cap>1 GROUP BY SoHieu Having (SoHieu LIKE '511*')"
    
    GauGe.Value = 1
    XDCapTK "511"
    frmMain.Rpt.ReportFileName = "DOANHTHU.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(5) = "PSCo=" + DoiDau(SoPSTK("511", tdau, tcuoi, 1))
    GauGe.Value = 2
End Sub

Private Sub ToKhaiVAT2(thang As Integer, taikhoan As ClsTaikhoan)
    Dim i As Integer, prg As Long, TK As New ClsTaikhoan, sql As String, shct As String, vr As Double, VV As Double, sotien As Double, vvx As Double
    Dim dkn As Double, dkc As Double, dknt As Double, KT As Double, v As Double, v1 As Double, tsv As Double, nkv As Double
    Dim f1 As Integer, mvpath As String
    
    f1 = 0
    mvpath = KiemTraMaVach(2, CurrentDrive + "\Hcmtax\PTKM1-2005\VAT.exe")
    If Len(Dir(mvpath)) = 0 Then GoTo KT
    
    prg = Shell(mvpath, vbNormalFocus)
    If prg = 0 Then Exit Sub
    AppActivate prg
    AppIdle w1
B:
    SendKeys "{F9}"
    AppIdle w1
    ' chon thang
    SendKeys "{PGUP}"
    For i = 1 To thang - 1
        SendKeys "{DOWN}"
    Next
    AppIdle w1
    SendKeys "{TAB}{PGUP}"
    For i = 1 To pNamTC - 2000
        SendKeys "{DOWN}"
    Next
    AppIdle w1
    SendKeys "{TAB}"
    AppIdle w1
    SendKeys Chr(32)
    AppIdle w1
    SendKeys Chr(32)
    AppIdle w1
    
    If f1 = 1 Then GoTo a
    SendKeys "{F12}"
    AppIdle w1
    SendKeys Chr(32)
    AppIdle w1
    SendKeys Chr(32)
    AppIdle w1
    f1 = 1
    GoTo B
    
a:
    SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
    AppIdle w1
    
    ' tong ban ra
    TK.InitTaikhoanSohieu "33312"
    shct = taikhoan.SHChiTiet()
    If TK.tk_id = GTGTKT_ID Then
        sql = "DELETE HoaDon.*  FROM ((HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo " _
            & " WHERE (HethongTK.SoHieu LIKE '" + pVATV + "*') AND (TK.SoHieu LIKE '33312*'  OR InStr(ChungTu.GhiChu,'33312')>0) AND HoaDon.Loai=-1"
        ExecuteSQL5 sql
    End If
    
    KT = DKToKhai(thang, shct)
    If KT > 0 Then
        SendKeys CStr(KT)
        AppIdle w1
    End If
    SendKeys "{TAB}"
    AppIdle w1
    
    ' Tong mua vao
    KT = SoNKChuaThue(thang, thang, shct)
    If Len(shct) > 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND DC=0 AND " + WThang("ThangCT", thang, thang) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    Else
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND DC=0 AND " + WThang("ThangCT", thang, thang)
    End If
    vvx = SelectSQL(sql, v1) + KT
    SendKeys CStr(vvx) + "{TAB}"
    AppIdle w1
    v = vvx
        
    If TK.tk_id = GTGTKT_ID Then
        SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKNC(-1) + " WHERE (HethongTK.SoHieu LIKE '521*' OR HethongTK.SoHieu LIKE '531*') AND " + WThang("ThangCT", thang, thang)
        sql = "SELECT SUM(SoPS) AS F1 FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT " _
            & " WHERE IsNull(MienTru.MaCT) AND (HethongTK.SoHieu LIKE '33312*') AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "' AND " + WThang("ThangCT", thang, thang) + " AND (TK.SoHieu LIKE '11*' OR TK.SoHieu LIKE '136*')"
        v1 = v1 - SelectSQL(sql) + PSDu(pVATV, "33312", thang, thang, shct)
    End If
    SendKeys CStr(v1) + "{TAB}"
    AppIdle w1
    
    ' Nhap khau
    If Len(shct) > 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND NK=1 AND " + WThang("ThangCT", thang, thang) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    Else
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND NK=1 AND " + WThang("ThangCT", thang, thang)
    End If
    SendKeys CStr(SelectSQL(sql, nkv)) + "{TAB}" + CStr(nkv) + "{TAB}"
    AppIdle w1
    
    ' Tai san
    If Len(shct) > 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND TS=1 AND " + WThang("ThangCT", thang, thang) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    Else
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND TS=1 AND " + WThang("ThangCT", thang, thang)
    End If
    SendKeys CStr(SelectSQL(sql, tsv)) + "{TAB}" + CStr(tsv) + "{TAB}"
    AppIdle w1
    
    ' Tong vao
    If Len(shct) > 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1, SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND DC=0 AND KCT=0 AND " + WThang("ThangCT", thang, thang) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    Else
        sql = "SELECT SUM(ThanhTien) AS F1, SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND DC=0 AND KCT=0 AND " + WThang("ThangCT", thang, thang)
    End If
    vvx = SelectSQL(sql, v1) + KT
    SendKeys CStr(vvx) + "{TAB}" + CStr(v1) + "{TAB}"
    AppIdle w1
    
    sql = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
            & " WHERE (HethongTK.SoHieu LIKE '142*' OR HethongTK.SoHieu LIKE '242*' OR HethongTK.SoHieu LIKE '6*') AND (TK.SoHieu LIKE '" + pVATV + "*') AND RIGHT(TK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "' AND " + WThang("ThangCT", thang, thang)
    v1 = v1 - SelectSQL(sql)
    SendKeys CStr(v1) + "{TAB}"
    AppIdle w1
    
    'Dieu chinh
    If Len(shct) > 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND DC=1 AND " + WThang("ThangCT", thang, thang) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    Else
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND DC=1 AND " + WThang("ThangCT", thang, thang)
    End If
    vvx = SelectSQL(sql, VV)
    If VV >= 0 Then
        SendKeys CStr(vvx) + "{TAB}" + CStr(VV) + "{TAB}{TAB}{TAB}{TAB}"
    Else
        SendKeys "{TAB}{TAB}" + CStr(-vvx) + "{TAB}" + CStr(-VV) + "{TAB}{TAB}"
    End If
    AppIdle w1
        
    If v = 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", thang, thang) + " AND DC=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
        If SelectSQL(sql) = 0 Then
            SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}" + Chr(32)
            AppIdle w1
            GoTo KT
        End If
    Else
        SendKeys Chr(32)
        AppIdle w1
        SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
        AppIdle w1
    End If
    
    sql = "SELECT SUM(ThanhTien) AS F1 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", thang, thang) + " AND KCT=1 AND DC=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    SendKeys CStr(SelectSQL(sql)) + "{TAB}"
    AppIdle w1
    SendKeys CStr(SelectSQL("SELECT Sum(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND DC=0 AND " + WThang("ThangCT", thang, thang) + " AND KCT=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "' AND TyLe=0")) + "{TAB}"
    AppIdle w1
    SendKeys CStr(SelectSQL("SELECT Sum(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND DC=0 AND " + WThang("ThangCT", thang, thang) + " AND KCT=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "' AND TyLe=5", v)) + "{TAB}"
    AppIdle w1
    SendKeys CStr(v) + "{TAB}"
    AppIdle w1
    SendKeys CStr(SelectSQL("SELECT Sum(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND DC=0 AND " + WThang("ThangCT", thang, thang) + " AND KCT=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "' AND TyLe=10", v)) + "{TAB}"
    AppIdle w1
    SendKeys CStr(v) + "{TAB}"
    AppIdle w1
    'Dieu chinh
    sql = "SELECT Sum(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND DC=1 AND " + WThang("ThangCT", thang, thang) + " AND KCT=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    vr = SelectSQL(sql, nkv)
    If nkv >= 0 Then
        SendKeys CStr(vr) + "{TAB}" + CStr(nkv) + "{TAB}{TAB}{TAB}{TAB}{TAB}"
        AppIdle w1
    Else
        SendKeys "{TAB}{TAB}" + CStr(-vr) + "{TAB}" + CStr(-nkv) + "{TAB}{TAB}{TAB}"
        AppIdle w1
    End If

    sql = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
        & " WHERE (TK.SoHieu LIKE '" + pVATV + "112*' OR TK.SoHieu LIKE '" + pVATV + "212*') AND RIGHT(TK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "' AND " + WThang("ThangCT", thang, thang) + " AND (HethongTK.SoHieu LIKE '336*' OR HethongTK.SoHieu LIKE '11*')"
    SendKeys CStr(SelectSQL(sql))
    AppIdle w1
KT:
    Set TK = Nothing
End Sub

Private Sub InVATDauRaMV(thang As Integer, tl As Integer, TK As ClsTaikhoan)
    Dim prg As Long, i As Integer, rs As Recordset, sql As String, j As Integer, mvpath As String
    
    mvpath = KiemTraMaVach(1, CurrentDrive + "\Program Files\P2D\P2D.EXE")
    If Len(Dir(mvpath)) = 0 Then Exit Sub
    
    If TK.MaSo = 0 Then
        Select Case tl
            Case -3:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND ThangCT=" + CStr(thang) + " AND KCT=0 AND HoaDon.DC=0 ORDER BY NgayPH"
            Case -2:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND ThangCT=" + CStr(thang) + " AND KCT=1 AND HoaDon.DC=0 ORDER BY NgayPH"
            Case -1:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND ThangCT=" + CStr(thang) + " AND HoaDon.DC=0 ORDER BY NgayPH"
            Case 30:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND ThangCT=" + CStr(thang) + " AND KCT=0 AND HoaDon.DC=0 ORDER BY NgayPH"
            Case Else
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND ThangCT=" + CStr(thang) + " AND KCT=0 AND HoaDon.DC=0 ORDER BY NgayPH"
        End Select
    Else
        Select Case tl
            Case -3:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND ThangCT=" + CStr(thang) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HoaDon.DC=0 ORDER BY NgayPH"
            Case -2:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND ThangCT=" + CStr(thang) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HoaDon.DC=0 ORDER BY NgayPH"
            Case -1:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND ThangCT=" + CStr(thang) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HoaDon.DC=0 ORDER BY NgayPH"
            Case 30:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND ThangCT=" + CStr(thang) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HoaDon.DC=0 ORDER BY NgayPH"
            Case Else
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND ThangCT=" + CStr(thang) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HoaDon.DC=0 ORDER BY NgayPH"
        End Select
    End If
    
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.RecordCount = 0 Then GoTo KetThuc
    
    Recycle pCurDir + "BKVAT2.XLS"
    On Error GoTo KetThuc
    FileCopy Left(mvpath, Len(mvpath) - 7) + "Data\Test02.xls", pCurDir + "BKVAT2.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "BKVAT2.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(1, 2) = thang
    xlsheet.Cells(2, 2) = pNamTC
    
    For i = 5 To 16
        For j = 1 To 12
            xlsheet.Cells(i, j) = ""
        Next
    Next
    
    i = 4
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(i, 1) = rs!MauSo
        xlsheet.Cells(i, 2) = rs!KyHieu
        xlsheet.Cells(i, 3) = rs!sohd
        xlsheet.Cells(i, 4) = Format(rs!NgayPH, "dd/mm/yyyy")
        xlsheet.Cells(i, 5) = rs!Ten
        xlsheet.Cells(i, 6) = rs!mst
        xlsheet.Cells(i, 7) = rs!MatHang
        xlsheet.Cells(i, 8) = rs!ThanhTien
        xlsheet.Cells(i, 9) = rs!TyLe
        xlsheet.Cells(i, 10) = rs!thue
        rs.MoveNext
    Loop
    
    xlsheet.Cells(i + 2, 2) = "END OF FILE"
    
    'xlsheet.Range("D5", "D" + CStr(i + 1)).a = "dd/mm/yyyy"
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    prg = Shell(mvpath, vbNormalFocus)
    If prg = 0 Then Exit Sub
    AppActivate prg
    AppIdle w1
    SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}~"
    AppIdle w1
    SendKeys "{F8}{TAB}~"
    AppIdle w1
    SendKeys pCurDir + "BKVAT2.XLS"
    AppIdle w1
    SendKeys "~{TAB}~~"
KetThuc:
    rs.Close
    Set rs = Nothing
End Sub

Private Sub InVATDauVaoMV(thang As Integer, tl As Integer, HD As Integer, TK As ClsTaikhoan)
    Dim prg As Long, i As Integer, rs As Recordset, sql As String, j As Integer, tk2 As New ClsTaikhoan, fname As String, dgia As Double
    Dim mvpath As String
    
    mvpath = KiemTraMaVach(1, CurrentDrive + "\Program Files\P2D\P2D.EXE")
    If Len(Dir(mvpath)) = 0 Then Exit Sub
    
    tk2.InitTaikhoanSohieu "33312"
    If tk2.tk_id = GTGTKT_ID Then
        sql = "DELETE HoaDon.*  FROM ((HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo " _
            & " WHERE (HethongTK.SoHieu LIKE '" + pVATV + "*') AND (TK.SoHieu LIKE '33312*'  OR InStr(ChungTu.GhiChu,'33312')>0) AND HoaDon.Loai=-1"
        ExecuteSQL5 sql
    End If
    
    If TK.MaSo = 0 Then
        Select Case tl
            Case -3:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND ThangCT=" + CStr(thang) + " AND HDBL=1 AND HoaDon.DC=0 ORDER BY NgayPH,MaCT"
            Case -2:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND ThangCT=" + CStr(thang) + " AND KCT=1 AND HoaDon.DC=0 ORDER BY NgayPH,MaCT"
            Case -1:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND ThangCT=" + CStr(thang) + " AND HoaDon.DC=0 ORDER BY NgayPH,MaCT"
            Case 30:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND (TyLe=5 OR TyLe=10) AND HD=" + CStr(HD) + " AND ThangCT=" + CStr(thang) + " AND KCT=0 AND HoaDon.DC=0 ORDER BY NgayPH,MaCT"
            Case Else
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND TyLe=" + CStr(tl) + " AND HD=" + CStr(HD) + " AND ThangCT=" + CStr(thang) + " AND KCT=0 AND HDBL=0 AND HoaDon.DC=0 ORDER BY NgayPH,MaCT"
        End Select
    Else
        Select Case tl
            Case -3:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND ThangCT=" + CStr(thang) + " AND HDBL=1 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HoaDon.DC=0 ORDER BY NgayPH,MaCT"
            Case -2:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND ThangCT=" + CStr(thang) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HoaDon.DC=0 ORDER BY NgayPH,MaCT"
            Case -1:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND ThangCT=" + CStr(thang) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HoaDon.DC=0 ORDER BY NgayPH,MaCT"
            Case 30:
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND (TyLe=5 OR TyLe=10) AND HD=" + CStr(HD) + " AND ThangCT=" + CStr(thang) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HoaDon.DC=0 ORDER BY NgayPH,MaCT"
            Case Else
                sql = "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND TyLe=" + CStr(tl) + " AND HD=" + CStr(HD) + " AND ThangCT=" + CStr(thang) + " AND KCT=0 AND HDBL=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HoaDon.DC=0 ORDER BY NgayPH,MaCT"
        End Select
    End If
    
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.RecordCount = 0 Then GoTo KetThuc
    
    If HD = 0 Then
        fname = "TEST04.XLS"
        sql = "B¶ng kª 04/GTGT"
    Else
        If (tl > 0 And tl < 5) Or tl = -3 Then
            fname = "TEST05.XLS"
            sql = "B¶ng kª 05/GTGT"
        Else
            fname = "TEST03.XLS"
            sql = "B¶ng kª 03/GTGT"
        End If
    End If
    
    Recycle pCurDir + "BKVATV.XLS"
    On Error GoTo KetThuc
    FileCopy Left(mvpath, Len(mvpath) - 7) + "Data\" + fname, pCurDir + "BKVATV.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "BKVATV.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    
    xlsheet.Cells(1, 2) = thang
    xlsheet.Cells(2, 2) = pNamTC
    
    For i = 5 To 60
        For j = 1 To 12
            xlsheet.Cells(i, j) = ""
        Next
    Next
    
    i = 4
    Do While Not rs.EOF
        i = i + 1
        Select Case fname
            Case "TEST03.XLS":
                xlsheet.Cells(i, 1) = rs!MauSo
                xlsheet.Cells(i, 2) = rs!KyHieu
                xlsheet.Cells(i, 3) = rs!sohd
                xlsheet.Cells(i, 4) = Format(rs!NgayPH, "dd/mm/yyyy")
                xlsheet.Cells(i, 5) = rs!Ten
                xlsheet.Cells(i, 6) = rs!mst
                xlsheet.Cells(i, 7) = rs!MatHang
                xlsheet.Cells(i, 8) = rs!ThanhTien
                xlsheet.Cells(i, 9) = rs!TyLe
                xlsheet.Cells(i, 10) = rs!sops
            Case "TEST04.XLS":
                If rs!SoLuong <> 0 Then dgia = Fix(0.5 + Mask_N * rs!ThangTien / rs!SoLuong) / Mask_N Else dgia = 0
                xlsheet.Cells(i, 1) = Format(rs!NgayPH, "dd/mm/yyyy")
                xlsheet.Cells(i, 2) = rs!Ten
                xlsheet.Cells(i, 3) = rs!mst
                xlsheet.Cells(i, 4) = rs!MatHang
                xlsheet.Cells(i, 5) = rs!SoLuong
                xlsheet.Cells(i, 6) = dgia
                xlsheet.Cells(i, 7) = rs!ThanhTien
            Case "TEST05.XLS":
                xlsheet.Cells(i, 1) = rs!MauSo
                xlsheet.Cells(i, 2) = rs!KyHieu
                xlsheet.Cells(i, 3) = rs!sohd
                xlsheet.Cells(i, 4) = Format(rs!NgayPH, "dd/mm/yyyy")
                xlsheet.Cells(i, 5) = rs!Ten
                xlsheet.Cells(i, 6) = rs!mst
                xlsheet.Cells(i, 7) = rs!ThanhTien
        End Select
        rs.MoveNext
    Loop
    
    xlsheet.Cells(i + 2, 2) = "END OF FILE"
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    prg = Shell(mvpath, vbNormalFocus)
    If prg = 0 Then Exit Sub
    AppActivate prg
    AppIdle w1
    SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}" + sql + "{TAB}~"
    
    AppIdle w1
    SendKeys "{F8}{TAB}~"
    AppIdle w1
    SendKeys pCurDir + "BKVATV.XLS"
    AppIdle w1
    SendKeys "~{TAB}~~"
KetThuc:
    rs.Close
    Set rs = Nothing
    Set tk2 = Nothing
End Sub

Private Sub MedNgay_GotFocus(Index As Integer)
'    AutoSelect MedNgay(Index)
End Sub

Private Sub MedNgay_LostFocus(Index As Integer)
    If IsDate(MedNgay(Index).Text) Then
        ngay(Index) = CDate(MedNgay(Index).Text)
    Else
        RFocus MedNgay(Index)
    End If
    
     If Index = 0 Then
       MedNgay(1).Text = MedNgay(0).Text
    End If
End Sub

Private Sub InCTKQKD(tdau As Integer, tcuoi As Integer)
    Dim sql As String, kn As Double
    Dim rs_lailo  As Recordset
    
    GauGe.Max = 4
    ExecuteSQL5 "UPDATE KQKDCT SET KyNay = 0"
    Set rs_lailo = DBKetoan.OpenRecordset("SELECT DISTINCTROW * FROM KQKDCT WHERE TKNo<>'0' AND TKCo<>'0' ORDER BY MaSo", dbOpenSnapshot)
    Do While Not rs_lailo.EOF
        sql = "SELECT DISTINCTROW Sum(ChungTu.SoPS) AS F1 " _
            & "FROM (HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) INNER JOIN HeThongTK AS HeThongTK_1 ON ChungTu.MaTKCo = HeThongTK_1.MaSo " _
            & "WHERE HethongTK.SoHieu LIKE '" + rs_lailo!tkno + "*' AND HethongTK_1.SoHieu LIKE '" + rs_lailo!TkCo + "*' AND " + WThang("ThangCT", tdau, tcuoi)
        kn = SelectSQL(sql)
        sql = "SELECT DISTINCTROW Sum(ChungTu.SoPS) AS F1 " _
            & "FROM (HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) INNER JOIN HeThongTK AS HeThongTK_1 ON ChungTu.MaTKCo = HeThongTK_1.MaSo " _
            & "WHERE HethongTK_1.SoHieu LIKE '" + rs_lailo!tkno + "*' AND HethongTK.SoHieu LIKE '" + rs_lailo!TkCo + "*' AND " + WThang("ThangCT", tdau, tcuoi)
        kn = kn - SelectSQL(sql)
        If rs_lailo!dau > 0 Then
            ExecuteSQL5 "UPDATE KQKDCT SET KyNay = " + DoiDau(kn) + " WHERE Ma=" + CStr(rs_lailo!ma)
        Else
            ExecuteSQL5 "UPDATE KQKDCT SET KyNay = " + DoiDau(-kn) + " WHERE Ma=" + CStr(rs_lailo!ma)
        End If
        rs_lailo.MoveNext
    Loop
    GauGe.Value = 1
    Set rs_lailo = DBKetoan.OpenRecordset("SELECT DISTINCTROW * FROM KQKDCT WHERE TKNo<>'0' AND TKCo='0' ORDER BY MaSo", dbOpenSnapshot)
    Do While Not rs_lailo.EOF
        sql = "SELECT DISTINCTROW Sum(ChungTu.SoPS) AS F1 " _
            & "FROM HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo " _
            & "WHERE HethongTK.SoHieu LIKE '" + rs_lailo!tkno + "*' AND " + WThang("ThangCT", tdau, tcuoi)
        kn = SelectSQL(sql)

        If rs_lailo!dau > 0 Then
            ExecuteSQL5 "UPDATE KQKDCT SET KyNay = " + DoiDau(kn) + " WHERE Ma=" + CStr(rs_lailo!ma)
        Else
            ExecuteSQL5 "UPDATE KQKDCT SET KyNay = " + DoiDau(-kn) + " WHERE Ma=" + CStr(rs_lailo!ma)
        End If
        rs_lailo.MoveNext
    Loop
    
    Set rs_lailo = DBKetoan.OpenRecordset("SELECT DISTINCTROW * FROM KQKDCT WHERE TKNo='0' AND TKCo<>'0' ORDER BY MaSo", dbOpenSnapshot)
    Do While Not rs_lailo.EOF
        sql = "SELECT DISTINCTROW Sum(ChungTu.SoPS) AS F1 " _
            & "FROM HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKCo " _
            & "WHERE HethongTK.SoHieu LIKE '" + rs_lailo!TkCo + "*' AND " + WThang("ThangCT", tdau, tcuoi)
        kn = SelectSQL(sql)

        If rs_lailo!dau > 0 Then
            ExecuteSQL5 "UPDATE KQKDCT SET KyNay = " + DoiDau(kn) + " WHERE Ma=" + CStr(rs_lailo!ma)
        Else
            ExecuteSQL5 "UPDATE KQKDCT SET KyNay = " + DoiDau(-kn) + " WHERE Ma=" + CStr(rs_lailo!ma)
        End If
        rs_lailo.MoveNext
    Loop
    
    GauGe.Value = 2
    
    sql = "SELECT SUM(DuCo_0-DuNo_0) AS F1 FROM HethongTK WHERE SoHieu='4211'"                      ' OR SoHieu='112'
    kn = SelectSQL(sql)
    If kn < 0 Then ExecuteSQL5 "UPDATE KQKDCT SET KyNay= " + DoiDau(kn) + " WHERE MaSo=36"
    GauGe.Value = 1
    
    InKqkd tdau, tcuoi, 0, nn
    kn = SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=11 OR MaSo=24 OR MaSo=25 OR MaSo=22 OR MaSo=32")
    kn = kn - SelectSQL("SELECT Sum(KyNay) AS F1 FROM KqkdCT WHERE MaSo>=12 AND MaSo<20")
    ExecuteSQL5 "UPDATE KQKDCT SET KyNay= " + DoiDau(kn) + " WHERE MaSo=20"
    
    ' Cong don
    Set rs_lailo = DBKetoan.OpenRecordset("SELECT KyNay,MaSoCha FROM KQKDCT WHERE TongHop=0 And MaSoCha<>0 ORDER BY MaSo", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_lailo.EOF
         If rs_lailo!MaSoCha > 0 Then
                 ExecuteSQL5 "UPDATE KQKDCT SET KyNay=KyNay+" + DoiDau(rs_lailo!KyNay) + " WHERE MaSo=" + CStr(rs_lailo!MaSoCha)
        Else
                 ExecuteSQL5 "UPDATE KQKDCT SET KyNay=KyNay-" + DoiDau(rs_lailo!KyNay) + " WHERE MaSo=" + CStr(-rs_lailo!MaSoCha)
        End If
        rs_lailo.MoveNext
    Loop
    
    Set rs_lailo = DBKetoan.OpenRecordset("SELECT KyNay,MaSoCha FROM KQKDCT WHERE MaSo=16", dbOpenSnapshot, dbForwardOnly)
    ExecuteSQL5 "UPDATE KQKDCT SET KyNay=KyNay+" + DoiDau(rs_lailo!KyNay) + " WHERE MaSo=12"
    
    Set rs_lailo = DBKetoan.OpenRecordset("SELECT KyNay,MaSoCha FROM KQKDCT WHERE TongHop>0 And MaSoCha<>0 ORDER BY MaSo", dbOpenDynaset, dbForwardOnly)
    Do While Not rs_lailo.EOF
        If rs_lailo!MaSoCha > 0 Then
            ExecuteSQL5 "UPDATE KQKDCT SET KyNay=KyNay+" + DoiDau(rs_lailo!KyNay) + " WHERE MaSo=" + CStr(rs_lailo!MaSoCha)
        Else
            ExecuteSQL5 "UPDATE KQKDCT SET KyNay=KyNay-" + DoiDau(rs_lailo!KyNay) + " WHERE MaSo=" + CStr(-rs_lailo!MaSoCha)
        End If
        rs_lailo.MoveNext
    Loop
    rs_lailo.Close
    Set rs_lailo = Nothing
    GauGe.Value = 4
    
    frmMain.Rpt.ReportFileName = "KQKDCT.RPT"
    RptSetDate NgayCuoiThang(pNamTC, 12), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(pThangDauKy, ThangCuoiNamTC, nn) + "'"
End Sub

Private Sub GTGTCT()
    Dim i As Integer, kq As Double
    
    For i = 1 To 12
        frmMain.Rpt.Formulas(3 + i) = "R" + CStr(i) + "=" + DoiDau(SoPSTK("3331", i, i, 1))
        kq = SoPSTK(pVATV, i, i, -1)
        frmMain.Rpt.Formulas(15 + i) = "V" + CStr(i) + "=" + DoiDau(kq)
        frmMain.Rpt.Formulas(27 + i) = "K" + CStr(i) + "=" + DoiDau(kq - PSDu("6", pVATV, i, i))
        frmMain.Rpt.Formulas(39 + i) = "H" + CStr(i) + "=" + DoiDau(PSDu("11", pVATV + "112", i, i) + PSDu("11", pVATV + "212", i, i))
        frmMain.Rpt.Formulas(51 + i) = "N" + CStr(i) + "=" + DoiDau(PSDu("33311", "11", i, i))
    Next
    
    frmMain.Rpt.ReportFileName = "THUE2.RPT"
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(pThangDauKy, ThangCuoiNamTC, nn) + "'"
    frmMain.Rpt.Formulas(70) = "SoDN=" + DoiDau(SoDuTKSH("3331", 0) - SoDuTKSH(pVATV, 0))
End Sub

Private Sub ToKhaiVAT(tdau As Integer, tcuoi As Integer, taikhoan As ClsTaikhoan)
    Dim sql As String, vr As Double, VV As Double, vvx As Double
    Dim KT As Double, shct As String, Fx As Integer
    Dim dkn As Double, dkc As Double, dknt As Double, TK As New ClsTaikhoan
    
    Fx = IIf(Left(frmMain.LbCty(8).Caption, 2) = "35" Or Left(frmMain.LbCty(3).Caption, 2) = "64", 35, 0)
    
    TK.InitTaikhoanSohieu "33312"
    If TK.tk_id = GTGTKT_ID Then
        sql = "DELETE HoaDon.*  FROM ((HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo " _
            & " WHERE (HethongTK.SoHieu LIKE '" + pVATV + "*') AND (TK.SoHieu LIKE '33312*'  OR InStr(ChungTu.GhiChu,'33312')>0) AND HoaDon.Loai=-1"
        ExecuteSQL5 sql
    End If
    
    GauGe.Max = 6
    frmMain.Rpt.Formulas(3) = "Thang = " + CStr(tdau)
    frmMain.Rpt.Formulas(4) = "ThangCuoi = " + CStr(tcuoi)
    If taikhoan.MaSo = 0 Then
        frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    Else
        frmMain.Rpt.Formulas(1) = "TenCN = '" + taikhoan.Ten + "'"
        frmMain.Rpt.Formulas(5) = "DiaChi = '" + taikhoan.GhiChu + "'"
    End If
    frmMain.Rpt.Formulas(6) = "MSThue = '" + frmMain.LbCty(8).Caption + "'"
    frmMain.Rpt.Formulas(7) = "SoHieuTK = '" + vatr + "'"
    
    frmMain.Rpt.Formulas(30) = "Quan = '" + frmMain.LbCty(10).Caption + "'"
    frmMain.Rpt.Formulas(31) = "TP = '" + frmMain.LbCty(11).Caption + "'"
    frmMain.Rpt.Formulas(32) = "Tel = '" + frmMain.LbCty(3).Caption + "'"
    frmMain.Rpt.Formulas(33) = "Fax = '" + frmMain.LbCty(4).Caption + "'"
    frmMain.Rpt.Formulas(34) = "Email = '" + frmMain.LbCty(9).Caption + "'"
    
    shct = taikhoan.SHChiTiet()
    vr = SelectSQL("SELECT SUM(ThanhTien) AS F1 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND DC=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'")
    frmMain.Rpt.Formulas(8) = "DTKCT=" + DoiDau(vr)
    
    GauGe.Value = 1
    frmMain.Rpt.Formulas(23) = "Vat0DT=" + DoiDau(SelectSQL("SELECT Sum(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND DC=0 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "' AND TyLe=0", VV))
    frmMain.Rpt.Formulas(24) = "Vat0=" + DoiDau(VV)
    frmMain.Rpt.Formulas(11) = "Vat5DT=" + DoiDau(SelectSQL("SELECT Sum(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND DC=0 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "' AND TyLe=5", VV))
    frmMain.Rpt.Formulas(12) = "Vat5=" + DoiDau(VV)
    frmMain.Rpt.Formulas(13) = "Vat10DT=" + DoiDau(SelectSQL("SELECT Sum(ThanhTien) AS F1,SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND DC=0 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "' AND TyLe=10", VV))
    frmMain.Rpt.Formulas(14) = "Vat10=" + DoiDau(VV)
    'Dieu chinh
    sql = "SELECT Sum(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND DC=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    vr = SelectSQL(sql, VV)
    If VV >= 0 Then
        frmMain.Rpt.Formulas(37) = "RaT=" + DoiDau(vr)
        frmMain.Rpt.Formulas(38) = "RaTV=" + DoiDau(VV)
    Else
        frmMain.Rpt.Formulas(37) = "RaG=" + DoiDau(Abs(vr))
        frmMain.Rpt.Formulas(38) = "RaGV=" + DoiDau(Abs(VV))
    End If
    
    GauGe.Value = 2
    KT = 0 'SoNKChuaThue(tdau, tcuoi, shct)
    If Len(shct) > 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND HD=1 AND DC=0 AND KCT=0 AND " + WThang("ThangCT", tdau, tcuoi) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    Else
        sql = "SELECT SUM(ThanhTien) AS F1 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND HD=1 AND DC=0 AND KCT=0 AND " + WThang("ThangCT", tdau, tcuoi)
    End If
    vvx = SelectSQL(sql) + KT
    frmMain.Rpt.Formulas(43) = "TongVaoV=" + DoiDau(vvx)
    If Len(shct) > 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND DC=0 AND " + WThang("ThangCT", tdau, tcuoi) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    Else
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND DC=0 AND " + WThang("ThangCT", tdau, tcuoi)
    End If
    vvx = SelectSQL(sql, VV) + KT
    frmMain.Rpt.Formulas(19) = "TongVao=" + DoiDau(vvx)
    
    sql = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
            & " WHERE (HethongTK.SoHieu LIKE '142*' OR HethongTK.SoHieu LIKE '242*' OR HethongTK.SoHieu LIKE '6*') AND (TK.SoHieu LIKE '" + pVATV + "*') AND RIGHT(TK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "' AND " + WThang("ThangCT", tdau, tcuoi)
    frmMain.Rpt.Formulas(17) = "TongVATV=" + DoiDau(VV - SelectSQL(sql))
    
    If TK.tk_id = GTGTKT_ID Then
        SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKNC(-1) + " WHERE (HethongTK.SoHieu LIKE '521*' OR HethongTK.SoHieu LIKE '531*') AND " + WThang("ThangCT", tdau, tcuoi)
        sql = "SELECT SUM(SoPS) AS F1 FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT " _
            & " WHERE IsNull(MienTru.MaCT) AND (HethongTK.SoHieu LIKE '33312*') AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "' AND " + WThang("ThangCT", tdau, tcuoi) + " AND (TK.SoHieu LIKE '11*' OR TK.SoHieu LIKE '136*')"
        frmMain.Rpt.Formulas(18) = "TongVATx=" + DoiDau(VV - SelectSQL(sql) + PSDu(pVATV, "33312", tdau, tcuoi, shct))               '- SelectSQL(sql)
    Else
        frmMain.Rpt.Formulas(18) = "TongVATx=" + DoiDau(VV)
    End If
    GauGe.Value = 3
    'Dieu chinh
    If Len(shct) > 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND DC=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    Else
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND DC=1 AND " + WThang("ThangCT", tdau, tcuoi)
    End If
    vvx = SelectSQL(sql, VV)
    If VV >= 0 Then
        frmMain.Rpt.Formulas(35) = "VaoT=" + DoiDau(vvx)
        frmMain.Rpt.Formulas(36) = "VaoTV=" + DoiDau(VV)
    Else
        frmMain.Rpt.Formulas(35) = "VaoG=" + DoiDau(Abs(vvx))
        frmMain.Rpt.Formulas(36) = "VaoGV=" + DoiDau(Abs(VV))
    End If
    ' Nhap khau
    If Len(shct) > 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND NK=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    Else
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND NK=1 AND " + WThang("ThangCT", tdau, tcuoi)
    End If
    vvx = SelectSQL(sql, VV)
    frmMain.Rpt.Formulas(39) = "NK=" + DoiDau(vvx)
    frmMain.Rpt.Formulas(40) = "NKV=" + DoiDau(VV)
    ' Tai san
    If Len(shct) > 0 Then
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND TS=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
    Else
        sql = "SELECT SUM(ThanhTien) AS F1,SUM(SoPS) AS F2 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND TS=1 AND " + WThang("ThangCT", tdau, tcuoi)
    End If
    vvx = SelectSQL(sql, VV)
    frmMain.Rpt.Formulas(41) = "TS=" + DoiDau(vvx)
    frmMain.Rpt.Formulas(42) = "TSV=" + DoiDau(VV)
    
    KT = DKToKhai(tdau, shct)
    If KT > 0 Then frmMain.Rpt.Formulas(20) = "KyTruoc=" + DoiDau(KT)
    
    sql = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
        & " WHERE (TK.SoHieu LIKE '" + pVATV + "112*' OR TK.SoHieu LIKE '" + pVATV + "212*') AND RIGHT(TK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "' AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HethongTK.SoHieu LIKE '336*' OR HethongTK.SoHieu LIKE '11*' OR HethongTK.SoHieu LIKE '138*')"
    VV = SelectSQL(sql)
    frmMain.Rpt.Formulas(44) = "HoanTra=" + DoiDau(VV)
    
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    GauGe.Value = 4
    
    frmMain.Rpt.ReportFileName = "TOKHAI.RPT"
    Set TK = Nothing
End Sub

Private Function KiemTraMaVach(loai As Integer, dd As String) As String
    Dim st As String
    
    st = GetSetting(IniPath, "VAT", GetFileName(pDataPath) + "_MV" + CStr(loai))
    
    On Error Resume Next
    If Len(st) = 0 Or Len(Dir(st)) = 0 Then
        st = FrmGetStr.GetString("§­êng dÉn ch­¬ng tr×nh m· v¹ch:", App.ProductName, dd)
        If Len(Dir(st)) > 0 Then
            SaveSetting IniPath, "VAT", GetFileName(pDataPath) + "_MV" + CStr(loai), st
            KiemTraMaVach = st
        End If
    Else
        KiemTraMaVach = st
    End If
    On Error GoTo 0
End Function

Private Function DKToKhai(thang As Integer, shct As String)
    Dim TK As ClsTaikhoan, duno As Double, duco As Double, dunt As Double, sql As String, X As Double
    Dim tdau As Integer
    
    If thang = pThangDauKy Then
        Set TK = New ClsTaikhoan
        
        TK.InitTaikhoanSohieu pVATV, shct
        TK.SoDuTK 0, duno, duco, dunt
        duno = duno - duco
        
        TK.InitTaikhoanSohieu vatr, shct
        TK.SoDuTK 0, X, duco, dunt
        duno = duno + X - duco
        
        Set TK = Nothing
    Else
        tdau = ThangTruoc(thang)
        duno = DKToKhai(tdau, shct)
        If Len(shct) > 0 Then
            sql = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKHD(-1) + " WHERE HoaDon.Loai=-1 AND " + WThang("ThangCT", tdau, tdau) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
        Else
            sql = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=-1 AND " + WThang("ThangCT", tdau, tdau)
        End If
        duno = duno + SelectSQL(sql)
        
        ' Kiem tra dau ra co bi am thi bo sung
        If Len(shct) > 0 Then
            sql = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tdau) + " AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
        Else
            sql = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKHD(0) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tdau)
        End If
        X = SelectSQL(sql)
        If X < 0 Then duno = duno + Abs(X)
        
        sql = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
                & " WHERE (HethongTK.SoHieu LIKE '142*' OR HethongTK.SoHieu LIKE '242*' OR HethongTK.SoHieu LIKE '6*') AND (TK.SoHieu LIKE '" + pVATV + "*') AND RIGHT(TK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "' AND " + WThang("ThangCT", tdau, tdau)
        duno = duno - SelectSQL(sql)
        
        sql = "SELECT SUM(IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS)) AS F1 FROM " + ChungTu2TKHD(1) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tdau) + " AND KCT=0 AND RIGHT(HethongTK.SoHieu," + CStr(Len(shct)) + ") = '" + shct + "'"
        duno = duno - SelectSQL(sql)
        
        sql = "SELECT SUM(SoPS) AS F1 FROM " + ChungTu2TKNC(0) _
            & " WHERE (TK.SoHieu LIKE '" + pVATV + "112*' OR TK.SoHieu LIKE '" + pVATV + "212*') AND RIGHT(TK.SoHieu," + CStr(Len(shct)) + ")='" + shct + "' AND " + WThang("ThangCT", tdau, tdau) + " AND (HethongTK.SoHieu LIKE '336*' OR HethongTK.SoHieu LIKE '11*' OR HethongTK.SoHieu LIKE '138*')"
        duno = duno - SelectSQL(sql)
    End If
    If duno < 0 Then duno = 0
    DKToKhai = duno
End Function

Private Sub BangKeTHPS(x2 As Worksheet, shtk As String, tdau As Integer, tcuoi As Integer, cap As Integer, loaips As Integer)
    Dim rs As Recordset, TK As New ClsTaikhoan, i As Integer, j As Integer, k As Integer, L As Integer, sql As String
    
    TK.InitTaikhoanSohieu shtk
    If TK.MaSo = 0 Then GoTo KT
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = "B¶ng tæng hîp ph¸t sinh tµi kho¶n " + shtk
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi, nn)
    
    ' Liet ke tai khoan chi tiet
    If TK.tkcon = 0 Then
        x2.Cells(8, 1) = TK.sohieu
        i = 1
    Else
        i = 0
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW SoHieu FROM HethongTK WHERE Cap=" + CStr(TK.cap + 1) + " AND SoHieu LIKE '" + shtk + "*' ", dbOpenSnapshot)
        Do While Not rs.EOF
            i = i + 1
            x2.Cells(7 + i, 1) = rs!sohieu
            rs.MoveNext
        Loop
        rs.Close
    End If
    x2.Cells(8 + i, 1) = "Tæng céng"
    
    ' Liet ke doi ung
    j = 0
    ' Lay danh sach tai khoan tai chinh
    sql = "SELECT BaoCaoCP.SoHieu AS SH FROM (" + ChungTu2TKNC(0) + ") LEFT JOIN BaoCaoCP ON " + IIf(loaips > 0, "Hethong", "") + "TK.SoHieu LIKE BaoCaoCP.SoHieu+'*'" _
        & " WHERE SoPS<>0 AND LEFT(" + IIf(loaips < 0, "Hethong", "") + "TK.SoHieu," + CStr(Len(TK.sohieu)) + ")='" + TK.sohieu + "' AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        j = j + 1
        x2.Cells(6, 2 + j) = rs!sh
        rs.MoveNext
    Loop
    rs.Close
    x2.Cells(6, 3 + j) = "Tæng céng"
    
    For k = 1 To i
        For L = 1 To j
            If loaips < 0 Then
                x2.Cells(7 + k, 2 + L) = PSDu(x2.Cells(7 + k, 1), x2.Cells(6, 2 + L), tdau, tcuoi, , 1)
            Else
                x2.Cells(7 + k, 2 + L) = PSDu(x2.Cells(6, 2 + L), x2.Cells(7 + k, 1), tdau, tcuoi, , 1)
            End If
        Next
        If j > 0 Then x2.Range(XLSCol(3 + j) + CStr(7 + k)).Formula = "=SUM(C" + CStr(7 + k) + ":" + XLSCol(2 + j) + CStr(7 + k) + ")"
    Next
    
    For k = 1 To j + 1
        If j > 0 Then x2.Range(XLSCol(2 + k) + CStr(8 + i)).Formula = "=SUM(" + XLSCol(2 + k) + "8:" + XLSCol(2 + k) + CStr(7 + i) + ")"
    Next
    
    xlsheet.Range("A3", XLSCol(3 + j) + "3").MergeCells = True
    xlsheet.Range("A4", XLSCol(3 + j) + "4").MergeCells = True
    
    xlsheet.Range("A6", XLSCol(3 + j) + "7").Font.Bold = True
    xlsheet.Range("A" + CStr(8 + i), XLSCol(3 + j) + CStr(8 + i)).Font.Bold = True
    xlsheet.Range(XLSCol(3 + j) + "6", XLSCol(3 + j) + CStr(8 + i)).Font.Bold = True
    With xlsheet.Range("A6", XLSCol(3 + j) + CStr(8 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
KT:
    Set rs = Nothing
    Set TK = Nothing
End Sub

Private Sub PTDTCP(tdau As Integer, tcuoi As Integer)
    Dim r As Integer, X As Integer, j As Integer, ps As Double, f As Integer, v As Double, pslk As Double, vlk As Double
    Dim TK As New ClsTaikhoan

    Recycle pCurDir + "DTCP.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\DTCP.XLS", pCurDir + "DTCP.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "DTCP.XLS"
    On Error GoTo 0
    
    For X = 1 To 3
        Set xlsheet = xlapp.Worksheets(X)
        
        xlsheet.Cells(1, 1) = pTenCty
        xlsheet.Cells(2, 1) = pTenCn
        xlsheet.Cells(5, 1) = ABCtoVNI(ThoiGian(tdau, tcuoi, nn))
        
        For r = 8 To 50
            ps = 0
            pslk = 0
            f = 0
            j = 12
            
            If X > 1 Then
                Do While Len(xlsheet.Cells(r, j)) > 0
                    TK.InitTaikhoanSohieu xlsheet.Cells(r, j)
                    If TK.MaSo > 0 Then
                        ps = ps + SoPSTK2(TK, tdau, tcuoi)
                        pslk = pslk + SoPSTK2(TK, pThangDauKy, tcuoi)
                        f = f + 1
                    End If
                    j = j + 1
                Loop
                If f > 0 Then xlsheet.Cells(r, 3) = ps
                If f > 0 Then xlsheet.Cells(r, 4) = pslk
            Else
                If Len(xlsheet.Cells(r, 12)) > 0 Then
                    ps = DoanhThuTK(xlsheet.Cells(r, 12), tdau, tcuoi, v)
                    pslk = DoanhThuTK(xlsheet.Cells(r, 12), pThangDauKy, tcuoi, vlk)
                    xlsheet.Cells(r, 3) = ps + v
                    xlsheet.Cells(r, 4) = ps
                    xlsheet.Cells(r, 6) = pslk + vlk
                    xlsheet.Cells(r, 7) = pslk
                    If Len(xlsheet.Cells(r, 13)) > 0 Then
                        xlsheet.Cells(r, 5) = SoPSTK(TK.sohieu, tdau, tcuoi, -1)
                        xlsheet.Cells(r, 8) = SoPSTK(TK.sohieu, pThangDauKy, tcuoi, -1)
                    End If
                End If
            End If
        Next
    Next
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    
    CallExcel "DTCP.XLS"
KetThuc:
    Set TK = Nothing
End Sub

Private Function BangTHCTuGoc(mgs As Long, shtk As String, ndau As Date, ncuoi As Date, loaips As Integer) As Boolean
    Dim rs As Recordset, sql As String, i As Integer
    
    BangTHCTuGoc = False
    If loaips < 0 Then
        sql = "SELECT DISTINCTROW LEFT(TK.SoHieu,3) AS TKDU FROM " + ChungTu2TKNC(10) _
            & " WHERE CTGS=" + CStr(mgs) + " AND (HethongTK.SoHieu LIKE '" + shtk + "*') AND " + WNgay("NgayGS", ndau, ncuoi) + " AND SoPS<>0  GROUP BY LEFT(TK.SoHieu,3) HAVING Sum(ChungTu.SoPS)<>0"
    Else
        sql = "SELECT DISTINCTROW LEFT(HethongTK.SoHieu,3) AS TKDU FROM " + ChungTu2TKNC(10) _
            & " WHERE CTGS=" + CStr(mgs) + " AND (TK.SoHieu LIKE '" + shtk + "*') AND " + WNgay("NgayGS", ndau, ncuoi) + " AND SoPS<>0  GROUP BY LEFT(HethongTK.SoHieu,3) HAVING Sum(ChungTu.SoPS)<>0"
    End If
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.RecordCount = 0 Then GoTo KT
    
    Do While Not rs.EOF And i < 10
        i = i + 1
        frmMain.Rpt.Formulas(2 + i) = "TK" + CStr(i) + "='" + rs!tkdu + "'"
        rs.MoveNext
    Loop
    frmMain.Rpt.Formulas(3 + i) = "SoHieuTK='" + shtk + "'"
    frmMain.Rpt.Formulas(4 + i) = "ThoiGian='" + ThoiGianN(ndau, ncuoi, nn) + "'"
    frmMain.Rpt.Formulas(5 + i) = "TKNo='" + IIf(loaips < 0, "Ghi nî tµi kho¶n " + shtk + ", ghi cã c¸c tµi kho¶n", "Ghi cã tµi kho¶n " + shtk + ", ghi nî c¸c tµi kho¶n") + "'"
    frmMain.Rpt.Formulas(6 + i) = "DG1='Sè: " + CStr((Day(ndau) - 1) \ 7 + 1) + "'"
    frmMain.Rpt.Formulas(7 + i) = "DG2='" + SelectSQL("SELECT SoHieu+' - '+DienGiai AS F1 FROM CTGhiSo WHERE MaSo=" + CStr(mgs)) + "'"
        
    If loaips < 0 Then
        sql = "SELECT DISTINCTROW ChungTu.SoHieu, NgayCT, NgayGS, DienGiai" + IIf(nn > 0, "E", "") + " AS DG, SoPS, LEFT(TK.SoHieu,3) AS TKDU FROM " + ChungTu2TKNC(10) _
            & " WHERE CTGS=" + CStr(mgs) + " AND HethongTK.SoHieu LIKE '" + shtk + "*' AND " + WNgay("NgayGS", ndau, ncuoi) + " AND SoPS<>0 ORDER BY NgayGS, ChungTu.SoHieu"
    Else
        sql = "SELECT DISTINCTROW ChungTu.SoHieu, NgayCT, NgayGS, DienGiai" + IIf(nn > 0, "E", "") + " AS DG, SoPS, LEFT(HethongTK.SoHieu,3) AS TKDU FROM " + ChungTu2TKNC(10) _
            & " WHERE CTGS=" + CStr(mgs) + " AND TK.SoHieu LIKE '" + shtk + "*' AND " + WNgay("NgayGS", ndau, ncuoi) + " AND SoPS<>0 ORDER BY NgayGS, ChungTu.SoHieu"
    End If
    SetSQL "QSoCai", sql
        
    frmMain.Rpt.ReportFileName = "THCTGOC.RPT"
    RptSetDate ncuoi, nn
    BangTHCTuGoc = True
KT:
    rs.Close
    Set rs = Nothing
End Function

Private Function CTGhiSo2(mgs As Long, tdau As Integer, tcuoi As Integer) As Boolean
    Dim sql As String
    
    sql = "SELECT  IIF(ThangCT<10,'0','')+Cstr(100*ThangCT)+Cstr((Day(NgayGS)-1)\7+1) AS TT, Last(ChungTu.SoHieu) AS SH, Last(ChungTu.NgayCT) AS NCT, Last(ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ") AS DG, LEFT(HeThongTK.SoHieu,3) AS TKNo, LEFT(HeThongTK_1.SoHieu,3) AS TKCo, Sum(ChungTu.SoPS) AS PS, Last(NgayGS) AS NGS " _
        & " FROM (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) LEFT JOIN HeThongTK AS HeThongTK_1 ON ChungTu.MaTKTCCo = HeThongTK_1.MaSo WHERE SoPS<>0 AND " + WThang("ThangCT", tdau, tcuoi) + " AND CTGS=" + CStr(mgs) + " AND (IsNull(HethongTK.MaSo) OR HethongTK.Loai>0) AND (IsNull(HethongTK_1.MaSo) OR HethongTK_1.Loai>0) " _
        & " GROUP BY IIF(ThangCT<10,'0','')+Cstr(100*ThangCT)+Cstr((Day(NgayGS)-1)\7+1), LEFT(HeThongTK.SoHieu,3),LEFT(HeThongTK_1.SoHieu,3) ORDER BY IIF(ThangCT<10,'0','')+Cstr(100*ThangCT)+Cstr((Day(NgayGS)-1)\7+1), Last(NgayGS)"
        
    SetSQL "QNhatKy", sql
    frmMain.Rpt.ReportFileName = "CTGHISO2.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    
    frmMain.Rpt.Formulas(4) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(6) = "DG='" + SelectSQL("SELECT SoHieu+' - '+DienGiai AS F1 FROM CTGhiSo WHERE MaSo=" + CStr(mgs)) + "'"
End Function

Private Sub InBKTheoTK(loai As Integer, tdau As Integer, tcuoi As Integer, shtk As String, shdu As String)
    Dim sql As String
    
    GauGe.Max = 1
    
    SetSQL "MienTru", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT,HethongTK.SoHieu AS TKDU FROM (" + ChungTu2TKHD(10) + ") LEFT JOIN HethongTK ON ChungTu.MaTK" + IIf(loai < 0, "No", "Co") + "=HethongTK.MaSo WHERE HoaDon.Loai=" + CStr(loai) + " AND " + WThang("ThangCT", tdau, tcuoi) + " ORDER BY NgayPH,MaCT"
    'SetSQL "MienTru", "SELECT DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,chungtu.mahs as DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT,HethongTK.SoHieu AS TKDU FROM (" + ChungTu2TKHD(10) + ") LEFT JOIN HethongTK ON ChungTu.MaTK" + IIf(loai < 0, "No", "Co") + "=HethongTK.MaSo WHERE HoaDon.Loai=" + CStr(loai) + " AND " + WThang("ThangCT", tdau, tcuoi) + " ORDER BY NgayPH,MaCT"
    
    SetSQL "QChitiet", "SELECT DISTINCTROW MaCT FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTK" + IIf(loai < 0, "No", "Co") + "=HethongTK.MaSo WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.SoHieu LIKE '" + shtk + "*' GROUP BY MaCT"
    SetSQL "QCdt", "SELECT DISTINCTROW MaCT FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTK" + IIf(loai < 0, "Co", "No") + "=HethongTK.MaSo WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.SoHieu LIKE '" + shdu + "*' GROUP BY MaCT"
    SetSQL "QNhatKy", "SELECT MienTru.* FROM (MienTru INNER JOIN QChitiet ON MienTru.MaCT=QChitiet.MaCT) INNER JOIN QCdt ON MienTru.MaCT=QCdt.MaCT ORDER BY NgayPH,MienTru.MaCT"
    
    GauGe.Value = 1
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\BANGKE" + IIf(loai < 0, "V", "R") + "7.RPT"
    
    RptSetDate NgayCuoiThang(pNamTC, tcuoi)
    frmMain.Rpt.Formulas(3) = "thang = " + CStr(tdau)
    frmMain.Rpt.Formulas(4) = "ThangCuoi = " + CStr(tcuoi)
    frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    frmMain.Rpt.Formulas(6) = "MSThue = '" + frmMain.LbCty(8).Caption + "'"
    frmMain.Rpt.Formulas(8) = "TK = '" + IIf(Len(shtk) > 0, "Tµi kho¶n: " + shtk, "") + IIf(Len(shdu) > 0, " - §èi øng: " + shdu, "") + "'"
End Sub

Private Sub InCdts_HCSN(tdau As Integer, tcuoi As Integer, cap As Integer)
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
        
    SetSQL "QTongHopCT", "SELECT DISTINCTROW SoHieu, Loai, First(Cap) AS FirstOfCap, First(Ten" + IIf(nn > 0, "E", "") + ") AS FirstOfTen, First(Kieu) AS FirstOfKieu, First(Loai) AS FirstOfLoai, First(MaTC = MaSo OR MaTC = 0) AS FirstOfChuanLa, First(TKCon) AS FirstOfTKCon, First(CapDuoi) AS CD," _
        & " Sum(DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkNo, Sum(DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkCo, Sum(" + sqln + ") AS PsNo, Sum(" + sqlc + ") AS PsCo, Sum(" + sqlnx + ") AS PsNoX, Sum(" + sqlcx + ") AS PsCoX, Sum(HeThongTK.DuNo_" + st + ") AS CkNo, Sum(HeThongTK.DuCo_" + st + ") AS CkCo" _
        & " From HeThongTK WHERE  Cap<=" + CStr(cap) + " GROUP BY SoHieu, Loai Having First(Cap > 0) ORDER BY IIF(Loai>0,Loai,Loai+10),SoHieu"

    GauGe.Value = 2
    XDCapTK ""
    frmMain.Rpt.ReportFileName = "CDTK_HC.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    
    GauGe.Value = 3
End Sub
'
'
'' in bang ke dau vao exl
'
'Private Function InVATDauvao_exel(tdau As Integer, tcuoi As Integer, tl As Integer, HD As Integer, TK As ClsTaikhoan) As Boolean
'    Dim sql, pDataPath As String
'  Dim rs
'  Dim UVowels As String
'  Dim mDOMVowels As Class1
' 'Set MyUnicodeText = New Class1
'  Set mDOMVowels = New Class1
'
'        ' Read Unicode Text from file txtFileName and display in TextBox1(0)
'       'TextBox1(0).Text = MyUnicodeText.ReadUnicode(txtFileName)
'          'UVowels = mDOMVowels.ReadUnicode(pCurDir & "UnicodeVowels.xml")
'          UVowels = mDOMVowels.ReadUnicode(GetLocalDirectory & "UnicodeVowels.xml")
'
'           PopulateListBoxFromFile lstVNI, "VNIVowelMap.txt", False
'           lstVNI.Font.Name = "Courier"
'    GauGe.Max = 1
'   ' Recycle pCurDir + "bangke.XLS"
'    On Error GoTo KetThuc
'
'             dlgCommonDialog.FileName = "D:\BangKeDauVao" + CStr(tdau) + IIf(tcuoi < 10, "0", "") + CStr(tcuoi) + CStr(pNamTC Mod 100) + ".XLS" '"*.XLS"
'          '   dlgCommonDialog.
'          '  On Error GoTo QuitSAS
'            dlgCommonDialog.ShowSave
'            On Error GoTo 0
'            pDataPath = dlgCommonDialog.FileName
'   ' pDataPath = "bangke.XLS"
'    FileCopy pCurDir + "REPORTS\Bangkemuavao.XLS", pDataPath '+ "bangkera.XLS"
'
'    Set xlapp = CreateObject("Excel.Application")
'    xlapp.Workbooks.Open pDataPath ' + "bangkera.XLS"
'    Set xlsheet = xlapp.Worksheets(1)
'
'   Dim Vowels As String
'
'    On Error GoTo 0
'
'
'    Dim tk2 As New ClsTaikhoan, Fx As Integer
'
'    Fx = IIf(Left(frmMain.LbCty(8).Caption, 2) = "35" Or Left(frmMain.LbCty(3).Caption, 2) = "64", 35, 0)
'
'    tk2.InitTaikhoanSohieu "33312"
'    If tk2.tk_id = GTGTKT_ID Then
'        ExecuteSQL5 "DELETE HoaDon.*  FROM ((HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo " _
'            & " WHERE (HethongTK.SoHieu LIKE '" + pVATV + "*') AND (TK.SoHieu LIKE '33312*'  OR InStr(ChungTu.GhiChu,'33312')>0) AND HoaDon.Loai=-1"
'
'    End If
'
'    GauGe.Max = 1
'
'    If TK.MaSo = 0 Then
'        Select Case tl
'            Case -3:
'                If Fx = 0 Then
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieuSoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HDBL=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                Else
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HDBL=1 AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                End If
'            Case -2:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'            Case -1:
'                If Fx = 0 Then
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                Else
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HDBL=0 OR KCT=0) AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                End If
'            Case 30:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND (TyLe=5 OR TyLe=10) AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'            Case Else
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND TyLe=" + CStr(tl) + " AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HDBL=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'        End Select
''        frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
'    Else
'        Select Case tl
'            Case -3:
'                If Fx = 0 Then
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND HDBL=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                Else
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND HDBL=1 AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                End If
'            Case -2:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'            Case -1:
'                If Fx = 0 Then
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                Else
'                    SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HDBL=0 OR KCT=0) AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'                End If
'            Case 30:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND (TyLe=5 OR TyLe=10) AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'            Case Else
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND TyLe=" + CStr(tl) + " AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HDBL=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
'        End Select
' '       frmMain.Rpt.Formulas(1) = "TenCn='" + TK.Ten + "'"
'  '      frmMain.Rpt.Formulas(5) = "DiaChi = '" + TK.GhiChu + "'"
'    End If
'
'    GauGe.Value = 1
'    If HD = 0 Then
'   '     frmMain.Rpt.ReportFileName = "BANGKEV3.RPT"
'    Else
'        If (tl > 0 And tl < 5) Or tl = -3 Then
'    '        frmMain.Rpt.ReportFileName = "BANGKEV4.RPT"
'        Else
'            If GetSetting(IniPath, "Invoice", "ListDetail", 0) = 1 Then
'
'                SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKHD(0) + " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HoaDon.Loai=-1 GROUP BY MaCT"
'                SetSQL "QCdt", "SELECT MaCT FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaVattu>0 AND MaLoai=1 GROUP BY MaCT HAVING Count(MaVattu)>1"
'                SetSQL "QDuPhong", "SELECT ChungTu.MaCT,TenVattu,SoPS2No AS SL,SoPS AS TT FROM (((ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN QCdt ON ChungTu.MaCT=QCdt.MaCT WHERE (TK_ID=" + CStr(TKVT_ID) + " OR HethongTK.Loai=6)"
'                SetSQL "QChitiet", "SELECT QNhatKy.*,TenVattu,SL,TT FROM QNhatKy LEFT JOIN QDuPhong ON QNhatKy.MaCT=QDuPhong.MaCT ORDER BY TyLe,NgayPH,SoHD"
'            Else
''                frmMain.Rpt.ReportFileName = "BANGKEV2.RPT"
'            End If
'        End If
'    End If
'    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
'
'
'    Dim i
'    Dim tongcong
'    tongcong = 0
'
'    Dim stt As Integer
'
'
'    i = 18
'    Dim sodong As String
'    sodong = i
'    stt = 1
'
'     Dim tt
'   ' Hang hoa dich vu khong chiu thue GTGT Tyle = -2
'    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
'    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,SoPS,ThanhTien,KhachHang.Ten,MST,SoHieu,sops,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky  where httt ='1' or httt ='...' order by NgayPH,SoHD ", dbOpenSnapshot) 'where tyle <>0
'    If rs.RecordCount <> 0 Then 'GoTo KetThuc where HTTT = '1'
'
'            Do While Not rs.EOF
'                sodong = i
'                xlsheet.Range("B" + sodong, "L" + sodong).Insert
'                xlsheet.Cells(i, 2) = stt
'                xlsheet.Cells(i, 3) = rs!KyHieu
'                xlsheet.Cells(i, 4) = rs!sohd
'
'               tt = rs!NgayPH
'                xlsheet.Cells(i, 5) = tt 'rs!NgayPH
'                xlsheet.Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'                xlsheet.Cells(i, 6) = MulticharToUnicode(UVowels, rs!Ten)
'                xlsheet.Cells(i, 7) = rs!mst
'                xlsheet.Cells(i, 8) = MulticharToUnicode(UVowels, rs!MatHang)
'                xlsheet.Cells(i, 9) = rs!ThanhTien
'                xlsheet.Cells(i, 10) = rs!TyLe
'                xlsheet.Cells(i, 11) = rs!sops
'                tongcong = tongcong + rs!ThanhTien
'                rs.MoveNext
'                i = i + 1
'                stt = stt + 1
'            Loop
'
'        '    xlsheet.Cells(i + 1, 9) = tongcong 'i + 1
'            sodong = sodong + 1
'         '    xlsheet.Range("B" + sodong, "L" + sodong).Delete
'         'i = i - 1
'    End If
'    i = i + 3
'    tongcong = 0
'    'stt = 1
'    ' chiu thue 0%
'    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
'    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,sops,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where httt ='2' order by NgayPH,SoHD ", dbOpenSnapshot) 'tyle = 0
'    If rs.RecordCount <> 0 Then ' GoTo KetThuc
'            Do While Not rs.EOF
'            sodong = i
'                xlsheet.Range("B" + sodong, "L" + sodong).Insert
'                xlsheet.Cells(i, 2) = stt
'                xlsheet.Cells(i, 3) = rs!KyHieu
'                xlsheet.Cells(i, 4) = rs!sohd
'
'               tt = rs!NgayPH
'                xlsheet.Cells(i, 5) = tt 'rs!NgayPH
'                xlsheet.Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'                xlsheet.Cells(i, 6) = MulticharToUnicode(UVowels, rs!Ten)
'                xlsheet.Cells(i, 7) = rs!mst
'                xlsheet.Cells(i, 8) = MulticharToUnicode(UVowels, rs!MatHang)
'                xlsheet.Cells(i, 9) = rs!ThanhTien
'                xlsheet.Cells(i, 10) = rs!TyLe
'                xlsheet.Cells(i, 11) = rs!sops
'                tongcong = tongcong + rs!ThanhTien
'                rs.MoveNext
'                i = i + 1
'               stt = stt + 1
'            Loop
'          '  xlsheet.Cells(i + 1, 9) = tongcong
'            sodong = sodong + 1
'          '  xlsheet.Range("B" + sodong, "L" + sodong).Delete
'          '  i = i - 1
'    End If
'    i = i + 3
'    'stt = 1
'    ' chiu thue 5%
'    tongcong = 0
'    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
'    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,sops,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where httt ='3' order by NgayPH,SoHD", dbOpenSnapshot) ' tyle = 5
'    If rs.RecordCount <> 0 Then ' GoTo KetThuc
'            Do While Not rs.EOF
'            sodong = i
'            xlsheet.Range("B" + sodong, "L" + sodong).Insert
'                xlsheet.Cells(i, 2) = stt
'                xlsheet.Cells(i, 3) = rs!KyHieu
'                xlsheet.Cells(i, 4) = rs!sohd
'
'               tt = rs!NgayPH
'                xlsheet.Cells(i, 5) = tt 'rs!NgayPH
'                xlsheet.Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'                xlsheet.Cells(i, 6) = MulticharToUnicode(UVowels, rs!Ten)
'                xlsheet.Cells(i, 7) = rs!mst
'                xlsheet.Cells(i, 8) = MulticharToUnicode(UVowels, rs!MatHang)
'                xlsheet.Cells(i, 9) = rs!ThanhTien
'                xlsheet.Cells(i, 10) = rs!TyLe
'                xlsheet.Cells(i, 11) = rs!sops
'                 tongcong = tongcong + rs!ThanhTien
'                rs.MoveNext
'                i = i + 1
'                stt = stt + 1
'              Loop
'         '   xlsheet.Cells(i + 1, 9) = tongcong
'            sodong = sodong + 1
'          '  xlsheet.Range("B" + sodong, "L" + sodong).Delete
'           ' i = i - 1
'    End If
'    i = i + 3
'    tongcong = 0
'
'   '  chiu thue 10%
'    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
'    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,sops,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where httt ='4' order by NgayPH,SoHD ", dbOpenSnapshot) 'tyle = 10
'    If rs.RecordCount <> 0 Then ' GoTo KetThuc
'                Do While Not rs.EOF
'                sodong = i
'                    xlsheet.Range("B" + sodong, "L" + sodong).Insert
'                    xlsheet.Cells(i, 2) = stt
'                    xlsheet.Cells(i, 3) = rs!KyHieu
'                    xlsheet.Cells(i, 4) = rs!sohd
'
'                    tt = rs!NgayPH
'
'                    xlsheet.Cells(i, 5) = tt 'rs!NgayPH
'                    xlsheet.Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'                    xlsheet.Cells(i, 6) = MulticharToUnicode(UVowels, rs!Ten)
'                    xlsheet.Cells(i, 7) = rs!mst
'                    xlsheet.Cells(i, 8) = MulticharToUnicode(UVowels, rs!MatHang)
'                    xlsheet.Cells(i, 9) = rs!ThanhTien
'                    xlsheet.Cells(i, 10) = rs!TyLe
'                    xlsheet.Cells(i, 11) = rs!sops
'                    'MsgBox rs!thue
'                    tongcong = tongcong + rs!ThanhTien
'                    rs.MoveNext
'                    i = i + 1
'                   'MsgBox rs!ThanhTien
'                    stt = stt + 1
'                Loop
'            '    xlsheet.Cells(i + 1, 9) = tongcong
'                sodong = sodong + 1
'            '    xlsheet.Range("B" + sodong, "L" + sodong).Delete
'              '  i = i - 1
'    End If
'    i = i + 3
'    tongcong = 0
' '   stt = 1
'   '  chiu thue tren 10%
'    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
'    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,sops,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where httt= '5' order by NgayPH,SoHD ", dbOpenSnapshot) 'tyle >10'
'    If rs.RecordCount <> 0 Then ' GoTo KetThuc
'            Do While Not rs.EOF
'            sodong = i
'                xlsheet.Range("B" + sodong, "L" + sodong).Insert
'                xlsheet.Cells(i, 2) = stt
'                xlsheet.Cells(i, 3) = rs!KyHieu
'                xlsheet.Cells(i, 4) = rs!sohd
'
'                tt = rs!NgayPH
'                xlsheet.Cells(i, 5) = tt 'rs!NgayPH
'                xlsheet.Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'                xlsheet.Cells(i, 6) = MulticharToUnicode(UVowels, rs!Ten)
'                xlsheet.Cells(i, 7) = rs!mst
'                xlsheet.Cells(i, 8) = MulticharToUnicode(UVowels, rs!MatHang)
'                xlsheet.Cells(i, 9) = rs!ThanhTien
'                xlsheet.Cells(i, 10) = rs!TyLe
'                xlsheet.Cells(i, 11) = rs!sops
'                tongcong = tongcong + rs!ThanhTien
'                rs.MoveNext
'                i = i + 1
'                 stt = stt + 1
'            Loop
'           ' xlsheet.Cells(i + 1, 9) = tongcong
'            sodong = sodong + 1
''            xlsheet.Range("B" + sodong, "L" + sodong).Delete
''            i = i - 1
'    End If
'    i = i + 3
'
'rs = Null
'KT:
'    xlapp.Workbooks(1).Save
'    xlapp.Workbooks.Close
'
'    Shell "EXPLORER.EXE " & pDataPath
'
'
'    Set xlsheet = Nothing
'    Set xlapp = Nothing
'
'KetThuc:
'GauGe.Value = 1
'InVATDauvao_exel = False
'End Function
'
'Private Function InVATDauRa_exel(tdau As Integer, tcuoi As Integer, tl As Integer, TK As ClsTaikhoan) As Boolean
'On Error GoTo ErrorHandler
'  Dim sql, pDataPath As String
'  Dim rs
'  Dim UVowels As String
'  Dim mDOMVowels As Class1
' 'Set MyUnicodeText = New Class1
'  Set mDOMVowels = New Class1
'
'        ' Read Unicode Text from file txtFileName and display in TextBox1(0)
'       'TextBox1(0).Text = MyUnicodeText.ReadUnicode(txtFileName)
'         ' UVowels = mDOMVowels.ReadUnicode(pCurDir & "UnicodeVowels.xml")
'          UVowels = mDOMVowels.ReadUnicode(GetLocalDirectory & "UnicodeVowels.xml")
'
'
'          PopulateListBoxFromFile lstVNI, "VNIVowelMap.txt", False
'           lstVNI.Font.Name = "Courier"
'    GauGe.Max = 1
'   ' Recycle pCurDir + "bangke.XLS"
'    On Error GoTo KetThuc
'
'             dlgCommonDialog.FileName = "D:\BangKeDauRa" + CStr(tdau) + IIf(tcuoi < 10, "0", "") + CStr(tcuoi) + CStr(pNamTC Mod 100) + ".XLS" '"*.XLS"
'          '   dlgCommonDialog.
'          '  On Error GoTo QuitSAS
'            dlgCommonDialog.ShowSave
'            On Error GoTo 0
'            pDataPath = dlgCommonDialog.FileName
'   ' pDataPath = "bangke.XLS"
'
'
'    FileCopy pCurDir + "REPORTS\bangkera.XLS", pDataPath '+ "bangkera.XLS"
'
'
'
'
'    Set xlapp = CreateObject("Excel.Application")
'    xlapp.Workbooks.Open pDataPath ' + "bangkera.XLS"
'    Set xlsheet = xlapp.Worksheets(1)
'
'
'
'
'    Dim Vowels As String
'
'    On Error GoTo 0
'
'
'    GauGe.Max = 1
'    If TK.MaSo = 0 Then
'        Select Case tl
'            Case -3:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -2:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -1:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case 30:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case Else
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'        End Select
'    Else
'        Select Case tl
'            Case -3:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -2:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case -1:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case 30:
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'            Case Else
'                SetSQL "QNhatky", "SELECT DISTINCTROW KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.SoHieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
'        End Select
'    End If
'    GauGe.Value = 1
'    If GetSetting(IniPath, "Invoice", "ListDetail", 0) = 1 Then
'        SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKHD(0) + " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HoaDon.Loai=1 GROUP BY MaCT"
'        SetSQL "QCdt", "SELECT MaCT FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaVattu>0 AND MaLoai=8 GROUP BY MaCT HAVING Count(MaVattu)>1"
'        SetSQL "QDuPhong", "SELECT ChungTu.MaCT,TenVattu AS MH,SoPS2Co AS SL,SoPS AS TT FROM (((ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN QCdt ON ChungTu.MaCT=QCdt.MaCT WHERE TK_ID=" + CStr(TKDT_ID)
'        SetSQL "QChitiet", "SELECT QNhatKy.*,QDuPhong.* FROM QNhatKy LEFT JOIN QDuPhong ON QNhatKy.MaCT=QDuPhong.MaCT ORDER BY TyLe,NgayPH,SoHD"
'        frmMain.Rpt.ReportFileName = "BANGKER6.RPT"
'    Else
'       ' frmMain.Rpt.ReportFileName = "BANGKER.RPT"
'    End If
'    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
'
'
'    Dim i
'    Dim tongcong
'    tongcong = 0
'
'    Dim stt As Integer
'
'     stt = 1
'    i = 18
'    Dim sodong As String
'    sodong = i
'    Dim tt
'   ' Hang hoa dich vu khong chiu thue GTGT Tyle = -2
'    'Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where tyle = -2", dbOpenSnapshot)
'    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where   HTTT = '2' or (HTTT = '3' and tyle =0 ) order by val(SoHD),NgayPH ", dbOpenSnapshot) 'HTTT = '2' or (HTTT = '3' and
'
'    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
'    If rs.RecordCount <> 0 Then 'GoTo KetThuc
'
'            Do While Not rs.EOF
'                sodong = i
'                xlsheet.Range("B" + sodong, "L" + sodong).Insert
'                xlsheet.Cells(i, 2) = stt
'                xlsheet.Cells(i, 3) = rs!KyHieu
'                xlsheet.Cells(i, 4) = rs!sohd
'                  tt = rs!NgayPH
'                xlsheet.Cells(i, 5) = tt 'rs!NgayPH
'                xlsheet.Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'                xlsheet.Cells(i, 6) = MulticharToUnicode(UVowels, rs!Ten)
'                xlsheet.Cells(i, 7) = rs!mst
'                xlsheet.Cells(i, 8) = MulticharToUnicode(UVowels, rs!MatHang)
'                xlsheet.Cells(i, 9) = rs!ThanhTien
'                xlsheet.Cells(i, 10) = rs!thue
'                tongcong = tongcong + rs!thue
'                rs.MoveNext
'                i = i + 1
'                stt = stt + 1
'            Loop
'            xlsheet.Cells(i + 1, 10) = tongcong
'            sodong = sodong + 1
'            xlsheet.Range("B" + sodong, "L" + sodong).Delete
'            i = i - 1
'    End If
'    i = i + 3
'    tongcong = 0
'   ' stt = 1
'    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
'    ' chiu thue 0%
'    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where (httt ='1' or httt = '3' or httt = '...') and tyle = 0 order by val(SoHD),NgayPH ", dbOpenSnapshot)
'    If rs.RecordCount <> 0 Then ' GoTo KetThuc
'            Do While Not rs.EOF
'            sodong = i
'                xlsheet.Range("B" + sodong, "L" + sodong).Insert
'                xlsheet.Cells(i, 2) = stt
'                xlsheet.Cells(i, 3) = rs!KyHieu
'                xlsheet.Cells(i, 4) = rs!sohd
'                tt = rs!NgayPH
'                xlsheet.Cells(i, 5) = tt
'                xlsheet.Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'                xlsheet.Cells(i, 6) = MulticharToUnicode(UVowels, rs!Ten)
'                xlsheet.Cells(i, 7) = rs!mst
'                xlsheet.Cells(i, 8) = MulticharToUnicode(UVowels, rs!MatHang)
'                xlsheet.Cells(i, 9) = rs!ThanhTien
'                xlsheet.Cells(i, 10) = rs!thue
'                tongcong = tongcong + rs!thue
'                rs.MoveNext
'                i = i + 1
'               stt = stt + 1
'            Loop
'            xlsheet.Cells(i + 1, 10) = tongcong
'            sodong = sodong + 1
'            xlsheet.Range("B" + sodong, "L" + sodong).Delete
'            i = i - 1
'    End If
'    i = i + 3
'   ' stt = 1
'    ' chiu thue 5%
'    tongcong = 0
'    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
'    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where (httt ='1' or httt = '3' or httt = '...' ) and tyle = 5 order by val(SoHD),NgayPH ", dbOpenSnapshot)
'    If rs.RecordCount <> 0 Then ' GoTo KetThuc
'            Do While Not rs.EOF
'            sodong = i
'            xlsheet.Range("B" + sodong, "L" + sodong).Insert
'                  xlsheet.Cells(i, 2) = stt
'                xlsheet.Cells(i, 3) = rs!KyHieu
'                xlsheet.Cells(i, 4) = rs!sohd
'                tt = rs!NgayPH
'                xlsheet.Cells(i, 5) = tt
'                xlsheet.Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'                xlsheet.Cells(i, 6) = MulticharToUnicode(UVowels, rs!Ten)
'                xlsheet.Cells(i, 7) = rs!mst
'                xlsheet.Cells(i, 8) = MulticharToUnicode(UVowels, rs!MatHang)
'                xlsheet.Cells(i, 9) = rs!ThanhTien
'                xlsheet.Cells(i, 10) = rs!thue
'                 tongcong = tongcong + rs!thue
'                rs.MoveNext
'                i = i + 1
'                 stt = stt + 1
'              Loop
'            xlsheet.Cells(i + 1, 10) = tongcong
'            sodong = sodong + 1
'            xlsheet.Range("B" + sodong, "L" + sodong).Delete
'            i = i - 1
'    End If
'    i = i + 3
'    tongcong = 0
'   ' stt = 1
'    ' chiu thue 10%
'    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
'    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where (httt ='1' or httt = '3'  or httt = '4' or httt = '...') and tyle = 10 order by val(SoHD),NgayPH ", dbOpenSnapshot)
'    If rs.RecordCount <> 0 Then ' GoTo KetThuc
'                Do While Not rs.EOF
'                sodong = i
'                    xlsheet.Range("B" + sodong, "L" + sodong).Insert
'                    xlsheet.Cells(i, 2) = stt
'                    xlsheet.Cells(i, 3) = rs!KyHieu
'                    xlsheet.Cells(i, 4) = rs!sohd
'                    tt = rs!NgayPH
'                    xlsheet.Cells(i, 5) = tt
'                      xlsheet.Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'                    xlsheet.Cells(i, 6) = MulticharToUnicode(UVowels, rs!Ten)
'                    xlsheet.Cells(i, 7) = rs!mst
'                    xlsheet.Cells(i, 8) = MulticharToUnicode(UVowels, rs!MatHang)
'                    xlsheet.Cells(i, 9) = rs!ThanhTien
'                    xlsheet.Cells(i, 10) = rs!thue
'                    'MsgBox rs!thue
'                    tongcong = tongcong + rs!thue
'                    rs.MoveNext
'                    i = i + 1
'                   'MsgBox rs!ThanhTien
'                    stt = stt + 1
'                Loop
'                xlsheet.Cells(i + 1, 10) = tongcong
'                sodong = sodong + 1
'                xlsheet.Range("B" + sodong, "L" + sodong).Delete
'                i = i - 1
'    End If
'    i = i + 3
'    tongcong = 0
'    stt = 1
'    ' chiu thue tren 10%
'    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
'    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where (httt ='1' or httt = '3' or httt = '...' ) and tyle >10 order by val(SoHD),NgayPH ", dbOpenSnapshot)
'    If rs.RecordCount = 0 Then ' GoTo KetThuc
'            Do While Not rs.EOF
'            sodong = i
'                xlsheet.Range("B" + sodong, "L" + sodong).Insert
'                xlsheet.Cells(i, 2) = stt
'                xlsheet.Cells(i, 3) = rs!KyHieu
'                xlsheet.Cells(i, 4) = rs!sohd
'                tt = rs!NgayPH
'                xlsheet.Cells(i, 5) = tt
'                xlsheet.Cells(i, 5).NumberFormat = "dd/mm/yyyy"
'                xlsheet.Cells(i, 6) = MulticharToUnicode(UVowels, rs!Ten)
'                xlsheet.Cells(i, 7) = rs!mst
'                xlsheet.Cells(i, 8) = MulticharToUnicode(UVowels, rs!MatHang)
'                xlsheet.Cells(i, 9) = rs!ThanhTien
'                xlsheet.Cells(i, 10) = rs!thue
'                tongcong = tongcong + rs!thue
'                rs.MoveNext
'                i = i + 1
'                stt = stt + 1
'            Loop
'            xlsheet.Cells(i + 1, 10) = tongcong
'            sodong = sodong + 1
'           ' xlsheet.Range("B" + sodong, "L" + sodong).Delete
'            i = i - 1
'    End If
'    i = i + 3
'
'rs = Null
'KT:
'    xlapp.Workbooks(1).Save
'    xlapp.Workbooks.Close
'
'    Shell "EXPLORER.EXE " & pDataPath
'
'
'    Set xlsheet = Nothing
'    Set xlapp = Nothing
'
'KetThuc:
'
'GauGe.Value = 1
'InVATDauRa_exel = False
' Exit Function
'ErrorHandler:
'    ' MsgBox "thong bao"
'     InVATDauRa_exel = False
'End Function
'

Private Function InVATDauRa_exel(tdau As Integer, tcuoi As Integer, tl As Integer, TK As ClsTaikhoan) As Boolean
On Error GoTo ErrorHandler
  Dim sql, pDataPath As String
  Dim rs
  Dim UVowels As String
  Dim mDOMVowels As Class1
 'Set MyUnicodeText = New Class1
  Set mDOMVowels = New Class1

        ' Read Unicode Text from file txtFileName and display in TextBox1(0)
       'TextBox1(0).Text = MyUnicodeText.ReadUnicode(txtFileName)
         ' UVowels = mDOMVowels.ReadUnicode(pCurDir & "UnicodeVowels.xml")
          UVowels = mDOMVowels.ReadUnicode(GetLocalDirectory & "UnicodeVowels.xml")
   
          
          PopulateListBoxFromFile lstVNI, "VNIVowelMap.txt", False
           lstVNI.Font.Name = "Courier"
    GauGe.Max = 1
   ' Recycle pCurDir + "bangke.XLS"
    On Error GoTo KetThuc

             dlgCommonDialog.FileName = "D:\BangKeDauRa" + CStr(tdau) + IIf(tcuoi < 10, "0", "") + CStr(tcuoi) + CStr(pNamTC Mod 100) + ".XLS" '"*.XLS"
          '   dlgCommonDialog.
          '  On Error GoTo QuitSAS
            dlgCommonDialog.ShowSave
            On Error GoTo 0
            pDataPath = dlgCommonDialog.FileName
   ' pDataPath = "bangke.XLS"
  
   
    FileCopy pCurDir + "REPORTS\bangkera.XLS", pDataPath '+ "bangkera.XLS"
   
 


    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pDataPath ' + "bangkera.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    



    Dim Vowels As String

    On Error GoTo 0


    GauGe.Max = 1
    If TK.MaSo = 0 Then
        Select Case tl
            Case -3:
                SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case -2:
                SetSQL "QNhatky", "SELECT DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case -1:
                SetSQL "QNhatky", "SELECT DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case 30:
                SetSQL "QNhatky", "SELECT DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case Else
                SetSQL "QNhatky", "SELECT DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
        End Select
    Else
        Select Case tl
            Case -3:
                SetSQL "QNhatky", "SELECT DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND HDBL=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case -2:
                SetSQL "QNhatky", "SELECT DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case -1:
                SetSQL "QNhatky", "SELECT DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case 30:
                SetSQL "QNhatky", "SELECT DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND (TyLe=5 OR TyLe=10) AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
            Case Else
                SetSQL "QNhatky", "SELECT DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,IIF(TK_ID=" + CStr(GTGTPN_ID) + ",SoPS,-SoPS) AS Thue,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,KCT FROM " + ChungTu2TKHD(2) + " WHERE HoaDon.Loai=1 AND TyLe=" + CStr(tl) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH"
        End Select
    End If
    GauGe.Value = 1
    If GetSetting(IniPath, "Invoice", "ListDetail", 0) = 1 Then
        SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKHD(0) + " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HoaDon.Loai=1 GROUP BY MaCT"
        SetSQL "QCdt", "SELECT MaCT FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaVattu>0 AND MaLoai=8 GROUP BY MaCT HAVING Count(MaVattu)>1"
        SetSQL "QDuPhong", "SELECT ChungTu.MaCT,TenVattu AS MH,SoPS2Co AS SL,SoPS AS TT FROM (((ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN QCdt ON ChungTu.MaCT=QCdt.MaCT WHERE TK_ID=" + CStr(TKDT_ID)
        SetSQL "QChitiet", "SELECT QNhatKy.*,QDuPhong.* FROM QNhatKy LEFT JOIN QDuPhong ON QNhatKy.MaCT=QDuPhong.MaCT ORDER BY TyLe,NgayPH,SoHD"
        frmMain.Rpt.ReportFileName = "BANGKER6.RPT"
    Else
       ' frmMain.Rpt.ReportFileName = "BANGKER.RPT"
    End If
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn

   
    Dim i
    Dim tongcong
    tongcong = 0
   
    Dim stt As Integer
    
     stt = 1
    i = 18
    Dim sodong As String
    sodong = i
    Dim tt
   ' Hang hoa dich vu khong chiu thue GTGT Tyle = -2
    'Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where tyle = -2", dbOpenSnapshot)
    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where   HTTT = '2' or (HTTT = '3' and tyle =0 ) order by val(SoHD),NgayPH ", dbOpenSnapshot) 'HTTT = '2' or (HTTT = '3' and
    
    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
    If rs.RecordCount <> 0 Then 'GoTo KetThuc
     
          Do While Not rs.EOF
                sodong = i
                    xlsheet.Range("B" + sodong, "N" + sodong).Insert
                    xlsheet.Cells(i, 2) = stt
                    
                '    xlsheet.Cells(i, 3) = rs!LoaiHoaDon
                    xlsheet.Cells(i, 3) = rs!MauSoHD
                    
                    xlsheet.Cells(i, 4) = rs!KyHieu
                    xlsheet.Cells(i, 5) = rs!sohd
                    xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                    tt = rs!NgayPH
                    xlsheet.Cells(i, 6) = tt
                      xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                    xlsheet.Cells(i, 7) = MulticharToUnicode(UVowels, rs!Ten)
                    xlsheet.Cells(i, 8) = rs!mst
                    xlsheet.Cells(i, 9) = MulticharToUnicode(UVowels, rs!MatHang)
                    xlsheet.Cells(i, 10) = rs!ThanhTien
                    xlsheet.Cells(i, 11) = rs!thue
                    'MsgBox rs!thue
                    tongcong = tongcong + rs!thue
                    rs.MoveNext
                    i = i + 1
                   'MsgBox rs!ThanhTien
                    stt = stt + 1
                Loop
                xlsheet.Cells(i + 1, 11) = tongcong
            sodong = sodong + 1
            xlsheet.Range("B" + sodong, "N" + sodong).Delete
            i = i - 1
    End If
    i = i + 3
    tongcong = 0
   ' stt = 1
    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
    ' chiu thue 0%
    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where (httt ='1' or httt = '3' or httt = '...') and tyle = 0 order by val(SoHD),NgayPH ", dbOpenSnapshot)
    If rs.RecordCount <> 0 Then ' GoTo KetThuc
                Do While Not rs.EOF
                sodong = i
                    xlsheet.Range("B" + sodong, "N" + sodong).Insert
                    xlsheet.Cells(i, 2) = stt
                    
                '    xlsheet.Cells(i, 3) = rs!LoaiHoaDon
                    xlsheet.Cells(i, 3) = rs!MauSoHD
                    
                    xlsheet.Cells(i, 4) = rs!KyHieu
                    xlsheet.Cells(i, 5) = rs!sohd
                    xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                    tt = rs!NgayPH
                    xlsheet.Cells(i, 6) = tt
                      xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                    xlsheet.Cells(i, 7) = MulticharToUnicode(UVowels, rs!Ten)
                    xlsheet.Cells(i, 8) = rs!mst
                    xlsheet.Cells(i, 9) = MulticharToUnicode(UVowels, rs!MatHang)
                    xlsheet.Cells(i, 10) = rs!ThanhTien
                    xlsheet.Cells(i, 11) = rs!thue
                    'MsgBox rs!thue
                    tongcong = tongcong + rs!thue
                    rs.MoveNext
                    i = i + 1
                   'MsgBox rs!ThanhTien
                    stt = stt + 1
                Loop
                xlsheet.Cells(i + 1, 11) = tongcong
            sodong = sodong + 1
            xlsheet.Range("B" + sodong, "N" + sodong).Delete
            i = i - 1
    End If
    i = i + 3
   ' stt = 1
    ' chiu thue 5%
    tongcong = 0
    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where (httt ='1' or httt = '3' or httt = '...' ) and tyle = 5 order by val(SoHD),NgayPH ", dbOpenSnapshot)
    If rs.RecordCount <> 0 Then ' GoTo KetThuc
                Do While Not rs.EOF
                sodong = i
                    xlsheet.Range("B" + sodong, "N" + sodong).Insert
                    xlsheet.Cells(i, 2) = stt
                    
                '    xlsheet.Cells(i, 3) = rs!LoaiHoaDon
                    xlsheet.Cells(i, 3) = rs!MauSoHD
                    
                    xlsheet.Cells(i, 4) = rs!KyHieu
                    xlsheet.Cells(i, 5) = rs!sohd
                    xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                    tt = rs!NgayPH
                    xlsheet.Cells(i, 6) = tt
                      xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                    xlsheet.Cells(i, 7) = MulticharToUnicode(UVowels, rs!Ten)
                    xlsheet.Cells(i, 8) = rs!mst
                    xlsheet.Cells(i, 9) = MulticharToUnicode(UVowels, rs!MatHang)
                    xlsheet.Cells(i, 10) = rs!ThanhTien
                    xlsheet.Cells(i, 11) = rs!thue
                    'MsgBox rs!thue
                    tongcong = tongcong + rs!thue
                    rs.MoveNext
                    i = i + 1
                   'MsgBox rs!ThanhTien
                    stt = stt + 1
                Loop
                xlsheet.Cells(i + 1, 11) = tongcong
            sodong = sodong + 1
            xlsheet.Range("B" + sodong, "N" + sodong).Delete
            i = i - 1
    End If
    i = i + 3
    tongcong = 0
   ' stt = 1
    ' chiu thue 10%
    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where (httt ='1' or httt = '3'  or httt = '4' or httt = '...') and tyle = 10 order by val(SoHD),NgayPH ", dbOpenSnapshot)
    If rs.RecordCount <> 0 Then ' GoTo KetThuc
                     Do While Not rs.EOF
                sodong = i
                    xlsheet.Range("B" + sodong, "N" + sodong).Insert
                    xlsheet.Cells(i, 2) = stt
                    
                '    xlsheet.Cells(i, 3) = rs!LoaiHoaDon
                    xlsheet.Cells(i, 3) = rs!MauSoHD
                    
                    xlsheet.Cells(i, 4) = rs!KyHieu
                    xlsheet.Cells(i, 5) = rs!sohd
                    xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                    tt = rs!NgayPH
                    xlsheet.Cells(i, 6) = tt
                      xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                    xlsheet.Cells(i, 7) = MulticharToUnicode(UVowels, rs!Ten)
                    xlsheet.Cells(i, 8) = rs!mst
                    xlsheet.Cells(i, 9) = MulticharToUnicode(UVowels, rs!MatHang)
                    xlsheet.Cells(i, 10) = rs!ThanhTien
                    xlsheet.Cells(i, 11) = rs!thue
                    'MsgBox rs!thue
                    tongcong = tongcong + rs!thue
                    rs.MoveNext
                    i = i + 1
                   'MsgBox rs!ThanhTien
                    stt = stt + 1
                Loop
                xlsheet.Cells(i + 1, 11) = tongcong
                sodong = sodong + 1
                xlsheet.Range("B" + sodong, "N" + sodong).Delete
                i = i - 1
    End If
    i = i + 3
    tongcong = 0
    stt = 1
    ' chiu thue tren 10%
    xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW  ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,Thue,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where (httt ='1' or httt = '3' or httt = '...' ) and tyle >10 order by val(SoHD),NgayPH ", dbOpenSnapshot)
    If rs.RecordCount = 0 Then ' GoTo KetThuc
                 Do While Not rs.EOF
                sodong = i
                    xlsheet.Range("B" + sodong, "N" + sodong).Insert
                    xlsheet.Cells(i, 2) = stt
                    
                '    xlsheet.Cells(i, 3) = rs!LoaiHoaDon
                    xlsheet.Cells(i, 3) = rs!MauSoHD
                    
                    xlsheet.Cells(i, 4) = rs!KyHieu
                    xlsheet.Cells(i, 5) = rs!sohd
                    xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                    tt = rs!NgayPH
                    xlsheet.Cells(i, 6) = tt
                      xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                    xlsheet.Cells(i, 7) = MulticharToUnicode(UVowels, rs!Ten)
                    xlsheet.Cells(i, 8) = rs!mst
                    xlsheet.Cells(i, 9) = MulticharToUnicode(UVowels, rs!MatHang)
                    xlsheet.Cells(i, 10) = rs!ThanhTien
                    xlsheet.Cells(i, 11) = rs!thue
                    'MsgBox rs!thue
                    tongcong = tongcong + rs!thue
                    rs.MoveNext
                    i = i + 1
                   'MsgBox rs!ThanhTien
                    stt = stt + 1
                Loop
                xlsheet.Cells(i + 1, 11) = tongcong
           ' xlsheet.Range("B" + sodong, "L" + sodong).Delete
            i = i - 1
    End If
    i = i + 3
       
rs = Null
KT:
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
 
    Shell "EXPLORER.EXE " & pDataPath
    
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
  
KetThuc:

GauGe.Value = 1
InVATDauRa_exel = False
 Exit Function
ErrorHandler:
    ' MsgBox "thong bao"
     InVATDauRa_exel = False
End Function


' in bang ke dau vao exl

Private Function InVATDauvao_exel(tdau As Integer, tcuoi As Integer, tl As Integer, HD As Integer, TK As ClsTaikhoan) As Boolean
    Dim sql, pDataPath As String
  Dim rs
  Dim UVowels As String
  Dim mDOMVowels As Class1
 'Set MyUnicodeText = New Class1
  Set mDOMVowels = New Class1

        ' Read Unicode Text from file txtFileName and display in TextBox1(0)
       'TextBox1(0).Text = MyUnicodeText.ReadUnicode(txtFileName)
          'UVowels = mDOMVowels.ReadUnicode(pCurDir & "UnicodeVowels.xml")
          UVowels = mDOMVowels.ReadUnicode(GetLocalDirectory & "UnicodeVowels.xml")
   
           PopulateListBoxFromFile lstVNI, "VNIVowelMap.txt", False
           lstVNI.Font.Name = "Courier"
    GauGe.Max = 1
   ' Recycle pCurDir + "bangke.XLS"
    On Error GoTo KetThuc

             dlgCommonDialog.FileName = "D:\BangKeDauVao" + CStr(tdau) + IIf(tcuoi < 10, "0", "") + CStr(tcuoi) + CStr(pNamTC Mod 100) + ".XLS" '"*.XLS"
          '   dlgCommonDialog.
          '  On Error GoTo QuitSAS
            dlgCommonDialog.ShowSave
            On Error GoTo 0
            pDataPath = dlgCommonDialog.FileName
   ' pDataPath = "bangke.XLS"
    FileCopy pCurDir + "REPORTS\Bangkemuavao.XLS", pDataPath '+ "bangkera.XLS"

    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pDataPath ' + "bangkera.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    
   Dim Vowels As String

    On Error GoTo 0


    Dim tk2 As New ClsTaikhoan, Fx As Integer
    
    Fx = IIf(Left(frmMain.LbCty(8).Caption, 2) = "35" Or Left(frmMain.LbCty(3).Caption, 2) = "64", 35, 0)
    
    tk2.InitTaikhoanSohieu "33312"
    If tk2.tk_id = GTGTKT_ID Then
        ExecuteSQL5 "DELETE HoaDon.*  FROM ((HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo " _
            & " WHERE (HethongTK.SoHieu LIKE '" + pVATV + "*') AND (TK.SoHieu LIKE '33312*'  OR InStr(ChungTu.GhiChu,'33312')>0) AND HoaDon.Loai=-1"
        
    End If
    
    GauGe.Max = 1
    
    If TK.MaSo = 0 Then
        Select Case tl
            Case -3:
                If Fx = 0 Then
                    SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieuSoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HDBL=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                Else
                    SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HDBL=1 AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                End If
            Case -2:
                SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
            Case -1:
                If Fx = 0 Then
                    SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                Else
                    SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND (HDBL=0 OR KCT=0) AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                End If
            Case 30:
                SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND (TyLe=5 OR TyLe=10) AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
            Case Else
                SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM " + ChungTu2TKHD(10) + " WHERE Loai=-1 AND TyLe=" + CStr(tl) + " AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HDBL=0 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
        End Select
'        frmMain.Rpt.Formulas(5) = "DiaChi = '" + frmMain.LbCty(2).Caption + "'"
    Else
        Select Case tl
            Case -3:
                If Fx = 0 Then
                    SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HDBL=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                Else
                    SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND HDBL=1 AND KCT=1 AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                End If
            Case -2:
                SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=1 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
            Case -1:
                If Fx = 0 Then
                    SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                Else
                    SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HDBL=0 OR KCT=0) AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
                End If
            Case 30:
                SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND (TyLe=5 OR TyLe=10) AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
            Case Else
                SetSQL "QNhatky", "SELECT DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,ChungTu.NgayCT as NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,KhachHang.MST,ChungTu.SoHieu,SoPS,KhachHang.DiaChi,TyLe,HTTT,MauSo,MaCT,HoaDon.MaSo,KCT FROM (" + ChungTu2TKHD(10) + ") INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE HoaDon.Loai=-1 AND TyLe=" + CStr(tl) + " AND HD=" + CStr(HD) + " AND " + WThang("ThangCT", tdau, tcuoi) + " AND KCT=0 AND HDBL=0 AND HethongTK.Sohieu LIKE '" + TK.sohieu + "*' AND (HoaDon.DC=0 OR HD=1) ORDER BY NgayPH,MaCT"
        End Select
 '       frmMain.Rpt.Formulas(1) = "TenCn='" + TK.Ten + "'"
  '      frmMain.Rpt.Formulas(5) = "DiaChi = '" + TK.GhiChu + "'"
    End If

    GauGe.Value = 1
    If HD = 0 Then
   '     frmMain.Rpt.ReportFileName = "BANGKEV3.RPT"
    Else
        If (tl > 0 And tl < 5) Or tl = -3 Then
    '        frmMain.Rpt.ReportFileName = "BANGKEV4.RPT"
        Else
            If GetSetting(IniPath, "Invoice", "ListDetail", 0) = 1 Then
            
                SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKHD(0) + " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND HoaDon.Loai=-1 GROUP BY MaCT"
                SetSQL "QCdt", "SELECT MaCT FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaVattu>0 AND MaLoai=1 GROUP BY MaCT HAVING Count(MaVattu)>1"
                SetSQL "QDuPhong", "SELECT ChungTu.MaCT,TenVattu,SoPS2No AS SL,SoPS AS TT FROM (((ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN QCdt ON ChungTu.MaCT=QCdt.MaCT WHERE (TK_ID=" + CStr(TKVT_ID) + " OR HethongTK.Loai=6)"
                SetSQL "QChitiet", "SELECT QNhatKy.*,TenVattu,SL,TT FROM QNhatKy LEFT JOIN QDuPhong ON QNhatKy.MaCT=QDuPhong.MaCT ORDER BY TyLe,NgayPH,SoHD"
            Else
'                frmMain.Rpt.ReportFileName = "BANGKEV2.RPT"
            End If
        End If
    End If
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn

   
    Dim i
    Dim tongcong
    tongcong = 0
   
    Dim stt As Integer
    
    
    i = 18
    Dim sodong As String
    sodong = i
    stt = 1
    
     Dim tt As Date
   ' Hang hoa dich vu khong chiu thue GTGT Tyle = -2
   ' xlsheet.Cells(i - 1, 5).NumberFormat = "dd/mm/yyyy"
    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,NgayPH,MatHang,SoLuong,SoPS,ThanhTien,KhachHang.Ten,MST,SoHieu,sops,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky  where httt ='1' or httt ='...' order by NgayPH,SoHD ", dbOpenSnapshot) 'where tyle <>0
    If rs.RecordCount <> 0 Then 'GoTo KetThuc where HTTT = '1'
     
            Do While Not rs.EOF
                sodong = i
                xlsheet.Range("B" + sodong, "N" + sodong).Insert
                xlsheet.Cells(i, 2) = stt
             '   xlsheet.Cells(i, 3) = rs!LoaiHoaDon
                  xlsheet.Cells(i, 3) = rs!MauSoHD
                xlsheet.Cells(i, 4) = rs!KyHieu
                xlsheet.Cells(i, 5) = rs!sohd
               
                tt = rs!NgayPH
                xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                xlsheet.Cells(i, 6) = tt 'rs!NgayPH
                
                xlsheet.Cells(i, 7) = MulticharToUnicode(UVowels, rs!Ten)
                xlsheet.Cells(i, 8) = rs!mst
                xlsheet.Cells(i, 9) = MulticharToUnicode(UVowels, rs!MatHang)
                  xlsheet.Cells(i, 10).NumberFormat = "##,###,###,###"
                xlsheet.Cells(i, 10) = Int(rs!ThanhTien)
                xlsheet.Cells(i, 11) = Int(rs!TyLe)
                 xlsheet.Cells(i, 12).NumberFormat = "##,###,###,###"
                xlsheet.Cells(i, 12) = rs!sops
                tongcong = tongcong + rs!ThanhTien
                rs.MoveNext
                i = i + 1
                stt = stt + 1
            Loop
                
               
        '    xlsheet.Cells(i + 1, 9) = tongcong 'i + 1
            sodong = sodong + 1
         '    xlsheet.Range("B" + sodong, "L" + sodong).Delete
         'i = i - 1
         '   xlsheet.Cells(i, 3) = rs!LoaiHoaDon
         '           xlsheet.Cells(i, 4) = rs!MauSoHD
    End If
    i = i + 3
    tongcong = 0
    'stt = 1
    ' chiu thue 0%

    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,sops,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where httt ='2' order by NgayPH,SoHD ", dbOpenSnapshot) 'tyle = 0
    If rs.RecordCount <> 0 Then ' GoTo KetThuc
           Do While Not rs.EOF
            sodong = i
                xlsheet.Range("B" + sodong, "N" + sodong).Insert
                    xlsheet.Cells(i, 2) = stt
             '   xlsheet.Cells(i, 3) = rs!LoaiHoaDon
                 xlsheet.Cells(i, 3) = rs!MauSoHD
                xlsheet.Cells(i, 4) = rs!KyHieu
                xlsheet.Cells(i, 5) = rs!sohd
               
                tt = rs!NgayPH
                xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                xlsheet.Cells(i, 6) = tt 'rs!NgayPH
                
                xlsheet.Cells(i, 7) = MulticharToUnicode(UVowels, rs!Ten)
                xlsheet.Cells(i, 8) = rs!mst
                xlsheet.Cells(i, 9) = MulticharToUnicode(UVowels, rs!MatHang)
                  xlsheet.Cells(i, 10).NumberFormat = "##,###,###,###"
                xlsheet.Cells(i, 10) = Int(rs!ThanhTien)
                xlsheet.Cells(i, 11) = Int(rs!TyLe)
                 xlsheet.Cells(i, 12).NumberFormat = "##,###,###,###"
                xlsheet.Cells(i, 12) = rs!sops
                tongcong = tongcong + rs!ThanhTien
                rs.MoveNext
                i = i + 1
                stt = stt + 1
            Loop
                
          '  xlsheet.Cells(i + 1, 9) = tongcong
            sodong = sodong + 1
          '  xlsheet.Range("B" + sodong, "L" + sodong).Delete
          '  i = i - 1
    End If
    i = i + 3
    'stt = 1
    ' chiu thue 5%
    tongcong = 0
    
    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,sops,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where httt ='3' order by NgayPH,SoHD", dbOpenSnapshot) ' tyle = 5
    If rs.RecordCount <> 0 Then ' GoTo KetThuc
           Do While Not rs.EOF
           sodong = i
                xlsheet.Range("B" + sodong, "N" + sodong).Insert
                    xlsheet.Cells(i, 2) = stt
             '   xlsheet.Cells(i, 3) = rs!LoaiHoaDon
               xlsheet.Cells(i, 3) = rs!MauSoHD
                xlsheet.Cells(i, 4) = rs!KyHieu
                xlsheet.Cells(i, 5) = rs!sohd
               
                tt = rs!NgayPH
                xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                xlsheet.Cells(i, 6) = tt 'rs!NgayPH
                
                xlsheet.Cells(i, 7) = MulticharToUnicode(UVowels, rs!Ten)
                xlsheet.Cells(i, 8) = rs!mst
                xlsheet.Cells(i, 9) = MulticharToUnicode(UVowels, rs!MatHang)
                  xlsheet.Cells(i, 10).NumberFormat = "##,###,###,###"
                xlsheet.Cells(i, 10) = Int(rs!ThanhTien)
                xlsheet.Cells(i, 11) = Int(rs!TyLe)
                 xlsheet.Cells(i, 12).NumberFormat = "##,###,###,###"
                xlsheet.Cells(i, 12) = rs!sops
                tongcong = tongcong + rs!ThanhTien
                rs.MoveNext
                i = i + 1
                stt = stt + 1
            Loop
                
         '   xlsheet.Cells(i + 1, 9) = tongcong
            sodong = sodong + 1
          '  xlsheet.Range("B" + sodong, "L" + sodong).Delete
           ' i = i - 1
    End If
    i = i + 3
    tongcong = 0
    
   '  chiu thue 10%
  
    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,sops,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where httt ='4' order by NgayPH,SoHD ", dbOpenSnapshot) 'tyle = 10
    If rs.RecordCount <> 0 Then ' GoTo KetThuc
           Do While Not rs.EOF
            sodong = i
                xlsheet.Range("B" + sodong, "N" + sodong).Insert
                xlsheet.Cells(i, 2) = stt
                'xlsheet.Cells(i, 3) = rs!LoaiHoaDon
                 xlsheet.Cells(i, 3) = rs!MauSoHD
                xlsheet.Cells(i, 4) = rs!KyHieu
                xlsheet.Cells(i, 5) = rs!sohd
               
                tt = rs!NgayPH
                xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                xlsheet.Cells(i, 6) = tt 'rs!NgayPH
                
                xlsheet.Cells(i, 7) = MulticharToUnicode(UVowels, rs!Ten)
                xlsheet.Cells(i, 8) = rs!mst
                xlsheet.Cells(i, 9) = MulticharToUnicode(UVowels, rs!MatHang)
                  xlsheet.Cells(i, 10).NumberFormat = "##,###,###,###"
                xlsheet.Cells(i, 10) = Int(rs!ThanhTien)
                xlsheet.Cells(i, 11) = Int(rs!TyLe)
                 xlsheet.Cells(i, 12).NumberFormat = "##,###,###,###"
                xlsheet.Cells(i, 12) = rs!sops
                tongcong = tongcong + rs!ThanhTien
                rs.MoveNext
                i = i + 1
                stt = stt + 1
            Loop
                'xlsheet.Cells(i + 1, 12) = tongcong
            '    xlsheet.Cells(i + 1, 9) = tongcong
                sodong = sodong + 1
            '    xlsheet.Range("B" + sodong, "L" + sodong).Delete
              '  i = i - 1
    End If
    i = i + 3
    tongcong = 0
 '   stt = 1
   '  chiu thue tren 10%
    Set rs = DBKetoan.OpenRecordset("select  DISTINCTROW ChungTu.MauSoHD ,ChungTu.LoaiHoaDon,KyHieu,SoHD,NgayPH,MatHang,SoLuong,ThanhTien,KhachHang.Ten,MST,SoHieu,sops,DiaChi,TyLe,HTTT,MauSo,MaCT,KCT  from QNhatky where httt= '5' order by NgayPH,SoHD ", dbOpenSnapshot) 'tyle >10'
    If rs.RecordCount <> 0 Then ' GoTo KetThuc
   Do While Not rs.EOF
             sodong = i
                xlsheet.Range("B" + sodong, "N" + sodong).Insert
                xlsheet.Cells(i, 2) = stt
                'xlsheet.Cells(i, 3) = rs!LoaiHoaDon
               xlsheet.Cells(i, 3) = rs!MauSoHD
                xlsheet.Cells(i, 4) = rs!KyHieu
                xlsheet.Cells(i, 5) = rs!sohd
               
                tt = rs!NgayPH
                xlsheet.Cells(i, 6).NumberFormat = "dd/mm/yyyy"
                xlsheet.Cells(i, 6) = tt 'rs!NgayPH
                
                xlsheet.Cells(i, 7) = MulticharToUnicode(UVowels, rs!Ten)
                xlsheet.Cells(i, 8) = rs!mst
                xlsheet.Cells(i, 9) = MulticharToUnicode(UVowels, rs!MatHang)
                  xlsheet.Cells(i, 10).NumberFormat = "##,###,###,###"
                xlsheet.Cells(i, 10) = Int(rs!ThanhTien)
                xlsheet.Cells(i, 11) = Int(rs!TyLe)
                 xlsheet.Cells(i, 12).NumberFormat = "##,###,###,###"
                xlsheet.Cells(i, 12) = rs!sops
                tongcong = tongcong + rs!ThanhTien
                rs.MoveNext
                i = i + 1
                stt = stt + 1
            Loop
                'xlsheet.Cells(i + 1, 12) = tongcong
           ' xlsheet.Cells(i + 1, 9) = tongcong
            sodong = sodong + 1
'            xlsheet.Range("B" + sodong, "L" + sodong).Delete
'            i = i - 1
    End If
    i = i + 3

rs = Null
KT:
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
 
    Shell "EXPLORER.EXE " & pDataPath
    
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
  
KetThuc:
GauGe.Value = 1
InVATDauvao_exel = False
End Function






























