VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmDsCT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Li÷t k™ ch¯ng tı"
   ClientHeight    =   5205
   ClientLeft      =   5310
   ClientTop       =   1620
   ClientWidth     =   9075
   ClipControls    =   0   'False
   Icon            =   "Frmdsct.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5205
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Voucher List"
   Begin VB.CommandButton CmdPhieu 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&2 In Phi’u NX"
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
      Left            =   9240
      TabIndex        =   75
      Tag             =   "0"
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "phieuchi"
      Height          =   255
      Left            =   9120
      TabIndex        =   74
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "phieu thu"
      Height          =   255
      Left            =   9240
      TabIndex        =   73
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton CmdPhieu 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&1 In thu chi"
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
      Left            =   9120
      TabIndex        =   72
      Tag             =   "0"
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ng©n hµng"
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
      TabIndex        =   71
      Tag             =   "Export"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton CmdTK 
      Height          =   375
      Index           =   6
      Left            =   5280
      Picture         =   "Frmdsct.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   1875
      Width           =   375
   End
   Begin VB.TextBox txtShTk 
      Height          =   315
      Index           =   6
      Left            =   3840
      LinkItem        =   "SË hi÷u vÀt t≠ ho∆c ph©n loπi vÀt t≠ c«n theo d‚i nhÀp xu t"
      MaxLength       =   20
      TabIndex        =   65
      Tag             =   "0"
      Top             =   1875
      Width           =   1335
   End
   Begin VB.CheckBox ChkTaikhoan 
      BackColor       =   &H00FFFFC0&
      Caption         =   "C´ng tr◊nh, s∂n ph»m"
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
      Left            =   2040
      TabIndex        =   64
      Tag             =   "Liability Code"
      Top             =   1875
      Width           =   1815
   End
   Begin VB.CheckBox ChkTaikhoan 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ch≠a thanh to∏n Æ›ch danh"
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
      Left            =   8760
      TabIndex        =   63
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vµo &sÊ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Xem b∏o c∏o"
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "B∏o gi∏"
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
      Left            =   8160
      TabIndex        =   7
      Tag             =   "Quotation"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox CboN 
      Height          =   330
      Index           =   3
      ItemData        =   "Frmdsct.frx":5C5C
      Left            =   9240
      List            =   "Frmdsct.frx":5C84
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   5280
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CheckBox ChkTaikhoan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ng≠Íi nhÀp"
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
      Left            =   9120
      TabIndex        =   51
      Tag             =   "User Name"
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtShTk 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   3840
      LinkItem        =   "SË hi÷u vÀt t≠ ho∆c ph©n loπi vÀt t≠ c«n theo d‚i nhÀp xu t"
      MaxLength       =   20
      TabIndex        =   19
      Tag             =   "0"
      Text            =   "0"
      Top             =   2595
      Width           =   1305
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   4
      Left            =   7800
      Picture         =   "Frmdsct.frx":5CAF
      Style           =   1  'Graphical
      TabIndex        =   61
      Tag             =   "&Print"
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   3
      Left            =   7680
      Picture         =   "Frmdsct.frx":7111
      Style           =   1  'Graphical
      TabIndex        =   60
      Tag             =   "&View"
      Top             =   4720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox CboN 
      Height          =   330
      Index           =   0
      ItemData        =   "Frmdsct.frx":8283
      Left            =   3840
      List            =   "Frmdsct.frx":82AB
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   2955
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox CboN 
      Height          =   330
      Index           =   1
      ItemData        =   "Frmdsct.frx":82D6
      Left            =   3840
      List            =   "Frmdsct.frx":82FE
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   3315
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   4
      Left            =   3840
      LinkItem        =   "SË hi÷u vÀt t≠ ho∆c ph©n loπi vÀt t≠ c«n theo d‚i nhÀp xu t"
      MaxLength       =   20
      TabIndex        =   30
      Tag             =   "0"
      Top             =   2235
      Width           =   1335
   End
   Begin VB.CheckBox ChkTaikhoan 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CT, Ho∏ Æ¨n sË"
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
      Left            =   2040
      TabIndex        =   29
      Tag             =   "V. Code"
      Top             =   2235
      Width           =   1695
   End
   Begin VB.ComboBox CboN 
      Height          =   330
      Index           =   2
      ItemData        =   "Frmdsct.frx":8329
      Left            =   3840
      List            =   "Frmdsct.frx":8351
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   3675
      Width           =   1815
   End
   Begin VB.CheckBox ChkTaikhoan 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ch¯ng tı ghi sÊ"
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
      Left            =   2040
      TabIndex        =   49
      Tag             =   "V. Book"
      Top             =   3675
      Width           =   1575
   End
   Begin VB.CheckBox ChkTaikhoan 
      BackColor       =   &H00FFFFC0&
      Caption         =   "C´ng nÓ li™n quan"
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
      Left            =   2040
      TabIndex        =   26
      Tag             =   "Liability Code"
      Top             =   1515
      Width           =   1695
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   3
      Left            =   3840
      LinkItem        =   "SË hi÷u vÀt t≠ ho∆c ph©n loπi vÀt t≠ c«n theo d‚i nhÀp xu t"
      MaxLength       =   20
      TabIndex        =   27
      Tag             =   "0"
      Top             =   1515
      Width           =   1335
   End
   Begin VB.CommandButton CmdTK 
      Height          =   375
      Index           =   3
      Left            =   5280
      Picture         =   "Frmdsct.frx":837C
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1470
      Width           =   375
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   975
      Left            =   7320
      TabIndex        =   43
      Top             =   3120
      Width           =   1575
      Begin VB.OptionButton SSOpt 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ngµy GS, SËhi÷u"
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
         Left            =   0
         TabIndex        =   46
         Tag             =   "B. Date, V. Code"
         Top             =   0
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton SSOpt 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ngµy CT, SËhi÷u"
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
         Left            =   0
         TabIndex        =   45
         Tag             =   "V. Date, V. Code"
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton SSOpt 
         BackColor       =   &H00FFFFC0&
         Caption         =   "SË hi÷u CT"
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
         Left            =   0
         TabIndex        =   44
         Tag             =   "V. Code"
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.OptionButton OptLK 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tı ngµy"
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
      Left            =   4680
      TabIndex        =   3
      Tag             =   "From"
      Top             =   4800
      Width           =   975
   End
   Begin VB.OptionButton OptLK 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tı th∏ng"
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
      Left            =   4680
      TabIndex        =   0
      Tag             =   "From"
      Top             =   4320
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "B∏n hµng"
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
      TabIndex        =   10
      Tag             =   "Sale"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tr›ch kh u hao"
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
      Left            =   120
      TabIndex        =   15
      Tag             =   "Depreciation"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ß∏nh gi∏ lπi"
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
      TabIndex        =   14
      Tag             =   "Revaluation"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Gi∂m TSCß"
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
      TabIndex        =   13
      Tag             =   "Dec. Assets"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "T®ng TSCß"
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
      TabIndex        =   12
      Tag             =   "Inc. Assets"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton CmdTK 
      Height          =   375
      Index           =   2
      Left            =   5280
      Picture         =   "Frmdsct.frx":87F6
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1110
      Width           =   375
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   2
      Left            =   3840
      LinkItem        =   "SË hi÷u vÀt t≠ ho∆c ph©n loπi vÀt t≠ c«n theo d‚i nhÀp xu t"
      MaxLength       =   20
      TabIndex        =   24
      Tag             =   "0"
      Top             =   1155
      Width           =   1335
   End
   Begin VB.CheckBox ChkTaikhoan 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tµi s∂n ph∏t sinh"
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
      Left            =   2040
      TabIndex        =   23
      Tag             =   "Assets Code"
      Top             =   1155
      Width           =   1695
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   2
      Left            =   2040
      Picture         =   "Frmdsct.frx":8C70
      Style           =   1  'Graphical
      TabIndex        =   40
      Tag             =   "&Return"
      Top             =   4605
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   1
      Left            =   8640
      Picture         =   "Frmdsct.frx":A092
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "&Select"
      Top             =   4720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   0
      Left            =   3360
      Picture         =   "Frmdsct.frx":B4F4
      Style           =   1  'Graphical
      TabIndex        =   38
      Tag             =   "&List"
      Top             =   4605
      Width           =   1095
   End
   Begin VB.CommandButton CmdTK 
      Height          =   375
      Index           =   1
      Left            =   5280
      Picture         =   "Frmdsct.frx":C8BE
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   735
      Width           =   375
   End
   Begin VB.CommandButton CmdTK 
      Height          =   375
      Index           =   0
      Left            =   5280
      Picture         =   "Frmdsct.frx":CD38
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   375
      Width           =   375
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   0
      Left            =   3840
      LinkItem        =   "SË hi÷u tµi kho∂n ho∆c chi ti’t c«n theo d‚i ph∏t sinh"
      MaxLength       =   20
      TabIndex        =   17
      Tag             =   "0"
      Top             =   435
      Width           =   1335
   End
   Begin VB.ComboBox CboThang 
      Height          =   330
      Index           =   1
      ItemData        =   "Frmdsct.frx":D1B2
      Left            =   7800
      List            =   "Frmdsct.frx":D1B4
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4290
      Width           =   1095
   End
   Begin VB.CheckBox ChkTaikhoan 
      BackColor       =   &H00FFFFC0&
      Caption         =   "VÀt t≠ ph∏t sinh"
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
      Left            =   2040
      TabIndex        =   20
      Tag             =   "Inventory Code"
      Top             =   795
      Width           =   1695
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   1
      Left            =   3840
      LinkItem        =   "SË hi÷u vÀt t≠ ho∆c ph©n loπi vÀt t≠ c«n theo d‚i nhÀp xu t"
      MaxLength       =   20
      TabIndex        =   21
      Tag             =   "0"
      Top             =   795
      Width           =   1335
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "K’t chuy”n"
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
      TabIndex        =   11
      Tag             =   "Conversion"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Xu t vt, gi∏ vËn"
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
      TabIndex        =   9
      Tag             =   "Export"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "NhÀp vÀt t≠"
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
      TabIndex        =   8
      Tag             =   "Import"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox ChkLoai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "TÊng hÓp"
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
      TabIndex        =   6
      Tag             =   "Common"
      Top             =   480
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox ChkTaikhoan 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tµi kho∂n ph∏t sinh"
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
      Left            =   2040
      TabIndex        =   16
      Tag             =   "Account"
      Top             =   435
      Width           =   1815
   End
   Begin VB.ComboBox CboThang 
      Height          =   330
      Index           =   0
      ItemData        =   "Frmdsct.frx":D1B6
      Left            =   5760
      List            =   "Frmdsct.frx":D1B8
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4290
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   0
      Left            =   5760
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   1
      Left            =   7800
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSGrid.Grid GrdChungtu 
      Height          =   3615
      Left            =   360
      TabIndex        =   68
      Tag             =   "1"
      Top             =   6000
      Width           =   9585
      _Version        =   65536
      _ExtentX        =   16916
      _ExtentY        =   6376
      _StockProps     =   77
      BackColor       =   16777215
      BorderStyle     =   0
      Rows            =   30
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   78
      Tag             =   "1"
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sæp theo th¯ t˘"
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
      Left            =   7440
      TabIndex        =   77
      Tag             =   "Sort by"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   4
      X1              =   0
      X2              =   0
      Y1              =   4200
      Y2              =   -480
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Th™m Æi“u ki÷n t◊m ki’m"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   2040
      TabIndex        =   76
      Tag             =   "Sort by"
      Top             =   80
      Width           =   2655
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   70
      Tag             =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   69
      Tag             =   "1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   67
      Tag             =   "1"
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SË ph∏t sinh"
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
      Left            =   2280
      TabIndex        =   62
      Tag             =   "Amount"
      Top             =   2595
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
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
      Left            =   6360
      TabIndex        =   58
      Tag             =   "Number of Voucher"
      Top             =   3360
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label LbSoCT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   1080
      TabIndex        =   57
      Top             =   7680
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   9000
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ph©n loπi"
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
      Left            =   2280
      TabIndex        =   56
      Tag             =   "Class"
      Top             =   3315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Kho hµng"
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
      Left            =   2280
      TabIndex        =   55
      Tag             =   "Store"
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TÊng Ph∏t sinh"
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
      Left            =   8280
      TabIndex        =   48
      Tag             =   "Amount"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   47
      Tag             =   "1"
      Top             =   1515
      Width           =   3135
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Æ’n ngµy"
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
      Left            =   6840
      TabIndex        =   42
      Tag             =   "to"
      Top             =   4800
      Width           =   975
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1800
      X2              =   1800
      Y1              =   4200
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   2
      X1              =   9000
      X2              =   9000
      Y1              =   4200
      Y2              =   0
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   41
      Tag             =   "1"
      Top             =   1155
      Width           =   3135
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Æ’n th∏ng"
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
      Index           =   6
      Left            =   6840
      TabIndex        =   37
      Tag             =   "to"
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "MÙc c«n t◊m"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   36
      Tag             =   "Sort by"
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00FFFFC0&
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
      Height          =   735
      Index           =   1
      Left            =   5760
      TabIndex        =   35
      Tag             =   "1"
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Di‘n gi∂i"
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
      Index           =   4
      Left            =   3480
      TabIndex        =   34
      Tag             =   "Description"
      Top             =   5640
      Width           =   4815
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ngµy GS"
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
      Index           =   3
      Left            =   2520
      TabIndex        =   33
      Tag             =   "B. Date"
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ngµy CT"
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
      Index           =   2
      Left            =   1560
      TabIndex        =   32
      Tag             =   "V. Date"
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SË CT"
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
      Index           =   1
      Left            =   240
      TabIndex        =   31
      Tag             =   "V. Code"
      Top             =   5640
      Width           =   1335
   End
End
Attribute VB_Name = "FrmDsCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MaCTChon As Long
Dim LietKe As Boolean
Dim Row As Integer
Dim ord As Integer
Dim ngay(0 To 1) As Date
Dim chuoidieukien_intoanbo As String

Private Sub ChkLoai_Click(Index As Integer)
    Dim vs As Boolean, i As Integer
    
    If ChkLoai(7).Value = 1 Then
        For i = 0 To 3
            ChkLoai(i).Value = 0
            ChkLoai(i).Enabled = False
        Next
        For i = 8 To 12
            ChkLoai(i).Value = 0
            ChkLoai(i).Enabled = False
        Next
    Else
        For i = 0 To 3
            ChkLoai(i).Enabled = True
        Next
        For i = 8 To 12
            ChkLoai(i).Enabled = True
        Next
    End If
    
    Command(5).Enabled = (ChkLoai(7).Value = 0)
    vs = (ChkLoai(1).Value = 1 Or ChkLoai(2).Value = 1 Or ChkLoai(8).Value = 1)
    Label(8).Visible = vs
    Label(9).Visible = vs
    CboN(0).Visible = vs
    CboN(1).Visible = vs
    
    KiemTraUser
End Sub

Private Sub CmdPhieu_Click(Index As Integer)
FrmChungtu.CmdPhieu_Click (Index)
End Sub

'======================================================================================
' Li÷t k™, Ch‰n ch¯ng tı
'======================================================================================
Public Sub Command_Click(Index As Integer)

If OptLK(0).Value = False Then
  If IsDate(MedNgay(0).Text) And IsDate(MedNgay(1).Text) Then
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
    
    
    Select Case Index
        Case 0:
            If ChkTaikhoan(0).Value = 1 And txtShTk(0).tag = 0 Then
                ErrMsg er_SHTaiKhoan1
                RFocus txtShTk(0)
                Exit Sub
            End If
            If ChkTaikhoan(1).Value = 1 And txtShTk(1).tag = 0 Then
                ErrMsg er_SHVattu
                RFocus txtShTk(1)
                Exit Sub
            End If
            If ChkTaikhoan(2).Value = 1 And txtShTk(2).tag = 0 Then
                ErrMsg er_SHTaiSan
                RFocus txtShTk(2)
                Exit Sub
            End If
            If ChkTaikhoan(3).Value = 1 And txtShTk(3).tag = 0 Then
                ErrMsg er_SHKhachHang
                RFocus txtShTk(3)
                Exit Sub
            End If
            If ChkTaikhoan(5).Value = 1 And Len(txtShTk(4).Text) = 0 Then
                MsgBox "Thi’u sË ho∏ Æ¨n !", vbExclamation, App.ProductName
                RFocus txtShTk(4)
                Exit Sub
            End If
            If CboThang(1).ListIndex < CboThang(0).ListIndex Then CboThang(1).ListIndex = CboThang(0).ListIndex
            LietKeChungtu IIf(ChkTaikhoan(0).Value = 1, txtShTk(0).Text, ""), IIf(ChkTaikhoan(1).Value = 1, txtShTk(1).tag, 0), IIf(ChkTaikhoan(2).Value = 1, txtShTk(2).tag, 0), IIf(ChkTaikhoan(3).Value = 1, txtShTk(3).tag, 0), IIf(ChkTaikhoan(5).Value = 1, txtShTk(4).Text, "")
'             GrdChungtu.col = 5
'            If Len(GrdChungtu.Text) = 0 Then Exit Sub
'            MaCTChon = CLng5(GrdChungtu.Text)
'            LietKe = True
'           Unload Me

        Case 1:
            GrdChungtu.col = 5
            If Len(GrdChungtu.Text) = 0 Then Exit Sub
            MaCTChon = CLng5(GrdChungtu.Text)
            Hide
            LietKe = True
        Case 2:
            MaCTChon = 0
            Hide
            LietKe = True
        Case 3, 4:            DSCTu Index - 3
        Case 5:
            Dim f As Form, ms As Long, i As Integer
            Set f = New FrmChungtu
            Load f
            With GrdChungtu
                .col = 5
                For i = .SelStartRow To .SelEndRow
                    .Row = i
                    ms = CLng5(.Text)
                    If ms > 0 Then
                        pPhieu = 0
                        f.VaoSoNK CLng5(GrdChungtu.Text)
                    End If
                Next
            End With
            Set f = Nothing
            pPhieu = 1
            Command_Click 0
    End Select
    If GrdChungtu.Rows > 1 Then
  GrdChungtu.RemoveItem (0)
  End If
End Sub
'======================================================================================
' Hi÷n cˆa sÊ ch‰n tµi kho∂n
'======================================================================================
Private Sub cmdtk_Click(Index As Integer)
    Me.MousePointer = 11
    Select Case Index
        Case 0:
            txtShTk(0).Text = FrmTaikhoan.ChonTk(txtShTk(0).Text)
        Case 1:
            txtShTk(1).Text = FrmVattu.ChonVattu(txtShTk(1).Text)
        Case 2:
            txtShTk(2).Text = frmDSTaiSan.ChonTaiSan(txtShTk(2).Text, CboThang(0).ItemData(CboThang(0).ListIndex), CboThang(0).ItemData(CboThang(1).ListIndex))
        Case 3:
            txtShTk(3).Text = FrmKhachHang.ChonKhachHang(txtShTk(3).Text)
        Case 6:
            txtShTk(6).Text = FrmTP.ChonTP(txtShTk(6).Text)
    End Select
    Me.MousePointer = 0
    RFocus txtShTk(Index)
End Sub

Private Sub Command1_Click()
'in toan bo phieu thu
Dim tungay As Date
Dim denngay As Date
tungay = "01/" + CboThang(0).Text
denngay = "1/" + CboThang(1).Text
If OptLK(0).Value = True Then
GetDate2.MedNgay(0) = tungay
GetDate2.MedNgay(1) = denngay
Else
GetDate2.MedNgay(0) = MedNgay(0).Text
GetDate2.MedNgay(1) = MedNgay(1).Text
End If
'RFocus GetDate2.MedNgay(1)
'GetDate2.MedNgay_KeyPress 1, 13
FrmChungtu.mnDD_Click (100)
End Sub

Private Sub Command2_Click()
'in toan bo phieu chi
Dim tungay As Date
Dim denngay As Date
tungay = "01/" + CboThang(0).Text
denngay = "1/" + CboThang(1).Text
If OptLK(0).Value = True Then
GetDate2.MedNgay(0) = tungay
GetDate2.MedNgay(1) = denngay
Else
GetDate2.MedNgay(0) = MedNgay(0).Text
GetDate2.MedNgay(1) = MedNgay(1).Text
End If
FrmChungtu.mnDD_Click (101)
End Sub

Private Sub Form_Activate()
    If LietKe Then
        'Me.Refresh
       ' Command_Click 0
        LietKe = False
    End If
    KiemTraUser
     Dim ngaytt As Date
 ngaytt = "01/01/" + Mid(str(pNamTC), 4, 5)
 MedNgay(0).Text = ngaytt
 MedNgay(1).Text = ngaytt

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
            Select Case KeyCode
                Case vbKeyL:
                    RFocus Command(0)
                    Command_Click 0
                Case vbKeyC:
                    RFocus Command(1)
                    Command_Click 1
                Case vbKeyV:
                    RFocus Command(2)
                    Command_Click 2
            End Select
    End If
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyA Then GridSelAll GrdChungtu
    If KeyCode = vbKeyEscape Then Command_Click 2
End Sub
'======================================================================================
' KhÎi tπo cˆa sÊ
'======================================================================================
Private Sub Form_Load()
    Dim chi_so As Integer
    
    If pPhieu > 0 Then Me.Caption = "Danh s∏ch c∏c phi’u ch≠a vµo sÊ"
    Command(5).Visible = (pPhieu > 0)
    
    ChkLoai(7).Enabled = pBaoGia > 0 And ((frmMain.Command(4).Visible And pPhieu = 1) Or (Not frmMain.Command(4).Visible And pPhieu = 0))
    
    ColumnSetUp GrdChungtu, 0, 1300, 2
    ColumnSetUp GrdChungtu, 1, 940, 2
    ColumnSetUp GrdChungtu, 2, 940, 2
    ColumnSetUp GrdChungtu, 3, 4780, 0
    ColumnSetUp GrdChungtu, 4, 1300, 1
    ColumnSetUp GrdChungtu, 5, 1, 0
    
    AddMonthToCbo CboThang(0)
    AddMonthToCbo CboThang(1)
    
    For chi_so = 0 To 1
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next

    MaCTChon = 0

    GrdChungtu.Row = 0
    GrdChungtu_Click
    LietKe = False

    Caption = Caption + " - " + CStr(pNamTC)
    ord = 0
        
    Int_RecsetToCbo "SELECT MaSo As F2,SoHieu As F1 FROM CTGhiSo ORDER BY SoHieu", CboN(2)
    SetSQL "MienTru", "SELECT MaSo As F2,SoHieu + ' - ' + DienGiai As F1 FROM NguonNhapXuat ORDER BY SoHieu  UNION SELECT 0,'" + ABCtoVNI(" Toµn bÈ") + "' FROM License"
    Int_RecsetToCbo "SELECT * FROM MienTru ORDER BY F1", CboN(1)
    SetSQL "MienTru", "SELECT MaSo As F2,TenKho As F1 FROM KhoHang ORDER BY TenKho  UNION SELECT 0,'" + ABCtoVNI(" Toµn bÈ") + "' FROM License"
    Int_RecsetToCbo "SELECT * FROM MienTru ORDER BY F1", CboN(0)
    Int_RecsetToCbo "SELECT MaSo As F2, TenNSD As F1 FROM Users ORDER BY TenNSD", CboN(3)
    
    ChkTaikhoan(4).Visible = (pSoKT Mod 100 >= 10)
    CboN(2).Visible = (pSoKT Mod 100 >= 10)
    
    ChkTaikhoan(8).Enabled = (pDTTP <> 0)
    txtShTk(6).Enabled = (pDTTP <> 0)
    LbTenTk(6).Enabled = (pDTTP <> 0)
    
    SetFont Me
LoiNgay:
End Sub

'======================================================================================
' NhÌ m∑ sË CT Æ≠Óc ch‰n
'======================================================================================
Private Sub GrdChungtu_Click()
 'SendKeys "{Home}", True
    Row = GrdChungtu.Row
    SetGridIndex GrdChungtu, Row
    With GrdChungtu
        .col = 5
        If Len(.Text) = 0 Then
        MaCTChon = 0
        Else
        MaCTChon = CLng5(.Text)
        'MaCTChon = CLng5(GrdChungtu.Text)
         FrmChungtu.MaSoCT = MaCTChon
         FrmChungtu.HienPhieuTrenManHinh (0)
         If (FrmChungtu.CmdPhieu(0).Visible = True) Then
         CmdPhieu(0).Visible = True
         Else
         CmdPhieu(0).Visible = False
         End If
         
         If (FrmChungtu.CmdPhieu(1).Visible = True) Then
         CmdPhieu(1).Visible = True
         CmdPhieu(1).Caption = FrmChungtu.CmdPhieu(1).Caption
         Else
         CmdPhieu(1).Visible = False
         End If
        End If
        .col = 0
        
         
    End With
'    SendKeys "{Home}", True
'    Row = GrdChungtu.Row
'    SetGridIndex GrdChungtu, Row
'    With GrdChungtu
'        .col = 5
'        If Len(.Text) = 0 Then
'        MaCTChon = 0
'        CmdPhieu(0).Visible = False
'        Else
'
'        MaCTChon = CLng5(.Text)
'         FrmChungtu.MaSoCT = MaCTChon
'         FrmChungtu.HienPhieuTrenManHinh (0)
'         If (FrmChungtu.CmdPhieu(0).Visible = True) Then
'          CmdPhieu(0).Visible = True
'          Else
'          CmdPhieu(0).Visible = False
'        End If
'
'         If (FrmChungtu.CmdPhieu(1).Visible = True) Then
'         CmdPhieu(1).Visible = True
'         CmdPhieu(1).Caption = FrmChungtu.CmdPhieu(1).Caption
'         Else
'         CmdPhieu(1).Visible = False
'         End If
'
'        .col = 0
'    End With
End Sub

Private Sub GrdChungtu_KeyPress(KeyAscii As Integer)
    GrdChungtu_Click
    If KeyAscii = 13 Then Command_Click 1
End Sub

Private Sub GrdChungtu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdChungtu, GrdChungtu.col
    End If
End Sub

Private Sub MedNgay_Change(Index As Integer)
 
  If IsDate(CboThang(0).Text) And IsDate(CboThang(1).Text) Then
        If CDate(CboThang(1).Text) < CDate(CboThang(0).Text) Then
            CboThang(1).Text = CboThang(0).Text
          Exit Sub
        End If
  End If
End Sub

Private Sub OptLK_Click(Index As Integer)
Dim ngay As Date
ngay = "01/01/" + Mid(str(pNamTC), 4, 5)
 MedNgay(0).Text = ngay
 MedNgay(1).Text = ngay
 End Sub

Private Sub SSOpt_Click(Index As Integer)
    ord = Index
End Sub

Private Sub txtShTk_GotFocus(Index As Integer)
    AutoSelect txtShTk(Index)
End Sub

Private Sub txtShTk_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 5 Then KeyProcess txtShTk(Index), KeyAscii
End Sub
'======================================================================================
' Ki”m tra sË hi÷u tµi kho∂n
'======================================================================================
Private Sub txtShTk_LostFocus(Index As Integer)
    Dim mtk As Long, st As String, id As Long
    
    mtk = 0
    Select Case Index
        Case 0:
             LbTenTk(0).Caption = tentk(txtShTk(0).Text, mtk)
            id = GetTK_ID("", mtk)
            ChkTaikhoan(7).Visible = (id = TKCNKH_ID Or id = TKCNPT_ID)
        Case 1:
            LbTenTk(1).Caption = TenVT(txtShTk(1).Text, mtk, st)
''        Case 2:
        Case 3:
            LbTenTk(3).Caption = TenKH(txtShTk(3).Text, mtk)
        Case 5:
            txtShTk(5).Text = Format(txtShTk(5).Text, Mask_0)
        Case 6:
            LbTenTk(6).Caption = TenTP(txtShTk(6).Text, mtk)
    End Select
    txtShTk(Index).tag = mtk
End Sub
'======================================================================================
' Li÷t k™ ch¯ng tı
'======================================================================================



Public Sub LietKeChungtu(shtk As String, mvt As Long, mts As Long, mcn As Long, shd As String)
    Dim sql As String, loaict As String, i As Integer, mct As Long, uid As Long, mct1 As Long, mloai As Integer
    Dim rs_chungtu As Recordset, st As Double, ovr As Integer, sh As String
    
    Me.MousePointer = 11
    st = Cdbl5(txtShTk(5).Text)
    sh = IIf(ChkLoai(7).Value = 1 Or pPhieu > 0, "P", "")
   ' sql = "SELECT MaCT,Sum(IIF(MaTKCo>0,SoPS,0)) AS TPS FROM ChungTu" + sh + " WHERE "
    sql = "SELECT MaCT,Sum(IIF(MaTKCo>0,SoPS,0)) AS TPS,sum(TLCK+val(iif(isnull(phantramchietkhau),0,phantramchietkhau))) as tylechietkhau,sum(CK+val(iif(isnull(sotienchietkhau) ,0,sotienchietkhau))) as chietkhau FROM ChungTu" + sh + " WHERE "
    
    If OptLK(0).Value Then
        sql = sql + WThang("ThangCT", CboThang(0).ItemData(CboThang(0).ListIndex), CboThang(1).ItemData(CboThang(1).ListIndex)) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " GROUP BY MaCT"
    Else
        sql = sql + WNgay("NgayGS", ngay(0), ngay(1)) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " GROUP BY MaCT"
    End If
    SetSQL "MienTru", sql
    
    sql = "SELECT DISTINCTROW ChungTu" + sh + ".MaCT, ChungTu" + sh + ".SoHieu, NgayCT, NgayGS, ChungTu" + sh + ".DienGiai" + IIf(pNN = 1, "E", "") + " AS DG,TPS,tylechietkhau,chietkhau,User_ID,ChungTu" + sh + ".MaLoai FROM ((((ChungTu" + sh + " INNER JOIN MienTru ON ChungTu" + sh + ".MaCT=MienTru.MaCT) LEFT JOIN HoaDon" + sh + " ON ChungTu" + sh + ".MaSo=HoaDon" + sh + ".MaSo) LEFT JOIN CTTaiSan ON ChungTu" + sh + ".MaCT = CTTaiSan.MaCTKT) LEFT JOIN HeThongTK ON ChungTu" + sh + ".MaTKNo = HeThongTK.MaSo) LEFT JOIN HeThongTK AS HeThongTK_1 ON ChungTu" + sh + ".MaTKCo = HeThongTK_1.MaSo WHERE (TRUE)"
    
    If Len(shd) > 0 Then sql = sql + " AND (HoaDon" + sh + ".SoHD LIKE '" + shd + "*' OR ChungTu" + sh + ".SoHieu LIKE '" + shd + "*')"
    If mvt > 0 Then sql = sql + " AND (ChungTu" + sh + ".MaLoai = 1 OR ChungTu" + sh + ".MaLoai = 2 OR ChungTu" + sh + ".MaLoai = 8) AND (MaVattu = " + CStr(mvt) + ")"
    If mts > 0 Then sql = sql + " AND (CTTaiSan.MaTS = " + CStr(mts) + ")"
    If mcn > 0 Then sql = sql + " AND (MaKH = " + CStr(mcn) + " OR MaKhachHang = " + CStr(mcn) + " OR MaKHC=" + CStr(mcn) + ")"
    If Len(shtk) > 0 Then
        sql = sql + " AND (HethongTK.SoHieu LIKE '" + shtk + "*' OR HethongTK_1.SoHieu LIKE '" + shtk + "*')" + IIf(st <> 0, " AND SoPS=" + DoiDau(st), "")
        If ChkTaikhoan(7).Value = 1 Then sql = sql + " AND ((HethongTK.TK_ID=" + CStr(TKCNPT_ID) + " OR HethongTK_1.TK_ID=" + CStr(TKCNKH_ID) + ") AND CT_ID=0)"
    End If
        
    If ChkTaikhoan(4).Value = 1 And CboN(2).ListIndex >= 0 Then
        sql = sql + " AND ChungTu" + sh + ".CTGS=" + CStr(CboN(2).ItemData(CboN(2).ListIndex))
    End If
    If ChkTaikhoan(8).Value = 1 And txtShTk(6).tag > 0 Then
        sql = sql + " AND ChungTu" + sh + ".MaTP=" + CStr(txtShTk(6).tag)
    End If
    If ChkTaikhoan(6).Value = 1 Then sql = sql + " AND (ChungTu" + sh + ".User_ID=" + CStr(CboN(3).ItemData(CboN(3).ListIndex)) + ")"
    
    loaict = ""
    For i = 0 To 12
        'If i < 4 Or i > 6 Then
        'loai = 4 loai ct moi them
        If i <= 4 Or i > 6 Then
            If ChkLoai(i).Value = 1 Then loaict = loaict + "ChungTu" + sh + ".MaLoai = " + CStr(i) + " OR "
        End If
    Next
    
    If Len(loaict) = 0 Then GoTo KT
    If Len(loaict) > 0 Then sql = sql + " AND (" + Left(loaict, Len(loaict) - 4) + ")"
        
    If (ChkLoai(1).Value = 1 Or ChkLoai(2).Value = 1 Or ChkLoai(8).Value = 1) And CboN(0).ListIndex > 0 Then sql = sql + " AND (ChungTu" + sh + ".MaKho=" + CStr(CboN(0).ItemData(CboN(0).ListIndex)) + ")"
    If (ChkLoai(1).Value = 1 Or ChkLoai(2).Value = 1 Or ChkLoai(8).Value = 1) And CboN(1).ListIndex > 0 Then sql = sql + " AND (ChungTu" + sh + ".MaNguon=" + CStr(CboN(1).ItemData(CboN(1).ListIndex)) + ")"
    
    If pProcessMode = 1 Then sql = sql + " AND XuLy<2 "
    
    sql = sql + " GROUP BY ChungTu" + sh + ".MaCT,User_ID,ChungTu" + sh + ".MaLoai,ChungTu" + sh + ".SoHieu, ChungTu" + sh + ".NgayCT, ChungTu" + sh + ".NgayGS, ChungTu" + sh + ".DienGiai" + IIf(pNN = 1, "E", "") + ",TPS,tylechietkhau,chietkhau "
        
    Select Case ord
        Case 0:
            sql = sql + "ORDER BY NgayGS asc ,val( ChungTu" + sh + ".SoHieu) asc "
        Case 1:
            sql = sql + "ORDER BY NgayCT asc, val(ChungTu" + sh + ".SoHieu) asc "
        Case 2:
            sql = sql + "ORDER BY  NgayCT asc,val(ChungTu" + sh + ".SoHieu) asc"
    End Select
   
    ClearGrid GrdChungtu, GrdChungtu.tag
    ClearGrid FrmChungtu.Grid2, FrmChungtu.Grid2.tag
    Dim so
    so = GrdChungtu.Rows - 1
If so > 0 Then
     Do Until so = 1
            GrdChungtu.RemoveItem (so)
           so = so - 1
      Loop
End If

'lay dieu kien de loc du lieu
Dim so_cong
so_cong = 0
chuoidieukien_intoanbo = sql

    Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs_chungtu.EOF
        If mct <> rs_chungtu!MaCT Then
            mct = rs_chungtu!MaCT
            uid = rs_chungtu!User_ID
        Else
            If uid <> rs_chungtu!User_ID Or mloai <> rs_chungtu!maloai Then
                mct1 = Lng_MaxValue("MaCT", "ChungTu" + sh) + 1
                ExecuteSQL5 "UPDATE ChungTu" + sh + " SET MaCT=" + CStr(mct1) + " WHERE MaCT=" + CStr(mct) + " AND User_ID=" + CStr(rs_chungtu!User_ID)
                mct = mct1
                uid = rs_chungtu!User_ID
                mloai = rs_chungtu!maloai
            End If
        End If
       If GrdChungtu.Rows < MaxGridRow Then
    ' MsgBox rs_chungtu!SoHieu
            GrdChungtu.AddItem rs_chungtu!sohieu + Chr(9) + Format(rs_chungtu!NgayCT, Mask_D) + Chr(9) _
                + Format(rs_chungtu!NgayGS, Mask_D) + Chr(9) + rs_chungtu!dg + Chr(9) + Format(rs_chungtu!tps, Mask_0) + Chr(9) + CStr(mct) ', 0
                    FrmChungtu.Grid2.AddItem rs_chungtu!sohieu + Chr(9) + Format(rs_chungtu!NgayCT, Mask_D) + Chr(9) _
                + Format(rs_chungtu!NgayGS, Mask_D) + Chr(9) + rs_chungtu!dg + Chr(9) + Format(rs_chungtu!tps, Mask_0) + Chr(9) + CStr(mct) + Chr(9) + Format(rs_chungtu!tylechietkhau, Mask_0) + Chr(9) + Format(rs_chungtu!chietkhau, Mask_0), 0
   
                so_cong = so_cong + 1
        Else
            ovr = 1
        End If
        rs_chungtu.MoveNext
    Loop
    FrmChungtu.Label(28).Caption = "TÊng sË ch¯ng tı: " + str(rs_chungtu.RecordCount)
        Dim kk
    kk = 0
    Do While kk <> 10
            GrdChungtu.AddItem "" + Chr(9) + "" + Chr(9) _
                + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" ', 0
                kk = kk + 1
    Loop
   If so_cong = 0 Then
        MsgBox "Kh´ng c„ ch¯ng tı ph∏t sinh"
   End If
    Do While so_cong < 15
                     FrmChungtu.Grid2.AddItem " " + Chr(9) + "" + Chr(9) _
                + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + ""
   
               so_cong = so_cong + 1
    Loop
    
    
    If rs_chungtu.RecordCount > 0 Then
        If Row >= rs_chungtu.RecordCount Then Row = rs_chungtu.RecordCount - 1
    Else
        Row = 0
    End If
  
'    GrdChungtu.Rows = IIf(rs_chungtu.RecordCount > GrdChungtu.tag, IIf(rs_chungtu.RecordCount > MaxGridRow, MaxGridRow, rs_chungtu.RecordCount), GrdChungtu.tag)' chuyen thanh maxrows
    LbSoCT.Caption = CStr(rs_chungtu.RecordCount)
    GrdChungtu.col = 0
    rs_chungtu.Close
    Set rs_chungtu = Nothing
    On Error Resume Next
    GrdChungtu.Row = Row
    'GrdChungtu_Click
    ' them
     SetGridIndex GrdChungtu, Row + 1
    With GrdChungtu
        .col = 5
        If Len(.Text) = 0 Then MaCTChon = 0 Else MaCTChon = CLng5(.Text)
        .col = 0
    End With
    
    If Not GrdChungtu.RowIsVisible(Row) Then GrdChungtu.TopRow = Row - 8
    On Error GoTo 0
    If ovr > 0 Then ErrMsg er_NhieuCT
     If so_cong <> 0 Then Hide
KT:
    Me.MousePointer = 0
End Sub


Public Sub LietKeChungtu_1(shtk As String, mvt As Long, mts As Long, mcn As Long, shd As String)
    Dim sql As String, loaict As String, i As Integer, mct As Long, uid As Long, mct1 As Long, mloai As Integer
    Dim rs_chungtu As Recordset, st As Double, ovr As Integer, sh As String
    
    Me.MousePointer = 11
'    st = Cdbl5(txtShTk(5).Text)
'    sh = IIf(ChkLoai(7).Value = 1 Or pPhieu > 0, "P", "")
'   ' sql = "SELECT MaCT,Sum(IIF(MaTKCo>0,SoPS,0)) AS TPS FROM ChungTu" + sh + " WHERE "
'    sql = "SELECT MaCT,Sum(IIF(MaTKCo>0,SoPS,0)) AS TPS,sum(TLCK+val(iif(isnull(phantramchietkhau),0,phantramchietkhau))) as tylechietkhau,sum(CK+val(iif(isnull(sotienchietkhau) ,0,sotienchietkhau))) as chietkhau FROM ChungTu" + sh + " WHERE "
'
'    If OptLK(0).Value Then
'        'sql = sql + WThang("ThangCT", CboThang(0).ItemData(CboThang(0).ListIndex), CboThang(1).ItemData(CboThang(1).ListIndex)) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " GROUP BY MaCT"
'        sql = sql + WThang("ThangCT", CboThang(0).ItemData(CboThang(0).ListIndex), CboThang(1).ItemData(CboThang(1).ListIndex)) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " GROUP BY MaCT"
'    Else
'        sql = sql + WNgay("NgayGS", ngay(0), ngay(1)) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " GROUP BY MaCT"
'       ' sql = sql + IIf(pProcessMode = 1, " AND XuLy<2", "") + " GROUP BY MaCT"
'    End If
'    SetSQL "MienTru", sql
'
'    sql = "SELECT DISTINCTROW  ChungTu" + sh + ".MaCT, ChungTu" + sh + ".SoHieu, NgayCT, NgayGS, ChungTu" + sh + ".DienGiai" + IIf(pNN = 1, "E", "") + " AS DG,TPS,tylechietkhau,chietkhau,User_ID,ChungTu" + sh + ".MaLoai FROM ((((ChungTu" + sh + " INNER JOIN MienTru ON ChungTu" + sh + ".MaCT=MienTru.MaCT) LEFT JOIN HoaDon" + sh + " ON ChungTu" + sh + ".MaSo=HoaDon" + sh + ".MaSo) LEFT JOIN CTTaiSan ON ChungTu" + sh + ".MaCT = CTTaiSan.MaCTKT) LEFT JOIN HeThongTK ON ChungTu" + sh + ".MaTKNo = HeThongTK.MaSo) LEFT JOIN HeThongTK AS HeThongTK_1 ON ChungTu" + sh + ".MaTKCo = HeThongTK_1.MaSo WHERE (TRUE)"
'
'    If Len(shd) > 0 Then sql = sql + " AND (HoaDon" + sh + ".SoHD LIKE '" + shd + "*' OR ChungTu" + sh + ".SoHieu LIKE '" + shd + "*')"
'    If mvt > 0 Then sql = sql + " AND (ChungTu" + sh + ".MaLoai = 1 OR ChungTu" + sh + ".MaLoai = 2 OR ChungTu" + sh + ".MaLoai = 8) AND (MaVattu = " + CStr(mvt) + ")"
'    If mts > 0 Then sql = sql + " AND (CTTaiSan.MaTS = " + CStr(mts) + ")"
'    If mcn > 0 Then sql = sql + " AND (MaKH = " + CStr(mcn) + " OR MaKhachHang = " + CStr(mcn) + " OR MaKHC=" + CStr(mcn) + ")"
'    If Len(shtk) > 0 Then
'        sql = sql + " AND (HethongTK.SoHieu LIKE '" + shtk + "*' OR HethongTK_1.SoHieu LIKE '" + shtk + "*')" + IIf(st <> 0, " AND SoPS=" + DoiDau(st), "")
'        If ChkTaikhoan(7).Value = 1 Then sql = sql + " AND ((HethongTK.TK_ID=" + CStr(TKCNPT_ID) + " OR HethongTK_1.TK_ID=" + CStr(TKCNKH_ID) + ") AND CT_ID=0)"
'    End If
'
'    If ChkTaikhoan(4).Value = 1 And CboN(2).ListIndex >= 0 Then
'        sql = sql + " AND ChungTu" + sh + ".CTGS=" + CStr(CboN(2).ItemData(CboN(2).ListIndex))
'    End If
'    If ChkTaikhoan(8).Value = 1 And txtShTk(6).tag > 0 Then
'        sql = sql + " AND ChungTu" + sh + ".MaTP=" + CStr(txtShTk(6).tag)
'    End If
'    If ChkTaikhoan(6).Value = 1 Then sql = sql + " AND (ChungTu" + sh + ".User_ID=" + CStr(CboN(3).ItemData(CboN(3).ListIndex)) + ")"
'
'    loaict = ""
'    For i = 0 To 12
'        'If i < 4 Or i > 6 Then
'        'loai = 4 loai ct moi them
'        If i <= 4 Or i > 6 Then
'            If ChkLoai(i).Value = 1 Then loaict = loaict + "ChungTu" + sh + ".MaLoai = " + CStr(i) + " OR "
'        End If
'    Next
'
'    If Len(loaict) = 0 Then GoTo KT
'    If Len(loaict) > 0 Then sql = sql + " AND (" + Left(loaict, Len(loaict) - 4) + ")"
'
'    If (ChkLoai(1).Value = 1 Or ChkLoai(2).Value = 1 Or ChkLoai(8).Value = 1) And CboN(0).ListIndex > 0 Then sql = sql + " AND (ChungTu" + sh + ".MaKho=" + CStr(CboN(0).ItemData(CboN(0).ListIndex)) + ")"
'    If (ChkLoai(1).Value = 1 Or ChkLoai(2).Value = 1 Or ChkLoai(8).Value = 1) And CboN(1).ListIndex > 0 Then sql = sql + " AND (ChungTu" + sh + ".MaNguon=" + CStr(CboN(1).ItemData(CboN(1).ListIndex)) + ")"
'
'    If pProcessMode = 1 Then sql = sql + " AND XuLy<2 "
'
'    sql = sql + " GROUP BY ChungTu" + sh + ".MaCT,User_ID,ChungTu" + sh + ".MaLoai,ChungTu" + sh + ".SoHieu, ChungTu" + sh + ".NgayCT, ChungTu" + sh + ".NgayGS, ChungTu" + sh + ".DienGiai" + IIf(pNN = 1, "E", "") + ",TPS,TPS,tylechietkhau,chietkhau "
        
    
'    Select Case ord
'        Case 0:
'            sql = sql + "ORDER BY NgayGS desc , ChungTu" + sh + ".SoHieu desc "
'        Case 1:
'            sql = sql + "ORDER BY NgayGS desc, ChungTu" + sh + ".SoHieu desc "
'        Case 2:
'            sql = sql + "ORDER BY NgayGS desc,ChungTu" + sh + ".SoHieu desc"
'    End Select
   
'
      st = Cdbl5(txtShTk(5).Text)
    sh = IIf(ChkLoai(7).Value = 1 Or pPhieu > 0, "P", "")
   ' sql = "SELECT MaCT,Sum(IIF(MaTKCo>0,SoPS,0)) AS TPS FROM ChungTu" + sh + " WHERE "
    sql = "SELECT MaCT,Sum(IIF(MaTKCo>0,SoPS,0)) AS TPS,sum(TLCK+val(iif(isnull(phantramchietkhau),0,phantramchietkhau))) as tylechietkhau,sum(CK+val(iif(isnull(sotienchietkhau) ,0,sotienchietkhau))) as chietkhau FROM ChungTu" + sh + " WHERE "
    
    If OptLK(0).Value Then
        sql = sql + WThang("ThangCT", CboThang(0).ItemData(CboThang(0).ListIndex), CboThang(1).ItemData(CboThang(1).ListIndex)) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " GROUP BY MaCT"
    Else
        sql = sql + WNgay("NgayGS", ngay(0), ngay(1)) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " GROUP BY MaCT"
    End If
    SetSQL "MienTru", sql
    
    sql = "SELECT DISTINCTROW ChungTu" + sh + ".MaCT, ChungTu" + sh + ".SoHieu, NgayCT, NgayGS, ChungTu" + sh + ".DienGiai" + IIf(pNN = 1, "E", "") + " AS DG,TPS,tylechietkhau,chietkhau,User_ID,ChungTu" + sh + ".MaLoai FROM ((((ChungTu" + sh + " INNER JOIN MienTru ON ChungTu" + sh + ".MaCT=MienTru.MaCT) LEFT JOIN HoaDon" + sh + " ON ChungTu" + sh + ".MaSo=HoaDon" + sh + ".MaSo) LEFT JOIN CTTaiSan ON ChungTu" + sh + ".MaCT = CTTaiSan.MaCTKT) LEFT JOIN HeThongTK ON ChungTu" + sh + ".MaTKNo = HeThongTK.MaSo) LEFT JOIN HeThongTK AS HeThongTK_1 ON ChungTu" + sh + ".MaTKCo = HeThongTK_1.MaSo WHERE (TRUE)"
    
    If Len(shd) > 0 Then sql = sql + " AND (HoaDon" + sh + ".SoHD LIKE '" + shd + "*' OR ChungTu" + sh + ".SoHieu LIKE '" + shd + "*')"
    If mvt > 0 Then sql = sql + " AND (ChungTu" + sh + ".MaLoai = 1 OR ChungTu" + sh + ".MaLoai = 2 OR ChungTu" + sh + ".MaLoai = 8) AND (MaVattu = " + CStr(mvt) + ")"
    If mts > 0 Then sql = sql + " AND (CTTaiSan.MaTS = " + CStr(mts) + ")"
    If mcn > 0 Then sql = sql + " AND (MaKH = " + CStr(mcn) + " OR MaKhachHang = " + CStr(mcn) + " OR MaKHC=" + CStr(mcn) + ")"
    If Len(shtk) > 0 Then
        sql = sql + " AND (HethongTK.SoHieu LIKE '" + shtk + "*' OR HethongTK_1.SoHieu LIKE '" + shtk + "*')" + IIf(st <> 0, " AND SoPS=" + DoiDau(st), "")
        If ChkTaikhoan(7).Value = 1 Then sql = sql + " AND ((HethongTK.TK_ID=" + CStr(TKCNPT_ID) + " OR HethongTK_1.TK_ID=" + CStr(TKCNKH_ID) + ") AND CT_ID=0)"
    End If
        
    If ChkTaikhoan(4).Value = 1 And CboN(2).ListIndex >= 0 Then
        sql = sql + " AND ChungTu" + sh + ".CTGS=" + CStr(CboN(2).ItemData(CboN(2).ListIndex))
    End If
    If ChkTaikhoan(8).Value = 1 And txtShTk(6).tag > 0 Then
        sql = sql + " AND ChungTu" + sh + ".MaTP=" + CStr(txtShTk(6).tag)
    End If
    If ChkTaikhoan(6).Value = 1 Then sql = sql + " AND (ChungTu" + sh + ".User_ID=" + CStr(CboN(3).ItemData(CboN(3).ListIndex)) + ")"
    
    loaict = ""
    For i = 0 To 12
        'If i < 4 Or i > 6 Then
        'loai = 4 loai ct moi them
        If i <= 4 Or i > 6 Then
            If ChkLoai(i).Value = 1 Then loaict = loaict + "ChungTu" + sh + ".MaLoai = " + CStr(i) + " OR "
        End If
    Next
    
    If Len(loaict) = 0 Then GoTo KT
    If Len(loaict) > 0 Then sql = sql + " AND (" + Left(loaict, Len(loaict) - 4) + ")"
        
    If (ChkLoai(1).Value = 1 Or ChkLoai(2).Value = 1 Or ChkLoai(8).Value = 1) And CboN(0).ListIndex > 0 Then sql = sql + " AND (ChungTu" + sh + ".MaKho=" + CStr(CboN(0).ItemData(CboN(0).ListIndex)) + ")"
    If (ChkLoai(1).Value = 1 Or ChkLoai(2).Value = 1 Or ChkLoai(8).Value = 1) And CboN(1).ListIndex > 0 Then sql = sql + " AND (ChungTu" + sh + ".MaNguon=" + CStr(CboN(1).ItemData(CboN(1).ListIndex)) + ")"
    
    If pProcessMode = 1 Then sql = sql + " AND XuLy<2 "
    
    sql = sql + " GROUP BY ChungTu" + sh + ".MaCT,User_ID,ChungTu" + sh + ".MaLoai,ChungTu" + sh + ".SoHieu, ChungTu" + sh + ".NgayCT, ChungTu" + sh + ".NgayGS, ChungTu" + sh + ".DienGiai" + IIf(pNN = 1, "E", "") + ",TPS,tylechietkhau,chietkhau "
        
    Select Case ord
        Case 0:
            sql = sql + "ORDER BY NgayGS asc ,val( ChungTu" + sh + ".SoHieu) asc "
        Case 1:
            sql = sql + "ORDER BY NgayCT asc, val(ChungTu" + sh + ".SoHieu) asc "
        Case 2:
            sql = sql + "ORDER BY  NgayCT asc,val(ChungTu" + sh + ".SoHieu) asc"
    End Select
   
    ClearGrid GrdChungtu, GrdChungtu.tag
    ClearGrid FrmChungtu.Grid2, FrmChungtu.Grid2.tag
    Dim so
    so = GrdChungtu.Rows - 1
If so > 0 Then
     Do Until so = 1
            GrdChungtu.RemoveItem (so)
           so = so - 1
      Loop
End If

'lay dieu kien de loc du lieu

chuoidieukien_intoanbo = sql
Dim so_cong As Integer
so_cong = 0
sql = " select a.* from (select top 500 tong.* from (" + sql + ") tong) a order by a.NgayGS asc,val(a.sohieu) asc"
    Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs_chungtu.EOF
        If mct <> rs_chungtu!MaCT Then
            mct = rs_chungtu!MaCT
            uid = rs_chungtu!User_ID
        Else
            If uid <> rs_chungtu!User_ID Or mloai <> rs_chungtu!maloai Then
                mct1 = Lng_MaxValue("MaCT", "ChungTu" + sh) + 1
                ExecuteSQL5 "UPDATE ChungTu" + sh + " SET MaCT=" + CStr(mct1) + " WHERE MaCT=" + CStr(mct) + " AND User_ID=" + CStr(rs_chungtu!User_ID)
                mct = mct1
                uid = rs_chungtu!User_ID
                mloai = rs_chungtu!maloai
            End If
        End If
    If GrdChungtu.Rows < MaxGridRow Then
    ' MsgBox rs_chungtu!SoHieu
    so_cong = so_cong + 1
   ' If so_cong <= 50 Then
            GrdChungtu.AddItem rs_chungtu!sohieu + Chr(9) + Format(rs_chungtu!NgayCT, Mask_D) + Chr(9) _
                + Format(rs_chungtu!NgayGS, Mask_D) + Chr(9) + rs_chungtu!dg + Chr(9) + Format(rs_chungtu!tps, Mask_0) + Chr(9) + CStr(mct), 0
                
                    FrmChungtu.Grid2.AddItem rs_chungtu!sohieu + Chr(9) + Format(rs_chungtu!NgayCT, Mask_D) + Chr(9) _
                + Format(rs_chungtu!NgayGS, Mask_D) + Chr(9) + rs_chungtu!dg + Chr(9) + Format(rs_chungtu!tps, Mask_0) + Chr(9) + CStr(mct) + Chr(9) + Format(rs_chungtu!tylechietkhau, Mask_0) + Chr(9) + Format(rs_chungtu!chietkhau, Mask_0), 0
   '
    '            End If
        Else
            ovr = 1
        End If
        rs_chungtu.MoveNext
    Loop
    FrmChungtu.Label(28).Caption = "TÊng sË ch¯ng tı: " + str(rs_chungtu.RecordCount)
    
    Dim kk
    kk = 0
    Do While so_cong < 15
'            FrmChungtu.Grid2.AddItem "" + Chr(9) + "" + Chr(9) _
'                + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" ', 0
'
                       FrmChungtu.Grid2.AddItem " " + Chr(9) + "" + Chr(9) _
                + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + ""
   
                so_cong = so_cong + 1
    Loop
    If rs_chungtu.RecordCount > 0 Then
        If Row >= rs_chungtu.RecordCount Then Row = rs_chungtu.RecordCount - 1
    Else
        Row = 0
    End If
  
'    GrdChungtu.Rows = IIf(rs_chungtu.RecordCount > GrdChungtu.tag, IIf(rs_chungtu.RecordCount > MaxGridRow, MaxGridRow, rs_chungtu.RecordCount), GrdChungtu.tag)
    LbSoCT.Caption = CStr(rs_chungtu.RecordCount)
    GrdChungtu.col = 0
    rs_chungtu.Close
    Set rs_chungtu = Nothing
    On Error Resume Next
    GrdChungtu.Row = Row
    'GrdChungtu_Click
    
     SetGridIndex GrdChungtu, Row + 1
    With GrdChungtu
        .col = 5
        If Len(.Text) = 0 Then MaCTChon = 0 Else MaCTChon = CLng5(.Text)
        .col = 0
    End With
    
    If Not GrdChungtu.RowIsVisible(Row) Then GrdChungtu.TopRow = Row - 8
    On Error GoTo 0
    If ovr > 0 Then ErrMsg er_NhieuCT
  
KT:
    Me.MousePointer = 0
End Sub

'======================================================================================
' Hµm hi÷n cˆa sÊ danh s∏ch vµ tr∂ v“ m∑ sË CT Æ≠Óc ch‰n
'======================================================================================
Public Function ChonCT(p As Integer) As Long
    Show 1
    p = IIf(ChkLoai(7).Value = 1 Or pPhieu > 0, 1, 0)
    ChonCT = MaCTChon
End Function

Private Sub KiemTraUser()
    ChkLoai(3).Enabled = (frmMain.tag Mod 10 >= 1)
    ChkLoai(0).Enabled = (frmMain.tag Mod 10 >= 1)
    
    If Not (frmMain.tag Mod 10 >= 1) Then
        ChkLoai(0).Value = 0
        
        ChkLoai(1).Enabled = (frmMain.tag Mod 100 >= 10)
        ChkLoai(2).Enabled = (frmMain.tag Mod 1000 >= 100)
    
        ChkLoai(9).Enabled = (frmMain.tag Mod 10000 >= 1000)
        ChkLoai(10).Enabled = (frmMain.tag Mod 10000 >= 1000)
        ChkLoai(11).Enabled = (frmMain.tag Mod 10000 >= 1000)
        ChkLoai(12).Enabled = (frmMain.tag Mod 10000 >= 1000)
        
        ChkLoai(7).Enabled = (frmMain.tag Mod 100000 >= 10000)
        ChkLoai(8).Enabled = (frmMain.tag Mod 100000 >= 10000)
    End If
End Sub

Private Sub MedNgay_GotFocus(Index As Integer)
    AutoSelect MedNgay(Index)
End Sub

Private Sub MedNgay_LostFocus(Index As Integer)
    If IsDate(MedNgay(Index).Text) Then
        ngay(Index) = CDate(MedNgay(Index).Text)
    Else
        RFocus MedNgay(Index)
    End If
    Dim n As Date
    If Index = 0 Then
    n = MedNgay(0).Text
     MedNgay(1).Text = n
     End If
End Sub

Private Sub DSCTu(dest As Integer)
    Dim i As Integer, sql As String
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    With GrdChungtu
        For i = 0 To .Rows - 1
            .Row = i
            .col = 5
            If Len(.Text) = 0 Then Exit For
            ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu) VALUES (" + CStr(.Text) + ",'" + .Text + "')"
        Next
    End With
    
    sql = "SELECT -1 AS LoaiCT,MaCT,ChungTu.SoHieu AS SHCT,NgayCT,DienGiai" + IIf(pNN = 1, "E", "") + ",Sum(SoPS) AS PS,Sum(SoPS2No) AS Luong,HethongTK.SoHieu AS SHTK,First(HethongTK.Ten) AS TenTK,IIF(TK_ID=" + CStr(TKVT_ID) + ",Vattu.SoHieu,'') AS SHVT, First(TenVattu) AS TenVT, IIF(TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ",KhachHang.SoHieu,'') AS SHKH,First(KhachHang.Ten) AS TenKH FROM (((" + ChungTu2TKNC(-1) + ") INNER JOIN BaoCaoCP2 ON ChungTu.MaCT=BaoCaoCP2.MaSo) LEFT JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) LEFT JOIN KhachHang ON ChungTu.MaKH=KhachHang.MaSo GROUP BY MaCT,ChungTu.SoHieu,NgayCT,DienGiai,HethongTK.SoHieu,IIF(TK_ID=" + CStr(TKVT_ID) + ",Vattu.SoHieu,''),IIF(TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ",KhachHang.SoHieu,'')" _
        & " UNION SELECT 1 AS LoaiCT,MaCT,ChungTu.SoHieu AS SHCT,NgayCT,DienGiai,Sum(SoPS) AS PS,Sum(SoPS2Co) AS Luong,HethongTK.SoHieu AS SHTK,First(HethongTK.Ten) AS TenTK,IIF(TK_ID=" + CStr(TKVT_ID) + ",Vattu.SoHieu,'') AS SHVT, First(TenVattu) AS TenVT, IIF(TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ",KhachHang.SoHieu,'') AS SHKH,First(KhachHang.Ten) AS TenKH FROM (((" + ChungTu2TKNC(1) + ") INNER JOIN BaoCaoCP2 ON ChungTu.MaCT=BaoCaoCP2.MaSo) LEFT JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) LEFT JOIN KhachHang ON ChungTu.MaKHC=KhachHang.MaSo GROUP BY MaCT,ChungTu.SoHieu,NgayCT,DienGiai,HethongTK.SoHieu,IIF(TK_ID=" + CStr(TKVT_ID) + ",Vattu.SoHieu,''),IIF(TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ",KhachHang.SoHieu,'')" ' them chung tu
    
    SetSQL "QNhatKy", sql
    
    SetRptInfo
    frmMain.Rpt.ReportFileName = "CHUNGTU3.RPT"
    
    InBaoCaoRPT
End Sub







