VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmChungtu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "NhËp chøng tõ"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   18180
   ClipControls    =   0   'False
   Icon            =   "Fchungtu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Voucher"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8685
   ScaleWidth      =   18180
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer timer1542 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11880
      Top             =   600
   End
   Begin VB.Timer timer154 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   11280
      Top             =   600
   End
   Begin VB.Timer timerImport 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12600
      Top             =   5640
   End
   Begin VB.TextBox txtNgaychungtu 
      Height          =   285
      Left            =   11040
      TabIndex        =   169
      Text            =   "Text2"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtNgayghiso 
      Height          =   285
      Left            =   11040
      TabIndex        =   168
      Text            =   "Text2"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
      Caption         =   "T¶i hãa ®¬n"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   395
      Index           =   5
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   167
      Tag             =   "35"
      ToolTipText     =   "Depreciation"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Timer timer3311 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9240
      Top             =   5880
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   10080
      Top             =   600
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9600
      Top             =   600
   End
   Begin VB.CommandButton btnOpenexe 
      Caption         =   "mt"
      Height          =   315
      Left            =   11760
      TabIndex        =   166
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4320
      Top             =   5880
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   11640
      TabIndex        =   165
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9360
      TabIndex        =   164
      Text            =   "Text1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   12240
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   12960
      Top             =   0
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B¶ng nhËp xuÊt tån"
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
      Index           =   14
      Left            =   120
      TabIndex        =   158
      Tag             =   "Fluctuation of inventories"
      Top             =   6820
      Width           =   1695
   End
   Begin VB.OptionButton OptVAT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B¶ng kª ho¸ ®¬n  ra"
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
      TabIndex        =   153
      Tag             =   "VAT Ouput Table"
      Top             =   5580
      Width           =   1815
   End
   Begin VB.OptionButton OptVAT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B¶ng kª ho¸ ®¬n vµo"
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
      TabIndex        =   152
      Tag             =   "VAT Input Table"
      Top             =   5245
      Width           =   1935
   End
   Begin VB.OptionButton OptVAT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tê khai thuÕ GTGT"
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
      TabIndex        =   151
      Tag             =   "Monthly VAT Declaration Form"
      Top             =   5910
      Width           =   1815
   End
   Begin VB.ComboBox CboThang 
      Height          =   315
      ItemData        =   "Fchungtu.frx":57E2
      Left            =   3120
      List            =   "Fchungtu.frx":57E4
      Style           =   2  'Dropdown List
      TabIndex        =   148
      Top             =   0
      Width           =   1095
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B¶ng c©n ®èi ph¸t sinh"
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
      TabIndex        =   146
      Tag             =   "Account Balance Report 2"
      Top             =   6560
      Width           =   2055
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
      Index           =   100
      Left            =   120
      TabIndex        =   145
      Tag             =   "Journal Ledger"
      Top             =   6230
      Width           =   1815
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sæ chi tiÕt c«ng nî "
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
      Index           =   36
      Left            =   120
      TabIndex        =   144
      Tag             =   "Detail Report of Payable and Receivable"
      Top             =   7840
      Width           =   1815
   End
   Begin VB.TextBox txtshkh 
      Height          =   315
      Index           =   0
      Left            =   1920
      LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
      MaxLength       =   20
      TabIndex        =   143
      Tag             =   "0"
      Top             =   7770
      Width           =   1095
   End
   Begin VB.CommandButton cmdkh 
      Height          =   255
      Index           =   0
      Left            =   3120
      Picture         =   "Fchungtu.frx":57E6
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   7800
      Width           =   255
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sæ chi tiÕt vËt t­"
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
      TabIndex        =   141
      Tag             =   "Inventory detail report"
      Top             =   7220
      Width           =   1815
   End
   Begin VB.TextBox txtShVT 
      Height          =   315
      Index           =   0
      Left            =   1920
      LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
      MaxLength       =   20
      TabIndex        =   147
      Tag             =   "0"
      Top             =   7140
      Width           =   1095
   End
   Begin VB.CommandButton cmdvt 
      Height          =   255
      Index           =   0
      Left            =   3120
      Picture         =   "Fchungtu.frx":5C60
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   7160
      Width           =   255
   End
   Begin VB.TextBox txtShTk 
      Height          =   315
      Index           =   0
      Left            =   1920
      LinkItem        =   "Sè hiÖu chi tiÕt cÇn xem"
      MaxLength       =   12
      TabIndex        =   159
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   7460
      Width           =   1095
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sæ chi tiÕt tµi kho¶n"
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
      TabIndex        =   139
      Tag             =   "Account Detail Report"
      Top             =   7530
      Width           =   1815
   End
   Begin VB.CommandButton cmdtk 
      Height          =   255
      Index           =   0
      Left            =   3120
      Picture         =   "Fchungtu.frx":60DA
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   7480
      Width           =   255
   End
   Begin VB.ComboBox CboThang1 
      Height          =   315
      Index           =   2
      ItemData        =   "Fchungtu.frx":6554
      Left            =   1920
      List            =   "Fchungtu.frx":6556
      Style           =   2  'Dropdown List
      TabIndex        =   137
      Top             =   8130
      Width           =   1095
   End
   Begin VB.ComboBox CboThang1 
      Height          =   315
      Index           =   1
      ItemData        =   "Fchungtu.frx":6558
      Left            =   480
      List            =   "Fchungtu.frx":655A
      Style           =   2  'Dropdown List
      TabIndex        =   136
      Top             =   8130
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Xem"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   135
      Top             =   8160
      Width           =   615
   End
   Begin VB.ComboBox CboNT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   1
      ItemData        =   "Fchungtu.frx":655C
      Left            =   8160
      List            =   "Fchungtu.frx":655E
      TabIndex        =   132
      Text            =   "CboNT"
      ToolTipText     =   "§¬n gi¸ mÆc ®Þnh"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1570
   End
   Begin VB.ComboBox CboNT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   3
      ItemData        =   "Fchungtu.frx":6560
      Left            =   9840
      List            =   "Fchungtu.frx":6588
      TabIndex        =   131
      Tag             =   "0"
      Text            =   "CboNT"
      ToolTipText     =   "Ngo¹i tÖ ph¸t sinh"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1499
   End
   Begin VB.ComboBox CboNT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   0
      ItemData        =   "Fchungtu.frx":65B3
      Left            =   6360
      List            =   "Fchungtu.frx":65B5
      Style           =   2  'Dropdown List
      TabIndex        =   130
      ToolTipText     =   "Ngo¹i tÖ ph¸t sinh"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox Checkinbangkevahoadon 
      BackColor       =   &H00E0E0E0&
      Caption         =   "In  bang ke"
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
      Left            =   13680
      TabIndex        =   129
      Tag             =   "Direct Export"
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox checkinbangke 
      BackColor       =   &H00E0E0E0&
      Caption         =   "In  bang ke"
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
      Left            =   13680
      TabIndex        =   128
      Tag             =   "Direct Export"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox CheckBox1 
      Caption         =   "Check3"
      Height          =   375
      Left            =   14280
      TabIndex        =   126
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox CheckBox2 
      Caption         =   "Check3"
      Height          =   375
      Left            =   14400
      TabIndex        =   125
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox CheckBox3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   15240
      TabIndex        =   124
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox sochungtu 
      Height          =   315
      Left            =   14040
      LinkItem        =   "Sè hiÖu chøng tõ"
      MaxLength       =   20
      TabIndex        =   123
      Tag             =   "11"
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox thoihanthanhtoan 
      Height          =   315
      Left            =   14040
      LinkItem        =   "Sè hiÖu chøng tõ"
      MaxLength       =   20
      TabIndex        =   122
      Tag             =   "11"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox hinhthucthanhtoan 
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   14040
      LinkItem        =   "Sè hiÖu chøng tõ"
      MaxLength       =   20
      TabIndex        =   121
      Tag             =   "11"
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copy"
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
      Left            =   14280
      TabIndex        =   120
      Tag             =   "Direct Export"
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XuÊt th¼ng"
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
      Left            =   14040
      TabIndex        =   119
      Tag             =   "Direct Export"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   14520
      Style           =   2  'Dropdown List
      TabIndex        =   118
      Top             =   1200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox txtVT 
      Height          =   315
      Index           =   2
      Left            =   18120
      MaxLength       =   20
      TabIndex        =   162
      Tag             =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txttrunggian 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   15120
      MaxLength       =   20
      TabIndex        =   115
      Tag             =   "14"
      Text            =   "0"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox CboNguon 
      Appearance      =   0  'Flat
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
      Index           =   3
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   112
      Top             =   360
      Width           =   1695
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
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
      ForeColor       =   &H00400000&
      Height          =   395
      Index           =   4
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   110
      Tag             =   "Conversion"
      Top             =   1110
      Width           =   1575
   End
   Begin VB.CommandButton CmdChitiet 
      DragIcon        =   "Fchungtu.frx":65B7
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13080
      Picture         =   "Fchungtu.frx":BD99
      TabIndex        =   108
      Tag             =   "-1"
      ToolTipText     =   "Ghi ph¸t sinh"
      Top             =   2160
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.TextBox txtchungtu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   10
      Left            =   13920
      MaxLength       =   20
      TabIndex        =   19
      Tag             =   "14"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   2040
      Width           =   1180
   End
   Begin VB.TextBox txtchungtu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   11520
      MaxLength       =   20
      TabIndex        =   17
      Tag             =   "14"
      Text            =   "0"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   2040
      Width           =   1650
   End
   Begin VB.TextBox txtchungtu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   9795
      MaxLength       =   20
      TabIndex        =   16
      Tag             =   "14"
      Text            =   "0"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtchungtu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   8230
      MaxLength       =   20
      TabIndex        =   15
      Tag             =   "14"
      Text            =   "0"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   2040
      Width           =   1590
   End
   Begin VB.TextBox txtchungtu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   7155
      MaxLength       =   20
      TabIndex        =   14
      Tag             =   "14"
      Text            =   "0"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   2040
      Width           =   1090
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000B&
      Caption         =   "Command2"
      Height          =   375
      Left            =   14400
      TabIndex        =   105
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   14400
      TabIndex        =   104
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtVT 
      Height          =   315
      Index           =   1
      Left            =   5040
      MaxLength       =   20
      TabIndex        =   3
      Tag             =   "1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox CboLoai 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   8760
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1440
      Width           =   2145
   End
   Begin VB.TextBox txtVT 
      Height          =   315
      Index           =   0
      Left            =   5040
      MaxLength       =   500
      TabIndex        =   4
      Tag             =   "14"
      ToolTipText     =   "Gõ mã khách hàng nêu ch?a taïo maõ mã thì  t? gõ mã m?i vào ?ây và ô MST, tên khách hàng, ??ai ch? ch??ng trình s? t? ghi l?i."
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtVT 
      Height          =   315
      Index           =   7
      Left            =   3120
      MaxLength       =   500
      TabIndex        =   6
      Tag             =   "16"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   1080
      Width           =   5400
   End
   Begin VB.TextBox txtVT 
      Height          =   315
      Index           =   8
      Left            =   8760
      MaxLength       =   200
      TabIndex        =   7
      Tag             =   "17"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox txtVT 
      Height          =   315
      Index           =   9
      Left            =   6840
      MaxLength       =   20
      TabIndex        =   5
      Tag             =   "15"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   720
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   14280
      TabIndex        =   71
      Top             =   6000
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtVT 
         Height          =   315
         Index           =   3
         Left            =   480
         MaxLength       =   20
         TabIndex        =   127
         Tag             =   "1"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox Chk 
         BackColor       =   &H80000013&
         Caption         =   "&B¸o gi¸"
         Height          =   255
         Left            =   2640
         TabIndex        =   92
         Tag             =   "&Quotation"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1320
         LinkItem        =   "DiÔn gi¶i chøng tõ"
         MaxLength       =   150
         TabIndex        =   90
         Top             =   3720
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.TextBox txtchungtu 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1800
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   88
         Tag             =   "0"
         Text            =   "Fchungtu.frx":C273
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox CboNguon 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Tag             =   """"""
         Top             =   2040
         Width           =   3015
      End
      Begin VB.ComboBox CboNT 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         ItemData        =   "Fchungtu.frx":C275
         Left            =   1440
         List            =   "Fchungtu.frx":C277
         Style           =   2  'Dropdown List
         TabIndex        =   82
         ToolTipText     =   "Danh s¸ch ®¬n vÞ tÝnh"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtchungtu 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   7560
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   81
         Tag             =   "0"
         Text            =   "Fchungtu.frx":C279
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtchungtu 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   4440
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   80
         Tag             =   "0"
         Text            =   "Fchungtu.frx":C27B
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox CboNguon 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   78
         ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
         Top             =   840
         Width           =   4455
      End
      Begin VB.ComboBox CboVV 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   74
         ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox CboVV 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   73
         ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox CboVV 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   72
         ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label 
         BackColor       =   &H80000013&
         Caption         =   "Description"
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   91
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "Tû gi¸"
         Height          =   255
         Index           =   17
         Left            =   840
         TabIndex        =   89
         Tag             =   "Ex. Rate"
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label LbKho 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "Ph©n lo¹i"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   87
         Tag             =   "Class"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "§.v.t"
         Height          =   255
         Index           =   12
         Left            =   600
         TabIndex        =   85
         Top             =   1515
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label 
         BackColor       =   &H80000013&
         Caption         =   "H¹n thanh to¸n"
         Enabled         =   0   'False
         Height          =   255
         Index           =   22
         Left            =   6240
         TabIndex        =   84
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "Ph¸t sinh USD"
         Height          =   255
         Index           =   25
         Left            =   3120
         TabIndex        =   83
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label 
         BackColor       =   &H80000013&
         Caption         =   "Bé phËn"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   79
         Tag             =   "Index"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LbTT 
         BackColor       =   &H80000013&
         Caption         =   "Th«ng tin 1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   77
         Tag             =   "Index"
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label LbTT 
         BackColor       =   &H80000013&
         Caption         =   "Th«ng tin 2"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   76
         Tag             =   "Index"
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label LbTT 
         BackColor       =   &H80000013&
         Caption         =   "Th«ng tin 3"
         Height          =   255
         Index           =   2
         Left            =   6720
         TabIndex        =   75
         Tag             =   "Index"
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2640
      TabIndex        =   57
      Top             =   3480
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton cmd 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3000
         Picture         =   "Fchungtu.frx":C27D
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtsh 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1560
         LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
         MaxLength       =   20
         TabIndex        =   63
         Tag             =   "0"
         Top             =   525
         Width           =   1335
      End
      Begin VB.CommandButton cmd 
         Height          =   375
         Index           =   0
         Left            =   3000
         Picture         =   "Fchungtu.frx":C6F7
         Style           =   1  'Graphical
         TabIndex        =   61
         Tag             =   "0"
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txtsh 
         Height          =   300
         Index           =   0
         Left            =   1560
         LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
         MaxLength       =   20
         TabIndex        =   60
         Tag             =   "0"
         Top             =   165
         Width           =   1335
      End
      Begin VB.Label lb 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   65
         Tag             =   "1"
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label lb 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   62
         Tag             =   "1"
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C«ng tr×nh"
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
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   59
         Tag             =   "Object"
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ghi nî tµi kho¶n"
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
         Left            =   120
         TabIndex        =   58
         Tag             =   "Deb. Account"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Barcode"
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
      Left            =   12600
      TabIndex        =   70
      Tag             =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtchungtu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   9
      Left            =   13380
      MaxLength       =   5
      TabIndex        =   18
      Tag             =   "0"
      Text            =   "0"
      Top             =   2040
      Width           =   550
   End
   Begin VB.CommandButton CmdPhieu 
      Caption         =   "&4 TCNH"
      Height          =   375
      Index           =   3
      Left            =   10680
      TabIndex        =   51
      Tag             =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   255
      Index           =   3
      Left            =   15480
      LinkItem        =   "Sè hiÖu chøng tõ"
      MaxLength       =   100
      TabIndex        =   11
      Tag             =   "11"
      Top             =   2880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CheckBox chkXT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XuÊt c«ng tr×nh"
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
      Left            =   2160
      TabIndex        =   56
      Tag             =   "Direct Export"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdDanhSach 
      Height          =   375
      Index           =   0
      Left            =   7200
      Picture         =   "Fchungtu.frx":CB71
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton CmdPhieu 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&3 UNC"
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
      Left            =   11760
      TabIndex        =   50
      Tag             =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   4
      Left            =   4800
      Picture         =   "Fchungtu.frx":DF3B
      Style           =   1  'Graphical
      TabIndex        =   52
      Tag             =   "&Print"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.CommandButton CmdDanhSach 
      Height          =   375
      Index           =   1
      Left            =   16200
      Picture         =   "Fchungtu.frx":F39D
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton CmdPhieu 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&2 Ho¸ ®¬n"
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
      Left            =   9600
      TabIndex        =   49
      Tag             =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CmdPhieu 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&1 PhiÕu TC"
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
      Left            =   10680
      TabIndex        =   48
      Tag             =   "0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
      Caption         =   "&B¸n hµng"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   395
      Index           =   8
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "Sales Invoice"
      Top             =   2370
      Width           =   1575
   End
   Begin VB.CommandButton SSCmdV 
      Caption         =   "0"
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
      Left            =   14160
      TabIndex        =   37
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
      Caption         =   "§¸nh gi¸ l¹&i"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   395
      Index           =   11
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   26
      Tag             =   "34"
      ToolTipText     =   "Assets Revaluation"
      Top             =   4070
      Width           =   1575
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
      Caption         =   "Gi¶&m TSC§"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   395
      Index           =   10
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "33"
      ToolTipText     =   "Assets Decreasing"
      Top             =   3630
      Width           =   1575
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
      Caption         =   "T¨ng T&SC§"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   395
      Index           =   9
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "32"
      ToolTipText     =   "Assets Increasing"
      Top             =   3210
      Width           =   1575
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
      Caption         =   " Tæng &hîp"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   395
      HelpContextID   =   800
      Index           =   0
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "Common"
      Top             =   680
      Width           =   1575
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
      Caption         =   "&NhËp vËt t­"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   395
      Index           =   1
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "Import Inventory"
      Top             =   1530
      Width           =   1575
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
      Caption         =   "X&uÊt vËt t­"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   395
      Index           =   2
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   107
      Tag             =   "Export Inventory"
      Top             =   1950
      Width           =   1575
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   0
      Left            =   5040
      LinkItem        =   "Sè hiÖu chøng tõ"
      MaxLength       =   20
      TabIndex        =   2
      Tag             =   "106"
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   1
      Left            =   3120
      LinkItem        =   "DiÔn gi¶i chøng tõ"
      MaxLength       =   500
      TabIndex        =   8
      Tag             =   "18"
      Top             =   1440
      Width           =   5400
   End
   Begin VB.ComboBox CboNguon 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      ItemData        =   "Fchungtu.frx":F817
      Left            =   10920
      List            =   "Fchungtu.frx":F819
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Tag             =   "19"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
      Caption         =   "&KÕt chuyÓn"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   395
      Index           =   3
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "Conversion"
      Top             =   2790
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   3
      Left            =   8400
      Picture         =   "Fchungtu.frx":F81B
      Style           =   1  'Graphical
      TabIndex        =   34
      Tag             =   "&Return"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   2
      Left            =   6000
      Picture         =   "Fchungtu.frx":10C3D
      Style           =   1  'Graphical
      TabIndex        =   33
      Tag             =   "&Delete"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   3615
      Picture         =   "Fchungtu.frx":1211F
      Style           =   1  'Graphical
      TabIndex        =   31
      Tag             =   "&Save"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   2400
      Picture         =   "Fchungtu.frx":1354D
      Style           =   1  'Graphical
      TabIndex        =   32
      Tag             =   "&Add"
      Top             =   4680
      Width           =   1080
   End
   Begin VB.TextBox txtchungtu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   2
      Left            =   5835
      MaxLength       =   20
      TabIndex        =   13
      Tag             =   "0"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   2040
      Width           =   1340
   End
   Begin VB.TextBox txtchungtu 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   0
      Left            =   2160
      LinkItem        =   "Sè hiÖu tµi kho¶n hoÆc chi tiÕt cã ph¸t sinh (nhÊn ENTER ®Ó xem danh s¸ch)"
      MaxLength       =   20
      TabIndex        =   12
      Tag             =   "21"
      Text            =   "?"
      ToolTipText     =   "NhÊn phÝm ? hoÆc click ®Ó hiÖn th«ng tin"
      Top             =   2040
      Width           =   795
   End
   Begin VB.TextBox txtchungtu 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Index           =   1
      Left            =   2940
      LinkItem        =   "Tªn tµi kho¶n hoÆc chi tiÕt"
      MaxLength       =   50
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2950
   End
   Begin VB.OptionButton OptLoai 
      BackColor       =   &H0080FF80&
      Caption         =   "TrÝch khÊu h&ao"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   395
      Index           =   12
      Left            =   120
      MaskColor       =   &H80000001&
      Style           =   1  'Graphical
      TabIndex        =   27
      Tag             =   "35"
      ToolTipText     =   "Depreciation"
      Top             =   4500
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   375
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
      Left            =   3120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSGrid.Grid Grid2 
      Height          =   3015
      Left            =   3600
      TabIndex        =   134
      Tag             =   "1"
      Top             =   5160
      Width           =   9795
      _Version        =   65536
      _ExtentX        =   17277
      _ExtentY        =   5318
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Rows            =   30
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSGrid.Grid GrdChungtu 
      Height          =   2220
      Left            =   1800
      TabIndex        =   160
      Tag             =   "20"
      Top             =   2350
      Width           =   11505
      _Version        =   65536
      _ExtentX        =   20294
      _ExtentY        =   3916
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   20
      Cols            =   29
      FixedRows       =   0
      HighLight       =   0   'False
      MousePointer    =   1
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Import"
      Height          =   375
      Left            =   17160
      TabIndex        =   163
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label txttinh_gia_ban 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
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
      Left            =   16080
      TabIndex        =   161
      Tag             =   "Month"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tæng sè chøng tõ"
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
      Left            =   9840
      TabIndex        =   157
      Tag             =   "Month"
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sè dßng"
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
      Index           =   27
      Left            =   12360
      TabIndex        =   156
      Tag             =   "Month"
      Top             =   8400
      Width           =   1335
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Left            =   6960
      TabIndex        =   155
      Top             =   4080
      Width           =   1455
      BackColor       =   14737632
      Size            =   "2566;661"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   17
      Left            =   1800
      TabIndex        =   154
      Tag             =   "Bill Code"
      Top             =   2040
      Width           =   11295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "®Õn"
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
      Left            =   1560
      TabIndex        =   150
      Tag             =   "Address"
      Top             =   8200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tõ"
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
      Left            =   200
      TabIndex        =   149
      Tag             =   "Address"
      Top             =   8200
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Index           =   8
      Left            =   0
      TabIndex        =   133
      Tag             =   "Bill Code"
      Top             =   4560
      Width           =   18255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "M· h®"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   7
      Left            =   13680
      TabIndex        =   117
      Tag             =   "Bill Code"
      Top             =   480
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MÉu sè"
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
      Height          =   255
      Index           =   6
      Left            =   18120
      TabIndex        =   116
      Tag             =   "Bill Code"
      Top             =   360
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   26
      Left            =   6240
      TabIndex        =   114
      Tag             =   "V. Code"
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "CT GS"
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
      Left            =   6300
      TabIndex        =   113
      Tag             =   "Voucher Type"
      Top             =   465
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4575
      Index           =   5
      Left            =   0
      TabIndex        =   111
      Tag             =   "Bill Code"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label LbKho 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Thu chi"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   3
      Left            =   16500
      TabIndex        =   109
      Tag             =   "Store"
      Top             =   2520
      Width           =   15
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "M· k.hµng"
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
      Left            =   4250
      TabIndex        =   97
      Tag             =   "Liability Code"
      Top             =   800
      Width           =   855
   End
   Begin VB.Label LbKho 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LiÖt kª chøng tõ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   2
      Left            =   14760
      TabIndex        =   106
      Tag             =   "Store"
      Top             =   4920
      Width           =   75
   End
   Begin VB.Label Label1 
      Caption         =   "Thµnh tiÒn tr­íc thuÕ"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   103
      Tag             =   "Amount before Tax"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Tû lÖ thuÕ (%)"
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   102
      Tag             =   "Tax Rate"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Gi¸ tÝnh thuÕ"
      Height          =   255
      Index           =   12
      Left            =   4560
      TabIndex        =   101
      Tag             =   "Taxable Amount"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "H×nh thøc thanh to¸n"
      Height          =   255
      Index           =   13
      Left            =   4920
      TabIndex        =   100
      Tag             =   "Payment Type"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ký hiÖu"
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
      Left            =   4440
      TabIndex        =   99
      Tag             =   "Bill Code"
      Top             =   465
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "§Þa chØ"
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
      Left            =   8760
      TabIndex        =   98
      Tag             =   "Address"
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tªn kh¸ch hµng"
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
      Left            =   1890
      TabIndex        =   96
      Tag             =   "Description"
      Top             =   1170
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chän môc nhËp"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   95
      Tag             =   "Address"
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MST"
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
      Left            =   6450
      TabIndex        =   94
      Tag             =   "Tax Code"
      Top             =   820
      Width           =   375
   End
   Begin VB.Label LBNV 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   13440
      TabIndex        =   93
      Tag             =   "1"
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TiÒn CK"
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
      Index           =   24
      Left            =   13920
      TabIndex        =   69
      Top             =   1800
      Width           =   1180
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CK"
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
      Left            =   13380
      TabIndex        =   68
      Top             =   1800
      Width           =   550
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "NV B¸n hµng"
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
      Left            =   15480
      TabIndex        =   67
      Tag             =   "Salesman"
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STT"
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
      Index           =   20
      Left            =   1800
      TabIndex        =   66
      Tag             =   "No."
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label LbUser 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   14520
      TabIndex        =   55
      Tag             =   "1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ng­êi nhËp"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Index           =   11
      Left            =   14760
      TabIndex        =   54
      Tag             =   "Input by"
      Top             =   1680
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
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
      Index           =   13
      Left            =   8235
      TabIndex        =   53
      Tag             =   "Unit price"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   13440
      X2              =   13440
      Y1              =   0
      Y2              =   8160
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   3600
      X2              =   3600
      Y1              =   8160
      Y2              =   5160
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th¸ng nhËp "
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
      Left            =   1925
      TabIndex        =   38
      Tag             =   "Month"
      Top             =   100
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ngµy chøng tõ"
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
      Left            =   1920
      TabIndex        =   39
      Tag             =   "V. Date"
      Top             =   465
      Width           =   1200
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ngµy ghi sæ"
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
      Left            =   1920
      TabIndex        =   35
      Tag             =   "B. Date"
      Top             =   810
      Width           =   960
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
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
      Index           =   3
      Left            =   4440
      TabIndex        =   36
      Tag             =   "V. Code"
      Top             =   90
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Néi dung"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   1920
      TabIndex        =   47
      Tag             =   "Desc. (V)"
      Top             =   1500
      Width           =   735
   End
   Begin VB.Label LbKho 
      BackColor       =   &H00E0E0E0&
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
      Height          =   15
      Index           =   0
      Left            =   13560
      TabIndex        =   46
      Tag             =   "Store"
      Top             =   1560
      Visible         =   0   'False
      Width           =   15
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ph¸t sinh cã"
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
      Index           =   10
      Left            =   11475
      TabIndex        =   45
      Tag             =   "Credit"
      Top             =   1800
      Width           =   1650
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ph¸t sinh nî"
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
      Index           =   9
      Left            =   9795
      TabIndex        =   44
      Tag             =   "Debit"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
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
      Index           =   8
      Left            =   7155
      TabIndex        =   43
      Tag             =   "Quantity"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   5835
      TabIndex        =   42
      Tag             =   "Code"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DiÔn gi¶i"
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
      Left            =   2940
      TabIndex        =   41
      Tag             =   "Description"
      Top             =   1800
      Width           =   2955
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tµi kho¶n"
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
      Index           =   5
      Left            =   2160
      TabIndex        =   40
      Tag             =   "Account"
      Top             =   1800
      Width           =   795
   End
   Begin VB.Menu mnPU 
      Caption         =   "&Danh ®iÓm"
      Visible         =   0   'False
      Begin VB.Menu mnDD 
         Caption         =   "Ph©n lo¹i vËt t­, hµng ho¸..."
         Index           =   0
         Tag             =   "Inventory classification..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "Ph©n lo¹i tµi s¶n cè ®Þnh..."
         Index           =   1
         Tag             =   "Fixed Assets Classification..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "Ph©n lo¹i c«ng nî..."
         Index           =   2
         Tag             =   "Receivable and Payable classification..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnDD 
         Caption         =   "HÖ thèng tµi kho¶n..."
         Index           =   4
         Tag             =   "Chart of Account..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "Danh ®iÓm vËt t­..."
         Index           =   5
         Tag             =   "Inventory List..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "Danh ®iÓm TSC§..."
         Index           =   6
         Tag             =   "Fixed Asset List..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "Danh ®iÓm c«ng nî"
         Index           =   7
         Tag             =   "Receivabe and Payable List ..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnDD 
         Caption         =   "Nguån nhËp xuÊt..."
         Index           =   9
         Tag             =   "Inventory Source..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "Vô viÖc liªn quan..."
         Index           =   10
         Tag             =   "Other Index..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "Hîp ®ång kinh tÕ..."
         Index           =   11
         Tag             =   "Contract List..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnDD 
         Caption         =   "B¸o c¸o chi tiÕt..."
         Index           =   13
         Tag             =   "Detail Report..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "B¸o c¸o tæng hîp..."
         Index           =   14
         Tag             =   "Summary Report..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnDD 
         Caption         =   "Sè hiÖu chøng tõ..."
         Index           =   16
         Tag             =   "Invoice Default Code..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "Danh s¸ch chøng tõ"
         Index           =   17
         Tag             =   "Invoice List"
      End
      Begin VB.Menu mnDD 
         Caption         =   "NhËt ký chung"
         Index           =   18
         Tag             =   "Journal Ledger"
      End
      Begin VB.Menu mnDD 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu mnDD 
         Caption         =   "In toµn bé phiÕu thu"
         Index           =   20
         Tag             =   "Print Receipt Voucher"
      End
      Begin VB.Menu mnDD 
         Caption         =   "In toµn bé phiÕu chi"
         Index           =   21
         Tag             =   "Print Payment Voucher"
      End
      Begin VB.Menu mnDD 
         Caption         =   "In toµn bé phiÕu nhÊp"
         Index           =   22
         Tag             =   "Print Inventory Import Voucher"
      End
      Begin VB.Menu mnDD 
         Caption         =   "In toµn bé phiÕu xuÊt"
         Index           =   23
         Tag             =   "Print Inventory Export Voucher"
      End
      Begin VB.Menu mnDD 
         Caption         =   "-"
         Index           =   24
      End
      Begin VB.Menu mnDD 
         Caption         =   "§¬n gi¸ nhËp míi nhÊt"
         Index           =   25
         Tag             =   "Unit price List"
         Visible         =   0   'False
      End
      Begin VB.Menu mnDD 
         Caption         =   "-"
         Index           =   26
         Visible         =   0   'False
      End
      Begin VB.Menu mnDD 
         Caption         =   "Th«ng tin chøng tõ 1"
         Index           =   27
         Visible         =   0   'False
      End
      Begin VB.Menu mnDD 
         Caption         =   "Th«ng tin chøng tõ 2"
         Index           =   28
         Visible         =   0   'False
      End
      Begin VB.Menu mnDD 
         Caption         =   "Th«ng tin chøng tõ 3"
         Index           =   29
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmChungtu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 ' Ð?u tiên, khai báo các API c?n thi?t
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
 

' Bi?n toàn c?c d? luu handle c?a c?a s? ?ng d?ng dã m?
Dim hWndApp As Long
 
 
 
 
Public fileImportList As Collection


Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Const TM = "111"
Const NH = "112"

Dim IsImport As Boolean

Dim tempchungtu As String

Dim rs_ktra152 As Recordset
Dim rs_ktra154c As Recordset
Dim rs_ktra711 As Recordset
Dim stt As Integer
Dim IndexFirst As Integer
Dim IdDuyet As Integer
Dim item As ClsFileImport
Dim i As Integer
Dim displayInfo As String
Dim xmlDoc As Object
Dim fDialog As Object
Dim dlhDonNode As Object
Dim ttChungNode As Object
Dim ndhDonNode As Object
Dim mstNode As Object
Dim TTNode As Object
Dim tenNode As Object
Dim DChiNode As Object
Dim convertedDate As Date
Dim FilePath As String
Dim cbbThang As String

Public MaSoCT As Long
Dim ngay(0 To 1) As Date                              ' Ngµy chøng tõ
Dim loaict As Integer                                        ' Lo¹i chøng tõ
Dim taikhoan As New ClsTaikhoan            ' Tµi kho¶n ®ang nhËp ph¸t sinh
Dim vattu As New ClsVattu                            ' VËt t­ ®ang ®­îc nhËp ph¸t sinh
Dim ckh As New ClsKhachHang
Dim tp As New Cls154
Dim MaNhap As Long
Dim nhieunoco As Boolean
Dim VTEnable As Boolean
Dim KhongNhapTS As Boolean
Dim SetLoaiEnable As Boolean
Dim shct As String
Dim xddu As Boolean
Dim TenTC As String, DiachiTC As String, ctgoc As String, TenNX As String, DiaChiNX As String, TenBH As String, DiaChiBH As String, MSTBH As String, unc1 As String, unc2 As String, unc3 As String, MaKHBH As Long, HanTT As Date
Attribute DiachiTC.VB_VarUserMemId = 1073938468
Attribute ctgoc.VB_VarUserMemId = 1073938468
Attribute TenNX.VB_VarUserMemId = 1073938468
Attribute DiaChiNX.VB_VarUserMemId = 1073938468
Attribute TenBH.VB_VarUserMemId = 1073938468
Attribute DiaChiBH.VB_VarUserMemId = 1073938468
Attribute MSTBH.VB_VarUserMemId = 1073938468
Attribute unc1.VB_VarUserMemId = 1073938468
Attribute unc2.VB_VarUserMemId = 1073938468
Attribute unc3.VB_VarUserMemId = 1073938468
Attribute MaKHBH.VB_VarUserMemId = 1073938468
Attribute HanTT.VB_VarUserMemId = 1073938468
Dim HD() As tpHoaDon, hdcount As Integer
Attribute HD.VB_VarUserMemId = 1073938459
Attribute hdcount.VB_VarUserMemId = 1073938459
Dim Ppthu As Integer, Ppchi As Integer, Ppunc As Integer
Attribute Ppthu.VB_VarUserMemId = 1073938461
Attribute Ppchi.VB_VarUserMemId = 1073938461
Attribute Ppunc.VB_VarUserMemId = 1073938461
Dim pVAT1 As Integer, pVAT2 As Integer, vBH As Integer
Attribute pVAT1.VB_VarUserMemId = 1073938464
Attribute pVAT2.VB_VarUserMemId = 1073938464
Attribute vBH.VB_VarUserMemId = 1073938464
Dim P_1 As Integer, LC As Integer
Attribute P_1.VB_VarUserMemId = 1073938467
Attribute LC.VB_VarUserMemId = 1073938467
Dim MaTS(0 To 9) As Long, tscount As Integer
Attribute MaTS.VB_VarUserMemId = 1073938469
Attribute tscount.VB_VarUserMemId = 1073938469
Dim pMaBG As Long
Attribute pMaBG.VB_VarUserMemId = 1073938471

Dim pGhi As Integer
Attribute pGhi.VB_VarUserMemId = 1073938472
Dim pRate As Double
Attribute pRate.VB_VarUserMemId = 1073938473

Dim bcstop As Integer
Attribute bcstop.VB_VarUserMemId = 1073938474

Public cho_hien_vat As Boolean
Attribute cho_hien_vat.VB_VarUserMemId = 1073938475
Dim cho_hien_thongbao As Boolean
Attribute cho_hien_thongbao.VB_VarUserMemId = 1073938476
Dim dathuchien As Boolean
Attribute dathuchien.VB_VarUserMemId = 1073938477

Dim hien_bang_tinh As Boolean
Attribute hien_bang_tinh.VB_VarUserMemId = 1073938478
Public tongtientruoc As Double
Attribute tongtientruoc.VB_VarUserMemId = 1073938479
Public Sub AutoCLickLoai()
    OptLoai(0).Value = True
    OptLoai_LostFocus 0
    RFocus CboThang
    DisplayFileImportList
End Sub
Public Sub AddImportData(ByVal id As String, ByVal Name As String, ByVal mst As String, ByVal sohd As String, ByVal khHD As String, ByVal ngay As Date, ByVal types As String, ByVal path As String, ByVal tkno As String, ByVal TkCo As String, ByVal tkThue As String, ByVal diengiai As String, ByVal TongTien As String, ByVal VAT As String, ByVal sohieutp As String)
    Dim fileImport As ClsFileImport
    Set fileImport = New ClsFileImport

    ' Gán giá tr? cho các thu?c tính
    fileImport.id = id
    fileImport.Name = Name
    fileImport.mst = mst
    fileImport.sohd = sohd
    fileImport.khHD = khHD
    fileImport.ngay = ngay
    fileImport.types = types
    fileImport.patTH = path
    fileImport.cotk = TkCo
    fileImport.notk = tkno
    fileImport.ThueTK = tkThue
    fileImport.diengiai = diengiai
    fileImport.TongTien = TongTien
    fileImport.VAT = VAT
    fileImport.sohieutp = sohieutp
    fileImportList.Add fileImport

End Sub
Public Sub DoneSetup()
     Timer1.Enabled = True
End Sub

Public Sub DisplayFileImportList()
    
    IdDuyet = 1
    Set item = fileImportList(IdDuyet)
    DuyetItemList item.patTH
    
    ' Hi?n th? thông tin
End Sub
Private Sub DuyetItemList(ByVal fname As String)
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    FilePath = fname
    If xmlDoc.Load(FilePath) Then
        ' L?y các node
        Set dlhDonNode = xmlDoc.selectSingleNode("/HDon/DLHDon")
        Set ttChungNode = xmlDoc.selectSingleNode("/HDon/DLHDon/TTChung")
        Set ndhDonNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon")
        Set mstNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/MST")
        Set tenNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/Ten")
        Set DChiNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/DChi")
        Set TTNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/TToan/TgTCThue")
        ' Hi?n th? thông tin
        If Not dlhDonNode Is Nothing Then
            ' txtID.Text = dlhDonNode.Attributes.getNamedItem("Id").Text
        Else
            MsgBox "Không tìm th?y DLHDon."
        End If

        If Not ttChungNode Is Nothing Then
            Dim shDonNode As Object
            Dim shKHHDNode As Object
            Dim shNLapNode As Object

            Set shDonNode = ttChungNode.getElementsByTagName("SHDon")(0)
            Set shKHHDNode = ttChungNode.getElementsByTagName("KHHDon")(0)
            Set shNLapNode = ttChungNode.getElementsByTagName("NLap")(0)

            If Not shDonNode Is Nothing Then
                txt(0).Text = shDonNode.Text
                txtVT(1).Text = shKHHDNode.Text

                If Not shNLapNode Is Nothing Then
                    convertedDate = CDate(shNLapNode.Text)
                    MedNgay(0).Text = Format(convertedDate, "dd/mm/yy")
                    cbbThang = CboThang.Text    ' L?y giá tr? du?c ch?n t? ComboBox
                    Dim monthValue As Integer
                    Dim monthValue2 As Integer
                    monthValue = Month(CDate(cbbThang))
                    monthValue2 = Month(CDate(convertedDate))
                    ' So sánh tháng l?y du?c v?i tháng hi?n t?i
                    If monthValue <> monthValue2 Then
                        Dim dateString As String
                        dateString = "1/" & monthValue & "/" & Year(Date)
                        MedNgay(1).Text = Format(dateString, "dd/mm/yy")
                    End If
                End If
                If Not mstNode Is Nothing Then
                    txtVT(9).Text = mstNode.Text
                    GetcustomerByMST txtVT(9).Text, tenNode.Text, DChiNode.Text
                Else
                    MsgBox "Không tìm th?y MST."
                End If
                With fileImportList(IdDuyet)
                    txt(1).Text = .diengiai
                End With

                ' txtchungtu(0).Text = 6422
                With fileImportList(IdDuyet)
                    txtchungtu(0).Text = .notk
                End With

                txtChungtu_LostFocus (0)
                txtchungtu(5).Text = TTNode.Text
                RFocus txtchungtu(6)
                txtChungtu_KeyPress 6, 13

                'txtchungtu(0).Text = 1331
                With fileImportList(IdDuyet)
                    If .ThueTK <> "" Then
                        txtchungtu(0).Text = .ThueTK
                    Else
                        If .notk = 5111 Then
                        txtchungtu(0).Text = 33311
                        Else
                        txtchungtu(0).Text = 1331
                        End If
                        
                    End If

                End With
                txtChungtu_LostFocus (0)
                txtchungtu(2).Text = 8
                txtChungtu_LostFocus (2)
                RFocus txtchungtu(6)
                txtChungtu_KeyPress 6, 13

                'txtchungtu(0).Text = 1111
                With fileImportList(IdDuyet)
                    txtchungtu(0).Text = .cotk
                End With
                FThuChi.FThuChiForm = 1
                If stt < 2 Then
                    txtChungtu_LostFocus (0)
                End If

                stt = stt + 1

            Else
                MsgBox "Không tìm th?y SHDon."
            End If
        Else
            MsgBox "Không tìm th?y TTChung."
        End If

        If Not ndhDonNode Is Nothing Then
            ' X? lý ndhDonNode n?u c?n
        End If
    Else
        MsgBox "L?i khi t?i file XML: " & xmlDoc.parseError.reason
    End If
End Sub


Private Sub btnImport_Click()
    Set fileImportList = New Collection

    IsImport = True
    ' Duyet du lieu tu tb_import
    Dim rs_ktra As Recordset
    Dim Query As String
    Dim rst As String

    Query = "select * from tbimport where Status=0"
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? các b?n ghi
        Do While Not rs_ktra.EOF
            ' L?y s? lu?ng tru?ng
            AddImportData rs_ktra!id, rs_ktra!Ten, rs_ktra!mst, rs_ktra!SHDon, rs_ktra!KHHDon, rs_ktra!NLap, "", "", rs_ktra!tkno, rs_ktra!TkCo, rs_ktra!tkThue, rs_ktra!Noidung, rs_ktra!TongTien, rs_ktra!VAT, rs_ktra!sohieutp
            rs_ktra.MoveNext
        Loop
    End If

    ' Xu ly phan tu dau tien

    'Neu list rong thi ko thuc hien
    If fileImportList.count = 0 Then
        MsgBox "Khong co data can xu ly!"
        Exit Sub
    End If

    IndexFirst = 1
    Set item = fileImportList(IndexFirst)
    Dim notk As String
    With fileImportList(IndexFirst)
        Xulyimport item
    End With
End Sub
Private Sub Another()

    rs_ktra152.MoveNext
    Timer4.Enabled = True

End Sub
Private Sub Xulyimport(ByVal item As ClsFileImport)

    Dim QueryUpdate As String
    Dim rstUPdate As String
    QueryUpdate = "UPDATE tbimport SET Status = 1 where ID= " & item.id & ""
    'Set rstUPdate = DBKetoan.OpenRecordset(QueryUpdate, dbOpenSnapshot)
    ExecuteSQL5 QueryUpdate


    ' Do data tu tbimport len form
    If item.notk = "711" Then
        OptLoai(0).Value = True
        OptLoai_LostFocus 0
        RFocus CboThang

    End If

    ' If item.notk = "6422" Or item.notk = "6421" Then
    If item.notk Like "642*" Then
        OptLoai(0).Value = True
        OptLoai_LostFocus 0
        RFocus CboThang
    End If

    If (item.notk Like "15*") And (item.notk <> "154") Then
        'If item.notk = "152" Or item.notk = "156" Or item.notk = "153" Or item.notk = "155" Then
        OptLoai(1).Value = True
        OptLoai_LostFocus 1
        RFocus CboThang
    End If
    If item.notk = "154" Then
        OptLoai(1).Value = True
        OptLoai_LostFocus 1
        RFocus CboThang
    End If

    If item.notk Like "511*" Then
        ' If item.notk = "5111" Then
        OptLoai(8).Value = True
        OptLoai_LostFocus 8
        RFocus CboThang
    End If

    Dim myDate As Date
    myDate = CDate(item.ngay)
    txt(0).Text = item.sohd
    txtVT(1).Text = item.khHD

    CboThang.Text = Month(myDate) & "/" & Year(myDate)
    MedNgay(0).Text = Format(item.ngay, "dd/mm/yy")

    ' If Month(myDate) <> Month(Date) Then
    ' MedNgay(1).Text = DateSerial(Year(Date), Month(Date), 1)
    'Else
    MedNgay(1).Text = Format(item.ngay, "dd/mm/yy")
    'End If

    Dim rs_ktra As Recordset
    Dim Query As String
    Dim rst As String
    Query = "SELECT Ten, DiaChi, MST FROM KhachHang WHERE MST = '" & item.mst & "'"
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? các b?n ghi
        Do While Not rs_ktra.EOF
            ' L?y s? lu?ng tru?ng
            txtVT(9).Text = item.mst
            rs_ktra.MoveNext

        Loop
    End If
    With fileImportList(IndexFirst)
        txt(1).Text = .diengiai
    End With


    ' txtchungtu(0).Text = 6422
    With fileImportList(IndexFirst)
        txtchungtu(0).Text = .notk
        tempchungtu = .notk
    End With


    If txtchungtu(0).Text = "711" Then
        FThuChi.FThuChiForm = 2

        txtChungtu_LostFocus (0)
        ' T?o truy v?n SQL d? l?y thông tin khách hàng theo MST
        With fileImportList(IndexFirst)
            Query = "SELECT * FROM tbimportdetail WHERE ParentId='" & .id & "' AND DVT = 'Exception'"
        End With
        Set rs_ktra711 = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
        If Not rs_ktra711.EOF Then
            txtchungtu(6).Text = rs_ktra711!dongia
            RFocus txtchungtu(6)
            txtChungtu_KeyPress 6, 13
            '1331 thue
            txtchungtu(0).Text = item.ThueTK
            txtChungtu_LostFocus (0)
            txtchungtu(2).Text = item.VAT
            Dim number As Long
            Dim VAT As Integer
            VAT = CInt(item.VAT)
            number = CLng(Replace(txtchungtu(5).Text, ",", ""))
            number = number * VAT / 100
            txtchungtu(5).Text = number * (-1)
            RFocus txtchungtu(6)
            txtChungtu_KeyPress 6, 13
            '3311
            RFocus txtchungtu(0)
            txtchungtu(0).Text = item.cotk
            txtChungtu_LostFocus (0)
            RFocus txtchungtu(6)
            txtChungtu_KeyPress 6, 13
            rs_ktra711.MoveNext
            Timer5_Timer
        End If
    End If

    If txtchungtu(0).Text Like "511*" Then
        'If (txtchungtu(0).Text = "5111") Or (txtchungtu(0).Text = "5112") Then
        FThuChi.FThuChiForm = 2

        txtChungtu_LostFocus (0)
        ' T?o truy v?n SQL d? l?y thông tin khách hàng theo MST
        With fileImportList(IndexFirst)
            Query = "SELECT * FROM tbimportdetail WHERE ParentId='" & .id & "' AND DVT <> 'Exception'"
        End With

        ' M? Recordset d? l?y thông tin khách hàng
        Set rs_ktra152 = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

        ' Ki?m tra xem Recordset có d? li?u không
        If Not rs_ktra152.EOF Then

            'MsgBox rs_ktra152!sohieu
            txtchungtu(2).Text = rs_ktra152!sohieu
            txtChungtu_LostFocus (2)
            txtchungtu(3).Text = rs_ktra152!SoLuong
            txtChungtu_LostFocus (3)
            RFocus txtchungtu(4)
            txtchungtu(4).Text = rs_ktra152!dongia
            txtChungtu_LostFocus (4)
            RFocus txtchungtu(6)
            txtChungtu_KeyPress 6, 13
            Another

        Else
            'Neu la 153
            If txtchungtu(0).Text = "5113" Then
                MsgBox "dung lai"
            Exit Sub
        End If

        Timer4.Enabled = True
    End If


End If

If (txtchungtu(0).Text = "154") Then
    txtChungtu_LostFocus (0)
    RFocus txtchungtu(1)
    'MsgBox item.sohieutp
    With fileImportList(IndexFirst)
        Query = "SELECT * FROM TP154 WHERE SoHieu='" & .sohieutp & "'"
    End With
    Dim rs_ktra154 As Recordset
    ' M? Recordset d? l?y thông tin khách hàng
    Set rs_ktra154 = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    Dim hasdata As Boolean
    hasdata = False
    If Not rs_ktra154.EOF Then
        hasdata = True
        RFocus txtchungtu(2)
        txtchungtu(2).Text = rs_ktra154!sohieu
        txtChungtu_LostFocus (2)
        txtchungtu(5).Text = item.TongTien
        txtChungtu_LostFocus (5)
        RFocus txtchungtu(6)
        txtChungtu_KeyPress 6, 13

        rs_ktra154.MoveNext
    End If
    If hasdata = True Then
        '1331
        txtchungtu(0).Text = 1331
        txtChungtu_LostFocus (0)
        txtchungtu(2).Text = item.VAT
        txtChungtu_LostFocus (2)
        RFocus txtchungtu(6)
        txtChungtu_KeyPress 6, 13
        '1111
        With fileImportList(IndexFirst)
            txtchungtu(0).Text = .cotk
        End With
        FThuChi.FThuChiForm = 1
        'If stt < 2 Then
        ' txtChungtu_KeyPress 6, 13
        'End If
        txtChungtu_KeyPress 6, 13
        stt = stt + 1
    Else
        'Xu ly1 con cua 154
        With fileImportList(IndexFirst)
            Query = "SELECT * FROM tbimportdetail WHERE ParentId='" & .id & "'"
        End With
        Set rs_ktra154c = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

        Xuly154

    End If  'cua hasdata
End If      ' end cua 154
If (txtchungtu(0).Text Like "15*") And (txtchungtu(0).Text <> "154") Then
    ' If (txtchungtu(0).Text = "152" Or txtchungtu(0).Text = "156" Or txtchungtu(0).Text = "153" Or txtchungtu(0).Text = "155") Then
    FThuChi.FThuChiForm = 2
    ' Duyet con ben trong
    txtChungtu_LostFocus (0)
    ' T?o truy v?n SQL d? l?y thông tin khách hàng theo MST
    With fileImportList(IndexFirst)
        Query = "SELECT * FROM tbimportdetail WHERE ParentId='" & .id & "' AND DVT <> 'Exception' ORDER BY DonGia ASC"
    End With

    ' M? Recordset d? l?y thông tin khách hàng
    Set rs_ktra152 = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

    ' Ki?m tra xem Recordset có d? li?u không
    If Not rs_ktra152.EOF Then
        'MsgBox rs_ktra152!sohieu
        txtchungtu(2).Text = rs_ktra152!sohieu
        txtChungtu_LostFocus (2)
        txtchungtu(3).Text = rs_ktra152!SoLuong
        txtChungtu_LostFocus (3)
        RFocus txtchungtu(4)

        txtchungtu(4).Text = rs_ktra152!dongia
        txtChungtu_LostFocus (4)
        RFocus txtchungtu(5)
        RFocus txtchungtu(6)
        txtChungtu_KeyPress 6, 13
        Another
    Else
        Timer4.Enabled = True
    End If

End If

If txtchungtu(0).Text Like "642*" Then
    'If (txtchungtu(0).Text = "6422") Then
    With fileImportList(IndexFirst)
        MedNgay(0).Text = Format(.ngay, "dd/mm/yy")
        MedNgay(1).Text = Format(.ngay, "dd/mm/yy")
        txtNgaychungtu.Text = Format(.ngay, "dd/mm/yy")
        txtNgayghiso.Text = Format(.ngay, "dd/mm/yy")
    End With

    txtChungtu_LostFocus (0)
    With fileImportList(IndexFirst)
        txtchungtu(5).Text = .TongTien

    End With
    RFocus txtchungtu(6)
    txtChungtu_KeyPress 6, 13

    'txtchungtu(0).Text = 1331
    With fileImportList(IndexFirst)
        If .ThueTK <> "" Then
            txtchungtu(0).Text = .ThueTK
        Else
            txtchungtu(0).Text = 1331
        End If

    End With
    txtChungtu_LostFocus (0)
    With fileImportList(IndexFirst)
        txtchungtu(2).Text = .VAT
    End With

    txtChungtu_LostFocus (2)
    RFocus txtchungtu(6)
    txtChungtu_KeyPress 6, 13

    'txtchungtu(0).Text = 1111
    With fileImportList(IndexFirst)
        txtchungtu(0).Text = .cotk
        If .cotk = "3311" Or .cotk = "3388" Then

            FThuChi.FThuChiForm = 1
            If stt < 2 Then
                txtChungtu_LostFocus (0)
            End If
            stt = stt + 1
            timer3311.Enabled = True
        Else
            FThuChi.FThuChiForm = 1
            If stt < 2 Then
                txtChungtu_LostFocus (0)
            End If

            stt = stt + 1
        End If

    End With

End If
'MedNgay(0).Text = txtNgaychungtu.Text
'MedNgay(1).Text = txtNgayghiso.Text

End Sub

Private Sub btnOpenexe_Click()
    Dim exePath As String
    exePath = App.path & "\\Tools\\Debug\SaovietTax.exe"

    'MsgBox exePath
    ' exePath = "C:\TCP\Saoviet\SaovietWF\SaovietWF\bin\Debug\SaovietWF.exe"    ' Thay b?ng du?ng d?n th?c t?
    Shell exePath, vbNormalFocus  ' Thay d?i du?ng d?n t?i ?ng d?ng c?a b?n
    ' Ð?i m?t chút d? ?ng d?ng m?
    DoEvents
    ' Thêm th?i gian ch? d? ?ng d?ng có th?i gian kh?i d?ng
    Sleep 2000  ' Ngh? 1 giây (1000 milliseconds)
    
    ' L?y handle c?a c?a s? ?ng d?ng
    hWndApp = FindWindow(vbNullString, "frmMain")    ' Thay d?i tiêu d? c?a ?ng d?ng
    ' Ki?m tra xem handle có h?p l? hay không tru?c khi ki?m tra liên t?c
    If hWndApp = 0 Then
        MsgBox "Khong tim thay ung dung."
    Else
        ' B?t d?u ki?m tra c? quay
        CheckWindow
    End If
End Sub
Private Sub CheckWindow()
' Ki?m tra liên t?c xem c?a s? còn t?n t?i hay không
    Do
        If IsWindow(hWndApp) = 0 Then
            'btnImport_Click
            Dim FilePath As String
            Dim FileNum As Integer
            Dim fileContent As String
            FilePath = App.path & "\\Hoadon\\status.txt"

            ' FilePath = "C:\S.T.E 25\S.T.E 25\Hoadon\status.txt"

            FileNum = FreeFile    'L?y s? file t? do

            Dim lineText As String
            Dim allText As String

            ' M? file d? d?c
            Open FilePath For Input As #FileNum

            ' Ð?c t?ng dòng d?n h?t file
            Do Until EOF(FileNum)
                Line Input #FileNum, lineText
                allText = allText & lineText & vbCrLf    'N?i dòng và xu?ng dòng
            Loop

            ' Ðóng file
            Close #FileNum
            Dim textss As String
            textss = "ButtonClicked"
            Dim textss2 As String
            textss2 = SuperTrim(allText)
            If textss = textss2 Then
                timerImport.Enabled = True

            End If

            Exit Do
        End If
        DoEvents  ' Cho phép ?ng d?ng x? lý các s? ki?n khác
    Loop
End Sub
Function SuperTrim(ByVal s As String) As String
    ' Xóa t?t c? ký t? tr?ng (kho?ng tr?ng, tab, xu?ng dòng)
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCrLf, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    SuperTrim = Trim(s)  ' Xóa kho?ng tr?ng d?u/cu?i (ASCII 32)
End Function
Private Sub CboLoai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And KHDetail Then
        'RFocus txtVT(7)
        RFocus txtchungtu(0)
    End If
End Sub

Private Sub CboLoai_LostFocus()
    If Len(Replace(Trim(txtVT(7).Text), ".", "")) <= 0 Then
        txtVT(7).Text = "..."

    End If
    'RFocus txtVT(7)
End Sub

Private Sub CboNT_Click(Index As Integer)
    Dim gia As Double, hsqd As Double, luong As Double, tien As Double, tien2 As Double

    Select Case Index
    Case 0:
        If CboNT(0).ItemData(CboNT(0).ListIndex) = 0 Then
            txtchungtu(2).tag = 0
            txtchungtu(3).Text = "0"
            txtchungtu(3).Enabled = False
            txtchungtu(4).Text = "0"
        Else
            txtchungtu(2).tag = 1
            txtchungtu(3).Enabled = True
            txtchungtu(4).Enabled = True
            If pTyGiaBQ = 0 Then
                txtchungtu(4).Text = Format(TyGiaNT(CboNT(0).ItemData(CboNT(0).ListIndex)), Mask_0)
            Else
                txtchungtu(4).Text = Format(TyGiaBQ(taikhoan.sohieu, CboNT(0).ItemData(CboNT(0).ListIndex), ngay(0)), Mask_0)
            End If
        End If
    Case 1:
        txtchungtu(4).Text = CboNT(1).Text
        txtChungtu_LostFocus 4
    Case 2:
        If vattu.MaSo > 0 Then
            txtchungtu(1).Text = vattu.TenVattu + " - " + ABCtoVNI("§.v.t: ") + CboNT(2).Text

            gia = GiaBanQD(CboNT(2).ItemData(CboNT(2).ListIndex), hsqd)
            If hsqd = 0 Then hsqd = 1

            'If OutCost = 0 Then
            luong = SoTonKho(CboThang.ItemData(CboThang.ListIndex), CboNguon(1).ItemData(CboNguon(1).ListIndex), 0, vattu.MaSo, tien, tien2)

            luong = luong / hsqd

            txtchungtu(3).Text = Format(luong, Mask_2)
            txtchungtu(3).tag = luong
            txtchungtu(6).Text = Format(tien, Mask_2)

            If pGiaUSD > 0 Then txtchungtu(11).Text = Format(tien2, Mask_2)

            txtchungtu(6).tag = tien
            txtchungtu(5).tag = tien2

            If luong <> 0 Then
                If pGiaUSD > 0 Then
                    txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien2 / luong) / Mask_N, Mask_2)
                Else
                    txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien / luong) / Mask_N, Mask_2)
                End If
            End If
            'End If

            If CboNT(2).ItemData(CboNT(2).ListIndex) > 0 Then

                If gia > 0 Then
                    txtchungtu(4).Text = Format(gia, Mask_2)
                    CboNT(1).Text = txtchungtu(4).Text
                Else
                    If luong <> 0 Then gia = tien / luong
                    If gia > 0 Then
                        txtchungtu(4).Text = Format(gia, Mask_2)
                        CboNT(1).Text = txtchungtu(4).Text
                    End If
                End If
                ' LK Ton


            Else
                If CboNT(1).ListCount > 0 Then CboNT(1).ListIndex = 0
                ' LK Ton

            End If

            HienThongBao "Sè l­îng tån kho: " + txtchungtu(3).Text + " - Thµnh tiÒn: " + Format(txtchungtu(6).tag, Mask_0), 1
        End If
    Case 3:
        If CboNT(3).ListIndex > 0 Then
            If CboNT(3).tag = 1 Then
                txtchungtu(2).Text = MaSo2SoHieu(CboNT(3).ItemData(CboNT(3).ListIndex), "Vattu")
            Else
                txtchungtu(2).Text = MaSo2SoHieu(CboNT(3).ItemData(CboNT(3).ListIndex), "KhachHang")
            End If
            txtChungtu_LostFocus 2
        End If
    End Select
End Sub

Private Sub CboNguon_Click(Index As Integer)
    Select Case Index:
    Case 1:
        If loaict = 2 Then
            MaSoCT = 0
            ClearGrid GrdChungtu, GrdChungtu.tag
            If vattu.MaSo > 0 Then txtChungtu_LostFocus 2
        End If
    End Select
End Sub

Private Sub CboNT_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then
        Select Case KeyAscii
        Case 13:
            CboNT_LostFocus 3
        Case 32:
            LaySohieuDoiTuong2 CboNT(3).tag, CboNT(3).Text
            KeyAscii = 0
        End Select
    End If
End Sub

Private Sub CboNT_LostFocus(Index As Integer)
    Dim sh As String

    Select Case Index
    Case 1:
        txtchungtu(4).Text = CboNT(1).Text
        txtChungtu_LostFocus 4
    Case 3:
        If CboNT(3).ListIndex < 0 Then SetListIndex2 CboNT(3), CboNT(3).Text
        If CboNT(3).ListIndex >= 0 Then
            txtchungtu(2).Text = SelectSQL("SELECT SoHieu AS F1 FROM " + IIf(CboNT(3).tag = 1, "Vattu", "KhachHang") + " WHERE MaSo=" + CStr(CboNT(3).ItemData(CboNT(3).ListIndex)))
            txtChungtu_LostFocus 2
        Else
            If CboNT(3).Text <> "" Then
                Select Case CboNT(3).tag
                Case 1:
                    sh = FrmVattu.ThemVattu(CboNT(3).Text)
                Case 2:
                    sh = FrmKhachHang.ThemKhachHang(CboNT(3).Text)
                End Select
                CboNT(3).Text = sh
                LaySohieuDoiTuong2 CboNT(3).tag, CboNT(3).Text
            End If
        End If
    End Select
End Sub

Private Sub CboNguon_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt(3).Visible = True Then
            RFocus txt(3)
        Else
            RFocus txtchungtu(0)

        End If
    End If
End Sub

Private Sub cboThang_Click()
    Label(26).Caption = ""
    If loaict = 2 Or loaict = 6 Then
        ClearGrid GrdChungtu, GrdChungtu.tag
        MaSoCT = 0
        If vattu.MaSo > 0 Then txtChungtu_LostFocus 2
    End If
    '  danh_sach_chung_tu
End Sub

Private Sub CboThang_DblClick()
    Label(26).Caption = ""
End Sub

Private Sub CboThang_DragDrop(source As Control, X As Single, Y As Single)
    Label(26).Caption = ""
End Sub

Private Sub CboThang_DropDown()
    Label(26).Caption = ""
End Sub

Private Sub CboThang_KeyPress(KeyAscii As Integer)
' Label(26).Casption = ""
    If KeyAscii = 13 Then
        RFocus MedNgay(0)

    End If
End Sub

Private Sub CboThang_LostFocus()
    Label(26).Caption = ""
    Dim st As String
    st = Day(Now)
    If (st > 28) Then
        If (Month("01/0" + CboThang.Text) = 2) Then
            If (Year("01/0" + CboThang.Text) Mod 4 = 0) Then
                st = 29
            Else
                st = 28
            End If
        End If
    End If


    If Day(Now) < 10 Then st = "0" + Day(Now)

    Dim ngay As Date
    If Len(CboThang.Text) <= 5 Then
        ngay = st + "/0" + CboThang.Text
    Else
        If (st > 28) Then
            ngay = "28/" + CboThang.Text
        Else
            ngay = st + "/" + CboThang.Text

        End If
    End If
    If MaSoCT = 0 Then
        MedNgay(0).Text = ngay
        MedNgay(1).Text = ngay
    End If


End Sub

Private Sub CboVV_Click(Index As Integer)
    If Index = 0 Then CboVVClick CboVV(0), CboVV(1)
End Sub

Private Sub Chk_Click()
    txt_LostFocus 0
    CmdPhieu(1).Caption = IIf(Chk.Value = 1, "&2 B¸o gi¸", "&2 Ho¸ ®¬n")
    CmdPhieu(1).tag = Chk.Value
End Sub

Private Sub ChkXT_Click()
    Frame.Visible = (chkXT.Value = 1)
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
    Case 0:
        txtsh(0).Text = FrmTaikhoan.ChonTk(txtsh(0).Text)
        RFocus txtsh(0)
    Case 1:
        If cmd(0).tag = 1 Then
            txtsh(1).Text = FrmKhachHang.ChonKhachHang(txtsh(1).Text)
            RFocus txtsh(1)
        End If
        If cmd(0).tag = 2 Then
            txtsh(1).Text = FrmTP.ChonTP(txtsh(1).Text)
            RFocus txtsh(1)
        End If
        If Left(txtsh(0).Text, 3) = "154" Then
            txtsh(1).Text = FrmTP.ChonTP(txtsh(1).Text)
            RFocus txtsh(1)
        End If
    End Select

End Sub

Private Sub CmdBC_Click()
    LayXuatKho loaict
End Sub

'====================================================================================================
' Ghi dßng ph¸t sinh vµo Grid
'====================================================================================================
Public Sub CmdChitiet_chon()
    If Len(Trim(txtchungtu(1).Text)) <= 0 Then    'kiem tra co ten tai khoan
        Exit Sub
    End If

    Dim no As Double, co As Double, nt As Double, i As Integer, shnt As String, tien As Double, tien1 As Double, m As Long
    Dim n As Double, c As Double, X As Double, ctgg As Boolean, dvt As Long, msk As String, sott As Integer, thang As Integer
    ReDim mct1(1 To 1) As Long
    ReDim tien1x(1 To 1) As Double

    msk = IIf(Left(taikhoan.sohieu, 3) = "007", Mask_2, Mask_0)
    thang = CboThang.ItemData(CboThang.ListIndex)
    nt = Cdbl5(txtchungtu(3).Text)

    If Left(taikhoan.sohieu, 3) <> "007" And pTien = 0 Then
        no = Fix(Cdbl5(txtchungtu(5).Text))
        co = Fix(Cdbl5(txtchungtu(6).Text))
    Else
        no = Cdbl5(txtchungtu(5).Text)
        co = Cdbl5(txtchungtu(6).Text)
    End If

    If taikhoan.MaSo = 0 Or taikhoan.tkcon > 0 Then
        ErrMsg er_SHTaiKhoan1
        RFocus txtchungtu(0)
        Exit Sub
    End If

    If pPQTK > 0 Then
        If Not taikhoan.ChoNhap Then
            MsgBox "Ch­a ®¨ng ký ng­êi sö dông cho tµi kho¶n nµy!", vbExclamation, App.ProductName
            RFocus txtchungtu(0)
            GoTo KT
        End If
    End If

    If (loaict <> 1 And loaict <> 2 And loaict <> 4 And loaict <> 9) And (taikhoan.tk_id = TKVT_ID) And STDetail Then
        MsgBox "Kh«ng nhËp trùc tiÕp ph¸t sinh cho tµi kho¶n vËt t­, hµng hãa. H·y vµo nhËp xuÊt !", vbExclamation, App.ProductName
        RFocus txtchungtu(0)
        Exit Sub
    End If

    If (loaict = 1) And (taikhoan.tk_id = TKVT_ID) And (co <> 0) Then
        MsgBox "Ghi ph¸t sinh nî khi nhËp vËt t­ !", vbExclamation, App.ProductName
        RFocus txtchungtu(5)
        Exit Sub
    End If

    If pDTTP <> 0 And (Left(taikhoan.sohieu, 3) = "621" Or Left(taikhoan.sohieu, 3) = "622" Or Left(taikhoan.sohieu, 3) = "623" Or Left(taikhoan.sohieu, 3) = "627") And (tp.MaSo = 0) Then
        If MsgBox("Kh«ng cã m· sè c«ng tr×nh, tiÕp tôc?", vbCritical + vbYesNo, App.ProductName) = vbNo Then
            RFocus txtchungtu(2)
            Exit Sub
        End If
    End If

    If pDTTP <> 0 And (Left(taikhoan.sohieu, 3) = "621" Or Left(taikhoan.sohieu, 3) = "622" Or Left(taikhoan.sohieu, 3) = "623" Or Left(taikhoan.sohieu, 3) = "627") And (tp.MaSo > 0) Then
        If tp.SoDT(pThangDauKy, thang) > 0 And tp.GiaThanhCK(thang) = 0 And tp.ChiPhiTP(thang) > 0 Then
            If MsgBox("C«ng tr×nh hoÆc s¶n phÈm ®· kÕt chuyÓn gi¸ thµnh, tiÕp tôc?", vbCritical + vbYesNo, App.ProductName) = vbNo Then
                RFocus txtchungtu(5)
                Exit Sub
            End If
        End If
    End If

    If (loaict = 2) And (taikhoan.tk_id = TKVT_ID) And (no <> 0) And (vattu.MaSo > 0) Then
        MsgBox "Ghi ph¸t sinh cã khi xuÊt vËt t­ !", vbExclamation, App.ProductName
        RFocus txtchungtu(6)
        Exit Sub
    End If

    If (loaict = 1 Or loaict = 2) And (taikhoan.tk_id = TKVT_ID And vattu.MaSo = 0) And STDetail Then
        ErrMsg er_SHVattu
        RFocus txtchungtu(2)
        Exit Sub
    End If

    If (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) And ckh.MaSo = 0 And KHDetail Then
        ErrMsg er_SHKhachHang
        RFocus txtchungtu(2)
        Exit Sub
    End If

    If (loaict = 2 Or loaict = 8) And (vattu.MaSo > 0) And (co >= 0) And (nt > txtchungtu(3).tag) And STDetail And Left(taikhoan.sohieu, 4) <> "5113" And Chk.Value = 0 And Me.Visible Then
        'MsgBox "§· xuÊt qu¸ l­îng tån!", vbCritical, App.ProductName
        'Exit Sub
        If IsImport = False Then
            If MsgBox("§· xuÊt qu¸ l­îng tån! TiÕp tôc ?", vbYesNo + vbCritical, App.ProductName) <> vbYes Then
                RFocus txtchungtu(3)
                Exit Sub
            End If
        End If

    End If

    If (loaict = 2) And (taikhoan.tk_id = TKVT_ID) And (txtchungtu(3).tag = 0 And txtchungtu(6).tag = 0) And (nt = 0) And STDetail Then
        If MsgBox("VËt t­ kh«ng cã tån kho! TiÕp tôc ?", vbYesNo + vbCritical, App.ProductName) <> vbYes Then
            RFocus txtchungtu(2)
            Exit Sub
        End If
    End If

    If (loaict = 1 Or loaict = 2) And (taikhoan.tk_id = TKVT_ID) And (vattu.TonMin > 0 Or vattu.TonMax > 0) And STDetail Then
        If (loaict = 1) And (vattu.TonMax > 0) And (txtchungtu(3).tag + nt > vattu.TonMax) Then
            If MsgBox("§· nhËp qu¸ l­îng tån kho tèi ®a cho phÐp! TiÕp tôc ?", vbYesNo + vbCritical, App.ProductName) <> vbYes Then
                RFocus txtchungtu(3)
                Exit Sub
            End If
        End If

        If (loaict = 2) And (vattu.TonMin > 0) And (txtchungtu(3).tag - nt < vattu.TonMin) Then
            If MsgBox("§· xuÊt qu¸ l­îng tån tèi thiÓu cho phÐp! TiÕp tôc ?", vbYesNo + vbCritical, App.ProductName) <> vbYes Then
                RFocus txtchungtu(3)
                Exit Sub
            End If
        End If
    End If

    If pVAT2 > 0 And loaict = 8 And vattu.MaSo > 0 And vBH > 0 And vattu.VAT > 0 And vBH <> vattu.VAT Then
        If MsgBox("MÆt hµng kh«ng cïng thuÕ suÊt VAT! TiÕp tôc ?", vbYesNo + vbCritical, App.ProductName) <> vbYes Then
            RFocus txtchungtu(2)
            Exit Sub
        End If
    End If

    If KhongNhapTS And FADetail And (taikhoan.tk_id = TSCD_ID Or taikhoan.tk_id = KHTSCD_ID) Then
        MsgBox "Kh«ng nhËp trùc tiÕp ph¸t sinh vµo tµi kho¶n TSC§ !", vbInformation, App.ProductName
        RFocus txtchungtu(0)
        Exit Sub
    End If

    If taikhoan.MaNT <> 0 And CboNT(0).ListIndex >= 0 Then
        shnt = CboNT(0).Text
        taikhoan.InitTaikhoanMaSo MaTKNguyenTe(taikhoan.sohieu, CboNT(0).ItemData(CboNT(0).ListIndex))
    Else
        shnt = txtchungtu(2).Text
    End If

    ' KiÓm tra chi tiÕt ®· cã ph¸t sinh ?
    If taikhoan.tk_id <> TSCD_ID And taikhoan.tk_id <> TKDT_ID And taikhoan.tk_id <> GTGTKT_ID And taikhoan.tk_id <> GTGTPN_ID And tp.MaSo = 0 Then
        With GrdChungtu
            For i = 0 To .Rows - 1
                .Row = i
                .col = 8
                If Len(.Text) = 0 Then Exit For
                If CLng5(.Text) = taikhoan.MaSo Then
                    .col = 18
                    If (vattu.MaSo = 0 And ckh.MaSo = 0) And (CInt5(.Text) < 0) And taikhoan.loai <> 6 Then
                        MsgBox "Chi tiÕt ®· cã ph¸t sinh trong chøng tõ !", vbExclamation, App.ProductName
                        RFocus txtchungtu(0)
                        Exit Sub
                    End If
                End If
            Next
        End With
    End If

    If (no = 0 And taikhoan.tk_id <> GTGTKT_ID And taikhoan.tk_id <> TKVT_ID) And (co = 0 And taikhoan.tk_id <> GTGTPN_ID And taikhoan.tk_id <> TKVT_ID And taikhoan.tk_id <> TKDT_ID And Left(taikhoan.sohieu, 5) <> "33332") Then
        If IsImport = False Then
            MsgBox "ThiÕu sè ph¸t sinh !", vbExclamation, App.ProductName
            RFocus txtchungtu(5)
            Exit Sub
        Else
            'Cap nhat status tbimport thanh 2
            With fileImportList(IndexFirst)
                ExecuteSQL5 "UPDATE tbimport SET Status = 2 WHERE ID = " & .id
            End With
        End If
    End If

    If no <> 0 And co <> 0 Then
        MsgBox "ChØ ghi ph¸t sinh nî hoÆc cã !", vbExclamation, App.ProductName
        RFocus txtchungtu(5)
        Exit Sub
    End If

    taikhoan.KtraPhatsinh thang, IIf(no > 0, no, co), IIf(no > 0, -1, 1)
    If co > 0 And ((Left(taikhoan.sohieu, Len(TM)) = TM) Or (Left(taikhoan.sohieu, Len(NH)) = NH)) Then
        taikhoan.SoDuNgay ngay(0), n, c, X
        If n - c < co And IsImport = False Then
            If MsgBox("Chi v­ît sè d­! TiÕp tôc ?", vbYesNo + vbCritical, App.ProductName) <> vbYes Then
                RFocus txtchungtu(6)
                Exit Sub
            End If
        End If
    End If

    If (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) And ckh.MaSo > 0 And Left(ckh.sohieu, 1) = "#" Then
        MsgBox "Kh«ng nhËp ph¸t sinh cho kh¸ch v·ng lai!", vbCritical, App.ProductName
        RFocus txtchungtu(2)
        Exit Sub
    End If

    If ckh.MaSo > 0 And ((no > 0 And taikhoan.tk_id = TKCNKH_ID) Or (co > 0 And taikhoan.tk_id = TKCNPT_ID)) Then
        Label(22).Enabled = True
        txtchungtu(8).Enabled = True
    End If

    If ckh.MaSo > 0 And no > 0 And ckh.DuMax > 0 And taikhoan.tk_id = TKCNKH_ID Then
        ckh.SoDuKH ThangCuoiNamTC, n, c, X
        If n - c + no > ckh.DuMax Then
            If MsgBox("V­ît qu¸ h¹n møc sè d­! TiÕp tôc ?", vbYesNo + vbExclamation, App.ProductName) <> vbYes Then
                RFocus txtchungtu(5)
                Exit Sub
            End If
        End If
    End If

    If (Left(taikhoan.sohieu, Len(TM)) = TM) Then
        If co > 0 Then FThuChi.tag = 1
        If MaSoCT = 0 And hdcount >= 0 And TenTC = "..." Then
            TenTC = HD(0).TenKH
            DiachiTC = HD(0).DiaChiKH
        End If
        FThuChi.GetPhieu TenTC, DiachiTC, ctgoc, MaKHBH
        CmdPhieu(0).Visible = True
    End If

    If (Left(taikhoan.sohieu, Len(NH)) = NH) And KiemTraMaSoThue(frmMain.LbCty(8).Caption, "04") Then
        If co > 0 Then FThuChi.tag = 1
        If MaSoCT = 0 And hdcount >= 0 And TenTC = "..." Then
            TenTC = HD(0).TenKH
            DiachiTC = HD(0).DiaChiKH
        End If
        FThuChi.GetPhieu TenTC, DiachiTC, ctgoc, MaKHBH
        CmdPhieu(3).Visible = True
    End If

    If (Left(taikhoan.sohieu, Len(NH)) = NH Or taikhoan.tk_id2 = CLng(NH)) And co > 0 Then
        FThuChi.tag = 2
        If MaSoCT = 0 And hdcount >= 0 And unc1 = "..." Then
            unc1 = HD(0).TenKH
        End If
        FThuChi.GetPhieu unc1, unc2, unc3, MaKHBH
        CmdPhieu(2).tag = taikhoan.sohieu
        CmdPhieu(2).Visible = True
    End If

    ctgg = CTGiamGia
    If (taikhoan.tk_id = TKVT_ID Or taikhoan.tk_id = TKDT_ID Or taikhoan.tk_id = TSCD_ID Or taikhoan.tk_id = TKGT_ID) Then CmdPhieu(1).Visible = True

    If ((taikhoan.tk_id = GTGTKT_ID And no <> 0) Or ((taikhoan.tk_id = GTGTPN_ID Or taikhoan.tk_id = TTDB_ID) And co <> 0) Or ((taikhoan.tk_id = GTGTKT_ID Or taikhoan.tk_id = GTGTPN_ID) And no = 0 And co = 0) Or (ctgg And no <> 0 And taikhoan.tk_id = GTGTPN_ID)) And pNoiBo = 0 Then
        h.TyLe = CInt5(txtchungtu(2).Text)
        If h.MaSo = 0 Then
            If CmdChitiet.tag >= 0 Then
                h.loai = IIf(taikhoan.tk_id = GTGTKT_ID Or Left(taikhoan.sohieu, 5) = "33312", -1, IIf(taikhoan.tk_id = TTDB_ID, 2, 1))
                GoTo htp
            End If
            If hdcount >= 0 And h.MaKhachHang = 0 Then
                h.MaKhachHang = HD(hdcount).MaKhachHang
                h.MauSo = HD(hdcount).MauSo
                h.KyHieu = HD(hdcount).KyHieu
                h.sohd = HD(hdcount).sohd
                h.NgayPH = HD(hdcount).NgayPH
                h.tygia = HD(hdcount).tygia
            Else
                If taikhoan.tk_id = GTGTPN_ID Then
                    h.MauSo = SelectSQL("SELECT TOP 1 MauSo AS F1 FROM HoaDon WHERE Loai=1 ORDER BY NgayPH DESC")
                    h.KyHieu = SelectSQL("SELECT TOP 1 KyHieu AS F1 FROM HoaDon WHERE Loai=1 ORDER BY NgayPH DESC")
                End If
            End If
            '////////////////////////////////////////////////////////////////


            'Chuyen sang form hoa don
            '    Dim maso_khachhang
            '
            '                'If maso_khachhang = 0 Then Command1_Click  ' Luu khach hang
            '                'lay ma so khach hang lai
            '    maso_khachhang = SelectSQL("SELECT TOP 1 Maso AS F1 FROM KhachHang WHERE sohieu = '" + txtVT(0).Text + "'")
            '   If maso_khachhang >= 0 Then ' neu khach hang nay chua co
            '     ExecuteSQL5 "UPDATE KhachHang SET ten = '" + txtVT(7).Text + "',diachi = '" + txtVT(8).Text + "' where maso = " + CStr(maso_khachhang)
            '    End If
            '                         If Len(Trim(txtVT(0))) > 0 Then
            '                                Command1_Click ' Kiem tra ma khach  hang co trong khong luu thong tin khach hang
            '                                 maso_khachhang = SelectSQL("SELECT TOP 1 Maso AS F1 FROM KhachHang WHERE sohieu = '" + txtVT(0).Text + "'")
            '                        End If
            '                Else
            '                    Dim tt
            '                    ExecuteSQL5 "UPDATE KhachHang SET ten = '" + txtVT(7).Text + "',diachi = '" + txtVT(8).Text + "' where maso = " + CStr(maso_khachhang)
            '
            '                End If
            '
            '                'lay lai ma so moi sinh ra trong luc luu
            '
            '                h.MaKhachHang = maso_khachhang
            'h.MaSo = FrmKhachHang.ChonKhachHang(txtVT(0).Text)

            ' them moi khi sua hoa dom
            '                h.KyHieu = txtVT(1).Text
            '                h.sohd = txt(0).Text 'so hoa don
            '                h.MatHang = txt(1).Text
            '                h.DiaChiKH = txtVT(8).Text
            ' --------------------
            ' vua them vao sau
            If hdcount <= 0 Then
                h.sohd = txt(0).Text    'so hoa don
            End If
            h.KyHieu = txtVT(1).Text
            h.MatHang = txt(1).Text
            h.DiaChiKH = txtVT(8).Text
            ' ket thuc them vao sau

            '                T(0).Text = ckh.SoHieu
            '        T(7).Text = ckh.Ten
            '        T(8).Text = ckh.DiaChi
            '        T(9).Text = ckh.mst
            '    FVAT.T(0).Text = txtVT(0).Text
            '                FVAT.T(7).Text = txtVT(7).Text
            '                FVAT.T(8).Text = txtVT(8).Text
            '                FVAT.T(9).Text = txtVT(9).Text
            If h.TyLe > 0 Then
                If taikhoan.tk_id = GTGTKT_ID Then tien = (no - co) Else tien = (co - no)
                If taikhoan.tk_id = GTGTKT_ID Or taikhoan.tk_id = GTGTPN_ID Then
                    If h.TyLe > 0 And h.TyLe < 5 Then
                        tien = 100 * tien / h.TyLe - tien
                    Else
                        tien = 100 * tien / h.TyLe
                    End If
                    tien = RoundMoney(tien)
                    tien1 = GiaTriTruocThue
                    If (pTien = 0 And Abs(tien - tien1) < 1000) Or (pTien > 0 And Abs(tien - tien1) < 0.1) Then
                        h.ThanhTien = GiaTriTruocThue
                    Else
                        h.ThanhTien = tien
                    End If
                Else
                    tien1 = GiaTriTruocThue
                    h.ThanhTien = tien1 - tien
                End If
            Else
                h.ThanhTien = GiaTriTruocThue

            End If
            h.NgayPH = ngay(0)
            h.MaSo = hdcount + 1
            If taikhoan.tk_id = TTDB_ID Then FVAT.tag = tien1
            h.loai = IIf(taikhoan.tk_id = GTGTKT_ID Or Left(taikhoan.sohieu, 5) = "33312", -1, IIf(taikhoan.tk_id = TTDB_ID, 2, 1))
            h.NK = IIf(h.loai = -1 And (CoPSTK("33312") Or Left(taikhoan.sohieu, 5) = "33312"), 1, 0)
            h.ts = IIf(h.loai = -1 And loaict = 9, 1, 0)
            If pTygia > 0 Then h.tygia = Cdbl5(txtchungtu(7).Text)
htp:
            'NEU NHAN SHIFT + ? THI CHO HIEN

            '                FVAT.T(0).Text = txtVT(0).Text
            '                FVAT.T(7).Text = txtVT(7).Text
            '                FVAT.T(8).Text = txtVT(8).Text
            '                FVAT.T(9).Text = txtVT(9).Text
            'If cho_hien_vat Then
            '                ckh.SoHieu = txtVT(0).Text
            '                ckh.mst = txtVT(9).Text
            '                ckh.Ten = txtVT(7).Text
            '                ckh.DiaChi = txtVT(8).Text
            If txtVT(0).Visible = True Then    ' neu co khach hang hien moi cho nhap
                If h.MaKhachHang > 0 Then    'neu da co khach hang do roi(da co hoa don)
                    FVAT.GetPhieu taikhoan.tk_id = TTDB_ID    ' mo form Fvat da duoc lap day thong tin
                Else    'chua co hoa don,
                    FVAT.GetPhieu True
                    If cho_hien_vat = False Then FVAT.Command_Click

                End If
            End If
            FrmChungtu.cho_hien_vat = False    ' tra ve trang thai khong cho hien
            If h.MaKhachHang = 0 And KHDetail Then Exit Sub
            If ctgg And h.ThanhTien > 0 And (Not CoPSTK("511", 1)) Then h.ThanhTien = -h.ThanhTien
            hdcount = hdcount + 1
            ReDim Preserve HD(0 To hdcount) As tpHoaDon

            ' h.MaKhachHang = 9999999
            ' cho nay co the dua truc tiep thong tin de luu hoa don

            CopyHD h, HD(hdcount)
            h.MaSo = 0
            i = hdcount
        Else
            FVAT.GetPhieu taikhoan.tk_id = TTDB_ID
            If KHDetail And h.MaKhachHang = 0 Then Exit Sub
            i = 0

            If hdcount >= 0 Then
                Do While HD(i).MaSo <> h.MaSo
                    i = i + 1
                    If i >= hdcount Then Exit Do
                Loop
            Else
                i = 0
            End If

            If i >= hdcount Then
                hdcount = hdcount + 1
                ReDim Preserve HD(0 To hdcount) As tpHoaDon
            End If

            CopyHD h, HD(i)
            HD(i).loai = IIf(taikhoan.tk_id = GTGTKT_ID Or Left(taikhoan.sohieu, 5) = "33312", -1, IIf(taikhoan.tk_id = TTDB_ID, 2, 1))
        End If
        If h.MaKhachHang <> 0 Then    ' them vao de ghi thong tin ra ma hinh chung tu
            ckh.InitKhachHangMaSo h.MaKhachHang
            TenBH = ckh.Ten
            txtVT(7).Text = ckh.Ten
            DiaChiBH = ckh.DiaChi
            txtVT(8).Text = ckh.DiaChi
            txtVT(9).Text = ckh.mst
            MSTBH = ckh.mst
            txtVT(1).Text = h.KyHieu
            txtVT(0).Text = ckh.sohieu
            MaKHBH = h.MaKhachHang

            '   ckh.InitKhachHangMaSo 0
        End If
        XoaHD
    Else
        i = -1
    End If

    If taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID Then
        m = ckh.MaSo
    Else
        m = 0
    End If

    If (loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8) And vattu.Dvt2 > 0 And vattu.MaSo > 0 Then
        dvt = CboNT(2).ItemData(CboNT(2).ListIndex)
    Else
        dvt = 0
    End If

    If loaict = 2 And taikhoan.loai > 0 And (OutCost = 2 Or OutCost = 3) And taikhoan.tk_id = TKVT_ID And vattu.MaSo > 0 Then
        Dim luongx() As Double, tienx() As Double, id() As Long, cx As Integer, tienx2() As Double

        If OutCost = 2 Then
            cx = GiaXuatKhoFIFO(CboNguon(1).ItemData(CboNguon(1).ListIndex), taikhoan.MaSo, vattu.MaSo, nt, luongx, tienx, id, tienx2)
        Else
            cx = GiaXuatKhoLIFO(CboNguon(1).ItemData(CboNguon(1).ListIndex), taikhoan.MaSo, vattu.MaSo, nt, luongx, tienx, id, tienx2)
        End If

        For i = 1 To cx
            If luongx(i) <> 0 Then X = tienx(i) / luongx(i) Else X = 0
            GrdChungtu.AddItem "" + Chr(9) + txtchungtu(0).Text + Chr(9) + txtchungtu(1).Text + Chr(9) + shnt _
                             + Chr(9) + Format(luongx(i), Mask_2) + Chr(9) + IIf(luongx(i) <> 0, Format(X, Mask_2), "") + Chr(9) + "" + Chr(9) + Format(tienx(i), msk) + Chr(9) _
                             + CStr(taikhoan.MaSo) + Chr(9) + CStr(vattu.MaSo) + Chr(9) + IIf(taikhoan.loai > 0, "0", "1") _
                             + Chr(9) + CStr(taikhoan.MaTC) + Chr(9) + CStr(id(i)) + Chr(9) + CStr(taikhoan.tk_id) _
                             + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + "-1" + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) _
                             + Format(tienx2(i), Mask_2), NewRowIndex(GrdChungtu, 1)    'NewRowIndex(GrdChungtu, 1)
        Next
        Erase luongx
        Erase tienx
        Erase tienx2
        Erase id

        GoTo KT
    End If

    If pCongNoHD > 0 Then
        If (taikhoan.tk_id = TKCNKH_ID And taikhoan.kieu < 0) Or (taikhoan.tk_id = TKCNPT_ID And taikhoan.kieu > 0) Then
            sott = 0
            If co <> 0 And taikhoan.tk_id = TKCNKH_ID Then
                FDsHD.tag = ckh.MaSo
                FDsHD.ThanhToanDichDanh taikhoan.MaSo, ckh.Ten, co, 0, mct1, tien1x(), sott
            End If
            If no <> 0 And taikhoan.tk_id = TKCNPT_ID Then
                FDsHD.tag = ckh.MaSo
                FDsHD.ThanhToanDichDanh taikhoan.MaSo, ckh.Ten, no, 1, mct1(), tien1x(), sott
            End If
            If sott = 1 Then MaNhap = mct1(1)
            If sott > 1 Then
                For i = 1 To sott
                    GrdChungtu.AddItem "" + Chr(9) + txtchungtu(0).Text + Chr(9) + txtchungtu(1).Text + Chr(9) + shnt _
                                     + Chr(9) + IIf(nt <> 0, Format(nt, Mask_2), "") + Chr(9) + IIf(nt <> 0, txtchungtu(4).Text, "") + Chr(9) + IIf(taikhoan.tk_id = TKCNPT_ID, Format(tien1x(i), msk), "") + Chr(9) + IIf(taikhoan.tk_id = TKCNKH_ID, Format(tien1x(i), msk), "") + Chr(9) _
                                     + CStr(taikhoan.MaSo) + Chr(9) + CStr(0) + Chr(9) + IIf(taikhoan.loai > 0, "0", "1") _
                                     + Chr(9) + CStr(taikhoan.MaTC) + Chr(9) + CStr(mct1(i)) + Chr(9) + CStr(taikhoan.tk_id) _
                                     + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + CStr(IIf(taikhoan.tk_id = TKCNPT_ID, m, 0)) + Chr(9) + CStr(-1) + Chr(9) _
                                     + "" + Chr(9) + CStr(IIf(taikhoan.tk_id = TKCNKH_ID, m, 0)) + Chr(9) + CStr(0) + Chr(9) + "" + Chr(9) + CStr(0), NewRowIndex(GrdChungtu, 1)
                Next
                GoTo a
            End If
        End If
    End If

    'x = 0
    'If loaict = 8 And pGiaUSD > 0 And vattu.MaSo > 0 And taikhoan.tk_id = TKDT_ID Then
    '    x = Cdbl5(txtchungtu(3).Text) * Cdbl5(txtchungtu(4).Text)
    'End If
    If MaNhap = 0 And pFunction = 10 And pCT_ID > 0 Then
        MaNhap = pCT_ID
    End If

    '        Dim SoLo, HanDung
    '        SoLo = ""
    '        HanDung = ""
    '        If nt <> 0 Then
    '            frmSoLo.Show 1
    '            SoLo = frmSoLo.txtsolo.Text
    '            HanDung = frmSoLo.txtngaynhap.Text
    '
    '        End If
    '
    '       GrdChungtu.AddItem "" + Chr(9) + txtchungtu(0).Text + Chr(9) + txtchungtu(1).Text + Chr(9) + shnt _
            '            + Chr(9) + IIf(nt <> 0, Format(nt, Mask_2), "") + Chr(9) + IIf(nt <> 0, txtchungtu(4).Text, "") + Chr(9) + Format(no, msk) + Chr(9) + Format(co, msk) + Chr(9) _
            '            + CStr(taikhoan.MaSo) + Chr(9) + CStr(vattu.MaSo) + Chr(9) + IIf(taikhoan.loai > 0, "0", "1") _
            '            + Chr(9) + CStr(taikhoan.MaTC) + Chr(9) + CStr(MaNhap) + Chr(9) + CStr(taikhoan.tk_id) _
            '            + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + DoiDau(IIf(no <> 0, m, 0)) + Chr(9) + CStr(i) + Chr(9) _
            '            + "" + Chr(9) + DoiDau(IIf(co <> 0, m, 0)) + Chr(9) + CStr(tp.MaSo) + Chr(9) + "" + Chr(9) + CStr(dvt) + Chr(9) _
            '            + Format(txtchungtu(11).Text, Mask_2) + Chr(9) + Format(txtchungtu(9).Text, IIf(Cdbl5(txtchungtu(9).Text) * 100 Mod 100 <> 0, Mask_2, Mask_0)) + Chr(9) + Format(txtchungtu(10).Text, Mask_0) + Chr(9) + SoLo + Chr(9) + HanDung, IIf(CmdChitiet.tag < 0, NewRowIndex(GrdChungtu, 1), CmdChitiet.tag)
    If SelectSQL("SELECT banthuoc as f1 from license ") = 1 Then
        If nt <> 0 Then
            If OptLoai(1).Value = True Then frmSoLo.Show 1
        Else
            frmSoLo.txtsolo.Text = ""
            frmSoLo.txtngaynhap.Text = "01/01/90"
        End If
    End If
    GrdChungtu.AddItem "" + Chr(9) + txtchungtu(0).Text + Chr(9) + txtchungtu(1).Text + Chr(9) + shnt _
                     + Chr(9) + IIf(nt <> 0, Format(nt, Mask_2), "") + Chr(9) + IIf(nt <> 0, txtchungtu(4).Text, "") + Chr(9) + Format(no, msk) + Chr(9) + Format(co, msk) + Chr(9) _
                     + CStr(taikhoan.MaSo) + Chr(9) + CStr(vattu.MaSo) + Chr(9) + IIf(taikhoan.loai > 0, "0", "1") _
                     + Chr(9) + CStr(taikhoan.MaTC) + Chr(9) + CStr(MaNhap) + Chr(9) + CStr(taikhoan.tk_id) _
                     + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + DoiDau(IIf(no <> 0, m, 0)) + Chr(9) + CStr(i) + Chr(9) _
                     + "" + Chr(9) + DoiDau(IIf(co <> 0, m, 0)) + Chr(9) + CStr(tp.MaSo) + Chr(9) + "" + Chr(9) + CStr(dvt) + Chr(9) _
                     + Format(txtchungtu(11).Text, Mask_2) + Chr(9) + Format(txtchungtu(9).Text, IIf(Cdbl5(txtchungtu(9).Text) * 100 Mod 100 <> 0, Mask_2, Mask_0)) + Chr(9) + Format(txtchungtu(10).Text, Mask_0) + Chr(9) + frmSoLo.txtsolo.Text + Chr(9) + frmSoLo.txtngaynhap.Text, IIf(CmdChitiet.tag < 0, NewRowIndex(GrdChungtu, 1), CmdChitiet.tag)


a:
    GrdChungtu.Row = GrdChungtu.Rows - 1
    GrdChungtu.col = 1
    If Len(GrdChungtu.Text) = 0 Then GrdChungtu.RemoveItem GrdChungtu.Row

KT:
    GrdChungtu.Row = 0
    CmdChitiet.tag = -1
    If (loaict = 1 Or loaict = 2) Then vattu.InitVattuMaSo 0
    If loaict = 2 Or loaict = 8 Then
        xddu = SetDoiUng(1)
        If Not xddu Then xddu = SetDoiUng
    Else
        xddu = SetDoiUng
        If Not xddu Then xddu = SetDoiUng(1)
    End If
    If ckh.MaSo > 0 Then
        DiaChiNX = ckh.DiaChi
        TenBH = ckh.Ten
        DiaChiBH = ckh.DiaChi
        MSTBH = ckh.mst
        MaKHBH = ckh.MaSo
    End If
    If pVAT2 > 0 And loaict = 8 And vattu.MaSo > 0 And vattu.VAT > 0 Then vBH = vattu.VAT
    NhapDongMoi taikhoan.sohieu

    Erase mct1
    Erase tien1x
End Sub

Public Sub CmdChitiet_Click()
    cho_hien_vat = True
    If SelectSQL("SELECT banthuoc as f1 from license ") = 1 Then
        If (OptLoai(2).Value = True Or OptLoai(8).Value = True) And (Mid(txtchungtu(0).Text, 1, 3) = "156" Or Mid(txtchungtu(0).Text, 1, 3) = "154" Or Mid(txtchungtu(0).Text, 1, 3) = "155" Or Mid(txtchungtu(0).Text, 1, 3) = "511") Then
            Dim bang_sd As Recordset
            Dim ma As String
            Dim st As String
            Dim n
            ma = txtchungtu(2).Text
            txttinh_gia_ban.Caption = txtchungtu(4).Text

            If IsMissing(LO_XXXX) Then LO_XXXX = ""
            If Len(LO_XXXX) > 0 Then
                st = "select * from DanhSachVatTu where sohieu = '" + ma + "' And solo='" + LO_XXXX + "' and conlai > 0 order by handung asc"
            Else
                st = "select * from DanhSachVatTu where sohieu = '" + ma + "' and conlai > 0 order by handung asc"
            End If

            Set bang_sd = DBKetoan.OpenRecordset(st, dbOpenSnapshot)
            If bang_sd.recordCount > 0 Then
                Dim so_luong, i As Integer
                so_luong = CInt(txtchungtu(3).Text)
                n = bang_sd.recordCount
                If n = 1 Then n = n - 1
                For i = 0 To n
                    txtChungtu_KeyPress 0, 13
                    txtChungtu_LostFocus (0)
                    txtchungtu(2).Text = ma
                    '  txtChungtu_KeyPress 2, 13
                    txtChungtu_LostFocus (2)

                    frmSoLo.txtsolo.Text = IIf(IsNull(bang_sd!solo), "", bang_sd!solo)
                    frmSoLo.txtngaynhap.Text = IIf(IsNull(bang_sd!handung), "01/01/11", bang_sd!handung)
                    If so_luong > bang_sd!conlai Then
                        txtchungtu(3).Text = CStr(bang_sd!conlai)
                        txtChungtu_KeyPress 3, 13
                        txtChungtu_LostFocus (3)
                        so_luong = so_luong - bang_sd!conlai
                        CmdChitiet_chon
                        ' lam quanh thu 2
                        bang_sd.MoveNext
                    Else
                        txtchungtu(3).Text = CStr(so_luong)
                        txtChungtu_KeyPress 3, 13
                        txtChungtu_LostFocus (3)
                        frmSoLo.txtsolo.Text = bang_sd!solo
                        frmSoLo.txtngaynhap.Text = bang_sd!handung
                        CmdChitiet_chon
                        txttinh_gia_ban.Caption = "0"
                        Exit Sub
                        bang_sd.MoveNext
                    End If

                Next
            Else
                MsgBox "B¹n ph¶i nhËp sè l« vµ h¹n dïng cña mÆt hµng nµy!"
            End If

        Else
            If Len(Trim(txtchungtu(1).Text)) > 0 Then CmdChitiet_chon
        End If
    Else
        If Len(Trim(txtchungtu(1).Text)) > 0 Then CmdChitiet_chon
    End If
    'ap dung cho binh thuong
    'If Len(Trim(txtchungtu(1).Text)) > 0 Then CmdChitiet_chon

    txttinh_gia_ban.Caption = "0"

End Sub
Sub in_hoa_don_tong_hop(Index As Integer, in_hd As Integer)
    frmMain.Rpt.Reset
    frmMain.Rpt.WindowState = crptMaximized
    Dim sotien As String, i As Integer, k As Integer, xxx As String, sodu As Integer, v As Double, lp As Integer, ms As Long, mv As String
    Dim tien As Double, ttien As Double, luong As Double, tkno As String, TkCo As String, TK As New ClsTaikhoan, tiennt As Double
    Dim ts As clsTaiSan, HTTT As String, tl As Integer, thue As Double, v338 As Double, v521 As Double, X As Double, shtk As String, vt As New ClsVattu
    Dim dn As Double, DC As Double, dnt As Double, CK As Double, somh As Integer, tp As New Cls154, lanin As Integer, stt As Integer, loaitien As String
    Dim chophep_in As Integer
    chophep_in = 0
    SetRptInfo

    Select Case Index
    Case 0, 3:
        shtk = IIf(Index = 0, TM, NH)
        lp = LoaiPhieuThuChi(shtk)
        tiennt = 0

        If lp < 0 Then

            frmMain.Rpt.ReportFileName = IIf(Index = 0, "PHIEUTHU.RPT", "THUNH.RPT")

            With GrdChungtu
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 1
                    xxx = .Text
                    If Len(xxx) = 0 Then Exit For
                    .col = 6
                    v = Cdbl5(.Text)
                    If v <> 0 Then
                        If Left(xxx, Len(shtk)) = shtk Then ttien = ttien + v
                        .col = 4
                        X = Cdbl5(.Text)
                        If Cdbl5(.Text) <> 0 Then
                            tiennt = tiennt + X
                            .col = 3
                            loaitien = .Text
                        End If
                        frmMain.Rpt.Formulas(80 + stt) = "TKNo" + CStr(stt) + "='" + xxx + "'"
                        frmMain.Rpt.Formulas(90 + stt) = "PSNo" + CStr(stt) + "=" + DoiDau(v)
                        stt = stt + 1
                    Else
                        .col = 7
                        tien = Cdbl5(.Text)
                        If tien <> 0 Then
                            .col = 1
                            xxx = .Text
                            For k = 0 To i - 1
                                .Row = k
                                If .Text = xxx Then
                                    .col = 19
                                    If IsNumeric(.Text) Then
                                        sotien = frmMain.Rpt.Formulas(7 + 2 * CInt5(.Text))
                                        frmMain.Rpt.Formulas(7 + 2 * CInt5(.Text)) = sotien + "+" + DoiDau(tien)
                                        GoTo A1
                                    Else
                                        GoTo B1
                                    End If
                                End If
                            Next
B1:
                            .Row = i
                            sodu = sodu + 1
                            frmMain.Rpt.Formulas(6 + 2 * sodu) = "TKCo" + CStr(sodu) + "='" + xxx + "'"
                            frmMain.Rpt.Formulas(7 + 2 * sodu) = "PSCo" + CStr(sodu) + "=" + DoiDau(tien)
                            .col = 19
                            .Text = CStr(sodu)
A1:
                        End If
                    End If
                Next
            End With
        Else
C1:
            frmMain.Rpt.ReportFileName = IIf(Index = 0, "PHIEUCHI.RPT", "CHINH.RPT")
            ttien = 0
            With GrdChungtu
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 1
                    xxx = .Text
                    If Len(xxx) = 0 Then Exit For
                    .col = 7
                    X = Cdbl5(.Text)
                    If X <> 0 Then
                        If Left(xxx, Len(shtk)) = shtk Then ttien = ttien + X
                        frmMain.Rpt.Formulas(80 + stt) = "TKCo" + CStr(stt) + "='" + xxx + "'"
                        frmMain.Rpt.Formulas(90 + stt) = "PSCo" + CStr(stt) + "=" + DoiDau(X)
                        stt = stt + 1
                        .col = 4
                        X = Cdbl5(.Text)
                        If Cdbl5(.Text) <> 0 Then
                            tiennt = tiennt + X
                            .col = 3
                            loaitien = .Text
                        End If
                    Else
                        .col = 6
                        tien = Cdbl5(.Text)
                        If tien <> 0 Then
                            .col = 1
                            xxx = .Text
                            For k = 0 To i - 1
                                .Row = k
                                If .Text = xxx Then
                                    .col = 19
                                    If IsNumeric(.Text) Then
                                        sotien = frmMain.Rpt.Formulas(7 + 2 * CInt5(.Text))
                                        frmMain.Rpt.Formulas(7 + 2 * CInt5(.Text)) = sotien + "+" + DoiDau(tien)
                                    Else
                                        GoTo c
                                    End If
                                    GoTo B
                                End If
                            Next
c:
                            .Row = i
                            sodu = sodu + 1
                            frmMain.Rpt.Formulas(6 + 2 * sodu) = "TKNo" + CStr(sodu) + "='" + xxx + "'"
                            frmMain.Rpt.Formulas(7 + 2 * sodu) = "PSNo" + CStr(sodu) + "=" + DoiDau(tien)
                            .col = 19
                            .Text = CStr(sodu)
B:
                        End If
                    End If
                Next
            End With
        End If
        If tiennt > 0 Then
            sotien = ToVNText(Fix(Abs(tiennt))) + IIf(UCase(loaitien) = "USD", " dollars ", " " + loaitien + " ")
            If tiennt - Fix(tiennt) > 0 Then sotien = sotien + " vµ " + ToVNText(Fix(0.5 + 100 * (tiennt - Fix(tiennt)))) + " cents"
            frmMain.Rpt.Formulas(50) = "LoaiT='" + loaitien + "'"
            frmMain.Rpt.Formulas(70) = "x='" + Format(tiennt, Mask_2) + "'"
        Else
            sotien = ToVNText(Abs(ttien)) + " ®ång ch½n"
            frmMain.Rpt.Formulas(70) = "x='" + Format(ttien, Mask_0) + "'"
        End If

        ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
        For i = 1 To IIf(lp < 0, Ppthu, Ppchi)
            ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten) SELECT " + CStr(i) + ",'" + CStr(i) + "',DiaChi FROM License"
        Next

        frmMain.Rpt.Formulas(3) = "SoPhieu='" + LaySH(txt(0).Text, 1) + "'"
        frmMain.Rpt.Formulas(4) = "DiaChi='" + DiachiTC + "'"
        frmMain.Rpt.Formulas(5) = "CTGoc='" + ctgoc + "'"
        frmMain.Rpt.Formulas(41) = "DiaChiDN='" + frmMain.LbCty(2).Caption + "'"
        frmMain.Rpt.Formulas(42) = "TelDN='" + frmMain.LbCty(3).Caption + "'"
        frmMain.Rpt.Formulas(44) = "Ngay='Ngµy " + Format(ngay(1), Mask_DR) + "'"
        frmMain.Rpt.Formulas(45) = "BangChu='" + sotien + "'"
        frmMain.Rpt.Formulas(46) = "TenNV='" + TenTC + "'"
        frmMain.Rpt.Formulas(47) = "LyDo='" + txt(1).Text + "'"

        i = LaySohieuDoiTuong(xxx)
        If i > 1 Then
            frmMain.Rpt.Formulas(48) = "TenKH='" + xxx + "'"
        Else
            If i < 1 Then
                If hdcount >= 0 Then
                    ms = HD(hdcount).MaKhachHang
                Else
                    ms = MaKHBH
                    If ms = 0 Then ms = LayMaKH(IIf(lp = 0, -1, 1))
                End If
                If ms > 0 Then
                    frmMain.Rpt.Formulas(48) = "TenKH='" + TenKH(xxx, ms) + "'"
                    frmMain.Rpt.Formulas(49) = "MaSoKH='" + IIf(Left(xxx, 1) = "#", "...", xxx) + "'"
                End If
            End If
            If i = 1 Then
                frmMain.Rpt.Formulas(48) = "TenKH='" + TenKH(xxx, 0) + "'"
                frmMain.Rpt.Formulas(49) = "MaSoKH='" + xxx + "'"
            End If
        End If
    Case 1:
        v = 0
        sodu = 0
        tl = -1
        thue = 0
        ExecuteSQL5 "DELETE * FROM PhieuNX"
        Dim solo As String, handung As String
        With GrdChungtu
            For i = 0 To .Rows - 1
                .Row = i
                .col = 1
                If Len(.Text) = 0 Then Exit For
                TK.InitTaikhoanSohieu .Text

                If (loaict = 9 Or loaict = 10) And TK.tk_id = TSCD_ID Then
                    .col = IIf(loaict = 9, 6, 7)
                    tien = Cdbl5(.Text)
                    ttien = ttien + tien
                    Set ts = New clsTaiSan
                    If pMaTaiSan > 0 Then
                        ts.ChiDinh pMaTaiSan, 0
                    Else
                        .col = 3
                        ts.ChiDinh SoHieu2MaSo(.Text, "TaiSan"), 0
                    End If
                    ExecuteSQL5 "INSERT INTO PhieuNX (MaSo, SoCT,DienGiaiCT,SoHieu,DienGiai,SoLuong,ThanhTien) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhieuNX") + 1) + ",'" + LaySH(txt(0).Text, 2) _
                              + "','" + txt(1).Text + "','" + ts.sohieu + "','" + ts.Ten + "',1," + DoiDau(tien) + ")"
                    Set ts = Nothing
                End If

                If TK.tk_id = TKDT_ID Or TK.tk_id = TKVT_ID Then
                    If loaict = 1 Or TK.tk_id = TKGT_ID Then
                        If InStr(tkno, TK.sohieu) = 0 Then tkno = tkno + IIf(Len(tkno) > 0, ", ", "") + TK.sohieu
                    Else
                        If InStr(TkCo, TK.sohieu) = 0 Then TkCo = TkCo + IIf(Len(TkCo) > 0, ", ", "") + TK.sohieu
                    End If

                    .col = 3
                    If Len(.Text) > 0 Then
                        sodu = 1
                        If TK.tk_id2 = TKDT_ID Then
                            tp.InitTPSohieu .Text
                            mv = tp.DonVi
                        Else
                            vt.InitVattuSohieu .Text
                            .col = 23
                            If Not KtraDVT(vt.MaSo, CLng5(.Text), mv) And (loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8) Then
                                mv = vt.DonVi
                            End If
                            If vt.MaSo > 0 Then XDTyLeQD vt.MaSo
                        End If
                        .col = 4
                        luong = Cdbl5(.Text)
                        .col = 24
                        tiennt = Cdbl5(.Text)
                        .col = 26
                        X = Cdbl5(.Text)
                        CK = CK + X
                        .col = IIf(loaict = 1 Or TK.tk_id = TKGT_ID, 6, 7)
                        tien = Cdbl5(.Text)
                        If loaict = 8 Then tien = Abs(tien)
                        ttien = ttien + tien
                        .col = 5
                        tiennt = Cdbl5(.Text)
                        If SelectSQL("select banthuoc as f1 from License") = 1 Then
                            .col = 27
                            solo = .Text
                            .col = 28
                            If (Len(solo) > 0) Then
                                solo = " " + solo + "\" + .Text
                            Else
                                handung = ""
                                solo = ""
                            End If
                        End If
                        '                                  If (SelectSQL("SELECT count(*) as f1 from PhieuNX Where soct = '" + LaySH(txt(0).Text, 2) + "' and SoHieu = '" + IIf(vt.MaSo > 0, vt.sohieu, tp.sohieu) + "' ") > 0) Then
                        '                                          ExecuteSQL5 "update PhieuNX set  " _
                                                                   '                                            & "SoLuong = SoLuong  + " + DoiDau(luong) + "," _
                                                                   '                                            & "ThanhTien = ThanhTien + " + DoiDau(tien) + "," _
                                                                   '                                            & "DonGia = DonGia + " + DoiDau(tiennt) + "," _
                                                                   '                                            & "Thue = Thue + " + DoiDau(thue) + "," _
                                                                   '                                            & "ThanhTien2 = ThanhTien2 + " + DoiDau(tiennt) + "," _
                                                                   '                                            & "CK = CK +   " + DoiDau(X) + "" _
                                                                   '                                            & " Where soct = '" + LaySH(txt(0).Text, 2) + "'" _
                                                                   '                                            & " and SoHieu = '" + IIf(vt.MaSo > 0, vt.sohieu, tp.sohieu) + "'" _
                                                                   '                                            & " and Diengiai = '" + IIf(vt.MaSo > 0, vt.TenVattu + solo, tp.TenVattu + solo) + "'"
                        '                               Else

                        ExecuteSQL5 "INSERT INTO PhieuNX (MaSo, SoCT,DienGiaiCT,SoHieu,DienGiai,SoLuong,ThanhTien,DVT,DonGia,TyLe,Thue,ThanhTien2,CK) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhieuNX") + 1) + ",'" + LaySH(txt(0).Text, 2) + "','" + txt(1).Text _
                                  + "','" + IIf(vt.MaSo > 0, vt.sohieu, tp.sohieu) + "','" + IIf(vt.MaSo > 0, vt.TenVattu + solo, tp.TenVattu + solo) + "'," + DoiDau(luong) + "," + DoiDau(tien) + ",'" + mv + "'," + DoiDau(tiennt) + "," + CStr(tl) + "," + DoiDau(thue) + "," + DoiDau(tiennt) + "," + DoiDau(X) + ")"
                        '                                End If

                        '                                ExecuteSQL5 "INSERT INTO PhieuNX (MaSo, SoCT,DienGiaiCT,SoHieu,DienGiai,SoLuong,ThanhTien,DVT,DonGia,TyLe,Thue,ThanhTien2,CK) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhieuNX") + 1) + ",'" + LaySH(txt(0).Text, 2) + "','" + txt(1).Text _
                                                         '                                        + "','" + IIf(vt.MaSo > 0, vt.sohieu, tp.sohieu) + "','" + IIf(vt.MaSo > 0, vt.TenVattu, tp.TenVattu) + "'," + DoiDau(luong) + "," + DoiDau(tien) + ",'" + mv + "'," + DoiDau(tiennt) + "," + CStr(tl) + "," + DoiDau(thue) + "," + DoiDau(tiennt) + "," + DoiDau(X) + ")"

                        somh = somh + 1
                    Else
                        If TK.tk_id = TKDT_ID Then
                            .col = 7
                            tiennt = tiennt + Cdbl5(.Text)
                        End If
                    End If
                Else
                    If loaict <> 8 Then
                        .col = 6
                        luong = Cdbl5(.Text)
                        If luong <> 0 And TK.tk_id = GTGTKT_ID Then v = v + luong
                        If luong <> 0 And InStr(tkno, TK.sohieu) = 0 Then tkno = tkno + IIf(Len(tkno) > 0, ", ", "") + TK.sohieu
                        .col = 7
                        luong = Cdbl5(.Text)
                        If luong <> 0 And TK.tk_id = GTGTPN_ID Then v = v + luong
                        If luong <> 0 And InStr(TkCo, TK.sohieu) = 0 Then TkCo = TkCo + IIf(Len(TkCo) > 0, ", ", "") + TK.sohieu
                        If TK.tk_id <> TKVT_ID And (Left(TK.sohieu, 3) = "138" Or Left(TK.sohieu, 3) = "338") Then
                            .col = 6
                            v338 = v338 + Cdbl5(.Text)
                        End If
                    Else
                        If TK.tk_id = GTGTPN_ID Then
                            .col = 19
                            TkCo = .Text
                            .col = 7
                            thue = Abs(Cdbl5(.Text))
                            .col = 6
                            thue = thue + Cdbl5(.Text)
                            v = v + thue
                            .col = 3
                            tl = CInt5(.Text)
                            .col = 18
                            k = CInt5(.Text)
                            If k >= 0 And k <= hdcount Then
                                If HD(CInt5(.Text)).KCT = 1 Then tl = -1
                            End If
                        End If
                        If TK.tk_id <> TKDT_ID And (Left(TK.sohieu, 3) = "138" Or Left(TK.sohieu, 3) = "338") Then
                            .col = 7
                            v338 = v338 + Cdbl5(.Text)
                        End If
                    End If
                End If
            Next
        End With
        Select Case loaict
        Case 1, 2:
            If P_1 = 0 Then
                TenNX = FrmGetStr.GetString("Tªn ng­êi giao nhËn:", "PhiÕu nhËp xuÊt", IIf(MaSoCT > 0, TenNX, ""))
                DiaChiNX = FrmGetStr.GetString("§Þa chØ ng­êi giao nhËn:", "PhiÕu nhËp xuÊt", IIf(MaSoCT > 0, DiaChiNX, ""))
            End If
            frmMain.Rpt.Formulas(12) = "NCC='" + IIf(loaict = 1, "§¬n vÞ cung cÊp:", "§¬n vÞ nhËn hµng:") + "'"
            If hdcount >= 0 Then
                If loaict = 1 Then
                    If hdcount > 0 Then
                        frmMain.Rpt.Formulas(13) = "TyLe='ThuÕ GTGT:'"
                    Else
                        frmMain.Rpt.Formulas(13) = "TyLe='ThuÕ GTGT (" + CStr(HD(hdcount).TyLe) + "%):'"
                    End If
                    frmMain.Rpt.Formulas(18) = "HoaDon='" + ABCtoVNI("Ho¸ ®¬n : ") + HD(hdcount).MauSo + ABCtoVNI(" - Sè : ") + HD(hdcount).sohd + ABCtoVNI(" - Sè : ") + HD(hdcount).KyHieu + ABCtoVNI(" - Ngµy : ") + Format(HD(hdcount).NgayPH, Mask_DR) + "'"
                End If
                ms = HD(hdcount).MaKhachHang
            Else
                ms = LayMaKH(IIf(loaict = 1, 1, -1))
            End If
            If ms > 0 Then
                frmMain.Rpt.Formulas(14) = "TenKH='" + TenKH(xxx, ms, mv) + "'"
                frmMain.Rpt.Formulas(15) = "MaKH='" + xxx + "'"
            End If
            frmMain.Rpt.ReportFileName = ""
            If loaict = 1 And Len(Dir(pCurDir + "REPORTS\PHIEUN.RPT")) > 0 Then frmMain.Rpt.ReportFileName = "PHIEUN.RPT"
            If loaict = 2 And Len(Dir(pCurDir + "REPORTS\PHIEUX.RPT")) > 0 Then frmMain.Rpt.ReportFileName = "PHIEUX.RPT"
            If loaict = 2 And CoPSTK("62") And Len(Dir(pCurDir + "REPORTS\PHIEUX62.RPT")) > 0 Then frmMain.Rpt.ReportFileName = "PHIEUX62.RPT"
            If loaict = 2 And CoPSTK("64") And Len(Dir(pCurDir + "REPORTS\PHIEUX64.RPT")) > 0 Then frmMain.Rpt.ReportFileName = "PHIEUX64.RPT"
            If Len(frmMain.Rpt.ReportFileName) = 0 Then frmMain.Rpt.ReportFileName = "PHIEUNX.RPT"
            frmMain.Rpt.Formulas(4) = "Kho='" + CboNguon(1).Text + "'"
            frmMain.Rpt.Formulas(6) = "TKno='" + tkno + "'"
            frmMain.Rpt.Formulas(7) = "TKco='" + TkCo + "'"
            frmMain.Rpt.Formulas(9) = "Sotien='" + ToVNText(ttien + v + v338) + " ®ång'"
            frmMain.Rpt.Formulas(10) = "DiaChi='" + DiaChiNX + "'"
            frmMain.Rpt.Formulas(8) = "TenNN='" + TenNX + "'"
            frmMain.Rpt.Formulas(11) = "LoaiCT=" + CStr(loaict)
            If v <> 0 Then frmMain.Rpt.Formulas(19) = "VAT=" + DoiDau(v)
            If v338 <> 0 Then frmMain.Rpt.Formulas(17) = "P=" + DoiDau(v338)
            frmMain.Rpt.PrinterCopies = 2
        Case 0, 7, 8:


            ''                       In_hoa_don1 sotien, i, k, xxx, sodu, v, lp, ms, mv, tien, ttien, luong, tkno, TkCo, TK, tiennt, ts, HTTT, tl, thue, v338, v521, X, shtk, vt, dn, DC, dnt, CK, somh, tp, lanin, stt, loaitien
            ''                     '   If CBINBANGKE.Value = True Then
            ''                        In_hoa_don2 sotien, i, k, xxx, sodu, v, lp, ms, mv, tien, ttien, luong, tkno, TkCo, TK, tiennt, ts, HTTT, tl, thue, v338, v521, X, shtk, vt, dn, DC, dnt, CK, somh, tp, lanin, stt, loaitien
            ''                     '   End If
            If in_hd = 0 Then
                In_hoa_don1 sotien, i, k, xxx, sodu, v, lp, ms, mv, tien, ttien, luong, tkno, TkCo, TK, tiennt, ts, HTTT, tl, thue, v338, v521, X, shtk, vt, dn, DC, dnt, CK, somh, tp, lanin, stt, loaitien
            End If
            If in_hd = 1 Then
                In_hoa_don2 sotien, i, k, xxx, sodu, v, lp, ms, mv, tien, ttien, luong, tkno, TkCo, TK, tiennt, ts, HTTT, tl, thue, v338, v521, X, shtk, vt, dn, DC, dnt, CK, somh, tp, lanin, stt, loaitien
            End If

            chophep_in = 1
        Case 9, 10:
            If P_1 = 0 Then TenNX = FrmGetStr.GetString("Tªn ng­êi giao nhËn:", "PhiÕu nhËp xuÊt", IIf(MaSoCT > 0, TenNX, ""))
            ms = LayMaKH(IIf(loaict = 9, 1, -1))
            If hdcount >= 0 Then
                If loaict = 9 Then frmMain.Rpt.Formulas(13) = "TyLe='ThuÕ GTGT (" + CStr(HD(hdcount).TyLe) + "%):'"
                If ms = 0 Then ms = HD(hdcount).MaKhachHang
            End If
            If ms > 0 Then
                frmMain.Rpt.Formulas(14) = "TenKH='" + TenKH(xxx, ms) + "'"
                frmMain.Rpt.Formulas(15) = "MaKH='" + xxx + "'"
            End If
            frmMain.Rpt.ReportFileName = "PHIEUTS.RPT"
            frmMain.Rpt.Formulas(4) = "Kho='" + CboNguon(0).Text + "'"
            frmMain.Rpt.Formulas(6) = "TKno='" + tkno + "'"
            frmMain.Rpt.Formulas(7) = "TKco='" + TkCo + "'"
            frmMain.Rpt.Formulas(9) = "Sotien='" + ToVNText(ttien + v) + " ®ång'"
            frmMain.Rpt.Formulas(8) = "TenNN='" + TenNX + "'"
            frmMain.Rpt.Formulas(11) = "LoaiCT=" + CStr(loaict)
            frmMain.Rpt.Formulas(18) = "VAT=" + DoiDau(v)
        End Select
        frmMain.Rpt.Formulas(5) = "Ngay='" + Format(ngay(0), Mask_DR) + "'"
    Case 2:
        ttien = 0
        k = 0
        With GrdChungtu
            For i = 0 To .Rows - 1
                .Row = i
                .col = 1
                If Len(.Text) = 0 Then Exit For
                If .Text = CmdPhieu(2).tag Then
                    TK.InitTaikhoanSohieu .Text
                    .col = 3
                    If Len(.Text) > 0 Then
                        mv = .Text
                        .col = 4
                        ttien = Cdbl5(.Text)
                        k = 1
                    Else
                        .col = 7
                        ttien = ttien + Cdbl5(.Text)
                    End If
                End If
            Next
        End With

        If k = 0 Then
            sotien = ToVNText(ttien)
            xxx = "SELECT 1 AS 1,#" + Format(ngay(0), Mask_DB) + "# AS Ngay,'" + unc1 + "' AS TenNV,'" + txt(1).Text + "' AS LyDo,'" + sotien + " ®ång ch½n' AS SoTien," + DoiDau(Abs(ttien)) + " AS XTien FROM License"
            For i = 2 To Ppunc
                xxx = xxx + " UNION SELECT " + CStr(i) + " AS 1,#" + Format(ngay(0), Mask_DB) + "# AS Ngay,'" + unc1 + "' AS TenNV,'" + txt(1).Text + "' AS LyDo,'" + sotien + " ®ång ch½n' AS SoTien," + DoiDau(Abs(ttien)) + " AS XTien FROM License"
            Next
            frmMain.Rpt.ReportFileName = "UNC.RPT"
        Else
            sotien = ToVNText(Fix(ttien)) + " " + mv
            If ttien - Fix(ttien) > 0 Then sotien = sotien + " vµ " + ToVNText(Fix(0.5 + 100 * (ttien - Fix(ttien)))) + " cents"
            xxx = "SELECT 1 AS 1,#" + Format(ngay(0), Mask_DB) + "# AS Ngay,'" + unc1 + "' AS TenNV,'" + txt(1).Text + "' AS LyDo,'" + sotien + "' AS SoTien," + DoiDau(Abs(ttien)) + " AS XTien FROM License"
            For i = 2 To Ppunc
                xxx = xxx + " UNION SELECT " + CStr(i) + " AS 1,#" + Format(ngay(0), Mask_DB) + "# AS Ngay,'" + unc1 + "' AS TenNV,'" + txt(1).Text + "' AS LyDo,'" + sotien + "' AS SoTien," + DoiDau(Abs(ttien)) + " AS XTien FROM License"
            Next
            frmMain.Rpt.ReportFileName = "UNC2.RPT"
            frmMain.Rpt.Formulas(10) = "LoaiT='" + mv + "'"
        End If
        SetSQL "QNhatKy", xxx

        frmMain.Rpt.Formulas(3) = "SoPhieu='" + LaySH(txt(0).Text, 1) + "'"
        frmMain.Rpt.Formulas(4) = "SoTK='" + unc2 + "'"
        xxx = LaySH(unc3, 1, "-")
        If Len(xxx) = 0 Then xxx = unc3
        frmMain.Rpt.Formulas(5) = "NH='" + xxx + "'"
        frmMain.Rpt.Formulas(6) = "SoTK2='" + TK.GhiChu + "'"
        frmMain.Rpt.Formulas(7) = "NH2='" + TK.Ten + "'"
        xxx = LaySH(unc3, 2, "-")
        If Len(xxx) > 0 Then frmMain.Rpt.Formulas(11) = "TP='" + xxx + "'"
    End Select
    If chophep_in <> 1 Then
        frmMain.Rpt.WindowTitle = frmMain.Rpt.ReportFileName
        InBaoCaoRPT pNN
    End If
    If (Index = 0 Or Index = 3) And lanin = 0 Then
        lanin = lanin + 1
        lp = IIf(lp < 0, 1, -1)
        If CoPSTK(shtk, lp, X) Then
            SetRptInfo
            tiennt = 0
            sodu = 0
            GoTo C1
        End If
    End If
KT:
    Set TK = Nothing
    Set vt = Nothing
    Set tp = Nothing
End Sub



Private Sub cmdkh_Click(Index As Integer)
    Me.MousePointer = 11
    txtshkh(Index).Text = FrmKhachHang.ChonKhachHang(txtshkh(Index).Text)
    Me.MousePointer = 0
    RFocus txtshkh(Index)
End Sub

Public Sub CmdPhieu_Click(Index As Integer)
    ngay(0) = CVDate(MedNgay(0).Text)
    ngay(1) = CVDate(MedNgay(1).Text)
    '    If Checkinbangkevahoadon.Value = 1 Then
    '        in_hoa_don_tong_hop Index, 0
    '        in_hoa_don_tong_hop Index, 1
    '    ElseIf checkinbangke = 1 Then
    '        in_hoa_don_tong_hop Index, 1
    '    Else
    in_hoa_don_tong_hop Index, 0
    '   End If
End Sub
Sub In_hoa_don2(sotien As String, i As Integer, k As Integer, xxx As String, sodu As Integer, v As Double, lp As Integer, ms As Long, mv As String, tien As Double, ttien As Double, luong As Double, tkno As String, TkCo As String, TK As ClsTaikhoan, tiennt As Double, ts As clsTaiSan, HTTT As String, tl As Integer, thue As Double, v338 As Double, v521 As Double, X As Double, shtk As String, vt As ClsVattu, dn As Double, DC As Double, dnt As Double, CK As Double, somh As Integer, tp As Cls154, lanin As Integer, stt As Integer, loaitien As String)
    tiennt = 0
    v = 0
    sodu = 0
    tl = -1
    thue = 0
    ExecuteSQL5 "DELETE * FROM PhieuNX"
    With GrdChungtu
        For i = 0 To .Rows - 1
            .Row = i
            .col = 1
            If Len(.Text) = 0 Then Exit For
            TK.InitTaikhoanSohieu .Text

            If (loaict = 9 Or loaict = 10) And TK.tk_id = TSCD_ID Then
                .col = IIf(loaict = 9, 6, 7)
                tien = Cdbl5(.Text)
                ttien = ttien + tien
                Set ts = New clsTaiSan
                If pMaTaiSan > 0 Then
                    ts.ChiDinh pMaTaiSan, 0
                Else
                    .col = 3
                    ts.ChiDinh SoHieu2MaSo(.Text, "TaiSan"), 0
                End If
                ExecuteSQL5 "INSERT INTO PhieuNX (MaSo, SoCT,DienGiaiCT,SoHieu,DienGiai,SoLuong,ThanhTien) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhieuNX") + 1) + ",'" + LaySH(txt(0).Text, 2) _
                          + "','" + txt(1).Text + "','" + ts.sohieu + "','" + ts.Ten + "',1," + DoiDau(tien) + ")"
                Set ts = Nothing
            End If

            If TK.tk_id = TKDT_ID Or TK.tk_id = TKVT_ID Then
                If loaict = 1 Or TK.tk_id = TKGT_ID Then
                    If InStr(tkno, TK.sohieu) = 0 Then tkno = tkno + IIf(Len(tkno) > 0, ", ", "") + TK.sohieu
                Else
                    If InStr(TkCo, TK.sohieu) = 0 Then TkCo = TkCo + IIf(Len(TkCo) > 0, ", ", "") + TK.sohieu
                End If

                .col = 3
                If Len(.Text) > 0 Then
                    sodu = 1
                    If TK.tk_id2 = TKDT_ID Then
                        tp.InitTPSohieu .Text
                        mv = tp.DonVi
                    Else
                        vt.InitVattuSohieu .Text
                        .col = 23
                        If Not KtraDVT(vt.MaSo, CLng5(.Text), mv) And (loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8) Then
                            mv = vt.DonVi
                        End If
                        If vt.MaSo > 0 Then XDTyLeQD vt.MaSo
                    End If
                    .col = 4
                    luong = Cdbl5(.Text)
                    .col = 24
                    tiennt = Cdbl5(.Text)
                    .col = 26
                    X = Cdbl5(.Text)
                    CK = CK + X
                    .col = IIf(loaict = 1 Or TK.tk_id = TKGT_ID, 6, 7)
                    tien = Cdbl5(.Text)
                    If loaict = 8 Then tien = Abs(tien)
                    ' ttien = ttien + tien ' bo di vi ly do in ra bang ke tong tien tang gap doi
                    .col = 5
                    tiennt = Cdbl5(.Text)
                    ExecuteSQL5 "INSERT INTO PhieuNX (MaSo, SoCT,DienGiaiCT,SoHieu,DienGiai,SoLuong,ThanhTien,DVT,DonGia,TyLe,Thue,ThanhTien2,CK) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhieuNX") + 1) + ",'" + LaySH(txt(0).Text, 2) + "','" + txt(1).Text _
                              + "','" + IIf(vt.MaSo > 0, vt.sohieu, tp.sohieu) + "','" + IIf(vt.MaSo > 0, vt.TenVattu, tp.TenVattu) + "'," + DoiDau(luong) + "," + DoiDau(tien) + ",'" + mv + "'," + DoiDau(tiennt) + "," + CStr(tl) + "," + DoiDau(thue) + "," + DoiDau(tiennt) + "," + DoiDau(X) + ")"
                    somh = somh + 1
                Else
                    If TK.tk_id = TKDT_ID Then
                        .col = 7
                        tiennt = tiennt + Cdbl5(.Text)
                    End If
                End If
            Else
                If loaict <> 8 Then
                    .col = 6
                    luong = Cdbl5(.Text)
                    If luong <> 0 And TK.tk_id = GTGTKT_ID Then v = v + luong
                    If luong <> 0 And InStr(tkno, TK.sohieu) = 0 Then tkno = tkno + IIf(Len(tkno) > 0, ", ", "") + TK.sohieu
                    .col = 7
                    luong = Cdbl5(.Text)
                    If luong <> 0 And TK.tk_id = GTGTPN_ID Then v = v + luong
                    If luong <> 0 And InStr(TkCo, TK.sohieu) = 0 Then TkCo = TkCo + IIf(Len(TkCo) > 0, ", ", "") + TK.sohieu
                    If TK.tk_id <> TKVT_ID And (Left(TK.sohieu, 3) = "138" Or Left(TK.sohieu, 3) = "338") Then
                        .col = 6
                        v338 = v338 + Cdbl5(.Text)
                    End If
                Else
                    If TK.tk_id = GTGTPN_ID Then
                        .col = 19
                        TkCo = .Text
                        .col = 7
                        thue = Abs(Cdbl5(.Text))
                        .col = 6
                        thue = thue + Cdbl5(.Text)
                        v = v + thue
                        .col = 3
                        tl = CInt5(.Text)
                        .col = 18
                        k = CInt5(.Text)
                        If k >= 0 And k <= hdcount Then
                            If HD(CInt5(.Text)).KCT = 1 Then tl = -1
                        End If
                    End If
                    If TK.tk_id <> TKDT_ID And (Left(TK.sohieu, 3) = "138" Or Left(TK.sohieu, 3) = "338") Then
                        .col = 7
                        v338 = v338 + Cdbl5(.Text)
                    End If
                End If
            End If
        Next
    End With
    Select Case loaict

    Case 0, 7, 8:
        If hdcount >= 0 Then
            frmMain.Rpt.Formulas(11) = "TyLe=" + CStr(HD(hdcount).TyLe)
            frmMain.Rpt.Formulas(16) = "SoHD='" + HD(hdcount).sohd + "'"
            ms = HD(hdcount).MaKhachHang
            HTTT = HD(hdcount).HTTT
            frmMain.Rpt.Formulas(17) = "TyGia=" + DoiDau(HD(hdcount).tygia)

            CoPSTK "", 1, X
            CoPSTK "11", -1, luong

            SoDuKHNgay ms, ngay(0), dn, DC, dnt

            frmMain.Rpt.Formulas(20) = "DuDK=" + DoiDau(IIf(MaSoCT > 0, dn - DC - X + luong, dn - DC))
            frmMain.Rpt.Formulas(21) = "PS=" + DoiDau(X)
            frmMain.Rpt.Formulas(22) = "11=" + DoiDau(luong)
        Else
            ms = LayMaKH(IIf(loaict = 1, 1, -1))
        End If
        If P_1 = 0 Then
            FThuChi.tag = 3
            xxx = ""
            '  FThuChi.GetPhieu TenBH, DiaChiBH, "...", 0, HanTT, xxx
            If Len(xxx) > 0 And xxx <> "..." Then frmMain.Rpt.Formulas(19) = "SoDH='" + xxx + "'"
        End If
        If ms > 0 Then
            frmMain.Rpt.Formulas(14) = "TenKH='" + TenKH(xxx, ms, mv) + "'"
            frmMain.Rpt.Formulas(15) = "MaKH='" + xxx + "'"
        End If
        '     frmMain.Rpt.ReportFileName = IIf(Chk.Value = 0, "HOADON" + IIf(pGiaUSD > 0, "X", "") + IIf(somh > 10, "2", "") + ".RPT", "BAOGIA" + IIf(pGiaUSD > 0, "X", "") + ".RPT")
        frmMain.Rpt.ReportFileName = IIf(Chk.Value = 0, "BANGKE" + IIf(pGiaUSD > 0, "X", "") + IIf(somh > 10, "", "") + ".RPT", "BAOGIA" + IIf(pGiaUSD > 0, "X", "") + ".RPT")
        frmMain.Rpt.Formulas(3) = "DC1='" + frmMain.LbCty(2).Caption + "'"
        frmMain.Rpt.Formulas(4) = "DiaChi='" + DiaChiBH + "'"
        frmMain.Rpt.Formulas(6) = "MS1='" + frmMain.LbCty(8).Caption + "'"
        frmMain.Rpt.Formulas(7) = "MS2='" + MSTBH + "'"
        frmMain.Rpt.Formulas(8) = "TenNN='" + txtVT(1).Text + "'"
        frmMain.Rpt.Formulas(10) = "HTTT='" + HTTT + "'"
        frmMain.Rpt.Formulas(12) = "Kho='" + CboNguon(1).Text + "'"
        If Year(HanTT) > 1900 Then
            frmMain.Rpt.Formulas(13) = "HanTT='" + Format(HanTT, Mask_DR) + "'"
        End If
        CoPSTK "521", -1, v521
        If sodu = 0 Then
            ExecuteSQL5 "INSERT INTO PhieuNX (MaSo,SoCT,DienGiaiCT,SoHieu,DienGiai,SoLuong,ThanhTien) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhieuNX") + 1) + ",'" + LaySH(txt(0).Text, 2) + "','" + txt(1).Text _
                      + "','...','" + txt(1).Text + "',0," + DoiDau(tiennt) + ")"
            frmMain.Rpt.Formulas(9) = "Sotien='" + ToVNText(tiennt + v + v338 - v521) + " ®ång.'"
        Else
            frmMain.Rpt.Formulas(9) = "Sotien='" + ToVNText(ttien + v + v338 - v521) + " ®ång.'"
        End If
        If v <> 0 Then frmMain.Rpt.Formulas(50) = "VAT=" + DoiDau(v)
        If v338 <> 0 Then frmMain.Rpt.Formulas(19) = "P=" + DoiDau(v338)
        frmMain.Rpt.Formulas(18) = "CK=" + IIf(v521 <> 0, DoiDau(v521), "SUM({PhieuNX.CK})")
        If pGiaUSD > 0 Then frmMain.Rpt.Formulas(17) = "TyGia=" + DoiDau(pRate)
        CoPSTK "51", 1, ttien
        If ttien < 0 Then frmMain.Rpt.Formulas(20) = "GhiChu='" + txt(1).Text + "'"
    End Select
    frmMain.Rpt.Formulas(5) = "Ngay='" + Format(ngay(0), Mask_DR) + "'"
    frmMain.Rpt.WindowTitle = frmMain.Rpt.ReportFileName
    SetSQL "bangkebanhang", "SELECT hethongtk.sohieu, IIf(isnull(vattu.tenvattu),hethongtk.ten,vattu.tenvattu) AS ten, vattu.DonVi, chungtu.SoPS2Co, IIf(chungtu.SoPS2Co>0,chungtu.sops/chungtu.SoPS2Co,0) AS dongia, chungtu.sops AS tongtien, 0 AS so1, 0 AS so2 FROM (chungtu LEFT JOIN hethongtk ON chungtu.MaTKCo=hethongtk.maso) LEFT JOIN vattu ON vattu.maso=chungtu.MaVattu WHERE left(hethongtk.sohieu,3) ='511' and chungtu.sohieu='" + txt(0).Text + "'"

    InBaoCaoRPT pNN

End Sub

Sub In_hoa_don1(sotien As String, i As Integer, k As Integer, xxx As String, sodu As Integer, v As Double, lp As Integer, ms As Long, mv As String, tien As Double, ttien As Double, luong As Double, tkno As String, TkCo As String, TK As ClsTaikhoan, tiennt As Double, ts As clsTaiSan, HTTT As String, tl As Integer, thue As Double, v338 As Double, v521 As Double, X As Double, shtk As String, vt As ClsVattu, dn As Double, DC As Double, dnt As Double, CK As Double, somh As Integer, tp As Cls154, lanin As Integer, stt As Integer, loaitien As String)
    Dim sql
    sql = "select * from license"
    Dim lisen
    Set lisen = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If lisen.recordCount > 0 Then
        frmMain.Rpt.Formulas(1) = "TenCty='" + lisen!TenCty + "'"    ' 'lisen!Tenhoadon
        frmMain.Rpt.Formulas(2) = "TenCn='" + lisen!tencn + "'"
        frmMain.Rpt.Formulas(3) = "DC1='" + lisen!DiaChi + "'"
        frmMain.Rpt.Formulas(4) = "Fax='" + lisen!Fax + "'"
        frmMain.Rpt.Formulas(999) = "dienthoai='" + lisen!Tel + "'"
        frmMain.Rpt.Formulas(6) = "taikhoantienviet='" + lisen!TaiKhoanVN + "'"
        frmMain.Rpt.Formulas(7) = "mstcty='" + lisen!masothue + "'"
        frmMain.Rpt.Formulas(150) = "email='" + lisen!email + "'"
    End If
    frmMain.Rpt.Formulas(8) = "Kyhieu='" + txtVT(1).Text + "'"
    frmMain.Rpt.Formulas(888) = "SOHD='" + txt(0).Text + "'"

    '                    frmMain.Rpt.Formulas(10) = "copy='" + "1" + "'"

    Dim ngay_hd As Date
    ngay_hd = CDate(MedNgay(0).Text)

    frmMain.Rpt.Formulas(11) = "NAMHD='" + str(Year(ngay_hd)) + "'"
    frmMain.Rpt.Formulas(12) = "NGAYHD= '" + str(Day(ngay_hd)) + "'"
    frmMain.Rpt.Formulas(13) = "THANGHD='" + str(Month(ngay_hd)) + "'"


    frmMain.Rpt.Formulas(14) = "mavach='" + "" + "'"
    frmMain.Rpt.Formulas(15) = "mv1='" + "" + "'"
    frmMain.Rpt.Formulas(16) = "mv2='" + "" + "'"
    frmMain.Rpt.Formulas(17) = "mv3='" + "" + "'"


    '   frmMain.Rpt.Formulas(18) = "TenNN='" + txt(3).Text + "'"
    '   frmMain.Rpt.Formulas(19) = "TenKH='" + txtVT(7).Text + "'"
    frmMain.Rpt.Formulas(20) = "Ms2='" + txtVT(9).Text + "'"
    frmMain.Rpt.Formulas(21) = "Diachi='" + txtVT(8).Text + "'"

    frmMain.Rpt.Formulas(23) = "taikhoan='" + "" + "'"




    If hdcount >= 0 Then
        '                            frmMain.Rpt.Formulas(11) = "TyLe=" + CStr(HD(hdcount).TyLe)
        '                            frmMain.Rpt.Formulas(16) = "SoHD='" + HD(hdcount).sohd + "'"
        ms = HD(hdcount).MaKhachHang
        HTTT = HD(hdcount).HTTT
        '                            frmMain.Rpt.Formulas(17) = "TyGia=" + DoiDau(HD(hdcount).tygia)

        CoPSTK "", 1, X
        CoPSTK "11", -1, luong

        SoDuKHNgay ms, ngay(0), dn, DC, dnt
        '
        '                            frmMain.Rpt.Formulas(20) = "DuDK=" + DoiDau(IIf(MaSoCT > 0, dn - DC - X + luong, dn - DC))
        '                            frmMain.Rpt.Formulas(21) = "PS=" + DoiDau(X)
        '                            frmMain.Rpt.Formulas(22) = "11=" + DoiDau(luong)
    Else
        ms = LayMaKH(IIf(loaict = 1, 1, -1))
    End If
    If P_1 = 0 Then
        FThuChi.tag = 3
        xxx = ""
        FThuChi.GetPhieu TenBH, DiaChiBH, "...", 0, HanTT, xxx
        ' If Len(xxx) > 0 And xxx <> "..." Then frmMain.Rpt.Formulas(19) = "SoDH='" + xxx + "'"
    End If


    If Check2.Value = 1 Then
        frmMain.Rpt.Formulas(400) = "copy = '" + "copy" + "'"
        frmMain.Rpt.Formulas(401) = "banmau = '" + "B¶n mÉu" + "'"
    Else
        frmMain.Rpt.Formulas(400) = "copy = '" + "" + "'"
        frmMain.Rpt.Formulas(401) = "banmau = '" + "" + "'"
    End If
    frmMain.Rpt.Formulas(405) = "hantt= '" + thoihanthanhtoan.Text + " " + "'"
    frmMain.Rpt.Formulas(406) = "sophieu= '" + sochungtu.Text + " " + "'"
    frmMain.Rpt.Formulas(550) = "thanhtoan = '" + hinhthucthanhtoan.Text + "'"
    frmMain.Rpt.ReportFileName = IIf(Chk.Value = 0, "HOADON" + IIf(pGiaUSD > 0, "X", "") + IIf(somh > 10, "2", "") + ".RPT", "BAOGIA" + IIf(pGiaUSD > 0, "X", "") + ".RPT")

    CoPSTK "521", -1, v521
    If sodu = 0 Then
        ExecuteSQL5 "INSERT INTO PhieuNX (MaSo,SoCT,DienGiaiCT,SoHieu,DienGiai,SoLuong,ThanhTien) VALUES (" + CStr(Lng_MaxValue("MaSo", "PhieuNX") + 1) + ",'" + LaySH(txt(0).Text, 2) + "','" + txt(1).Text _
                  + "','...','" + txt(1).Text + "',0," + DoiDau(tiennt) + ")"
        frmMain.Rpt.Formulas(9) = "Sotien='" + ToVNText(tiennt + v + v338 - v521) + " ®ång'"
    Else
        frmMain.Rpt.Formulas(9) = "Sotien='" + ToVNText(ttien + v + v338 - v521) + " ®ång'"
    End If

    frmMain.Rpt.Formulas(600) = "dong1= '" + " " + "'"
    frmMain.Rpt.Formulas(601) = "dong2= '" + " " + "'"
    frmMain.Rpt.Formulas(602) = "dong3= '" + " " + "'"
    frmMain.Rpt.Formulas(603) = "dong4= '" + " " + "'"
    frmMain.Rpt.Formulas(604) = "dong5= '" + " " + "'"
    frmMain.Rpt.Formulas(605) = "dong6= '" + " " + "'"
    frmMain.Rpt.Formulas(606) = "dong7= '" + " " + "'"
    frmMain.Rpt.Formulas(607) = "dong8= '" + " " + "'"
    frmMain.Rpt.Formulas(608) = "dong9= '" + " " + "'"
    frmMain.Rpt.Formulas(609) = "dong10= '" + " " + "'"
    frmMain.Rpt.Formulas(610) = "dong11= '" + " " + "'"
    frmMain.Rpt.Formulas(611) = "dong12= '" + " " + "'"

    sql = "select * from PhieuNX order by ThanhTien asc "
    Dim chitiet_sp
    Set chitiet_sp = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If chitiet_sp.recordCount > 0 Then
        Dim dem, dem1 As Integer
        Dim stt_dem As String
        Dim gach
        dem1 = 100
        dem = 90
        Dim g, kk As Integer
        Dim TongTien As Double
        Dim thoat As Boolean
        thoat = True
        kk = 1
        TongTien = 0
        Dim mang_stt(100) As String
        Dim mang_diengiai(100) As String
        Dim mang_dvt(100) As String
        Dim mang_soluong(100) As String
        Dim mang_dongia(100) As String
        Dim mang_thanhtien(100) As String
        Dim mang_trunggian() As String
        Dim hhh As Integer
        hhh = 1
        If (Checkinbangkevahoadon.Value = 1) Then
            Check1.Value = 1
        End If
        If (Check1.Value = 1) Then
            If (Checkinbangkevahoadon.Value = 1) Then
                mang_trunggian = Form2.mang_chuoi("Baùn haøng theo baûng keâ.")
            Else
                mang_trunggian = Form2.mang_chuoi(txt(1).Text)
            End If
            For i = 1 To UBound(mang_trunggian)
                If i = 1 Then
                    mang_stt(i) = "1"
                Else
                    mang_stt(i) = ""
                End If
                mang_diengiai(i) = mang_trunggian(i)
                If i = 1 Then
                    mang_dvt(i) = "..."
                    mang_soluong(i) = "..."
                    mang_dongia(i) = "..."
                Else
                    mang_dvt(i) = ""
                    mang_soluong(i) = ""
                    mang_dongia(i) = ""
                End If
                If (i = 1) Then
                    mang_thanhtien(i) = chuyenso(str(ttien + v + v338 - v521 - Int(DoiDau(v))))

                Else
                    mang_thanhtien(i) = chuyenso("")
                End If
                hhh = hhh + 1
                If Len(mang_trunggian(i)) <= 0 Then
                    Exit For
                End If
            Next
            hhh = hhh - 1
        Else
            Do While Not chitiet_sp.EOF
                mang_stt(hhh) = CStr(hhh)
                mang_diengiai(hhh) = CStr(chitiet_sp!diengiai)

                mang_dvt(hhh) = CStr(chitiet_sp!dvt)
                mang_soluong(hhh) = chuyenso(chitiet_sp!SoLuong)
                mang_dongia(hhh) = chuyenso(chitiet_sp!dongia)
                mang_thanhtien(hhh) = chuyenso(chitiet_sp!ThanhTien)
                hhh = hhh + 1
                chitiet_sp.MoveNext
            Loop

        End If
        'chitiet_sp.MoveFirst

        ' If (Check1.Value = 1) Then
        'Do While Not chitiet_sp.EOF And thoat
        For kk = 1 To hhh - 1
            If kk = 1 Then stt_dem = "1"
            If kk = 2 Then stt_dem = "2"
            If kk = 3 Then stt_dem = "3"
            If kk = 4 Then stt_dem = "4"
            If kk = 5 Then stt_dem = "5"
            If kk = 6 Then stt_dem = "6"
            If kk = 7 Then stt_dem = "7"
            If kk = 8 Then stt_dem = "8"
            If kk = 9 Then stt_dem = "9"
            If kk = 10 Then stt_dem = "10"
            If kk = 11 Then stt_dem = "11"
            If kk = 12 Then stt_dem = "12"
            frmMain.Rpt.Formulas(dem) = "stt" + stt_dem + "='" + mang_stt(kk) + "'"
            dem = dem + 1
            gach = mang_diengiai(kk)
            frmMain.Rpt.Formulas(dem) = "diengiai" + stt_dem + "='" + mang_diengiai(kk) + "'"


            dem = dem + 1
            frmMain.Rpt.Formulas(dem) = "dvt" + stt_dem + "='" + mang_dvt(kk) + "'"
            dem = dem + 1
            Dim siiii
            If (Len(mang_soluong(kk)) > 0) Then
                frmMain.Rpt.Formulas(dem) = "soluong" + stt_dem + "='" + mang_soluong(kk) + "'"
            Else
                frmMain.Rpt.Formulas(dem) = "soluong" + stt_dem + "='" + "" + "'"
            End If
            dem = dem + 1
            If (Len(mang_dongia(kk)) > 0) Then
                frmMain.Rpt.Formulas(dem) = "dongia" + stt_dem + "='" + mang_dongia(kk) + "'"
            Else

                frmMain.Rpt.Formulas(dem) = "dongia" + stt_dem + "='" + "" + "'"
            End If
            dem = dem + 1
            If (Len(mang_thanhtien(kk)) > 0) Then
                'TongTien = TongTien + CDbl(mang_thanhtien(kk))
                frmMain.Rpt.Formulas(dem) = "thanhtien" + stt_dem + "='" + mang_thanhtien(kk) + "'"    '
            Else
                frmMain.Rpt.Formulas(dem) = "thanhtien" + stt_dem + "='" + "" + "'"
            End If
            dem = dem + 1
            '  kk = kk + 1
            ' chitiet_sp.MoveNext
            If kk > 13 Then
                'thoat = False
                MsgBox "B¹n ®· v­ît qu¸ sè dßng quy ®Þnh cña hãa ®¬n."
                Exit For
            End If
        Next
        If kk = 1 Then
            stt_dem = "1"
            '     frmMain.Rpt.Formulas(600) = "dong1= '" + " " + "'"
        End If
        If kk = 2 Then
            stt_dem = "1"
            frmMain.Rpt.Formulas(601) = "dong1= '" + "_________________________________________________________________________________________         " + "'"
        End If
        If kk = 3 Then
            stt_dem = "2"
            frmMain.Rpt.Formulas(602) = "dong2= '" + "_________________________________________________________________________________________" + "'"
        End If
        If kk = 4 Then
            stt_dem = "3"
            frmMain.Rpt.Formulas(603) = "dong3= '" + "_________________________________________________________________________________________" + "'"
        End If
        If kk = 5 Then
            stt_dem = "4"
            frmMain.Rpt.Formulas(604) = "dong4= '" + "____________________________________________________________________________________________" + "'"
        End If
        If kk = 6 Then
            frmMain.Rpt.Formulas(605) = "dong5= '" + "____________________________________________________________________________________________" + " '"
            stt_dem = "5"
        End If
        If kk = 7 Then
            frmMain.Rpt.Formulas(606) = "dong6= '" + "____________________________________________________________________________________________" + "'"
            stt_dem = "6"
        End If
        If kk = 8 Then
            frmMain.Rpt.Formulas(607) = "dong7= '" + "____________________________________________________________________________________________" + "'"
            stt_dem = "7"
        End If
        If kk = 9 Then
            frmMain.Rpt.Formulas(608) = "dong8= '" + "____________________________________________________________________________________________" + "'"
            stt_dem = "8"
        End If
        If kk = 10 Then
            stt_dem = "9"
            frmMain.Rpt.Formulas(609) = "dong9= '" + "____________________________________________________________________________________________" + "'"
        End If
        If kk = 11 Then
            frmMain.Rpt.Formulas(610) = "dong10= '" + "___________________________________________________________________________________________" + "'"
            stt_dem = "10"
        End If
        If kk = 12 Then
            frmMain.Rpt.Formulas(610) = "dong11= '" + "___________________________________________________________________________________________" + "'"
            stt_dem = "11"
        End If

    End If
    chitiet_sp.Close
    Set chitiet_sp = Nothing
    '  frmMain.Rpt.Formulas(9) = "Sotien='" + ToVNText(tiennt + v + v338 - v521) + " ®ång'"
    ' Else
    '   frmMain.Rpt.Formulas(28) = "Sotien='" + ToVNText(ttien + v + v338 - v521) + " ®ång'"
    ' End If
    '                        If v <> 0 Then frmMain.Rpt.Formulas(50) = "VAT=" + DoiDau(v)
    frmMain.Rpt.Formulas(19) = "P=" + "0"
    If v338 <> 0 Then frmMain.Rpt.Formulas(19) = "P=" + DoiDau(v338)

    ' frmMain.Rpt.Formulas(18) = "CK=" + IIf(v521 <> 0, DoiDau(v521), "SUM({PhieuNX.CK})")
    '                        If pGiaUSD > 0 Then frmMain.Rpt.Formulas(17) = "TyGia=" + DoiDau(pRate)
    CoPSTK "51", 1, ttien
    '                        If ttien < 0 Then frmMain.Rpt.Formulas(20) = "GhiChu='" + txt(1).Text + "'"
    '
    ' frmMain.Rpt.Formulas(24) = "tyle='" + CStr(HD(hdcount).TyLe) + "'"
    ''  If (hdcount > 0) Then
    ''    frmMain.Rpt.Formulas(24) = "tyle='" + CStr(HD(hdcount).TyLe) + "'"
    '' Else
    ''   frmMain.Rpt.Formulas(24) = "tyle='" + "0" + "'"
    '' End If
    frmMain.Rpt.Formulas(24) = "tyle='" + CStr(SelectSQL("SELECT tyle AS F1 FROM hoadon where sohd = '" + Trim(txt(0).Text) + "'")) + "'"

    frmMain.Rpt.Formulas(27) = "vat='" + chuyenso(DoiDau(v)) + "'"
    ' frmMain.Rpt.Formulas(27) = "vat='" + DoiDau(v) + "'"

    If (Check1.Value = 1) Then
        frmMain.Rpt.Formulas(500) = "thanhtien1 ='" + chuyenso(str(ttien + v + v338 - v521 - Int(DoiDau(v)))) + "'"  '
    End If
    frmMain.Rpt.Formulas(28) = "sotien='" + ToVNText(ttien + v + v338 - v521) + " ®ång." + "'"
    frmMain.Rpt.Formulas(26) = "t1='" + chuyenso(str(ttien + v + v338 - v521 - Int(DoiDau(v)))) + "'"
    frmMain.Rpt.Formulas(25) = "tt='" + chuyenso(str(ttien + v + v338 - v521)) + "'"
    ' If (Check1.Value = 1) Then
    '      frmMain.Rpt.Formulas(dem) = "thanhtien1='" + Format(Str(ttien + v + v338 - v521 - Int(DoiDau(v))), Mask_0) + "'"
    'End If
    Dim sql22 As String
    Dim kaka As Recordset
    sql22 = "SELECT iif(Nguoimuahang is null ,'...',Nguoimuahang) as aa1 from chungtu where sohieu = '" + FrmChungtu.txt(0).Text + "'"
    Set kaka = DBKetoan.OpenRecordset(sql22, dbOpenSnapshot)
    If kaka.recordCount > 0 Then
        frmMain.Rpt.Formulas(200) = "TenNN='" + kaka!AA1 + "'"
    Else
        frmMain.Rpt.Formulas(200) = "TenNN='" + "..." + "'"
    End If

    ' frmMain.Rpt.Formulas(200) = "TenNN='" + FThuChi.T(3).Text + "'"
    frmMain.Rpt.Formulas(300) = "TenKH='" + txtVT(7).Text + "'"

    frmMain.Rpt.Formulas(5) = "Ngay='" + Format(ngay(0), Mask_DR) + "'"

    'Set DataReport1.Sections(3).Controls("Image1").Picture = LoadPicture(App.path & ImageFile)

    frmMain.Rpt.WindowTitle = frmMain.Rpt.ReportFileName



    If Checkinbangkevahoadon.Value = 1 Then
        If CheckBox3.Value = 1 Then
            frmMain.Rpt.Formulas(505) = "Lien='" + "Lieân 3: Noäi boä" + "'"
            InBaoCaoRPT pNN
        End If
        If CheckBox2.Value = 1 Then
            frmMain.Rpt.Formulas(505) = "Lien='" + "Lieân 2: Giao cho ngöôøi mua  " + "'"
            InBaoCaoRPT pNN
        End If
        If CheckBox1.Value = 1 Then
            frmMain.Rpt.Formulas(505) = "Lien='" + "Lieân 1: Löu " + "'"
            InBaoCaoRPT pNN
        End If
        in_hoa_don_tong_hop 1, 1

    ElseIf checkinbangke = 1 Then
        in_hoa_don_tong_hop 1, 1
    Else
        If CheckBox3.Value = 1 Then
            frmMain.Rpt.Formulas(505) = "Lien='" + "Lieân 3: Noäi boä" + "'"
            InBaoCaoRPT pNN
        End If
        If CheckBox2.Value = 1 Then
            frmMain.Rpt.Formulas(505) = "Lien='" + "Lieân 2: Giao cho ngöôøi mua  " + "'"
            InBaoCaoRPT pNN
        End If
        If CheckBox1.Value = 1 Then
            frmMain.Rpt.Formulas(505) = "Lien='" + "Lieân 1: Löu " + "'"
            InBaoCaoRPT pNN
        End If

    End If


    '                  If CheckBox3.Value = 1 Then
    '                   frmMain.Rpt.Formulas(505) = "Lien='" + "Lieân 3: Noäi boä" + "'"
    '                   InBaoCaoRPT pNN
    '                 End If
    '                '  frmMain.Rpt.Formulas(10) = "copy='" + "" + "'"
    '              '    InBaoCaoRPT pNN
    '                  If CheckBox2.Value = 1 Then
    '                   frmMain.Rpt.Formulas(505) = "Lien='" + "Lieân 2: Giao cho ngöôøi mua  " + "'"
    '                   InBaoCaoRPT pNN
    '                  End If
    ''                  frmMain.Rpt.Formulas(10) = "copy='" + "" + "'"
    ''                  InBaoCaoRPT pNN
    '
    '                    If CheckBox1.Value = 1 Then
    '                      frmMain.Rpt.Formulas(505) = "Lien='" + "Lieân 1: Löu " + "'"
    '                    ' InBaoCaoRPT pNN
    '                    End If
    ''                  frmMain.Rpt.Formulas(10) = "copy='" + "" + "'"
    ''                  InBaoCaoRPT pNN
    '


End Sub
Private Function chuyenso(st As String) As String
    Dim i, j, dem, k
    Dim chuoi1 As String
    st = Trim(st)
    j = Len(st)
    chuoi1 = " "
    dem = 0

    For i = 1 To j
        If (Mid(st, i, 1) = ".") Or (Mid(st, i, 1) = ",") Then
            dem = 1
        End If
    Next
    k = 0
    If dem = 1 Then    'neu co so thap phan
        For i = 0 To Len(st)

            If (dem = 1) Then
                If (Mid(st, j, 1) = ".") Or (Mid(st, j, 1) = ",") Then
                    If j > 0 Then chuoi1 = "," + chuoi1
                    dem = 2
                Else
                    If j > 0 Then chuoi1 = Mid(st, j, 1) + chuoi1
                End If
            Else
                If k = 3 Then
                    If j > 0 Then chuoi1 = "." + chuoi1
                    k = 1
                    If j > 0 Then chuoi1 = Mid(st, j, 1) + chuoi1
                Else
                    If j > 0 Then chuoi1 = Mid(st, j, 1) + chuoi1
                    k = k + 1
                End If
            End If

            j = j - 1
        Next
    Else    ' khong co so thap phan
        For i = 0 To Len(st)
            If k = 3 Then
                If j > 0 Then chuoi1 = "." + chuoi1
                k = 1
                If j > 0 Then chuoi1 = Mid(st, j, 1) + chuoi1
            Else
                If j > 0 Then chuoi1 = Mid(st, j, 1) + chuoi1
                k = k + 1
            End If

            j = j - 1
        Next

    End If
    If (Len(chuoi1) > 0) Then
        ' MsgBox Str(Mid(chuoi1, 1, 1))
        If (Mid(chuoi1, 1, 1) = "." Or Mid(chuoi1, 1, 1) = ",") Then
            chuoi1 = Mid(2, Len(chuoi1))
        End If
    End If
    chuyenso = chuoi1
End Function
Sub tinhkyhieu()
    If OptLoai(8).Value = True Then
        Dim sql
        sql = "SELECT kyhieu as F1 from hoadon where maso in (select max(maso) from hoadon where maso in (select maso from chungtu where maloai = 8))"
        Dim rs_chungtu As Recordset
        Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        If rs_chungtu.recordCount > 0 Then txtVT(1).Text = rs_chungtu!f1
        rs_chungtu.Close

    End If
End Sub
Public Sub Test()
    Command_Click 1

End Sub
Private Sub Combo2_Change()

End Sub

Private Sub cmdtk_Click(Index As Integer)
    Me.MousePointer = 11
    ' txtShTk(Index).Text = FrmTaikhoan.ChonTk(txtShTk(Index).Text)
    FBcKt.txtShTk(Index).Text = FrmTaikhoan.ChonTk(txtShTk(Index).Text)
    txtShTk(Index).Text = FBcKt.txtShTk(Index).Text
    RFocus txtShTk(Index)
    Me.MousePointer = 0
End Sub

Private Sub cmdvt_Click(Index As Integer)
    Me.MousePointer = 11
    txtShVT(Index).Text = FrmVattu.ChonVattu(txtShVT(Index).Text)
    Me.MousePointer = 0
    RFocus txtShVT(Index)
End Sub
Function kiem_tra_333_133() As Integer
    Dim so_333
    Dim so_133
    Dim i
    so_333 = 0
    so_133 = 0
    With GrdChungtu
        For i = 0 To .Row
            .Row = i
            .col = 1
            If (Left(.Text, 3) = "133") Then so_133 = 1
            If (Left(.Text, 3) = "333") Then so_333 = 1
        Next
    End With
    kiem_tra_333_133 = so_133 + so_333
End Function
'====================================================================================================
' C¸c chøc n¨ng thªm, ghi, xãa
'====================================================================================================
Public Sub Command_Click(Index As Integer)
    Dim so_dem
    so_dem = kiem_tra_333_133
    If IsImport = False Then
        ngay(0) = CVDate(MedNgay(0).Text)
        ngay(1) = CVDate(MedNgay(1).Text)
    Else
        If IndexFirst <= fileImportList.count Then
            With fileImportList(IndexFirst)
                ngay(0) = .ngay
                ngay(1) = .ngay
            End With
        End If
    End If



    Dim chungtu As New ClsChungtu, mtk As Long, mvt As Long, mtk2 As Long, sops As Double, psnt As Double, psnt2 As Double, mtc As Long, mtc2 As Long
    Dim MaCT As Long, GhiChu As String, loai As Integer, mhdx As Integer, MaTP As Long, j As Integer, sh As String, bg As Boolean
    Dim rs_chungtu As Recordset, i As Integer, mn As Long, mk As Long, sql As String, X As New ClsTaikhoan, ctdx As Long, m As Long
    Dim so_kiem_tra
    Dim SoLoNhap As String, handung As String
    SoLoNhap = ""
    handung = ""

    so_kiem_tra = 0
    If MaSoCT > 0 Then so_kiem_tra = 1
    'Chuyen sang form hoa don
    ExecuteSQL5 "UPDATE KhachHang SET ten = '" + txtVT(7).Text + "',diachi = '" + txtVT(8).Text + "' where sohieu = '" + txtVT(0).Text + "'"

    ' sua lai thong tin khach hang


    Me.MousePointer = 11

    Dim tong_tien_
    tong_tien_ = 0
    With GrdChungtu
        For i = 0 To .Row - 1
            '  For j = 0 To .Cols - 1
            .Row = i
            .col = 7
            '.col = 6
            If Len(.Text) > 0 Then
                tong_tien_ = tong_tien_ + Int(.Text)
            End If
            'MsgBox (.Text)
            '  Next
        Next
    End With


    Select Case Index
    Case 0:
        kiem_tra_so_dong

        MaSoCT = 0
        XoaPhieuTrenManHinh
        '  Command_Click 2
        If Len(shct) > 0 Then
            txt(0).Text = SHCtuMoi(shct)
            txt(0).Text = "..."
        End If

        If OptLoai(8).Value = True Then
            tinhkyhieu
            Mo_thong_tin
        End If
        If OptLoai(8).Value = True Then
            hien_thong_tin_mau_HD
            '            txtVT(2).Text = SelectSQL("SELECT MauSoHD AS F1 FROM chungtu WHERE  maso in (select max(maso) from chungtu)")

        End If
        RFocus CboThang
    Case 1:
        If (kiem_tra_so_dong() = False) Then
            GoTo XongCT
        End If
        If Not KiemTraChungtu Then
            MsgBox "Cã tµi kho¶n chi tiÕt"
            GoTo XongCT
        End If
        LockDB
        If MaSoCT > 0 Then
            If (pPhieu = 0 And pMaBG = 0) Then
                Set rs_chungtu = DBKetoan.OpenRecordset("SELECT MaSo FROM ChungTu WHERE MaCT=" + CStr(MaSoCT), dbOpenSnapshot, dbForwardOnly)
                Do While Not rs_chungtu.EOF
                    chungtu.InitChungtu rs_chungtu!MaSo, 0, "", 0, ngay(0), ngay(1), 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
                    chungtu.XoaChungtu
                    rs_chungtu.MoveNext
                Loop
                rs_chungtu.Close
                Set rs_chungtu = Nothing
            Else
                XoaPhieu MaSoCT
            End If

            If loaict > 8 Then
                SuaChungtuTS MaSoCT
            End If
            MaCT = MaSoCT
            MaSoCT = 0
        Else
            MaCT = Lng_MaxValue("MaCT", "ChungTu" + IIf((pBaoGia = 1 And Chk.Value = 1) Or (pPhieu > 0), "P", "")) + 1
        End If

        bg = Fix(SoPSConLai * Mask_N) <> 0 And loaict = 7
        mhdx = -1
        With GrdChungtu
            For i = 0 To .Rows - 1
                .Row = i
                .col = 1
                sql = .Text
                .col = 8
                If Len(.Text) = 0 Then Exit For
                mtk = CLng5(.Text)
                .col = 21
                MaTP = CLng5(.Text)
                .col = 9
                mn = CLng5(.Text)
                .col = 14

                If Len(.Text) > 0 Or ((pHachToan = 0 Or mn > 0 Or bg) And Not xddu) Or (((Left(sql, 4) = "3331") Or (Left(sql, Len(pVATV)) = pVATV)) And (Len(.Text) > 0 Or loaict = 8)) Then
                    .col = 11
                    mtc = CLng5(.Text)
                    .col = 9
                    mvt = CLng5(.Text)
                    If mvt = 0 And (Not nhieunoco) Then
                        .col = 15
                        mvt = CLng5(.Text)
                    End If
                    .col = 4
                    psnt = Cdbl5(.Text)
                    .col = 6
                    sops = Cdbl5(.Text)
                    If sops = 0 Then
                        loai = 1
                        .col = 7
                        sops = Cdbl5(.Text)
                        If sops = 0 Then
                            If ((Left(sql, Len(pVATV)) = pVATV Or (Left(sql, 2) = "15") And loaict = 1)) Then loai = -1
                        End If
                    Else
                        loai = -1
                    End If
                    .col = 14

                    If nhieunoco Then    'And kiem_tra_333_133 < 2 Then
                        GhiChu = .Text
                        mtk2 = 0
                        mtc2 = 0
                    Else
                        GhiChu = "..."
                        mtc2 = CInt5(.Text)
                        If mtc2 = -1 Then
                            MsgBox "Kh«ng x¸c ®Þnh ®­îc ®èi øng !", vbExclamation, App.ProductName
                            GoTo XongCT
                        End If
                        .col = 13
                        mtk2 = CLng5(.Text)
                        .col = 16
                        psnt2 = Cdbl5(.Text)
                    End If


                    If loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8 Then mn = CboNguon(0).ItemData(CboNguon(0).ListIndex) Else mn = 0
                    If loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8 Or loaict = 9 Then mk = CboNguon(1).ItemData(CboNguon(1).ListIndex) Else mk = 0
                    If loai < 0 Then
                        ' dua thong tin vao chung tu chuan bi luu
                        chungtu.InitChungtu 0, loaict, txt(0).Text, CboThang.ItemData(CboThang.ListIndex), ngay(0), ngay(1), _
                                            mn, mk, txt(1).Text, mtk, mtk2, sops, psnt, psnt2, mvt, GhiChu, CboNguon(2).ItemData(CboNguon(2).ListIndex), txtVT(2).Text, txtVT(3).Text, SoLoNhap, handung
                    Else

                        chungtu.InitChungtu 0, loaict, txt(0).Text, CboThang.ItemData(CboThang.ListIndex), ngay(0), ngay(1), _
                                            mn, mk, txt(1).Text, mtk2, mtk, sops, psnt2, psnt, mvt, GhiChu, CboNguon(2).ItemData(CboNguon(2).ListIndex), txtVT(2).Text, txtVT(3).Text, SoLoNhap, handung
                    End If
                    chungtu.MaCT = MaCT
                Else
                    .col = 10
                    If CInt5(.Text) = 1 Or bg Then
                        .col = 21
                        MaTP = CLng5(.Text)
                        .col = 11
                        mtc = CLng5(.Text)
                        .col = 9
                        mvt = CLng5(.Text)
                        .col = 4
                        psnt = Cdbl5(.Text)
                        .col = 6
                        sops = Cdbl5(.Text)
                        If sops = 0 Then
                            loai = 1
                            .col = 7
                            sops = Cdbl5(.Text)
                        Else
                            loai = -1
                        End If


                        If loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8 Then mn = CboNguon(0).ItemData(CboNguon(0).ListIndex) Else mn = 0
                        If loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8 Then mk = CboNguon(1).ItemData(CboNguon(1).ListIndex) Else mk = 0

                        If loai < 0 Then
                            chungtu.InitChungtu 0, loaict, txt(0).Text, CboThang.ItemData(CboThang.ListIndex), ngay(0), ngay(1), _
                                                mn, mk, txt(1).Text, mtk, 0, sops, psnt, 0, mvt, "...", CboNguon(2).ItemData(CboNguon(2).ListIndex), txtVT(2).Text, txtVT(3).Text, SoLoNhap, handung
                        Else
                            chungtu.InitChungtu 0, loaict, txt(0).Text, CboThang.ItemData(CboThang.ListIndex), ngay(0), ngay(1), _
                                                mn, mk, txt(1).Text, 0, mtk, sops, 0, psnt, mvt, "...", CboNguon(2).ItemData(CboNguon(2).ListIndex), txtVT(2).Text, txtVT(3).Text, SoLoNhap, handung
                        End If
                        chungtu.MaCT = MaCT
                    Else
                        chungtu.MaCT = 0
                    End If
                End If
                If chungtu.MaCT > 0 Then
                    If loaict = 13 Then
                        chungtu.CT_ID = CboNguon(0).ItemData(CboNguon(0).ListIndex)
                    Else
                        .col = 12
                        chungtu.CT_ID = CLng5(.Text)
                        If (mvt > 0 And loaict = 2) Or (chungtu.TkCo.tk_id = TKCNKH_ID) Or (chungtu.tkno.tk_id = TKCNPT_ID) Then chungtu.CT_ID = -Abs(chungtu.CT_ID)
                    End If
                    .col = 17
                    chungtu.makh = CLng5(.Text)
                    .col = 20
                    chungtu.MaKHC = CLng5(.Text)
                    chungtu.CTGS = CboNguon(3).ItemData(CboNguon(3).ListIndex)
                    chungtu.MaTP = MaTP
                    .col = 23
                    m = CLng5(.Text)
                    If mvt > 0 And m > 0 Then
                        If KtraDVT(mvt, m, sql) Then
                            If chungtu.tkno.tk_id = TKVT_ID And loaict = 1 Then
                                chungtu.SoPS2No = QuyDoiTheoDVT1(mvt, m, chungtu.SoPS2No)
                            End If
                            If (((chungtu.TkCo.tk_id = TKDT_ID Or chungtu.TkCo.tk_id = TKGT_ID) And loaict = 8) Or (chungtu.TkCo.tk_id = TKVT_ID And loaict = 2)) Then
                                chungtu.SoPS2Co = QuyDoiTheoDVT1(mvt, m, chungtu.SoPS2Co)
                            End If
                            If (loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8) And chungtu.MaVattu > 0 Then chungtu.dvt = m
                        End If
                    End If
                    chungtu.User_ID = UserID
                    If pCongNoHD > 0 Then chungtu.HanTT = CInt5(txtchungtu(8).Text)
                    If pSongNgu Then chungtu.DienGiaiE = txt(2).Text
                    If pGiaUSD > 0 And (loaict = 1 Or loaict = 2 Or loaict = 8) And mvt > 0 Then
                        .col = 24
                        chungtu.PSUSD = Cdbl5(.Text)
                    End If
                    If pTygia > 0 Then chungtu.tygia = Cdbl5(txtchungtu(7).Text)
                    If loaict = 8 And pNVBH > 0 Then chungtu.MaNV = txt(3).tag
                    If loaict = 8 And Chk.Value = 1 Then chungtu.maloai = 7
                    If loaict = 7 And Chk.Value = 0 Then chungtu.maloai = 8
                    If loaict = 8 And (chungtu.tkno.tk_id = TKGT_ID And chungtu.MaVattu = 0) Then
                        .col = 3
                        chungtu.TLCK = Cdbl5(.Text)
                    End If
                    If (loaict = 8) And chungtu.MaVattu > 0 Then
                        .col = 25
                        chungtu.TLCK = Cdbl5(.Text)
                        .col = 26
                        chungtu.CK = Cdbl5(.Text)
                    End If
                    ' ap dung cho phieu nhap co chiet khau
                    If (loaict = 1) And chungtu.MaVattu > 0 Then
                        .col = 25
                        chungtu.phantramchietkhau = Cdbl5(.Text)
                        .col = 26
                        chungtu.sotienchietkhau = Cdbl5(.Text)
                    Else
                        .col = 25
                        chungtu.phantramchietkhau = ""
                        .col = 26
                        chungtu.sotienchietkhau = ""
                    End If

                    If pSoVV > 0 And CboVV(0).ListIndex >= 0 Then chungtu.MaDT1 = CboVV(0).ItemData(CboVV(0).ListIndex)
                    If pSoVV > 1 And CboVV(1).ListIndex >= 0 Then chungtu.MaDT2 = CboVV(1).ItemData(CboVV(1).ListIndex)
                    If pSoVV > 2 And CboVV(2).ListIndex >= 0 Then chungtu.MaDT3 = CboVV(2).ItemData(CboVV(2).ListIndex)
                    If IsImport = False Then
                        chungtu.NgayCT = MedNgay(0).Text
                        chungtu.NgayGS = MedNgay(1).Text
                    End If

                    ' GHI CHU CHUNG TU ---------------------------------------------------------


                    If chungtu.GhiChungtu(pPhieu) <> 0 Then
                        'RFocus CboThang

                        GoTo XongCT

                    End If
                    'MsgBox CStr(chungtu.MaCT)
                    If pPhieu = 0 And loaict <> 7 Then
                        If mvt > 0 And loaict = 1 Then
                            .col = 22
                            If IsNumeric(.Text) Then
                                ExecuteSQL5 "UPDATE ChungTu SET SoXuat=" + DoiDau(Cdbl5(.Text)) + " WHERE MaSo=" + CStr(chungtu.MaSo)
                            End If
                        End If
                    End If
                    .col = 18
                    If Len(.Text) > 0 Then
                        '  If Int(.Text) > 0 Then
                        j = CInt5(.Text)
                        If mhdx <> j And j <= hdcount Then
                            mhdx = j
                            If mhdx >= 0 Then
                                LayHoaDon HD, j
                                h.MaSo = chungtu.MaSo
                                ' ghi thong tin hoa don
                                ' If (SelectSQL("select count(*) as f1 from hethongtk where maso in (" + str(mtk) + "," + str(mtk2) + ") and (left(sohieu,3) ='133' or left(sohieu,3) ='333')") > 0) Then

                                '      GhiHoaDon IIf((loaict = 7) Or pPhieu > 0, 1, 0)
                                '   End If
                                If (SelectSQL("select count(*) as f1 from hethongtk where maso in (" + str(mtk) + "," + str(mtk2) + ") and (left(sohieu,3) ='133' or left(sohieu,3) ='333')") > 0) Then
                                    If so_dem < 2 Then
                                        GhiHoaDon IIf((loaict = 7) Or pPhieu > 0, 1, 0)
                                    Else
                                        If (SelectSQL("select count(*) as f1 from hethongtk where maso in (" + str(mtk) + "," + str(mtk2) + ") and left(sohieu,3) ='133' ") > 0) Then GhiHoaDon IIf((loaict = 7) Or pPhieu > 0, 1, 0)
                                    End If
                                End If

                            End If
                        End If
                    End If
                    If chkXT.Value = 1 And pPhieu = 0 Then chungtu.XuatThang txtsh(0).tag, txtsh(1).tag
                End If
            Next
            chungtu.MaCT = MaCT

            Dim tss As String
            Dim Makhachhang_lay

            '                    If Len(txtVT(0).Text) > 0 Then
            '                    Makhachhang_lay = SelectSQL("SELECT TOP 1 maso AS F1 FROM khachhang WHERE sohieu = '" + txtVT(0).Text + "'")
            '                    tss = "Update Hoadon set KyHieu ='" + txtVT(1).Text + "',SoHD = '" + txt(0).Text + "',MaKhachHang = " + CStr(Makhachhang_lay) + " where maso in (select maso from chungtu where MaCT = " + CStr(chungtu.MaCT) + ")"
            '                     ExecuteSQL5 tss
            '                    End If
            If Len(txtVT(0).Text) > 0 Then
                Makhachhang_lay = SelectSQL("SELECT TOP 1 maso AS F1 FROM khachhang WHERE sohieu = '" + txtVT(0).Text + "'")
                If SelectSQL("SELECT COUNT(MASO) AS F1 FROM HOADON where maso in (select maso from chungtu where MaCT = " + CStr(chungtu.MaCT) + ")") = 1 Then
                    tss = "Update Hoadon set KyHieu ='" + txtVT(1).Text + "',SoHD = '" + txt(0).Text + "',MaKhachHang = " + CStr(Makhachhang_lay) + " where maso in (select maso from chungtu where MaCT = " + CStr(chungtu.MaCT) + ")"
                Else
                    tss = "Update Hoadon set KyHieu ='" + txtVT(1).Text + "',MaKhachHang = " + CStr(Makhachhang_lay) + " where maso in (select maso from chungtu where MaCT = " + CStr(chungtu.MaCT) + ")"
                End If
                ExecuteSQL5 tss
            End If


            If CmdPhieu(0).Visible Or CmdPhieu(3).Visible Then chungtu.GhiThongtinCT 0, TenTC, DiachiTC, ctgoc, MaKHBH, IIf(pPhieu > 0 Or loaict = 7, 1, 0)
            If CmdPhieu(1).Visible And loaict <> 7 And loaict <> 8 Then chungtu.GhiThongtinCT 1, TenNX, DiaChiNX, "...", 0, IIf(pPhieu > 0 Or loaict = 7, 1, 0)
            If CmdPhieu(1).Visible And (loaict = 7 Or loaict = 8) Then chungtu.GhiThongtinCT 2, TenBH, DiaChiBH, Format(HanTT, Mask_D), 0, IIf(pPhieu > 0 Or loaict = 7, 1, 0)
            If CmdPhieu(2).Visible Then chungtu.GhiThongtinCT 3, unc1, unc2, unc3, 0, IIf(pPhieu > 0 Or loaict = 7, 1, 0)
            shct = chungtu.sohieu
        End With


        If loaict > 8 Then GhiChungtuTS MaCT
        If loaict = 8 And pBaoGia > 0 And pMaBG > 0 And Chk.Value = 0 Then XoaPhieu pMaBG

        'tat ca da ghi xong dua so so va han dung vao trong database
        Dim stt_dong As Integer
        Dim mataikhoan_dung As Integer, masp_dung As Integer
        Dim solo_dung As String, hadung_dung As String, chuoi_dung As String
        ' MsgBox CStr(chungtu.MaCT)
        If SelectSQL("SELECT banthuoc as f1 from license ") = 1 Then
            With GrdChungtu

                For stt_dong = 0 To .Rows - 1
                    .Row = stt_dong
                    .col = 1
                    mataikhoan_dung = SelectSQL("select maso as f1 from hethongtk where sohieu = '" + .Text + "'")
                    .col = 3
                    masp_dung = SelectSQL("select maso as f1 from vattu where sohieu = '" + .Text + "'")
                    .col = 27
                    solo_dung = .Text
                    .col = 28
                    hadung_dung = .Text
                    .col = 4
                    '                 If Len(Trim(.Text)) > 0 Then
                    '                chuoi_dung = "update chungtu set solo = '" + solo_dung + "',handung ='" + hadung_dung + "' where SoPS2No + SoPS2co = " + Replace(.Text, ",", "") + " AND  Mact  = " + CStr(chungtu.MaCT) + " and (MaTKTCNo =" + CStr(mataikhoan_dung) + " or MaTKTCCo =" + CStr(mataikhoan_dung) + ") and mavattu =" + CStr(masp_dung)
                    '                  ExecuteSQL5 chuoi_dung
                    '                  chuoi_dung = "update chungtu set solo = '" + solo_dung + "',handung ='" + hadung_dung + "' where  maso in (select top 1 maso from chungtu where SoPS2No + SoPS2co = " + Replace(.Text, ",", "") + " and sohieu  = '" + chungtu.sohieu + "' and mavattu =" + CStr(masp_dung) + " and len(solo) = 0) "
                    '                If OptLoai(8).Value = True Then
                    '                    ExecuteSQL5 chuoi_dung
                    '                End If
                    '                End If
                    If Len(Trim(.Text)) > 0 Then
                        chuoi_dung = "update chungtu set solo = '" + solo_dung + "',handung ='" + hadung_dung + "' where  maso in (select top 1 maso from chungtu where SoPS2No + SoPS2co = " + CStr(Cdbl5(.Text)) + " and sohieu  = '" + chungtu.sohieu + "' and mavattu =" + CStr(masp_dung) + " and len(solo) = 0) "    'Replace(.Text, ",", "")
                        ExecuteSQL5 chuoi_dung
                        If OptLoai(8).Value = True Then
                            '    ExecuteSQL5 chuoi_dung
                            Dim so_1
                            so_1 = 0    ' SelectSQL("SELECT Sum(Tien_" + CStr(CThangDB(ThangCuoiNamTC)) + ") AS f1 FROM TonKho INNER JOIN KhoHang ON TonKho.MaSoKho=KhoHang.MaSo WHERE TonKho.MaSoKho = " + CStr(CboNguon(1).ItemData(CboNguon(1).ListIndex)) + " and MaVattu=" + CStr(masp_dung) + " GROUP BY TenKho HAVING Sum(Luong_" + CStr(CThangDB(ThangCuoiNamTC)) + ")=0 and abs(Sum(Tien_" + CStr(CThangDB(ThangCuoiNamTC)) + ")) = 1")
                            ' ExecuteSQL5 "update chungtu set sops = sops +" + CStr(so_1) + ", solo = '" + solo_dung + "',handung ='" + hadung_dung + "' where  maso in (select top 1 maso from chungtu where SoPS2No + SoPS2co = " + CStr(Cdbl5(.Text)) + " and sohieu  = '" + chungtu.sohieu + "GV" + "' and mavattu =" + CStr(masp_dung) + " and len(solo) = 0) " 'Replace(.Text, ",", "")
                            ExecuteSQL5 "update chungtu set solo = '" + solo_dung + "',handung ='" + hadung_dung + "' where  maso in (select top 1 maso from chungtu where SoPS2No + SoPS2co = " + CStr(Cdbl5(.Text)) + " and sohieu  = '" + chungtu.sohieu + "GV" + "' and mavattu =" + CStr(masp_dung) + " and len(solo) = 0) "    'Replace(.Text, ",", "")
                        End If
                    End If
                Next
            End With
        End If
        '" + sohieu + "GV'
        '-------------------------------------
        UnlockDB


        If so_kiem_tra > 0 Then
            With Grid2
                For i = 0 To .Cols - 1
                    .col = i
                    If i = 0 Then .Text = txt(0).Text
                    If i = 1 Then .Text = MedNgay(0).Text
                    If i = 2 Then .Text = MedNgay(1).Text
                    If i = 3 Then .Text = txt(1).Text
                    If i = 4 Then .Text = Format(tong_tien_, Mask_0)
                Next
            End With
        Else
            Grid2.AddItem txt(0).Text + Chr(9) + Format(MedNgay(0).Text, Mask_D) + Chr(9) _
                        + Format(MedNgay(1).Text, Mask_D) + Chr(9) + txt(1).Text + Chr(9) + Format(tong_tien_, Mask_0) + Chr(9) + CStr(chungtu.MaCT) + Chr(9) + "0" + Chr(9) + "0", 0

            'bo tam thoi---------------------------------------------------
            'danh_sach_chung_tu
        End If

        If loaict = 0 Then
            'danh_sach_chung_tu
        End If

        XoaPhieuTrenManHinh
        If Len(shct) > 0 Then txt(0).Text = SHCtuMoi(shct)
        'DI CHUYEN LEN TREN

        RFocus MedNgay(0)    ' di chuyen con tro len ngay nhap chung tu
        If OptLoai(8).Value = True Then
            'bo tam thoi-------------------------------------------------


            tinhkyhieu    ' lay so ky hieu cuoi cung cua hoa don , hien tu dong tren form nhap, chi ap dung voi muc ban hang
            Enable_thong_tin
        End If
        ' RFocus txt(0)
        If OptLoai(8).Value = False Then
            txtVT(9).Visible = False
            txtVT(8).Visible = False
            txtVT(7).Visible = False
            txtVT(0).Visible = False
            CboLoai.Visible = False
            CboNguon(1).Visible = False
            Label1(0).Visible = False
            Label1(1).Visible = False
            Label1(3).Visible = False
            Label1(16).Visible = False
        End If
        '  danh dau
        '    danh_sach_chung_tu


        '            With Grid2
        '                    .col = 0
        '                    .Text = txt(0).Text
        '                    .col = 1
        '                    .Text = MedNgay(0).Text
        '                    .col = 2
        '                     .Text = MedNgay(1).Text
        '                    .col = 3
        '            End With
        '    Grid2.Text = "test"


        If pPhieu = 0 And pGhi > 0 Then Me.Hide
    Case 2:
        If MaSoCT > 0 Then
            If loaict = 1 And chkXT.Value = 0 Then
                If Not XoaCTOK(MaSoCT) Then
                    MsgBox "V©t t­ nhËp ®· xuÊt hÕt, kh«ng xo¸ chøng tõ!", vbCritical, App.ProductName
                    GoTo XongCT
                End If
            End If
            If MsgBox("B¹n ®· ch¾c ch¾n xãa chøng tõ nµy ?", vbYesNo + vbCritical, App.ProductName) = vbYes Then
                If pPhieu > 0 Or pMaBG > 0 Then
                    XoaPhieu MaSoCT
                Else
                    If loaict > 8 Then XoaChungtuTS loaict, MaSoCT
                    LockDB
                    Set rs_chungtu = DBKetoan.OpenRecordset("SELECT MaLoai, MaSo, SoPS2Co, MaVattu FROM ChungTu WHERE MaCT=" + CStr(MaSoCT), dbOpenSnapshot)
                    If loaict <> 9 Then
                        If rs_chungtu!maloai = 1 And OutCost > 0 Then
                            Do While Not rs_chungtu.EOF
                                If rs_chungtu!MaVattu > 0 And rs_chungtu!SoPS2Co > 0 Then
                                    rs_chungtu.Close
                                    Set rs_chungtu = Nothing
                                    MsgBox "Chøng tõ nhËp ®· tÝnh gi¸ xuÊt, kh«ng xãa !", vbExclamation, App.ProductName
                                    GoTo XongCT
                                Else
                                    rs_chungtu.MoveNext
                                End If
                            Loop
                            rs_chungtu.MoveFirst
                        End If
                    End If
                    Do While Not rs_chungtu.EOF
                        chungtu.InitChungtu rs_chungtu!MaSo, 0, "", 0, ngay(0), ngay(1), 0, 0, "", 0, 0, 0, 0, 0, 0, "", 0, "", "", "", ""
                        chungtu.XoaChungtu
                        rs_chungtu.MoveNext
                    Loop
                    rs_chungtu.Close
                    Set rs_chungtu = Nothing
                    UnlockDB
                End If
            Else
                GoTo XongCT
            End If
            MaSoCT = 0
        End If

        XoaPhieuTrenManHinh
        GrdChungtu.Row = 0
        If loaict > 8 Then OptLoai(0).Value = True
        'RFocus txt(0)
        RFocus MedNgay(0)
        If pPhieu = 0 And pGhi > 0 Then Me.Hide

        ' danh_sach_chung_tu
        With Grid2
            Grid2.RemoveItem (Grid2.Row)
        End With

    Case 3:
        Unload Me
    Case 4:
        If Not KiemTraChungtu Then
            FrmDsCT.Command_Click (3)    ' in danh sach chung tu
            GoTo XongCT
        End If
        ExecuteSQL5 "DELETE * FROM BaoCaoCP"
        With GrdChungtu
            sql = ""
            For i = 0 To .Rows - 1
                .Row = i
                .col = 1
                If Len(.Text) = 0 Then Exit For
                X.InitTaikhoanSohieu .Text

                .col = 6
                sops = Cdbl5(.Text)
                If sops <> 0 Then
                    mk = -1
                Else
                    .col = 7
                    sops = Cdbl5(.Text)
                    mk = 1
                End If
                If sops <> 0 Then
                    sql = "INSERT INTO BaoCaoCP (MaSo,SoHieu,Ten,Cap, Kq1) VALUES (" + CStr(i) + ",'" + CStr(i) + "','" + X.sohieu + "'+' - '+'" + X.Ten + "'," + CStr(mk) + "," + DoiDau(sops) + ")"
                    ExecuteSQL5 sql
                End If
            Next
        End With

        frmMain.Rpt.ReportFileName = "CHUNGTU.RPT"
        SetRptInfo
        frmMain.Rpt.Formulas(4) = "SoCT='" + txt(0).Text + "'"
        frmMain.Rpt.Formulas(5) = "NgayCT='" + Format(ngay(0), Mask_DR) + "'"
        frmMain.Rpt.Formulas(6) = "NgayGS='" + Format(ngay(1), Mask_DR) + "'"
        frmMain.Rpt.Formulas(7) = "DienGiai='" + txt(1).Text + "'"
        InBaoCaoRPT
    End Select
XongCT:
    Me.Caption = "NhËp chøng tõ kÕ to¸n" + " - " + CStr(pNamTC)
    Me.MousePointer = 0
KT:
    Set chungtu = Nothing
    Set X = Nothing
    If OptLoai(1).Value = True Or OptLoai(2).Value = True Then
        CboNguon(1).Visible = True
        CboLoai.Visible = True
    End If

End Sub
'====================================================================================================
' HiÓn thÞ cöa sæ danh s¸ch chøng tõ vµ hiÓn thÞ chøng tõ ®­îc chän
'====================================================================================================
Public Sub CmdDanhsach_Click(Index As Integer)
    Command_Click (0)
    Label(26).Caption = ""
    cho_hien_thongbao = True

    Dim p As Integer
ChonCT:
    If Index = 0 Then
        MaSoCT = FrmDsCT.ChonCT(p)
    Else
        p = pPhieu
        MaSoCT = FrmDsTC.ChonCT
    End If
    'Set FrmDsTC =
    MaSoCT = 0
    If MaSoCT > 0 Then
        Me.Refresh
        Me.MousePointer = 11
        ' If HienPhieuTrenManHinh(p) < 0 Then GoTo ChonCT
        Me.Caption = "Söa ®æi néi dung chøng tõ"
        Me.MousePointer = 0

    End If
End Sub


Private Sub Command2_Click()
    ExecuteSQL5 "Delete from khachhang where sohieu ='#'"
End Sub

Private Sub danh_sach_chung_tu()

'        FrmDsCT.CboThang(0).Text = "1/" + CStr(pNamTC)
'        FrmDsCT.CboThang(1).Text = "12/" + CStr(pNamTC)
'
'       If OptLoai(0).Value = True Then
'            FrmDsCT.ChkLoai(0).Value = 1
'         Else
'            FrmDsCT.ChkLoai(0).Value = 0
'        End If
'
'
'        If OptLoai(1).Value = True Then
'            FrmDsCT.ChkLoai(1).Value = 1
'         Else
'            FrmDsCT.ChkLoai(1).Value = 0
'        End If
'
'
'
'        If OptLoai(2).Value = True Then
'            FrmDsCT.ChkLoai(2).Value = 1
'         Else
'            FrmDsCT.ChkLoai(2).Value = 0
'        End If
'
'
'        If OptLoai(3).Value = True Then
'         FrmDsCT.ChkLoai(3).Value = 1
'         Else
'         FrmDsCT.ChkLoai(3).Value = 0
'        End If
'
'
'        If OptLoai(4).Value = True Then
'         FrmDsCT.ChkLoai(4).Value = 1
'         Else
'         FrmDsCT.ChkLoai(4).Value = 0
'        End If
'
'
'        If OptLoai(8).Value = True Then
'         FrmDsCT.ChkLoai(8).Value = 1
'         Else
'         FrmDsCT.ChkLoai(8).Value = 0
'        End If
'
'        If OptLoai(9).Value = True Then
'         FrmDsCT.ChkLoai(9).Value = 1
'         Else
'         FrmDsCT.ChkLoai(9).Value = 0
'        End If
'
'
'        If OptLoai(10).Value = True Then
'         FrmDsCT.ChkLoai(10).Value = 1
'         Else
'         FrmDsCT.ChkLoai(10).Value = 0
'        End If
'
'        If OptLoai(11).Value = True Then
'         FrmDsCT.ChkLoai(11).Value = 1
'         Else
'         FrmDsCT.ChkLoai(11).Value = 0
'        End If
'
'
'        If OptLoai(12).Value = True Then
'         FrmDsCT.ChkLoai(12).Value = 1
'         Else
'         FrmDsCT.ChkLoai(12).Value = 0
'        End If
    FrmDsCT.LietKeChungtu_1 "", 0, 0, 0, ""
    '   AddMonthToCbo CboThang
End Sub

Private Sub Command3_Click()
' FBcKt.txtShTk(0).Text = txtShTk(0).Text
' FBcKt.txtShVT(0).Text = txtShVT(0).Text
' FBcKt.OptBC(0).Value = OptBC(0).Value
' FBcKt.OptBC(100).Value = OptBC(100).Value
' FBcKt.OptBC(12).Value = OptBC(12).Value
' FBcKt.CboThang(0).Text = CboThang1(1).Text
' FBcKt.CboThang(1).Text = CboThang1(2).Text
' FBcKt.OptBC(36).Value = OptBC(36).Value
' FBcKt.OptTG(0).Value = True
' FBcKt.txtshkh(0).Text = txtshkh(0).Text
' FBcKt.txtshkh_LostFocus (0)
' 'FBcKt.txtShTk(0).tag = txtShTk(0).tag
' FBcKt.txtShTk_LostFocus (0)
' FBcKt.txtShVT_LostFocus (0)
' 'FBcKt.txtShVT(0).tag = 1
' FBcKt.Command_Click (0)
    FBcTC.Form_Load
    FBcKt.txtShTk(0).Text = txtShTk(0).Text
    FBcKt.txtShVT(0).Text = txtShVT(0).Text
    FBcKt.OptBC(0).Value = OptBC(0).Value
    FBcKt.OptBC(100).Value = OptBC(100).Value
    FBcKt.OptBC(12).Value = OptBC(12).Value
    FBcKt.CboThang(0).Text = CboThang1(1).Text
    FBcKt.CboThang(1).Text = CboThang1(2).Text
    FBcKt.OptBC(36).Value = OptBC(36).Value
    FBcKt.OptBC(9).Value = OptBC(9).Value
    FBcKt.OptBC(14).Value = OptBC(14).Value
    FBcKt.OptTG(0).Value = True
    FBcKt.txtshkh(0).Text = txtshkh(0).Text
    FBcKt.txtshkh_LostFocus (0)
    'FBcKt.txtShTk(0).tag = txtShTk(0).tag
    FBcKt.txtShTk_LostFocus (0)
    FBcKt.txtShVT_LostFocus (0)
    'FBcKt.txtShVT(0).tag = 1
    If OptVAT(5).Value = True Or OptVAT(3).Value = True Or OptVAT(4).Value = True Then
        FBcTC.OptVAT(5).Value = OptVAT(5).Value
        FBcTC.OptVAT(4).Value = OptVAT(4).Value
        FBcTC.OptVAT(3).Value = OptVAT(3).Value
        FBcTC.CboThang(0).Text = CboThang1(1).Text
        FBcTC.CboThang(1).Text = CboThang1(2).Text
        ' FBcTC.OptTG(0).Value = True

        FBcTC.Command_Click (0)
    Else
        FBcKt.Command_Click (0)
    End If

End Sub


Sub GetcustomerByMST(ByVal mst As String, ByVal Name As String, ByVal Address As String)
    Dim rs_ktra As Recordset
    Dim Query As String

    ' T?o truy v?n SQL d? l?y thông tin khách hàng theo MST
    Query = "SELECT Ten, DiaChi, MST FROM KhachHang WHERE MST = '" & mst & "'"

    ' M? Recordset d? l?y thông tin khách hàng
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

    ' Ki?m tra xem Recordset có d? li?u không
    If Not rs_ktra.EOF Then
        ' Hi?n th? thông tin khách hàng
        'MsgBox "Tên: " & rs_ktra.Fields("Ten").Value & vbCrLf & _
         "Ð?a Ch?: " & rs_ktra.Fields("DiaChi").Value & vbCrLf & _
         "MST: " & rs_ktra.Fields("MST").Value"

    Else
        'Khong tim thay khach hang
        Dim getMst As String
        getMst = Right(txtVT(9).Text, 4)
        txtVT(0).Text = getMst
        txtVT(7).Text = Name
        txtVT(8).Text = Address
    End If

    ' Ðóng Recordset
    rs_ktra.Close
    Set rs_ktra = Nothing
End Sub

Public Sub ImportData()
Dim FilePath As String
    Dim xmlDoc As Object
    Dim fDialog As Object
    Dim dlhDonNode As Object
    Dim ttChungNode As Object
    Dim ndhDonNode As Object
    Dim mstNode As Object
    Dim TTNode As Object
    Dim tenNode As Object
    Dim DChiNode As Object
    Dim convertedDate As Date
    Dim cbbThang As String

    ' T?o h?p tho?i m? file
    Set fDialog = CreateObject("MSComDlg.CommonDialog")
    fDialog.ShowOpen
    FilePath = fDialog.fileName

    ' Kh?i t?o MSXML
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False

    ' T?i file XML
    If xmlDoc.Load(FilePath) Then
        ' L?y các node
        Set dlhDonNode = xmlDoc.selectSingleNode("/HDon/DLHDon")
        Set ttChungNode = xmlDoc.selectSingleNode("/HDon/DLHDon/TTChung")
        Set ndhDonNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon")
        Set mstNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/MST")
        Set tenNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/Ten")
        Set DChiNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/DChi")
        Set TTNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/TToan/TgTCThue")
        ' Hi?n th? thông tin
        If Not dlhDonNode Is Nothing Then
            ' txtID.Text = dlhDonNode.Attributes.getNamedItem("Id").Text
        Else
            MsgBox "Không tìm th?y DLHDon."
        End If

        If Not ttChungNode Is Nothing Then
            Dim shDonNode As Object
            Dim shKHHDNode As Object
            Dim shNLapNode As Object

            Set shDonNode = ttChungNode.getElementsByTagName("SHDon")(0)
            Set shKHHDNode = ttChungNode.getElementsByTagName("KHHDon")(0)
            Set shNLapNode = ttChungNode.getElementsByTagName("NLap")(0)

            If Not shDonNode Is Nothing Then
                txt(0).Text = shDonNode.Text
                txtVT(1).Text = shKHHDNode.Text

                If Not shNLapNode Is Nothing Then
                    convertedDate = CDate(shNLapNode.Text)
                    MedNgay(0).Text = Format(convertedDate, "dd/mm/yy")
                    cbbThang = CboThang.Text    ' L?y giá tr? du?c ch?n t? ComboBox
                    Dim monthValue As Integer
                    Dim monthValue2 As Integer
                    monthValue = Month(CDate(cbbThang))
                    monthValue2 = Month(CDate(convertedDate))
                    ' So sánh tháng l?y du?c v?i tháng hi?n t?i
                    If monthValue <> monthValue2 Then
                        Dim dateString As String
                        dateString = "1/" & monthValue & "/" & Year(Date)
                        MedNgay(1).Text = Format(dateString, "dd/mm/yy")
                    End If
                End If
                If Not mstNode Is Nothing Then
                    txtVT(9).Text = mstNode.Text
                    GetcustomerByMST txtVT(9).Text, tenNode.Text, DChiNode.Text
                Else
                    MsgBox "Không tìm th?y MST."
                End If
                txt(1).Text = "Noi dung chi phi quan ly"
                txtchungtu(0).Text = 6422
                txtChungtu_LostFocus (0)
                txtchungtu(5).Text = TTNode.Text
                RFocus txtchungtu(6)
                txtChungtu_KeyPress 6, 13

                txtchungtu(0).Text = 1331
                txtChungtu_LostFocus (0)
                txtchungtu(2).Text = 8
                txtChungtu_LostFocus (2)
                RFocus txtchungtu(6)
                txtChungtu_KeyPress 6, 13

                txtchungtu(0).Text = 1111
                FThuChi.tag = 20
                FThuChi.FThuChiForm = 1

                Timer1.Enabled = True
                txtChungtu_LostFocus (0)

            Else
                MsgBox "Không tìm th?y SHDon."
            End If
        Else
            MsgBox "Không tìm th?y TTChung."
        End If

        If Not ndhDonNode Is Nothing Then
            ' X? lý ndhDonNode n?u c?n
        End If

    Else
        MsgBox "L?i khi t?i file XML: " & xmlDoc.parseError.reason
    End If

End Sub

Public Sub Capnhatlist(ByVal rowIndex As Integer, ByVal colIndex As Integer, ByVal Text As String)

    With fileImportList(rowIndex)
        If colIndex = 5 Then
            .notk = Text
        End If
        If colIndex = 6 Then
            .cotk = Text
        End If
        If colIndex = 3 Then
            .diengiai = Text
        End If
    End With
End Sub
Private Sub Command4_Click()
    Load frmLocImport
    frmLocImport.Show vbModal
End Sub
Public Sub MultiImportData()
OptLoai_Click 0
End Sub

Public Sub cmdReset_Click()
    ' Làm r?ng danh sách fileImportList
End Sub
Private Sub Form_Activate()
    
    Dim ithang As Integer
    'Add item cho cbbThang
    'For i=0

    ' ExecuteSQL5_Themmoi (" delete from chungtu where maso = 3195")
    'MsgBox SelectSQL("SELECT maso AS F1 FROM chungtu where sohieu ='HYY'")


    'Me.Move 3300, 2900
    'If pPhieu = 0 And pGhi = 2 And pFunction <> 10 Then Command_Click 1
    If pFunction = 10 Then
        pFunction = 0
        RFocus txt(0)
    End If
    If (Len(CboLoai.Text) <= 0) Then
        Int_RecsetToCbo "SELECT DISTINCTROW MaSo As F2,SoHieu + ' - '  + TenPhanLoai As F1 FROM PhanLoaiKhachHang WHERE PLCon=0 AND LEFT(SoHieu,1)<>'#' ORDER BY SoHieu", CboLoai
    End If
    If Len(txtVT(0).Text) <= 0 Then txtVT(0).Text = "..."
    If Len(txtVT(1).Text) <= 0 Then txtVT(1).Text = "..."
    If Len(txtVT(7).Text) <= 0 Then txtVT(7).Text = "..."
    If Len(txtVT(8).Text) <= 0 Then txtVT(8).Text = "..."
    If Len(txtVT(9).Text) <= 0 Then txtVT(9).Text = "..."
    If Len(Replace(txtVT(2).Text, ".", "")) <= 0 Then
        '        If OptLoai.item(8).Value = False Then
        txtVT(2).Text = "01GTKT3/001"
    End If
    '  End If
    hien_thong_tin_mau_HD
    If Len(txtVT(3).Text) <= 0 Then txtVT(3).Text = "01GTKT"
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD Nguoimuahang text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD Nguoimuahang text")
    ' che thong tin khi khoi tao form lam viec

    If Len(Replace(Trim(txtVT(0).Text), ".", "")) <= 0 Then
        '  Dong_thong_tin
    End If
    If (Len(txtchungtu(0).Text) <= 0) Then txtchungtu(0).Text = "..."

    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD MauSoHD text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD MauSoHD text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD LoaiHoaDon text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD LoaiHoaDon text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD hinhthucthanhtoan text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD hinhthucthanhtoan text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD sophieudathang text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD sophieudathang text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD chondiengiai text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD chondiengiai text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD solo text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD solo text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD handung datetime")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD handung datetime")
    ' If OptLoai.item(8).Value = True Then
    '    If Len(txtVT(2).Text) <= 0 Then
    '    hien_thong_tin_mau_HD
    '  txtVT(2).Text = SelectSQL("SELECT MauSoHD AS F1 FROM chungtu WHERE  maso in (select max(maso) from chungtu where maloai = 8)")
    '     End If
    '  End If
    '   ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP DROP COLUMN LoaiHD")
    '   ExecuteSQL5_Themmoi ("ALTER TABLE chungtu DROP COLUMN LoaiHD")

    kiem_tra_so_dong
    ' RFocus Command4

    '    If (boolean_kiemtra() = False) Then
    '        If (SelectSQL("SELECT count(*) as F1 FROM ChungTu ") > 2) Then ' so nghiep vu gioi han
    '            Command(0).Enabled = False
    '            Command(1).Enabled = False
    '            Else
    '            Command(0).Enabled = True
    '            Command(1).Enabled = True
    '        End If
    '    End If
End Sub
Function kiem_tra_so_dong() As Boolean
    Dim KT As Boolean
    Dim so
    KT = True
    If (boolean_kiemtra() = False) Then
        Dim rss As Recordset

        Set rss = DBKetoan.OpenRecordset("SELECT DISTINCTROW License.* FROM License", dbOpenSnapshot)
        ' If (SelectSQL("SELECT count(*) as F1 FROM ChungTu ") + rss!sodong > 300) Then ' so nghiep vu gioi han
        If (ban_quyen = 1) Then
            Command(0).Enabled = False
            Command(1).Enabled = False
            KT = False
        Else
            Command(0).Enabled = True
            Command(1).Enabled = True
        End If
        '        so = SelectSQL("SELECT  DateDiff('d',min(NgayCT ), max(NgayCT ))  as F1 from chungtu")
        '        If (so > 90) Then
        '            Command(0).Enabled = False
        '            Command(1).Enabled = False
        '            KT = False
        '            Else
        '            Command(0).Enabled = True
        '            Command(1).Enabled = True
        '        End If
    End If
    kiem_tra_so_dong = KT
End Function
'====================================================================================================
' Xö lý c¸c phÝm nãng

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (Shift And vbAltMask) > 0 Then
        i = -1
        Select Case KeyCode
        Case vbKeyT: i = 0
        Case vbKeyG: i = 1
        Case vbKeyX: i = 2
        Case vbKeyV: i = 3
        Case vbKeyI: i = 4
        Case vbKeyL:
            If mnDD(25).Visible Then mnDD_Click 25
        End Select
        If i >= 0 Then
            If Command(i).Enabled Then
                RFocus Command(i)
                Command_Click i
            End If
        End If
    End If
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyP Then frmMain.mnHT_Click 8
    If KeyCode = vbKeyEscape Then
        bcstop = 1
        Unload Me
    End If
End Sub
Public Sub ListReset()
Set fileImportList = New Collection
End Sub
'====================================================================================================
' Khëi t¹o cöa sæ nhËp
'====================================================================================================
Private Sub Form_Load()
    
    stt = 1
    ListReset
    ColumnSetUp Grid2, 0, 1300, 2
    ColumnSetUp Grid2, 1, 940, 2
    ColumnSetUp Grid2, 2, 940, 2
    ColumnSetUp Grid2, 3, 4610, 0
    ColumnSetUp Grid2, 4, 1690, 1
    ColumnSetUp Grid2, 5, 1, 0
    ColumnSetUp Grid2, 6, 340 + 370, 1
    ColumnSetUp Grid2, 7, 940 + 240, 1
    '   OptLoai(0).BackColor = 8438015
    Dim chi_so As Integer
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD MauSoHD text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtuP  ADD MauSoHD text")

    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD phantramchietkhau text")
    ExecuteSQL5_Themmoi ("ALTER TABLE chungtu  ADD sotienchietkhau text")

    ColumnSetUp GrdChungtu, 0, 340, 2  '340, 2
    ColumnSetUp GrdChungtu, 1, 1060 + 20 - 300, 2
    ColumnSetUp GrdChungtu, 2, 2260 + 600, 0
    ColumnSetUp GrdChungtu, 3, 1300, 2
    ColumnSetUp GrdChungtu, 4, 1060, 1
    ColumnSetUp GrdChungtu, 5, 1300 + 250, 1
    ColumnSetUp GrdChungtu, 6, 1300 + 350, 1
    ColumnSetUp GrdChungtu, 7, 1300 + 350, 1
    ColumnSetUp GrdChungtu, 8, 1, 0                     ' Cét chøa m· TK
    ColumnSetUp GrdChungtu, 9, 1, 0                     ' Cét chøa m· VT
    ColumnSetUp GrdChungtu, 10, 1, 0                     ' "0" nÕu trong b¶ng, "1" nÕu ngoµi b¶ng
    ColumnSetUp GrdChungtu, 11, 1, 0                     ' M· TKTC
    ColumnSetUp GrdChungtu, 12, 1, 0                     ' M· phiÕu nhËp
    ColumnSetUp GrdChungtu, 13, 1, 0                     ' M· TK chi tiÕt ®èi øng
    ColumnSetUp GrdChungtu, 14, 1, 0                     ' M· TKTC doi ung
    ColumnSetUp GrdChungtu, 15, 1, 0                     ' M· vËt t­
    ColumnSetUp GrdChungtu, 16, 1, 0                     ' Sè l­îng
    ColumnSetUp GrdChungtu, 17, 1, 0                     ' MaKH No
    ColumnSetUp GrdChungtu, 18, 1, 0                     ' Ghi chu
    ColumnSetUp GrdChungtu, 19, 1, 0                     ' Ghi chu
    ColumnSetUp GrdChungtu, 20, 1, 0                     ' MaKHCo
    ColumnSetUp GrdChungtu, 21, 1, 0                     ' Ghi chu
    ColumnSetUp GrdChungtu, 22, 1, 0                     ' Ghi chu
    ColumnSetUp GrdChungtu, 23, 1, 0                     ' §.v.t
    ColumnSetUp GrdChungtu, 24, 1, 0                     ' price by usd
    ColumnSetUp GrdChungtu, 25, 340 + 200, 1
    ColumnSetUp GrdChungtu, 26, 940 + 200, 1

    dathuchien = False    ' da thuc hien luu khoi tao

    AddMonthToCbo CboThang
    AddMonthToCbo CboThang1(1)
    AddMonthToCbo CboThang1(2)
    For chi_so = 0 To 1
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next
    SetLoaiEnable = True
    SetLoaiChungtu 0
    MaSoCT = 0

    ' LiÖt kª danh s¸ch kho hµng
    If STDetail Then Int_RecsetToCbo "SELECT MaSo As F2,TenKho As F1 FROM KhoHang ORDER BY TenKho", CboNguon(1)
    If User_Right = 0 Then
a:
        Int_RecsetToCbo "SELECT MaSo As F2,SoHieu+ ' - '+DienGiai As F1 FROM CTGhiSo ORDER BY SoHieu", CboNguon(3)
    Else
        Int_RecsetToCbo "SELECT MaSo As F2,SoHieu As F1 FROM CTGhiSo INNER JOIN User2 ON CTGhiSo.MaSo=User2.CTGS WHERE User2.User=" + CStr(UserID) + " ORDER BY SoHieu", CboNguon(3)
        If CboNguon(3).ListCount = 0 Then GoTo a
    End If
    Int_RecsetToCbo "SELECT DoituongCT.MaSo As F2,(IIF(DoituongCT.MaKhachHang>0,KhachHang.Ten+' - '+DoituongCT.Sohieu+' - ','')+DienGiai) As F1 FROM DoituongCT LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo ORDER BY  KhachHang.Ten,DoituongCT.SoHieu,DienGiai", CboNguon(2)

    VTEnable = True
    Caption = Caption + " - " + CStr(pNamTC)
    OptLoai(1).Enabled = STDetail
    OptLoai(2).Enabled = STDetail
    OptLoai(8).Enabled = STDetail

    OptLoai(9).Enabled = FADetail
    OptLoai(10).Enabled = FADetail
    OptLoai(11).Enabled = FADetail
    OptLoai(12).Enabled = FADetail

    KhongNhapTS = True

    KiemTraUser
    pVAT1 = GetSetting(IniPath, "Invoice", "VAT1", 0)
    pVAT2 = GetSetting(IniPath, "Invoice", "VATCheck", 0)

    Ppthu = GetSetting(IniPath, "Environment", "DInvoice", 2)
    Ppchi = GetSetting(IniPath, "Environment", "CInvoice", 2)
    Ppunc = GetSetting(IniPath, "Environment", "UNC", 2)

    hdcount = -1

    Label(15).Visible = (pSoKT Mod 100 >= 10)
    CboNguon(3).Visible = (pSoKT Mod 100 >= 10)

    Label(16).Visible = pSongNgu
    txt(2).Visible = pSongNgu

    Label(17).Visible = (pTygia > 0)
    txtchungtu(7).Visible = (pTygia > 0)
    pRate = TyGiaNT(0)
    txtchungtu(7).Text = Format(pRate, Mask_2)

    'If pGiaUSD = 0  Then pRate = 1

    Command(1).Enabled = ChoNhapTiep

    If pPhieu > 0 Then Me.Caption = IIf(pNN = 0, "NhËp phiÕu", "Template Voucher")

    If frmMain.Command(4).Visible Then
        If pPhieu = 1 Then
            OptLoai(3).Enabled = False
            For chi_so = 9 To 12
                OptLoai(chi_so).Enabled = False
            Next
        Else
            For chi_so = 0 To 2
                CmdPhieu(chi_so).Enabled = False
            Next
        End If
    End If

    txtchungtu(8).Visible = (pCongNoHD > 0)
    Label(22).Visible = (pCongNoHD > 0)



    If pNVBH = 0 Then txt(1).Width = 7935

    txt(1).Width = 5400

    mnDD(26).Visible = (pSoVV > 0)

    For chi_so = 1 To pSoVV
        LbTT(chi_so - 1).Visible = True
        CboVV(chi_so - 1).Visible = True
        mnDD(26 + chi_so).Visible = True
        Int_RecsetToCbo "SELECT MaSo As F2,DienGiai As F1 FROM DoituongCT" + CStr(chi_so) + " ORDER BY DoituongCT" + CStr(chi_so) + ".DienGiai", CboVV(chi_so - 1)
    Next
    hien_thong_tin_mau_HD
    SetFont Me

LoiNgay:


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnPU
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i



    For i = 0 To 4
        If OptLoai(i).Value = False Then OptLoai(i).BackColor = &H80FF80       '&H80000003
    Next
    For i = 8 To 12
        If OptLoai(i).Value = False Then OptLoai(i).BackColor = &H80FF80       '&H80000003

    Next

    ' For i = 0 To 4
    '        If OptLoai(i).Value = False Then OptLoai(i).BackColor = &HC0FFC0    '&H80000003
    '        Next
    '        For i = 8 To 12
    '        If OptLoai(i).Value = False Then OptLoai(i).BackColor = &HC0FFC0    '&H80000003
    '
    '        Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    If pNghiepVu = NV_TANG Then
        For i = 0 To tscount
            XoaTaiSan MaTS(i)
        Next
    End If

    pGhichungtu = 0
    pNghiepVu = 0
    pMaTaiSan = 0
    Set taikhoan = Nothing
    Set vattu = Nothing
    Set ckh = Nothing
    Set tp = Nothing
    Erase HD
    On Error Resume Next
    Unload FrmDsCT
    Unload FrmDsTC
    On Error GoTo 0
End Sub

Private Sub GrdChungtu_DblClick()
    Dim i As Integer, r As Integer

    With GrdChungtu
        .col = 1
        r = .Row
        If Len(.Text) = 0 Then Exit Sub
        CmdChitiet.tag = .Row
        taikhoan.InitTaikhoanSohieu .Text
        If FADetail And (taikhoan.tk_id = TSCD_ID Or taikhoan.tk_id = KHTSCD_ID) Then Exit Sub
        For i = 0 To 6
            .col = i + 1
            txtchungtu(i).Text = .Text
        Next
        If taikhoan.tk_id = GTGTKT_ID Or taikhoan.tk_id = GTGTPN_ID Or taikhoan.tk_id = TTDB_ID Then
            .col = 18
            If IsNumeric(.Text) Then
                LayHoaDon HD, CInt5(.Text)
                '                BotHoaDon HD, CInt5(.Text), hdcount
                h.MaSo = 0
            End If
        End If
        txtChungtu_LostFocus 2
        .col = 23
        If loaict = 1 And CInt5(.Text) > 0 Then CboNT(2).ListIndex = 1
        If loaict = 8 And pChietKhau > 0 Then
            .col = 25
            txtchungtu(9).Text = .Text
            .col = 26
            txtchungtu(10).Text = .Text
        End If
        If pGiaUSD > 0 Then
            .col = 24
            txtchungtu(11).Text = .Text
        End If
        .col = 27
        frmSoLo.txtsolo.Text = .Text
        .col = 28
        frmSoLo.txtngaynhap.Text = IIf(Len(.Text) <= 0, "01/01/11", .Text)

        .Row = r
        .RemoveItem .Row
        If .Rows < .tag Then .Rows = .tag
        If loaict = 2 Or loaict = 8 Then
            xddu = SetDoiUng(1)
            If Not xddu Then xddu = SetDoiUng
        Else
            xddu = SetDoiUng
            If Not xddu Then xddu = SetDoiUng(1)
        End If
    End With
    RFocus txtchungtu(0)
End Sub

Private Sub GrdChungtu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then GrdChungtu_DblClick
End Sub

Private Sub GrdChungtu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdChungtu, GrdChungtu.col
    End If
End Sub


Private Sub Grid2_Click()
    Dim MaCTChon
    With Grid2
        .col = 5
        If Len(.Text) = 0 Then
            MaCTChon = 0
            Command_Click (0)
        Else
            MaCTChon = CLng5(.Text)
            MaSoCT = MaCTChon
            HienPhieuTrenManHinh (0)
        End If
        .col = 0
        Label(27).Caption = "Dßng sè: " + str(.Row + 1)

    End With

End Sub

Private Sub MedNgay_GotFocus(Index As Integer)
    AutoSelect MedNgay(Index)
End Sub


Private Sub MedNgay_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Index
        Case 0:
            RFocus MedNgay(1)
        Case 1:
            If Len(Replace(Trim(txt(0).Text), ".", "")) <= 0 Then txt(0).Text = "..."
            RFocus txt(0)

        End Select
    End If
End Sub

Private Sub MedNgay_LostFocus(Index As Integer)

    Dim ngay As Date
    Dim Ngaykt As Date
    Dim st As String
    If Index = 0 Then
        st = Day(MedNgay(0).Text)
        If Day(MedNgay(0).Text) < 10 Then
            st = "0" + Day(MedNgay(0).Text)
        End If

        If Len(CboThang.Text) <= 6 Then
            If IsDate(st + "/0" + CboThang.Text) Then
                ngay = st + "/0" + CboThang.Text
            Else
                ngay = "01/0" + CboThang.Text
            End If
        Else
            If Not IsDate(st + "/" + CboThang.Text) Then
                ngay = "01/" + CboThang.Text
            Else
                ngay = st + "/" + CboThang.Text
            End If
        End If
        If MaSoCT = 0 Then MedNgay(1).Text = ngay
    End If

    st = Day(MedNgay(0).Text)
    If Day(MedNgay(0).Text) < 10 Then
        st = "0" + Day(MedNgay(0).Text)

    End If


    If Len(CboThang.Text) <= 6 Then
        If IsDate(st + "/0" + CboThang.Text) Then

            Ngaykt = st + "/0" + CboThang.Text
        Else
            ngay = "01/0" + CboThang.Text
        End If
    Else
        If Not IsDate(st + "/" + CboThang.Text) Then
            Ngaykt = "01" + "/" + CboThang.Text
        Else

            Ngaykt = st + "/" + CboThang.Text
        End If
    End If


    '           Dim ngay As Date1
    '           If Index = 0 Then
    '           If Len(CboThang.Text) <= 6 Then
    '           ngay = "01/0" + CboThang.Text
    '           Else
    '           ngay = "01/" + CboThang.Text
    '           End If
    '           MedNgay(1).Text = ngay
    '           End If
    '/////////////////
    Dim ngayx As Date
    Label(26).Caption = ""
    If IsDate(MedNgay(Index).Text) Then
        ngayx = CDate(MedNgay(Index).Text)
        If Year(ngayx) <> pNamTC Then
            Label(26).Caption = "Ngµy chøng tõ kh¸c n¨m tµi chÝnh !"
            Ngaykt = "01/" + CboThang.Text
            MedNgay(1).Text = Ngaykt
            'MsgBox "Ngµy chøng tõ kh¸c n¨m tµi chÝnh !", vbExclamation, App.ProductName
            '            If Index = 1 Then RFocus txtVT(1)
        Else
            If (Month(MedNgay(Index).Text) <> Month(Ngaykt)) Then
                Label(26).Caption = "Ngµy chøng tõ kh¸c ngµy ghi sæ !"
                Ngaykt = "01/" + CboThang.Text
                MedNgay(1).Text = Ngaykt
            End If

        End If

        '            ngay(Index) = ngayx
        '            If Index = 0 Then
        '                MedNgay(1).Text = MedNgay(0).Text
        '                ngay(1) = ngay(0)
        '                If NgayDauThangMoi > 0 Then
        '                     If Day(ngay(0)) >= NgayDauThangMoi And Month(ngay(0)) < 12 Then CboThang.ListIndex = Month(ngay(1)) Else CboThang.ListIndex = Month(ngay(1)) - 1
        '                End If
        '            End If
        ' Else
        '    RFocus MedNgay(Index)
    End If


    '//////////////////////////////
    If Index = 1 Then
        If Len(Replace(Trim(txt(0).Text), ".", "")) <= 0 Then txt(0).Text = "..."
        ' RFocus txt(0)

        'End If
    End If

End Sub

'====================================================================================================
' §Æt chÕ ®é nhËp cho lo¹i phiÕu t­¬ng øng
'====================================================================================================
Public Sub OptLoai_Click(Index As Integer)
   

    ' chon nut khau hao
    If Index = 5 Then
        btnOpenexe_Click
    End If
    txtVT(2).Text = ""
    OptLoai(Index).Value = True
    Dim i
    i = 0
    For i = 0 To 4
        OptLoai(i).BackColor = &H80FF80    '&HC0FFC0    '&H80000003
    Next
    For i = 8 To 12

        OptLoai(i).BackColor = &H80FF80    ' &HC0FFC0    ' &H80000003
    Next
    OptLoai(Index).BackColor = 8438015

    ' set thong so static
    '  Me.Move 3300, 2900
    SetLoaiChungtu Index
    RFocus CboThang
    '  RFocus MedNgay(0)
    ' che thong tin cua khach hang
    ' loc them loai chung tu
    Dong_thong_tin
    txtchungtu(0).Text = "..."
    txtchungtu(0).SelStart = 0
    txtchungtu(0).SelLength = Len(txtchungtu(0).Text)
    If Index = 8 Then
        Dim sql
        sql = "SELECT kyhieu as F1 from hoadon where maso in (select max(maso) from hoadon where maso in (select maso from chungtu where maloai = 8))"
        Dim rs_chungtu As Recordset
        Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        If rs_chungtu.recordCount > 0 Then txtVT(1).Text = rs_chungtu!f1
        Enable_thong_tin
        ' da bo txtVT(2).Text = SelectSQL("SELECT MauSoHD AS F1 FROM chungtu WHERE  maso in (select max(maso) from chungtu where maloai = 8)")
        rs_chungtu.Close

        ' tam thoi thi de the nay
        '       txtVT(1).Text = "..."
    End If
    txtVT(2).Text = SelectSQL("SELECT MauSoHD AS F1 FROM chungtu WHERE  maso in (select max(maso) from chungtu)")
    If OptLoai(1).Value = True Then
        CboLoai.Visible = True
    End If
    ' danh_sach_chung_tu

    RFocus CboThang

End Sub
Private Sub OptLoai_LostFocus(Index As Integer)
    FrmDsCT.CboThang(0).Text = "1/" + CStr(pNamTC)
    FrmDsCT.CboThang(1).Text = "12/" + CStr(pNamTC)

    If OptLoai(0).Value = True Then
        FrmDsCT.ChkLoai(0).Value = 1
    Else
        FrmDsCT.ChkLoai(0).Value = 0
    End If


    If OptLoai(1).Value = True Then
        FrmDsCT.ChkLoai(1).Value = 1
    Else
        FrmDsCT.ChkLoai(1).Value = 0
    End If



    If OptLoai(2).Value = True Then
        FrmDsCT.ChkLoai(2).Value = 1
    Else
        FrmDsCT.ChkLoai(2).Value = 0
    End If


    If OptLoai(3).Value = True Then
        FrmDsCT.ChkLoai(3).Value = 1
    Else
        FrmDsCT.ChkLoai(3).Value = 0
    End If


    If OptLoai(4).Value = True Then
        FrmDsCT.ChkLoai(4).Value = 1
    Else
        FrmDsCT.ChkLoai(4).Value = 0
    End If


    If OptLoai(8).Value = True Then
        FrmDsCT.ChkLoai(8).Value = 1
    Else
        FrmDsCT.ChkLoai(8).Value = 0
    End If

    If OptLoai(9).Value = True Then
        FrmDsCT.ChkLoai(9).Value = 1
    Else
        FrmDsCT.ChkLoai(9).Value = 0
    End If


    If OptLoai(10).Value = True Then
        FrmDsCT.ChkLoai(10).Value = 1
    Else
        FrmDsCT.ChkLoai(10).Value = 0
    End If

    If OptLoai(11).Value = True Then
        FrmDsCT.ChkLoai(11).Value = 1
    Else
        FrmDsCT.ChkLoai(11).Value = 0
    End If


    If OptLoai(12).Value = True Then
        FrmDsCT.ChkLoai(12).Value = 1
    Else
        FrmDsCT.ChkLoai(12).Value = 0
    End If
    danh_sach_chung_tu

    AddMonthToCbo CboThang
End Sub

Private Sub OptLoai_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OptLoai(Index).Value = True
    RFocus CboThang
End Sub

Private Sub OptLoai_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i
    '        For i = 0 To 4
    '        If OptLoai(i).Value = False Then OptLoai(i).BackColor = &HC0FFC0    ' &H80000003
    '        Next
    '        For i = 8 To 12
    '        If OptLoai(i).Value = False Then OptLoai(i).BackColor = &HC0FFC0    '&H80000003
    '
    '        Next
    For i = 0 To 4
        If OptLoai(i).Value = False Then OptLoai(i).BackColor = &H80FF80       ' &H80000003
    Next
    For i = 8 To 12
        If OptLoai(i).Value = False Then OptLoai(i).BackColor = &H80FF80       '&H80000003

    Next

    OptLoai(Index).BackColor = 8438015
End Sub

Private Sub SSCmdV_Click()
    Dim CoIn As Boolean

    Me.MousePointer = 11
    CoIn = False
    SetRptInfo
    If vattu.MaSo > 0 Then
        CoIn = InTheKho2(CboNguon(1).ItemData(CboNguon(1).ListIndex), vattu.MaSo, CboThang.ItemData(CboThang.ListIndex), CboThang.ItemData(CboThang.ListIndex), True, 0, "", 0, pNN)
    Else
        If taikhoan.MaSo > 0 Then
            If pPQTK > 0 And User_Right <> 0 Then
                If Not taikhoan.ChoNhap Then GoTo KT
            End If
            If (taikhoan.MaSo = taikhoan.MaTC Or taikhoan.MaTC = 0) Then
                CoIn = InSocaiTk(taikhoan, CboThang.ItemData(CboThang.ListIndex), CboThang.ItemData(CboThang.ListIndex), ngay(0), ngay(1), True, "", 0, 0, pNN)
            Else
                CoIn = InSoChitiet(taikhoan, CboThang.ItemData(CboThang.ListIndex), CboThang.ItemData(CboThang.ListIndex), ngay(0), ngay(1), True, "", 0, 0, pNN)
            End If
        End If
    End If
    If CoIn Then InBaoCaoRPT
KT:
    Me.MousePointer = 0
End Sub

Private Sub Timer1_Timer()
' Khi Timer dã h?t th?i gian (sau 2 giây)
    Command_Click 1

    Timer1.Enabled = False
    Command_Click 0
    Timer3.Enabled = True
End Sub

Private Sub Xuly154()
    If Not rs_ktra154c.EOF Then

        txtchungtu(0).Text = rs_ktra154c!tkno
        txtChungtu_LostFocus (0)
        'xu ly 154
        If rs_ktra154c!MaCT <> "" Then
            Dim recordCount As Long
            recordCount = rs_ktra154c.recordCount
            txtchungtu(2).Text = rs_ktra154c!MaCT
            txtChungtu_LostFocus (2)
            txtchungtu(5).Text = rs_ktra154c!SoLuong * rs_ktra154c!dongia
            txtChungtu_LostFocus (5)
            ' RFocus txtchungtu(6)
            txtChungtu_KeyPress 6, 13
            rs_ktra154c.MoveNext
            timer154.Enabled = True
            'Xu ly 152
        Else
            txtchungtu(2).Text = rs_ktra154c!sohieu
            txtChungtu_LostFocus (2)
            txtchungtu(3).Text = rs_ktra154c!SoLuong
            txtChungtu_LostFocus (3)
            RFocus txtchungtu(4)

            txtchungtu(4).Text = rs_ktra154c!dongia
            txtChungtu_LostFocus (4)
            RFocus txtchungtu(5)
            txtChungtu_LostFocus (5)
            txtChungtu_KeyPress 6, 13
            rs_ktra154c.MoveNext
            timer154.Enabled = True
        End If
    End If
End Sub
Private Sub timer154_Timer()
    timer154.Enabled = False
    If Not rs_ktra154c.EOF Then
        Xuly154
    Else
        With fileImportList(IndexFirst)
            If .notk <> "5111" Then
                'txtchungtu(0) = .ThueTK
                txtchungtu(0).Text = "1331"
            Else
                txtchungtu(0).Text = "33311"
            End If

            Dim myDate As Date
            myDate = CDate(.ngay)
            CboThang.Text = Month(myDate) & "/" & Year(myDate)
            MedNgay(0).Text = .ngay
            MedNgay(0).Text = .ngay
            If Month(myDate) <> Month(Now) Then
                MedNgay(1).Text = DateSerial(Year(Date), Month(Date), 1)
            Else
                MedNgay(1).Text = Format(.ngay, "dd/mm/yy")
            End If

            txtChungtu_LostFocus (0)
            txtchungtu(2).Text = .VAT
            txtChungtu_LostFocus (2)
            txtChungtu_KeyPress 6, 13
            txtchungtu(0) = .cotk
            txtChungtu_LostFocus (0)
            FThuChi.FThuChiForm = 1
            txtChungtu_KeyPress 6, 13
        End With

    End If
End Sub

Private Sub timer1542_Timer()
    timer1542.Enabled = False
    FThuChi.Command_Click
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    IdDuyet = IdDuyet + 1
    Dim item2 As ClsFileImport

    If IdDuyet <= fileImportList.count Then
        Set item2 = fileImportList(IdDuyet)
        DuyetItemList item2.patTH
    Else
        ' Code to execute if the condition is not met
        MsgBox "Duyet xong"
        FThuChi.FThuChiForm = 0
        
    End If
End Sub



Private Sub Timer3_Timer()
    Timer3.Enabled = False
    IndexFirst = IndexFirst + 1
    Dim item2 As ClsFileImport

     
    If IndexFirst <= fileImportList.count Then
        Set item2 = fileImportList(IndexFirst)
        Xulyimport item2
    Else
        ' Code to execute if the condition is not met

        MsgBox "Duyet xong"
        FThuChi.FThuChiForm = 0
    End If
End Sub

Private Sub timer3311_Timer()
    timer3311.Enabled = False
    DoneSetup
    
End Sub

Private Sub Timer4_Timer()

    If Not rs_ktra152.EOF Then
        Timer4.Enabled = False
        txtchungtu(0).Text = tempchungtu
        txtChungtu_LostFocus (0)
        RFocus txtchungtu(2)
        txtchungtu(2).Text = rs_ktra152!sohieu
        txtChungtu_LostFocus (2)
        txtchungtu(3).Text = rs_ktra152!SoLuong
        txtChungtu_LostFocus (3)
        RFocus txtchungtu(4)
        txtchungtu(4).Text = rs_ktra152!dongia
        txtChungtu_LostFocus (4)
        RFocus txtchungtu(5)
        txtChungtu_LostFocus (5)
        'RFocus txtchungtu(6)
        txtChungtu_KeyPress 6, 13
        rs_ktra152.MoveNext
        Timer4.Enabled = True
    Else
        Timer4.Enabled = False
        'Xu li tai khoan chiec khau .....
        Dim Query As String
        With fileImportList(IndexFirst)
            Query = "SELECT * FROM tbimportdetail WHERE ParentId='" & .id & "' AND DVT = 'Exception' ORDER BY DonGia ASC"
        End With
        Set rs_ktra152 = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
        If Not rs_ktra152.EOF Then
            ' Duy?t qua t?t c? các b?n ghi
            Do While Not rs_ktra152.EOF
                RFocus txtchungtu(0)
                txtchungtu(0).Text = rs_ktra152!sohieu
                txtChungtu_LostFocus (0)
                RFocus txtchungtu(2)
                If rs_ktra152!sohieu = "6422" Then
                    RFocus txtchungtu(5)
                    txtchungtu(5).Text = rs_ktra152!dongia
                    txtChungtu_LostFocus (5)
                    txtChungtu_KeyPress 6, 13
                End If
                If rs_ktra152!sohieu = "711" Then
                    RFocus txtchungtu(6)
                    txtchungtu(6).Text = rs_ktra152!dongia
                    txtChungtu_KeyPress 6, 13
                End If

                rs_ktra152.MoveNext
            Loop
        End If

        'Xu ly cac  tai khoan 1331, 1111
        With fileImportList(IndexFirst)
            If .notk <> "5111" Then
                'txtchungtu(0) = .ThueTK
                txtchungtu(0).Text = "1331"
            Else
                txtchungtu(0).Text = "33311"
            End If

            Dim myDate As Date
            myDate = CDate(.ngay)
            CboThang.Text = Month(myDate) & "/" & Year(myDate)
            MedNgay(0).Text = .ngay
            MedNgay(0).Text = .ngay
            If Month(myDate) <> Month(Now) Then
                MedNgay(1).Text = DateSerial(Year(Date), Month(Date), 1)
            Else
                MedNgay(1).Text = Format(.ngay, "dd/mm/yy")
            End If

            txtChungtu_LostFocus (0)
            txtchungtu(2).Text = .VAT
            txtChungtu_LostFocus (2)
            txtChungtu_KeyPress 6, 13
            txtchungtu(0) = .cotk
            txtChungtu_LostFocus (0)
            txtChungtu_KeyPress 6, 13
            Timer5.Enabled = True
        End With

    End If

End Sub

Private Sub Timer5_Timer()
    Timer5.Enabled = False
    Command_Click 1
    Timer3.Enabled = True
End Sub

Private Sub timerImport_Timer()
    timerImport.Enabled = False
    btnImport_Click
End Sub

Private Sub txt_Click(Index As Integer)
    Label(26).Caption = ""
End Sub

Private Sub txt_GotFocus(Index As Integer)
    AutoSelect txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Label(26).Caption = ""
    Select Case Index
    Case 0:
        If KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 42 Then KeyAscii = 0
    Case 3:
        If KeyAscii = 63 Then
            txt(3).Text = FrmNhanVien.ChonNV(txt(3).Text)
            ' RFocus txt(3)
        End If
    End Select

    If KeyAscii = 13 And KHDetail Then
        Select Case Index
        Case 0:
            If Len(Replace(Trim(txtVT(1).Text), ".", "")) <= 0 Then txtVT(1).Text = "..."
            Dim ttt
            ttt = "select kyhieu from hoadon where max(maso) "
            RFocus txtVT(1)
        Case 1:
            If CboNguon(1).Visible = True Then
                RFocus CboNguon(1)
            ElseIf txt(3).Visible = True Then
                RFocus txt(3)
            Else:
                Disnable_thong_tin    'an thong tin truoc khi chuyen xuong dien giai
                If (CboLoai.Visible = True) Then
                    RFocus CboLoai
                Else
                    RFocus txtchungtu(0)
                End If

            End If
        Case 3:
            Disnable_thong_tin    'an thong tin truoc khi chuyen xuong dien giai
            RFocus txtchungtu(0)
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Label(26).Caption = ""


    Dim L As Long, sh As String
    Select Case Index
    Case 0:
        txt(0).Text = UCase(txt(0).Text)
        If Len(txt(Index).Text) = 0 Then
            txt(Index).Text = "..."
        Else
            sh = IIf(Chk.Value = 1, "P", "")
            L = Len(txt(0).Text)
            If Index = 0 And L > 0 And MaSoCT = 0 Then
                If Not IsNumeric(txt(0).Text) Then
                    shct = SelectSQL("SELECT TOP 1 SoHieu AS F1 FROM ChungTu" + sh + " WHERE Len(SoHieu)>" + CStr(L) + " AND IsNumeric(Right(SoHieu,Len(SoHieu)-" + CStr(L) + ")) AND SoHieu LIKE'" + txt(0).Text + "*' AND ThangCT=" + CStr(CboThang.ItemData(CboThang.ListIndex)) + " ORDER BY SoHieu DESC")
                    If shct <> "0" Then txt(0).Text = SHCtuMoi(shct)
                End If
            End If
            If Index = 0 And txt(0).Text <> "..." And MaSoCT = 0 Then
                If SelectSQL("SELECT DISTINCTROW Count(MaSo) AS F1 FROM ChungTu" + sh + " WHERE SoHieu = '" + txt(0).Text + "' AND MaCT<>" + CStr(MaSoCT) + IIf(pTrungSoHieuKhacThang = 0, "", " AND ThangCT=" + CStr(CboThang.ItemData(CboThang.ListIndex))), dbOpenSnapshot) > 0 Then
                    ErrMsg er_SHChTu
                    RFocus txt(0)
                End If
            End If
        End If
        If MaSoCT = 0 Then
            If OptLoai(8).Value = True Then
                Dim rs_chungtu As Recordset
                Dim sql
                sql = "SELECT kyhieu as F1 from hoadon where maso in (select max(maso) from hoadon where maso in (select maso from chungtu where maloai = 8))"

                Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
                If rs_chungtu.recordCount > 0 Then txtVT(1).Text = rs_chungtu!f1
                rs_chungtu.Close
                'rs_chungtu = Null
            End If
        End If
        If Len(Replace(Trim(txtVT(1).Text), ".", "")) <= 0 Then txtVT(1).Text = "..."
        ' moi them vao
        If Len(Replace(txt(Index).Text, ".", "")) = 0 Then
            MsgBox "B¹n ph¶i nhËp sè hãa ®¬n hoÆc sè hiÖu b¹n tù lËp"
            RFocus txt(Index)

        End If

    Case 3:
        LBNV.Caption = TenNV(txt(3).Text, L)
        txt(3).tag = L
    Case 1:
        ' Disnable_thong_tin
        '  RFocus txtchungtu(0)
    End Select

End Sub

Private Sub txtchungtu_Change(Index As Integer)
    If IsNumeric(txtchungtu(6).Text) Or Len(Trim(txtchungtu(6).Text)) <= 0 Then txttrunggian.Text = txtchungtu(6).Text
    txtchungtu(0).Text = Replace(txtchungtu(0).Text, "?", "")
    txtchungtu(2).Text = Replace(txtchungtu(2).Text, "?", "")
    If (txtchungtu(6).Text = "+") Then txtchungtu(6).Text = txttrunggian.Text    ' Format(txttrunggian.Text, Mask_2)
End Sub

Private Sub txtchungtu_DblClick(Index As Integer)
    Select Case Index
    Case 0:
        txtchungtu(0).Text = FrmTaikhoan.ChonTk(txtchungtu(0).Text)
        Me.MousePointer = 0
        txtChungtu_LostFocus 0
        ' ham xu ly chung tu
        If txtchungtu(3).Enabled = True Then
            RFocus txtchungtu(3)
        Else
            RFocus txtchungtu(5)
        End If
    Case 2:
        If (taikhoan.tk_id <> TKCNKH_ID And taikhoan.tk_id <> TKCNPT_ID) Then
            txtchungtu(2).Text = FrmVattu.ChonVattu(txtchungtu(2).Text)
        Else
            txtchungtu(2).Text = FrmKhachHang.ChonKhachHang(txtchungtu(2).Text)
        End If
        Me.MousePointer = 0
        txtChungtu_LostFocus 2
        VTEnable = True
        ' bat chuc nang
        If txtchungtu(4).Enabled = True Then
            RFocus txtchungtu(4)
        Else
            RFocus txtchungtu(5)
        End If
        RFocus txtchungtu(5)

    Case 6, 8:
        If (Left(txtchungtu(0).Text, 4) = "1331") Then
            cho_hien_vat = True
            If Len(Trim(txtchungtu(1).Text)) > 0 Then
                CmdChitiet_chon
            End If
        End If
    End Select
End Sub

Private Sub txtChungtu_GotFocus(Index As Integer)
    AutoSelect txtchungtu(Index)
    If Len(Trim(txtchungtu(0).Text)) <= 0 Then

        '  RFocus txtchungtu(0)
    End If
End Sub
Private Sub xuly_ham_enter_chungtu(Index As Integer)
    Dim luong As Double, tien As Double, i As Integer, j As Integer, v As Double, sh As String, tien2 As Double
    Select Case Index
    Case 0:    ' So hieu tai khoan
        taikhoan.InitTaikhoanSohieu txtchungtu(0).Text
        txtchungtu(1).Text = IIf(pNN = 0, taikhoan.Ten, taikhoan.TenE)
        If (taikhoan.tk_id = TTDB_ID) Or (taikhoan.tk_id = GTGTKT_ID) Or (taikhoan.tk_id = GTGTPN_ID) Or (taikhoan.tk_id = TKCNKH_ID) Or (taikhoan.tk_id = TKCNPT_ID) Then
            If ((taikhoan.tk_id = TKCNKH_ID) Or (taikhoan.tk_id = TKCNPT_ID)) And hdcount >= 0 Then
                ckh.InitKhachHangMaSo HD(hdcount).MaKhachHang
                txtchungtu(2).Text = ckh.sohieu
                If txtchungtu(3).Enabled = True Then
                End If
            End If
            CboNT(0).Visible = False
            txtchungtu(4).Enabled = False
            txtchungtu(3).Enabled = False
            CboNT(3).Visible = False

            If ((taikhoan.tk_id = TKCNKH_ID) Or (taikhoan.tk_id = TKCNPT_ID)) And (Not KHDetail) Then
                txtchungtu(2).Enabled = False
            Else
                If pNhapDoiTuong > 0 And ((taikhoan.tk_id = TKCNKH_ID) Or (taikhoan.tk_id = TKCNPT_ID)) Then
                    txtchungtu(2).Enabled = False
                    CboNT(3).Visible = True
                    LaySohieuDoiTuong2 2, ""
                    If hdcount >= 0 Then
                        CboNT(3).Text = ckh.sohieu + " - " + ckh.Ten
                    End If
                Else
                    txtchungtu(2).Enabled = True
                    RFocus txtchungtu(2)
                End If
            End If
        Else
            '////////////////////
            txtchungtu(4).Enabled = False
            txtchungtu(2).Enabled = False
            CboNT(3).Visible = False
            If taikhoan.MaSo > 0 Then
                If (taikhoan.tk_id <> TKVT_ID And taikhoan.tk_id <> TKDT_ID And taikhoan.tk_id <> TKGT_ID) Or (Not STDetail) Or (taikhoan.tk_id = TKDT_ID And loaict <> 8) Or (taikhoan.tk_id = TKVT_ID And loaict <> 1 And loaict <> 2 And loaict <> 8) Then
                    txtchungtu(2).tag = IIf(taikhoan.MaNT > 0, 1, 0)
                    txtchungtu(2).Text = ""
                    If taikhoan.MaNT = 0 Then
                        CboNT(0).Visible = False
                        txtchungtu(3).Text = "0"
                        txtchungtu(3).Enabled = False
                        CboNT(1).Visible = False
                    Else
                        Int_RecsetToCbo "SELECT NguyenTe.MaSo As F2, NguyenTe.KyHieu As F1 FROM NguyenTe INNER JOIN" _
                                      & " HethongTK ON NguyenTe.MaSo = HethongTK.MaNT WHERE HethongTK.SoHieu = '" _
                                      + taikhoan.sohieu + "' ORDER BY NguyenTe.KyHieu", CboNT(0)
                        CboNT(0).AddItem pTienStr, 0
                        CboNT(0).ItemData(0) = 0
                        CboNT(0).ListIndex = 0
                        CboNT(0).Visible = True
                        txtchungtu(4).Enabled = True
                        RFocus CboNT
                    End If
                    txtchungtu(3).Text = "0"
                Else
                    CboNT(0).Visible = False
                    If ((taikhoan.tk_id = TKVT_ID Or taikhoan.tk_id = TKDT_ID Or taikhoan.tk_id = TKGT_ID) And (loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8)) Then
                        If pNhapDoiTuong > 0 Then
                            txtchungtu(2).Enabled = False
                            CboNT(3).Visible = True
                            LaySohieuDoiTuong2 1, ""
                        Else
                            txtchungtu(2).Enabled = True
                        End If
                        txtchungtu(3).Enabled = True
                    Else
                        txtchungtu(2).Enabled = False
                        txtchungtu(3).Enabled = False
                    End If
                    If loaict = 1 Or loaict = 8 Then
                        txtchungtu(4).Enabled = True
                    Else
                        CboNT(1).Visible = False
                    End If
                    RFocus txtchungtu(2)
                End If
            Else
                txtchungtu(2).Enabled = False
                txtchungtu(3).Enabled = False
            End If
        End If

        If (pDTTP <> 0 And (loaict <> 3 And (Left(taikhoan.sohieu, 2) = "64" Or Left(taikhoan.sohieu, 3) = "621" Or Left(taikhoan.sohieu, 3) = "622" Or Left(taikhoan.sohieu, 3) = "623" Or Left(taikhoan.sohieu, 3) = "627" Or Left(taikhoan.sohieu, 3) = "154" Or Left(taikhoan.sohieu, 3) = "911"))) Or (loaict = 8 And Left(taikhoan.sohieu, 3) = "521") Then
            txtchungtu(2).Enabled = True
            RFocus txtchungtu(2)
        End If

    Case 2:    ' So hieu vat tu, nguyen te
        Label(12).Visible = False
        CboNT(2).Visible = False
        CboNT(1).Visible = False
        txtchungtu(4).Enabled = False

        If pDTTP <> 0 And (loaict <> 3 And ((Left(taikhoan.sohieu, 2) = "51" And taikhoan.tk_id2 = TKDT_ID) Or Left(taikhoan.sohieu, 2) = "64" Or Left(taikhoan.sohieu, 3) = "621" Or Left(taikhoan.sohieu, 3) = "622" Or Left(taikhoan.sohieu, 3) = "623" Or Left(taikhoan.sohieu, 3) = "627" Or Left(taikhoan.sohieu, 3) = "154" Or Left(taikhoan.sohieu, 3) = "911")) And Len(txtchungtu(2).Text) > 0 Then
            tp.InitTPSohieu txtchungtu(2).Text
            txtchungtu(1).Text = tp.TenVattu
            If Left(taikhoan.sohieu, 3) = "154" And tp.MaSo > 0 Then
                tien = tp.GiaThanhCK(CboThang.ItemData(CboThang.ListIndex))
                If tien <> 0 Then
                    '   txtchungtu(5).Text = ""
                    txtchungtu(6).Text = Format(tien, Mask_2)
                End If
            End If
        Else
            tp.InitTPMaSo 0
        End If

        If ((taikhoan.tk_id = TKCNKH_ID) Or (taikhoan.tk_id = TKCNPT_ID)) Then
            ckh.InitKhachHangSohieu txtchungtu(2).Text
            txtchungtu(1).Text = ckh.Ten
            txtchungtu(3).Enabled = (ckh.MaNT > 0)
            txtchungtu(2).tag = IIf(ckh.MaNT > 0, 1, 0)
            If ckh.MaNT = 0 Then
                txtchungtu(3).Text = "0"
            Else
                txtchungtu(4).Enabled = True
                txtchungtu(4).Text = Format(TyGiaNT(ckh.MaNT), Mask_0)
                RFocus txtchungtu(3)
            End If
        End If

        If (((loaict = 1 Or loaict = 2 Or loaict = 8) And (taikhoan.tk_id = TKVT_ID)) Or ((loaict = 7 Or loaict = 8) And (taikhoan.tk_id = TKDT_ID Or taikhoan.tk_id = TKGT_ID))) And STDetail And VTEnable And tp.MaSo = 0 Then
            vattu.InitVattuSohieu txtchungtu(2).Text
            If vattu.MaSo > 0 Then
                VTEnable = False
                Label(25).Visible = (pGiaUSD > 0) And (loaict = 1 Or loaict = 2 Or loaict = 8)
                txtchungtu(11).Visible = (pGiaUSD > 0) And (loaict = 1 Or loaict = 2 Or loaict = 8)
                If (taikhoan.loai = 0 Or (OutCost > 0 And OutCost <> 2)) And loaict = 2 Then
                    FDsNhap.tag = vattu.MaSo
                    MaNhap = FDsNhap.XuatDichDanh(CboThang.ItemData(CboThang.ListIndex), vattu.sohieu + " - " + vattu.TenVattu + ABCtoVNI(" - §.v.t: ") + vattu.DonVi, CboNguon(1).ItemData(CboNguon(1).ListIndex), luong, tien)
                    If luong = 0 Then
                        luong = SoTonKho(CboThang.ItemData(CboThang.ListIndex), CboNguon(1).ItemData(CboNguon(1).ListIndex), taikhoan.MaSo, vattu.MaSo, tien, tien2)
                    End If
                    txtchungtu(3).tag = luong
                    txtchungtu(5).tag = tien2
                    txtchungtu(6).tag = tien
                    txtchungtu(3).Text = Format(luong, Mask_2)
                    txtchungtu(6).Text = Format(tien, Mask_2)
                    If pGiaUSD > 0 Then txtchungtu(11).Text = Format(tien2, Mask_2)
                    If luong <> 0 Then
                        txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien / luong) / Mask_N, Mask_2)
                    End If
                    HienThongBao "Sè l­îng tån kho: " + txtchungtu(3).Text + " - Thµnh tiÒn: " + txtchungtu(6).Text, 1
                Else
                    If loaict <> 8 Then
                        luong = SoTonKho(CboThang.ItemData(CboThang.ListIndex), CboNguon(1).ItemData(CboNguon(1).ListIndex), taikhoan.MaSo, vattu.MaSo, tien, tien2)
                    Else
                        luong = SoTonKho(CboThang.ItemData(CboThang.ListIndex), CboNguon(1).ItemData(CboNguon(1).ListIndex), 0, vattu.MaSo, tien, tien2)
                    End If
                    txtchungtu(3).Text = Format(luong, Mask_2)
                    If loaict = 1 Then
                        txtchungtu(5).Text = Format(tien, Mask_2)
                    Else
                        txtchungtu(3).tag = luong
                        txtchungtu(6).Text = Format(tien, Mask_2)
                    End If
                    If pGiaUSD > 0 Then txtchungtu(11).Text = Format(tien2, Mask_2)
                    txtchungtu(6).tag = tien
                    txtchungtu(5).tag = tien2
                    If luong <> 0 Then
                        If pGiaUSD > 0 Then
                            txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien2 / luong) / Mask_N, Mask_2)
                        Else
                            txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien / luong) / Mask_N, Mask_2)
                        End If
                    End If
                End If
                If loaict = 8 And (vattu.GiaBan1 > 0 Or vattu.GiaBan2 > 0 Or vattu.GiaBan3 > 0 Or CDbl(txttinh_gia_ban.Caption) > 0) Then
                    LayGiaBan
                Else
                    If luong <> 0 Then
                        If pGiaUSD > 0 Then
                            txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien2 / luong) / Mask_N, Mask_2)
                        Else
                            txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien / (luong * IIf(pGiaUSD > 0, pRate, 1))) / Mask_N, Mask_2)
                        End If
                    End If
                End If
                If loaict = 1 And pGiaHT > 0 And vattu.GiaHT > 0 And Left(taikhoan.sohieu, Len(ShTkTP)) = ShTkTP Then
                    txtchungtu(4).Text = Format(vattu.GiaHT, Mask_2)
                End If
                HienThongBao "Sè l­îng tån kho: " + txtchungtu(3).Text + " - Thµnh tiÒn: " + Format(txtchungtu(6).tag, Mask_0), 1

                txtchungtu(1).Text = vattu.TenVattu
                txtchungtu(2).tag = 1
                RFocus txtchungtu(3)
                VTEnable = True
            End If
        End If
        With GrdChungtu
            tien = 0
            luong = 0
            If (taikhoan.tk_id = GTGTKT_ID) Then
                On Error Resume Next
                j = Abs(CLng5(txtchungtu(2).Text))
                On Error GoTo 0
                If j = 0 Then
                    txtchungtu(5).Text = "0"
                    txtchungtu(6).Text = "0"
                    Exit Sub
                End If
                If Cdbl5(txtchungtu(5).Text) = 0 Then
                    If j > 0 And j < 5 And Right(txtchungtu(2).Text, 1) <> "-" Then txtchungtu(2).Text = txtchungtu(2).Text + "-"
                    For i = 0 To .Rows - 1
                        .Row = i
                        .col = 1
                        sh = .Text
                        If Len(sh) = 0 Then Exit For
                        If Left(sh, Len(pVATV)) = pVATV Then Exit For
                        .col = 6
                        If Right(txtchungtu(2).Text, 1) = "-" And Left(sh, 3) <> "211" Then
                            tien = Cdbl5(.Text)
                            If j > 0 And j < 5 Then
                                v = tien * j / 100
                            Else
                                v = tien * j / (100 + j)
                            End If
                            v = RoundMoney(v)
                            .Text = Format(tien - v, Mask_0)
                            luong = luong + v
                        Else
                            tien = tien + Cdbl5(.Text)
                        End If
                        If Left(sh, 3) = "338" Then
                            .col = 7
                            tien = tien - Cdbl5(.Text)
                        End If
                    Next
                    If Right(txtchungtu(2).Text, 1) <> "-" Then
                        luong = RoundMoney(tien * j / 100)
                    Else
                        txtchungtu(2).Text = CStr(j)
                    End If
                    txtchungtu(5).Text = Format(luong, Mask_0)
                    txtchungtu(6).Text = "0"
                End If
            End If
            If (taikhoan.tk_id = GTGTPN_ID Or taikhoan.tk_id = TTDB_ID) Then
                If Cdbl5(txtchungtu(2).Text) = 0 Then
                    txtchungtu(5).Text = "0"
                    txtchungtu(6).Text = "0"
                    Exit Sub
                End If
                On Error Resume Next
                j = Abs(CLng5(txtchungtu(2).Text))
                On Error GoTo 0
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 1
                    sh = .Text
                    If Len(sh) = 0 Then Exit For
                    If Left(sh, 4) = "3331" And taikhoan.tk_id = GTGTPN_ID Then Exit For
                    If Left(sh, 2) <> "11" And Left(sh, 4) <> "3331" And Left(sh, 3) <> "521" Then
                        .col = 7
                        If Right(txtchungtu(2).Text, 1) <> "-" And pVAT1 > 0 And taikhoan.tk_id = GTGTPN_ID Then txtchungtu(2).Text = txtchungtu(2).Text + "-"
                        If Right(txtchungtu(2).Text, 1) = "-" And taikhoan.tk_id = GTGTPN_ID Then
                            tien = Cdbl5(.Text)
                            If j > 0 And j < 5 Then
                                v = tien * j / 100
                            Else
                                v = tien * j / (100 + j)
                            End If
                            v = RoundMoney(v)
                            .Text = Format(tien - v, Mask_0)
                            luong = luong + v
                        Else
                            If loaict = 1 Then .col = 6
                            tien = tien + Cdbl5(.Text)
                        End If
                    End If

                    If Left(sh, Len(pSHPT)) <> pSHPT And Left(sh, 2) <> "11" Then            'Left(sh, 3) <> "521" And
                        .col = 6
                        tien = tien - Cdbl5(.Text)
                    End If
                Next
                If taikhoan.tk_id = GTGTPN_ID Then
                    If Right(txtchungtu(2).Text, 1) <> "-" Then
                        luong = RoundMoney(tien * j / 100)
                    Else
                        txtchungtu(2).Text = CStr(j)
                    End If
                Else
                    luong = RoundMoney(tien * j / (j + 100))
                End If
                If luong >= 0 Then
                    txtchungtu(6).Text = Format(luong, Mask_2)
                    txtchungtu(5).Text = "0"
                Else
                    txtchungtu(5).Text = Format(-luong, Mask_2)
                    txtchungtu(6).Text = "0"
                End If
                If luong <> 0 Then
                    '    RFocus txtchungtu(6)
                End If
            End If
        End With

        If loaict = 8 And Left(taikhoan.sohieu, 3) = "521" Then
            luong = Cdbl5(txtchungtu(2).Text)
            If luong > 0 And luong < 100 Then
                txtchungtu(5).Text = Format(RoundMoney(GiaTriTruocThue * luong / 100), Mask_0)
            Else
                txtchungtu(5).Text = Format(SoChietKhau, Mask_0)
            End If
        End If

        If (loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8) And vattu.MaSo > 0 And vattu.Dvt2 > 0 Then
            Label(12).Visible = True
            CboNT(2).Visible = True
            Int_RecsetToCbo "SELECT MaSo AS F2, DonVi AS F1 FROM DVTVattu WHERE MaVattu=" + CStr(vattu.MaSo) + " ORDER BY DonVi", CboNT(2)
            CboNT(2).AddItem vattu.DonVi, 0
            CboNT(2).ListIndex = 0
            RFocus CboNT(2)
        End If

        txtchungtu(9).Text = ""
        txtchungtu(10).Text = ""
        txtchungtu(4).Enabled = ((loaict = 1 Or loaict = 7 Or loaict = 8) And (vattu.MaSo > 0)) Or (ckh.MaNT > 0)
        If loaict = 8 And vattu.MaSo > 0 And vattu.CK > 0 And pChietKhau > 0 Then
            txtchungtu(9).Text = Format(vattu.CK, IIf(vattu.CK * 100 Mod 100 <> 0, Mask_2, Mask_0))
            TinhCKCT
        End If
    Case 3:    ' So luong
        txtchungtu(3).Text = Format(txtchungtu(3).Text, Mask_2)
        If (loaict = 2 Or loaict = 7 Or loaict = 8) And vattu.MaSo > 0 Then
            luong = Cdbl5(txtchungtu(3).Text)
            txtchungtu(5).Text = "0"
            If loaict = 2 Then
                If luong <> txtchungtu(3).tag And txtchungtu(3).tag <> 0 Then
                    txtchungtu(6).Text = Format((luong * txtchungtu(6).tag) / txtchungtu(3).tag, Mask_2)
                    If pGiaUSD > 0 Then txtchungtu(11).Text = Format((luong * txtchungtu(5).tag) / txtchungtu(3).tag, Mask_2)
                Else
                    txtchungtu(6).Text = Format(txtchungtu(6).tag, Mask_2)
                    If pGiaUSD > 0 Then txtchungtu(11).Text = Format(txtchungtu(5).tag, Mask_2)
                End If
            Else
                If pGiaUSD > 0 Then txtchungtu(11).Text = Format(luong * Cdbl5(txtchungtu(4).Text), Mask_2)
                txtchungtu(6).Text = Format(luong * Cdbl5(txtchungtu(4).Text) * IIf(pGiaUSD > 0, pRate, 1), Mask_2)
            End If
        End If
        If (taikhoan.MaNT <> 0 And CboNT(0).ListIndex > 0) Or ckh.MaNT > 0 Then
            luong = Cdbl5(txtchungtu(3).Text)
            If Cdbl5(txtchungtu(4).Text) <> 0 Then
                If SoPSConLai < 0 Then
                    txtchungtu(5).Text = Format(luong * Cdbl5(txtchungtu(4).Text), Mask_2)
                Else
                    txtchungtu(6).Text = Format(luong * Cdbl5(txtchungtu(4).Text), Mask_2)
                End If
            End If
        End If
        If loaict = 8 And vattu.MaSo > 0 Then TinhCKCT
    Case 4, 5, 6:    ' PS No, Co
        Dim psnt As Boolean, tygia As Double, m As String, nt As Double

        m = IIf(Left(taikhoan.sohieu, 3) = "007", Mask_2, Mask_0)
        If Index = 4 Then
            txtchungtu(Index).Text = Format(txtchungtu(Index).Text, Mask_2)
        Else
            txtchungtu(Index).Text = Format(Cdbl5(txtchungtu(Index).Text), m)
        End If

        If Index = 5 Or Index = 6 Then
            If Cdbl5(txtchungtu(Index).Text) <> 0 Then txtchungtu(11 - Index).Text = "0"
        End If

        psnt = (CboNT(0).Visible And CboNT(0).ListIndex > 0) Or ckh.MaNT > 0
        If ((loaict = 1 Or loaict = 7 Or loaict = 8) And (vattu.MaSo > 0)) Or psnt Then
            luong = Cdbl5(txtchungtu(3).Text)
            Select Case Index
            Case 5:
                If loaict = 1 Or psnt Then
                    If luong > 0 And Cdbl5(txtchungtu(5).Text) <> 0 Then
                        If loaict = 1 Or pTien = 0 Then
                            tygia = Cdbl5(txtchungtu(5).Text) / luong
                        Else
                            tygia = luong / Cdbl5(txtchungtu(5).Text)
                        End If
                        If pGiaUSD > 0 And pRate > 0 Then
                            txtchungtu(4).Text = Format(tygia / pRate, Mask_2)
                        Else
                            txtchungtu(4).Text = Format(tygia, Mask_2)
                        End If
                        If psnt Then
                            If ckh.MaNT > 0 Then
                                CapNhatTyGia ckh.MaNT, tygia
                            Else
                                CapNhatTyGia CboNT(0).ItemData(CboNT(0).ListIndex), tygia
                            End If
                        End If
                    End If
                End If
            Case 6:
                If loaict = 7 Or loaict = 8 Or psnt Then
                    If luong > 0 And Cdbl5(txtchungtu(6).Text) <> 0 Then
                        If loaict = 7 Or loaict = 8 Or pTien = 0 Then
                            tygia = Cdbl5(txtchungtu(6).Text) / luong
                            If loaict = 8 And pGiaUSD > 0 Then tygia = tygia / pRate
                        Else
                            If Cdbl5(txtchungtu(5).Text) <> 0 Then tygia = luong / Cdbl5(txtchungtu(5).Text)
                        End If
                        txtchungtu(4).Text = Format(tygia, Mask_2)
                        If psnt Then
                            If ckh.MaNT > 0 Then
                                CapNhatTyGia ckh.MaNT, tygia
                            Else
                                CapNhatTyGia CboNT(0).ItemData(CboNT(0).ListIndex), tygia
                            End If
                        End If
                    End If
                End If
            Case 4:
                tygia = Cdbl5(txtchungtu(4).Text)
                If psnt Then
                    tien = SoPSConLai
                    nt = DoiRaTien(luong, tygia)
                    If tien < 0 Or (tien = 0 And taikhoan.loai < 0) Then
                        txtchungtu(5).Text = Format(nt, Mask_2)
                    Else
                        txtchungtu(6).Text = Format(nt, Mask_2)
                    End If
                    If ckh.MaNT > 0 Then
                        CapNhatTyGia ckh.MaNT, tygia
                    Else
                        CapNhatTyGia CboNT(0).ItemData(CboNT(0).ListIndex), tygia
                    End If
                Else
                    If loaict = 1 Or taikhoan.tk_id = TKGT_ID Then
                        If luong > 0 Then
                            If luong = txtchungtu(3).tag Then
                                txtchungtu(4).Text = Format(txtchungtu(6).tag, Mask_0)
                            Else
                                txtchungtu(5).Text = Format(tygia * luong * IIf(pGiaUSD > 0, pRate, 1), Mask_0)
                                If pGiaUSD > 0 Then txtchungtu(11).Text = Format(tygia * luong, Mask_2)
                            End If
                        Else
                            txtchungtu(5).Text = txtchungtu(4).Text
                        End If
                    Else
                        If luong > 0 Then
                            If pGiaUSD > 0 Then txtchungtu(11).Text = Format(tygia * luong, Mask_2)
                            txtchungtu(6).Text = Format(tygia * luong * IIf(pGiaUSD > 0, pRate, 1), Mask_2)
                            If Cdbl5(txtchungtu(6).Text) <> 0 Then txtchungtu(5).Text = "0"
                        Else
                            If pGiaUSD > 0 Then
                                txtchungtu(6).Text = Format(Cdbl5(txtchungtu(4).Text) * pRate, Mask_2)
                            Else
                                txtchungtu(6).Text = txtchungtu(4).Text
                            End If
                        End If
                    End If
                End If
            End Select
        End If
        If (loaict = 7 Or loaict = 8) And vattu.MaSo > 0 Then TinhCKCT
    Case 7:
        If pTygia > 0 Then pRate = Cdbl5(txtchungtu(7).Text)
        txtchungtu(Index).Text = Format(txtchungtu(Index).Text, Mask_2)
    Case 9:
        txtchungtu(Index).Text = Format(txtchungtu(Index).Text, Mask_2)
        If loaict = 8 And vattu.MaSo > 0 Then TinhCKCT
    Case 10:
        txtchungtu(Index).Text = Format(txtchungtu(Index).Text, Mask_0)
    End Select
End Sub
Private Function Kiemtrataikhoanchitiet(taikhoan As String) As Boolean
    Dim sql
    Dim rs_chungtu
    sql = " select * from hethongtk where SoHieu = '" + Trim(taikhoan) + "' and tkcon = 0 "
    Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs_chungtu.recordCount = 0 Then
        Kiemtrataikhoanchitiet = False
    Else
        Kiemtrataikhoanchitiet = True
    End If

    ' Kiemtrataikhoanchitiet = True
End Function


'====================================================================================================
' Xö lý phÝm bÊm trªn c¸c « nhËp
'====================================================================================================
Private Sub txtChungtu_KeyPress(Index As Integer, KeyAscii As Integer)


    Dim str As String

    Select Case Index
    Case 0:
        ' If KeyAscii = vbKeyReturn Then 'phim enter
        If KeyAscii = 63 Then
            Me.MousePointer = 11
            txtchungtu(0).Text = FrmTaikhoan.ChonTk(txtchungtu(0).Text)
            If Kiemtrataikhoanchitiet(txtchungtu(0).Text) = False Then Exit Sub
            Me.MousePointer = 0
            txtChungtu_LostFocus 0
            ' ham xu ly chung tu
            If txtchungtu(3).Enabled = True Then
                RFocus txtchungtu(3)
            Else
                RFocus txtchungtu(5)
            End If
        Else
        End If
        '//////////////////////////////////

        '''''''''''''''''''''''''
        If KeyAscii = 13 Then
            If Len(Replace(Trim(txtchungtu(0).Text), ".", "")) <= 0 Or Kiemtrataikhoanchitiet(txtchungtu(0).Text) = False Then
                txtchungtu(0).Text = FrmTaikhoan.ChonTk(txtchungtu(0).Text)
            Else
                RFocus txtchungtu(5)
            End If
        End If
    Case 1:
        KeyAscii = 0
        Beep
    Case 2:

        If (((taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) And KHDetail) Or (((taikhoan.tk_id = TKVT_ID And (loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8 Or loaict = 9)) Or ((taikhoan.tk_id = TKDT_ID Or taikhoan.tk_id = TKGT_ID) And (loaict = 7 Or loaict = 8) And taikhoan.tk_id2 <> TKDT_ID)) And STDetail)) And (KeyAscii = vbKeyReturn) Then
            Me.MousePointer = 11
            VTEnable = False
            If (taikhoan.tk_id <> TKCNKH_ID And taikhoan.tk_id <> TKCNPT_ID) Then
                txtchungtu(2).Text = FrmVattu.ChonVattu(txtchungtu(2).Text)
                txtchungtu(3).Enabled = True
                txtchungtu(4).Enabled = True
                RFocus txtchungtu(3)
            Else
                '  If KeyAscii = 63 Then
                txtchungtu(2).Text = FrmKhachHang.ChonKhachHang(txtchungtu(2).Text)
                RFocus txtchungtu(5)
                ' End If
                ' If KeyAscii = 13 Then RFocus txtchungtu(5)

            End If
            Me.MousePointer = 0

            txtChungtu_LostFocus 2
            VTEnable = True
        Else
            '  If pDTTP <> 0 And (loaict <> 3 And (Left(taikhoan.sohieu, 2) = "64" Or Left(taikhoan.sohieu, 3) = "621" Or Left(taikhoan.sohieu, 3) = "622" Or Left(taikhoan.sohieu, 3) = "623" Or Left(taikhoan.sohieu, 3) = "627" Or Left(taikhoan.sohieu, 3) = "154" Or Left(taikhoan.sohieu, 3) = "911" Or taikhoan.tk_id2 = TKDT_ID)) Then
            If pDTTP <> 0 And ((Left(taikhoan.sohieu, 2) = "64" Or Left(taikhoan.sohieu, 3) = "621" Or Left(taikhoan.sohieu, 3) = "622" Or Left(taikhoan.sohieu, 3) = "623" Or Left(taikhoan.sohieu, 3) = "627" Or Left(taikhoan.sohieu, 3) = "154" Or Left(taikhoan.sohieu, 3) = "911" Or taikhoan.tk_id2 = TKDT_ID)) Then
                If (KeyAscii = vbKeyReturn) Then
                    txtchungtu(2).Text = FrmTP.ChonTP(txtchungtu(2).Text)
                    If txtchungtu(3).Enabled = True Then
                        RFocus txtchungtu(3)
                    Else
                        RFocus txtchungtu(5)
                    End If
                End If
            Else
                ' If (taikhoan.tk_id = GTGTKT_ID Or taikhoan.tk_id = GTGTPN_ID Or taikhoan.tk_id = TTDB_ID) Then KeyProcess txtchungtu(Index), KeyAscii, True
                If KeyAscii = 13 Then
                    If txtchungtu(3).Enabled = True Then
                        RFocus txtchungtu(3)
                    Else
                        RFocus txtchungtu(5)
                    End If
                End If

                If ((loaict <> 1 And loaict <> 2 And loaict <> 7 And loaict <> 8 And loaict <> 9) Or (Not STDetail)) And Not ((taikhoan.tk_id = GTGTKT_ID) Or (taikhoan.tk_id = GTGTPN_ID) Or (taikhoan.tk_id = TTDB_ID)) And ((taikhoan.tk_id <> TKCNKH_ID And taikhoan.tk_id <> TKCNPT_ID) Or Not KHDetail) Then KeyAscii = 0
            End If
        End If

    Case 3:

        If KeyAscii = 13 Then
            If Index = 3 Then
                txtchungtu(4).Enabled = True
                If txtchungtu(4).Enabled = True Then
                    RFocus txtchungtu(4)
                Else
                    RFocus txtchungtu(5)
                End If
            End If
        Else
            If txtchungtu(2).tag = 1 Then KeyProcess txtchungtu(Index), KeyAscii, True Else KeyAscii = 0
        End If
    Case 4:

        If KeyAscii = 13 Then
            If txtchungtu(5).Enabled = True Then
                RFocus txtchungtu(5)
            Else
                RFocus txtchungtu(6)
            End If
        Else
            If txtchungtu(2).tag = 1 Then KeyProcess txtchungtu(Index), KeyAscii, True Else KeyAscii = 0
        End If
    Case 5:
        If KeyAscii = vbKeyReturn Then
            'CmdChitiet_chon
            RFocus txtchungtu(6)
        Else: KeyProcess txtchungtu(Index), KeyAscii, True
        End If
        'luu xuong nut ghi
        If txtchungtu(0).Text <> "007" Then
            Di_chuyen_con_tro_xuong_nut_ghi
        End If

    Case 6:

        Dim so

        If IsNumeric(txtchungtu(6).Text) Then
            If Int(txtchungtu(6).Text) > 0 Then
                If KeyAscii = 13 Then
                    If (IsNumeric(txtchungtu(3).Text)) Then
                        If Int(txtchungtu(3).Text) > 0 Then    ' tinh lai gia
                            txtChungtu_LostFocus (6)
                            '   txtchungtu(4).Text = Format(CDbl(txtchungtu(6).Text) / CDbl(txtchungtu(3).Text), Mask_2)
                        End If
                    End If
                End If

            End If
        End If

        'so = txtchungtu(6).Text
        If KeyAscii <> 13 And KeyAscii <> 63 Then    ' 43
            hien_bang_tinh = False
            '   tongtientruoc = txtchungtu(6).Text
            KeyProcess txtchungtu(Index), KeyAscii
            'FrmCal.Show 1
            'RFocus CmdChitiet
        Else
            If KeyAscii = vbKeyReturn Then
                dathuchien = True    ' thuc hien luu de focus
                cho_hien_vat = False
                ' If Len(Trim(txtchungtu(1).Text)) > 0 Then CmdChitiet_chon 'CmdChitiet_Click  'sua theo tieu chuan nha thuoc
                If Len(Trim(txtchungtu(1).Text)) > 0 Then
                    If SelectSQL("SELECT banthuoc as f1 from license ") = 1 Then
                        CmdChitiet_Click
                    Else
                        CmdChitiet_chon    'CmdChitiet_Click  'sua theo tieu chuan nha thuoc
                    End If
                End If
            ElseIf KeyAscii = 63 Then
                dathuchien = True    ' thuc hien luu de focus
                cho_hien_vat = True
                'If Len(Trim(txtchungtu(1).Text)) > 0 Then CmdChitiet_chon ' CmdChitiet_Click  sua theo tieu chuan nha thuoc
                If Len(Trim(txtchungtu(1).Text)) > 0 Then
                    If SelectSQL("SELECT banthuoc as f1 from license ") = 1 Then
                        CmdChitiet_Click
                    Else
                        CmdChitiet_chon  'CmdChitiet_Click 'sua theo tieu chuan nha thuoc
                    End If
                End If
            Else
                If loaict = 2 And vattu.MaSo > 0 And FCost Then
                    If Cdbl5(txtchungtu(3).Text) <> 0 Then
                        KeyAscii = 0
                    Else
                        '  KeyProcess txtchungtu(Index), KeyAscii 'ghi chu
                    End If
                Else
                    'KeyProcess txtchungtu(Index), KeyAscii, True
                End If

            End If
            'luu xuong nut ghi
            '  Dim SO, SO1
            If txtchungtu(0).Text <> "007" Then
                Di_chuyen_con_tro_xuong_nut_ghi
            End If
        End If
    Case 7, 8:
        KeyProcess txtchungtu(Index), KeyAscii
    Case 9, 10, 11:
        If KeyAscii = 63 Then
            cho_hien_vat = True
            If Len(Trim(txtchungtu(1).Text)) > 0 Then txtChungtu_LostFocus (6)    'CmdChitiet_chon
        Else
            If KeyAscii = vbKeyReturn Then
                cho_hien_vat = False
                If Len(Trim(txtchungtu(1).Text)) > 0 Then txtChungtu_LostFocus (6)    ' CmdChitiet_chon
            Else
                KeyProcess txtchungtu(Index), KeyAscii
            End If
        End If
        If KeyAscii = 13 Then CmdChitiet_chon
    End Select

End Sub
' dung di chuyen con tro xuong nut ghi
Sub Di_chuyen_con_tro_xuong_nut_ghi()
    Dim tongbenco
    Dim tongbenno
    tongbenco = 0
    tongbenno = 0
    Dim hh
    For hh = 0 To GrdChungtu.Rows - 1
        GrdChungtu.Row = hh
        GrdChungtu.col = 7
        If Len(GrdChungtu.Text) > 0 Then
            tongbenco = tongbenco + Int(GrdChungtu.Text)
        End If
        GrdChungtu.col = 6
        If Len(GrdChungtu.Text) > 0 Then
            tongbenno = tongbenno + Int(GrdChungtu.Text)
        End If
    Next
    If tongbenco <> 0 Or tongbenno <> 0 Then
        If tongbenco - tongbenno = 0 Then
            RFocus Command(1)
        End If
    End If
End Sub

'====================================================================================================
' KiÓm tra d÷ liÖu nhËp t¹i c¸c « nhËp
'====================================================================================================
Public Sub txtChungtu_LostFocus(Index As Integer)
    Label(26).Caption = ""
    Dim luong As Double, tien As Double, i As Integer, j As Integer, v As Double, sh As String, tien2 As Double
    If Len(Trim(txtchungtu(2).Text)) > 0 Then txtchungtu(2).Enabled = True
    Select Case Index
    Case 0:    ' So hieu tai khoan

        taikhoan.InitTaikhoanSohieu txtchungtu(0).Text
        txtchungtu(1).Text = IIf(pNN = 0, taikhoan.Ten, taikhoan.TenE)
        If (taikhoan.tk_id = TTDB_ID) Or (taikhoan.tk_id = GTGTKT_ID) Or (taikhoan.tk_id = GTGTPN_ID) Or (taikhoan.tk_id = TKCNKH_ID) Or (taikhoan.tk_id = TKCNPT_ID) Then
            If ((taikhoan.tk_id = TKCNKH_ID) Or (taikhoan.tk_id = TKCNPT_ID)) And hdcount >= 0 Then
                ckh.InitKhachHangMaSo HD(hdcount).MaKhachHang
                txtchungtu(2).Text = ckh.sohieu
                If txtchungtu(3).Enabled = True Then
                End If
            End If
            CboNT(0).Visible = False
            txtchungtu(4).Enabled = False
            txtchungtu(3).Enabled = False
            CboNT(3).Visible = False

            If ((taikhoan.tk_id = TKCNKH_ID) Or (taikhoan.tk_id = TKCNPT_ID)) And (Not KHDetail) Then
                txtchungtu(2).Enabled = False
            Else
                If pNhapDoiTuong > 0 And ((taikhoan.tk_id = TKCNKH_ID) Or (taikhoan.tk_id = TKCNPT_ID)) Then
                    txtchungtu(2).Enabled = False
                    CboNT(3).Visible = True
                    LaySohieuDoiTuong2 2, ""
                    If hdcount >= 0 Then
                        CboNT(3).Text = ckh.sohieu + " - " + ckh.Ten
                    End If
                Else
                    txtchungtu(2).Enabled = True
                    If Len(Trim(txtchungtu(2).Text)) <= 0 Then RFocus txtchungtu(2)
                    '=====================================

                End If
            End If
        Else
            '////////////////////
            txtchungtu(4).Enabled = False
            'txtchungtu(2).Enabled = False
            CboNT(3).Visible = False
            If taikhoan.MaSo > 0 Then
                If (taikhoan.tk_id <> TKVT_ID And taikhoan.tk_id <> TKDT_ID And taikhoan.tk_id <> TKGT_ID) Or (Not STDetail) Or (taikhoan.tk_id = TKDT_ID And loaict <> 8) Or (taikhoan.tk_id = TKVT_ID And loaict <> 1 And loaict <> 2 And loaict <> 8) Then
                    txtchungtu(2).tag = IIf(taikhoan.MaNT > 0, 1, 0)
                    txtchungtu(2).Text = ""
                    If taikhoan.MaNT = 0 Then
                        CboNT(0).Visible = False
                        txtchungtu(3).Text = "0"
                        txtchungtu(3).Enabled = False
                        CboNT(1).Visible = False
                    Else
                        Int_RecsetToCbo "SELECT NguyenTe.MaSo As F2, NguyenTe.KyHieu As F1 FROM NguyenTe INNER JOIN" _
                                      & " HethongTK ON NguyenTe.MaSo = HethongTK.MaNT WHERE HethongTK.SoHieu = '" _
                                      + taikhoan.sohieu + "' ORDER BY NguyenTe.KyHieu", CboNT(0)
                        CboNT(0).AddItem pTienStr, 0
                        CboNT(0).ItemData(0) = 0
                        CboNT(0).ListIndex = 0
                        CboNT(0).Visible = True
                        txtchungtu(4).Enabled = True
                        RFocus CboNT
                    End If
                    txtchungtu(3).Text = "0"
                Else
                    CboNT(0).Visible = False
                    If ((taikhoan.tk_id = TKVT_ID Or taikhoan.tk_id = TKDT_ID Or taikhoan.tk_id = TKGT_ID) And (loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8)) Then
                        If pNhapDoiTuong > 0 Then
                            txtchungtu(2).Enabled = False
                            CboNT(3).Visible = True
                            LaySohieuDoiTuong2 1, ""
                        Else
                            txtchungtu(2).Enabled = True
                        End If
                        txtchungtu(3).Enabled = True
                    Else
                        txtchungtu(2).Enabled = False
                        txtchungtu(3).Enabled = False
                    End If
                    If loaict = 1 Or loaict = 8 Then
                        txtchungtu(4).Enabled = True
                    Else
                        CboNT(1).Visible = False
                    End If
                    'bo/////////////////////////////////////
                    'If Len(Trim(txtchungtu(2).Text)) <= 0 Then RFocus txtchungtu(2)
                    If txtchungtu(2).Enabled = True Then
                        If IsImport = False Then RFocus txtchungtu(2)
                    End If
                End If
            Else
                txtchungtu(2).Enabled = False
                txtchungtu(3).Enabled = False
            End If
        End If

        '   If (pDTTP <> 0 And (loaict <> 3 And (Left(taikhoan.sohieu, 2) = "64" Or Left(taikhoan.sohieu, 3) = "621" Or Left(taikhoan.sohieu, 3) = "622" Or Left(taikhoan.sohieu, 3) = "623" Or Left(taikhoan.sohieu, 3) = "627" Or Left(taikhoan.sohieu, 3) = "154" Or Left(taikhoan.sohieu, 3) = "911"))) Or (loaict = 8 And Left(taikhoan.sohieu, 3) = "521") Or (loaict = 9 And Left(taikhoan.sohieu, 3) = "153") Then
        If (pDTTP <> 0 And ((Left(taikhoan.sohieu, 2) = "64" Or Left(taikhoan.sohieu, 3) = "621" Or Left(taikhoan.sohieu, 3) = "622" Or Left(taikhoan.sohieu, 3) = "623" Or Left(taikhoan.sohieu, 3) = "627" Or Left(taikhoan.sohieu, 3) = "154" Or Left(taikhoan.sohieu, 3) = "911"))) Or (loaict = 8 And Left(taikhoan.sohieu, 3) = "521") Or (loaict = 9 And Left(taikhoan.sohieu, 3) = "153") Then
            txtchungtu(2).Enabled = True
            RFocus txtchungtu(2)
        End If
    Case 2:    ' So hieu vat tu, nguyen te
        Label(12).Visible = False
        CboNT(2).Visible = False
        CboNT(1).Visible = False
        txtchungtu(4).Enabled = False

        '  If pDTTP <> 0 And (loaict <> 3 And ((Left(taikhoan.sohieu, 2) = "51" And taikhoan.tk_id2 = TKDT_ID) Or Left(taikhoan.sohieu, 2) = "64" Or Left(taikhoan.sohieu, 3) = "621" Or Left(taikhoan.sohieu, 3) = "622" Or Left(taikhoan.sohieu, 3) = "623" Or Left(taikhoan.sohieu, 3) = "627" Or Left(taikhoan.sohieu, 3) = "154" Or Left(taikhoan.sohieu, 3) = "911")) And Len(txtchungtu(2).Text) > 0 Then
        If pDTTP <> 0 And (((Left(taikhoan.sohieu, 2) = "51" And taikhoan.tk_id2 = TKDT_ID) Or Left(taikhoan.sohieu, 2) = "64" Or Left(taikhoan.sohieu, 3) = "621" Or Left(taikhoan.sohieu, 3) = "622" Or Left(taikhoan.sohieu, 3) = "623" Or Left(taikhoan.sohieu, 3) = "627" Or Left(taikhoan.sohieu, 3) = "154" Or Left(taikhoan.sohieu, 3) = "911")) And Len(txtchungtu(2).Text) > 0 Then
            tp.InitTPSohieu txtchungtu(2).Text
            txtchungtu(1).Text = tp.TenVattu
            If Left(taikhoan.sohieu, 3) = "154" And tp.MaSo > 0 Then
                tien = tp.GiaThanhCK(CboThang.ItemData(CboThang.ListIndex))
                If tien <> 0 Then
                    txtchungtu(5).Text = ""
                    txtchungtu(6).Text = Format(tien, Mask_2)
                End If
            End If
        Else
            tp.InitTPMaSo 0
        End If

        If ((taikhoan.tk_id = TKCNKH_ID) Or (taikhoan.tk_id = TKCNPT_ID)) Then
            ckh.InitKhachHangSohieu txtchungtu(2).Text
            txtchungtu(1).Text = ckh.Ten
            txtchungtu(3).Enabled = (ckh.MaNT > 0)
            txtchungtu(2).tag = IIf(ckh.MaNT > 0, 1, 0)
            If ckh.MaNT = 0 Then
                txtchungtu(3).Text = "0"
            Else
                txtchungtu(4).Enabled = True
                txtchungtu(4).Text = Format(TyGiaNT(ckh.MaNT), Mask_0)
                RFocus txtchungtu(3)
            End If
        End If
        If loaict = 9 Then VTEnable = True
        If (((loaict = 1 Or loaict = 2 Or loaict = 8 Or loaict = 9) And (taikhoan.tk_id = TKVT_ID)) Or ((loaict = 7 Or loaict = 8) And (taikhoan.tk_id = TKDT_ID Or taikhoan.tk_id = TKGT_ID))) And STDetail And VTEnable And tp.MaSo = 0 Then
            vattu.InitVattuSohieu txtchungtu(2).Text
            If vattu.MaSo > 0 Then
                VTEnable = False
                Label(25).Visible = (pGiaUSD > 0) And (loaict = 1 Or loaict = 2 Or loaict = 8)
                txtchungtu(11).Visible = (pGiaUSD > 0) And (loaict = 1 Or loaict = 2 Or loaict = 8)
                If (taikhoan.loai = 0 Or (OutCost > 0 And OutCost <> 2)) And loaict = 2 Then
                    FDsNhap.tag = vattu.MaSo
                    MaNhap = FDsNhap.XuatDichDanh(CboThang.ItemData(CboThang.ListIndex), vattu.sohieu + " - " + vattu.TenVattu + ABCtoVNI(" - §.v.t: ") + vattu.DonVi, CboNguon(1).ItemData(CboNguon(1).ListIndex), luong, tien)
                    If luong = 0 Then
                        luong = SoTonKho(CboThang.ItemData(CboThang.ListIndex), CboNguon(1).ItemData(CboNguon(1).ListIndex), taikhoan.MaSo, vattu.MaSo, tien, tien2)
                    End If
                    txtchungtu(3).tag = luong
                    txtchungtu(5).tag = tien2
                    txtchungtu(6).tag = tien
                    txtchungtu(3).Text = Format(luong, Mask_2)
                    txtchungtu(6).Text = Format(tien, Mask_2)
                    If pGiaUSD > 0 Then txtchungtu(11).Text = Format(tien2, Mask_2)
                    If luong <> 0 Then
                        txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien / luong) / Mask_N, Mask_2)
                    End If
                    HienThongBao "Sè l­îng tån kho: " + txtchungtu(3).Text + " - Thµnh tiÒn: " + txtchungtu(6).Text, 1
                Else
                    If loaict <> 8 Then
                        luong = SoTonKho(CboThang.ItemData(CboThang.ListIndex), CboNguon(1).ItemData(CboNguon(1).ListIndex), taikhoan.MaSo, vattu.MaSo, tien, tien2)
                    Else
                        luong = SoTonKho(CboThang.ItemData(CboThang.ListIndex), CboNguon(1).ItemData(CboNguon(1).ListIndex), 0, vattu.MaSo, tien, tien2)
                    End If
                    If IsMissing(LO_XXXX) Then LO_XXXX = ""
                    If Len(LO_XXXX) > 0 Then luong = SL_XXXX
                    If IsImport = False Then
                        txtchungtu(3).Text = Format(luong, Mask_2)
                    End If
                    If loaict = 1 Then
                        txtchungtu(5).Text = Format(tien, Mask_2)
                    Else
                        txtchungtu(3).tag = luong
                        txtchungtu(6).Text = Format(tien, Mask_2)
                    End If
                    If pGiaUSD > 0 Then txtchungtu(11).Text = Format(tien2, Mask_2)
                    txtchungtu(6).tag = tien
                    txtchungtu(5).tag = tien2
                    If luong <> 0 Then
                        If pGiaUSD > 0 Then
                            txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien2 / luong) / Mask_N, Mask_2)
                        Else
                            txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien / luong) / Mask_N, Mask_2)
                        End If
                    End If
                End If
                If loaict = 8 And (vattu.GiaBan1 > 0 Or vattu.GiaBan2 > 0 Or vattu.GiaBan3 > 0 Or CDbl(txttinh_gia_ban.Caption) > 0) Then
                    LayGiaBan
                Else
                    If luong <> 0 Then
                        If pGiaUSD > 0 Then
                            txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien2 / luong) / Mask_N, Mask_2)
                        Else
                            txtchungtu(4).Text = Format(Fix(0.5 + Mask_N * tien / (luong * IIf(pGiaUSD > 0, pRate, 1))) / Mask_N, Mask_2)
                        End If
                    End If
                End If
                If loaict = 1 And pGiaHT > 0 And vattu.GiaHT > 0 And Left(taikhoan.sohieu, Len(ShTkTP)) = ShTkTP Then
                    txtchungtu(4).Text = Format(vattu.GiaHT, Mask_2)
                End If
                HienThongBao "Sè l­îng tån kho: " + txtchungtu(3).Text + " - Thµnh tiÒn: " + Format(txtchungtu(6).tag, Mask_0), 1

                txtchungtu(1).Text = vattu.TenVattu
                txtchungtu(2).tag = 1
                If IsImport = False Then
                    RFocus txtchungtu(3)
                End If



                VTEnable = True
            End If
        End If
        With GrdChungtu
            tien = 0
            luong = 0
            If (taikhoan.tk_id = GTGTKT_ID) Then
                On Error Resume Next
                j = Abs(CLng5(txtchungtu(2).Text))
                On Error GoTo 0
                '                        If j = 0 Then
                '                            txtchungtu(5).Text = "0"
                '                            txtchungtu(6).Text = "0"
                '                            Exit Sub
                '                        End If
                If Cdbl5(txtchungtu(5).Text) = 0 Then
                    If j > 0 And j < 5 And Right(txtchungtu(2).Text, 1) <> "-" Then txtchungtu(2).Text = txtchungtu(2).Text + "-"
                    For i = 0 To .Rows - 1
                        .Row = i
                        .col = 1
                        sh = .Text
                        If Len(sh) = 0 Then Exit For
                        If Left(sh, Len(pVATV)) = pVATV Then Exit For
                        .col = 6
                        If Right(txtchungtu(2).Text, 1) = "-" And Left(sh, 3) <> "211" Then
                            tien = Cdbl5(.Text)
                            If j > 0 And j < 5 Then
                                v = tien * j / 100
                            Else
                                v = tien * j / (100 + j)
                            End If
                            v = RoundMoney(v)
                            .Text = Format(tien - v, Mask_0)
                            luong = luong + v
                        Else
                            tien = tien + Cdbl5(.Text)
                        End If
                         If Left(sh, 3) = "711" Then
                            .col = 7
                            tien = tien - Cdbl5(.Text)
                        End If
                        If Left(sh, 3) = "338" Then
                            .col = 7
                            tien = tien - Cdbl5(.Text)
                        End If
                    Next
                    If Right(txtchungtu(2).Text, 1) <> "-" Then
                        luong = RoundMoney(tien * j / 100)
                    Else
                        txtchungtu(2).Text = CStr(j)
                    End If
                    txtchungtu(5).Text = Format(luong, Mask_0)
                    txtchungtu(6).Text = "0"
                End If
            End If
            If (taikhoan.tk_id = GTGTPN_ID Or taikhoan.tk_id = TTDB_ID) Then
                If Cdbl5(txtchungtu(2).Text) = 0 Then
                    txtchungtu(5).Text = "0"
                    txtchungtu(6).Text = "0"
                    Exit Sub
                End If
                On Error Resume Next
                j = Abs(CLng5(txtchungtu(2).Text))
                On Error GoTo 0
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 1
                    sh = .Text
                    If Len(sh) = 0 Then Exit For
                    If Left(sh, 4) = "3331" And taikhoan.tk_id = GTGTPN_ID Then Exit For
                    If Left(sh, 2) <> "11" And Left(sh, 4) <> "3331" And Left(sh, 3) <> "521" Then
                        .col = 7
                        If Right(txtchungtu(2).Text, 1) <> "-" And pVAT1 > 0 And taikhoan.tk_id = GTGTPN_ID Then txtchungtu(2).Text = txtchungtu(2).Text + "-"
                        If Right(txtchungtu(2).Text, 1) = "-" And taikhoan.tk_id = GTGTPN_ID Then
                            tien = Cdbl5(.Text)
                            If j > 0 And j < 5 Then
                                v = tien * j / 100
                            Else
                                v = tien * j / (100 + j)
                            End If
                            v = RoundMoney(v)
                            .Text = Format(tien - v, Mask_0)
                            luong = luong + v
                        Else
                            If loaict = 1 Then .col = 6
                            tien = tien + Cdbl5(.Text)
                        End If
                    End If

                    If Left(sh, Len(pSHPT)) <> pSHPT And Left(sh, 2) <> "11" Then            'Left(sh, 3) <> "521" And
                        .col = 6
                        tien = tien - Cdbl5(.Text)
                    End If
                Next
                If taikhoan.tk_id = GTGTPN_ID Then
                    If Right(txtchungtu(2).Text, 1) <> "-" Then
                        luong = RoundMoney(tien * j / 100)
                    Else
                        txtchungtu(2).Text = CStr(j)
                    End If
                Else
                    luong = RoundMoney(tien * j / (j + 100))
                End If
                If luong >= 0 Then
                    txtchungtu(6).Text = Format(luong, Mask_2)
                    txtchungtu(5).Text = "0"
                Else
                    txtchungtu(5).Text = Format(-luong, Mask_2)
                    txtchungtu(6).Text = "0"
                End If
                If luong <> 0 Then
                    '   RFocus txtchungtu(6)
                End If
            End If
        End With

        If loaict = 8 And Left(taikhoan.sohieu, 3) = "521" Then
            luong = Cdbl5(txtchungtu(2).Text)
            If luong > 0 And luong < 100 Then
                txtchungtu(5).Text = Format(RoundMoney(GiaTriTruocThue * luong / 100), Mask_0)
            Else
                txtchungtu(5).Text = Format(SoChietKhau, Mask_0)
            End If
        End If

        If (loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8) And vattu.MaSo > 0 And vattu.Dvt2 > 0 Then
            Label(12).Visible = True
            CboNT(2).Visible = True
            Int_RecsetToCbo "SELECT MaSo AS F2, DonVi AS F1 FROM DVTVattu WHERE MaVattu=" + CStr(vattu.MaSo) + " ORDER BY DonVi", CboNT(2)
            CboNT(2).AddItem vattu.DonVi, 0
            CboNT(2).ListIndex = 0
            RFocus CboNT(2)
        End If

        txtchungtu(9).Text = ""
        txtchungtu(10).Text = ""
        txtchungtu(4).Enabled = ((loaict = 1 Or loaict = 7 Or loaict = 8) And (vattu.MaSo > 0)) Or (ckh.MaNT > 0)
        If loaict = 8 And vattu.MaSo > 0 And vattu.CK > 0 And pChietKhau > 0 Then
            txtchungtu(9).Text = Format(vattu.CK, IIf(vattu.CK * 100 Mod 100 <> 0, Mask_2, Mask_0))
            TinhCKCT
        End If
    Case 3:    ' So luong
        txtchungtu(3).Text = Format(txtchungtu(3).Text, Mask_2)
        If (loaict = 2 Or loaict = 7 Or loaict = 8 Or loaict = 9) And vattu.MaSo > 0 Then
            luong = Cdbl5(txtchungtu(3).Text)
            txtchungtu(5).Text = "0"
            If loaict = 2 Then
                If luong <> txtchungtu(3).tag And txtchungtu(3).tag <> 0 Then
                    txtchungtu(6).Text = Format((luong * txtchungtu(6).tag) / txtchungtu(3).tag, Mask_2)
                    If pGiaUSD > 0 Then txtchungtu(11).Text = Format((luong * txtchungtu(5).tag) / txtchungtu(3).tag, Mask_2)
                Else
                    txtchungtu(6).Text = Format(txtchungtu(6).tag, Mask_0)
                    If pGiaUSD > 0 Then txtchungtu(11).Text = Format(txtchungtu(5).tag, Mask_2)
                End If
            Else
                If pGiaUSD > 0 Then txtchungtu(11).Text = Format(luong * Cdbl5(txtchungtu(4).Text), Mask_2)
                txtchungtu(6).Text = Format(luong * Cdbl5(txtchungtu(4).Text) * IIf(pGiaUSD > 0, pRate, 1), Mask_0)
            End If
        End If
        If (taikhoan.MaNT <> 0 And CboNT(0).ListIndex > 0) Or ckh.MaNT > 0 Then
            luong = Cdbl5(txtchungtu(3).Text)
            If Cdbl5(txtchungtu(4).Text) <> 0 Then
                If SoPSConLai < 0 Then
                    txtchungtu(5).Text = Format(luong * Cdbl5(txtchungtu(4).Text), Mask_2)
                Else
                    txtchungtu(6).Text = Format(luong * Cdbl5(txtchungtu(4).Text), Mask_2)
                End If
            End If
        End If
        If loaict = 8 And vattu.MaSo > 0 Then TinhCKCT
    Case 4, 5, 6:    ' PS No, Co
        Dim psnt As Boolean, tygia As Double, m As String, nt As Double

        m = IIf(Left(taikhoan.sohieu, 3) = "007", Mask_2, Mask_0)
        If Index = 4 Then
            txtchungtu(Index).Text = Format(txtchungtu(Index).Text, Mask_2)
        Else
            txtchungtu(Index).Text = Format(Cdbl5(txtchungtu(Index).Text), m)
        End If

        If Index = 5 Or Index = 6 Then
            If Cdbl5(txtchungtu(Index).Text) <> 0 Then txtchungtu(11 - Index).Text = "0"
        End If

        psnt = (CboNT(0).Visible And CboNT(0).ListIndex > 0) Or ckh.MaNT > 0
        If ((loaict = 1 Or loaict = 7 Or loaict = 8) And (vattu.MaSo > 0)) Or psnt Then
            luong = Cdbl5(txtchungtu(3).Text)
            Select Case Index
            Case 5:
                If loaict = 1 Or psnt Then
                    If luong > 0 And Cdbl5(txtchungtu(5).Text) <> 0 Then
                        If loaict = 1 Or pTien = 0 Then
                            tygia = Cdbl5(txtchungtu(5).Text) / luong
                        Else
                            tygia = luong / Cdbl5(txtchungtu(5).Text)
                        End If
                        If pGiaUSD > 0 And pRate > 0 Then
                            txtchungtu(4).Text = Format(tygia / pRate, Mask_2)
                        Else
                            txtchungtu(4).Text = Format(tygia, Mask_2)
                        End If
                        If psnt Then
                            If ckh.MaNT > 0 Then
                                CapNhatTyGia ckh.MaNT, tygia
                            Else
                                CapNhatTyGia CboNT(0).ItemData(CboNT(0).ListIndex), tygia
                            End If
                        End If
                    End If
                End If
            Case 6:
                If loaict = 7 Or loaict = 8 Or psnt Then
                    If luong > 0 And Cdbl5(txtchungtu(6).Text) <> 0 Then
                        If loaict = 7 Or loaict = 8 Or pTien = 0 Then
                            tygia = Cdbl5(txtchungtu(6).Text) / luong
                            If loaict = 8 And pGiaUSD > 0 Then tygia = tygia / pRate
                        Else
                            If Cdbl5(txtchungtu(5).Text) <> 0 Then tygia = luong / Cdbl5(txtchungtu(5).Text)
                        End If
                        ' tinh lai gia
                        txtchungtu(4).Text = Format(tygia, Mask_2)
                        If psnt Then
                            If ckh.MaNT > 0 Then
                                CapNhatTyGia ckh.MaNT, tygia
                            Else
                                CapNhatTyGia CboNT(0).ItemData(CboNT(0).ListIndex), tygia
                            End If
                        End If
                    End If
                End If
            Case 4:
                tygia = Cdbl5(txtchungtu(4).Text)
                If psnt Then
                    tien = SoPSConLai
                    nt = DoiRaTien(luong, tygia)
                    If tien < 0 Or (tien = 0 And taikhoan.loai < 0) Then
                        txtchungtu(5).Text = Format(nt, Mask_2)
                    Else
                        txtchungtu(6).Text = Format(nt, Mask_2)
                    End If
                    If ckh.MaNT > 0 Then
                        CapNhatTyGia ckh.MaNT, tygia
                    Else
                        CapNhatTyGia CboNT(0).ItemData(CboNT(0).ListIndex), tygia
                    End If
                Else
                    If loaict = 1 Or taikhoan.tk_id = TKGT_ID Then
                        If luong > 0 Then
                            If luong = txtchungtu(3).tag Then
                                txtchungtu(4).Text = Format(txtchungtu(6).tag, Mask_0)
                            Else
                                txtchungtu(5).Text = Format(tygia * luong * IIf(pGiaUSD > 0, pRate, 1), Mask_0)
                                If pGiaUSD > 0 Then txtchungtu(11).Text = Format(tygia * luong, Mask_2)
                            End If
                        Else
                            txtchungtu(5).Text = txtchungtu(4).Text
                        End If
                    Else
                        If luong > 0 Then
                            If pGiaUSD > 0 Then txtchungtu(11).Text = Format(tygia * luong, Mask_2)
                            txtchungtu(6).Text = Format(tygia * luong * IIf(pGiaUSD > 0, pRate, 1), Mask_2)
                            If Cdbl5(txtchungtu(6).Text) <> 0 Then txtchungtu(5).Text = "0"
                        Else
                            If pGiaUSD > 0 Then
                                txtchungtu(6).Text = Format(Cdbl5(txtchungtu(4).Text) * pRate, Mask_2)
                            Else
                                txtchungtu(6).Text = txtchungtu(4).Text
                            End If
                        End If
                    End If
                End If
            End Select
        End If
        If (loaict = 7 Or loaict = 8) And vattu.MaSo > 0 Then TinhCKCT

        '/////////////////////////////////////////// ky thuat dung co va chua cac nut lenh
        'Neu nhan phim khac enter thi se bat thong tin bang tinh bang tay
        If Index = 6 Then
            If txtchungtu(9).Visible = False Then
                If dathuchien = False Then
                    cho_hien_vat = False
                    If Len(Trim(txtchungtu(1).Text)) > 0 And Index = 6 Then
                        If hien_bang_tinh = True Then
                            CmdChitiet_Click
                            ''''''''sua lai theo tieu chuan c
                            'CmdChitiet_chon
                        End If
                        hien_bang_tinh = True
                    End If
                Else
                    dathuchien = False
                End If
            End If
        End If
        '///////////////////////////////////////////

    Case 7:
        If pTygia > 0 Then pRate = Cdbl5(txtchungtu(7).Text)
        txtchungtu(Index).Text = Format(txtchungtu(Index).Text, Mask_2)
    Case 9:
        txtchungtu(Index).Text = Format(txtchungtu(Index).Text, Mask_2)
        If (loaict = 8 Or loaict = 1) And vattu.MaSo > 0 Then TinhCKCT    '
    Case 10:
        txtchungtu(Index).Text = Format(txtchungtu(Index).Text, Mask_0)
        If txtchungtu(9).Visible = True Then
            If dathuchien = False Then
                cho_hien_vat = False
                If Len(Trim(txtchungtu(1).Text)) > 0 And Index = 10 Then
                    'If hien_bang_tinh = True Then
                    CmdChitiet_Click
                    ''''''''sua lai theo tieu chuan c
                    'CmdChitiet_chon
                    ' End If
                    '    hien_bang_tinh = True
                End If
            Else
                dathuchien = False
            End If
        End If
    End Select
    '////////////
    dathuchien = False
End Sub    '====================================================================================================
' §Æt chÕ ®é nhËp cho lo¹i chøng tõ
'====================================================================================================
Public Sub SetLoaiChungtu(loai As Integer)
    Dim vis As Boolean, i As Integer

    If Not SetLoaiEnable Then Exit Sub

    Me.MousePointer = 11

    '    Me.Width = IIf(loai = 8 And pChietKhau > 0, 12030, 10500)
    '   GrdChungtu.Width = IIf(loai = 8 And pChietKhau > 0, 11895, 10335)
    '    CmdChitiet.Left = IIf(loai = 8 And pChietKhau > 0, 11640, 10080)

    '   Me.Width = IIf(loai = 8 And pChietKhau > 0, 12030 + 1800 - 460, 12030 - 200)
    Me.Width = IIf((loai = 8 Or loai = 1) And pChietKhau > 0, 12030 + 1800 - 460 + 1800 + 300, 12030 - 200 + 1680)
    '  Me.Width = IIf(loai = 8 And pChietKhau > 0, 12030 + 1800 - 460 + 1800, 12030 - 200 + 1800)

    GrdChungtu.Width = IIf((loai = 8 Or loai = 1) And pChietKhau > 0, 13211 + 350, 11600)    '11895 + 1950, 11895 + 400)
    CmdChitiet.Left = IIf((loai = 8 Or loai = 1) And pChietKhau > 0, 12030 + 920, 11380)    ' 12030 + 1800, 11895 + 15)

    Label(23).Visible = ((loai = 8 Or loai = 1) And pChietKhau > 0)
    txtchungtu(9).Visible = ((loai = 8 Or loai = 1) And pChietKhau > 0)
    Grid2.Width = IIf((loai = 8 Or loai = 1) And pChietKhau > 0, 9790 + 1700 + 250, 9790)
    Line1(3).Visible = IIf((loai = 8 Or loai = 1) And pChietKhau > 0, False, True)

    WCenter Me

    vis = (loai = 1 Or loai = 2 Or loai = 7 Or loai = 8) And STDetail
    LbKho(1).Visible = vis Or (loai > 8)
    LbKho(1).Caption = IIf(loai = 1 Or loai = 2 Or loai = 7 Or loai = 8, "Kªnh ph©n phèi", "Ph©n lo¹i")
    LbKho(0).Visible = vis
    CboNguon(0).Visible = vis Or (loai > 8)
    CboNguon(1).Visible = vis
    mnDD(25).Visible = vis Or (loai > 8)

    If loai > 8 Then
        Int_RecsetToCbo "SELECT MaSo As F2,SoHieu + ' - ' + Ten As F1" _
                      & " FROM LoaiChungTu WHERE (Cap = 2) AND (CapTren = " + CStr(OptLoai(loai).tag) + ") ORDER BY SoHieu", CboNguon(0)
    End If

    txtchungtu(2).tag = IIf(loai = 1 Or loai = 2, 1, 0)

    XoaPhieuTrenManHinh
    chkXT.Visible = (loai = 1) And pDTTP <> 0

    vis = (loai = 7 Or loai = 8) And (pBaoGia = 1) And ((frmMain.Command(4).Visible And pPhieu = 1) Or (Not frmMain.Command(4).Visible And pPhieu = 0))
    Chk.Visible = vis
    Chk.Value = 0
    pMaBG = 0

    vis = ((loai = 7 Or loai = 8) And pNVBH = 1)
    Label(21).Visible = vis
    txt(3).Visible = vis
    LBNV.Visible = vis

    CmdBC.Visible = (pBarCode > 0) And (loai = 2 Or loai = 8)

    Select Case loai
    Case 1, 2, 7, 8:
        Int_RecsetToCbo "SELECT MaSo As F2,SoHieu + ' - ' + DienGiai As F1 FROM NguonNhapXuat ORDER BY SoHieu", CboNguon(0)
        If CboNguon(1).ListCount = 0 And STDetail Then
            ErrMsg er_KhoHang
            Unload Me
            FrmKho.tag = 1
            FrmKho.Show 1
            Exit Sub
        End If
        If CboNguon(0).ListCount = 0 And STDetail Then
            ErrMsg er_NguonNX
            Unload Me
            FrmNguon.Show 1
            Exit Sub
        End If
        If loai = 1 Then
            chkXT.Value = 0
            ChkXT_Click
        End If
        Select Case loai
        Case 1, 2: CmdPhieu(1).Caption = "&2 PhiÕu NX"
        Case 7: CmdPhieu(1).Caption = "&2 B¸o gi¸"
        Case 8: CmdPhieu(1).Caption = "&2 Ho¸ ®¬n"
        Case 9, 10: CmdPhieu(1).Caption = "&2 PhiÕu TG"
        End Select
    Case 9:
        If MaSoCT = 0 Then
TTS:
            pNghiepVu = NV_TANG
            frmTaiSan.Show 1
            If pMaTaiSan > 0 Then
                ngay(0) = CVDate(MedNgay(0).Text)
                ngay(1) = CVDate(MedNgay(1).Text)
            Else
                XoaPhieuTrenManHinh
            End If
        End If
    Case 10:
        If MaSoCT = 0 Then
            pNghiepVu = NV_GIAM
            frmDSTaiSan.Show 1
        End If
    Case 11:
        If MaSoCT = 0 Then
            pNghiepVu = NV_DGLAI
            frmDSTaiSan.Show 1
        End If
    Case 12:
        If MaSoCT = 0 Then
            pNghiepVu = NV_TKHAO
            frmKhauHao.Show 1
        End If
    End Select

    If loai > 8 And (pGhichungtu = 1 Or MaSoCT > 0) Then
        For i = 0 To 12
            If (i < 4 Or i > 7) And (i <> loai) Then OptLoai(i).Enabled = False
        Next
        CboThang.Enabled = False
        If pGhichungtu = 1 Then
            hienctts
            tscount = tscount + 1
            MaTS(tscount) = pMaTaiSan
            If loai = 9 And MaSoCT = 0 And tscount < 9 Then
                If MsgBox("NhËp bæ sung tµi s¶n cïng chøng tõ?", vbYesNo + vbInformation, App.ProductName) = vbYes Then
                    pMaTaiSan = 0
                    GoTo TTS
                End If
            End If
        End If
    Else
        If loaict > 8 Or (loai > 8 And pGhichungtu = 0) Then
            SetLoaiEnable = False
            OptLoai(0).Value = True
            RFocus OptLoai(0)

            OptLoai(0).Enabled = True
            OptLoai(3).Enabled = True

            OptLoai(1).Enabled = STDetail
            OptLoai(2).Enabled = STDetail
            OptLoai(8).Enabled = STDetail

            OptLoai(9).Enabled = FADetail
            OptLoai(10).Enabled = FADetail
            OptLoai(11).Enabled = FADetail
            OptLoai(12).Enabled = FADetail

            CboThang.Enabled = True
            SetLoaiEnable = True
        End If
    End If
    vBH = 0
    MaSoCT = 0
    loaict = IIf((loai >= 8 And pGhichungtu = 1) Or (loai < 9), loai, 0)
    Me.MousePointer = 0

End Sub
'====================================================================================================
' Thñ tôc xãa néi dung phiÕu trªn cöa sæ
'====================================================================================================
Private Sub XoaPhieuTrenManHinh()
    Dim i As Integer

    If pNghiepVu = NV_TANG Then
        For i = 0 To tscount
            XoaTaiSan MaTS(i)
        Next
    End If
    pGhichungtu = 0
    pNghiepVu = 0
    pMaTaiSan = 0

    Me.Caption = "NhËp chøng tõ kÕ to¸n"
    ClearText Me
    ClearGrid GrdChungtu, GrdChungtu.tag
    taikhoan.InitTaikhoanMaSo 0
    vattu.InitVattuMaSo 0
    ckh.InitKhachHangMaSo 0
    txtchungtu(0).tag = 0
    txtchungtu(2).tag = 0
    For i = 3 To 6
        txtchungtu(i).Text = ""
    Next
    LbUser.Caption = UserName
    CboNT(1).Visible = False
    For i = 0 To 2
        txt(i).Text = "..."
        CmdPhieu(i).Visible = False
    Next
    If CboNguon(2).ListCount > 0 Then CboNguon(2).ListIndex = 0
    nhieunoco = False
    Label(12).Visible = False
    CboNT(2).Visible = False
    TenTC = "..."
    DiachiTC = "..."
    ctgoc = "..."
    TenNX = "..."
    DiaChiNX = "..."
    TenBH = "..."
    DiaChiBH = "..."
    MSTBH = "..."
    MaKHBH = 0
    HanTT = CVDate("01/01/1900")
    Erase HD
    hdcount = -1
    tscount = -1
    XoaHD
    vBH = 0
    If loaict = 1 Then
        chkXT.Value = 0
        ChkXT_Click
    End If
    CmdChitiet.tag = -1
    Command(1).Enabled = ChoNhapTiep
    Command(2).Enabled = True
    Label(22).Enabled = False
    txtchungtu(8).Enabled = False
    Chk.Value = 0
    pMaBG = 0

    txtchungtu(7).Text = Format(pRate, Mask_2)

    Label(25).Visible = False
    txtchungtu(11).Visible = False
    CboNT(3).Visible = False
End Sub
'====================================================================================================
' Thñ tôc hiÓn thÞ néi dung phiÕu trªn mµn h×nh
'====================================================================================================
Public Function HienPhieuTrenManHinh(p As Integer) As Integer
    Dim rs_chungtu, thongtinkhachhang As Recordset
    Dim sh As String, i As Integer, sodong As Integer, ps As Double, ThemDong As Boolean, mct As Long, mts As Long, uid As Long
    Dim ma As Long, diengiai As String, ms As Long, tl As Integer, mvt As Long, mk As Long, mtp As Long, psnt As Double, dgia As Double, luong As Double, st As String

    ma = MaSoCT
    diengiai = ""
    sh = IIf(p > 0, "P", "")
    Dim sql

    'sql = "SELECT ChungTu" + sh + ".*,HoaDon" + sh + ".MaKhachHang,HoaDon" + sh + ".Loai AS LoaiHD,KyHieu,HoaDon" + sh + ".SoHD AS SHD,NgayPH,MatHang,Soluong,Thanhtien,Tyle,HD,KCT,NK,TS,HoaDon" + sh + ".DC,HTTT,MauSo,KhachHang.Ten,KhachHang.DiaChi,KhachHang.MST,khachhang.sohieu as sohieukhachhang,HDBL,HoaDon" + sh + ".TyGia AS TG FROM (ChungTu" + sh + " LEFT JOIN HoaDon" + sh + " ON ChungTu" + sh + ".MaSo=HoaDon" + sh + ".MaSo) LEFT JOIN KhachHang ON ChungTu" + sh + ".MaKH=KhachHang.MaSo WHERE Chungtu" + sh + ".MaCT=" + CStr(MaSoCT) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " ORDER BY Chungtu" + sh + ".MaSo DESC"
    sql = "SELECT ChungTu" + sh + ".*,HoaDon" + sh + ".MaKhachHang,HoaDon" + sh + ".KyHieu as kyhieuhoadon ,HoaDon" + sh + ".Loai AS LoaiHD,KyHieu,HoaDon" + sh + ".SoHD AS SHD,NgayPH,MatHang,Soluong,Thanhtien,Tyle,HD,KCT,NK,TS,HoaDon" + sh + ".DC,HTTT,MauSo,KhachHang.Ten,KhachHang.DiaChi,KhachHang.MST,khachhang.sohieu as sohieukhachhang,HDBL,HoaDon" + sh + ".TyGia AS TG FROM (ChungTu" + sh + " LEFT JOIN HoaDon" + sh + " ON ChungTu" + sh + ".MaSo=HoaDon" + sh + ".MaSo) LEFT JOIN KhachHang ON ChungTu" + sh + ".MaKH =KhachHang.MaSo WHERE Chungtu" + sh + ".MaCT=" + CStr(MaSoCT) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " ORDER BY Chungtu" + sh + ".MaSo DESC"
    Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs_chungtu.recordCount = 0 Then
        MsgBox "PhiÕu ®· bÞ xo¸!", vbCritical, App.ProductName
        HienPhieuTrenManHinh = -1
        GoTo KetThuc
    End If

    OptLoai(IIf(rs_chungtu!maloai <> 7, rs_chungtu!maloai, 8)).Value = True
    XoaPhieuTrenManHinh
    loaict = rs_chungtu!maloai
    SetListIndex CboThang, rs_chungtu!ThangCT
    Chk.Value = IIf(rs_chungtu!maloai = 7, 1, 0)
    pMaBG = IIf(rs_chungtu!maloai = 7, ma, 0)
    ngay(0) = rs_chungtu!NgayCT
    ngay(1) = rs_chungtu!NgayGS
    MedNgay(0).Text = Format(ngay(0), Mask_D)
    MedNgay(1).Text = Format(ngay(1), Mask_D)
    txt(0).Text = rs_chungtu!sohieu
    txt(1).Text = rs_chungtu!diengiai


    If rs_chungtu!kyhieuhoadon <> Null Then txtVT(1).Text = rs_chungtu!kyhieuhoadon
    'lay ma khach hang dua r
    Dim mang1, mang2, mang3
    mang1 = "SELECT ChungTu" + sh + ".sohieu as sh,HoaDon" + sh + ".MaKhachHang,HoaDon" + sh + ".KyHieu as kyhieuhoadon ,khachhang.sohieu,khachhang.tel,khachhang.fax,KhachHang.Ten,KhachHang.DiaChi,KhachHang.MST,khachhang.sohieu as sohieukhachhang,HDBL,HoaDon" + sh + ".TyGia AS TG "
    mang2 = "FROM (ChungTu" + sh + " LEFT JOIN HoaDon" + sh + " ON ChungTu" + sh + ".MaSo=HoaDon" + sh + ".MaSo) "
    mang3 = " LEFT JOIN KhachHang ON hoadon" + sh + ".MaKhachhang = KhachHang.maso WHERE hoadon.makhachhang > 0 and Chungtu" + sh + ".MaCT=" + CStr(MaSoCT) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " ORDER BY Chungtu" + sh + ".MaSo DESC"

    'sql = "SELECT ChungTu" + sh + ".sohieu as sh,HoaDon" + sh + ".MaKhachHang,HoaDon" + sh + ".KyHieu as kyhieuhoadon ,khachhang.sohieu,khachhang.tel,khachhang.fax,KhachHang.Ten,KhachHang.DiaChi,KhachHang.MST,khachhang.sohieu as sohieukhachhang,HDBL,HoaDon" + sh + ".TyGia AS TG FROM (ChungTu" + sh + " LEFT JOIN HoaDon" + sh + " ON ChungTu" + sh + ".MaSo=HoaDon" + sh + ".MaSo) LEFT JOIN KhachHang ON hoadon" + sh + ".MaKhachhang = KhachHang.maso WHERE hoadon.makhachhang > 0 and Chungtu" + sh + ".MaCT=" + CStr(MaSoCT) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " ORDER BY Chungtu" + sh + ".MaSo DESC"
    sql = "SELECT ChungTu" + sh + ".sohieu as sh,HoaDon" + sh + ".MaKhachHang,HoaDon" + sh + ".KyHieu as kyhieuhoadon ,khachhang.sohieu,khachhang.tel,khachhang.fax,KhachHang.Ten,KhachHang.DiaChi,KhachHang.MST,khachhang.sohieu as sohieukhachhang,HDBL,HoaDon" + sh + ".TyGia AS TG FROM (ChungTu" + sh + " LEFT JOIN HoaDon" + sh + " ON ChungTu" + sh + ".MaSo=HoaDon" + sh + ".MaSo) LEFT JOIN KhachHang ON hoadon" + sh + ".MaKhachhang = KhachHang.maso WHERE hoadon.makhachhang > 0 and Chungtu" + sh + ".MaCT=" + CStr(ma) + IIf(pProcessMode = 1, " AND XuLy<2", "") + " ORDER BY Chungtu" + sh + ".MaSo DESC"
    Dim rs

    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.recordCount > 0 Then
        Mo_thong_tin
        txtVT(1).Text = rs!kyhieuhoadon
        txtVT(0).Text = CStr(rs!sohieu)
        txtVT(7).Text = rs!Ten

        If rs!Ten = "Coâng Ty TNHH MTV Vöôøn Phoá Vuõng Taøu" Then
            Text1.Text = rs!Ten
        End If

        txtVT(8).Text = rs!DiaChi
        txtVT(9).Text = rs!mst
        If rs_chungtu!MauSoHD <> "" Then
            txtVT(2).Text = rs_chungtu!MauSoHD
        End If
        If rs_chungtu!LoaiHoaDon <> "" Then
            txtVT(3).Text = rs_chungtu!LoaiHoaDon
        End If
    End If

    ' txtVT(8).Text = rs_chungtu!Ten

    '  End If


    If pSongNgu Then txt(2).Text = rs_chungtu!DienGiaiE
    SetListIndex CboNguon(3), rs_chungtu!CTGS
    uid = rs_chungtu!User_ID
    LbUser.Caption = TenUser(uid)
    If (loaict = 7 Or loaict = 8) And pNVBH > 0 Then
        LBNV.Caption = TenNV(sh, rs_chungtu!MaNV)
        txt(3).Text = sh
        txt(3).tag = rs_chungtu!MaNV
    End If
    If (rs_chungtu!maloai = 1 Or rs_chungtu!maloai = 2 Or rs_chungtu!maloai = 7 Or rs_chungtu!maloai = 8) And STDetail Then
        SetListIndex CboNguon(0), rs_chungtu!MaNguon
        SetListIndex CboNguon(1), rs_chungtu!MaKho
    End If
    If pTygia > 0 Then txtchungtu(7).Text = Format(rs_chungtu!tygia, Mask_2)
    If pGiaUSD > 0 Then pRate = rs_chungtu!tygia
    If (rs_chungtu!maloai > 8) And FADetail Then
        Dim rsts As Recordset

        Set rsts = DBKetoan.OpenRecordset("SELECT DISTINCTROW TOP 1 MaNhom FROM CTTaiSan WHERE MaCTKT = " + CStr(ma), dbOpenSnapshot)
        SetListIndex CboNguon(0), rsts!MaNhom
        rsts.Close
        Set rsts = Nothing
    End If

    SetListIndex CboNguon(2), rs_chungtu!MaDT
    sodong = -1

    LayThongtinCT rs_chungtu!MaCT, 0, TenTC, DiachiTC, ctgoc, MaKHBH, p
    LayThongtinCT rs_chungtu!MaCT, 1, TenNX, DiaChiNX, , , p
    LayThongtinCT rs_chungtu!MaCT, 2, TenBH, DiaChiBH, sh, , p
    LayThongtinCT rs_chungtu!MaCT, 3, unc1, unc2, unc3, , p
    If IsDate(sh) Then HanTT = IIf(Year(CVDate(sh)) >= pNamTC - 1, CVDate(sh), CVDate("01/01/1900")) Else HanTT = CVDate("01/01/1900")
    sh = ""

    If pCongNoHD > 0 Then txtchungtu(8).Text = CStr(rs_chungtu!HanTT)
    If pSoVV > 0 And rs_chungtu!MaDT1 > 0 Then SetListIndex CboVV(0), rs_chungtu!MaDT1
    If pSoVV > 1 And rs_chungtu!MaDT2 > 0 Then SetListIndex CboVV(1), rs_chungtu!MaDT2
    If pSoVV > 2 And rs_chungtu!MaDT3 > 0 Then SetListIndex CboVV(2), rs_chungtu!MaDT3

    Do While Not rs_chungtu.EOF
        If rs_chungtu!MaTkNo <> 0 Then
            taikhoan.InitTaikhoanMaSo rs_chungtu!MaTkNo
            If (taikhoan.tk_id = GTGTKT_ID) Or (taikhoan.tk_id = GTGTPN_ID) Then
                hdcount = hdcount + 1
                ReDim Preserve HD(0 To hdcount) As tpHoaDon
                With HD(hdcount)
                    .MaSo = rs_chungtu!MaSo
                    If Not IsNull(rs_chungtu!MaKhachHang) Then
                        .MaKhachHang = rs_chungtu!MaKhachHang
                        .loai = rs_chungtu!LoaiHD
                        .KyHieu = rs_chungtu!KyHieu
                        .sohd = rs_chungtu!shd
                        .NgayPH = rs_chungtu!NgayPH
                        .MatHang = rs_chungtu!MatHang
                        .SoLuong = rs_chungtu!SoLuong
                        .ThanhTien = rs_chungtu!ThanhTien
                        .TyLe = rs_chungtu!TyLe
                        .HD = rs_chungtu!HD
                        .KCT = rs_chungtu!KCT
                        .HTTT = rs_chungtu!HTTT
                        .MauSo = rs_chungtu!MauSo
                        .HDBL = rs_chungtu!HDBL
                        .NK = rs_chungtu!NK
                        .ts = rs_chungtu!ts
                        .DC = rs_chungtu!DC
                        .tygia = rs_chungtu!tg
                    End If
                    tl = .TyLe
                    If Not KHDetail Then
                        If Not IsNull(rs_chungtu!Ten) Then .TenKH = rs_chungtu!Ten
                        If Not IsNull(rs_chungtu!DiaChi) Then .DiaChiKH = rs_chungtu!DiaChi
                        If Not IsNull(rs_chungtu!mst) Then .MSTKH = rs_chungtu!mst
                    End If
                End With
                If Not IsNull(rs_chungtu!MaKhachHang) Then ms = hdcount Else ms = -1
            Else
                tl = 0
                ms = -1
            End If
            If rs_chungtu!sops <> 0 Or taikhoan.tk_id = GTGTKT_ID Or taikhoan.tk_id = GTGTPN_ID Or taikhoan.tk_id = TTDB_ID Or taikhoan.tk_id = TKVT_ID Or ((taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) And rs_chungtu!MaKHC > 0) Then
                If Not CmdPhieu(0).Visible Then CmdPhieu(0).Visible = (Left(taikhoan.sohieu, Len(TM)) = TM)
                If Not CmdPhieu(1).Visible Then CmdPhieu(1).Visible = (taikhoan.tk_id = TKVT_ID Or taikhoan.tk_id = TKDT_ID Or taikhoan.tk_id = TSCD_ID)
                If Not CmdPhieu(3).Visible Then CmdPhieu(3).Visible = (Left(taikhoan.sohieu, Len(NH)) = NH And KiemTraMaSoThue(frmMain.LbCty(8).Caption, "04"))

                If (((taikhoan.tk_id <> TKVT_ID) Or (Not STDetail)) And (taikhoan.tk_id <> GTGTKT_ID) And (taikhoan.tk_id <> TKDT_ID) And (taikhoan.tk_id <> TKGT_ID) And (taikhoan.tk_id <> TSCD_ID)) And (rs_chungtu!MaTP = 0 Or taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) Then         ' And taikhoan.TK_ID <> TKCNKH_ID And taikhoan.TK_ID <> TKCNPT_ID
                    ThemDong = Not PSDaCo(taikhoan, -1, rs_chungtu!sops, rs_chungtu!SoPS2No, rs_chungtu!makh)
                Else
                    ThemDong = True
                End If
                If ThemDong Then
                    sodong = sodong + 1
                    diengiai = ""

                    If (taikhoan.tk_id = TKVT_ID Or taikhoan.tk_id = TKDT_ID Or taikhoan.tk_id = TKGT_ID) And (rs_chungtu!MaVattu > 0) Then
                        If chkXT.Value = 0 And pDTTP <> 0 Then
                            sh = SelectSQL("SELECT HethongTK.SoHieu AS F1,IIF(MaTP>0,MaTP,MaKH) AS F2 FROM " + ChungTu2TKNC(-1) + " WHERE CT_ID=900000000+" + CStr(rs_chungtu!MaSo), mct)
                            If sh <> "0" Then
                                chkXT.Value = 1
                                ChkXT_Click
                                txtsh(0).Text = sh
                                txtsh_LostFocus 0
                                If cmd(0).tag = 1 And mct > 0 Then
                                    txtsh(1).Text = MaSo2SoHieu(mct, "KhachHang")
                                    txtsh_LostFocus 1
                                End If
                                ' bo chon chung tu
                                If cmd(0).tag = 2 And mct > 0 Then
                                    txtsh(1).Text = MaSo2SoHieu(mct, "TP154")
                                    txtsh_LostFocus 1
                                End If
                                If Left(txtsh(0).Text, 3) = "154" Then
                                    txtsh(1).Text = MaSo2SoHieu(mct, "TP154")
                                    txtsh_LostFocus 1
                                End If
                            End If
                        End If
                        vattu.InitVattuMaSo rs_chungtu!MaVattu
                        sh = vattu.sohieu
                        diengiai = vattu.TenVattu
                        diengiai = vattu.TenVattu + IIf((loaict = 1 Or loaict = 2 Or loaict = 8) And KtraDVT(vattu.MaSo, rs_chungtu!dvt, st), " - " + ABCtoVNI("§.v.t: ") + st, "")
                    Else
                        vattu.InitVattuMaSo 0
                        If (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) And (rs_chungtu!makh > 0) Then
                            ckh.InitKhachHangMaSo rs_chungtu!makh
                            sh = ckh.sohieu
                            diengiai = ckh.Ten
                        Else
                            ckh.InitKhachHangMaSo 0
                            If taikhoan.MaNT > 0 Then sh = SoHieuNT(taikhoan.MaNT) Else sh = IIf(tl > 0, CStr(tl), "")
                            diengiai = IIf(pNN = 0, taikhoan.Ten, taikhoan.TenE)
                        End If
                    End If
                    If rs_chungtu!MaTP <> 0 And (taikhoan.tk_id <> TKVT_ID) And ckh.MaSo = 0 And (taikhoan.loai = 6 Or Left(taikhoan.sohieu, 3) = "154" Or Left(taikhoan.sohieu, 3) = "911") Then
                        tp.InitTPMaSo rs_chungtu!MaTP
                        sh = tp.sohieu
                        diengiai = tp.TenVattu
                    End If
                    If (loaict = 9 Or loaict = 11) And taikhoan.tk_id = TSCD_ID Then
                        mts = SelectSQL("SELECT MaTS AS F1 FROM CTTaiSan WHERE MaTS>0 AND MaTS<>" + CStr(mts) + " AND MaCTKT=" + CStr(rs_chungtu!MaCT) + " AND (ABS(NG_NS+NG_TBS+NG_TD+NG_CNK)=" + DoiDau(rs_chungtu!sops) + " OR ABS(CL_NS+CL_TBS+CL_TD+CL_CNK)=" + DoiDau(rs_chungtu!sops) + ")")
                        If mts > 0 Then
                            sh = MaSo2SoHieu(mts, "TaiSan")
                            diengiai = TenTS("", mts)
                        End If
                    End If
                    If Len(diengiai) = 0 Then diengiai = IIf(pNN = 0, taikhoan.Ten, taikhoan.TenE)
                    If taikhoan.tk_id <> GTGTKT_ID And taikhoan.tk_id <> GTGTPN_ID And taikhoan.tk_id <> TKCNKH_ID And taikhoan.tk_id <> TKCNPT_ID And taikhoan.tk_id <> TKDT_ID And taikhoan.tk_id <> TSCD_ID Then
                        With GrdChungtu
                            For i = 0 To .Rows - 1
                                .Row = i
                                .col = 8
                                If Len(.Text) = 0 Then Exit For
                                If CLng5(.Text) = taikhoan.MaSo Then
                                    .col = 17
                                    mk = CLng5(.Text)
                                    .col = 21
                                    mtp = CLng5(.Text)
                                    .col = 9
                                    If tp.MaSo = mtp And taikhoan.tk_id = TKVT_ID Then mvt = CLng5(.Text) Else mvt = 0
                                    .col = 4
                                    luong = Cdbl5(.Text)
                                    If tp.MaSo = mtp And vattu.MaSo = mvt And ckh.MaSo = mk And (rs_chungtu!SoPS2No = 0 Or luong = 0) Then
                                        .Text = Format(rs_chungtu!SoPS2No + luong, Mask_2)
                                        .col = 6
                                        .Text = Format(rs_chungtu!sops + Cdbl5(.Text), Mask_0)
                                        GoTo KT1
                                    End If
                                End If
                            Next
                        End With
                    End If

                    psnt = IIf((loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8) And vattu.MaSo > 0 And rs_chungtu!dvt > 0, QuyDoiTheoDVT2(vattu.MaSo, rs_chungtu!dvt, rs_chungtu!SoPS2No), rs_chungtu!SoPS2No)
                    luong = IIf(taikhoan.tk_id = TKVT_ID Or taikhoan.tk_id = TKDT_ID Or taikhoan.MaNT > 0 Or ckh.MaNT > 0, psnt, 0)
                    If luong <> 0 Then dgia = Fix(0.5 + Mask_N * rs_chungtu!sops / luong) / Mask_N Else dgia = 0
                    If (loaict = 1 Or loaict = 2 Or loaict = 8) And pGiaUSD > 0 And pRate > 0 And rs_chungtu!MaVattu > 0 Then
                        ps = rs_chungtu!PSUSD
                    Else
                        ps = 0
                    End If
                    If Mid(taikhoan.sohieu, 1, 3) = "156" Or Mid(taikhoan.sohieu, 1, 3) = "154" Or Mid(taikhoan.sohieu, 1, 3) = "155" Or Mid(taikhoan.sohieu, 1, 3) = "511" Then

                        '                        GrdChungtu.AddItem "" + Chr(9) + taikhoan.sohieu + Chr(9) + diengiai + Chr(9) + sh _
                                                 '                            + Chr(9) + IIf(luong <> 0, Format(luong, Mask_2), "") + Chr(9) + IIf(dgia <> 0, Format(dgia, Mask_2), "") _
                                                 '                            + Chr(9) + Format(rs_chungtu!sops, Mask_0) + Chr(9) + "" + Chr(9) + CStr(taikhoan.MaSo) _
                                                 '                            + Chr(9) + CStr(rs_chungtu!MaVattu) + Chr(9) + IIf(taikhoan.loai > 0, "0", "1") + Chr(9) + CStr(taikhoan.MaTC) _
                                                 '                            + Chr(9) + CStr(Abs(rs_chungtu!CT_ID)) + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + CStr(rs_chungtu!makh) _
                                                 '                            + Chr(9) + CStr(ms) + Chr(9) + "" + Chr(9) + CStr(rs_chungtu!MaKHC) + Chr(9) + CStr(rs_chungtu!MaTP) + Chr(9) + Format(rs_chungtu!SoXuat, Mask_2) _
                                                 '                            + Chr(9) + CStr(rs_chungtu!dvt) + Chr(9) + Format(ps, Mask_2) + Chr(9) + "" + Chr(9) + "" + Chr(9) + IIf(IsNull(rs_chungtu!solo), "", rs_chungtu!solo) + Chr(9) + CStr(IIf(IsNull(rs_chungtu!handung), "", rs_chungtu!handung)), 0
                        GrdChungtu.AddItem "" + Chr(9) + taikhoan.sohieu + Chr(9) + diengiai + Chr(9) + sh _
                                         + Chr(9) + IIf(luong <> 0, Format(luong, Mask_2), "") + Chr(9) + IIf(dgia <> 0, Format(dgia, Mask_2), "") _
                                         + Chr(9) + Format(rs_chungtu!sops, Mask_0) + Chr(9) + "" + Chr(9) + CStr(taikhoan.MaSo) _
                                         + Chr(9) + CStr(rs_chungtu!MaVattu) + Chr(9) + IIf(taikhoan.loai > 0, "0", "1") + Chr(9) + CStr(taikhoan.MaTC) _
                                         + Chr(9) + CStr(Abs(rs_chungtu!CT_ID)) + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + CStr(rs_chungtu!makh) _
                                         + Chr(9) + CStr(ms) + Chr(9) + "" + Chr(9) + CStr(rs_chungtu!MaKHC) + Chr(9) + CStr(rs_chungtu!MaTP) + Chr(9) + Format(rs_chungtu!SoXuat, Mask_2) _
                                         + Chr(9) + CStr(rs_chungtu!dvt) + Chr(9) + Format(ps, Mask_2) + Chr(9) + Format(rs_chungtu!phantramchietkhau, Mask_2) + Chr(9) + Format(rs_chungtu!sotienchietkhau, Mask_0) + Chr(9) + IIf(IsNull(rs_chungtu!solo), "", rs_chungtu!solo) + Chr(9) + CStr(IIf(IsNull(rs_chungtu!handung), "", rs_chungtu!handung)), 0

                    Else
                        If (Mid(taikhoan.sohieu, 1, 3) = "133" Or Mid(taikhoan.sohieu, 1, 3) = "333") Or ((Mid(taikhoan.sohieu, 1, 3) <> "133" And Mid(taikhoan.sohieu, 1, 3) <> "333") And (rs_chungtu!sops + luong <> 0)) Then
                            GrdChungtu.AddItem "" + Chr(9) + taikhoan.sohieu + Chr(9) + diengiai + Chr(9) + sh _
                                             + Chr(9) + IIf(luong <> 0, Format(luong, Mask_2), "") + Chr(9) + IIf(dgia <> 0, Format(dgia, Mask_2), "") _
                                             + Chr(9) + Format(rs_chungtu!sops, Mask_0) + Chr(9) + "" + Chr(9) + CStr(taikhoan.MaSo) _
                                             + Chr(9) + CStr(rs_chungtu!MaVattu) + Chr(9) + IIf(taikhoan.loai > 0, "0", "1") + Chr(9) + CStr(taikhoan.MaTC) _
                                             + Chr(9) + CStr(Abs(rs_chungtu!CT_ID)) + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) + CStr(rs_chungtu!makh) _
                                             + Chr(9) + CStr(ms) + Chr(9) + "" + Chr(9) + CStr(rs_chungtu!MaKHC) + Chr(9) + CStr(rs_chungtu!MaTP) + Chr(9) + Format(rs_chungtu!SoXuat, Mask_2) _
                                             + Chr(9) + CStr(rs_chungtu!dvt) + Chr(9) + Format(ps, Mask_2) + Chr(9) + "" + Chr(9) + "" + Chr(9) + IIf(IsNull(rs_chungtu!solo), "", "") + Chr(9) + CStr(IIf(IsNull(rs_chungtu!handung), "", "")), 0
                        End If
                    End If
KT1:
                End If
            End If
        End If
        '            GrdChungtu.CellSelected = True
        ' GrdChungtu.SelStartRow = 8
        If rs_chungtu!MaTkCo <> 0 Then
            taikhoan.InitTaikhoanMaSo rs_chungtu!MaTkCo
            If (taikhoan.tk_id = GTGTPN_ID Or taikhoan.tk_id = TTDB_ID) Or (taikhoan.tk_id = GTGTKT_ID And rs_chungtu!sops = 0) Then
                hdcount = hdcount + 1
                ReDim Preserve HD(0 To hdcount) As tpHoaDon
                With HD(hdcount)
                    .MaSo = rs_chungtu!MaSo
                    If Not IsNull(rs_chungtu!MaKhachHang) Then
                        .MaKhachHang = rs_chungtu!MaKhachHang
                        .loai = rs_chungtu!LoaiHD
                        .KyHieu = rs_chungtu!KyHieu
                        .sohd = rs_chungtu!shd
                        .NgayPH = rs_chungtu!NgayPH
                        .MatHang = rs_chungtu!MatHang
                        .SoLuong = rs_chungtu!SoLuong
                        .ThanhTien = rs_chungtu!ThanhTien
                        .TyLe = rs_chungtu!TyLe
                        .HD = rs_chungtu!HD
                        .KCT = rs_chungtu!KCT
                        .HTTT = rs_chungtu!HTTT
                        .MauSo = rs_chungtu!MauSo
                        .HDBL = rs_chungtu!HDBL
                        .NK = rs_chungtu!NK
                        .ts = rs_chungtu!ts
                        .DC = rs_chungtu!DC
                        .tygia = rs_chungtu!tg
                    End If
                    tl = .TyLe
                    If Not KHDetail Then
                        If Not IsNull(rs_chungtu!Ten) Then .TenKH = rs_chungtu!Ten
                        If Not IsNull(rs_chungtu!DiaChi) Then .DiaChiKH = rs_chungtu!DiaChi
                        If Not IsNull(rs_chungtu!mst) Then .MSTKH = rs_chungtu!mst
                    End If
                    If .MaKhachHang <> 0 Then
                        ckh.InitKhachHangMaSo .MaKhachHang
                        TenBH = ckh.Ten
                        DiaChiBH = ckh.DiaChi
                        MSTBH = ckh.mst
                        MaKHBH = ckh.MaSo
                        ckh.InitKhachHangMaSo 0
                    End If
                End With
                If Not IsNull(rs_chungtu!MaKhachHang) Then ms = hdcount Else ms = -1
            Else
                tl = 0
                ms = -1
            End If
            If rs_chungtu!sops <> 0 Or taikhoan.tk_id = TKDT_ID Or taikhoan.tk_id = GTGTKT_ID Or taikhoan.tk_id = GTGTPN_ID Or taikhoan.tk_id = TTDB_ID Or taikhoan.tk_id = TKVT_ID Or ((taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) And rs_chungtu!MaKHC > 0) Then
                If Not CmdPhieu(0).Visible Then CmdPhieu(0).Visible = (Left(taikhoan.sohieu, Len(TM)) = TM)
                If Not CmdPhieu(1).Visible Then CmdPhieu(1).Visible = (taikhoan.tk_id = TKVT_ID Or taikhoan.tk_id = TKDT_ID Or taikhoan.tk_id = TSCD_ID)
                If Not CmdPhieu(2).Visible Then CmdPhieu(2).Visible = (Left(taikhoan.sohieu, Len(NH)) = NH Or taikhoan.tk_id2 = CLng(NH))
                CmdPhieu(2).tag = taikhoan.sohieu
                If Not CmdPhieu(3).Visible Then CmdPhieu(3).Visible = (Left(taikhoan.sohieu, Len(NH)) = NH And KiemTraMaSoThue(frmMain.LbCty(8).Caption, "04"))

                If ((taikhoan.tk_id <> TKVT_ID And taikhoan.tk_id <> TKDT_ID) Or (Not STDetail)) And (taikhoan.tk_id <> TKDT_ID) And (taikhoan.tk_id <> GTGTPN_ID) And (taikhoan.tk_id <> TTDB_ID) And (rs_chungtu!MaTP = 0 Or taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) Then       ' And taikhoan.TK_ID <> TKCNKH_ID And taikhoan.TK_ID <> TKCNPT_ID
                    ThemDong = Not PSDaCo(taikhoan, 1, rs_chungtu!sops, rs_chungtu!SoPS2Co, rs_chungtu!MaKHC)
                Else
                    ThemDong = True
                End If

                If ThemDong Then
                    sodong = sodong + 1

                    If (taikhoan.tk_id = TKVT_ID Or taikhoan.tk_id = TKDT_ID) And (rs_chungtu!MaVattu > 0) Then
                        vattu.InitVattuMaSo rs_chungtu!MaVattu
                        sh = vattu.sohieu
                        diengiai = vattu.TenVattu + IIf((loaict = 1 Or loaict = 2 Or loaict = 8) And KtraDVT(vattu.MaSo, rs_chungtu!dvt, st), " - " + ABCtoVNI("§.v.t: ") + st, "")
                    Else
                        vattu.InitVattuMaSo 0
                        If (taikhoan.tk_id = TKCNKH_ID Or taikhoan.tk_id = TKCNPT_ID) And (rs_chungtu!MaKHC > 0) Then
                            ckh.InitKhachHangMaSo rs_chungtu!MaKHC
                            sh = ckh.sohieu
                            diengiai = ckh.Ten
                        Else
                            ckh.InitKhachHangMaSo 0
                            If taikhoan.MaNT > 0 Then sh = SoHieuNT(taikhoan.MaNT) Else sh = IIf(tl > 0, CStr(tl), "")
                            diengiai = IIf(pNN = 0, taikhoan.Ten, taikhoan.TenE)
                        End If
                    End If
                    If rs_chungtu!MaTP <> 0 And (taikhoan.tk_id <> TKVT_ID) And ckh.MaSo = 0 And (taikhoan.tk_id2 = TKDT_ID Or taikhoan.loai = 6 Or Left(taikhoan.sohieu, 3) = "154") Then
                        tp.InitTPMaSo rs_chungtu!MaTP
                        sh = tp.sohieu
                        diengiai = tp.TenVattu
                    End If
                    If (loaict = 10 Or loaict = 11) And taikhoan.tk_id = TSCD_ID Then
                        mts = SelectSQL("SELECT MaTS AS F1 FROM CTTaiSan WHERE MaTS>0 AND MaTS<>" + CStr(mts) + " AND MaCTKT=" + CStr(rs_chungtu!MaCT) + " AND ABS(CL_NS+CL_TBS+CL_TD+CL_CNK)=" + DoiDau(rs_chungtu!sops))
                        If mts > 0 Then
                            sh = MaSo2SoHieu(mts, "TaiSan")
                            diengiai = TenTS("", mts)
                        End If
                    End If
                    If Len(diengiai) = 0 Then diengiai = IIf(pNN = 0, taikhoan.Ten, taikhoan.TenE)
                    If taikhoan.tk_id <> GTGTKT_ID And taikhoan.tk_id <> GTGTPN_ID And taikhoan.tk_id <> TKCNKH_ID And taikhoan.tk_id <> TKCNPT_ID And taikhoan.tk_id <> TKVT_ID And taikhoan.tk_id <> TKDT_ID Then           '  And (taikhoan.tk_id <> TKDT_ID Or rs_chungtu!sops <> 0)
                        With GrdChungtu
                            For i = 0 To .Rows - 1
                                .Row = i
                                .col = 8
                                If Len(.Text) = 0 Then Exit For
                                If CLng5(.Text) = taikhoan.MaSo Then
                                    .col = 20
                                    mk = CLng5(.Text)
                                    .col = 21
                                    mtp = CLng5(.Text)
                                    .col = 9
                                    If (tp.MaSo = mtp And taikhoan.tk_id = TKVT_ID) Or (taikhoan.tk_id = TKDT_ID And rs_chungtu!sops <> 0) Then mvt = CLng5(.Text) Else mvt = 0
                                    If tp.MaSo = mtp And vattu.MaSo = mvt And ckh.MaSo = mk Then
                                        .col = 7
                                        ps = Cdbl5(.Text)
                                        If ps <> 0 Then
                                            .Text = Format(rs_chungtu!sops + ps, Mask_0)
                                            .col = 4
                                            .Text = Format(rs_chungtu!SoPS2Co + Cdbl5(.Text), Mask_0)
                                            GoTo KT2
                                        End If
                                    End If
                                End If
                            Next
                        End With
                    End If
                    psnt = IIf((loaict = 1 Or loaict = 2 Or loaict = 7 Or loaict = 8) And vattu.MaSo > 0 And rs_chungtu!dvt > 0, QuyDoiTheoDVT2(vattu.MaSo, rs_chungtu!dvt, rs_chungtu!SoPS2Co), rs_chungtu!SoPS2Co)
                    luong = IIf(taikhoan.tk_id = TKVT_ID Or taikhoan.tk_id = TKDT_ID Or taikhoan.MaNT > 0 Or ckh.MaNT > 0, psnt, 0)
                    If luong <> 0 Then dgia = Fix(0.5 + Mask_N * rs_chungtu!sops / luong) / Mask_N Else dgia = 0
                    If (loaict = 1 Or loaict = 2 Or loaict = 8) And pGiaUSD > 0 And pRate > 0 And rs_chungtu!MaVattu > 0 Then
                        ps = rs_chungtu!PSUSD
                    Else
                        ps = 0
                    End If
                    If Mid(taikhoan.sohieu, 1, 3) = "156" Or Mid(taikhoan.sohieu, 1, 3) = "154" Or Mid(taikhoan.sohieu, 1, 3) = "155" Or Mid(taikhoan.sohieu, 1, 3) = "511" Then

                        GrdChungtu.AddItem "" + Chr(9) + taikhoan.sohieu + Chr(9) + diengiai + Chr(9) + sh _
                                         + Chr(9) + IIf(luong <> 0, Format(luong, Mask_2), "") + Chr(9) + IIf(dgia <> 0, Format(dgia, Mask_2), "") _
                                         + Chr(9) + "" + Chr(9) + Format(rs_chungtu!sops, Mask_0) + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) + CStr(rs_chungtu!MaVattu) + Chr(9) _
                                         + IIf(taikhoan.loai > 0, "0", "1") + Chr(9) + CStr(taikhoan.MaTC) + Chr(9) + CStr(Abs(rs_chungtu!CT_ID)) + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) _
                                         + "" + Chr(9) + CStr(rs_chungtu!makh) + Chr(9) + CStr(ms) + Chr(9) + "" + Chr(9) + CStr(rs_chungtu!MaKHC) + Chr(9) + CStr(rs_chungtu!MaTP) + Chr(9) _
                                         + Format(rs_chungtu!SoXuat, Mask_2) + Chr(9) + CStr(rs_chungtu!dvt) + Chr(9) + Format(ps, Mask_2) + Chr(9) + Format(rs_chungtu!TLCK, Mask_2) + Chr(9) + Format(rs_chungtu!CK, Mask_0) + Chr(9) + IIf(IsNull(rs_chungtu!solo), "", rs_chungtu!solo) + Chr(9) + CStr(IIf(IsNull(rs_chungtu!handung), "", rs_chungtu!handung)), 0
                    Else
                        If (Mid(taikhoan.sohieu, 1, 3) = "133" Or Mid(taikhoan.sohieu, 1, 3) = "333") Or ((Mid(taikhoan.sohieu, 1, 3) <> "133" And Mid(taikhoan.sohieu, 1, 3) <> "333") And (rs_chungtu!sops + luong <> 0)) Then
                            GrdChungtu.AddItem "" + Chr(9) + taikhoan.sohieu + Chr(9) + diengiai + Chr(9) + sh _
                                             + Chr(9) + IIf(luong <> 0, Format(luong, Mask_2), "") + Chr(9) + IIf(dgia <> 0, Format(dgia, Mask_2), "") _
                                             + Chr(9) + "" + Chr(9) + Format(rs_chungtu!sops, Mask_0) + Chr(9) + CStr(taikhoan.MaSo) + Chr(9) + CStr(rs_chungtu!MaVattu) + Chr(9) _
                                             + IIf(taikhoan.loai > 0, "0", "1") + Chr(9) + CStr(taikhoan.MaTC) + Chr(9) + CStr(Abs(rs_chungtu!CT_ID)) + Chr(9) + "" + Chr(9) + "" + Chr(9) + "" + Chr(9) _
                                             + "" + Chr(9) + CStr(rs_chungtu!makh) + Chr(9) + CStr(ms) + Chr(9) + "" + Chr(9) + CStr(rs_chungtu!MaKHC) + Chr(9) + CStr(rs_chungtu!MaTP) + Chr(9) _
                                             + Format(rs_chungtu!SoXuat, Mask_2) + Chr(9) + CStr(rs_chungtu!dvt) + Chr(9) + Format(ps, Mask_2) + Chr(9) + Format(rs_chungtu!TLCK, Mask_2) + Chr(9) + Format(rs_chungtu!CK, Mask_0) + Chr(9) + IIf(IsNull(rs_chungtu!solo), "", "") + Chr(9) + CStr(IIf(IsNull(rs_chungtu!handung), "", "")), 0
                        End If
                    End If
KT2:
                End If
            End If
        End If
        vattu.InitVattuMaSo 0
        ckh.InitKhachHangMaSo 0
        tp.InitTPMaSo 0
        rs_chungtu.MoveNext
    Loop
    GrdChungtu.Rows = IIf(sodong >= GrdChungtu.tag, sodong + 1, GrdChungtu.tag)
    'taikhoan.InitTaikhoanMaSo 0
    'vattu.InitVattuMaSo 0
    'ckh.InitKhachHangMaSo 0
    MaSoCT = ma
    If loaict = 2 Or loaict = 8 Then
        xddu = SetDoiUng(1)
        If Not xddu Then xddu = SetDoiUng
    Else
        xddu = SetDoiUng
        If Not xddu Then xddu = SetDoiUng(1)
    End If
    GrdChungtu.Row = 0
    Command(1).Enabled = (frmMain.tag Mod 10000000 >= 1000000) And ChoNhapTiep And (User_Right = 0 Or (UserID = uid))
    Command(2).Enabled = (frmMain.tag Mod 10000000 >= 1000000) And (User_Right = 0 Or (UserID = uid))
    If User_Right <> 0 Then
        If SelectSQL("SELECT Lock" + CStr(CboThang.ItemData(CboThang.ListIndex)) + " Mod 10 AS F1 FROM License") > 0 Then
            Command(1).Enabled = False
            Command(2).Enabled = False
        End If
    End If
    NhapDongMoi ""
    ' RFocus txt(0)
    RFocus CboThang
    HienPhieuTrenManHinh = 0
KetThuc:
    rs_chungtu.Close
    Set rs_chungtu = Nothing

End Function
'====================================================================================================
' NhËp dßng ph¸t sinh míi
'====================================================================================================
Private Sub NhapDongMoi(shtk As String)
    Dim ps As Double

    If pCongNoHD > 0 And ((taikhoan.tk_id = TKCNKH_ID And Cdbl5(txtchungtu(5).Text) > 0) Or (taikhoan.tk_id = TKCNPT_ID And Cdbl5(txtchungtu(6).Text) > 0)) Then
        RFocus txtchungtu(8)
    Else
        RFocus txtchungtu(0)
    End If

    CboNT(0).Visible = False

    txtchungtu(0).tag = 0
    txtchungtu(0).Text = shtk
    txtchungtu(1).Text = ""
    txtchungtu(2).Text = ""
    txtchungtu(3).Text = "0"
    txtchungtu(3).tag = 0
    txtchungtu(5).tag = 0
    txtchungtu(6).tag = 0
    txtchungtu(9).Text = ""
    txtchungtu(10).Text = ""

    taikhoan.InitTaikhoanMaSo 0
    vattu.InitVattuMaSo 0
    ckh.InitKhachHangMaSo 0
    tp.InitTPMaSo 0

    If loaict = 1 And STDetail Then txtchungtu(6).Text = "0"
    ps = SoPSConLai
    If ps < 0 Then
        txtchungtu(5).Text = Format(-ps, Mask_2)
        txtchungtu(6).Text = "0"
    Else
        txtchungtu(6).Text = Format(ps, Mask_2)
        txtchungtu(5).Text = "0"
    End If
    MaNhap = 0
    txtchungtu(4).Enabled = False
    CboNT(1).Visible = False
    CboNT(2).Visible = False

    Label(25).Visible = False
    txtchungtu(11).Visible = False

    CboNT(3).Visible = False
End Sub
'====================================================================================================
' Hµm tr¶ vÒ sè ph¸t sinh chªnh lÖch tÝnh theo sè PS nî
'====================================================================================================
Public Function SoPSConLai() As Double
    Dim tong As Double, i As Integer, ps As Double

    tong = 0
    With GrdChungtu
        For i = 0 To .Rows - 1
            .Row = i
            .col = 1
            If Len(.Text) = 0 Then Exit For
            .col = 10
            If .Text = "0" Then
                .col = 6
                ps = Cdbl5(.Text)
                tong = tong + ps
                .col = 7
                ps = Cdbl5(.Text)
                tong = tong - ps
            End If
        Next
    End With
    SoPSConLai = tong
End Function
'====================================================================================================
' Hµm tr¶ kiÓm tra d÷ liÖu nhËp
'====================================================================================================
Private Function KiemTraChungtu() As Boolean
    Dim sodu As Double, sodu2 As Double, st As String

    KiemTraChungtu = False
    If Len(txt(0).Text) = 0 Then txt(0).Text = "..."
    If (loaict = 1 Or loaict = 2 Or loaict = 8) And (CboNguon(0).ListIndex < 0 Or CboNguon(1).ListIndex < 0) Then
        ErrMsg er_NguonNX
        RFocus CboNguon(0)
        Exit Function
    End If
    GrdChungtu.Row = 0
    GrdChungtu.col = 1
    If Len(GrdChungtu.Text) = 0 Then
        RFocus txtchungtu(0)
        Exit Function
    End If
    If pHachToan <> 0 And Fix(SoPSConLai * Mask_N) <> 0 And ((loaict <> 8 And loaict <> 7) Or Chk.Value = 0) Then
        If Not PSTuDong(SoPSConLai) Then
            MsgBox "Sè ph¸t sinh nî cã ch­a c©n b»ng !", vbInformation, App.ProductName
            RFocus txtchungtu(0)
            Exit Function
        End If
    End If
    If loaict = 3 And nhieunoco Then
        MsgBox "Kh«ng nhËp chøng tõ kÕt chuyÓn nhiÒu nî, nhiÒu cã !", vbInformation, App.ProductName
        RFocus txtchungtu(0)
        Exit Function
    End If
    If loaict = 8 And (Not (CoPSTK("5", 0) Or CoPSTK("33", 0))) Then
        MsgBox "Chøng tõ b¸n hµng kh«ng cã tµi kho¶n doanh thu !", vbInformation, App.ProductName
        Exit Function
    End If
    If loaict = 3 And (Not (CoPSTK("9") Or CoPSTK("8") Or CoPSTK("7") Or CoPSTK("6") Or CoPSTK("5") Or CoPSTK("142") Or CoPSTK("421") Or CoPSTK(pVATV))) Then
        MsgBox "ChØ nhËp chøng tõ kÕt chuyÓn trong ph©n lo¹i chøng tõ nµy!", vbInformation, App.ProductName
        Exit Function
    End If
    If User_Right <> 0 Then
        If SelectSQL("SELECT Lock" + CStr(CboThang.ItemData(CboThang.ListIndex)) + " Mod 10 AS F1 FROM License") > 0 Then
            MsgBox "Th¸ng ®· bÞ kho¸ kh«ng cho nhËp sè liÖu!", vbCritical, App.ProductName
            Exit Function
        End If
    End If
    If SelectSQL("SELECT TOP 1 MaCT AS F1 FROM ChungTu WHERE MaLoai=3 AND CT_ID>300000000 AND CT_ID<310000000 AND ThangCT=" + CStr(CboThang.ItemData(CboThang.ListIndex))) > 0 Then
        st = CStr(CThangDB(CboThang.ItemData(CboThang.ListIndex)))
        sodu = SelectSQL("SELECT Sum(IIF(Loai=1 OR Loai=2,DuNo_" + st + "-DuCo_" + st + ",0)) AS F1,Sum(IIF(Loai=3 OR Loai=4,DuNo_" + st + "-DuCo_" + st + ",0)) AS F2 FROM HethongTK WHERE Cap=0", sodu2)
        If sodu = sodu2 Then
            If MsgBox("Th¸ng ®· kÕt chuyÓn, cho nhËp chøng tõ?", vbCritical + vbYesNo, App.ProductName) = vbNo Then Exit Function
        End If
    End If
    If chkXT.Value = 1 Then
        If txtsh(0).tag = 0 Then
            RFocus txtsh(0)
            Exit Function
        End If
        If cmd(0).tag = 1 And txtsh(1).tag = 0 Then
            RFocus txtsh(1)
            Exit Function
        End If
    End If
    KiemTraChungtu = True
End Function
'===========================================================================
' Thu tuc xac dinh TK Tai chinh doi ung
'===========================================================================
Private Function SetDoiUng(Optional dr As Integer = 0) As Boolean
    Dim sono As Integer, soco As Integer, j As Integer, sodong As Integer, sonox As Integer, socox As Integer
    Dim sodon As Double, chuyen As Boolean, chuyenvt As Boolean, chuyenkh As Boolean    ', st As String
    Dim i As Integer, X As Double, TK As String, TK1 As String, kq As Boolean, id As Long

    kq = False
    sono = 0
    soco = 0
    sonox = 0
    socox = 0
    ' NhËn c¸c dßng ph¸t sinh
    With GrdChungtu
        ReDim mtk(0 To .Rows - 1) As Long
        ReDim mtktc(0 To .Rows - 1) As Long
        ReDim loaips(0 To .Rows - 1) As Integer
        ReDim mvt(0 To .Rows - 1) As Long
        ReDim sops(0 To .Rows - 1) As Double
        ReDim sopsvt(0 To .Rows - 1) As Double
        ReDim SoPS2(0 To .Rows - 1) As Double
        ReDim nb(0 To .Rows - 1) As Integer
        ReDim mkh(0 To .Rows - 1) As Long
        ReDim mkhc(0 To .Rows - 1) As Long
        ReDim mhd(0 To .Rows - 1) As Long
        ReDim mtp(0 To .Rows - 1) As Long
        ReDim dvt(0 To .Rows - 1) As Long
        ReDim cid(0 To .Rows - 1) As Long
        ReDim CK1(0 To .Rows - 1) As Double
        ReDim ck2(0 To .Rows - 1) As Double
        sodong = -1
        For i = 0 To .Rows - 1
            .Row = i
            .col = 8
            If Len(.Text) = 0 Then Exit For
            sodong = sodong + 1
            mtk(i) = CLng5(.Text)
            .col = 0
            .Text = CStr(i + 1)
            .col = 1
            TK = .Text
            .col = 11
            mtktc(i) = CLng5(.Text)
            .col = 12
            cid(i) = CLng5(.Text)
            .col = 9
            mvt(i) = CLng5(.Text)
            .col = 24
            SoPS2(i) = Cdbl5(.Text)
            .col = 25
            If (loaict <> 8 And loaict <> 1) Or mvt(i) = 0 Then .Text = ""
            CK1(i) = Cdbl5(.Text)
            .col = 26
            If ((loaict <> 8 And loaict <> 1) Or mvt(i) = 0) And Left(TK, 4) <> "3331" Then .Text = ""
            ck2(i) = Cdbl5(.Text)

            .col = 17
            mkh(i) = CLng5(.Text)
            .col = 20
            mkhc(i) = CLng5(.Text)
            If Left(TK, 1) = "6" Or Left(TK, 3) = "154" Or Left(TK, 2) = "51" Or Left(TK, 2) = "91" Then
                .col = 21
                mtp(i) = CLng5(.Text)
            Else
                .col = 21
                .Text = ""
            End If
            .col = 3
            If mvt(i) > 0 And Len(.Text) > 0 Then
                If SelectSQL("SELECT DVTVattu.MaSo FROM DVTVattu INNER JOIN Vattu ON DVTVattu.MaVattu=Vattu.MaSo WHERE Vattu.SoHieu='" + .Text + "'") > 0 Then
                    .col = 23
                    dvt(i) = CLng5(.Text)
                End If
                'If Not KtraDVT(mvt(i), dvt(i), st) Then dvt(i) = 0
            End If
            If Left(TK, Len(pVATV)) = pVATV Or Left(TK, 4) = "3331" Or Left(TK, 4) = "3332" Then
                .col = 18
                mhd(i) = CInt5(.Text)
            Else
                mhd(i) = -1
            End If
            .col = 4
            sopsvt(i) = Cdbl5(.Text)
            .col = 6
            sops(i) = Cdbl5(.Text)
            .col = 7
            X = Cdbl5(.Text)
            .col = 10
            nb(i) = CInt5(.Text)
            If sops(i) <> 0 Or (sops(i) = 0 And X = 0 And (Left(TK, Len(pVATV)) = pVATV Or ((Left(TK, 2) = "15") Or (Left(TK, 5) = "33312")) And loaict = 1)) Then
a:
                loaips(i) = -1
                sonox = sonox + 1
                If (nb(i) = 0) And (Not MaDaCo(mtktc(i), -1, mtktc, loaips, sodong - 1)) Then sono = sono + 1
            Else
                .col = 7
                sops(i) = Cdbl5(.Text)
                If sops(i) = 0 And ((Left(TK, Len(pVATV)) = pVATV Or (Left(TK, 2) = "15") And loaict = 1)) Then GoTo a
                loaips(i) = 1
                socox = socox + 1
                If (nb(i) = 0) And (Not MaDaCo(mtktc(i), 1, mtktc, loaips, sodong - 1)) Then soco = soco + 1
            End If
            For j = 13 To 16
                .col = j
                .Text = ""
            Next
        Next
        GoTo DEF
        ' X¸c ®Þnh ®èi øng
        If sono > 1 And soco > 1 Then
ABC:
            Dim shtk As String, tkno As String, TkCo As String, mtcno As Long

            tkno = ""
            TkCo = ""
            nhieunoco = True
            For i = 0 To sodong
                If nb(i) = 0 And loaips(i) = -1 Then
                    shtk = MaSo2SoHieu(mtktc(i), "HethongTK")
                    If InStr(tkno, shtk) = 0 Then tkno = tkno + shtk + ","
                End If
                If nb(i) = 0 And loaips(i) = 1 Then
                    shtk = MaSo2SoHieu(mtktc(i), "HethongTK")
                    If InStr(TkCo, shtk) = 0 Then TkCo = TkCo + shtk + ","
                End If
            Next
            .col = 14
            For i = 0 To sodong
                .Row = i
                If nb(i) = 0 Then
                    .Text = IIf(loaips(i) = -1, TkCo, tkno)
                End If
            Next
        Else
DEF:
            nhieunoco = False
            If sono = 0 Or soco = 0 Then GoTo KT
            If sonox > 1 And socox > 1 Then
                i = IIf(dr = 0, 0, sodong)
                Do While IIf(dr = 0, i <= sodong, i >= 0)
                    If nb(i) = 0 Then
                        .Row = i
                        .col = 1
                        TK = .Text
                        .col = 14
                        If Len(.Text) = 0 Then
                            sodon = 0
                            chuyen = mvt(i) <> 0 Or sopsvt(i) <> 0
                            chuyenvt = True
                            j = IIf(dr = 0, 0, sodong)
                            Do While IIf(dr = 0, j <= sodong, j >= 0)
                                .Row = j
                                If nb(j) = 0 And (loaips(j) <> loaips(i) Or sops(j) = 0) And sops(j) <= sops(i) - sodon And Len(.Text) = 0 Then
                                    If chuyen Then
                                        .col = 16
                                        If chuyenvt Then
                                            .Text = CStr(sopsvt(i))
                                            .col = 24
                                            .Text = CStr(SoPS2(i))
                                        End If
                                        If mvt(i) = 0 Then
                                            chuyen = False
                                        Else
                                            chuyenvt = False
                                        End If
                                        .col = 15
                                        .Text = CStr(mvt(i))
                                        If dvt(i) > 0 Then
                                            .col = 23
                                            .Text = CStr(dvt(i))
                                        End If
                                    End If
                                    .col = 17                                                           ' KH
                                    If CLng5(.Text) = 0 And mkh(i) > 0 Then
                                        .Text = CStr(mkh(i))
                                    Else
                                        If mkh(i) > 0 And mkh(i) <> CLng5(.Text) Then GoTo ABC
                                    End If
                                    .col = 12
                                    If CLng5(.Text) = 0 Then .Text = CStr(cid(i))
                                    .col = 25
                                    If loaict = 8 And mvt(i) > 0 And Cdbl5(.Text) = 0 Then
                                        .Text = Format(CK1(i), Mask_2)
                                        .col = 26
                                        .Text = Format(ck2(i), Mask_0)
                                    End If
                                    .col = 20                                                           ' KH
                                    If CLng5(.Text) = 0 And mkhc(i) > 0 Then
                                        .Text = CStr(mkhc(i))
                                    Else
                                        If mkhc(i) > 0 And mkhc(i) <> CLng5(.Text) Then GoTo ABC
                                    End If
                                    .col = 1
                                    TK1 = .Text
                                    .col = 18
                                    If CInt5(.Text) < 0 And mhd(i) >= 0 And (Left(TK, Len(pVATV)) = pVATV Or Left(TK, 4) = "3331" Or Left(TK, 4) = "3332" Or Left(TK1, 3) = pVATV Or Left(TK1, 4) = "3331" Or Left(TK1, 4) = "3332") Then
                                        .Text = CStr(mhd(i))
                                    Else
                                        If mhd(i) >= 0 And mhd(i) <> CLng5(.Text) And CLng5(.Text) >= 0 Then GoTo ABC
                                    End If
                                    .col = 13
                                    .Text = CStr(mtk(i))
                                    If mtp(i) > 0 Then
                                        .col = 21
                                        If CLng5(.Text) = 0 Then .Text = CStr(mtp(i))
                                    End If
                                    .col = 14
                                    .Text = CStr(mtktc(i))
                                    sodon = sodon + sops(j)
                                    If j > 0 Then
                                        If sodon >= sops(i) And sops(j - 1) <> 0 Then Exit Do
                                    Else
                                        If sodon >= sops(i) Then Exit Do
                                    End If
                                End If
                                j = j + IIf(dr = 0, 1, -1)
                            Loop
                            If sodon < sops(i) And sodon <> 0 Then GoTo ABC
                        End If
                    End If
                    i = i + IIf(dr = 0, 1, -1)
                Loop
            Else
                If sonox = 1 Then
                    i = IIf(dr = 0, 0, sodong)
                    Do While IIf(dr = 0, i <= sodong, i >= 0)
                        .Row = i
                        .col = 1
                        TK = .Text
                        .col = 14
                        If nb(i) = 0 And loaips(i) < 0 Then
                            mtcno = mtktc(i)
                            .Text = ""
                            chuyen = mvt(i) <> 0 Or sopsvt(i) <> 0
                            chuyenvt = True
                            j = IIf(dr = 0, 0, sodong)
                            Do While IIf(dr = 0, j <= sodong, j >= 0)
                                .Row = j
                                If nb(j) = 0 And loaips(j) > 0 Then
                                    If chuyen And (sops(j) <> 0) Then
                                        .col = 16
                                        If chuyenvt Then
                                            .Text = CStr(sopsvt(i))
                                            .col = 24
                                            .Text = CStr(SoPS2(i))
                                        End If
                                        If mvt(i) = 0 Then
                                            chuyen = False
                                        Else
                                            chuyenvt = False
                                        End If
                                        .col = 15
                                        .Text = CStr(mvt(i))
                                        If dvt(i) > 0 Then
                                            .col = 23
                                            .Text = CStr(dvt(i))
                                        End If
                                    End If
                                    .col = 13
                                    .Text = CStr(mtk(i))
                                    .col = 14
                                    .Text = CStr(mtcno)
                                    .col = 17                                                           ' KH
                                    If CLng5(.Text) = 0 And mkh(i) > 0 Then
A2:
                                        .col = 17
                                        .Text = CStr(mkh(i))
                                    Else
                                        If mkh(i) > 0 And mkh(i) <> CLng5(.Text) Then
                                            .col = 1
                                            id = GetTK_ID(.Text, 0)
                                            If id <> TKCNKH_ID And id <> TKCNPT_ID Then GoTo A2
                                            GoTo ABC
                                        End If
                                    End If
                                    .col = 12
                                    If CLng5(.Text) = 0 Then .Text = CStr(cid(i))
                                    .col = 25
                                    If loaict = 8 And mvt(i) > 0 And Cdbl5(.Text) = 0 Then
                                        .Text = Format(CK1(i), Mask_2)
                                        .col = 26
                                        .Text = Format(ck2(i), Mask_0)
                                    End If
                                    .col = 20                                                           ' KH
                                    If CLng5(.Text) = 0 And mkhc(i) > 0 Then
                                        .Text = CStr(mkhc(i))
                                    Else
                                        If mkhc(i) > 0 And mkhc(i) <> CLng5(.Text) Then GoTo ABC
                                    End If
                                    If mtp(i) > 0 Then
                                        .col = 21
                                        If CLng5(.Text) = 0 Then .Text = CStr(mtp(i))
                                    End If
                                    .col = 1
                                    TK1 = .Text
                                    .col = 18
                                    If CInt5(.Text) < 0 And mhd(i) >= 0 And (Left(TK, Len(pVATV)) = pVATV Or Left(TK, 4) = "3331" Or Left(TK, 4) = "3332" Or Left(TK1, Len(pVATV)) = pVATV Or Left(TK1, 4) = "3331" Or Left(TK1, 4) = "3332") Then     ' And (Left(TK, len(PVATV)) = PVATV Or Left(TK, 4) = "3331")
                                        .Text = CStr(mhd(i))
                                    Else
                                        If mhd(i) >= 0 And mhd(i) <> CLng5(.Text) And CLng5(.Text) >= 0 Then GoTo ABC
                                    End If
                                End If
                                j = j + IIf(dr = 0, 1, -1)
                            Loop
                            Exit Do
                        End If
                        i = i + IIf(dr = 0, 1, -1)
                    Loop
                Else
                    i = IIf(dr = 0, 0, sodong)
                    Do While IIf(dr = 0, i <= sodong, i >= 0)
                        .Row = i
                        .col = 1
                        TK = .Text
                        .col = 14
                        If nb(i) = 0 And loaips(i) > 0 Then
                            mtcno = mtktc(i)
                            .Text = ""
                            chuyen = mvt(i) <> 0 Or sopsvt(i) <> 0
                            j = IIf(dr = 0, 0, sodong)
                            Do While IIf(dr = 0, j <= sodong, j >= 0)
                                .Row = j
                                If nb(j) = 0 And loaips(j) < 0 Then
                                    If chuyen And (sops(j) <> 0) Then
                                        .col = 16
                                        .Text = CStr(sopsvt(i))
                                        .col = 24
                                        .Text = CStr(SoPS2(i))
                                        chuyen = False
                                    End If
                                    .col = 15
                                    .Text = CStr(mvt(i))
                                    If dvt(i) > 0 Then
                                        .col = 23
                                        .Text = CStr(dvt(i))
                                    End If
                                    .col = 13
                                    .Text = CStr(mtk(i))
                                    .col = 14
                                    .Text = CStr(mtcno)
                                    If mtp(i) > 0 Then
                                        .col = 21
                                        If CLng5(.Text) = 0 Then .Text = CStr(mtp(i))
                                    End If
                                    .col = 17                                                           ' KH
                                    If CLng5(.Text) = 0 And mkh(i) > 0 Then
A1:
                                        .col = 17
                                        .Text = CStr(mkh(i))
                                    Else
                                        If mkh(i) > 0 And mkh(i) <> CLng5(.Text) Then
                                            .col = 1
                                            id = GetTK_ID(.Text, 0)
                                            If id <> TKCNKH_ID And id <> TKCNPT_ID Then GoTo A1
                                            GoTo ABC
                                        End If
                                    End If
                                    .col = 12
                                    If CLng5(.Text) = 0 Then .Text = CStr(cid(i))
                                    .col = 25
                                    If loaict = 8 And mvt(i) > 0 And Cdbl5(.Text) = 0 Then
                                        .Text = Format(CK1(i), Mask_2)
                                        .col = 26
                                        .Text = Format(ck2(i), Mask_0)
                                    End If
                                    .col = 20                                                           ' KH
                                    If CLng5(.Text) = 0 And mkhc(i) > 0 Then
                                        .Text = CStr(mkhc(i))
                                    Else
                                        If mkhc(i) > 0 And mkhc(i) <> CLng5(.Text) Then GoTo ABC
                                    End If
                                    .col = 1
                                    TK = .Text
                                    .col = 18
                                    If CInt5(.Text) < 0 And mhd(i) >= 0 And (Left(TK, Len(pVATV)) = pVATV Or Left(TK, 4) = "3331" Or Left(TK, 4) = "3332" Or Left(TK1, Len(pVATV)) = pVATV Or Left(TK1, 4) = "3331" Or Left(TK1, 4) = "3332") Then
                                        .Text = CStr(mhd(i))
                                    Else
                                        If mhd(i) >= 0 And mhd(i) <> CLng5(.Text) And CLng5(.Text) >= 0 Then GoTo ABC
                                    End If
                                End If
                                j = j + IIf(dr = 0, 1, -1)
                            Loop
                            Exit Do
                        End If
                        i = i + IIf(dr = 0, 1, -1)
                    Loop
                End If
            End If
            kq = True
        End If
KT:
        Erase mtk
        Erase mtktc
        Erase loaips
        Erase mvt
        Erase sopsvt
        Erase nb
        Erase sops
        Erase SoPS2
        Erase mkh
        Erase mkhc
        Erase mhd
        Erase mtp
        Erase dvt
        Erase cid
        Erase CK1
        Erase ck2
    End With
    SetDoiUng = kq
End Function
'=====================================================================================================
' Hµm kiÓm tra m· sè ®· cã trong cét cña Grid
'=====================================================================================================
Private Function MaDaCo(MaSo As Long, ps As Integer, BangMa() As Long, loaips() As Integer, sodong As Integer) As Boolean
    Dim i As Integer

    For i = 0 To sodong
        If BangMa(i) = MaSo And ps = loaips(i) Then
            MaDaCo = True
            Exit Function
        End If
    Next
    MaDaCo = False
End Function

Private Sub hienctts()
    Dim i As Integer

    KhongNhapTS = False
    For i = 0 To parSoPS
        If arPhatSinh(i).PS_SoLg <> 0 Then
            txtchungtu(0).Text = arPhatSinh(i).TK_SoHieu
            txtChungtu_LostFocus 0

            If Len(arPhatSinh(i).TS_SoHieu) > 0 Then
                txtchungtu(2).Text = arPhatSinh(i).TS_SoHieu
                txtchungtu(1).Text = TenTS(arPhatSinh(i).TS_SoHieu, 0)
            End If

            If pDTTP <> 0 And Len(arPhatSinh(i).ShTP) > 0 Then
                txtchungtu(2).Text = arPhatSinh(i).ShTP
                txtChungtu_LostFocus 2
            End If
            If arPhatSinh(i).PS_Loai = -1 Then
                txtchungtu(5).Text = Format(arPhatSinh(i).PS_SoLg, Mask_2)
                txtchungtu(6).Text = "0"
            Else
                txtchungtu(6).Text = Format(arPhatSinh(i).PS_SoLg, Mask_2)
                txtchungtu(5).Text = "0"
            End If
            If Len(Trim(txtchungtu(1).Text)) > 0 Then CmdChitiet_chon
        End If
    Next
    KhongNhapTS = True
End Sub

Private Sub GhiChungtuTS(MaCTKT As Long)
    Dim sql As String, i As Integer
    Select Case pNghiepVu
    Case NV_TANG:
        For i = 0 To tscount
            TinhGiaTriTaiSan MaTS(i), pThangTacDong, KH_KHONG
            sql = "INSERT INTO CTTaiSan (MaSo, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, " _
                & "MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, " _
                & "CL_NS, CL_TBS, CL_CNK, CL_TD, MaCTKT" + IIf(pSongNgu And Len(txt(2).Text) > 0, ",DienGiaiE", "") + ") VALUES (" + CStr(Lng_MaxValue("MaSo", "CTTaiSan") + 1) + ",'" + txt(0).Text + "'," + CStr(CboThang.ItemData(CboThang.ListIndex)) _
                + ",#" + Format(ngay(0), Mask_DB) + "#,#" + Format(ngay(1), Mask_DB) + "#,'" _
                + txt(1).Text + "'," + CStr(OptLoai(loaict).tag) + "," + CStr(CboNguon(0).ItemData(CboNguon(0).ListIndex)) + "," + CStr(MaTS(i)) + "," _
                + DoiDau(GiaTri.NG_NS) + "," + DoiDau(GiaTri.NG_TBS) + "," + DoiDau(GiaTri.NG_CNK) + "," + DoiDau(GiaTri.NG_TD) + "," _
                + DoiDau(GiaTri.CL_NS) + "," + DoiDau(GiaTri.CL_TBS) + "," + DoiDau(GiaTri.CL_CNK) + "," + DoiDau(GiaTri.CL_TD) + "," + CStr(MaCTKT) + IIf(pSongNgu And Len(txt(2).Text) > 0, ",'" + txt(2).Text + "'", "") + ")"
            ExecuteSQL5 (sql)
            sql = "UPDATE TaiSan SET SHCT='" + txt(0).Text + "',NCT=#" + Format(ngay(0), Mask_DB) + "# WHERE MaSo=" + CStr(MaTS(i))
            ExecuteSQL5 (sql)
        Next
    Case NV_GIAM:
        TacDongGiamTaiSan pMaTaiSan, CboThang.ItemData(CboThang.ListIndex), TD_GIAM
    Case NV_DGLAI:
    Case NV_TKHAO:
        XoaChungTuKhauHao CInt5(Me.tag), CboThang.ItemData(CboThang.ListIndex), CboNguon(0).ItemData(CboNguon(0).ListIndex), MaCTKT, CboNguon(0).tag
    End Select
    If pNghiepVu > 0 And pNghiepVu <> NV_TANG Then
        sql = "INSERT INTO CTTaiSan (MaSo, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, " _
            & "MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, " _
            & "CL_NS, CL_TBS, CL_CNK, CL_TD, MaCTKT" + IIf(pSongNgu And Len(txt(2).Text) > 0, ",DienGiaiE", "") + ")VALUES (" + CStr(Lng_MaxValue("MaSo", "CTTaiSan") + 1) + ",'" + txt(0).Text + "'," + CStr(CboThang.ItemData(CboThang.ListIndex)) _
            + ",#" + Format(ngay(0), Mask_DB) + "#,#" + Format(ngay(1), Mask_DB) + "#,'" _
            + txt(1).Text + "'," + CStr(OptLoai(loaict).tag) + "," + CStr(CboNguon(0).ItemData(CboNguon(0).ListIndex)) + "," + CStr(pMaTaiSan) + "," _
            + DoiDau(GiaTri.NG_NS) + "," + DoiDau(GiaTri.NG_TBS) + "," + DoiDau(GiaTri.NG_CNK) + "," + DoiDau(GiaTri.NG_TD) + "," _
            + DoiDau(GiaTri.CL_NS) + "," + DoiDau(GiaTri.CL_TBS) + "," + DoiDau(GiaTri.CL_CNK) + "," + DoiDau(GiaTri.CL_TD) + "," + CStr(MaCTKT) + IIf(pSongNgu And Len(txt(2).Text) > 0, ",'" + txt(2).Text + "'", "") + ")"
        ExecuteSQL5 (sql)
        If pNghiepVu = NV_DGLAI Then DieuChinhKH pMaTaiSan, CboThang.ItemData(CboThang.ListIndex)
    End If
    pGhichungtu = 0
    pMaTaiSan = 0
    tscount = -1
    OptLoai(0).Value = True
    SetLoaiChungtu 0
End Sub

Private Sub SuaChungtuTS(MaCTKT As Long)
    Dim sql As String
    sql = "UPDATE CTTaiSan SET SoHieu = '" + txt(0).Text + "', Thang = " + CStr(CboThang.ItemData(CboThang.ListIndex)) + ", VaoSo = #" + Format(ngay(0), Mask_DB) _
        + "#, NgayGhi = #" + Format(ngay(1), Mask_DB) + "#, DienGiai = '" + txt(1).Text + "', MaLoai = " + CStr(OptLoai(loaict).tag) _
        + ", MaNhom = " + CStr(CboNguon(0).ItemData(CboNguon(0).ListIndex)) + IIf(pSongNgu, ",DienGiaiE='" + txt(2).Text + "'", "") + " WHERE CTTaiSan.MaCTKT = " + CStr(MaCTKT)
    ExecuteSQL5 (sql)
End Sub
'=====================================================================================================
' Ham tra ve dong va loai ps cua tai khoan trong grid chung tu
'=====================================================================================================
Private Function PSDaCo(taikhoan As ClsTaikhoan, loaips As Integer, sops As Double, sopsnt As Double, mkh As Long) As Boolean
    Dim i As Integer, ps As Double, tien As Double

    PSDaCo = False
    With GrdChungtu
        For i = 0 To .Rows - 1
            .col = 8
            .Row = i
            If Len(.Text) = 0 Then Exit For
            If CLng5(.Text) = taikhoan.MaSo Then
                If mkh > 0 Then
                    .col = IIf(loaips = -1, 17, 20)
                    If CLng5(.Text) <> mkh Then GoTo a
                    .col = IIf(loaips = -1, 6, 7)
                    If Cdbl5(.Text) = 0 Then GoTo a
                End If
                .col = IIf(loaips = -1, 6, 7)
                ps = sops + Cdbl5(.Text)
                .Text = Format(ps, Mask_0)
                If taikhoan.MaNT > 0 Or KHMaNT(mkh) > 0 Then       ' Or KHMaNT(mkh) > 0
                    .col = 4
                    ps = sopsnt
                    ps = ps + Cdbl5(.Text)
                    .Text = IIf(ps <> 0, Format(ps, Mask_2), "")
                    If ps <> 0 Then
                        .col = 6
                        tien = Cdbl5(.Text)
                        If tien = 0 Then
                            .col = 7
                            tien = Cdbl5(.Text)
                        End If
                        .col = 5
                        .Text = Format(tien / ps, Mask_2)
                    End If
                End If
                PSDaCo = True
            End If
a:
        Next
    End With
End Function
' hien ra de in phieu thu chi
Public Sub mnDD_Click(Index As Integer)
    Select Case Index
    Case 0:    ' PLVT
        frmMain.mnVT_Click 0
    Case 1:    ' PLTS
        frmMain.mnTS_Click 0
    Case 2:    ' PLTS
        If KHDetail Then frmMain.mnCn_Click 0
    Case 4:    ' DDTK
        FrmTaikhoan.tag = 1
        FrmTaikhoan.Show 1
    Case 5:    ' DDVT
        If STDetail Then FrmVattu.Show 1
    Case 6:    ' DDTS
        If KHDetail Then
            pNghiepVu = NV_KHONG
            frmDSTaiSan.Show 1
        End If
    Case 7:    ' DDTS
        If FADetail Then FrmKhachHang.Show vbModal
    Case 9:    ' SDCT
        frmMain.mnVT_Click 1
        If loaict = 1 Or loaict = 2 Then Int_RecsetToCbo "SELECT MaSo As F2,SoHieu + ' - ' + DienGiai As F1 FROM NguonNhapXuat ORDER BY SoHieu", CboNguon(0)
    Case 10:
        'Load FrmKho
        FrmKho.tag = 5
        FrmKho.Show 1
        Int_RecsetToCbo "SELECT DoituongCT.MaSo As F2,(DienGiai+IIF(DoituongCT.MaKhachHang>0,' - HÑ: '+DoituongCT.SoHieu+' - Kyù ngaøy: '+ Format(NgayKy,'dd/mm/yy'),'')) As F1 FROM DoituongCT LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo ORDER BY DienGiai", CboNguon(2)
    Case 11:    ' SDCT
        FrmHD.Show vbModal
        Int_RecsetToCbo "SELECT DoituongCT.MaSo As F2,(DienGiai+IIF(DoituongCT.MaKhachHang>0,' - HÑ: '+DoituongCT.SoHieu+' - Kyù ngaøy: '+ Format(NgayKy,'dd/mm/yy'),'')) As F1 FROM DoituongCT LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo ORDER BY DienGiai", CboNguon(2)
    Case 13:    ' BCCT
        If User_Right = 3 Then Exit Sub
        FBcKt.Show 1
    Case 14:    ' BCTH
        If User_Right = 3 Then Exit Sub
        FBcTC.Show 1
    Case 16:    ' BCTH
        FrmDU.Show 1
        Set FrmDU = Nothing
    Case 17:
        InDSCtu
    Case 18:
        InNhatKy1
    Case 20, 21:
        InTC Index - 20
    Case 100, 101:
        Dim kk
        If Index = 100 Then
            InTC_in_toan_bo 0
        Else
            InTC_in_toan_bo 1
        End If
    Case 22, 23:
        InNX Index - 21
    Case 25:
        On Error Resume Next
        txtChungtu_LostFocus 2
        On Error GoTo 0
        DonGiaNhap vattu.MaSo
    Case 27, 28, 29:
        FrmKho.tag = Index - 17
        FrmKho.Show 1
        Int_RecsetToCbo "SELECT MaSo As F2,DienGiai As F1 FROM DoituongCT" + CStr(Index - 26) + " ORDER BY DoituongCT" + CStr(Index - 26) + ".DienGiai", CboVV(Index - 27)
    End Select
    Me.MousePointer = 0
End Sub

Private Sub KiemTraUser()
    OptLoai(3).Enabled = (frmMain.tag Mod 10 >= 1)
    mnDD(13).Enabled = (User_Right <> 1)
    mnDD(14).Enabled = (frmMain.tag Mod 10 >= 1) And (User_Right <> 1)
    mnDD(18).Enabled = (frmMain.tag Mod 10 >= 1)
    If Not (frmMain.tag Mod 10 >= 1) Then
        OptLoai(1).Enabled = (frmMain.tag Mod 100 >= 10)
        OptLoai(2).Enabled = (frmMain.tag Mod 1000 >= 100)

        OptLoai(9).Enabled = (frmMain.tag Mod 10000 >= 1000)
        OptLoai(10).Enabled = (frmMain.tag Mod 10000 >= 1000)
        OptLoai(11).Enabled = (frmMain.tag Mod 10000 >= 1000)
        OptLoai(12).Enabled = (frmMain.tag Mod 10000 >= 1000)

        OptLoai(8).Enabled = (frmMain.tag Mod 100000 >= 10000)

        CmdDanhSach(1).Enabled = (frmMain.tag Mod 1000000 >= 100000)

        mnDD(0).Enabled = (frmMain.tag Mod 100 >= 10 Or frmMain.tag Mod 1000 >= 100)
        mnDD(5).Enabled = (frmMain.tag Mod 100 >= 10 Or frmMain.tag Mod 1000 >= 100)
        mnDD(9).Enabled = (frmMain.tag Mod 100 >= 10 Or frmMain.tag Mod 1000 >= 100)
        mnDD(22).Enabled = (frmMain.tag Mod 100 >= 10)
        mnDD(23).Enabled = (frmMain.tag Mod 1000 >= 100)
        mnDD(25).Enabled = (frmMain.tag Mod 100 >= 10)

        mnDD(1).Enabled = (frmMain.tag Mod 10000 >= 1000)
        mnDD(6).Enabled = (frmMain.tag Mod 10000 >= 1000)

        mnDD(2).Enabled = (frmMain.tag Mod 100000 >= 10000)
        mnDD(7).Enabled = (frmMain.tag Mod 100000 >= 10000)
        mnDD(11).Enabled = (frmMain.tag Mod 100000 >= 10000)

        mnDD(20).Enabled = (frmMain.tag Mod 1000000 >= 100000)
        mnDD(21).Enabled = (frmMain.tag Mod 1000000 >= 100000)
    End If
End Sub

Private Function PSTuDong(ps As Double) As Boolean
    Dim sql As String, sh As String

    PSTuDong = False
    If ps > 0 Then
        txtchungtu(5).Text = 0
        txtchungtu(6).Text = Format(ps, Mask_2)
    Else
        txtchungtu(6).Text = 0
        txtchungtu(5).Text = Format(-ps, Mask_2)
    End If

    sql = "SELECT SHTK As F1 FROM SHChungTu WHERE SoHieu='" + Left(txt(0).Text, SHCT_Len) + "'"
    sh = SelectSQL(sql)
    If Len(sh) < 3 Then Exit Function
    txtchungtu(0).Text = sh
    txtChungtu_LostFocus 0
    If Len(Trim(txtchungtu(1).Text)) > 0 Then CmdChitiet_chon
    PSTuDong = (SoPSConLai = 0)
End Function

Private Sub InDSCtu()
    Dim d1 As Date, d2 As Date, sql As String

    If Not GetDate2.GetDate("In danh s¸ch chøng tõ theo ngµy", d1, d2) Then Exit Sub
    sql = "SELECT ChungTu.MaCT AS M, ChungTu.ThangCT AS T, ChungTu.SoHieu AS SH, ChungTu.NgayCT AS NCT, ChungTu.NgayGS AS NGS, ChungTu.DienGiai AS DG, HeThongTK.SoHieu AS TKNo,HethongTK.Ten AS TNo, HeThongTK_1.SoHieu AS TKCo,HethongTK_1.Ten AS TCo, ChungTu.SoPS AS PS" _
        & " FROM (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) LEFT JOIN HeThongTK AS HeThongTK_1 ON ChungTu.MaTKCo = HeThongTK_1.MaSo WHERE " + WNgay("NgayGS", d1, d2)
    SetSQL "QNhatKy", sql
    frmMain.Rpt.ReportFileName = "CHUNGTU2.RPT"
    RptSetDate d2
    SetRptInfo
    InBaoCaoRPT
End Sub

Private Sub InNhatKy1()
    Dim d1 As Date, d2 As Date

    If Not GetDate2.GetDate("Sæ nhËt ký theo ngµy", d1, d2) Then Exit Sub
    SetRptInfo
    If InNhatKy(0, 0, 0, 1, d1, d2, 0, pPhieu) Then InBaoCaoRPT
End Sub

Private Function LoaiPhieuThuChi(sh As String) As Integer
    Dim i As Integer

    LoaiPhieuThuChi = 0
    With GrdChungtu
        For i = 0 To .Rows - 1
            .col = 1
            .Row = i
            If Len(.Text) = 0 Then Exit Function
            If Left(.Text, Len(sh)) = sh Then
                .col = 6
                If LoaiPhieuThuChi > -1 Then LoaiPhieuThuChi = IIf(Cdbl5(.Text) > 0, -1, 1)
            End If
        Next
    End With
End Function

Private Function GiaTriTruocThue() As Double
    Dim tien As Double, i As Integer

    With GrdChungtu
        tien = 0
        If (taikhoan.tk_id = GTGTKT_ID) Then
            For i = 0 To .Rows - 1
                .Row = i
                .col = 1
                If Len(.Text) = 0 Then Exit For
                If Left(.Text, Len(pVATV)) = pVATV Then Exit For
                .col = 6
                tien = tien + Cdbl5(.Text)
            Next
        End If
        If (taikhoan.tk_id = GTGTPN_ID Or taikhoan.tk_id = TTDB_ID Or Left(taikhoan.sohieu, 3) = "521") Then
            For i = 0 To .Rows - 1
                .Row = i
                .col = 1
                If Len(.Text) = 0 Then Exit For
                If Left(.Text, 4) = "3331" And taikhoan.tk_id = GTGTPN_ID Then Exit For
                If Left(.Text, 2) <> "11" And Left(.Text, 4) <> "3331" Then
                    .col = IIf(taikhoan.tk_id = TTDB_ID And loaict = 1, 6, 7)
                    tien = tien + Cdbl5(.Text)
                End If
            Next
        End If
    End With
    GiaTriTruocThue = tien
End Function

Private Function LayMaKH(loaips As Integer) As Long
    Dim i As Integer

    With GrdChungtu
        For i = 0 To .Rows - 1
            .Row = i
            .col = 1
            If Len(.Text) = 0 Then Exit Function
            .col = IIf(loaips < 0, 17, 20)
            LayMaKH = CLng5(.Text)
            If LayMaKH > 0 Then Exit Function
        Next
    End With
End Function

Private Function CTGiamGia() As Boolean
    CTGiamGia = CoPSTK("521") Or CoPSTK("531") Or CoPSTK("532")
End Function

Private Sub InTC(loai As Integer)
    Dim d1 As Date, d2 As Date, sql As String
    Dim rs As Recordset

    If Not GetDate2.GetDate("In tõng phiÕu " + IIf(loai = 0, "thu", "chi") + " theo ngµy", d1, d2) Then Exit Sub
    sql = "SELECT MaCT FROM HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTK" + IIf(loai = 0, "N", "C") + "o WHERE " + WNgay("NgayGS", d1, d2) + " AND HethongTK.SoHieu LIKE '1111*' GROUP BY MaCT"
    frmMain.Rpt.Destination = crptToPrinter
    P_1 = 1
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        MaSoCT = rs!MaCT
        HienPhieuTrenManHinh 0
        CmdPhieu_Click 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Command_Click 0
    frmMain.Rpt.Destination = crptToWindow
    P_1 = 0
End Sub

Private Sub InTC_in_toan_bo(loai As Integer)
    Dim d1 As Date, d2 As Date, sql As String
    Dim rs As Recordset
    Dim dieukien As String
    d1 = GetDate2.MedNgay(0).Text
    d2 = GetDate2.MedNgay(1).Text

    ' If Not GetDate2.GetDate("In tõng phiÕu " + IIf(loai = 0, "thu", "chi") + " theo ngµy", d1, d2) Then Exit Sub
    ' dieukien = FrmDsCT
    sql = "SELECT MaCT FROM HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTK" + IIf(loai = 0, "N", "C") + "o WHERE " + WNgay("NgayGS", d1, d2) + " AND HethongTK.SoHieu LIKE '1111*' GROUP BY MaCT"
    frmMain.Rpt.Destination = crptToPrinter
    P_1 = 1
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        MaSoCT = rs!MaCT
        HienPhieuTrenManHinh 0
        CmdPhieu_Click 0
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Command_Click 0
    frmMain.Rpt.Destination = crptToWindow
    P_1 = 0
End Sub

Private Sub InNX(loai As Integer)
    Dim d1 As Date, d2 As Date, sql As String
    Dim rs As Recordset

    If Not GetDate2.GetDate("In tõng phiÕu " + IIf(loai = 1, "nhËp", "xuÊt") + " theo ngµy", d1, d2) Then Exit Sub
    sql = "SELECT MaCT FROM ChungTu WHERE MaLoai=" + CStr(loai) + " AND " + WNgay("NgayGS", d1, d2) + " AND MaVattu>0 GROUP BY MaCT"
    frmMain.Rpt.Destination = crptToPrinter
    P_1 = 1
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        MaSoCT = rs!MaCT
        HienPhieuTrenManHinh 0
        CmdPhieu_Click 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Command_Click 0
    frmMain.Rpt.Destination = crptToWindow
    P_1 = 0
End Sub

Private Sub DonGiaNhap(mvt As Long)
    Dim sql As String

    SetSQL "MienTru", "SELECT NgayCT,MaVattu,SoPS,SoPS2No FROM ChungTu WHERE ChungTu.MaLoai=1 AND SoPS>0 AND SoPS2No>0 AND MaVattu" + IIf(mvt > 0, "=" + CStr(mvt), ">0") + " ORDER BY NgayCT DESC"
    sql = "SELECT SoHieu,DonVi,TenVattu,First(NgayCT) AS Ngay,First(SoPS/SoPS2No) AS DonGia FROM MienTru INNER JOIN Vattu On MienTru.MaVattu=Vattu.MaSo GROUP BY SoHieu,DonVi,TenVattu"
    SetSQL "QLuyKe", sql

    SetRptInfo

    frmMain.Rpt.ReportFileName = "DONGIA.RPT"

    frmMain.Rpt.WindowTitle = "B¶ng ®¬n gi¸ nhËp kho míi nhÊt"
    frmMain.Rpt.Destination = crptToWindow
    Me.MousePointer = 0
    InBaoCaoRPT
End Sub

Private Sub LayGiaBan()
    If loaict <> 8 Then Exit Sub
    CboNT(1).Clear
    With vattu
        If .MaSo = 0 Then GoTo KT
        If loaict = 8 And CDbl(txttinh_gia_ban.Caption) > 0 Then CboNT(1).AddItem Format(txttinh_gia_ban.Caption)
        If .GiaBan1 > 0 Then CboNT(1).AddItem Format(.GiaBan1, Mask_2)
        If .GiaBan2 > 0 Then CboNT(1).AddItem Format(.GiaBan2, Mask_2)
        If .GiaBan3 > 0 Then CboNT(1).AddItem Format(.GiaBan3, Mask_2)
    End With
KT:
    CboNT(1).Visible = (CboNT(1).ListCount > 0)
    If CboNT(1).ListCount > 0 Then CboNT(1).ListIndex = 0
End Sub

Private Function CoPSTK(shtk As String, Optional loaips As Integer = 0, Optional tien As Double) As Boolean
    Dim i As Integer

    CoPSTK = False
    tien = 0
    With GrdChungtu
        For i = 0 To .Rows - 1
            .Row = i
            .col = 1
            If Len(.Text) = 0 Then Exit For
            If Left(.Text, Len(shtk)) = shtk Then
                Select Case loaips
                Case -1:
                    GrdChungtu.col = 6
                    tien = tien + Cdbl5(GrdChungtu.Text)
                    If tien <> 0 Then CoPSTK = True
                Case 0:
                    CoPSTK = True
                    Exit For
                Case 1:
                    GrdChungtu.col = 7
                    tien = tien + Cdbl5(GrdChungtu.Text)
                    If tien <> 0 Then CoPSTK = True
                End Select
            End If
        Next
    End With

End Function

Private Sub txtsh_GotFocus(Index As Integer)
    AutoSelect txtsh(Index)
End Sub

Private Sub txtsh_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmd_Click Index
End Sub

Private Sub txtsh_LostFocus(Index As Integer)
    Dim vis As Boolean, tkxt As ClsTaikhoan, khxt As ClsKhachHang, tpxt As Cls154

    Select Case Index
    Case 0:
        Set tkxt = New ClsTaikhoan
        tkxt.InitTaikhoanSohieu txtsh(0).Text
        txtsh(0).tag = IIf(tkxt.MaSo > 0 And tkxt.tkcon = 0, tkxt.MaSo, 0)
        lb(0).Caption = tkxt.Ten
        vis = (tkxt.tk_id = TKCNKH_ID Or tkxt.tk_id = TKCNPT_ID Or (tkxt.loai = 6 And pDTTP <> 0))
        If Left(txtsh(0).Text, 3) = "154" Then
            vis = True
        End If

        Label(19).Enabled = vis
        txtsh(1).Enabled = vis
        lb(1).Enabled = vis
        cmd(1).Enabled = vis
        cmd(0).tag = IIf(tkxt.tk_id = TKCNKH_ID Or tkxt.tk_id = TKCNPT_ID, 1, IIf(tkxt.loai = 6 And pDTTP <> 0, 2, 0))
        Set tkxt = Nothing
    Case 1:
        If cmd(0).tag = 1 Then
            Set khxt = New ClsKhachHang
            khxt.InitKhachHangSohieu txtsh(1).Text
            txtsh(1).tag = khxt.MaSo
            lb(1).Caption = khxt.Ten
            Set khxt = Nothing
        End If
        If cmd(0).tag = 2 Then
            Set tpxt = New Cls154
            tpxt.InitTPSohieu txtsh(1).Text
            txtsh(1).tag = tpxt.MaSo
            lb(1).Caption = tpxt.TenVattu
            Set tpxt = Nothing
        End If

        If Left(txtsh(0).Text, 3) = "154" Then
            Set tpxt = New Cls154
            tpxt.InitTPSohieu txtsh(1).Text
            txtsh(1).tag = tpxt.MaSo
            lb(1).Caption = tpxt.TenVattu
            Set tpxt = Nothing
        End If

    End Select
End Sub

Public Sub VaoSoNK(mct As Long)
    Dim xphieu As Integer, loaip As Integer

    pGhi = 1
    xphieu = pPhieu
    pPhieu = 0
    MaSoCT = mct
    HienPhieuTrenManHinh 1
    CmdDanhSach(0).Enabled = False
    CmdDanhSach(1).Enabled = False
    Me.Show 1
    If MaSoCT = 0 Then XoaPhieu mct
    pPhieu = xphieu
    pGhi = 0
End Sub

Private Function SoChietKhau() As Double
    Dim i As Integer, CK As Double

    With GrdChungtu
        For i = 0 To .Rows - 1
            .Row = i
            .col = 1
            If Len(.Text) = 0 Then Exit For
            If Left(.Text, 3) = "511" Then
                .col = 26
                CK = CK + Cdbl5(.Text)
            End If
        Next
    End With
    SoChietKhau = CK
End Function

Private Function SoPSTKCT(shtk As String, loaips As Integer) As Double
    Dim i As Integer

    With GrdChungtu
        For i = 0 To .Rows - 1
            .Row = i
            .col = 1
            If Len(.Text) = 0 Then Exit For
            If Left(.Text, Len(shtk)) = shtk Then
                Select Case loaips
                Case -1:
                    GrdChungtu.col = 6
                    SoPSTKCT = SoPSTKCT + Cdbl5(GrdChungtu.Text)
                Case 1:
                    GrdChungtu.col = 7
                    SoPSTKCT = SoPSTKCT + Cdbl5(GrdChungtu.Text)
                End Select
            End If
        Next
    End With
End Function

Private Sub TinhCKCT()
    Dim CK As Double
    If OptLoai(8).Value = True Then
        CK = RoundMoney(Cdbl5(txtchungtu(6).Text) * Cdbl5(txtchungtu(9).Text) / 100)
    ElseIf OptLoai(1).Value = True Then
        CK = RoundMoney(Cdbl5(txtchungtu(5).Text) * Cdbl5(txtchungtu(9).Text) / 100)
        txtchungtu(5).Text = CStr(CDbl(txtchungtu(5).Text) - CK)
    End If

    txtchungtu(10).Text = Format(CK, Mask_0)
End Sub

Private Function SoLanPSTK(shtk As String) As Integer
    Dim i As Integer, k As Integer

    With GrdChungtu
        For i = 0 To .Rows - 1
            .Row = i
            .col = 1
            If Len(.Text) = 0 Then Exit For
            If Left(.Text, Len(shtk)) = shtk Then k = k + 1
        Next
    End With
    SoLanPSTK = k
End Function


Private Sub HienThiHachToan()
    Dim i As Integer, stt1 As Integer, stt2 As Integer, shtk As String, tien As Double, j As Integer, daco As Boolean
    ReDim tkn(1 To 1) As String
    ReDim psn(1 To 1) As Double
    ReDim tkc(1 To 1) As String
    ReDim psc(1 To 1) As Double

    With GrdChungtu
        For i = 0 To .Rows - 1
            .Row = i
            .col = 1
            shtk = .Text
            If Len(shtk) = 0 Then GoTo KT
            .col = 6
            tien = Cdbl5(.Text)
            If tien <> 0 Then
                daco = False
                For j = 1 To stt1
                    If tkn(j) = shtk Then
                        psn(j) = psn(j) + tien
                        daco = True
                        Exit For
                    End If
                Next
                If Not daco Then
                    stt1 = stt1 + 1
                    ReDim Preserve tkn(1 To stt1) As String
                    ReDim Preserve psn(1 To stt1) As Double
                    tkn(stt1) = shtk
                    psn(stt1) = tien
                End If
            End If
            .col = 7
            tien = Cdbl5(.Text)
            If tien <> 0 Then
                daco = False
                For j = 1 To stt2
                    If tkc(j) = shtk Then
                        psc(j) = psc(j) + tien
                        daco = True
                        Exit For
                    End If
                Next
                If Not daco Then
                    stt2 = stt2 + 1
                    ReDim Preserve tkc(1 To stt2) As String
                    ReDim Preserve psc(1 To stt2) As Double
                    tkc(stt2) = shtk
                    psc(stt2) = tien
                End If
            End If
        Next
KT:
        For i = 1 To stt1
            frmMain.Rpt.Formulas(70 + i) = "TKN" + CStr(i) + " = '" + tkn(i) + "'"
            frmMain.Rpt.Formulas(80 + i) = "PSN" + CStr(i) + " = " + DoiDau(psn(i))
        Next
        For i = 1 To stt2
            frmMain.Rpt.Formulas(90 + i) = "TKC" + CStr(i) + " = '" + tkc(i) + "'"
            frmMain.Rpt.Formulas(100 + i) = "PSC" + CStr(i) + " = " + DoiDau(psc(i))
        Next
    End With

    Erase tkn
    Erase tkc
    Erase psn
    Erase psc
End Sub

Private Function LaySohieuDoiTuong(dg As String) As Integer
    Dim i As Integer, k As Integer, id As Long

    dg = ""
    With GrdChungtu
        For i = 0 To .Rows - 1
            .Row = i
            .col = 1
            If Len(.Text) = 0 Then Exit For
            id = GetTK_ID(.Text, 0)
            If id = TKCNKH_ID Or id = TKCNPT_ID Then
                k = k + 1
                .col = 3
                If Len(.Text) > 0 Then dg = dg + " - " + .Text
            End If
        Next
        If k > 0 Then dg = Right(dg, Len(dg) - 3)
        LaySohieuDoiTuong = k
    End With

End Function

Private Sub LaySohieuDoiTuong2(loaidt As Integer, sh As String)
    Select Case loaidt
    Case 1:
        Int_RecsetToCbo "SELECT DISTINCTROW MaSo AS F2, SoHieu + ' - ' + TenVattu AS F1 FROM Vattu WHERE SoHieu LIKE '" + sh + "*' ORDER BY SoHieu", CboNT(3), 1
    Case 2:
        Int_RecsetToCbo "SELECT DISTINCTROW MaSo AS F2, SoHieu + ' - ' + Ten AS F1 FROM KhachHang WHERE SoHieu LIKE '" + sh + "*' AND LEFT(SoHieu,1)<>'#' ORDER BY SoHieu", CboNT(3), 1
    End Select
    Me.Refresh
    CboNT(3).tag = loaidt
    RFocus CboNT(3)
    SendKeys "{F4}"
End Sub

Private Sub LayXuatKho(ml As Integer)
    Dim id As Double
    Dim FileNum As Integer
    Dim BytesNeeded As Long
    Dim Buffers As Long
    Dim i As Long, st As String, j As Integer, shtk As String, st2 As String, luong As Double, tien As Double, ms As Long, mtk As Long, mv As Long, sl As Double, T As Double
    Dim Buffer(32) As Byte

    If Len(Dir(pCurDir + "DOWNLOAD.EXE")) = 0 Then Exit Sub
    ChDir Left(pCurDir, Len(pCurDir) - 1)
    Recycle pCurDir + "BARCODE.FIL"
    On Error GoTo E
    id = Shell(pCurDir + "DOWNLOAD.EXE", vbMaximizedFocus)
    On Error GoTo 0
    Do While bcstop = 0
        AppIdle 1000
        i = i + 1
        If Len(Dir(pCurDir + "BARCODE.FIL")) > 0 Or i > 10000 Then Exit Do
        'If OpenProcess(PROCESS_ALL_ACCESS, 0&, id) = 0 Then Exit Do
    Loop
    bcstop = 0
    If Len(Dir(pCurDir + "BARCODE.FIL")) = 0 Then GoTo KT

    Me.MousePointer = 11
    XoaPhieuTrenManHinh

    FileNum = FreeFile
    Open pCurDir + "BARCODE.FIL" For Binary As #FileNum
    BytesNeeded = LOF(FileNum)

    Buffers = BytesNeeded \ 32
    For i = 0 To Buffers - 1
        Get #FileNum, , Buffer
        If i = 0 Then
            st = ""
            For j = 0 To 9
                st = st + Chr(Buffer(j))
            Next
            If IsDate(st) Then
                ngay(0) = CVDate(st)
                ngay(1) = ngay(0)
                MedNgay(0).Text = Format(ngay(0), Mask_D)
                MedNgay(1).Text = Format(ngay(1), Mask_D)
            End If
        End If

        st = ""
        j = 10
        Do While Buffer(j) <> 32 And j < 32
            'If Buffer(j) <> 42 Or pBCode <> 39 Then st = st + Chr(Buffer(j))
            If Buffer(j) <> 42 Then st = st + Chr(Buffer(j))
            j = j + 1
        Loop
        'If pBCode <> 39 Then
        '    If Not IsNumeric(st) Then GoTo E1
        'End If

        st2 = ""
        j = 26
        Do While Buffer(j) <> 32 And j < 32
            st2 = st2 + Chr(Buffer(j))
            j = j + 1
        Loop
        luong = Cdbl5(st2)

        If pBarCode = 1 Then
            vattu.InitVattuSohieu st

            If vattu.MaSo > 0 Then
                If Len(shtk) = 0 Then
                    If ml = 2 Then
                        shtk = SelectSQL("SELECT HethongTK.Sohieu AS F1 FROM TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo WHERE MaVattu=" + CStr(vattu.MaSo) + " AND MaSoKho=" + CStr(CboNguon(1).ItemData(CboNguon(1).ListIndex)))
                        Do While SoHieu2MaSo(shtk, "HethongTK") = 0
                            shtk = FrmGetStr.GetString("Sè hiÖu tµi kho¶n ghi cã:", "PhiÕu xuÊt kho")
                            If SelectSQL("SELECT TK_ID AS F1 FROM HethongTK WHERE TKCon=0 AND Sohieu='" + shtk + "'") <> TKVT_ID Then shtk = "0"
                        Loop
                    End If
                    If ml = 8 Then
                        shtk = SelectSQL("SELECT HethongTK.Sohieu AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE MaLoai=8 AND MaVattu=" + CStr(vattu.MaSo))
                        Do While SoHieu2MaSo(shtk, "HethongTK") = 0
                            shtk = FrmGetStr.GetString("Sè hiÖu tµi kho¶n ghi doanh thu:", "Ho¸ ®¬n b¸n hµng")
                            If SelectSQL("SELECT TK_ID AS F1 FROM HethongTK WHERE TKCon=0 AND Sohieu='" + shtk + "'") <> TKDT_ID Then shtk = "0"
                        Loop
                    End If
                End If
                FrmChungtu.txtchungtu(0).Text = shtk
                FrmChungtu.txtchungtu(2).Text = vattu.sohieu
                FrmChungtu.txtChungtu_LostFocus 0
                FrmChungtu.txtChungtu_LostFocus 2
                FrmChungtu.txtchungtu(3).Text = Format(luong, Mask_2)
                FrmChungtu.txtchungtu(5).Text = ""
                If ml = 2 Then FrmChungtu.txtChungtu_LostFocus 3
                If ml = 8 Then FrmChungtu.txtchungtu(6).Text = Format(Fix(0.5 + luong * vattu.GiaBan1), Mask_2)
                FrmChungtu.CmdChitiet_chon
            End If
        End If
E1:
    Next

    Close #FileNum

    Recycle pCurDir + "BARCODE1.FIL"
    Name pCurDir + "BARCODE.FIL" As pCurDir + "BARCODE1.FIL"
    Recycle pCurDir + "BARCODE.FIL"

    Me.MousePointer = 0
    Exit Sub
E:
    MsgBox Err.Description
KT:
    Me.MousePointer = 0
End Sub

' kiem tra xuat hien ma so
Private Sub TxtVT_Change(Index As Integer)
    Label(26).Caption = ""
    If Len(txtVT(0).Text) < 0 Then txtVT(0).Text = ""
    txtVT(0).Text = Replace(txtVT(0).Text, "?", "")
    txtVT(9).Text = Replace(txtVT(9).Text, "?", "")
    txtVT(7).Text = Replace(txtVT(7).Text, "?", "")
    txtVT(8).Text = Replace(txtVT(8).Text, "?", "")

    Dim rs As Recordset
    Dim sql
    Dim i
    i = 0
    sql = ""
    Dim qq
    qq = CStr(Replace(txtVT(0).Text, ".", ""))
    If (Trim(Replace(txtVT(0).Text, ".", "")) = "#") Then qq = qq + "@*"

    Select Case Index
    Case 0:    ' lay thong tin

        sql = "SELECT top 1 * FROM KhachHang WHERE SoHieu = '" + qq + "' and left(sohieu,1) <> '#' order by maso desc"
        Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        If rs.recordCount <= 0 Then
            sql = "SELECT top 1 * FROM KhachHang WHERE SoHieu = '" + qq + "' order by maso desc"

        End If

    Case 9:
        sql = "SELECT top 1 * FROM KhachHang WHERE MST = '" + CStr(Replace(txtVT(9).Text, ".", "")) + "' and left(sohieu,1) <> '#'  order by maso desc"
        Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        If rs.recordCount <= 0 Then
            sql = "SELECT top 1 * FROM KhachHang WHERE MST = '" + CStr(Replace(txtVT(9).Text, ".", "")) + "' order by maso desc"
        End If
    End Select


    If sql <> "" Then
        Enable_thong_tin
        Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        If rs.recordCount > 0 Then
            txtVT(7).Text = rs!Ten
            Text1.Text = rs!Ten
            txtVT(8).Text = rs!DiaChi
            If txtVT(9).Text <> rs!mst Then
                txtVT(9).Text = rs!mst
            End If
            Disnable_thong_tin
            If Index <> 0 Then txtVT(0).Text = rs!sohieu
            Do While i <> CboLoai.ListCount
                If CboLoai.ItemData(i) = rs!MaPhanLoai Then
                    CboLoai.ListIndex = i
                    i = CboLoai.ListCount
                Else
                    i = i + 1
                End If
            Loop
            ' du thong tin roi se xuong duoi note
            If OptLoai(8).Value = False Then
                Disnable_thong_tin

                RFocus txt(1)
            End If
        End If

    Else
        ' If Len(Replace(Trim(txtVT(0).Text), ".", "")) > 0 Then
        If OptLoai(8).Value = False Then
            '   If Len(Replace(Trim(txtVT(0).Text), ".", "")) > 0 Then
            '        txtVT(9).Visible = True
            '        txtVT(7).Visible = True
            '        txtVT(8).Visible = True
            '        txtVT(2).Visible = True
            '        txtVT(3).Visible = True
            '        Enable_thong_tin
            '        Else
            '        If (Index = 0) Then
            '        txtVT(9).Visible = False
            '        txtVT(7).Visible = False
            '        txtVT(8).Visible = False
            '        txtVT(2).Visible = True
            '        txtVT(3).Visible = True
            '        End If
            '    End If
        End If
        '    If Index = 0 Then
        '       txtVT(7).Text = "..."
        '       txtVT(8).Text = "..."
        '       txtVT(9).Text = "..."
        '       txtVT(2).Text = "..."
        '       txtVT(3).Text = "..."
        '    End If
    End If


    If Len(Replace(Trim(txtVT(1).Text), ".", "")) > 0 Then
        Mo_thong_tin
        sql = "SELECT top 1 * FROM KhachHang WHERE SoHieu = '" + qq + "' order by maso desc"
        Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
        '  If rs.RecordCount > 0 Then Disnable_thong_tin
    Else
        Dong_thong_tin

    End If

    ' txtVT(0).Text = UCase(txtVT(0).Text)
    '  txtVT(1).Text = UCase(txtVT(1).Text)

    '    If Index <> 1 Then
    '    If Len(Replace(Trim(txtVT(0).Text), ".", "")) <= 0 Then
    '    txtVT(1).Text = "..."
    '    '    Dong_thong_tin
    '    End If
    '    End If

End Sub

Private Sub txtVT_DblClick(Index As Integer)
    Select Case Index
    Case 0, 9, 7, 8:

        txtVT(0).Text = FrmKhachHang.ChonKhachHang(txtVT(0).Text)

    End Select
End Sub

Private Sub Txtvt_GotFocus(Index As Integer)

    AutoSelect txtVT(Index)
End Sub
Private Sub Mo_thong_tin()
    txtVT(9).Visible = True
    txtVT(8).Visible = True
    txtVT(7).Visible = True
    txtVT(0).Visible = True
    txtVT(2).Visible = True
    txtVT(3).Visible = True
    CboLoai.Visible = True
    ' CboNguon(1).Visible = True
    Label1(0).Visible = True
    Label1(1).Visible = True
    Label1(3).Visible = True
    Label1(16).Visible = True
    Label1(6).Visible = True
    Label1(7).Visible = True

End Sub
Private Sub Dong_thong_tin()
    txtVT(9).Visible = False
    txtVT(8).Visible = False
    txtVT(7).Visible = False
    txtVT(0).Visible = False
    txtVT(2).Visible = False
    txtVT(3).Visible = False
    Label1(0).Visible = False
    Label1(1).Visible = False
    Label1(3).Visible = False
    Label1(16).Visible = False
    Label1(6).Visible = False
    Label1(7).Visible = False
    CboLoai.Visible = False
    ' CboNguon(1).Visible = False
    txtVT(0).Text = "..."
    'txtVT(1).Text = ""
    txtVT(7).Text = "..."
    txtVT(8).Text = "..."
    If txtVT(9).Text <> "..." Then
        txtVT(9).Text = "..."
    End If
    ' txtVT(2).Text = "01GTKT3/001"
    hien_thong_tin_mau_HD
    txtVT(3).Text = "01GTKT"
End Sub
Sub hien_thong_tin_mau_HD()
' If OptLoai.item(8).Value = True Then
    If (Len(txtVT(2).Text) <= 0) Then
        txtVT(2).Text = SelectSQL("SELECT MauSoHD AS F1 FROM chungtu WHERE  maso in (select max(maso) from chungtu)")
    End If
    '             Else
    '             If (Len(txtVT(2).Text) <= 0) Then
    '           txtVT(2).Text = SelectSQL("SELECT MauSoHD AS F1 FROM chungtu WHERE  maso in (select max(maso) from chungtu where maloai <> 8)")
    '               End If
    'End If
End Sub
Private Sub Enable_thong_tin()
    txtVT(9).Enabled = True
    txtVT(8).Enabled = True
    txtVT(7).Enabled = True
    'txtVT(0).Enabled = True

End Sub
Private Sub Disnable_thong_tin()
    txtVT(9).Enabled = False
    txtVT(8).Enabled = False
    txtVT(7).Enabled = False
    '  txtVT(0).Enabled = False

End Sub
Sub them_dau_cham_txtVT(Index As Integer)
    If Len(Replace(Trim(txtVT(Index).Text), ".", "")) <= 0 Then txtVT(Index).Text = "..."
    RFocus txtVT(Index)
End Sub
Sub them_dau_cham_txt(Index As Integer)
    If Len(Replace(Trim(txt(Index).Text), ".", "")) <= 0 Then txt(Index).Text = "..."
    RFocus txt(Index)
End Sub
Private Sub TxtVT_KeyPress(Index As Integer, KeyAscii As Integer)
'txtVT(0).BackColor = ColorConstants.vbWhite
    Label(26).Caption = ""
    Select Case Index
    Case 0:
        If KeyAscii = 13 And KHDetail Then
            If Len(Replace(Trim(txtVT(0).Text), ".", "")) > 0 Then
                them_dau_cham_txtVT (9)    ' nhay sang ma so thue
                txtVT(0).Text = UCase(txtVT(0).Text)
                txtVT(1).Text = UCase(txtVT(1).Text)
            Else
                '////////////////////////
                If Len(Replace(Trim(txtVT(Index).Text), ".", "")) <= 0 Then
                    If KeyAscii = 13 And KHDetail Then    '63 neu nhan phim ? da chon dc khach hang
                        txtVT(0).Text = FrmKhachHang.ChonKhachHang(txtVT(0).Text)    ' lay thong tin khach hang tu bang khach hang
                        If Len(Replace(Trim(txtVT(Index).Text), ".", "")) <= 0 Then
                            them_dau_cham_txtVT (0)    ' chuyen xuong dien giai
                        Else

                            Disnable_thong_tin    'an thong tin truoc khi chuyen xuong dien giai
                            them_dau_cham_txt (1)    ' chuyen xuong dien giai
                        End If
                    End If
                Else
                    '////////////////////
                    txtVT(1).Text = "..."
                    them_dau_cham_txt (1)    ' xuong ghi chu chung tu
                    '  Dong_thong_tin ' che thong tin lai
                End If
            End If
            If (txtVT(9).Visible = False) Then them_dau_cham_txt (1)    ' xuong ghi chu chung tu
            txtVT(0).Text = UCase(txtVT(0).Text)
            txtVT(1).Text = UCase(txtVT(1).Text)
        End If

        If KeyAscii = 63 And KHDetail Then    '63 neu nhan phim ? da chon dc khach hang
            txtVT(0).Text = FrmKhachHang.ChonKhachHang(txtVT(0).Text)    ' lay thong tin khach hang tu bang khach hang
            Disnable_thong_tin    'an thong tin truoc khi chuyen xuong dien giai
            them_dau_cham_txt (1)    ' chuyen xuong dien giai

        End If

    Case 1:    ' neu dang o ky hieu
        If KeyAscii = 13 And KHDetail Then
            Enable_thong_tin
            If Len(Replace(Trim(txtVT(1).Text), ".", "")) > 0 Then
                them_dau_cham_txtVT (0)     'txtVT(0).BackColor = ColorConstants.vbGreen
            Else
                them_dau_cham_txt (1)
                Dong_thong_tin
            End If
            txtVT(1).Text = UCase(txtVT(1).Text)
        End If
    Case 2:    ' neu dang o ky hieu
        If KeyAscii = 13 And KHDetail Then
            '   them_dau_cham_txtVT (3)
            them_dau_cham_txtVT (0)
            txtVT(2).Text = UCase(txtVT(2).Text)
        End If
    Case 3:    ' neu dang o ky hieu
        If KeyAscii = 13 And KHDetail Then
            them_dau_cham_txtVT (0)
            txtVT(3).Text = UCase(txtVT(3).Text)
        End If
    Case 9:    ' neu dang o ma so thue

        If KeyAscii = 63 And KHDetail Then    'neu nhan phim ? da chon dc khach hang
            txtVT(0).Text = FrmKhachHang.ChonKhachHang(txtVT(0).Text)    ' lay thong tin khach hang tu bang khach hang
            them_dau_cham_txt (1)    ' chuyen xuong dien giai
            Disnable_thong_tin    'an thong tin truoc khi chuyen xuong dien giai

        Else
            ' them_dau_cham_txtVT (7)
            If KeyAscii = 13 And KHDetail Then
                them_dau_cham_txtVT (7)

            End If
        End If
    Case 7:    ' neu dang o ma so thue
        If KeyAscii = 13 And KHDetail Then
            them_dau_cham_txtVT (8)    ' chuyen xuong dien giai
        End If
    Case 8:    ' neu dang o ma so thue
        If KeyAscii = 13 And KHDetail Then
            them_dau_cham_txt (1)    ' chuyen xuong dien giai
        End If
    End Select

End Sub
Private Sub Nhay_Chuot(Index As Integer)
    Select Case Index
    Case 0:
        If Len(Replace(Trim(txtVT(0).Text), ".", "")) > 0 Then
            txtVT(0).Text = UCase(txtVT(0).Text)
            txtVT(1).Text = UCase(txtVT(1).Text)
            ' them_dau_cham_txtVT (9) ' nhay sang ma so thue
            If Len(Replace(Trim(txtVT(9).Text), ".", "")) <= 0 Then txtVT(9).Text = "..."
        Else
            'txtVT(1).Text = "..."
            If Len(Replace(Trim(txt(1).Text), ".", "")) <= 0 Then txt(1).Text = "..."
            '  Dong_thong_tin ' che thong tin lai
        End If


    Case 1:    ' neu dang o ky hieu
        Enable_thong_tin

        If Len(Replace(Trim(txtVT(1).Text), ".", "")) > 0 Then

            If Len(Replace(Trim(txtVT(0).Text), ".", "")) <= 0 Then txtVT(0).Text = "..."
        Else
            txtVT(1).Text = "..."
            If Len(Replace(Trim(txt(1).Text), ".", "")) <= 0 Then txt(1).Text = "..."
            ' Dong_thong_tin
        End If
        txtVT(1).Text = UCase(txtVT(1).Text)
    Case 9:    ' neu dang o ma so thue
        'RFocus CboLoai
        'If Len(Replace(Trim(txtVT(7).Text), ".", "")) <= 0 Then txtVT(7).Text = "..."
    Case 7:    ' neu dang o ma so thue
        If Len(Replace(Trim(txtVT(8).Text), ".", "")) <= 0 Then txtVT(8).Text = "..."

    Case 8:    ' neu dang o ma so thue
        If Len(Replace(Trim(txt(1).Text), ".", "")) <= 0 Then txt(1).Text = "..."
    End Select

End Sub
Private Sub TxtVT_LostFocus(Index As Integer)
    Label(26).Caption = ""
    Label(26).Caption = ""
    '  Nhay_Chuot (Index)
    txtVT(2).Text = UCase(txtVT(2).Text)
    Select Case Index
    Case 0:
        If Len(Replace(Trim(txtVT(Index).Text), ".", "")) <= 0 Then
            txtVT(0).Text = FrmKhachHang.ChonKhachHang(txtVT(0).Text)    ' lay thong tin khach hang tu bang khach hang
            If Len(Replace(Trim(txtVT(0).Text), ".", "")) <= 0 Then
                txtVT(0).Text = "..."
                ' RFocus txtVT(0)
            Else
                Disnable_thong_tin    'an thong tin truoc khi chuyen xuong dien giai
                them_dau_cham_txt (1)    ' chuyen xuong dien giai
            End If
            Exit Sub
        End If

        If Len(Replace(Trim(txtVT(0).Text), ".", "")) > 0 Then
            txtVT(0).Text = UCase(txtVT(0).Text)
            txtVT(1).Text = UCase(txtVT(1).Text)
            ' them_dau_cham_txtVT (9) ' nhay sang ma so thue
            If Len(Replace(Trim(txtVT(9).Text), ".", "")) <= 0 Then txtVT(9).Text = "..."
        Else
            'txtVT(1).Text = "..."
            If Len(Replace(Trim(txt(1).Text), ".", "")) <= 0 Then txt(1).Text = "..."
            '  Dong_thong_tin ' che thong tin lai
        End If


    Case 1:    ' neu dang o ky hieu
        Enable_thong_tin

        If Len(Replace(Trim(txtVT(1).Text), ".", "")) > 0 Then

            If Len(Replace(Trim(txtVT(0).Text), ".", "")) <= 0 Then txtVT(0).Text = "..."
        Else
            txtVT(1).Text = "..."
            If Len(Replace(Trim(txt(1).Text), ".", "")) <= 0 Then txt(1).Text = "..."
            ' Dong_thong_tin
        End If
        txtVT(1).Text = UCase(txtVT(1).Text)
    Case 9:    ' neu dang o ma so thue
        'RFocus CboLoai
        'If Len(Replace(Trim(txtVT(7).Text), ".", "")) <= 0 Then txtVT(7).Text = "..."
    Case 7:    ' neu dang o ma so thue
        If Len(Replace(Trim(txtVT(8).Text), ".", "")) <= 0 Then txtVT(8).Text = "..."

    Case 8:    ' neu dang o ma so thue
        If Len(Replace(Trim(txt(1).Text), ".", "")) <= 0 Then txt(1).Text = "..."
    End Select
End Sub



Private Sub Command1_Click()


'             FrmKhachHang.Command_Click 0
'
'             FrmKhachHang.txtVT(0).Text = txtVT(0).Text
'             FrmKhachHang.txtVT(1).Text = txtVT(7).Text
'             FrmKhachHang.txtVT(2).Text = txtVT(8).Text
'             FrmKhachHang.txtVT(3).Text = txtVT(9).Text
'             FrmKhachHang.txtVT(4).Text = "00000-000"
'             FrmKhachHang.txtVT(5).Text = "00000-000"
'             FrmKhachHang.txtVT(6).Text = " "
'             FrmKhachHang.txtVT(7).Text = " "
'             FrmKhachHang.txtVT(8).Text = " "
'             FrmKhachHang.txtVT(9).Text = " "
'             If Left(txtVT(0).Text, 0) = "#" Then
'             FrmKhachHang.txtVT(0).Text = "#" + CStr(Year(Date) - 2000) + CStr(Month(Date)) + CStr(Day(Date)) + CStr(Hour(Now)) + CStr(Minute(Now)) + CStr(Second(Now))
'                             '   ckh.MaPhanLoai = SelectSQL("SELECT MaSo AS F1 FROM PhanLoaiKhachHang WHERE LEFT(SoHieu,1)='#'")
'             FrmKhachHang.Command_Click 1

    FVAT.T(0).Text = txtVT(0).Text
    FVAT.T(7).Text = txtVT(7).Text
    FVAT.T(8).Text = txtVT(8).Text
    FVAT.T(9).Text = txtVT(9).Text
    FVAT.Command_Click



    '  End If


    '= txtVT(0).Text '.SoHieu
    '= txtVT(1).Text '.Ten
    'txtVT(2).Text '.DiaChi =
    'txtVT(3).Text '.mst =
    'txtVT(4).Text '.Tel =
    'txtVT(5).Text ' .Fax =
    'txtVT(6).Text '.email =
    'txtVT(7).Text '.DaiDien =
    'txtVT(8).Text '.taikhoan =
    'txtVT(9).Text '.GhiChu =
    'Cdbl5 (txtVT(10).Text) '.DuMax =
    'FrmKhachHangCboNT.ItemData (CboLoai.ListIndex) '.MaNT =

    ' If FrmKhachHang.KiemTraSoLieu = False Then
    '   RFocus txtVT(3)
    '  Else

    'ckh.GhiKhachHang

    '  End If
    '   Luu_hoa_don
End Sub
'    Private Sub Luu_hoa_don()
'        Dim i As Integer
'
'        If KHDetail And ckh.MaSo = 0 And Left(T(0).Text, 1) = "#" Then
'            ckh.Ten = T(7).Text
'            ckh.DiaChi = T(8).Text
'            ckh.mst = T(9).Text
'            ckh.SoHieu = "#" + CStr(Year(Date) - 2000) + CStr(Month(Date)) + CStr(Day(Date)) + CStr(Hour(Now)) + CStr(Minute(Now)) + CStr(Second(Now))
'            ckh.MaPhanLoai = SelectSQL("SELECT MaSo AS F1 FROM PhanLoaiKhachHang WHERE LEFT(SoHieu,1)='#'")
'            If ckh.GhiKhachHang <> 0 Then GoTo Er
'        End If
'        If KHDetail And ckh.MaSo = 0 Then
'Er:
'           ' RFocus T(0)
'            Exit Sub
'        End If
'        With h
'            .MaKhachHang = ckh.MaSo
'            .KyHieu = IIf(Len(T(1).Text) > 0, T(1).Text, "")
'            .sohd = IIf(Len(txt(0).Text) > 0, txt(0).Text, "")
'            .NgayPH = MedNgay(0).Text 'ngay
'            .MatHang = IIf(Len(T(3).Text) > 0, T(3).Text, "")
'            .SoLuong = Cdbl5(T(4).Text)
'            .ThanhTien = Cdbl5(T(5).Text)
'            .TyLe = CInt5(T(6).Text)
'            .HD = 1 'ChkV(0).Value
'            .KCT = 0 'ChkV(1).Value
'            .HDBL = 0 'ChkV(2).Value
'            .NK = 0 'ChkV(3).Value
'            .ts = 0 'ChkV(4).Value
'            .DC = 0 'ChkV(5).Value
'            .TenKH = T(7).Text
'            .DiaChiKH = T(8).Text
'            .MSTKH = T(9).Text
'            .HTTT = IIf(Len(T(11).Text) > 0, T(11).Text, "...")
'            .MauSo = IIf(Len(T(12).Text) > 0, T(12).Text, "...")
'            .tygia = Cdbl5(T(13).Text)
'        End With
'        Unload Me
'    End Sub
'

























