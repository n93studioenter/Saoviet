VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form CPGV 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ph©n bæ chi phÝ hµng nhËp khÈu"
   ClientHeight    =   6930
   ClientLeft      =   135
   ClientTop       =   1020
   ClientWidth     =   11295
   ClipControls    =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "Cpgv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6930
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin MSGrid.Grid Grd 
      Height          =   3135
      Left            =   0
      TabIndex        =   70
      Top             =   1800
      Width           =   11055
      _Version        =   65536
      _ExtentX        =   19500
      _ExtentY        =   5530
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
      Cols            =   16
   End
   Begin VB.OptionButton OptPB 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gi¸ trÞ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   10200
      TabIndex        =   22
      Top             =   0
      Width           =   975
   End
   Begin VB.OptionButton OptPB 
      BackColor       =   &H00E0E0E0&
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
      Height          =   195
      Index           =   0
      Left            =   9120
      TabIndex        =   68
      Top             =   0
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.ComboBox CboVV 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   64
      ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox CboVV 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   1
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   63
      ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox CboVV 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   2
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   62
      ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   13
      Left            =   10200
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "Cpgv.frx":57E2
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   12
      Left            =   9360
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "Cpgv.frx":57E4
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   11
      Left            =   8520
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "Cpgv.frx":57E6
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton CmdTK 
      Height          =   375
      Index           =   2
      Left            =   3600
      Picture         =   "Cpgv.frx":57E8
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   6480
      Width           =   375
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   2
      Left            =   2160
      LinkItem        =   "Sè hiÖu tµi kho¶n cÇn xem"
      MaxLength       =   12
      TabIndex        =   50
      Tag             =   "0"
      Text            =   "3388"
      Top             =   6525
      Width           =   1335
   End
   Begin VB.CommandButton CmdTK 
      Height          =   375
      Index           =   1
      Left            =   3600
      Picture         =   "Cpgv.frx":5966
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   1
      Left            =   2160
      LinkItem        =   "Sè hiÖu tµi kho¶n cÇn xem"
      MaxLength       =   12
      TabIndex        =   47
      Tag             =   "0"
      Text            =   "33332"
      Top             =   6165
      Width           =   1335
   End
   Begin VB.CommandButton CmdTK 
      Height          =   375
      Index           =   0
      Left            =   3600
      Picture         =   "Cpgv.frx":5AE4
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox txtShTk 
      Height          =   285
      Index           =   0
      Left            =   2160
      LinkItem        =   "Sè hiÖu tµi kho¶n cÇn xem"
      MaxLength       =   12
      TabIndex        =   44
      Tag             =   "0"
      Text            =   "1561"
      Top             =   5805
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Caption         =   "&Xö lý"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "&Done"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ph©n bæ tù ®éng"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Xem b¸o c¸o"
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   6840
      MaxLength       =   20
      TabIndex        =   12
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   6600
      MaxLength       =   20
      TabIndex        =   11
      Top             =   4920
      Width           =   255
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   10
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   4920
      MaxLength       =   20
      TabIndex        =   9
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   4080
      MaxLength       =   20
      TabIndex        =   8
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   3120
      MaxLength       =   20
      TabIndex        =   7
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   2760
      MaxLength       =   20
      TabIndex        =   6
      Top             =   4920
      Width           =   375
   End
   Begin VB.ComboBox Cbo 
      Height          =   330
      Index           =   0
      ItemData        =   "Cpgv.frx":5C62
      Left            =   960
      List            =   "Cpgv.frx":5C64
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   1
      Left            =   7560
      Picture         =   "Cpgv.frx":5C66
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Xem b¸o c¸o"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   2
      Left            =   8760
      Picture         =   "Cpgv.frx":6DD8
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "In b¸o c¸o ra m¸y in"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   14
      Left            =   6120
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Cpgv.frx":823A
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtTon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   7680
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "Cpgv.frx":823C
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdct 
      Caption         =   "­"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11040
      TabIndex        =   17
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   0
      Left            =   9960
      Picture         =   "Cpgv.frx":823E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   0
      MaxLength       =   20
      TabIndex        =   3
      Tag             =   "-1"
      Top             =   4920
      Width           =   735
   End
   Begin VB.ComboBox Cbo 
      Height          =   330
      Index           =   1
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ph©n bæ theo"
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
      Left            =   7920
      TabIndex        =   69
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chi phÝ 4"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   23
      Left            =   10200
      TabIndex        =   58
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chi phÝ 3"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   22
      Left            =   9360
      TabIndex        =   57
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chi phÝ 2"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   21
      Left            =   8520
      TabIndex        =   56
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chi phÝ 1"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   13
      Left            =   7680
      TabIndex        =   36
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè tiÒn"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   6840
      TabIndex        =   35
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   6600
      TabIndex        =   34
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ThuÕ NK"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   6600
      TabIndex        =   32
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thµnh tiÒn VND"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   5
      Left            =   5760
      TabIndex        =   28
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label LbTT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th«ng tin 1"
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
      TabIndex        =   67
      Tag             =   "Index"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LbTT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th«ng tin 2"
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
      Left            =   4080
      TabIndex        =   66
      Tag             =   "Index"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LbTT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th«ng tin 3"
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
      Left            =   7920
      TabIndex        =   65
      Tag             =   "Index"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label CP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   10275
      TabIndex        =   61
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label CP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   9435
      TabIndex        =   60
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label CP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   8595
      TabIndex        =   59
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chi phÝ kh¸c"
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
      Left            =   120
      TabIndex        =   55
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ThuÕ nhËp khÈu"
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
      TabIndex        =   54
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tµi kho¶n ghi nî hµng ho¸"
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
      TabIndex        =   53
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   52
      Tag             =   "1"
      Top             =   6525
      Width           =   4695
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   49
      Tag             =   "1"
      Top             =   6165
      Width           =   4695
   End
   Begin VB.Label LbTenTk 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   46
      Tag             =   "1"
      Top             =   5805
      Width           =   4695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PN"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   17
      Left            =   2760
      TabIndex        =   42
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chi phÝ ®· ph©n bæ"
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
      Left            =   6120
      TabIndex        =   41
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label CP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   7755
      TabIndex        =   40
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label CP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   9555
      TabIndex        =   39
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th¸ng"
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
      TabIndex        =   38
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tæng chi phÝ"
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
      Left            =   8040
      TabIndex        =   37
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VËt t­"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   33
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thµnh tiÒn NT"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   8
      Left            =   4920
      TabIndex        =   31
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§¬n gi¸ NT"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   7
      Left            =   3120
      TabIndex        =   30
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tû gi¸ tÝnh thuÕ"
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
      Left            =   4680
      TabIndex        =   29
      Top             =   0
      Width           =   1335
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   4080
      TabIndex        =   27
      Top             =   1320
      Width           =   855
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   26
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tªn vËt t­"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   25
      Top             =   1560
      Width           =   1575
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   24
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hîp ®ång"
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
      TabIndex        =   23
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "CPGV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vt As New ClsVattu

Private Sub Cbo_Click(Index As Integer)
    If Me.tag = 0 Then
        Select Case Index
            Case 0:
                Int_RecsetToCbo "SELECT DienGiai AS F1, MaSo as F2 FROM DoituongCT WHERE MaSo>1 ORDER BY DienGiai", Cbo(1)
                If Cbo(1).ListCount = 0 Then ClearGrid Grd, Grd.tag
            Case 1:
                LietKeChiPhiHD Cbo(1).ItemData(Cbo(1).ListIndex)
        End Select
    End If
End Sub

Private Sub CboVV_Click(Index As Integer)
    Dim sops As Double
    
    Select Case Index
        Case 0:      CboVVClick CboVV(0), CboVV(1)
        Case 1:
                            sops = SelectSQL("SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-1) + " WHERE MaDT2=" + CStr(CboVV(1).ItemData(CboVV(1).ListIndex)) + " AND HethongTK.SoHieu LIKE '1562*'")
                            If sops <> 0 Then
                                CP(1).Caption = Format(sops, Mask_0)
                                CP(0).Caption = Format(Cdbl5(CP(1).Caption) + Cdbl5(CP(2).Caption) + Cdbl5(CP(3).Caption) + Cdbl5(CP(4).Caption), Mask_0)
                                PhanBo 1
                            End If
    End Select
End Sub

Private Sub CmdCt_Click()
    Dim cpx(1 To 4) As Double, tt2 As Double, tnk As Double, tg As Double, luong As Double, pn As String, tt As Double, i As Integer
    Dim mdt(1 To 3) As Long
        
    If Me.tag = 0 And Cbo(1).ListIndex < 0 Then
        RFocus Cbo(1)
        Exit Sub
    End If
    If vt.MaSo = 0 Then
        RFocus txtTon(0)
        Exit Sub
    End If
    luong = Cdbl5(txtTon(5).Text)
    If luong = 0 Then
        RFocus txtTon(0)
        Exit Sub
    End If
    For i = 1 To 4
        cpx(i) = Cdbl5(txtTon(9 + i).Text)
    Next
    tt2 = Cdbl5(txtTon(6).Text)
    tt = Cdbl5(txtTon(7).Text)
    tnk = Cdbl5(txtTon(9).Text)
        
    tg = Cdbl5(txtTon(14).Text)
    If txtTon(3).Text <> "" Then pn = txtTon(3).Text Else pn = "..."
    
    If Me.tag = 0 Then
        For i = 1 To 3
            If CboVV(i - 1).ListIndex >= 0 Then mdt(i) = CboVV(i - 1).ItemData(CboVV(i - 1).ListIndex)
        Next
        
        ExecuteSQL5 "INSERT INTO CPGVHD (MaSo,MaDT, MaVattu, CP1, CP2, CP3, CP4, TT2, TNK, TyGia1, TT,  SL, PN, MaDT1, MaDT2, MaDT3) VALUES (" + CStr(Lng_MaxValue("MaSo", "CPGVHD") + 1) + "," + CStr(Cbo(1).ItemData(Cbo(1).ListIndex)) + "," + CStr(vt.MaSo) + "," + DoiDau(cpx(1)) + "," + DoiDau(cpx(2)) + "," + DoiDau(cpx(3)) + "," + DoiDau(cpx(4)) + "," + DoiDau(tt2) + "," + DoiDau(tnk) + "," + DoiDau(tg) + "," + DoiDau(tt) + "," + DoiDau(luong) + ",'" + pn + "'," + CStr(mdt(1)) + "," + CStr(mdt(2)) + "," + CStr(mdt(3)) + ")"
    End If
    Grd.AddItem vt.sohieu + Chr(9) + vt.TenVattu + Chr(9) + vt.DonVi + Chr(9) + pn + Chr(9) + Format(tt2 / luong, Mask_2) + Chr(9) + Format(luong, Mask_2) + Chr(9) + Format(tt2, Mask_2) _
        + Chr(9) + Format(tt, Mask_0) + Chr(9) + txtTon(8).Text + Chr(9) + Format(tnk, Mask_0) + Chr(9) + Format(cpx(1), Mask_0) + Chr(9) + Format(cpx(2), Mask_0) + Chr(9) + Format(cpx(3), Mask_0) + Chr(9) + Format(cpx(4), Mask_0) + Chr(9) + CStr(vt.MaSo) + Chr(9) + CStr(Lng_MaxValue("MaSo", "CPGVHD")), 0
    
    vt.InitVattuMaSo 0
    For i = 0 To 13
        txtTon(i).Text = ""
    Next
    TinhTien
    RFocus txtTon(0)
End Sub

Private Sub cmdtk_Click(Index As Integer)
    txtShTk(Index).Text = FrmTaikhoan.ChonTk(txtShTk(Index).Text)
End Sub

Private Sub Command_Click(Index As Integer)
    Dim i As Integer, k As Integer, sl As Double, tt As Double, mtk As Long, tnk As Double, cpk As Double, sh As String, mkh As Long, gt As Double
    
    Select Case Index
        Case 0:
            Unload Me
        Case 1, 2:
            ChiPhiHD Index - 1
        Case 3:
            With Grd
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 0
                    If .Text = "" Then Exit For
                    .col = 10
                    .Text = ""
                    .col = 11
                    .Text = ""
                    .col = 12
                    .Text = ""
                    .col = 13
                    .Text = ""
                Next
            End With
            For i = 1 To 4
                PhanBo i
            Next
        Case 4:
            If Me.tag = 0 Then
                If Cbo(1).ListIndex < 0 Then Exit Sub
                If GetTK_ID(txtShTk(0).Text, 0, , i) <> TKVT_ID Then
                    RFocus txtShTk(0)
                    Exit Sub
                End If
                For i = 0 To 2
                    mtk = 0
                    GetTK_ID txtShTk(i).Text, mtk, , k
                    If mtk = 0 Or k <> 0 Then
                        RFocus txtShTk(i)
                        Exit Sub
                    End If
                Next
            End If
            k = 0
            Load FrmChungtu
            FrmChungtu.OptLoai(1).Value = True
            If Me.tag = 0 Then
                For i = 1 To pSoVV
                    If CboVV(i - 1).ListIndex >= 0 Then SetListIndex FrmChungtu.CboVV(i - 1), CboVV(i - 1).ItemData(CboVV(i - 1).ListIndex)
                Next
            End If
            With Grd
                For i = 0 To .Rows - 1
                    .Row = i
                    .col = 0
                    If Len(.Text) = 0 Then Exit For
                    .col = 5
                    sl = Cdbl5(.Text)
                    .col = 7
                    tt = Cdbl5(.Text)
                    gt = tt
                    .col = 9
                    tnk = tnk + Cdbl5(.Text)
                    tt = tt + Cdbl5(.Text)
                    .col = 10
                    cpk = cpk + Cdbl5(.Text)
                    tt = tt + Cdbl5(.Text)
                    gt = gt + Cdbl5(.Text)
                    .col = 11
                    cpk = cpk + Cdbl5(.Text)
                    tt = tt + Cdbl5(.Text)
                    gt = gt + Cdbl5(.Text)
                    .col = 12
                    cpk = cpk + Cdbl5(.Text)
                    tt = tt + Cdbl5(.Text)
                    gt = gt + Cdbl5(.Text)
                    .col = 13
                    cpk = cpk + Cdbl5(.Text)
                    tt = tt + Cdbl5(.Text)
                    gt = gt + Cdbl5(.Text)
                    .col = 0
                    FrmChungtu.txtchungtu(0).Text = txtShTk(0).Text
                    FrmChungtu.txtchungtu(2).Text = .Text
                    FrmChungtu.txtChungtu_LostFocus 0
                    FrmChungtu.txtChungtu_LostFocus 2
                    FrmChungtu.txtchungtu(3).Text = Format(sl, Mask_2)
                    FrmChungtu.txtchungtu(5).Text = Format(IIf(Me.tag = 0, tt, gt), Mask_0)
                    If sl <> 0 Then FrmChungtu.txtchungtu(4).Text = Format(IIf(Me.tag = 0, tt, gt) / sl, Mask_2)
                    FrmChungtu.txtchungtu(6).Text = ""
                    FrmChungtu.CmdChitiet_Click
                    k = k + 1
                Next
            End With
            If k > 0 And Me.tag = 0 Then
                pFunction = 10
                pCT_ID = 1000000000 + Cbo(1).ItemData(Cbo(1).ListIndex)
                SetListIndex FrmChungtu.CboVV(2), Cbo(1).ItemData(Cbo(1).ListIndex)
                FrmChungtu.MaSoCT = SelectSQL("SELECT MaCT AS F1 FROM ChungTu WHERE CT_ID=" + CStr(pCT_ID))
                If FrmChungtu.MaSoCT > 0 Then
                    mkh = SelectSQL("SELECT Max(MaKHC) AS F1 FROM ChungTu WHERE MaCT=" + CStr(FrmChungtu.MaSoCT))
                    If mkh > 0 Then sh = MaSo2SoHieu(mkh, "KhachHang")
                End If
                If pTygia > 0 Then FrmChungtu.txtchungtu(7).Text = txtTon(14).Text
                If tnk <> 0 Then
                    FrmChungtu.txtchungtu(0).Text = txtShTk(1).Text
                    FrmChungtu.txtChungtu_LostFocus 0
                    FrmChungtu.txtchungtu(5).Text = ""
                    FrmChungtu.txtchungtu(6).Text = Format(tnk, Mask_0)
                    FrmChungtu.CmdChitiet_Click
                End If
                If cpk <> 0 Then
                    FrmChungtu.txtchungtu(0).Text = txtShTk(2).Text
                    FrmChungtu.txtChungtu_LostFocus 0
                    FrmChungtu.txtchungtu(5).Text = ""
                    FrmChungtu.txtchungtu(6).Text = Format(cpk, Mask_0)
                    FrmChungtu.CmdChitiet_Click
                End If
                FrmChungtu.txtchungtu(5).Text = ""
                FrmChungtu.txtchungtu(6).Text = ""
                FrmChungtu.txtchungtu(0).Text = "3311"
                FrmChungtu.txtChungtu_LostFocus 0
                If mkh > 0 Then
                    FrmChungtu.txtchungtu(2).Text = sh
                    FrmChungtu.txtChungtu_LostFocus 2
                End If
                FrmChungtu.txtchungtu(6).Text = Format(FrmChungtu.SoPSConLai, Mask_0)
            End If
            If k > 0 Then FrmChungtu.Show 1
    End Select
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    
    AddMonthToCbo Cbo(0)
    If Me.tag > 0 Then
        Me.Caption = "Gi¸ vèn hµng nhËp kho"
        txtTon(8).Enabled = False
        txtTon(9).Enabled = False
        
        Label(0).Visible = False
        Label(15).Visible = False
        Label(6).Visible = False
        
        Cbo(0).Visible = False
        Cbo(1).Visible = False
        
        txtTon(14).Visible = False
        
        Me.Height = 6120
    Else
        For i = 1 To pSoVV
            LbTT(i - 1).Visible = True
            CboVV(i - 1).Visible = True
            Int_RecsetToCbo "SELECT MaSo As F2,DienGiai As F1 FROM DoituongCT" + CStr(i) + " ORDER BY DoituongCT" + CStr(i) + ".DienGiai", CboVV(i - 1)
        Next
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    ColumnSetUp Grd, 0, 700, 2
    ColumnSetUp Grd, 1, 1540, 0
    ColumnSetUp Grd, 2, 460, 2
    ColumnSetUp Grd, 3, 340, 0
    ColumnSetUp Grd, 4, 940, 1
    ColumnSetUp Grd, 5, 820, 1
    ColumnSetUp Grd, 6, 820, 1
    ColumnSetUp Grd, 7, 820, 1
    ColumnSetUp Grd, 8, 220, 2
    ColumnSetUp Grd, 9, 820, 1
    ColumnSetUp Grd, 10, 820, 1
    ColumnSetUp Grd, 11, 820, 1
    ColumnSetUp Grd, 12, 820, 1
    ColumnSetUp Grd, 13, 820, 1
    ColumnSetUp Grd, 14, 1, 0
    ColumnSetUp Grd, 15, 1, 0
        
    For i = 0 To 2
        txtShTk_LostFocus i
    Next
    
    txtTon(14).Text = Format(TyGiaNT(0), Mask_2)
    
    SetFont Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask Then
        Select Case KeyCode
            Case vbKeyX:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyI:
                RFocus Command(2)
                Command_Click 2
            Case vbKeyV:
                RFocus Command(0)
                Command_Click 0
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub LietKeChiPhiHD(mhd As Long)
    Dim rs As Recordset, sql As String, dg2 As Double, cpx(1 To 4) As Double, i As Integer, tlthue As Double
    
    ClearGrid Grd, Grd.tag
    sql = "SELECT CPGVHD.*, Vattu.SoHieu AS SHVT, TenVattu, DonVi, ThueNK FROM CPGVHD INNER JOIN Vattu ON CPGVHD.MaVattu=Vattu.MaSo WHERE MaDT=" + CStr(mhd) + " ORDER BY Vattu.SoHieu DESC"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    
    If rs.EOF Then GoTo a
    
    SetListIndex CboVV(0), rs!MaDT1
    SetListIndex CboVV(1), rs!MaDT2
    SetListIndex CboVV(2), rs!MaDT3
    txtTon(14).Text = Format(rs!tygia1, Mask_2)
    Do While Not rs.EOF
        dg2 = rs!tt2 / rs!sl
        For i = 1 To 4
            cpx(i) = cpx(i) + rs.Fields("CP" + CStr(i))
        Next
        If rs!tt <> 0 Then tlthue = Fix(0.5 + 1000 * rs!tnk / rs!tt) / 10 Else tlthue = 0
        Grd.AddItem rs!shvt + Chr(9) + rs!TenVattu + Chr(9) + rs!DonVi + Chr(9) + rs!pn + Chr(9) + Format(dg2, Mask_2) + Chr(9) + Format(rs!sl, Mask_2) + Chr(9) + Format(rs!tt2, Mask_2) _
            + Chr(9) + Format(rs!tt, Mask_0) + Chr(9) + CStr(tlthue) + Chr(9) + Format(rs!tnk, Mask_0) + Chr(9) + Format(rs!cp1, Mask_0) + Chr(9) + Format(rs!cp2, Mask_0) + Chr(9) + Format(rs!cp3, Mask_0) + Chr(9) + Format(rs!cp4, Mask_0) + Chr(9) + CStr(rs!MaVattu) + Chr(9) + CStr(rs!MaSo), 0
        rs.MoveNext
    Loop
    Grd.Rows = IIf(rs.RecordCount >= Grd.tag, rs.RecordCount, Grd.tag)
a:
    For i = 1 To 4
        CP(i).Caption = Format(cpx(i), Mask_0)
    Next
    
    rs.Close
    Set rs = Nothing
    TinhTien
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set vt = Nothing
End Sub

Private Sub Grd_Click()
    Dim i As Integer
    
    With Grd
        .col = 0
        If .Text = "" Then Exit Sub
        For i = 0 To 13
            .col = i
            txtTon(i).Text = .Text
        Next
        txtTon_LostFocus 0
        .col = 15
        ExecuteSQL5 "DELETE * FROM CPGVHD WHERE MaSo=" + .Text
        .RemoveItem .Row
        If .Rows < .tag Then .Rows = .tag
        RFocus txtTon(10)
    End With
End Sub

Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , Grd, Grd.col
    End If
End Sub

Private Sub txtShTk_LostFocus(Index As Integer)
    Dim mtk As Long
    LbTenTk(Index).Caption = tentk(txtShTk(Index).Text, mtk)
End Sub

Private Sub txtTon_GotFocus(Index As Integer)
    AutoSelect txtTon(Index)
End Sub

Private Sub txtTon_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0:
            If KeyAscii = 13 Then txtTon(0).Text = FrmVattu.ChonVattu(txtTon(0).Text)
            RFocus txtTon(0)
        Case 7, 8, 9, 10:
            If KeyAscii = 13 Then CmdCt_Click Else KeyProcess txtTon(Index), KeyAscii
    End Select
End Sub

Private Sub txtTon_LostFocus(Index As Integer)
    Dim sl As Double, dg As Double, tt As Double
    
    Select Case Index
        Case 0:
            vt.InitVattuSohieu txtTon(0).Text
            txtTon(1).Text = vt.TenVattu
            txtTon(2).Text = vt.DonVi
            If Me.tag = 0 Then txtTon(8).Text = CStr(vt.ThueNK)
        Case 4, 5:
            txtTon(Index).Text = Format(txtTon(Index).Text, Mask_2)
            sl = Cdbl5(txtTon(4).Text)
            dg = Cdbl5(txtTon(5).Text)
            tt = sl * dg
            txtTon(6).Text = Format(tt, Mask_2)
a:
            tt = Fix(0.5 + Cdbl5(txtTon(14).Text) * tt)
            txtTon(7).Text = Format(tt, Mask_0)
B:
            tt = Fix(0.5 + tt * CInt5(txtTon(8).Text) / 100)
            txtTon(9).Text = Format(tt, Mask_0)
        Case 6:
            tt = Cdbl5(txtTon(6).Text)
            sl = Cdbl5(txtTon(4).Text)
            If sl <> 0 Then txtTon(5).Text = Format(tt / sl, Mask_2)
            GoTo a
        Case 7:
            txtTon(Index).Text = Format(txtTon(Index).Text, Mask_0)
            tt = Cdbl5(txtTon(7).Text)
            GoTo B
        Case 8:
            txtTon(Index).Text = Format(txtTon(Index).Text, Mask_0)
            tt = Cdbl5(txtTon(7).Text)
            GoTo B
        Case 9, 10, 11, 12, 13:
            txtTon(Index).Text = Format(txtTon(Index).Text, Mask_0)
        Case 14:
            txtTon(Index).Text = Format(txtTon(Index).Text, Mask_2)
            TinhTien
    End Select
End Sub

Private Sub TinhTien()
    Dim i As Integer, cpx(1 To 4) As Double, tt As Double, tnk As Double, tt2 As Double, j As Integer
    Dim tygia As Double, sl As Double, sodong As Integer, cpa(1 To 4) As Double, cp1 As Double
    
    tygia = Cdbl5(txtTon(14).Text)
    If tygia = 0 Then tygia = 1
a:
    With Grd
        For i = 0 To .Rows - 1
            .col = 0
            .Row = i
            If .Text = "" Then Exit For
            .col = 5
            sl = sl + Cdbl5(.Text)
            sodong = sodong + 1
            .col = 6
            tt = Fix(0.5 + Cdbl5(.Text) * tygia)
            .col = 7
            If tt <> 0 Then .Text = Format(tt, Mask_0) Else tt = Cdbl5(.Text)
            '.col = 8
            'tnk = Fix(0.5 + tt * CInt5(.Text) / 100)
            '.col = 9
            '.Text = Format(tnk, Mask_0)
            .col = 9
            tnk = Cdbl5(.Text)
            For j = 1 To 4
                .col = 9 + j
                cpa(j) = Cdbl5(.Text)
                cpx(j) = cpx(j) + cpa(j)
            Next
            .col = 15
            ExecuteSQL5 "UPDATE CPGVHD SET CP1=" + DoiDau(cpa(1)) + ",CP2=" + DoiDau(cpa(2)) + ",CP3=" + DoiDau(cpa(3)) + ",CP4=" + DoiDau(cpa(4)) + ",TT=" + DoiDau(tt) + ",TNK=" + DoiDau(tnk) + ",TyGia1=" + DoiDau(tygia) + " WHERE MaSo=" + CStr(CLng5(.Text))
        Next
    End With
    
    For i = 1 To 4
        CP(i).Caption = Format(cpx(i), Mask_0)
    Next
    CP(0).Caption = Format(cpx(1) + cpx(2) + cpx(3) + cpx(4), Mask_0)
End Sub

Private Sub ChiPhiHD(des As Integer)
    Dim sql As String, i As Integer, tt2  As Double, tnk As Double, cpx As Double
    
    If Cbo(1).ListIndex < 0 Then Exit Sub
    
    sql = "SELECT CPGVHD.MaVattu, Vattu.SoHieu AS SHVT, Vattu.TenVattu, DonVi, ThueNK, SL, TT, TyGia1, PN, CP1, CP2, CP3, CP4, TT2, TNK " _
        & " FROM CPGVHD INNER JOIN Vattu ON CPGVHD.MaVattu=Vattu.MaSo " _
        & " WHERE CPGVHD.MaDT=" + CStr(Cbo(1).ItemData(Cbo(1).ListIndex)) + " ORDER BY Vattu.SoHieu"
    SetSQL "QLuyKe", sql
    
    SetRptInfo
    frmMain.Rpt.Destination = des
    frmMain.Rpt.Formulas(3) = "DienGiai='" + Cbo(1).Text + "'"
    frmMain.Rpt.Formulas(4) = "TyGia=" + DoiDau(Cdbl5(txtTon(14).Text))
    frmMain.Rpt.ReportFileName = "CPGVHD.RPT"
    InBaoCaoRPT
End Sub

Private Sub txtShTk_GotFocus(Index As Integer)
    AutoSelect txtShTk(Index)
End Sub

Private Sub txtShTk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdtk_Click Index
End Sub

Private Sub PhanBo(id As Integer)
    Dim cp1 As Double, sl As Double, sodong As Integer, tt As Double, tt2 As Double, cpx As Double, i As Integer
    
    cp1 = Cdbl5(CP(id).Caption)
    With Grd
        For i = 0 To .Rows - 1
            .col = 0
            .Row = i
            If .Text = "" Then Exit For
            .col = IIf(OptPB(0).Value, 5, 7)
            sl = sl + Cdbl5(.Text)
            sodong = sodong + 1
        Next
                
        If sl > 0 And sodong > 0 Then
            tt2 = 0
            For i = 0 To sodong - 2
                .Row = i
                .col = IIf(OptPB(0).Value, 5, 7)
                tt = Cdbl5(.Text)
                tt = Fix(0.5 + cp1 * tt / sl)
                tt2 = tt2 + tt
                .col = 9 + id
                .Text = Format(tt, Mask_0)
                .col = 15
                ExecuteSQL5 "UPDATE CPGVHD SET CP" + CStr(id) + "=" + DoiDau(tt) + " WHERE MaSo=" + CStr(CLng5(.Text))
            Next
            .Row = sodong - 1
            tt = cp1 - tt2
            .col = 9 + id
            .Text = Format(tt, Mask_0)
            .col = 15
            ExecuteSQL5 "UPDATE CPGVHD SET CP" + CStr(id) + "=" + DoiDau(tt) + " WHERE MaSo=" + CStr(CLng5(.Text))
        End If
    End With
End Sub
