VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UNESCO Accounting SERVER"
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   7095
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
   Icon            =   "frmServer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "UNESCO Accounting"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "11"
   WindowState     =   1  'Minimized
   Begin VB.Timer CTTimer 
      Interval        =   1
      Left            =   -120
      Top             =   360
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4605
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   476
      SimpleText      =   ""
      ShowTips        =   0   'False
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5583
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1905
            MinWidth        =   1905
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1905
            MinWidth        =   1905
            TextSave        =   "21/11/05"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmServer.frx":030A
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   9120
      TabIndex        =   38
      Top             =   960
      Width           =   2775
      Begin VB.CheckBox chk 
         Caption         =   "SX"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox chk 
         Caption         =   "CDT"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox chk 
         Caption         =   "XD"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox chk 
         Caption         =   "TMDV"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   10920
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.MDB"
      FileName        =   "*.MDB"
      Filter          =   "TÖp d÷ liÖu (*.MDB)|*.MDB|TÊt c¶ (*.*)|*.*"
   End
   Begin Crystal.CrystalReport Rpt 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   19
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   ", CÇn Th¬"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   19
      Left            =   3600
      TabIndex        =   49
      ToolTipText     =   "148/217 ®­êng 3/2 - P. H­ng Lîi - Q. Ninh KiÒu, Tel: 071.740028. Fax: 071.740028, Email: ketoancantho@yahoo.com"
      Top             =   2040
      Width           =   1110
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   ", §µ N½ng"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   18
      Left            =   2520
      TabIndex        =   48
      ToolTipText     =   "259/4 NguyÔn V¨n Linh - Q. Thanh Khª. Tel: 0511.653057, Fax: 0511.653057, Email: ketoandanang@yahoo.com"
      Top             =   2040
      Width           =   1140
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "TP HCM"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   17
      Left            =   1560
      TabIndex        =   47
      ToolTipText     =   "15/1 §ång Khëi - T©n B×nh. Tel: 08.8110556, Fax: 08.8113624, Email: ketoansaigon@yahoo.com"
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "C¸c VP§D:"
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
      Height          =   240
      Index           =   13
      Left            =   240
      TabIndex        =   46
      Tag             =   "Branch:"
      Top             =   2040
      Width           =   1155
   End
   Begin VB.Label email 
      AutoSize        =   -1  'True
      Caption         =   "www.ucdit.com"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   1560
      MouseIcon       =   "frmServer.frx":0326
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   1800
      Width           =   1515
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Website:"
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
      Height          =   240
      Index           =   2
      Left            =   480
      TabIndex        =   44
      Top             =   1800
      Width           =   930
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Index           =   12
      Left            =   4800
      TabIndex        =   26
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Trung t©m UNESCO Ph¸t triÓn C«ng nghÖ Th«ng tin"
      BeginProperty Font 
         Name            =   ".VnTime"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   26
      Left            =   240
      TabIndex        =   43
      Tag             =   "UNESCO Center for Development of Information Technology"
      Top             =   960
      Width           =   6675
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      Caption         =   "..."
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
      Height          =   240
      Index           =   2
      Left            =   9120
      TabIndex        =   37
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Height          =   240
      Index           =   1
      Left            =   9120
      TabIndex        =   36
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Index           =   0
      Left            =   9120
      TabIndex        =   35
      Tag             =   "0"
      Top             =   240
      Width           =   225
   End
   Begin VB.Label Label 
      Caption         =   "Ngµnh nghÒ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   7800
      TabIndex        =   34
      Tag             =   "Profession"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "LÜnh vùc H§"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   7800
      TabIndex        =   33
      Tag             =   "Activities"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "Lo¹i h×nh DN"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   7800
      TabIndex        =   32
      Tag             =   "Class"
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "Phiªn b¶n"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   7800
      TabIndex        =   31
      Tag             =   "Version"
      Top             =   360
      Width           =   975
   End
   Begin VB.Label email 
      AutoSize        =   -1  'True
      Caption         =   "unesco@hn.vnn.vn"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   4920
      MouseIcon       =   "frmServer.frx":0630
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   1560
      Width           =   1950
   End
   Begin VB.Label Label 
      Caption         =   "B¶n quyÒn ch­¬ng tr×nh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   28
      Tag             =   "Copyright by"
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   22
      Left            =   240
      TabIndex        =   29
      Tag             =   "UNESCO Center for Development of Information Technology"
      Top             =   960
      Width           =   105
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Ch­¬ng tr×nh kÕ to¸n UNESCO"
      BeginProperty Font 
         Name            =   ".VnTime"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Index           =   14
      Left            =   240
      TabIndex        =   27
      Tag             =   "UNESCO Accounting Software"
      Top             =   120
      Width           =   6555
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   2520
      TabIndex        =   25
      Tag             =   "1"
      Top             =   4635
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   11
      Left            =   5760
      TabIndex        =   24
      Tag             =   "1"
      Top             =   4635
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "QuËn, huyÖn"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   23
      Tag             =   "District"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "TØnh, thµnh phè"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   22
      Tag             =   "Province"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Hµ Néi: L« 7 B20 Nam Thµnh C«ng - §èng §a "
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
      Height          =   240
      Index           =   15
      Left            =   240
      TabIndex        =   21
      Tag             =   "HaNoi City: B20 Nam Thanh Cong - Ba Dinh District"
      Top             =   1320
      Width           =   4725
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Tel: 04.7763162    Fax: 04.7763163    Email:"
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
      Height          =   240
      Index           =   16
      Left            =   480
      TabIndex        =   20
      Top             =   1560
      Width           =   4380
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   5760
      TabIndex        =   19
      Tag             =   "1"
      Top             =   5025
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5160
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   17
      Tag             =   "1"
      Top             =   6075
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "M· sè thuÕ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   16
      Tag             =   "Tax Code"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   15
      Tag             =   "1"
      Top             =   6075
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   14
      Tag             =   "1"
      Top             =   5760
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   13
      Tag             =   "1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   12
      Tag             =   "1"
      Top             =   5025
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   11
      Tag             =   "1"
      Top             =   5025
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label LbCty 
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   10
      Tag             =   "1"
      Top             =   4275
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label 
      Caption         =   "N¨m tµi chÝnh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   9
      Tag             =   "Financial Year"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Tµi kho¶n Ngo¹i tÖ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Tag             =   "Bank F.C. Account"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label 
      Caption         =   "Tµi kho¶n tiÒn ViÖt"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Tag             =   "Bank VND Account"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label 
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "§iÖn tho¹i"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Tag             =   "Tel"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "§Þa chØ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Tag             =   "Address"
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LbCty 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   240
   End
   Begin VB.Label LbCty 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Tag             =   "0"
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "Ch­¬ng tr×nh ®­îc cµi ®Æt t¹i"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Tag             =   "This product is licensed to"
      Top             =   3240
      Width           =   2655
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   11280
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   15
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":093A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":0C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":0F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":1288
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":15A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":18BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":1BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":1EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":220A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":2524
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":283E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":2B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":2E72
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":318C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmServer.frx":34A6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ret
Dim m_nonClientMetrics As NONCLIENTMETRICS
Dim m_logFont As LOGFONT

Dim m_fontCaption As String * 32

Dim m_fontSmCaption As String * 32
Dim m_fontMenu As String * 32
Dim m_fontMessage As String * 32
Dim m_fontStatus As String * 32
Dim m_fontIcon As String * 32
Dim pProcessEnable As Boolean

Private Const MaxNamTC = 9
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function LDBUser_GetUsers Lib "MSLDBUSR.DLL" (lpszUserBuffer() As String, ByVal lpszFilename As String, ByVal nOptions As Long) As Integer

Private Const OptLDBLoggedUsers = &H2
Dim n1 As Date, n2 As Date

Private Sub CTTimer_Timer()
    If pProcessEnable Then
        pProcessEnable = False
        XuLyChungtu
        StationList
        pProcessEnable = True
        If Me.WindowState <> 1 Then Me.WindowState = 1
    End If
End Sub

Private Sub email_Click(Index As Integer)
    Select Case Index
        Case 0:
            ShellExecute hwnd, "open", "mailto:" + email(Index).Caption, vbNullString, vbNullString, 0
        Case 1:
            ShellExecute ByVal 0&, "open", "www.ucdit.com", vbNullString, vbNullString, 3
        Case 2:
            ShellExecute hwnd, "open", "ypager.exe", vbNullString, Left(pWinDir, 2) + "\Program Files\Yahoo!\Messenger", 1
    End Select
End Sub

Private Sub Form_Activate()
    StationList
End Sub

Private Sub Form_Load()
    Dim X1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer
    
    GetLicense
    
    If m_IgnoreEvents Then GoTo XuLy
    If Not WillRunAtStartup(App.EXEName) Then
        If MsgBox("ThiÕt lËp m¸y chñ ?", vbYesNo, App.ProductName) = vbYes Then
            If FrmGetStr.GetMK(LbCty(8).Caption) Then
                SetRunAtStartup App.EXEName, App.Path
                
                m_IgnoreEvents = True
                WillRunAtStartup (App.EXEName)
                
                m_IgnoreEvents = False
            Else
                CloseUp
                End
            End If
        Else
            CloseUp
            End
        End If
    End If
XuLy:
    HienThongBao pDataPath, 2
    dlgCommonDialog.InitDir = pCurDir + "DATA"
        
    
    On Error Resume Next
    Rpt.WindowShowPrintSetupBtn = True
    Rpt.WindowShowGroupTree = True
    Rpt.WindowShowSearchBtn = True
    Rpt.WindowShowZoomCtl = True
    On Error GoTo 0
    
    setMDSettings
                
    Mask_D = GetShortDateFormat
    
    pProcessEnable = True
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Screen.MousePointer = 11
    HienThongBao "KÕt thóc ch­¬ng tr×nh kÕ to¸n!", 1
    
    Do While Not pProcessEnable
        Idle
    Loop
    
    CTTimer.Enabled = False
    
    CloseUp
            
    Recycle pCurDir + "DATA\ZIP.MDB"
    SetPsw pDataPath, pPSW, ""
    On Error Resume Next
    DBEngine.CompactDatabase pDataPath, pCurDir + "DATA\ZIP.MDB"
    On Error GoTo 0
    If Len(Dir(pCurDir + "DATA\ZIP.MDB")) > 0 Then
        Recycle pDataPath
        FileCopy pCurDir + "DATA\ZIP.MDB", pDataPath
        SetPsw pCurDir + "DATA\ZIP.MDB", "", pPSW
    End If
    SetPsw pDataPath, "", pPSW
    
    restoreSettings
    Screen.MousePointer = 0
        
    End
    Set App = Nothing
End Sub

Private Sub sbStatusBar_PanelClick(ByVal Panel As ComctlLib.Panel)
    Select Case Panel.Index
        Case 2:             Panel.Text = IIf(Panel.Tag = 0, "Data File Size: " + Format(FileLen(pDataPath) / 1048576, Mask_2) + " MB, Version: " + IIf(DBKetoan.Version < 4, "97", "2000"), pDataPath)
                                    Panel.Tag = 1 - Panel.Tag
                                    StationList
    End Select
End Sub

'======================================================================================
' Function GetLicense : Thí tñc lÃy tËn càng ty v¡ chi nh¤nh
'======================================================================================
Private Sub GetLicense()
    Dim rs_license As Recordset, i As Integer, k As Integer, sh As String
    
    DEMO = 1
    i = er_DBFile
On Error Resume Next
    Set rs_license = DBKetoan.OpenRecordset("License", dbOpenSnapshot)
    i = 0
    i = rs_license!Flag1 Mod 100
    If i > 0 Then
        If StationList() > i Then
            i = er_Connection
            Unload Me
            Exit Sub
        End If
    End If
    
    pTenCty = rs_license!TenCty
    pTenCn = rs_license!tencn
    
    LbCty(2).Caption = rs_license!DiaChi
    LbCty(3).Caption = rs_license!Tel
    LbCty(4).Caption = rs_license!Fax
    LbCty(5).Caption = rs_license!TaiKhoanVN
    LbCty(6).Caption = rs_license!TaiKhoanNT
    pNamTC = rs_license!NamTC
    pThangDauKy = rs_license!thang
    LbCty(7).Caption = CStr(pNamTC)
    LbCty(8).Caption = rs_license!MaSoThue
    
    pBaoGia = (rs_license!Flag1 Mod 1000) \ 100
    pNVBH = (rs_license!Flag1 Mod 10000) \ 1000
        
    i = 0
    i = (rs_license!Flag1 Mod 1000000000) \ 100000000
    Lb(0).Tag = i
    If (i < 3 Or i = 5) And pVersion = 0 Then ExecuteSQL5 "UPDATE License SET Flag1=400000000+Flag1 Mod 100000000", False
    Select Case i
        Case 1: Lb(1).Caption = "Doanh nghiÖp Nhµ n­íc"
                        Lb(0).Caption = "4.8.NN."
        Case 2: Lb(1).Caption = "Cæ phÇn - Liªn doanh"
                        Lb(0).Caption = "4.8.CP."
        Case 3: Lb(1).Caption = "C«ng ty TNHH"
                        Lb(0).Caption = "4.8.HH."
        Case 4: Lb(1).Caption = "Doanh nghiÖp t­ nh©n"
                        Lb(0).Caption = "4.8.TN."
        Case 5: Lb(1).Caption = "C¬ së ®µo t¹o"
                        Lb(0).Caption = "4.8.DT."
        Case 6:
                        Lb(1).Caption = "Hµnh chÝnh sù nghiÖp"
                        Lb(0).Caption = "4.8.HCSN."
                        Label(24).Visible = False
                        Label(25).Visible = False
                        Frame(1).Visible = False
        Case Else
                        Lb(0).Caption = "2.04"
    End Select
    If pVersion <> 3 Then Lb(0).Caption = Lb(0).Caption + IIf((rs_license!Flag1 Mod 100000000) \ 10000000 > 0, "1", "0") + IIf((rs_license!Flag1 Mod 10000000) \ 1000000 > 0, "1", "0") + IIf((rs_license!Flag1 Mod 1000000) \ 100000 > 0, "1", "0") + IIf((rs_license!Flag1 Mod 100000) \ 10000 > 0, "1", "0")
    chk(0).Value = (rs_license!Flag1 Mod 100000000) \ 10000000
    chk(1).Value = (rs_license!Flag1 Mod 10000000) \ 1000000
    chk(2).Value = (rs_license!Flag1 Mod 1000000) \ 100000
    chk(3).Value = (rs_license!Flag1 Mod 100000) \ 10000
        
    Command(6).Visible = ((rs_license!Flag1 Mod 1000000) \ 100000 > 0)
    
    Command(4).Visible = (rs_license!Flag1 \ 1000000000 > 0)
            
    pTygia = IIf(rs_license!tygia > 0, 1, 0)
    pHachToan = 1 - (rs_license!RptOrder Mod 10)
    pRpt = (rs_license!RptOrder Mod 100) \ 10
    OutCost = rs_license!OutCost
    FCost = rs_license!FixedOutCost
    STDetail = rs_license!STDetail
    FADetail = rs_license!FADetail
    KHDetail = rs_license!HDV
    pGiaHT = rs_license!GiaHT
    pGiaVon = (rs_license!STDetail Mod 100) \ 10
    pDTTP = (rs_license!STDetail Mod 1000) \ 100
    pDinhmuc = (rs_license!STDetail Mod 10000) \ 1000
        
    Command(5).Visible = ((rs_license!Lock0 Mod 100) \ 10 > 0)
    pCongNoHD = (rs_license!Lock0 Mod 1000) \ 100
    pPQTK = (rs_license!Lock0 Mod 10000) \ 1000
    pGiaUSD = (rs_license!Lock0 Mod 100000) \ 10000
    pChietKhau = (rs_license!Lock1 Mod 100) \ 10
    pKiemKeNgay = (rs_license!Lock1 Mod 1000) \ 100
    pNoiBo = (rs_license!Lock1 Mod 10000) \ 1000
    pSoVV = (rs_license!Lock1 Mod 100000) \ 10000
    pNhapKhau = (rs_license!Lock2 Mod 100) \ 10
    pBarCode = (rs_license!Lock2 Mod 1000) \ 100
    pNhapDoiTuong = (rs_license!Lock2 Mod 10000) \ 1000
       
    pTien = 0
    pTien = rs_license!loaitien
    If pTien > 0 Then
        Mask_0 = Mask_2
        pTienStr = TenNT(pTien)
    Else
        Mask_0 = GetSetting(IniPath, "Environment", "IntMask", "###,###,###,###")
        pTienStr = "VND"
    End If
    CTGS_GV = rs_license!CTGS_GV
    pFontName = rs_license!FontName
    pFontSize = rs_license!FontSize
    LbCty(0).FontName = pFontName
    LbCty(1).FontName = pFontName
    LbCty(10).Caption = rs_license!Quan
    LbCty(11).Caption = rs_license!ThanhPho
    frmMain.LbCty(9).Caption = rs_license!email
    pSoKT = rs_license!SoKT
    
    pSongNgu = False
    pSongNgu = (pSoKT Mod 100000 >= 10000)
    pMaVach = 0
    pMaVach = rs_license!mv Mod 10
    pTyGiaBQ = 0
    pTyGiaBQ = IIf(rs_license!mv Mod 10000 >= 1000, 1, 0)
    
    DEMO = IIf((rs_license!mv Mod 100000 >= 10000) And (rs_license!MKUP = pRev), 0, 1)
    NgayDauThangMoi = rs_license!NgayDauThang
    FontFlag = LoaiFont(pFontName)

    i = pNN
    pNN = 0
    SetFont Me
    pNN = i
        
    sh = SelectSQL("SELECT App1Path AS F1 FROM License")
    Command(3).Visible = Len(Dir(sh)) > 0
        
    If DEMO = 0 And pVersion <> 2 Then
        If ((Int_StrToCode(rs_license!MaSoThue) <> rs_license!MST_ID) Or (Int_StrToCode(pTenCty) <> rs_license!TenCty_ID) Or (Int_StrToCode(pTenCn) <> rs_license!tencn_id)) Then
            pTenCty = ABCtoVNI("Sao chÐp kh«ng b¶n quyÒn")
            pTenCn = ABCtoVNI("Sao chÐp kh«ng b¶n quyÒn")
            ExecuteSQL5 "UPDATE License SET MST_ID=-1"
            pSTOP = 1
        End If
    End If
    
    If pVersion = 3 Then
        pVATV = "3113"
        pSHPT = "3111"
    Else
        pVATV = "133"
        pSHPT = "131"
    End If
    
    LbCty(0).Tag = rs_license!TenCty_ID
    LbCty(0).Caption = pTenCty
    LbCty(1).Caption = pTenCn
    Frame(0).Visible = pSongNgu
                
    rs_license.Close
    Set rs_license = Nothing
    
    On Error GoTo 0
End Sub

Private Sub NoRight(id As Integer)
    Select Case id
        Case 0: HienThongBao "Kh«ng cã quyÒn truy cËp!", 1
        Case 1: HienThongBao "Kh«ng ®¨ng ký theo dâi chi tiÕt vËt t­!", 1
        Case 2: HienThongBao "Kh«ng ®¨ng ký theo dâi chi tiÕt TSC§!", 1
    End Select
    Beep
End Sub

Public Function ChonTenTep(title As String, f As Long, mask As String, act As Integer) As String
    With dlgCommonDialog
        .DialogTitle = title
        .Flags = f
        .Filename = mask
        .DefaultExt = mask
        .Filter = "TÖp d÷ liÖu (" + mask + ")|" + mask + "|TÊt c¶ (*.*)|*.*"
        On Error GoTo Xong
        Select Case act
            Case 1:            .ShowOpen
            Case 2:            .ShowSave
            Case 3:             .ShowPrinter
            Case 4:             .ShowFont
        End Select
        On Error GoTo 0
        If Len(.Filename) = 0 Or Left(.Filename, 1) = "*" Then GoTo Xong
        
        If act = 2 Then
            If Len(Dir(.Filename)) > 0 Then
                If .Filename = pDataPath Then
                    MsgBox "TÖp d÷ liÖu ®ang më !", vbCritical, App.ProductName
                    GoTo Xong
                End If
                If MsgBox("TÖp " + .Filename + " ®· tån t¹i, tiÕp tôc ? !", vbQuestion + vbYesNo, App.ProductName) = vbNo Then GoTo Xong
                If Recycle(.Filename) <> 0 Then
                    MsgBox "Kh«ng xo¸ ®­îc tÖp " + dlgCommonDialog.Filename + " !", vbExclamation, App.ProductName
                    GoTo Xong
                End If
            End If
        End If
        ChonTenTep = .Filename
    End With
Xong:
End Function

Private Function StationList() As Integer
    ' It is important that ReDim be used to define the array as the DLL,
' because the DLL depends on being able to redimension the array.
    ReDim msString(1) As String
' The array is 1-based rather than 0-based, regardless if Option Base 1
' is specified in the declarations section.
    Dim miLoop As Integer, i As Integer, LDBName As String, sql As String, U As String, X As String
 
    LDBName = Left(pDataPath, Len(pDataPath) - 3) + "LDB"
    miLoop = LDBUser_GetUsers(msString, LDBName, OptLDBLoggedUsers)
' The function calls cannot be combined and must be used individually.
' Get the first user in the selected .LDB file.
    For i = 0 To miLoop - 1
        If i >= LBound(msString, 1) And i <= UBound(msString, 1) Then
            U = SelectSQL("SELECT TenNSD AS F1 FROM Users WHERE WS='" + msString(i) + "' AND TenNSD<>'" + X + "'")
            If U <> "0" Then
                sql = sql + Chr(13) + msString(i) + " : " + U
                X = U
            End If
        End If
    Next
    If miLoop > 1 Then
        LbCty(12).Caption = "C¸c m¸y tr¹m: " + sql
    Else
        LbCty(12).Caption = ""
    End If
    
    StationList = miLoop
End Function

Private Sub XoaQuery()
    Dim q As String
    
    q = InputBox("Tªn query cÇn xo¸: ", App.ProductName)
    If Len(q) > 0 Then
        If QueryDaCo(q) Then DBKetoan.QueryDefs.Delete q
    End If
End Sub

Private Sub FontSetUp()
    Add32Font "VNTIME.TTF"
    Add32Font "VNTIMEB.TTF"
    Add32Font "VNTIMEBI.TTF"
    Add32Font "VNTIMEI.TTF"
    
    Add32Font "VHTIME.TTF"
    Add32Font "VHTIMEB.TTF"
    Add32Font "VHTIMEBI.TTF"
    Add32Font "VHTIMEI.TTF"
    
    Add32Font "VTIMESN.TTF"
    Add32Font "VTIMESB.TTF"
    Add32Font "VTIMESBI.TTF"
    Add32Font "VTIMESI.TTF"
End Sub

Public Sub SetUserRight()

End Sub

Private Sub XuLyChungtu()
    Dim rs_chungtu As Recordset
    Dim chungtu As New ClsChungtu
    
    Set rs_chungtu = DBKetoan.OpenRecordset("SELECT * FROM ChungTu WHERE XuLy=2 ORDER BY MaCT, MaSo", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_chungtu.EOF
        chungtu.InitChungtuRS rs_chungtu
        chungtu.XuLyCTXoa
        rs_chungtu.MoveNext
    Loop
    rs_chungtu.Close
    
    Set rs_chungtu = DBKetoan.OpenRecordset("SELECT * FROM ChungTu WHERE XuLy=1 ORDER BY MaCT, MaSo", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs_chungtu.EOF
        chungtu.InitChungtuRS rs_chungtu
        chungtu.XuLyCT
        rs_chungtu.MoveNext
    Loop
    rs_chungtu.Close
    
    Set rs_chungtu = Nothing
    Set chungtu = Nothing
    
    HienThongBao "", 1
End Sub

