VERSION 5.00
Begin VB.Form FrmOptions 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Th«ng tin doanh nghiÖp"
   ClientHeight    =   7845
   ClientLeft      =   660
   ClientTop       =   915
   ClientWidth     =   10320
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFC0&
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFC0&
   Icon            =   "Frmopt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Frmopt.frx":57E2
   ScaleHeight     =   7845
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Tag             =   "Options"
   Begin VB.CheckBox Chbanthuoc 
      BackColor       =   &H00FFFFC0&
      Caption         =   "§Æt thï ngµnh d­îc"
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
      Left            =   120
      TabIndex        =   105
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton active 
      Caption         =   "Active"
      Height          =   375
      Left            =   9120
      TabIndex        =   104
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   4
      Left            =   7440
      MaxLength       =   300
      TabIndex        =   103
      Text            =   "0"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text 
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
      Index           =   27
      Left            =   4080
      MaxLength       =   30
      TabIndex        =   92
      Text            =   "..."
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame 
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
      Height          =   4215
      Index           =   0
      Left            =   0
      TabIndex        =   56
      Top             =   120
      Width           =   9075
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Frame1"
         Height          =   1935
         Left            =   3840
         TabIndex        =   115
         Top             =   2160
         Width           =   5055
         Begin VB.TextBox Text2 
            Height          =   360
            Left            =   2880
            TabIndex        =   119
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Height          =   360
            Left            =   480
            TabIndex        =   118
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Theo n¨m"
            Height          =   240
            Left            =   360
            TabIndex        =   117
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "§¨ng ký vÜnh viÔn"
            Height          =   240
            Left            =   360
            TabIndex        =   116
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Sè l­îng chøng tõ"
            Height          =   255
            Left            =   2880
            TabIndex        =   120
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.ComboBox Combo 
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
         ItemData        =   "Frmopt.frx":AFC4
         Left            =   9120
         List            =   "Frmopt.frx":AFDC
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text 
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
         Index           =   28
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   112
         Text            =   "..."
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text 
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   110
         Text            =   "2"
         Top             =   3840
         Width           =   255
      End
      Begin VB.TextBox Text 
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
         Index           =   26
         Left            =   1800
         MaxLength       =   500
         TabIndex        =   102
         Text            =   "..."
         Top             =   480
         Visible         =   0   'False
         Width           =   7215
      End
      Begin VB.CheckBox Check 
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
         Index           =   28
         Left            =   2280
         TabIndex        =   88
         Top             =   3840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "B¸o c¸o néi bé"
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
         Index           =   55
         Left            =   9120
         TabIndex        =   95
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text 
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
         Left            =   9360
         MaxLength       =   20
         TabIndex        =   16
         Text            =   "..."
         Top             =   1080
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "CDT"
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
         Index           =   27
         Left            =   9120
         TabIndex        =   15
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox Check 
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
         Index           =   26
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   14
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check 
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
         Index           =   25
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   13
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text 
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
         Index           =   17
         Left            =   9120
         MaxLength       =   30
         TabIndex        =   23
         Text            =   "..."
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text 
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
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   20
         Text            =   "2"
         Top             =   3360
         Width           =   255
      End
      Begin VB.ComboBox Combo 
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
         Index           =   1
         ItemData        =   "Frmopt.frx":B040
         Left            =   9120
         List            =   "Frmopt.frx":B06B
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text 
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
         Index           =   18
         Left            =   3360
         MaxLength       =   1
         TabIndex        =   22
         Text            =   "2"
         Top             =   3360
         Width           =   255
      End
      Begin VB.TextBox Text 
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
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   21
         Text            =   "2"
         Top             =   3360
         Width           =   255
      End
      Begin VB.TextBox Text 
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
         Index           =   23
         Left            =   7680
         MaxLength       =   3
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text 
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
         HideSelection   =   0   'False
         Index           =   0
         Left            =   1800
         MaxLength       =   500
         TabIndex        =   0
         Text            =   "..."
         Top             =   240
         Width           =   7215
      End
      Begin VB.TextBox Text 
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
         Left            =   8520
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "..."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
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
         Left            =   1800
         MaxLength       =   500
         TabIndex        =   3
         Text            =   "..."
         Top             =   650
         Width           =   7215
      End
      Begin VB.TextBox Text 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "0"
         Top             =   1395
         Width           =   1695
      End
      Begin VB.TextBox Text 
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
         Left            =   1800
         MaxLength       =   5000
         TabIndex        =   9
         Text            =   "..."
         Top             =   1800
         Width           =   7215
      End
      Begin VB.TextBox Text 
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
         Index           =   6
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "..."
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text 
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
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   11
         Text            =   "..."
         Top             =   1020
         Width           =   1695
      End
      Begin VB.ComboBox Combo 
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
         Index           =   0
         ItemData        =   "Frmopt.frx":B096
         Left            =   4080
         List            =   "Frmopt.frx":B098
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1005
         Width           =   1215
      End
      Begin VB.TextBox Text 
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
         Index           =   14
         Left            =   9360
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "0"
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text 
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
         Index           =   16
         Left            =   7560
         MaxLength       =   30
         TabIndex        =   8
         Text            =   "..."
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text 
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
         Index           =   15
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "..."
         Top             =   1390
         Width           =   3015
      End
      Begin VB.TextBox Text 
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
         Index           =   19
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "..."
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text 
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
         Index           =   20
         Left            =   1800
         MaxLength       =   500
         TabIndex        =   5
         Text            =   "..."
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sè ch÷ sè thËp ph©n"
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
         TabIndex        =   111
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "LÜnh vùc H§"
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
         Left            =   8880
         TabIndex        =   85
         Tag             =   "Activities"
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lo¹i h×nh DN"
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
         Left            =   8880
         TabIndex        =   84
         Tag             =   "Class"
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "H¹ch to¸n theo:"
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
         Left            =   75
         TabIndex        =   78
         Tag             =   "Send data to default addr"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "chi"
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
         Left            =   2160
         TabIndex        =   77
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "UNC"
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
         Left            =   2880
         TabIndex        =   76
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sè lÇn in mçi phiÕu thu"
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
         Left            =   80
         TabIndex        =   75
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tªn c«ng ty: "
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
         Left            =   500
         TabIndex        =   68
         Tag             =   "Company"
         Top             =   250
         Width           =   975
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "§Þa chØ:"
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
         Left            =   500
         TabIndex        =   67
         Tag             =   "Address"
         Top             =   650
         Width           =   855
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "§iÖn tho¹i:"
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
         Left            =   500
         TabIndex        =   66
         Tag             =   "Tel"
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Fax"
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
         Left            =   3720
         TabIndex        =   65
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tµi kho¶n VN§:"
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
         Left            =   500
         TabIndex        =   64
         Tag             =   "Bank VND Account"
         Top             =   1850
         Width           =   1215
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sè ®Þa ®iÓm kinh doanh:"
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
         Left            =   80
         TabIndex        =   63
         Tag             =   "Bank F.C. Account"
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "M· sè thuÕ:"
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
         Left            =   500
         TabIndex        =   62
         Tag             =   "Tax Code"
         Top             =   1065
         Width           =   975
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "N¨m"
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
         Left            =   3720
         TabIndex        =   61
         Tag             =   "Year"
         Top             =   1050
         Width           =   375
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Lo¹i h×nh ho¹t ®éng:"
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
         Left            =   75
         TabIndex        =   60
         Tag             =   "From month"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ngµy ®Çu th¸ng"
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
         Index           =   17
         Left            =   9240
         TabIndex        =   59
         Tag             =   "Month from Date"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Email"
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
         Left            =   5520
         TabIndex        =   58
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SMTP"
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
         Left            =   6960
         TabIndex        =   57
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame 
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
      Height          =   3465
      Index           =   1
      Left            =   0
      TabIndex        =   69
      Top             =   4320
      Width           =   9050
      Begin VB.CheckBox ChkVT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Luü kÕ theo ngµy chØ kª vËt t­ cã ph¸t sinh"
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
         TabIndex        =   114
         Top             =   3000
         Width           =   3735
      End
      Begin VB.CheckBox ChkVT 
         BackColor       =   &H00FFFFC0&
         Caption         =   "KiÓm kª tån kho theo ngµy"
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
         TabIndex        =   113
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "NhËt ký chung"
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
         Left            =   4440
         TabIndex        =   108
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "CT ghi sæ"
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
         Left            =   4440
         TabIndex        =   107
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "NK chøng tõ"
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
         Left            =   4440
         TabIndex        =   106
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cho nhËp trïng sè hiÖu chøng tõ kh¸c th¸ng"
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
         Index           =   38
         Left            =   5520
         TabIndex        =   101
         Top             =   5280
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "NhËp theo tªn"
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
         Index           =   37
         Left            =   1680
         TabIndex        =   100
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "TÝnh gi¸ vèn hµng nhËp khÈu"
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
         Left            =   5520
         TabIndex        =   98
         Top             =   5040
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox Text 
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
         Index           =   25
         Left            =   9360
         MaxLength       =   2
         TabIndex        =   97
         Text            =   "0"
         Top             =   4200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sè th«ng tin chøng tõ bæ sung cÇn theo dâi"
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
         Index           =   35
         Left            =   5520
         TabIndex        =   96
         Top             =   4800
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ChiÕt khÊu ®Çu ra theo mÆt hµng"
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
         Index           =   34
         Left            =   120
         TabIndex        =   93
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "§¬n gi¸ hµng ho¸ vµ ho¸ ®¬n sö dông USD"
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
         Index           =   33
         Left            =   120
         TabIndex        =   91
         Top             =   4920
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ph©n quyÒn theo tµi kho¶n"
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
         Left            =   600
         TabIndex        =   90
         Top             =   3960
         Width           =   2415
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "C«ng nî theo ho¸ ®¬n"
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
         Left            =   5520
         TabIndex        =   89
         Top             =   3600
         Width           =   3015
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sö dông c¸c b¸o c¸o qu¶n trÞ"
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
         Index           =   30
         Left            =   5520
         TabIndex        =   87
         Top             =   4560
         Width           =   3015
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "T¸ch chøc n¨ng in phiÕu"
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
         Index           =   29
         Left            =   5520
         TabIndex        =   86
         Top             =   4320
         Width           =   3015
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Theo dâi nh©n viªn b¸n hµng"
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
         Left            =   5520
         TabIndex        =   83
         Top             =   4080
         Width           =   3015
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sö dung chøc n¨ng in b¸o gi¸"
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
         Index           =   23
         Left            =   120
         TabIndex        =   82
         Top             =   4680
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox Text 
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
         Left            =   9120
         MaxLength       =   20
         TabIndex        =   81
         Text            =   "..."
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Gi¸ thµnh s¶n xuÊt"
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
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tû gi¸ b×nh qu©n"
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
         Left            =   120
         TabIndex        =   32
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox Text 
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
         Index           =   24
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   79
         Tag             =   "0"
         Text            =   "0"
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cho sö dông chøc n¨ng tæng hîp sè liÖu"
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
         Left            =   5520
         TabIndex        =   74
         Top             =   3840
         Width           =   3615
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cho phÐp ®iÒu chØnh tªn chi nh¸nh"
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
         TabIndex        =   73
         Top             =   4440
         Width           =   3015
      End
      Begin VB.TextBox Text 
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
         Index           =   22
         Left            =   3960
         MaxLength       =   20
         TabIndex        =   51
         Text            =   "8.0"
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox Text 
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
         Index           =   21
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   50
         Text            =   "0"
         Top             =   4080
         Width           =   975
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Theo dâi tû gi¸ tõng chøng tõ víi tû gi¸ ®Çu n¨m lµ"
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
         TabIndex        =   49
         Top             =   4200
         Width           =   3975
      End
      Begin VB.ComboBox Combo 
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
         ItemData        =   "Frmopt.frx":B09A
         Left            =   3960
         List            =   "Frmopt.frx":B09C
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ph­¬ng ph¸p tÝnh gi¸ xuÊt kho"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Index           =   2
         Left            =   4320
         TabIndex        =   70
         Top             =   240
         Width           =   3975
         Begin VB.CheckBox ChkVT 
            BackColor       =   &H00FFFFC0&
            Caption         =   "In barcode"
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
            Left            =   6120
            TabIndex        =   99
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox ChkVT 
            BackColor       =   &H00FFFFC0&
            Caption         =   "KiÓm kª tån kho theo ngµy"
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
            TabIndex        =   94
            Top             =   1440
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.OptionButton OptVT 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Gi¸ b×nh qu©n"
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
            TabIndex        =   34
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptVT 
            BackColor       =   &H00FFFFC0&
            Caption         =   "XuÊt ®Ých danh"
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
            TabIndex        =   35
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton OptVT 
            BackColor       =   &H00FFFFC0&
            Caption         =   "FIFO"
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
            TabIndex        =   36
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton OptVT 
            BackColor       =   &H00FFFFC0&
            Caption         =   "LIFO"
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
            TabIndex        =   37
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox ChkVT 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Cè ®Þnh gi¸ xuÊt"
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
            TabIndex        =   38
            Top             =   240
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox ChkVT 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Sö dông gi¸ HT"
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
            TabIndex        =   39
            Top             =   480
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox ChkVT 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Luü kÕ theo ngµy chØ kª vËt t­ cã ph¸t sinh"
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
            TabIndex        =   40
            Top             =   1680
            Visible         =   0   'False
            Width           =   3735
         End
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Theo dâi chi tiÕt vËt t­, hµng ho¸"
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
         TabIndex        =   24
         Top             =   120
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Theo dâi chi tiÕt tµi s¶n cè ®Þnh"
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
         TabIndex        =   25
         Top             =   360
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Theo dâi chi tiÕt c«ng nî"
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
         TabIndex        =   26
         Top             =   600
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "H¹ch to¸n kÐp"
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
         TabIndex        =   27
         Top             =   840
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tù ®éng xuÊt gi¸ vèn hµng b¸n"
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
         TabIndex        =   28
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "TËp hîp gi¸ thµnh theo ®èi t­îng"
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
         TabIndex        =   30
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Trõ lïi thuÕ GTGT trªn ho¸ ®¬n "
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
         TabIndex        =   42
         Top             =   2280
         Width           =   3615
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "In chi tiÕt mÆt hµng trªn b¶ng kª"
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
         Left            =   2880
         TabIndex        =   43
         Top             =   5040
         Width           =   2655
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "KiÓm tra tû lÖ thuÕ c¸c mÆt hµng cïng ho¸ ®¬n"
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
         Left            =   2880
         TabIndex        =   44
         Top             =   5280
         Width           =   3735
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "In b¸o c¸o thuÕ cã m· v¹ch"
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
         TabIndex        =   41
         Top             =   5400
         Width           =   2415
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "§Þnh møc thµnh phÈm theo th¸ng"
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
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Song ng÷"
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
         Left            =   3240
         TabIndex        =   52
         Top             =   4560
         Width           =   1095
      End
      Begin VB.ComboBox CTGS 
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
         ItemData        =   "Frmopt.frx":B09E
         Left            =   2880
         List            =   "Frmopt.frx":B0C6
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Qu¶n lý quyÒn xem tõng b¸o c¸o"
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
         TabIndex        =   45
         Top             =   3720
         Width           =   2775
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cho sö dông chøc n¨ng ®æi font"
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
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   2775
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ch­¬ng tr×nh söa ®æi theo doanh nghiÖp"
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
         Index           =   17
         Left            =   120
         TabIndex        =   47
         Top             =   3960
         Width           =   3255
      End
      Begin VB.TextBox Text 
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   48
         Text            =   "..."
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sö dông c¸c sæ:"
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
         Left            =   4440
         TabIndex        =   109
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sè m¸y truy cËp tèi ®a trªn m¹ng"
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
         Left            =   600
         TabIndex        =   80
         Top             =   3720
         Visible         =   0   'False
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rev."
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
         Left            =   3120
         TabIndex        =   72
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label 
         BackColor       =   &H00E0E0E0&
         Caption         =   "H¹ch to¸n b»ng"
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
         Left            =   3960
         TabIndex        =   71
         Top             =   4800
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command 
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
      Height          =   375
      Index           =   1
      Left            =   9120
      Picture         =   "Frmopt.frx":B0F1
      Style           =   1  'Graphical
      TabIndex        =   55
      Tag             =   "&Return"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command 
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
      Height          =   375
      Index           =   0
      Left            =   9120
      Picture         =   "Frmopt.frx":C513
      Style           =   1  'Graphical
      TabIndex        =   54
      Tag             =   "&Save"
      Top             =   5880
      Width           =   1095
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ttVT As Integer
Dim mst As String
Dim suatencn As Integer
Dim kb As Integer
Dim typeRegistry As Integer

Private Sub active_Click()

    Dim st As String
    Dim rs As Recordset
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW License.* FROM License", dbOpenSnapshot)
    If FrmGetStr.GetMK(Text(7).Text) Then
        st = rs!CMP
        ExecuteSQL5 "UPDATE license SET CMG = " + str(Int_StrToCodes(st)) + ",namcode = " + str(Int_StrToCodes(str(rs!nam)))
        frmMain.txtdungthu.Caption = ""
        url_helper.Thong_tin Text(7).Text, Text(0) + " - " + Text(2).Text + " - " + Text(3).Text + " - " + Text(15).Text
        If Option1.Value = True Then
            MsgBox "Dang ky vinh vien thanh cong"
            ExecuteSQL5 "Update tbLicensekey set Type=1,Year=0,Totals=0"
        Else
            MsgBox "Dang ky " & Text1.Text & " nam thanh cong"
            ExecuteSQL5 "UPDATE tbLicensekey SET Type = 2, Year = '" & Text1.Text & "|" & pNamTC & "', Totals = '" & Text2.Text & "'"
        End If

    End If
    Unload Me
End Sub

Private Sub Check_Click(Index As Integer)
    Select Case Index
        Case 17:
            If Check(Index).Value = 0 Then Text(8).Text = "..."
        Case 18:
            If Check(18).Value = 0 Then
                Check(33).Enabled = False
                Check(33).Value = 0
            Else
                Check(33).Enabled = True
            End If
        Case 25, 26, 27, 28:
            PhanChucNang Combo(3).ListIndex + 1, Check(25).Value, Check(26).Value, Check(27).Value, Check(28).Value
    End Select
End Sub

Private Sub ChkVT_Click(Index As Integer)
    If Index = 4 And ChkVT(4).Value = 1 Then
        If Len(Dir(pCurDir + "DOWNLOAD.EXE")) = 0 Then
            MsgBox "Ch­a cµi ®Æt ch­¬ng tr×nh ®äc m· v¹ch!", vbCritical, App.ProductName
            ChkVT(4).Value = 0
        End If
    End If
End Sub

Private Sub Combo_Click(Index As Integer)
    If Index = 3 Then PhanChucNang Combo(Index).ListIndex + 1, Check(25).Value, Check(26).Value, Check(27).Value, Check(28).Value
End Sub

Private Sub Form_Activate()
    Option1.Value = True
    ' Combo(0).Enabled = False
    If (SelectSQL("select count(*) as f1 from chungtu") > 0) Then Combo(0).Enabled = False
    ActiveInfo

    Dim Types As Integer
    Types = SelectSQL("select Type AS f1 from  tbLicensekey")
    If Types <> 0 Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    End If

End Sub
Private Sub ActiveInfo()
    Dim rs_ktra As Recordset
    Dim Query As String
    Dim rst As String
    Query = "SELECT *  FROM tbLicensekey "
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? các b?n ghi
        Do While Not rs_ktra.EOF
            If rs_ktra!Type = 1 Then
                Option1.Value = True
                Option1.Enabled = False
                Option2.Enabled = False
                Text1.Enabled = False
                Text2.Enabled = False
                Text2.Text = rs_ktra!Totals
            End If
            If rs_ktra!Type = 2 Then
                Option2.Value = True
                Option1.Enabled = False
                Option2.Enabled = False
                Dim resultArray() As String
                resultArray = Split(rs_ktra!Year, "|")

                Text1.Text = resultArray(0)
                Text2.Text = rs_ktra!Totals
                Text2.Enabled = False
                Text1.Enabled = False
            End If
            rs_ktra.MoveNext

        Loop
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyG:
                RFocus Command(0)
                Command_Click 0
          
            Case vbKeyV:
                RFocus Command(1)
                Command_Click 1
        End Select
    End If
    
    If (Shift And vbAltMask) > 0 And (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyN Then
        kb = 1
        HienNoiBo
    End If
    
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim vis As Boolean
    
    ' LÊy l¹i c¸c gi¸ trÞ mÆc ®Þnh
    Int_RecsetToCbo "SELECT MaSo As F2,SoHieu+ ' - '+DienGiai As F1 FROM CTGhiSo ORDER BY SoHieu", CTGS
    
    SetFont Me
    Set Combo(3).Font = Me.Font
    mst = frmMain.LbCty(8).Caption
    If IsNumeric(mst) Then
        vis = (Cdbl5(mst) = 0)
    Else
        vis = False
    End If
    
    Frame(1).Enabled = vis
    Combo(3).Enabled = vis
    Check(25).Enabled = vis
    Check(26).Enabled = vis
    Check(27).Enabled = vis
    Check(28).Enabled = vis

    Int_RecsetToCbo "SELECT MaSo As F2, KyHieu As F1 FROM NguyenTe ORDER BY KyHieu", Combo(2)
    Combo(2).AddItem "VND", 0
    Combo(2).ItemData(0) = 0
    SetListIndex Combo(2), pTien
    
    LoadInfo
  
   Frame(1).Enabled = True
   Frame(2).Enabled = False
   If Len(Text(7).Text) >= 9 Then Chbanthuoc.Enabled = False
   
End Sub

Private Sub Option1_Click()
    typeRegistry = 1
End Sub

Private Sub Option2_Click()
    typeRegistry = 2
End Sub

Private Sub OptVT_Click(Index As Integer)
    ttVT = Index
End Sub

Private Sub Command_Click(Index As Integer)
    
    Dim i As Integer, tygia As Double, T As Long, mk As Long, Fx As Long, pctpath As String, F0 As Integer, f1 As Integer, k As Integer, F2 As Integer
    Dim KiemTra
       ExecuteSQL5_Themmoi ("ALTER TABLE license  ADD COLUMN banthuoc Number")
         ExecuteSQL5_Themmoi ("ALTER TABLE license  ADD tenhoadon text")
         ExecuteSQL5 ("ALTER TABLE license ALTER COLUMN TaiKhoanVN TEXT(255)")
          ExecuteSQL5 ("ALTER TABLE license ALTER COLUMN DiaChi TEXT(255)")
         ' ExecuteSQL5 ("ALTER TABLE license add COLUMN sofax TEXT(255)")
         ExecuteSQL5_Themmoi ("ALTER TABLE license add COLUMN sofax TEXT(255)")
       ' them moi sau
        ExecuteSQL5_Themmoi ("ALTER TABLE license ADD Lock13 TEXT(255)")
        ExecuteSQL5_Themmoi ("ALTER TABLE license ADD CMP TEXT(255)")
        ExecuteSQL5 ("ALTER TABLE license ALTER COLUMN CMP TEXT(255)")
         ExecuteSQL5_Themmoi ("ALTER TABLE license add COLUMN CMG Number")
         ExecuteSQL5_Themmoi ("ALTER TABLE license  ADD nam Number")
         ExecuteSQL5_Themmoi ("ALTER TABLE license  ADD namCode Number")
         ExecuteSQL5_Themmoi ("ALTER TABLE license ADD sodong Number")
         ExecuteSQL5_Themmoi ("ALTER TABLE license ADD sodongId Number")
     ' them moi sau
     
     
    If Index = 0 Then ' neu la nut ghi, nguoc lai thoat
    'neu test dung thi bo
    If (CInt(SelectSQL("SELECT nam as f1 from license ")) <= 0) Then ExecuteSQL5 "UPDATE license SET nam = " + str(pNamTC)
    
        For i = 0 To Text.count - 1
            Text_LostFocus i
        Next
        'If Not KiemTraMaSoThue(Text(7).Text, pTaxCode, 1) Then
        '    RFocus Text(7)
        '    GoTo KT
        'End If
     
        If IsNumeric(Left(App.LegalCopyright, 10)) And Len(App.LegalCopyright) >= 10 Then
            If Left(Text(7).Text, 10) <> Left(App.LegalCopyright, 10) Then GoTo KT
        End If
        If Combo(3).ListIndex < 0 Then Combo(3).ListIndex = Combo(3).ListCount - 1
          Dim Tr As String
               'Tr = Int_StrToCode(Text(0).Text)
              ' Me.Caption = "Th«ng tin ch­¬ng tr×nh   -  " + Tr
        If (Combo(3).ListIndex < 2 Or Combo(3).ListIndex > 4) And pVersion = 0 Then
            ErrMsg er_Version
            GoTo KT
        End If
        If Check(35).Value = 1 Then
            i = CInt5(Text(25).Text)
            If i < 1 Or i > 3 Then Text(25).Text = ""
        End If
        
      '  MsgBox FrmGetStr.GetMK(Text(7).Text)
        
        If Combo(2).ListIndex >= 0 Then T = Combo(2).ItemData(Combo(2).ListIndex) Else T = pTien
        If CInt5(Left(Text(Index).Text, 2)) <> 0 Then Check(55).Value = 0
        If ((((pTenCty = Text(0).Text And (pTenCn = Text(1).Text Or suatencn = 1) And (Check(19).Value = suatencn) And pMaVach = Check(9).Value And pDinhmuc = Check(13).Value And pSongNgu = (Check(14).Value = 1) And pRpt = Check(15).Value And pTygia = Check(18).Value And T = pTien And mk = 0) Or (DEMO = 1 And CLng5(Left(Text(7).Text, 2)) > 0)) And (mst = Text(7).Text Or (suatencn = 1 And Left(mst, 10) = Left(Text(7).Text, 10)))) Or Combo(3).ListIndex = 4 Or (Cdbl5(Left(Text(7).Text, 10)) = 0 And Cdbl5(Left(frmMain.LbCty(8).Caption, 10)) = 0)) And (pNoiBo = Check(55).Value) And (CInt5(Combo(0).Text) = pNamTC) Then GoTo a
        If (Len(pMST) > 0 And Left(Text(7).Text, Len(pMST)) = pMST) Then GoTo a
        If boolean_kiemtra() = False Then GoTo a ' kiem tra da active thi bat khung nhap ma so le
        If FrmGetStr.GetMK(Text(7).Text) Then
        frmMain.txtdungthu.Caption = ""
a:
            If ttVT <> OutCost And SelectSQL("SELECT TOP 1 MaCT AS F1 FROM ChungTu WHERE MaLoai=2 OR MaLoai=4") > 0 And ttVT <> 0 Then
                If MsgBox("§· cã chøng tõ xuÊt kho, thay ®æi ph­¬ng ph¸p tÝnh gi¸ xuÊt ?", vbCritical + vbYesNo, App.ProductName) = vbNo Then GoTo KT
            End If
            If Combo(2).ListIndex >= 0 Then pTien = Combo(2).ItemData(Combo(2).ListIndex)
            pMaVach = Check(9).Value + IIf(Check(19).Value = 1, 10, 0) + IIf(Check(20).Value = 1, 100, 0) + IIf(Check(21).Value = 1, 1000, 0) + IIf(DEMO = 0, 10000, 0)
            pSoKT = IIf(Check(10).Value = 1, 1, 0) + IIf(Check(11).Value = 1, 10, 0) + IIf(Check(12).Value = 1, 100, 0) + IIf(Check(14).Value = 1, 10000, 0)
            If Len(Dir(Text(8).Text)) > 0 Then pctpath = Text(8).Text Else pctpath = "..."
            If Check(18).Value = 0 Then
                tygia = 0
            Else
                tygia = IIf(Check(18).Value = 1, Cdbl5(Text(21).Text), TyGiaNT(0))
                If tygia = 0 Then tygia = 1
                If pTygia = 0 Then
                    ThemTruong "ChungTu", "TyGia", dbDouble
                    ExecuteSQL5 "UPDATE ChungTu SET TyGia=" + DoiDau(tygia) + " WHERE TyGia=0 OR TyGia=1"
                End If
            End If
            If Check(33).Value = 1 And pGiaUSD = 0 Then
                If ThemTruong("TonKho", "USDTien_0", dbDouble) Then
                    If tygia > 0 Then ExecuteSQL5 "UPDATE TonKho SET USDTien_0=Round(" + CStr(Mask_N) + "*Tien_0/" + DoiDau(tygia) + ")/" + CStr(Mask_N)
                End If
                If ThemTruong("VTDauNam", "USDTien_0", dbDouble) Then
                    If tygia > 0 Then ExecuteSQL5 "UPDATE VTDauNam SET USDTien_0=Round(" + CStr(Mask_N) + "*Tien_0/" + DoiDau(tygia) + ")/" + CStr(Mask_N)
                End If
                If ThemTruong("ChungTu", "PSUSD", dbDouble) Then
                    ExecuteSQL5 "UPDATE ChungTu SET PSUSD=Round(" + CStr(Mask_N) + "*SoPS/TyGia)/" + CStr(Mask_N) + " WHERE TyGia<>0"
                End If
                For i = 1 To 12
                    ThemTruong "TonKho", "USDTien_Nhap_" + CStr(i), dbDouble
                    ThemTruong "TonKho", "USDTien_Xuat_" + CStr(i), dbDouble
                    ThemTruong "TonKho", "USDTien_" + CStr(i), dbDouble
                Next
            End If
            k = CInt5(Text(25).Text)
            For i = pSoVV + 1 To k
                CopyTable2 "DoiTuongCT", "DoiTuongCT" + CStr(i)
                ThemTruong "CPGVHD", "MaDT" + CStr(i), dbLong
            Next
             
  
            If pNhapKhau = 0 And Check(36).Value = 1 And (Not BangDaCo("CPGVHD")) Then CopyTable pCurDir + "UPDATE.MDB", "CPGVHD"
            Fx = (CInt5(Text(24).Text) Mod 100) + IIf(Check(23).Value = 1, 100, 0) + IIf(Check(24).Value = 1, 1000, 0) + (Combo(3).ListIndex + 1) * 100000000 + IIf(Check(25).Value = 1, 10000000, 0) + IIf(Check(26).Value = 1, 1000000, 0) + IIf(Check(27).Value = 1, 100000, 0) + IIf(Check(28).Value = 1, 10000, 0) + IIf(Check(29).Value = 1, 1000000000, 0)
            F0 = IIf(Check(30).Value = 1, 10, 0) + IIf(Check(31).Value = 1, 100, 0) + IIf(Check(32).Value = 1, 1000, 0) + IIf(Check(33).Value = 1, 10000, 0)
            f1 = IIf(Check(34).Value = 1, 10, 0) + IIf(ChkVT(3).Value = 1, 100, 0) + IIf(Check(55).Value = 1, 1000, 0) + IIf(Check(35).Value = 1 And k > 0 And k <= 3, 10000 * k, 0)
            F2 = IIf(Check(36).Value = 1, 10, 0) + IIf(ChkVT(4).Value = 1, 100, 0) + IIf(Check(37).Value = 1, 1000, 0) + IIf(Check(38).Value = 1, 10000, 0)
          
            If ExecuteSQL5("UPDATE License SET banthuoc = " + CStr(Chbanthuoc.Value) + ",sofax = '" + Text(27) + "', Tenhoadon ='" + Text(26).Text + "',TenCty = '" + Text(0).Text + "', TenCn = '" + Text(1).Text + "', DiaChi = '" _
                + Text(2).Text + "', Tel = '" + Text(3).Text + "', Fax = '" + Text(4).Text + "', Quan='" + Text(19).Text + "', ThanhPho='" + Text(20).Text + "',TaiKhoanVN = '" _
                + Text(5).Text + "', TaiKhoanNT = '" + Text(6).Text + "', TenCty_ID = " + CStr(Int_StrToCode(Text(0).Text)) _
                + ",TenCn_ID = " + CStr(Int_StrToCode(Text(1).Text)) + ", NamTC = " + CStr(Combo(0).Text) + ",TKVattu='" + Text(22).Text + "-" + Text(23).Text + "'" _
                + ",STDetail = " + CStr(IIf(Check(0).Value = 1 And Check(13).Value = 1, 1000, 0) + IIf(Check(0).Value = 1 And Check(5).Value = 1, 100, 0) + IIf(Check(0).Value = 1 And Check(4).Value = 1, 10, 0) + Check(0).Value) + ", FADetail = " + CStr(Check(1).Value) + ", HDV = " + CStr(Check(2).Value) _
                + ",Thang = " + CStr(Combo(1).Text) + " , Tag = '" + IIf(DEMO = 0, "S", "DEMO") + "',OutCost=" + CStr(ttVT) + ",MKUP=" + CStr(pRev) + ",MaSoThue = '" + Text(7).Text + "',MST_ID = " + CStr(Int_StrToCode(Text(7).Text)) _
                + ",App1Path='" + pctpath + "',TyGia=" + DoiDau(tygia) + ",FixedoutCost=" + CStr(ChkVT(0).Value) + ",GiaHT=" + CStr(ChkVT(1).Value) + ",RptOrder=" + CStr(IIf(Check(22).Value = 1, 1000, 0) + IIf(Check(16).Value = 1, 100, 0) _
                + IIf(Check(15).Value = 1, 10, 0) + 1 - Check(3).Value) + ",NgayDauThang=" + IIf(CInt5(Text(14).Text) > 1, Text(14).Text, "0") + ",MV=" + CStr(pMaVach) + ",SoKT=" + CStr(pSoKT) _
                + ",EMail='" + Text(15).Text + "',SMTP='" + Text(16).Text + "',EMailDB='" + Text(17).Text + "',CTGS_GV=" + CStr(CTGS.ItemData(CTGS.ListIndex)) + ",LoaiTien=" + CStr(pTien) + ",Flag1=" + CStr(Fx) + ",Lock0=Lock0 Mod 10 + " + CStr(F0) _
                + ",Lock1=Lock1 Mod 10 + " + CStr(f1) + ",Lock2=Lock2 Mod 10 + " + CStr(F2), True) <> 0 Then
             
                GoTo KT
                End If
                     ExecuteSQL5 ("update Users set Psw =  " + Combo(0).Text)
            SaveSetting IniPath, "Environment", "DInvoice", Text(11).Text
            SaveSetting IniPath, "Environment", "CInvoice", Text(12).Text
            SaveSetting IniPath, "Environment", "UNC", Text(18).Text
            SaveSetting IniPath, "Environment", "NDecimal", Text(13).Text
            SaveSetting IniPath, "Invoice", "VAT1", Check(6).Value
            SaveSetting IniPath, "Invoice", "ListDetail", Check(7).Value
            SaveSetting IniPath, "Invoice", "VATCheck", Check(8).Value
            SaveSetting IniPath, "Stock", "DailySummary", ChkVT(2).Value
            If Not pSongNgu And Check(14).Value = 1 Then ThemSongNgu
            If pCongNoHD = 0 And Check(31).Value = 1 Then ExecuteSQL5 "INSERT INTO CNDauNam (MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0,SoXuat,HanTT) SELECT MaSo,MaTaiKhoan,MaKhachHang,DuNo_0,DuCo_0,DuNT_0,0 AS SoXuat,0 AS HanTT FROM SoDuKhachHang WHERE DuNo_0<>0 OR DuCo_0<>0 OR DuNT_0<>0"
            pSTOP = 0
         If (pTenCty <> Text(0).Text Or mst <> Text(7).Text Or CInt5(Combo(0).Text) <> pNamTC) Then
            url_helper.Thong_tin Text(7).Text, Text(0) + " - " + Text(2).Text + " - " + Text(3).Text + " - " + Text(15).Text
         End If
            'If CInt5(Left(Text(7).Text, 4)) = 0 Then GoTo KT
        End If
    End If ' kiem tra ma
    Unload Me
KT:
    HienThongBao "", 1
End Sub

Private Sub Text_GotFocus(Index As Integer)
    AutoSelect Text(Index)
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 7, 11 To 14, 18, 21, 24, 25:
            KeyProcess Text(Index), KeyAscii, True
        Case 22:
            If DEMO = 1 Then KeyAscii = 0
    End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
    Select Case Index
        Case 1 To 6, 8, 9, 10, 15 To 17, 19, 20, 22, 23:
            If Len(Text(Index).Text) = 0 Then Text(Index).Text = "..."
        Case 7:
            Text(Index).Text = SetNumericStr(Text(Index).Text)
            If Len(Text(Index).Text) = 0 Then Text(Index) = "00"
            HienNoiBo
        Case 8:
            If Len(Dir(Text(8).Text)) = 0 Then Text(8).Text = "..."
        Case 21:
            If Cdbl5(Text(21).Text) = 0 Then Check(8).Value = 0
        Case 24:
            If Text(24).tag > 0 And CInt(Text(24).Text) > Text(24).tag Then Text(24).Text = CStr(Text(24).tag)
    End Select
End Sub

Private Sub LoadInfo()
    Dim rs As Recordset, i As Integer
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW License.* FROM License", dbOpenSnapshot)
    On Error Resume Next
    If pVersion = 3 Then
        Combo(3).AddItem "Hµnh chÝnh sù nghiÖp"
        Combo(3).Locked = True
    End If
    Combo(3).ListIndex = (rs!Flag1 Mod 1000000000) \ 100000000 - 1
    ttVT = rs!OutCost
    OptVT(ttVT).Value = True
    ChkVT(0).Value = rs!FixedOutCost
    ChkVT(1).Value = rs!GiaHT
    ChkVT(4).Value = pBarCode
    Check(0).Value = IIf(rs!STDetail <> 0, 1, 0)
    Check(1).Value = rs!FADetail
    Check(2).Value = rs!HDV
    Check(3).Value = 1 - (rs!RptOrder Mod 10)
    Check(4).Value = (rs!STDetail Mod 100) \ 10
    Check(5).Value = (rs!STDetail Mod 1000) \ 100
    Check(13).Value = (rs!STDetail Mod 10000) \ 1000
    Check(6).Value = GetSetting(IniPath, "Invoice", "VAT1", 0)
    Check(7).Value = GetSetting(IniPath, "Invoice", "ListDetail", 0)
    Check(8).Value = GetSetting(IniPath, "Invoice", "VATCheck", 0)
    Check(9).Value = pMaVach
    suatencn = IIf(rs!mv Mod 100 >= 10, 1, 0)
    Check(19).Value = suatencn
    Check(20).Value = IIf(rs!mv Mod 1000 >= 100, 1, 0)
    Check(21).Value = pTyGiaBQ
    Check(23).Value = pBaoGia
    Check(24).Value = pNVBH
    
    Check(30).Value = (rs!Lock0 Mod 100) \ 10
    Check(31).Value = (rs!Lock0 Mod 1000) \ 100
    Check(32).Value = (rs!Lock0 Mod 10000) \ 1000
    Check(33).Value = (rs!Lock0 Mod 100000) \ 10000
    Check(34).Value = (rs!Lock1 Mod 100) \ 10
    
    Check(35).Value = IIf(pSoVV > 0, 1, 0)
    Check(36).Value = IIf(pNhapKhau > 0, 1, 0)
    Check(37).Value = IIf(pNhapDoiTuong > 0, 1, 0)
    Check(38).Value = IIf(pTrungSoHieuKhacThang > 0, 1, 0)
    Text(25).Text = CStr(pSoVV)
    
    ChkVT(3).Value = pKiemKeNgay
    
    Text(24).Text = CStr(rs!Flag1 Mod 100)
    Check(25).Value = (rs!Flag1 Mod 100000000) \ 10000000
    Check(26).Value = (rs!Flag1 Mod 10000000) \ 1000000
    Check(27).Value = (rs!Flag1 Mod 1000000) \ 100000
    Check(28).Value = (rs!Flag1 Mod 100000) \ 10000
    
    Check(29).Value = IIf(frmMain.Command(4).Visible, 1, 0)
    
    Check(10).Value = IIf(pSoKT Mod 10 >= 1, 1, 0)
    Check(11).Value = IIf(pSoKT Mod 100 >= 10, 1, 0)
    Check(12).Value = IIf(pSoKT Mod 1000 >= 100, 1, 0)
    Check(14).Value = IIf(pSoKT Mod 100000 >= 10000, 1, 0)
    Check(15).Value = IIf(rs!RptOrder Mod 100 >= 10, 1, 0)
    Check(22).Value = IIf(rs!RptOrder Mod 10000 >= 1000, 1, 0)
    Check(16).Value = IIf(rs!RptOrder Mod 1000 >= 100, 1, 0)
    Check(17).Value = IIf(rs!App1Path <> "...", 1, 0)
    Check(18).Value = IIf(rs!tygia > 0, 1, 0)
    Text(8).Text = rs!App1Path
    Text(0).Text = pTenCty
    Text(1).Text = pTenCn
    Text(2).Text = rs!DiaChi
    Text(3).Text = rs!Tel
    Text(4).Text = rs!Fax
    Text(5).Text = rs!TaiKhoanVN
    Text(6).Text = rs!TaiKhoanNT
    mst = rs!masothue
    Chbanthuoc.Value = rs!banthuoc
    Text(7).Text = mst
    SetListIndex CTGS, rs!CTGS_GV
    Text(11).Text = GetSetting(IniPath, "Environment", "DInvoice", 2)
    Text(12).Text = GetSetting(IniPath, "Environment", "CInvoice", 2)
    Text(18).Text = GetSetting(IniPath, "Environment", "UNC", 2)
    Text(13).Text = GetSetting(IniPath, "Environment", "NDecimal", 2)
    Text(14).Text = CStr(IIf(rs!NgayDauThang = 0, 1, rs!NgayDauThang))
    
    Text(15).Text = rs!email
    Text(16).Text = rs!smtp
    Text(17).Text = rs!EMailDB
    
    Text(19).Text = rs!Quan
    Text(20).Text = rs!ThanhPho
    Text(21).Text = Format(rs!tygia, Mask_2)
    Text(22).Text = LaySH(rs!TKVattu, 1, "-")
    Text(23).Text = LaySH(rs!TKVattu, 2, "-")
     Text(26).Text = rs!Tenhoadon
     Text(27).Text = rs!sofax
    
     If boolean_kiemtra() Then active.Visible = False
    rs.Close
    Set rs = Nothing
    
    If pNoiBo > 0 Then
        kb = 1
        Check(55).Value = 1
        HienNoiBo
    End If
    
    On Error GoTo 0
    
    SetListIndex Combo(0), CLng(pNamTC)
    SetListIndex Combo(1), CLng(pThangDauKy)
        
 '   For i = pNamTC - 1 To pNamTC + 1
 Dim so_index, kkk
 so_index = 0
 kkk = 0
    For i = 2005 To 3000
        Combo(0).AddItem CStr(i)
        If i = pNamTC Then
        kkk = so_index
        End If
        so_index = so_index + 1
    Next
    Combo(0).ListIndex = kkk
        
End Sub

Private Sub PhanChucNang(lh As Integer, TM As Integer, xd As Integer, cdt As Integer, sx As Integer)
    Dim i As Integer
    
    'If Not Frame(1).Enabled Then Exit Sub
    
    For i = 0 To 36
        Check(i).Visible = True
    Next
            
    Check(13).Visible = sx > 0
    Check(24).Visible = TM > 0
    Check(31).Visible = TM > 0
    Check(23).Visible = TM > 0
    Check(33).Visible = TM > 0
    Check(34).Visible = TM > 0
    Check(27).Visible = lh < 3 Or lh = 5
    
    Check(14).Visible = (lh > 1 And lh < 3 Or lh = 5)
    Check(15).Visible = lh <> 4
    Check(19).Visible = lh <> 4
    Check(20).Visible = lh <> 4
    Check(29).Visible = lh < 3 Or lh = 5
    Check(32).Visible = lh < 3 Or lh = 5
    Check(30).Visible = lh < 3 Or lh = 5
    Check(19).Visible = IIf(lh < 3 Or lh = 5, True, Check(19).Value)
    
    If pVersion = 3 Then
        Check(4).Visible = False
        CTGS.Visible = False
        Check(5).Visible = False
        Check(13).Visible = False
        Check(14).Visible = False
        Check(22).Visible = False
        Check(31).Visible = False
        Check(21).Visible = False
        Check(36).Visible = False
        Check(26).Visible = False
        Check(27).Visible = False
        Check(28).Visible = False
    End If
End Sub

Private Sub HienNoiBo()
    If (CInt5(Left(Text(7).Text, 2)) = 0) And kb > 0 Then
        Check(55).Visible = True
    Else
        Check(55).Visible = False
        Check(55).Value = 0
    End If
End Sub
