VERSION 5.00
Begin VB.Form FrmOptions 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Th«ng tin doanh nghiÖp"
   ClientHeight    =   5055
   ClientLeft      =   660
   ClientTop       =   345
   ClientWidth     =   10335
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
   ScaleHeight     =   5055
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Options"
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
      Left            =   4080
      MaxLength       =   300
      TabIndex        =   110
      Text            =   "0"
      Top             =   2160
      Width           =   1095
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
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   99
      Text            =   "..."
      Top             =   1110
      Width           =   1335
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
      Height          =   3255
      Index           =   0
      Left            =   0
      TabIndex        =   60
      Top             =   -40
      Width           =   9015
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
         TabIndex        =   109
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
         Left            =   6600
         TabIndex        =   95
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
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
         Height          =   255
         Index           =   55
         Left            =   6840
         TabIndex        =   102
         Top             =   2160
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
         Index           =   9
         Left            =   8640
         MaxLength       =   20
         TabIndex        =   16
         Text            =   "..."
         Top             =   2520
         Visible         =   0   'False
         Width           =   2055
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
         Height          =   255
         Index           =   27
         Left            =   7200
         TabIndex        =   15
         Top             =   2235
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
         Left            =   6600
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   14
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
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
         Left            =   6600
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   13
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
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
         Left            =   8160
         List            =   "Frmopt.frx":AFDC
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2160
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
         Index           =   17
         Left            =   8400
         MaxLength       =   30
         TabIndex        =   23
         Text            =   "..."
         Top             =   2880
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
         Top             =   2880
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
         Left            =   7560
         List            =   "Frmopt.frx":B06B
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2880
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
         Top             =   2880
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
         Top             =   2880
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
         Left            =   8520
         MaxLength       =   3
         TabIndex        =   2
         Top             =   480
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
         Top             =   720
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
         Top             =   1155
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
         Index           =   5
         Left            =   1800
         MaxLength       =   5000
         TabIndex        =   9
         Text            =   "..."
         Top             =   1560
         Width           =   7095
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
         Top             =   1875
         Width           =   7095
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
         Top             =   2160
         Width           =   2055
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
         Left            =   1800
         List            =   "Frmopt.frx":B098
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2475
         Width           =   1095
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
         Left            =   7680
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "0"
         Top             =   2640
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
         Top             =   1160
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
         Index           =   15
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "..."
         Top             =   1160
         Width           =   3375
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
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "..."
         Top             =   2840
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
         Index           =   20
         Left            =   4680
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "..."
         Top             =   2500
         Width           =   1815
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
         Left            =   5160
         TabIndex        =   92
         Tag             =   "Activities"
         Top             =   1890
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
         Left            =   7200
         TabIndex        =   91
         Tag             =   "Class"
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "H¹ch to¸n theo Q§ 48, 15:"
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
         Left            =   3840
         TabIndex        =   85
         Tag             =   "Send data to default addr"
         Top             =   2880
         Width           =   2055
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
         TabIndex        =   84
         Top             =   2880
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
         TabIndex        =   83
         Top             =   2880
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
         Left            =   120
         TabIndex        =   82
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tªn c«ng ty "
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
         TabIndex        =   73
         Tag             =   "Company"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tªn theo hãa ®¬n"
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
         TabIndex        =   72
         Tag             =   "Branch Name"
         Top             =   555
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "§Þa chØ, quËn, t. phè"
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
         TabIndex        =   71
         Tag             =   "Address"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "§iÖn tho¹i"
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
         TabIndex        =   70
         Tag             =   "Tel"
         Top             =   1155
         Width           =   1095
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
         Left            =   3240
         TabIndex        =   69
         Top             =   1155
         Width           =   375
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tµi kho¶n TiÒn ViÖt"
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
         TabIndex        =   68
         Tag             =   "Bank VND Account"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tµi kho¶n Ngo¹i tÖ"
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
         TabIndex        =   67
         Tag             =   "Bank F.C. Account"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "M· sè thuÕ"
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
         TabIndex        =   66
         Tag             =   "Tax Code"
         Top             =   2235
         Width           =   1575
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "N¨m tµi chÝnh"
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
         TabIndex        =   65
         Tag             =   "Year"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "LÜnh vùc:"
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
         Left            =   3840
         TabIndex        =   64
         Tag             =   "From month"
         Top             =   2520
         Width           =   855
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
         Left            =   6360
         TabIndex        =   63
         Tag             =   "Month from Date"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   5040
         TabIndex        =   62
         Top             =   1200
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
         TabIndex        =   61
         Top             =   1200
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
      Height          =   5460
      Index           =   1
      Left            =   0
      TabIndex        =   74
      Top             =   3240
      Width           =   9015
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
         Left            =   5880
         TabIndex        =   108
         Top             =   4680
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
         TabIndex        =   107
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
         Left            =   5880
         TabIndex        =   105
         Top             =   4440
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
         TabIndex        =   104
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
         Left            =   5880
         TabIndex        =   103
         Top             =   4200
         Visible         =   0   'False
         Width           =   3495
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
         TabIndex        =   54
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5880
         TabIndex        =   100
         Top             =   3960
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
         TabIndex        =   98
         Top             =   4680
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
         Left            =   5880
         TabIndex        =   97
         Top             =   2520
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
         Left            =   5880
         TabIndex        =   96
         Top             =   2760
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
         Left            =   5880
         TabIndex        =   94
         Top             =   3720
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
         Left            =   5880
         TabIndex        =   93
         Top             =   3480
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
         Left            =   5880
         TabIndex        =   90
         Top             =   3240
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
         TabIndex        =   89
         Top             =   4440
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
         TabIndex        =   88
         Text            =   "..."
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
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
         Top             =   4920
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
         Left            =   6960
         MaxLength       =   2
         TabIndex        =   86
         Tag             =   "0"
         Text            =   "0"
         Top             =   1320
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
         Left            =   5880
         TabIndex        =   81
         Top             =   3000
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
         TabIndex        =   80
         Top             =   4200
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
         Top             =   2760
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
         Top             =   3960
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
         Top             =   3960
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
         Left            =   3840
         List            =   "Frmopt.frx":B09C
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   2520
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
         Height          =   1620
         Index           =   2
         Left            =   0
         TabIndex        =   75
         Top             =   120
         Width           =   4215
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
            TabIndex        =   106
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
            Left            =   1680
            TabIndex        =   101
            Top             =   960
            Visible         =   0   'False
            Width           =   2295
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
            Left            =   1680
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
            Left            =   1680
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
            Left            =   1680
            TabIndex        =   40
            Top             =   1200
            Visible         =   0   'False
            Width           =   2175
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
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   52
         Text            =   "2"
         Top             =   1080
         Width           =   495
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
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tù ®éng trõ lïi thuÕ GTGT trªn ho¸ ®¬n b¸n"
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
         Top             =   2520
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
         Left            =   120
         TabIndex        =   43
         Top             =   2760
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
         Left            =   120
         TabIndex        =   44
         Top             =   3000
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
         Top             =   2280
         Width           =   2415
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
         TabIndex        =   53
         Top             =   360
         Value           =   1  'Checked
         Width           =   1455
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
         TabIndex        =   55
         Top             =   840
         Value           =   1  'Checked
         Width           =   1455
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
         Top             =   1440
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
         Left            =   8880
         TabIndex        =   56
         Top             =   2520
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
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CheckBox Check 
         BackColor       =   &H00E0E0E0&
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
         Top             =   3480
         Visible         =   0   'False
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
         Top             =   3720
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
         Top             =   3720
         Width           =   2295
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
         Left            =   4440
         TabIndex        =   87
         Top             =   1440
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
         Left            =   8280
         TabIndex        =   79
         Top             =   2520
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
         Left            =   3720
         TabIndex        =   78
         Top             =   2280
         Width           =   1575
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
         Left            =   4440
         TabIndex        =   77
         Top             =   1200
         Width           =   1695
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
         TabIndex        =   76
         Top             =   120
         Width           =   1335
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
      TabIndex        =   59
      Tag             =   "&Return"
      Top             =   4560
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
      TabIndex        =   58
      Tag             =   "&Save"
      Top             =   4080
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
End Sub

Private Sub Label_Click(Index As Integer)

End Sub

Private Sub OptVT_Click(Index As Integer)
    ttVT = Index
End Sub

Private Sub Command_Click(Index As Integer)
    
    Dim i As Integer, tygia As Double, T As Long, mk As Long, Fx As Long, pctpath As String, F0 As Integer, f1 As Integer, k As Integer, F2 As Integer
         ExecuteSQL5_Themmoi ("ALTER TABLE license  ADD tenhoadon text")
         ExecuteSQL5 ("ALTER TABLE license ALTER COLUMN TaiKhoanVN TEXT(200)")
          ExecuteSQL5 ("ALTER TABLE license ALTER COLUMN DiaChi TEXT(200)")
         ' ExecuteSQL5 ("ALTER TABLE license add COLUMN sofax TEXT(200)")
          ExecuteSQL5_Themmoi ("ALTER TABLE license add COLUMN sofax TEXT(200)")
    
             ExecuteSQL5 ("UPDATE License SET Tenhoadon ='" + Text(26).Text + "',TaiKhoanVN = '" + Text(5).Text + "'")
    If Index = 0 Then
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
        If Combo(2).ListIndex >= 0 Then T = Combo(2).ItemData(Combo(2).ListIndex) Else T = pTien
        If CInt5(Left(Text(Index).Text, 2)) <> 0 Then Check(55).Value = 0
        If ((((pTenCty = Text(0).Text And (pTenCn = Text(1).Text Or suatencn = 1) And (Check(19).Value = suatencn) And pMaVach = Check(9).Value And pDinhmuc = Check(13).Value And pSongNgu = (Check(14).Value = 1) And pRpt = Check(15).Value And pTygia = Check(18).Value And T = pTien And mk = 0) Or (DEMO = 1 And CLng5(Left(Text(7).Text, 2)) > 0)) And (mst = Text(7).Text Or (suatencn = 1 And Left(mst, 10) = Left(Text(7).Text, 10)))) Or Combo(3).ListIndex = 4 Or (Cdbl5(Left(Text(7).Text, 10)) = 0 And Cdbl5(Left(frmMain.LbCty(8).Caption, 10)) = 0)) And (pNoiBo = Check(55).Value) And (CInt5(Combo(0).Text) = pNamTC) Then GoTo a
        If (Len(pMST) > 0 And Left(Text(7).Text, Len(pMST)) = pMST) Then GoTo a
        If FrmGetStr.GetMK(Text(7).Text) Then
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
            If ExecuteSQL5("UPDATE License SET sofax = '" + Text(27) + "', Tenhoadon ='" + Text(26).Text + "',TenCty = '" + Text(0).Text + "', TenCn = '" + Text(1).Text + "', DiaChi = '" _
                + Text(2).Text + "', Tel = '" + Text(3).Text + "', Fax = '" + Text(4).Text + "', Quan='" + Text(19).Text + "', ThanhPho='" + Text(20).Text + "',TaiKhoanVN = '" _
                + Text(5).Text + "', TaiKhoanNT = '" + Text(6).Text + "', TenCty_ID = " + CStr(Int_StrToCode(Text(0).Text)) _
                + ",TenCn_ID = " + CStr(Int_StrToCode(Text(1).Text)) + ", NamTC = " + CStr(Combo(0).Text) + ",TKVattu='" + Text(22).Text + "-" + Text(23).Text + "'" _
                + ",STDetail = " + CStr(IIf(Check(0).Value = 1 And Check(13).Value = 1, 1000, 0) + IIf(Check(0).Value = 1 And Check(5).Value = 1, 100, 0) + IIf(Check(0).Value = 1 And Check(4).Value = 1, 10, 0) + Check(0).Value) + ", FADetail = " + CStr(Check(1).Value) + ", HDV = " + CStr(Check(2).Value) _
                + ",Thang = " + CStr(Combo(1).Text) + " , Tag = '" + IIf(DEMO = 0, "S", "DEMO") + "',OutCost=" + CStr(ttVT) + ",MKUP=" + CStr(pRev) + ",MaSoThue = '" + Text(7).Text + "',MST_ID = " + CStr(Int_StrToCode(Text(7).Text)) _
                + ",App1Path='" + pctpath + "',TyGia=" + DoiDau(tygia) + ",FixedoutCost=" + CStr(ChkVT(0).Value) + ",GiaHT=" + CStr(ChkVT(1).Value) + ",RptOrder=" + CStr(IIf(Check(22).Value = 1, 1000, 0) + IIf(Check(16).Value = 1, 100, 0) _
                + IIf(Check(15).Value = 1, 10, 0) + 1 - Check(3).Value) + ",NgayDauThang=" + IIf(CInt5(Text(14).Text) > 1, Text(14).Text, "0") + ",MV=" + CStr(pMaVach) + ",SoKT=" + CStr(pSoKT) _
                + ",EMail='" + Text(15).Text + "',SMTP='" + Text(16).Text + "',EMailDB='" + Text(17).Text + "',CTGS_GV=" + CStr(CTGS.ItemData(CTGS.ListIndex)) + ",LoaiTien=" + CStr(pTien) + ",Flag1=" + CStr(Fx) + ",Lock0=Lock0 Mod 10 + " + CStr(F0) _
                + ",Lock1=Lock1 Mod 10 + " + CStr(f1) + ",Lock2=Lock2 Mod 10 + " + CStr(F2), True) <> 0 Then GoTo KT
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
            'If CInt5(Left(Text(7).Text, 4)) = 0 Then GoTo KT
        End If
    End If
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
    mst = rs!MaSoThue
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
        
    For i = pNamTC - 1 To pNamTC + 1
        Combo(0).AddItem CStr(i)
    Next
    Combo(0).ListIndex = 1
        
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
