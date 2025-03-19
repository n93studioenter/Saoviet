VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Th«ng tin"
   ClientHeight    =   3990
   ClientLeft      =   2010
   ClientTop       =   1740
   ClientWidth     =   6120
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C000&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   3585
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label LbAbout 
      BackColor       =   &H00808000&
      Caption         =   "§­îc triÓn khai bëi C«ng Ty TNHH DV ThuÕ Sao ViÖt"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   7
      Left            =   240
      TabIndex        =   12
      Tag             =   "App Description"
      Top             =   840
      Width           =   3885
   End
   Begin VB.Label LbAbout 
      BackColor       =   &H00808000&
      Caption         =   "§­îc chuyÓn nh­îng bëi C«ng ty TNHH C«ng NghÖ Tin Häc ViÔn S¬n"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   6
      Left            =   240
      TabIndex        =   11
      Tag             =   "App Description"
      Top             =   600
      Width           =   5565
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H00808000&
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
      Height          =   345
      Left            =   480
      MouseIcon       =   "frmAbout.frx":57E2
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3600
      Width           =   3630
   End
   Begin VB.Label LbAbout 
      BackColor       =   &H00808000&
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
      Left            =   2040
      TabIndex        =   9
      Top             =   1200
      Width           =   2925
   End
   Begin VB.Label LbAbout 
      BackColor       =   &H00808000&
      Caption         =   "Kh«ng chÞu tr¸ch nhiÖm ®èi víi c¸c b¶n sao chÐp kh«ng ®¨ng ký"
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   4
      Left            =   1080
      TabIndex        =   8
      Tag             =   "App Description"
      Top             =   2760
      Width           =   3045
   End
   Begin VB.Label LbAbout 
      BackColor       =   &H00808000&
      Caption         =   "Mäi chi tiÕt xin liªn hÖ: 090 575 7799"
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   5
      Left            =   1080
      TabIndex        =   7
      Tag             =   "App Description"
      Top             =   3240
      Width           =   4365
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   5400
      X2              =   5400
      Y1              =   1740
      Y2              =   2420
   End
   Begin VB.Line Line 
      Index           =   2
      X1              =   1080
      X2              =   1080
      Y1              =   1760
      Y2              =   2420
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   1080
      X2              =   5400
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   1080
      X2              =   5400
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label LbCty 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Tag             =   "Application Title"
      Top             =   2115
      Width           =   150
   End
   Begin VB.Label LbCty 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      Caption         =   "No license"
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
      Left            =   1200
      TabIndex        =   5
      Tag             =   "Application Title"
      Top             =   1815
      Width           =   1125
   End
   Begin VB.Label LbAbout 
      BackColor       =   &H00808000&
      Caption         =   "Ch­¬ng tr×nh ®­îc cµi ®Æt t¹i"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   1080
      TabIndex        =   4
      Tag             =   "App Description"
      Top             =   1440
      Width           =   3885
   End
   Begin VB.Label LbAbout 
      BackColor       =   &H00808000&
      Caption         =   "Ch­¬ng tr×nh kÕ to¸n VietStar"
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
      Left            =   840
      TabIndex        =   3
      Tag             =   "UNESCO Accounting"
      Top             =   240
      Width           =   4245
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   240
      X2              =   5465
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      X1              =   240
      X2              =   5450
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Label LbAbout 
      BackColor       =   &H00808000&
      Caption         =   "Phiªn b¶n:"
      Height          =   225
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Tag             =   "Version"
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00808000&
      Caption         =   "Chó ý:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   255
      TabIndex        =   1
      Tag             =   "Warning: ..."
      Top             =   2805
      Width           =   630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdok_Click()
        Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 4350
Me.Width = 5745
    LbCty(0).Caption = pTenCty
    LbCty(1).Caption = pTenCn
    lblEmail.Left = 1050
    lblEmail.Top = 3720
    LbAbout(2).Caption = frmMain.Lb(0).Caption
    On Error Resume Next
    LbCty(0).FontName = pFontName
    LbCty(1).FontName = pFontName
    On Error GoTo 0
   ' If DEMO = 1 Then LbAbout(1).Caption = "Phiªn b¶n §µo t¹o                                   2003 - 2008"
End Sub

Private Sub lblEmail_Click()
   'ShellExecute ByVal 0&, "open", "www.gmail.com", vbNullString, vbNullString, 3
End Sub
