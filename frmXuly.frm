VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmXuly 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   735
   ClientLeft      =   5925
   ClientTop       =   5880
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "frmXuly.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin ComctlLib.ProgressBar Gau 
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "LÖnh Xö lý cho nh©n viªn PhÇn MÒm"
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
      Height          =   225
      Left            =   540
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2880
   End
End
Attribute VB_Name = "frmXuly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
