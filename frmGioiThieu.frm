VERSION 5.00
Begin VB.Form frmgioithieu 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Gioi thieu"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGioiThieu.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5460
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   8280
      Picture         =   "frmGioiThieu.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "&Return"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PhÇn mÒm hoµn toµn miÔn phÝ cho doanh nghiÖp nhá, cã sè chøng tõ nhá h¬n 500 vµ doanh thu nhá h¬n  05 tû."
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Xin vui lßng liªn hÖ  090 575 7799 Mr. H­ng. "
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2520
      TabIndex        =   11
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"frmGioiThieu.frx":6C04
      Height          =   615
      Index           =   6
      Left            =   360
      TabIndex        =   10
      Top             =   3840
      Width           =   9135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Giíi thiÖu"
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
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Index           =   5
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   9135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Lµm ®­îc nhiÒu c«ng ty trªn mét phÇn mÒm, thªm c«ng ty míi ®¬n gi¶n, phï hîp cho ng­êi dïng lµm dÞch vô kÕ to¸n."
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- ChuyÓn trùc tiÕp c¸c b¸o c¸o ra m· v¹ch."
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Më c¸c tµi kho¶n chi tiÕt ®¬n gi¶n, ®¨ng ký theo dâi ®èi t­îng nhanh."
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- PhÇn mÒm chØ nhËp mét lÇn lµ ra tÊt c¶ c¸c sæ, tæng hîp, chi tiÕt, vËt t­, c«ng nî, c«ng tr×nh..."
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Tæng hîp, ng©n hµng, vËt t­, b¸n hµng, TSC§, dông cô, kÕt chuyÓn tù ®éng..."
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "PhÇn mÒm gåm c¸c ph©n hÖ:"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"frmGioiThieu.frx":6CB7
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   9015
   End
End
Attribute VB_Name = "frmgioithieu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command_Click(Index As Integer)
Unload Me
End Sub

