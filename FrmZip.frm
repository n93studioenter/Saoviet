VERSION 5.00
Object = "{E88121A0-9FA9-11CF-9D9F-00AA003A3AA3}#1.0#0"; "ZLIBTOOL.OCX"
Begin VB.Form FrmZip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   240
   ClientLeft      =   1395
   ClientTop       =   9315
   ClientWidth     =   11160
   ClipControls    =   0   'False
   Icon            =   "FrmZip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin ZLIBTOOLLib.ZlibTool zip 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11175
      _Version        =   65536
      _ExtentX        =   19711
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin VB.Label lblHienthi 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Vui lßng chê trong giay l¸t/ Please waiting . . . . . . . . . ."
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   135
      Left            =   0
      TabIndex        =   0
      Tag             =   "Please waiting . . . . . . .  "
      Top             =   0
      Visible         =   0   'False
      Width           =   7515
   End
End
Attribute VB_Name = "FrmZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

