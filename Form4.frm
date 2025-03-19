VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{E88121A0-9FA9-11CF-9D9F-00AA003A3AA3}#1.0#0"; "ZLIBTOOL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11250
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   11250
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   2566
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   16761024
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form4.frx":57E2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form4.frx":57FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form4.frx":581A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
   End
   Begin ZLIBTOOLLib.ZlibTool ZlibTool1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   661
      _StockProps     =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
