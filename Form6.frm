VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDauky 
   Caption         =   "Dau ky"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7365
   LinkTopic       =   "Form4"
   ScaleHeight     =   2520
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox suasolieucu 
      Caption         =   "Dua ton dau san pham co san"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Xoa dau ky cua TSCD"
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TSCD"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "He thong TK"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cong no"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "San Pham"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin MSForms.OptionButton chon 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   1935
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "3413;661"
      Value           =   "0"
      Caption         =   "Them muc chua co"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton chon 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   1335
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "2355;661"
      Value           =   "0"
      Caption         =   "chuyen moi"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmDauky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
