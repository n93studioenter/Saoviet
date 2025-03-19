VERSION 5.00
Begin VB.Form FPsw 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MÀt kh»u"
   ClientHeight    =   705
   ClientLeft      =   5730
   ClientTop       =   5100
   ClientWidth     =   4065
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
   Icon            =   "FPsw.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Password"
   Begin VB.TextBox txtPsw 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MÀt kh»u"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "Password"
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FPsw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pswx As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13:
            pswx = txtPsw.Text
            Unload Me
        Case 27:
            Unload Me
    End Select
End Sub

Public Function GetPswX(Optional des As String = "MÀt kh»u") As String
    Me.Caption = des
    Me.Show vbModal
    GetPswX = pswx
End Function
