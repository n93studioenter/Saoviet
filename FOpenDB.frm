VERSION 5.00
Begin VB.Form FOpenDB 
   AutoRedraw      =   -1  'True
   Caption         =   "Open Database Utility"
   ClientHeight    =   1095
   ClientLeft      =   3540
   ClientTop       =   3660
   ClientWidth     =   3150
   ClipControls    =   0   'False
   Icon            =   "FOpenDB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   3150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      Caption         =   "&Close"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.PictureBox Dlg 
      Height          =   480
      Left            =   3120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   360
      Width           =   1200
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H8000000D&
      Caption         =   "&Select database"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "FOpenDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 0:
            With Dlg
                .DialogTitle = "Chän tÖp d÷ liÖu"
                .Flags = &H4&
                .FileName = "*.MDB"
                On Error GoTo XongMN
                .ShowOpen
                On Error GoTo 0
                If Len(.FileName) = 0 Or Left(.FileName, 1) = "*" Then GoTo XongMN
                OpenDB1 .FileName
            End With
        Case 1:
            Unload Me
    End Select
XongMN:
End Sub
