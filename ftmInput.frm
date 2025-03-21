VERSION 5.00
Begin VB.Form ftmInput 
   Caption         =   "ftmInput"
   ClientHeight    =   1125
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   3465
   LinkTopic       =   "Form4"
   ScaleHeight     =   1125
   ScaleWidth      =   3465
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "ftmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmLocImport.ChangeValueInpput Text1.Text
        'Cap nhat item luon cho List
        MsgBox frmLocImport.rowSelect
        Me.Hide
    End If

End Sub
