VERSION 5.00
Begin VB.Form FrmE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DiÔn gi¶i song ng÷"
   ClientHeight    =   4455
   ClientLeft      =   2985
   ClientTop       =   1680
   ClientWidth     =   6360
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
   Icon            =   "FrmE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Financial Report Description"
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   6135
   End
   Begin VB.ListBox Lst 
      Height          =   3570
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6135
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      ItemData        =   "FrmE.frx":57E2
      Left            =   120
      List            =   "FrmE.frx":57F8
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "FrmE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tbl As String
Dim fld As String

Private Sub Cbo_Click()
    Select Case Cbo.ListIndex
        Case 0:     tbl = "CDTS"
                            fld = "DienGiaiE"
        Case 1:     tbl = "KQKD"
                            fld = "TenE"
        Case 2:     tbl = "Thue"
                            fld = "DienGiaiE"
        Case 3:     tbl = "VAT"
                            fld = "DienGiaiE"
        Case 4:     tbl = "LCTT"
                            fld = "TenE"
        Case 5:     tbl = "KQKDCT"
                            fld = "TenE"
    End Select
    Int_RecsetToCbo "SELECT MaSo As F2,First(IIF(Left(" + fld + ",1)='.'," + Left(fld, Len(fld) - 1) + "," + fld + ")) As F1 FROM " + tbl + " GROUP BY MaSo", Lst
End Sub

Private Sub Form_Load()
    Cbo.ListIndex = 0
    Lst.FontName = pFontName
    txt.FontName = pFontName
End Sub

Private Sub Lst_Click()
    txt.Text = Lst.Text
    RFocus txt
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub txt_GotFocus()
    AutoSelect txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Lst.ListIndex >= 0 And Len(txt.Text) > 0 Then
        ExecuteSQL5 "UPDATE " + tbl + " SET " + fld + " = '" + txt.Text + "' WHERE MaSo=" + CStr(Lst.ItemData(Lst.ListIndex))
        On Error Resume Next
        DBKetoan.Execute "UPDATE " + tbl + " IN '" + pCurDir + "UPDATE.MDB' SET " + fld + " = '" + txt.Text + "' WHERE MaSo=" + CStr(Lst.ItemData(Lst.ListIndex))
        On Error GoTo 0
        Lst.List(Lst.ListIndex) = txt.Text
    End If
End Sub
