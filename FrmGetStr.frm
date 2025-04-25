VERSION 5.00
Begin VB.Form FrmGetStr 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3975
   ClientLeft      =   2430
   ClientTop       =   3870
   ClientWidth     =   6795
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
   Icon            =   "FrmGetStr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   3
      Left            =   4560
      MaxLength       =   4
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   3120
      MaxLength       =   4
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox List 
      Height          =   1035
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   6375
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   240
      MaxLength       =   100
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   3720
      Width           =   4575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Liªn hÖ nhµ cung cÊp ®Ó cã m· sè cµi ®Æt"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "FrmGetStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pms As String

Public Function GetString(des As String, title As String, Optional ans As String) As String
 Text(1).Visible = False
 Text(2).Visible = False
 Text(3).Visible = False
 Command1.Visible = False
  Text(0).MaxLength = 500
 Text(0).Width = 6375
 Label(1).Visible = False
    Me.Caption = title
    Label(0).Caption = des
    Text(0).Text = ans
    Me.Show vbModal
    GetString = Text(0).Text
    Unload Me
End Function

Private Sub Command1_Click()
 Clipboard.Clear
  Clipboard.SetText Trim(Label1.Caption)
End Sub

Private Sub Form_Activate()
    If Me.tag = 0 Then
        Me.Height = 2000 ' 1395
    Else
        Select Case Me.tag
            Case 1:
                Int_RecsetToCbo "SELECT DISTINCTROW MaSo AS F2, Sohieu+' - '+Ten AS F1 FROM HethongTK WHERE TK_ID=" + CStr(TKVT_ID) + " AND TKCon=0 ORDER BY SoHieu", List
            Case 2:
                Int_RecsetToCbo "SELECT DISTINCTROW MaSo AS F2, Sohieu+' - '+Ten AS F1 FROM HethongTK WHERE (TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ") AND TKCon=0 ORDER BY SoHieu", List
            Case 3:
                Int_RecsetToCbo "SELECT DISTINCTROW MaSo AS F2, Sohieu+' - '+Ten AS F1 FROM HethongTK WHERE TK_ID2=" + CStr(TKCPSX_ID) + " AND TKCon=0 ORDER BY SoHieu", List
            Case 4:
                Int_RecsetToCbo "SELECT DISTINCTROW MaSo AS F2, Sohieu+' - '+Ten AS F1 FROM HethongTK WHERE TK_ID2=" + CStr(TKDT_ID) + " AND TKCon=0 ORDER BY SoHieu", List
        End Select
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13:    Me.Hide
        Case 27:    Text(0).Text = ""
                             Me.Hide
    End Select
End Sub

Private Sub Form_Load()
    SetFont Me
End Sub

'Private Sub Text_GotFocus()
'    AutoSelect Text
'End Sub

Public Function GetMK(mst As String, Optional off As Integer) As Boolean
Dim Tr As Double
Dim str As String
Dim st As String
Dim a(5) As String
Dim TK
TK = SelectSQL("SELECT CMP as f1 from license ")
st = ""
     str = FrmOptions.Text(0).Text
     GetRandom ' lay random
     Tr = Int_StrToCodes(str + pms)
     st = pms
     a(0) = Mid(CStr(Tr), 1, 1) + Mid(st, 1, 4) '
     a(1) = Mid(CStr(Tr), 2, 1) + Mid(st, 5, 4) '
     a(2) = Mid(CStr(Tr), 3, 1) + Mid(st, 9, 4) '
     a(3) = Mid(CStr(Tr), 4, 6)
    Me.Caption = "MËt m· ch­¬ng tr×nh:       " + a(0) + " - " + a(1) + " - " + a(2) + " - " + a(3)
    Label1.Caption = mst + "*" + a(0) + " *" + a(1) + " *" + a(2) + "*" + a(3)
    Label(0).Caption = "NhËp m· sè cµi ®Æt"
    Me.Show vbModal
    GetMK = CStr(Int_StrToCodes(pms + TK) + Int_StrToCodes(mst + TK) + Tr + 2006) + CStr(Int_StrToCodes(CStr(Int_StrToCodes(pms + TK) + Int_StrToCodes(mst + TK) + Tr + 2006))) = Text(0).Text + Text(1).Text + Text(2).Text + Text(3).Text
    If (GetMK = False) Then MsgBox "KÝch ho¹t kh«ng thµnh c«ng!"
    Unload Me

End Function

Private Function GetRandom() As String
    Dim i As Integer, st As String
    
    pms = ""
    Randomize
    Do While Len(pms) < 12
       ' Initialize random-number generator.
       i = 0
       Do While Not ((i >= 65 And i <= 90 And i <> 79) Or (i >= 49 And i <= 57))
            i = Int((44 * Rnd) + 47)
       Loop
       pms = pms + Chr(i)
       st = st + Chr(i) + "  "
    Loop
    GetRandom = st
End Function
Private Function GetRandom4() As String
    Dim i As Integer, st As String
    
    pms = ""
    Randomize
    Do While Len(pms) < 4
       ' Initialize random-number generator.
       i = 0
       Do While Not ((i >= 65 And i <= 90 And i <> 79) Or (i >= 49 And i <= 57))
            i = Int((44 * Rnd) + 47)
       Loop
       pms = pms + Chr(i)
       st = st + Chr(i) + "  "
    Loop
    GetRandom4 = st
End Function

Private Sub Text_Change(Index As Integer)
Dim a() As String
 On Error GoTo Error_Handler
If (Text(1).Visible = True) Then
If (Len(Text(0).Text) > 13) Then
    a = Split(Text(0).Text, "*")
    Text(0).Text = a(0)
    Text(1).Text = a(1)
    Text(2).Text = a(2)
    Text(3).Text = a(3)
End If
End If
Error_Handler:
End Sub
