VERSION 5.00
Begin VB.Form FrmGetStr 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
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
   ScaleHeight     =   3195
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.ListBox List 
      Height          =   1620
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   6375
   End
   Begin VB.TextBox Text 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   6375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tµi kho¶n cã theo dâi chi tiÕt"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
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
    Me.Caption = title
    Label(0).Caption = des
    Text.Text = ans
    Me.Show vbModal
    GetString = Text.Text
    Unload Me
End Function

Private Sub Form_Activate()
    If Me.tag = 0 Then
        Me.Height = 1395
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
        Case 27:    Text.Text = ""
                             Me.Hide
    End Select
End Sub

Private Sub Form_Load()
    SetFont Me
End Sub

Private Sub Text_GotFocus()
    AutoSelect Text
End Sub

Public Function GetMK(mst As String, Optional off As Integer) As Boolean
Dim Tr As Double
Dim Str As String
     
      Str = FrmOptions.Text(0).Text
      
   
               Tr = Int_StrToCode(Str)
             
    Me.Caption = "M· sè ch­¬ng tr×nh:  " + GetRandom + "  -  " + CStr(Tr)
    Label(0).Caption = "NhËp m· sè cµi ®Æt"
    Me.Show vbModal

    GetMK = Int_StrToCode(pms) + Int_StrToCode(mst) + Tr + 2006 = CLng5(Text.Text)
    
    Unload Me
   
End Function

Private Function GetRandom() As String
    Dim i As Integer, st As String
    
    pms = ""
    Randomize
    Do While Len(pms) < 12
       ' Initialize random-number generator.
       i = 0
       Do While Not ((i >= 65 And i <= 90) Or (i >= 49 And i <= 57))
            i = Int((44 * Rnd) + 47)
       Loop
       pms = pms + Chr(i)
       st = st + Chr(i) + "  "
    Loop
    GetRandom = st
End Function

