VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form GetDate2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   555
   ClientLeft      =   3135
   ClientTop       =   3330
   ClientWidth     =   3675
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
   Icon            =   "GetDate2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "®Õn ngµy"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Tag             =   "to"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tõ ngµy"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Tag             =   "From"
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "GetDate2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ngay(0 To 1) As Date
Dim esc As Integer

Public Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27:
            esc = 1
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Dim chi_so As Integer
    esc = 0
    For chi_so = 0 To 1
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next
End Sub

Private Sub MedNgay_GotFocus(Index As Integer)
    AutoSelect MedNgay(Index)
End Sub

Public Sub MedNgay_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Index
            Case 0: RFocus MedNgay(1)
            Case 1:
                MedNgay_LostFocus 1
                Unload Me
        End Select
    End If
End Sub

Private Sub MedNgay_LostFocus(Index As Integer)
    If IsDate(MedNgay(Index).Text) Then
        ngay(Index) = CVDate(MedNgay(Index).Text)
    Else
        RFocus MedNgay(Index)
    End If
End Sub

Public Function GetDate(s As String, d1 As Date, d2 As Date, Optional setndau As Integer = 0) As Boolean
    Me.Caption = s
    If setndau > 0 Then
        ngay(0) = d1
        MedNgay(0).Text = Format(d1, Mask_D)
        MedNgay(0).Enabled = False
    End If
    Me.Show vbModal
    If esc = 1 Then
        GetDate = False
    Else
        d1 = ngay(0)
        d2 = ngay(1)
        GetDate = True
    End If
End Function

Public Function GetDate1(s As String, d1 As Date) As Boolean
    Me.Caption = s
    MedNgay(1).Visible = False
    Me.Show vbModal
    If esc = 1 Then
        GetDate1 = False
    Else
        d1 = ngay(0)
        GetDate1 = True
    End If
End Function

