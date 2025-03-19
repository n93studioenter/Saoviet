VERSION 5.00
Begin VB.Form FrmCal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "M¸y TÝnh"
   ClientHeight    =   2550
   ClientLeft      =   2175
   ClientTop       =   3480
   ClientWidth     =   3720
   ClipControls    =   0   'False
   Icon            =   "FrmCal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   -120
      TabIndex        =   28
      Top             =   -120
      Width           =   4335
      Begin VB.Label kq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   27
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M+"
      Height          =   255
      Index           =   26
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "P"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MS"
      Height          =   255
      Index           =   25
      Left            =   120
      TabIndex        =   25
      ToolTipText     =   "M"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MR"
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "R"
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MC"
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   23
      ToolTipText     =   "C"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      Height          =   255
      Index           =   22
      Left            =   1920
      TabIndex        =   22
      ToolTipText     =   "F12"
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CE"
      Height          =   255
      Index           =   21
      Left            =   1320
      TabIndex        =   21
      ToolTipText     =   "Del"
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Back"
      Height          =   255
      Index           =   20
      Left            =   720
      TabIndex        =   20
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "="
      Height          =   255
      Index           =   19
      Left            =   3120
      TabIndex        =   19
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1/x"
      Height          =   255
      Index           =   18
      Left            =   3120
      TabIndex        =   18
      ToolTipText     =   "End"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%"
      Height          =   255
      Index           =   17
      Left            =   3120
      TabIndex        =   17
      ToolTipText     =   "\"
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "sqrt"
      Height          =   255
      Index           =   16
      Left            =   3120
      TabIndex        =   16
      ToolTipText     =   "Home"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   15
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      Height          =   255
      Index           =   14
      Left            =   2520
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*"
      Height          =   255
      Index           =   13
      Left            =   2520
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "/"
      Height          =   255
      Index           =   12
      Left            =   2520
      TabIndex        =   12
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "."
      Height          =   255
      Index           =   11
      Left            =   1920
      TabIndex        =   11
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+-"
      Height          =   255
      Index           =   10
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "PageUp / PageDown"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      Height          =   255
      Index           =   9
      Left            =   1920
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   495
   End
End
Attribute VB_Name = "FrmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Status As Integer
Dim Operator As Integer
Dim sDecimal As String
Dim r1 As Double
Dim m As Double
Dim kb As Integer

Public Function Calc() As Double
    Me.Show 1
    Calc = Cdbl5(kq.Caption)
    Unload Me
End Function

Private Function Cdbl5(st As String) As Double
    If IsNumeric(st) Then Cdbl5 = CDbl(st) Else Cdbl5 = 0
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If kb = vbKeyReturn And KeyCode = vbKeyReturn Then
        Me.Hide
        Exit Sub
    End If
    Select Case KeyCode
        Case vbKey0 To vbKey9: If (Shift And vbShiftMask) = 0 Then Label_Click KeyCode - 48
        Case vbKeyNumpad0 To vbKeyNumpad9: Label_Click KeyCode - 96
        Case vbKeyPageUp, vbKeyPageDown: Label_Click 10
        Case vbKeyDecimal, Asc(sDecimal): Label_Click 11
        Case vbKeyDivide, 191: Label_Click 12
        Case vbKeyMultiply: Label_Click 13
        Case vbKeySubtract, 189: Label_Click 14
        Case vbKeyAdd: Label_Click 15
        Case vbKeyHome: Label_Click 16
        Case 220: Label_Click 17
        Case vbKeyEnd: Label_Click 18
        Case vbKeyReturn: Label_Click 19
        Case vbKeyBack: Label_Click 20
        Case vbKeyDelete: Label_Click 21
        Case vbKeyF12: Label_Click 22
        Case vbKeyC: Label_Click 23
        Case vbKeyR: Label_Click 24
        Case vbKeyM: Label_Click 25
        Case vbKeyP: Label_Click 26
        Case vbKeyEscape:  Me.Hide
    End Select
    kb = KeyCode
End Sub

Private Sub Form_Load()
    If Cdbl5("1,5") <> 1.5 Then sDecimal = "." Else sDecimal = ","
    Operator = 0
End Sub

Private Sub Label_Click(Index As Integer)
    Dim r As Double, i As Integer
    
    For i = 0 To Label.count - 1
        Label(i).FontBold = (i = Index)
    Next
    
    r = Cdbl5(kq.Caption)
    Select Case Index
        Case 0 To 9:
            kq.Caption = IIf(Status = 0 And (r <> 0 Or InStr(kq.Caption, sDecimal) > 0), kq.Caption + CStr(Index), CStr(Index))
            Status = 0
        Case 10:
            If r = 0 Then Exit Sub
            kq.Caption = IIf(Left(kq.Caption, 1) <> "-", "-" + kq.Caption, Right(kq.Caption, Len(kq.Caption) - 1))
        Case 11:
            If InStr(kq.Caption, sDecimal) = 0 Then kq.Caption = kq.Caption + sDecimal
        Case 12, 13, 14, 15:
            If Operator > 0 Then
                Label_Click 19
                r = Cdbl5(kq.Caption)
            End If
            r1 = r
            Operator = Index - 11
            Status = 1
            GoTo d
        Case 16:
            r = Sqr(r)
            GoTo d
        Case 17:
            r = r / 100
            GoTo d
        Case 18:
            If r = 0 Then GoTo E
            r = 1 / r
            GoTo d
        Case 19:
            Select Case Operator
                Case 1:
                    If r = 0 Then GoTo E
                    r = r1 / r
                Case 2:
                    r = r1 * r
                Case 3:
                    r = r1 - r
                Case 4:
                    r = r1 + r
                Case Else
                    Exit Sub
            End Select
            Status = 1
            Operator = 0
            GoTo d
        Case 20:
            If Len(kq.Caption) > 0 Then kq.Caption = Left(kq.Caption, Len(kq.Caption) - 1)
        Case 21:
            kq.Caption = "0"
            r = 0
        Case 22:
            kq.Caption = "0"
            Operator = 0
            r = 0
            Status = 0
            Operator = 0
        Case 23:
            m = 0
            Memory
            Status = 1
        Case 24:
            r = m
            Status = 1
            GoTo d
        Case 25:
            m = r
            Memory
            Status = 1
        Case 26:
            m = m + r
            Memory
            Status = 1
    End Select
    Exit Sub
E:
    kq.Caption = "E"
d:
    kq.Caption = CStr(r)
End Sub

Private Sub Memory()
    Label(27).Caption = IIf(m <> 0, "M", "")
End Sub
