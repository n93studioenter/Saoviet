VERSION 5.00
Begin VB.Form FrmDB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s∏ch t÷p d˜ li÷u l≠u tr˜ t˘ ÆÈng"
   ClientHeight    =   4800
   ClientLeft      =   1710
   ClientTop       =   1680
   ClientWidth     =   5265
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
   Icon            =   "FrmDB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   1
      Left            =   2880
      Picture         =   "FrmDB.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&View"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ß∆t thµnh t÷p m∆c Æﬁnh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "&Select"
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox txtMST 
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   2
      Left            =   4080
      Picture         =   "FrmDB.frx":6954
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "&Return"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   0
      Left            =   4320
      Picture         =   "FrmDB.frx":7D76
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&List"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtMST 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.ListBox LstDB 
      Height          =   2985
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   3735
   End
   Begin VB.ListBox LstDB 
      Height          =   2985
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "N®m tµi ch›nh"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "M∑ sË thu’"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "T™n t÷p"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   9
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ngµy l≠u d˜ li÷u"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "FrmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fndb As String

Private Sub Command_Click(Index As Integer)
    Dim fn As String
    
    Select Case Index
        Case 0:
            LoadDBList txtMST(0).Text, CInt5(txtMST(1).Text)
        Case 1:
            If LstDB(0).ListIndex >= 0 Then LstDB_DblClick 0
        Case 2:
            Unload Me
        Case 3:
            If LstDB(0).ListIndex >= 0 Then
                If MsgBox("Bπn Æ∑ chæc chæn muËn l y lπi t÷p backup ngµy " + LstDB(0).List(LstDB(0).ListIndex) + " lµm t÷p d˜ li÷u m∆c Æﬁnh ?", vbCritical + vbYesNo, App.ProductName) = vbYes Then
                    fn = pCurDir + "DATA\ERROR" + Format(Date, "DDMMYYYY") + ".SAS"
                    fndb = GetSetting(IniPath, "Environment", "Path", pCurDir + "DATA\KETOAN.MDB")
                    On Error Resume Next
                    DBKetoan.Close
                    Recycle fn
                    FileCopy fndb, fn
                    Recycle fndb
                    On Error GoTo 0
                    GianTepNen LstDB(1).List(LstDB(1).ListIndex), fndb
                    Unload Me
                End If
            End If
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Public Function ChonTepLuu(mst As String, nam As Integer) As String
    txtMST(0).Text = mst
    txtMST(1).Text = CStr(nam)
    LoadDBList mst, nam
    Screen.MousePointer = 0
    Me.Show 1
    ChonTepLuu = fndb
    Unload Me
End Function

Private Sub LstDB_Click(Index As Integer)
    LstDB(1 - Index).ListIndex = LstDB(Index).ListIndex
End Sub

Private Sub LoadDBList(mst As String, nam As Integer)
    Dim m As Integer, d As Integer, fn As String, m1 As Integer
    
    LstDB(0).Clear
    LstDB(1).Clear
    For m = CThangDB(Month(Date)) To 1 Step -1
        m1 = CThangFR(m)
        For d = SoNgayTrongThang(pNamTC, m) To 1 Step -1
            fn = pCurDir + "DATA\K" + IIf(m1 < 10, "0", "") + CStr(m1) + IIf(d < 10, "0", "") + CStr(d) + Right(CStr(nam), 2) + "_" + mst + ".SAS"
            If Len(Dir(fn)) > 0 Then
                LstDB(0).AddItem IIf(d < 10, "0", "") + CStr(d) + "/" + IIf(m1 < 10, "0", "") + CStr(m1) + "/" + CStr(nam)
                LstDB(1).AddItem fn
            End If
        Next
    Next
End Sub

Private Sub LstDB_DblClick(Index As Integer)
    Dim fn As String
    
    fn = LstDB(1).List(Index)
    fndb = Left(fndb, Len(fndb) - 3) + "MDB"
    
    If Len(Dir(fn)) > 0 Then
        Me.Hide
        GianTepNen fn, fndb
    Else
        fndb = ""
    End If
End Sub

Private Sub txtMST_GotFocus(Index As Integer)
    AutoSelect txtMST(Index)
End Sub
