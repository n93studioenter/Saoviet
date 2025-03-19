VERSION 5.00
Begin VB.Form FrmBaoCao 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch c¸c mÉu b¸o c¸o tù t¹o"
   ClientHeight    =   7080
   ClientLeft      =   750
   ClientTop       =   930
   ClientWidth     =   9855
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
   Icon            =   "FBaoCao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "User defined Reports"
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   4
      Left            =   8640
      Picture         =   "FBaoCao.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "&View"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CmdChitiet 
      Caption         =   "­"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8280
      TabIndex        =   19
      Tag             =   "-1"
      ToolTipText     =   "Ghi ph¸t sinh"
      Top             =   6720
      Width           =   255
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      Left            =   840
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   4200
      MaxLength       =   80
      TabIndex        =   16
      Tag             =   "0"
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   5280
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "X"
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6960
      MaxLength       =   80
      TabIndex        =   7
      Tag             =   "0"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   5640
      MaxLength       =   80
      TabIndex        =   6
      Tag             =   "0"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.PictureBox Grd 
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5595
      ScaleWidth      =   8355
      TabIndex        =   10
      Tag             =   "30"
      Top             =   1080
      Width           =   8415
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   8640
      Picture         =   "FBaoCao.frx":06A4
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "&Return"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   8640
      Picture         =   "FBaoCao.frx":09E6
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "&Delete"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   8640
      Picture         =   "FBaoCao.frx":0D28
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "&Add"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtNhap 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   600
      MaxLength       =   80
      TabIndex        =   4
      Top             =   6720
      Width           =   4695
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "0"
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Néi dung chi tiÕt:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Tag             =   "Content:"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label 
      Caption         =   "Sè hiÖu"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Tag             =   "Code"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "Tªn b¸o c¸o"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   14
      Tag             =   "Name"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TH"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   13
      Tag             =   "S"
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tµi kho¶n cã"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6960
      TabIndex        =   12
      Tag             =   "Credit Account"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tµi kho¶n nî"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   11
      Tag             =   "Debit Account"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DiÔn gi¶i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Tag             =   "Desciption"
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Tag             =   "Ord."
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "FrmBaoCao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ThemMoi As Integer

Private Sub Cbo_Click()
    NoiDungBaoCao Cbo.Text
End Sub

'====================================================================================================
' Thªm, Ghi, Xãa nguyªn tÖ
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0:
            XoaManHinh
            ThemMoi = 1
            RFocus Cbo
        Case 2:
            If Cbo.ListIndex >= 0 Then
                ExecuteSQL5 "DELETE * FROM BaoCao WHERE SoHieu='" + Cbo.Text + "'"
                Cbo.RemoveItem Cbo.ListIndex
            End If
            XoaManHinh
            RFocus Cbo
        Case 3:
            Unload Me
        Case 4:
            If Cbo.ListIndex >= 0 Then XemBaoCao Cbo.Text, Date, Date
    End Select
End Sub

'====================================================================================================
' Xö lý phÝm nãng
'====================================================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyT:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyG:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyX:
                RFocus Command(2)
                Command_Click 2
            Case vbKeyV:
                RFocus Command(3)
                Command_Click 3
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
'====================================================================================================
' Khëi t¹o cöa sæ
'====================================================================================================
Private Sub Form_Load()
    ColumnSetUp Grd, 0, 465, 2
    ColumnSetUp Grd, 1, 4660, 0
    ColumnSetUp Grd, 2, 340, 2
    ColumnSetUp Grd, 3, 1300, 0
    ColumnSetUp Grd, 4, 1300, 0
    ColumnSetUp Grd, 5, 1, 0
    
    Caption = Caption + " - " + CStr(pNamTC)
    
    Int_RecsetToCbo "SELECT DISTINCTROW 0 AS F2,SoHieu As F1 FROM BaoCao WHERE STT=0 GROUP BY SoHieu", Cbo
    
    SetFont Me
    
    Command_Click 0
End Sub

Private Sub Grd_click()
    Dim i As Integer
    
    SendKeys "{Home}", True
    SetGridIndex Grd, Grd.Row
    
    With Grd
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        For i = 0 To 4
            .col = i
            txtNhap(i).Text = .Text
        Next
        ThemMoi = 0
    End With
End Sub

Private Sub Grd_KeyPress(KeyAscii As Integer)
    SendKeys "{Home}", True
    SetGridIndex Grd, Grd.Row
    
    If KeyAscii = 13 Then Grd_click
End Sub

Private Sub txtNhap_GotFocus(Index As Integer)
    AutoSelect txtNhap(Index)
End Sub
'====================================================================================================
' HiÓn thÞ danh s¸ch nguyªn tÖ
'====================================================================================================
Private Sub NoiDungBaoCao(sh As String)
    Dim rs As Recordset
    
    XoaManHinh
    Set rs = DBKetoan.OpenRecordset("SELECT * FROM BaoCao WHERE SoHieu='" + sh + "' ORDER BY STT DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        Grd.AddItem IIf(rs!stt < 10, "0", "") + CStr(rs!stt) + Chr(9) + rs!diengiai + Chr(9) + rs!shtkno + Chr(9) + rs!shtkco, 0
        rs.MoveNext
    Loop
    Grd.Rows = IIf(rs.RecordCount > Grd.Tag, rs.RecordCount, Grd.Tag)
    Grd.Row = 0
    Grd.col = 0
    rs.Close
    Set rs = Nothing
    
    Grd_click
End Sub

Private Sub txtNhap_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0:         KeyProcess txtNhap(Index), KeyAscii
        Case 2:         txtNhap(Index).Text = IIf(Len(txtNhap(Index).Text) > 0, "", "X")
                                KeyAscii = 0
        Case 3, 4:
            If KeyAscii = vbKeyReturn Then
                Me.MousePointer = 11
                txtNhap(Index).Text = FrmTaikhoan.ChonTk(txtNhap(Index).Text)
                Me.MousePointer = 0
                txtNhap_LostFocus Index
            End If
    End Select
End Sub

Private Sub txtNhap_LostFocus(Index As Integer)
    Select Case Index
        Case 0:
            txtNhap(0).Text = Format(txtNhap(0).Text, Mask_0)
            If Len(txtNhap(0).Text) < 2 Then txtNhap(0).Text = "0" + txtNhap(0).Text
        Case 1, 3, 4, 5:
            If Len(txtNhap(Index).Text) = 0 Then txtNhap(Index).Text = "..."
    End Select
End Sub

Private Sub XoaManHinh()
    Dim i As Integer
    
    Cbo.ListIndex = -1
    txtNhap(0).Text = "1"
    For i = 1 To txtNhap.Count - 1
        txtNhap(i).Text = "..."
    Next
    ClearGrid Grd, Grd.Tag
End Sub
