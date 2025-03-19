VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmMauKC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Néi dung chøng tõ kÕt chuyÓn"
   ClientHeight    =   7080
   ClientLeft      =   9675
   ClientTop       =   2955
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
   ForeColor       =   &H00808000&
   Icon            =   "FrmMauKC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin MSGrid.Grid GrdNT 
      Height          =   5535
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Tag             =   "0"
      Top             =   1080
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   9763
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   30
      Cols            =   4
      FixedRows       =   0
   End
   Begin MSGrid.Grid GrdNT 
      Height          =   5535
      Index           =   1
      Left            =   5520
      TabIndex        =   13
      Top             =   1080
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   9763
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   30
      Cols            =   3
      FixedRows       =   0
   End
   Begin VB.CommandButton CmdChitiet 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9480
      TabIndex        =   11
      ToolTipText     =   "Ghi ph¸t sinh"
      Top             =   6720
      Width           =   255
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   5520
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "0"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txtNhap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Tag             =   "0"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label TK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TK1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Tag             =   "1"
      Top             =   120
      Width           =   9495
   End
   Begin VB.Label LbTK 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   10
      Tag             =   "1"
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label LbTK 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Tag             =   "0"
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ð"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   48
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   4440
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label TK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TK1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   7
      Tag             =   "1"
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chi tiÕt"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DiÔn gi¶i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label TK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TK1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Tag             =   "1"
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chi tiÕt"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DiÔn gi¶i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "FrmMauKC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim act As Integer

Private Sub CmdChitiet_Click()
    Dim i As Integer
    For i = 0 To 1
        If Len(LbTK(i).Caption) = 0 Then
            RFocus txtNhap(i)
            Exit Sub
        Else
            If txtNhap(i).tag > 0 And CLng5(Left(txtNhap(i).Text, Len(CStr(TK(i).tag)))) = TK(i).tag Then
                MsgBox "H·y nhËp chi tiÕt tµi kho¶n " + TK(i).Caption, vbCritical, App.ProductName
                RFocus txtNhap(i)
                Exit Sub
            End If
        End If
    Next
        
    If ExecuteSQL5("INSERT INTO KCChitiet (MaSo,MaKC,TK1,TK2) VALUES (" + CStr(Lng_MaxValue("MaSo", "KCChitiet") + 1) + "," + CStr(Me.tag) + "," + CStr(LbTK(0).tag) + "," + CStr(LbTK(1).tag) + ")") = 0 Then
        GrdNT(0).AddItem txtNhap(0).Text + Chr(9) + LbTK(0).Caption + Chr(9) + CStr(LbTK(0).tag) + Chr(9) + CStr(Lng_MaxValue("MaSo", "KCChitiet")), NewRowIndex(GrdNT(0), 0)
        GrdNT(1).AddItem txtNhap(1).Text + Chr(9) + LbTK(1).Caption + Chr(9) + CStr(LbTK(1).tag), NewRowIndex(GrdNT(1), 0)
        GrdNT(0).Row = GrdNT(0).Rows - 1
        If Len(GrdNT(0).Text) = 0 Then
            GrdNT(0).RemoveItem GrdNT(0).Rows - 1
            GrdNT(1).RemoveItem GrdNT(1).Rows - 1
        End If
    End If
End Sub

Private Sub Form_Activate()
    Dim rs As Recordset, s As String, T As String
    
    If act = 0 Then
        Set rs = DBKetoan.OpenRecordset("SELECT TK1 AS M1, HeThongTK.SoHieu AS S1, HeThongTK.Ten AS T1, TK2 AS M2, HeThongTK_1.SoHieu AS S2, HeThongTK_1.Ten AS T2, KCChitiet.MaSo " _
            & "FROM (KCChitiet LEFT JOIN HeThongTK ON KCChitiet.TK1 = HeThongTK.MaSo) LEFT JOIN HeThongTK AS HeThongTK_1 ON KCChitiet.TK2 = HeThongTK_1.MaSo " _
            & "WHERE KCChitiet.MaKC=" + CStr(Me.tag), dbOpenSnapshot, dbForwardOnly)
        Do While Not rs.EOF
            If Not IsNull(rs!s1) Then s = rs!s1 Else s = "..."
            If Not IsNull(rs!T1) Then T = rs!T1 Else T = "..."
            GrdNT(0).AddItem s + Chr(9) + T + Chr(9) + CStr(rs!m1) + Chr(9) + CStr(rs!MaSo), 0
            If Not IsNull(rs!s2) Then s = rs!s2 Else s = "..."
            If Not IsNull(rs!t2) Then T = rs!t2 Else T = "..."
            GrdNT(1).AddItem s + Chr(9) + T + Chr(9) + CStr(rs!m2), 0
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        act = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    ColumnSetUp GrdNT(0), 0, 1300, 0
    ColumnSetUp GrdNT(1), 0, 1300, 0
    ColumnSetUp GrdNT(0), 1, 2620, 0
    ColumnSetUp GrdNT(1), 1, 2620, 0
    ColumnSetUp GrdNT(0), 2, 1, 0
    ColumnSetUp GrdNT(1), 2, 1, 0
    ColumnSetUp GrdNT(0), 3, 1, 0
    
    SetFont Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    act = 0
End Sub

Private Sub GrdNT_DblClick(Index As Integer)
    If Len(GrdNT(0).Text) > 0 Then
        GrdNT(0).col = 3
        If ExecuteSQL5("DELETE FROM KCChitiet WHERE MaSo=" + GrdNT(0).Text) = 0 Then
            GrdNT(1).RemoveItem GrdNT(0).Row
            GrdNT(0).RemoveItem GrdNT(0).Row
            'MsgBox GrdNT(0).tag
            If GrdNT(0).Rows < GrdNT(0).tag Then
                GrdNT(0).Rows = GrdNT(0).tag
                GrdNT(1).Rows = GrdNT(1).tag
            End If
        End If
    End If
End Sub

Private Sub GrdNT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdNT(Index), GrdNT(Index).col
    End If
End Sub

Private Sub txtNhap_GotFocus(Index As Integer)
    AutoSelect txtNhap(Index)
End Sub

Private Sub txtNhap_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.MousePointer = 11
        txtNhap(Index).Text = FrmTaikhoan.ChonTk(txtNhap(Index).Text)
        Me.MousePointer = 0
        txtNhap_LostFocus Index
    End If
End Sub

Private Sub txtNhap_LostFocus(Index As Integer)
    Dim mtk As Long
    
    LbTK(Index).Caption = tentk(txtNhap(Index).Text, mtk, txtNhap(Index).tag)
    LbTK(Index).tag = mtk
End Sub
