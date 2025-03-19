VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch chøng tõ"
   ClientHeight    =   7965
   ClientLeft      =   5700
   ClientTop       =   510
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
   Icon            =   "FrmA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7965
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Voucher List"
   Begin MSGrid.Grid GrdChungtu 
      Height          =   6855
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   9375
      _Version        =   65536
      _ExtentX        =   16536
      _ExtentY        =   12091
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
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H8000000E&
      Caption         =   "Vµo &sæ"
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
      Index           =   1
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Xem b¸o c¸o"
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   8400
      Picture         =   "FrmA.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "&Return"
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label LbThang 
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vµo sæ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8520
      TabIndex        =   11
      Tag             =   "B. Date"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sè chøng tõ"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   9
      Tag             =   "Number of Voucher"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label LbSoCT 
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tæng Ph¸t sinh"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7200
      TabIndex        =   7
      Tag             =   "Amount"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Th¸ng"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   5
      Tag             =   "to"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DiÔn gi¶i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   4
      Tag             =   "Description"
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ngµy GS"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Tag             =   "B. Date"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ngµy CT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Tag             =   "V. Date"
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè CT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Tag             =   "V. Code"
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FrmA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileName As String
Dim thang As Integer
Dim act As Integer
'======================================================================================
' LiÖt kª, Chän chøng tõ
'======================================================================================
Private Sub Command_Click(Index As Integer)
    act = Index
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
            Select Case KeyCode
                Case vbKeyS:
                    RFocus Command(0)
                    Command_Click 0
                Case vbKeyV:
                    RFocus Command(1)
                    Command_Click 1
            End Select
    End If
    If KeyCode = vbKeyEscape Then Command_Click 1
End Sub
'======================================================================================
' Khëi t¹o cöa sæ
'======================================================================================
Private Sub Form_Load()
    ColumnSetUp GrdChungtu, 0, 1300, 2
    ColumnSetUp GrdChungtu, 1, 940, 2
    ColumnSetUp GrdChungtu, 2, 940, 2
    ColumnSetUp GrdChungtu, 3, 3820, 0
    ColumnSetUp GrdChungtu, 4, 1300, 1
    ColumnSetUp GrdChungtu, 5, 940, 2
    ColumnSetUp GrdChungtu, 6, 1, 2
        
    Caption = Caption + " - " + CStr(pNamTC)
    
    SetFont Me
End Sub
'======================================================================================
' Nhí m· sè CT ®­îc chän
'======================================================================================
Private Sub GrdChungtu_Click()
    With GrdChungtu
        .col = 5
        If Len(.Text) = 0 Then Exit Sub
        If .Text = "X" Then
            .Text = "-"
        Else
            If KtraChon Then .Text = "X" Else MsgBox "Kh«ng chän qu¸ " + CStr(Max1000) + " chøng tõ!", vbCritical, App.ProductName
        End If
    End With
End Sub

Private Sub GrdChungtu_KeyPress(KeyAscii As Integer)
    GrdChungtu_Click
    If KeyAscii = 13 Then Command_Click 1
End Sub
'======================================================================================
' LiÖt kª chøng tõ
'======================================================================================
Private Sub LietKeChungtu(thang As Integer, ndau As Date, ncuoi As Date)
    Dim i As Integer, mct As Long, uid As Long, mct1 As Long, mloai As Integer
    Dim rs_chungtu As Recordset, st As String, ovr As Integer
    Dim db As Database, sql As String
    
    Me.MousePointer = 11
    
    On Error GoTo KT
    Set db = WSpace.OpenDatabase(FileName, False, False, ";PWD=" + pPSW)
    db.QueryDefs("MienTru").sql = "SELECT MaCT,Sum(IIF(MaTKNo>0,SoPS,0)) AS TPS FROM ChungTu WHERE " + IIf(thang > 0, WThang("ThangCT", thang, thang), WNgay("NgayCT", ndau, ncuoi)) + " GROUP BY MaCT"
    
    sql = "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.SoHieu, NgayCT, NgayGS, ChungTu.DienGiai" + IIf(pNN = 1, "E", "") + " AS DG,TPS,User_ID,ChungTu.MaLoai,Max(IIF(IsNull(HoaDon.MaSo),'-','X')) AS Chon FROM (ChungTu INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) LEFT JOIN HoaDon ON ChungTu.MaSo = HoaDon.MaSo " _
        & " GROUP BY ChungTu.MaCT,User_ID,ChungTu.MaLoai,ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(pNN = 1, "E", "") + ",TPS,MaLoai " _
        & " ORDER BY NgayGS DESC, ChungTu.SoHieu DESC"
    
    ClearGrid GrdChungtu, GrdChungtu.tag
    Set rs_chungtu = db.OpenRecordset(sql, dbOpenSnapshot)
    On Error GoTo 0
    Do While Not rs_chungtu.EOF
        If mct <> rs_chungtu!MaCT Then
            mct = rs_chungtu!MaCT
            uid = rs_chungtu!User_ID
        Else
            If uid <> rs_chungtu!User_ID Or mloai <> rs_chungtu!maloai Then
                mct1 = Lng_MaxValue("MaCT", "ChungTu") + 1
                ExecuteSQL5 "UPDATE ChungTu SET MaCT=" + CStr(mct1) + " WHERE MaCT=" + CStr(mct) + " AND User_ID=" + CStr(rs_chungtu!User_ID)
                mct = mct1
                uid = rs_chungtu!User_ID
                mloai = rs_chungtu!maloai
            End If
        End If
        If GrdChungtu.Rows < MaxGridRow Then
            If rs_chungtu!maloai = 2 And Right(rs_chungtu!sohieu, 2) = "GV" And Len(rs_chungtu!sohieu) > 2 Then
                st = IIf(SelectSQL("SELECT TOP 1 HoaDon.MaSo AS F1 FROM HoaDon INNER JOIN ChungTu ON HoaDon.MaSo=ChungTu.MaSo WHERE ChungTu.MaLoai=8 AND ChungTu.SoHieu='" + Left(rs_chungtu!sohieu, Len(rs_chungtu!sohieu) - 2) + "' AND ChungTu.NgayCT=#" + Format(rs_chungtu!NgayCT, Mask_DB) + "#") > 0, "X", "0")
            Else
                st = rs_chungtu!chon
            End If
            GrdChungtu.AddItem rs_chungtu!sohieu + Chr(9) + Format(rs_chungtu!NgayCT, Mask_D) + Chr(9) _
                + Format(rs_chungtu!NgayGS, Mask_D) + Chr(9) + rs_chungtu!dg + Chr(9) + Format(rs_chungtu!tps, Mask_0) + Chr(9) + st + Chr(9) + CStr(mct), 0
        Else
            ovr = 1
        End If
        rs_chungtu.MoveNext
    Loop
    GrdChungtu.Rows = IIf(rs_chungtu.RecordCount > GrdChungtu.tag, IIf(rs_chungtu.RecordCount > MaxGridRow, MaxGridRow, rs_chungtu.RecordCount), GrdChungtu.tag)
    LbSoCT.Caption = CStr(rs_chungtu.RecordCount)
    GrdChungtu.col = 0
    rs_chungtu.Close
    Set rs_chungtu = Nothing
    If ovr > 0 Then ErrMsg er_NhieuCT
KT:
    Set db = Nothing
    Me.MousePointer = 0
End Sub

Public Function ChonCTSoA(fn As String, thangchon As Integer, ndau As Date, ncuoi As Date, mct() As Long) As Integer
    Dim i As Integer, c As Integer
    
    FileName = fn
    act = 0
    thang = thangchon
    If thangchon > 0 Then
        LbThang.Caption = CStr(thangchon) + "/" + CStr(IIf(thang >= pThangDauKy, pNamTC, pNamTC - 1))
    Else
        LbThang.Caption = "From " + Format(ndau, Mask_D) + " to " + Format(ncuoi, Mask_D)
    End If
    LietKeChungtu thangchon, ndau, ncuoi
    Me.Show vbModal
    Me.Refresh
    
    If act = 0 Then Exit Function
    
    With GrdChungtu
        .col = 5
        For i = 0 To .Rows - 1
            .Row = i
            If Len(.Text) = 0 Then Exit For
            If .Text = "X" Then
                If c < Max1000 Then
                    c = c + 1
                    .col = 6
                    mct(c) = CLng5(.Text)
                    .col = 5
                End If
            End If
        Next
    End With
    
    ChonCTSoA = c
    Unload Me
End Function

Private Function KtraChon() As Boolean
    Dim c As Integer, i As Integer
    
    With GrdChungtu
        .col = 5
        For i = 0 To .Rows - 1
            If Len(.Text) = 0 Then Exit For
            If .Text = "X" Then
                c = c + 1
            End If
        Next
        KtraChon = (c < Max1000)
    End With
End Function
