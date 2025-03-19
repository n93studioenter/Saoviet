VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FDsHD 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ho∏ Æ¨n ch≠a thanh to∏n"
   ClientHeight    =   7080
   ClientLeft      =   5205
   ClientTop       =   1230
   ClientWidth     =   11175
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
   Icon            =   "FHD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "0"
   Begin MSGrid.Grid GrdCT 
      Height          =   5775
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   10695
      _Version        =   65536
      _ExtentX        =   18865
      _ExtentY        =   10186
      _StockProps     =   77
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   1
      Cols            =   9
      FixedRows       =   0
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   7320
      Picture         =   "FHD.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "&Print"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8520
      Picture         =   "FHD.frx":6C44
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   9720
      Picture         =   "FHD.frx":80A6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thanh to∏n"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   9480
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hπn TT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   8880
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ß∑ thanh to∏n"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label LbVattu 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ten VT"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1200
      TabIndex        =   6
      Tag             =   "1"
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ßËi t≠Óng"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cﬂn nÓ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thµnh ti“n"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Di‘n gi∂i"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ngµy CT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SË CT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FDsHD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MaTK As Long
Dim f1 As Integer
Dim TongTien As Double
Dim tienx As Double
Dim Lsql As String
'====================================================================================================
' Ch‰n phi’u nhÀp
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Dim tien As Double, sono As Double
    
    Select Case Index
        Case 0:
            With GrdCT
                .col = 8
                If .Text = "" Then Exit Sub
                .col = 7
                tien = Cdbl5(.Text)
                If tien = 0 Then
                    .col = 5
                    sono = Cdbl5(.Text)
                    .col = 7
                    tien = IIf(sono <= TongTien - tienx, sono, TongTien - tienx)
                    .Text = Format(tien, Mask_0)
                    tienx = tienx + tien
                Else
                    tienx = tienx - tien
                    .Text = ""
                End If
            End With
        Case 1:
            Hide
        Case 2:
            If Len(Lsql) > 0 Then
                SetSQL "MienTru", Lsql
                SetRptInfo
                frmMain.Rpt.ReportFileName = "DSHDCTT.RPT"
                frmMain.Rpt.Formulas(3) = "KH='" + LbVattu.Caption + "'"
                frmMain.Rpt.Formulas(4) = "N='" + FrmChungtu.MedNgay(0).Text + "'"
                InBaoCaoRPT
            End If
    End Select
End Sub

Private Sub Form_Activate()
    If Me.tag > 0 Then
        Me.MousePointer = 11
        If f1 = 0 Then
            LietKePhieuNhap Me.tag
        Else
            LietKePhieuNhap2 Me.tag
        End If
        Me.tag = -Me.tag
        Me.MousePointer = 0
        HienThongBao Caption, 1
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask Then
        Select Case KeyCode
            Case vbKeyC:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyV:
                RFocus Command(1)
                Command_Click 1
        End Select
    End If
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        Hide
        On Error GoTo 0
    End If
End Sub

Private Sub Form_Load()
    ColumnSetUp GrdCT, 0, 1300, 2
    ColumnSetUp GrdCT, 1, 820, 2
    ColumnSetUp GrdCT, 2, 2620, 0
    ColumnSetUp GrdCT, 3, 1300, 1
    ColumnSetUp GrdCT, 4, 1300, 1
    ColumnSetUp GrdCT, 5, 1300, 1
    ColumnSetUp GrdCT, 6, 580, 2
    ColumnSetUp GrdCT, 7, 1300, 1
    ColumnSetUp GrdCT, 8, 1, 0
        
    Caption = Caption + " - " + CStr(pNamTC)
    
    KtraCongNo
    
    SetFont Me
End Sub

'Private Sub GrdCT_Click()
'    SendKeys "{Home}", True
'    SetGridIndex GrdCT, GrdCT.Row
'End Sub

Private Sub GrdCT_Click()
    Command_Click 0
End Sub
'======================================================================================
' Li÷t k™ phi’u nhÀp
'======================================================================================
Private Sub LietKePhieuNhap(mkh As Long)
    Dim rs_phieu As Recordset, dgia As Double
    
    Lsql = "SELECT 2000000000+MaSo AS MS,'DK' AS SoHieu,#01/01/1990# AS NgayCT,'" + ABCtoVNI("D≠ Æ«u k˙") + "' AS DienGiai,SoXuat,DuNo_0 AS SoPS,HanTT FROM CNDauNam" _
        & " WHERE (MaTaiKhoan =" + CStr(MaTK) + ") AND (MaKhachHang = " + CStr(mkh) + ") AND (DuNo_0>SoXuat)" _
        & " UNION SELECT MaCT, SoHieu, NgayCT, DienGiai, Sum(SoXuat), Sum(SoPS),HanTT FROM Chungtu" _
        & " WHERE MaTKNo =" + CStr(MaTK) + " AND (ChungTu.MaKH = " + CStr(mkh) + ")" _
        + " GROUP BY MaCT,SoHieu,NgayCT,DienGiai,HanTT HAVING Sum(SoPS)>Sum(SoXuat) ORDER BY NgayCT DESC, SoHieu DESC"
    Set rs_phieu = DBKetoan.OpenRecordset(Lsql, dbOpenSnapshot)
    Do While Not rs_phieu.EOF
        GrdCT.AddItem rs_phieu!sohieu + Chr(9) + IIf(rs_phieu!ms < 2000000000, Format(rs_phieu!NgayCT, Mask_D), "") + Chr(9) + rs_phieu!diengiai _
            + Chr(9) + Format(rs_phieu!sops, Mask_0) + Chr(9) + Format(rs_phieu!SoXuat, Mask_0) + Chr(9) _
            + Format(rs_phieu!sops - rs_phieu!SoXuat, Mask_0) + Chr(9) + CStr(rs_phieu!HanTT) + Chr(9) + "" + Chr(9) + CStr(rs_phieu!ms), 0
        rs_phieu.MoveNext
    Loop
    GrdCT.Rows = IIf(rs_phieu.RecordCount > GrdCT.tag, rs_phieu.RecordCount, GrdCT.tag)
    GrdCT.Row = 0
    GrdCT.col = 0
    rs_phieu.Close
    Set rs_phieu = Nothing
End Sub

Private Sub LietKePhieuNhap2(mkh As Long)
    Dim rs_phieu As Recordset, dgia As Double
    
    Lsql = "SELECT 2000000000+MaSo AS MS,'DK' AS SoHieu,#01/01/1990# AS NgayCT,'" + ABCtoVNI("D≠ Æ«u k˙") + "' AS DienGiai,SoXuat,DuCo_0 AS SoPS,HanTT FROM CNDauNam" _
        & " WHERE (MaTaiKhoan =" + CStr(MaTK) + ") AND (MaKhachHang = " + CStr(mkh) + ") AND (DuCo_0>SoXuat)" _
        & " UNION SELECT MaCT, SoHieu, NgayCT, DienGiai, Sum(SoXuat), Sum(SoPS),HanTT FROM Chungtu" _
        & " WHERE MaTKCo =" + CStr(MaTK) + " AND (ChungTu.MaKHC = " + CStr(mkh) + ")" _
        + " GROUP BY MaCT,SoHieu,NgayCT,DienGiai,HanTT HAVING Sum(SoPS)>Sum(SoXuat) ORDER BY NgayCT DESC, SoHieu DESC"
    Set rs_phieu = DBKetoan.OpenRecordset(Lsql, dbOpenSnapshot)
    Do While Not rs_phieu.EOF
        GrdCT.AddItem rs_phieu!sohieu + Chr(9) + IIf(rs_phieu!ms < 2000000000, Format(rs_phieu!NgayCT, Mask_D), "") + Chr(9) + rs_phieu!diengiai _
            + Chr(9) + Format(rs_phieu!sops, Mask_0) + Chr(9) + Format(rs_phieu!SoXuat, Mask_0) + Chr(9) _
            + Format(rs_phieu!sops - rs_phieu!SoXuat, Mask_0) + Chr(9) + CStr(rs_phieu!HanTT) + Chr(9) + "" + Chr(9) + CStr(rs_phieu!ms), 0
        rs_phieu.MoveNext
    Loop
    GrdCT.Rows = IIf(rs_phieu.RecordCount > GrdCT.tag, rs_phieu.RecordCount, GrdCT.tag)
    GrdCT.Row = 0
    GrdCT.col = 0
    rs_phieu.Close
    Set rs_phieu = Nothing
End Sub
'======================================================================================
' Hµm tr∂ v“ phi’u nhÀp Æ≠Óc ch‰n
'======================================================================================
Public Function ThanhToanDichDanh(mtk As Long, diengiai As String, ps As Double, f As Integer, mct() As Long, tien() As Double, sott As Integer) As Integer
    Dim i As Integer, mct1 As Long, tien1 As Double
        
    f1 = f
    MaTK = mtk
    TongTien = ps
    LbVattu.Caption = diengiai
    Show 1
    sott = 0
    tienx = 0
    With GrdCT
        For i = 0 To .Rows - 1
            .Row = i
            .col = 8
            mct1 = CLng5(.Text)
            If mct1 = 0 Then Exit For
            .col = 7
            tien1 = Cdbl5(.Text)
            If tien1 > 0 Then
                sott = sott + 1
                ReDim Preserve mct(1 To sott) As Long
                ReDim Preserve tien(1 To sott) As Double
                mct(sott) = mct1
                .col = 5
                tien(sott) = tien1
            End If
        Next
    End With
    ThanhToanDichDanh = sott
    Unload Me
End Function

Private Sub GrdCT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 32 Then GrdCT_Click
End Sub

Private Sub GrdCT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdCT, GrdCT.col
    End If
End Sub
