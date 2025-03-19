VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FDsNhap 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Phi’u nhÀp vÀt t≠ ch≠a t›nh gi∏ xu t"
   ClientHeight    =   7080
   ClientLeft      =   1245
   ClientTop       =   1125
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
   Icon            =   "Fdsnhap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Import Voucher List"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.CommandButton Command 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   7200
      Picture         =   "Fdsnhap.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "&Select"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   8400
      Picture         =   "Fdsnhap.frx":6C44
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "&Return"
      Top             =   6600
      Width           =   1095
   End
   Begin MSGrid.Grid GrdCT 
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Tag             =   "30"
      Top             =   720
      Width           =   9615
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   10186
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
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      MouseIcon       =   "Fdsnhap.frx":8066
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thµnh ti“n USD"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   9480
      TabIndex        =   11
      Tag             =   "Amount"
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ß¨n gi∏"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   10
      Tag             =   "Unit price"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label LbVattu 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ten VT"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Tag             =   "1"
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "VÀt t≠"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Tag             =   "Material"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thµnh ti“n"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   8160
      TabIndex        =   4
      Tag             =   "Amount"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SË l≠Óng"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   3
      Tag             =   "Quantity"
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
      Tag             =   "Description"
      Top             =   480
      Width           =   3255
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
      Tag             =   "Date"
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
      Tag             =   "Invoice No."
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FDsNhap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MaSo As Long
Dim SoLuong As Double
Dim ThanhTien As Double
Dim ThanhTien2 As Double
Dim ThangVT As Integer
Dim KhoVT As Long
'====================================================================================================
' Ch‰n phi’u nhÀp
'====================================================================================================
Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 0:
            With GrdCT
                .col = 6
                If Len(.Text) = 0 Then Exit Sub
                MaSo = CLng5(.Text)
                .col = 5
                ThanhTien = Cdbl5(.Text)
                .col = 6
                ThanhTien2 = Cdbl5(.Text)
                .col = 3
                SoLuong = Cdbl5(.Text)
                Unload Me
            End With
        Case 1:
            Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    If Me.tag > 0 Then
        Me.MousePointer = 11
        LietKePhieuNhap Me.tag
        Me.tag = -Me.tag
        Me.MousePointer = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyC:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyV:
                RFocus Command(1)
                Command_Click 1
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    ColumnSetUp GrdCT, 0, 1300, 2
    ColumnSetUp GrdCT, 1, 820, 2
    ColumnSetUp GrdCT, 2, 3220, 0
    ColumnSetUp GrdCT, 3, 1300, 1
    ColumnSetUp GrdCT, 4, 1300, 1
    ColumnSetUp GrdCT, 5, 1300, 1
    ColumnSetUp GrdCT, 6, 1, 0
    ColumnSetUp GrdCT, 7, 1300, 1
    
    MaSo = 0
    SoLuong = 0
    ThanhTien = 0
    Caption = Caption + " - " + CStr(pNamTC)
    
    If (pGiaUSD > 0) Then
        Label(7).Visible = True
        Me.Width = Me.Width + 1300
        FCenter Me
    End If
    
    SetFont Me
End Sub

Private Sub GrdCT_Click()
    SendKeys "{Home}", True
    SetGridIndex GrdCT, GrdCT.Row
End Sub

Private Sub GrdCT_DblClick()
    Command_Click 0
End Sub
'======================================================================================
' Li÷t k™ phi’u nhÀp
'======================================================================================
Private Sub LietKePhieuNhap(mvt As Long)
    Dim rs_phieu As Recordset, sql As String, dgia As Double
    
    SetSQL "QNhatKy", "SELECT MaCT,Sum(SoPS) AS SPS,MaVattu" + IIf(pGiaUSD > 0, ",Sum(PSUSD) AS SPS2", "") + " FROM " + ChungTu2TKNC(-1) + " WHERE MaLoai = 1 AND MaKho =" + CStr(KhoVT) + " AND (ThangCT <= " + CStr(ThangVT) + ") AND (MaVattu = " + CStr(mvt) + ") AND SoPS2No=0 AND TK_ID=" + CStr(TKVT_ID) + " GROUP BY MaCT,MaVattu"
    
    sql = "SELECT 2000000000+MaSo AS MS,'DK' AS SoHieu,#01/01/1990# AS NgayCT,'" + ABCtoVNI("SË tÂn kho Æ«u k˙") + "' AS DienGiai,Luong_0 AS SoPS2No,SoXuat,Tien_0 AS SoPS" + IIf(pGiaUSD > 0, ",USDTien_0 AS SoPS2", "") + " FROM VTDauNam" _
        & " WHERE (MaSoKho =" + CStr(KhoVT) + ") AND (MaVattu = " + CStr(mvt) + ") AND (Luong_0>SoXuat)" _
        & " UNION SELECT MaSo, SoHieu, NgayCT, DienGiai, IIF(MaTKNo>0,SoPS2No,SoPS2Co), SoXuat, SoPS+IIF(IsNull(SPS),0,SPS)" + IIf(pGiaUSD > 0, ",PSUSD+IIF(IsNull(SPS2),0,SPS2)", "") + " FROM Chungtu LEFT JOIN QNhatKy ON ChungTu.MaCT=QNhatKy.MaCT AND ChungTu.MaVattu=QNhatKy.MaVattu" _
        & " WHERE ((MaLoai = 1 AND MaKho =" + CStr(KhoVT) + ") OR (MaLoai = 4 AND MaNguon = " + CStr(KhoVT) + ")) AND (ThangCT <= " + CStr(ThangVT) + ") AND (ChungTu.MaVattu = " + CStr(mvt) + ") AND (IIF(MaTKNo>0,SoPS2No,SoPS2Co)>SoXuat)" _
        + IIf(OutCost < 3, " ORDER BY NgayCT DESC, SoHieu DESC", " ORDER BY NgayCT, SoHieu")
    
    Set rs_phieu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    
    Do While Not rs_phieu.EOF
        If rs_phieu!SoPS2No <> 0 Then dgia = Fix(0.5 + Mask_N * rs_phieu!sops / rs_phieu!SoPS2No) / Mask_N Else dgia = 0
        If pGiaUSD > 0 Then sql = Chr(9) + Format(rs_phieu!SoPS2, Mask_2) Else sql = ""
        GrdCT.AddItem rs_phieu!sohieu + Chr(9) + IIf(rs_phieu!ms < 2000000000, Format(rs_phieu!NgayCT, Mask_D), "") + Chr(9) + rs_phieu!diengiai _
            + Chr(9) + Format(rs_phieu!SoPS2No - rs_phieu!SoXuat, Mask_2) + Chr(9) + Format(dgia, Mask_2) + Chr(9) _
            + Format(((rs_phieu!SoPS2No - rs_phieu!SoXuat) * rs_phieu!sops / rs_phieu!SoPS2No), Mask_0) + Chr(9) + CStr(rs_phieu!ms) + sql, 0
        rs_phieu.MoveNext
    Loop
    GrdCT.Rows = IIf(rs_phieu.RecordCount > GrdCT.tag, rs_phieu.RecordCount, GrdCT.tag)
    GrdCT.Row = 0
    GrdCT.col = 0
    rs_phieu.Close
    Set rs_phieu = Nothing
    GrdCT_Click
End Sub
'======================================================================================
' Hµm tr∂ v“ phi’u nhÀp Æ≠Óc ch‰n
'======================================================================================
Public Function XuatDichDanh(thang As Integer, diengiai As String, mkho As Long, luong As Double, tien As Double, Optional tien2 As Double) As Long
    ThangVT = thang
    KhoVT = mkho
    luong = 0
    tien = 0
    LbVattu.Caption = diengiai
    Show 1
    luong = SoLuong
    tien = ThanhTien
    tien2 = ThanhTien2
    XuatDichDanh = MaSo
End Function

Private Sub GrdCT_KeyPress(KeyAscii As Integer)
    SendKeys "{Home}", True
    SetGridIndex GrdCT, GrdCT.Row
    If KeyAscii = 13 Then GrdCT_DblClick
End Sub

Private Sub GrdCT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdCT, GrdCT.col
    End If
End Sub
