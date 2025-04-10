VERSION 5.00
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Begin VB.Form frmPhanLoaiVT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5325
   ClientLeft      =   1605
   ClientTop       =   1050
   ClientWidth     =   6615
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
   Icon            =   "PhanloaiVT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Classification"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5325
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1"
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   5400
      MaxLength       =   15
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   4
      Left            =   5400
      Picture         =   "PhanloaiVT.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "&Print"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   3
      Left            =   5400
      Picture         =   "PhanloaiVT.frx":6C44
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "&Return"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   2
      Left            =   5400
      Picture         =   "PhanloaiVT.frx":8066
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Delete"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   5400
      Picture         =   "PhanloaiVT.frx":9548
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "&Save"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   5400
      Picture         =   "PhanloaiVT.frx":A976
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "&Add"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   4920
      Width           =   3975
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      Top             =   4920
      Width           =   1155
   End
   Begin MSOutl.Outline Outline 
      Height          =   4455
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   7858
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "PhanloaiVT.frx":BED0
      Style           =   2
      PicturePlus     =   "PhanloaiVT.frx":BEEC
      PictureMinus    =   "PhanloaiVT.frx":BFE6
      PictureLeaf     =   "PhanloaiVT.frx":C0E0
      PictureOpen     =   "PhanloaiVT.frx":C1DA
      PictureClosed   =   "PhanloaiVT.frx":C2D4
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   4800
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   5280
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   5280
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "frmPhanLoaiVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbl As String, tbl1 As String
Private Type tpPhanLoai
      MaSo As Long
      TenPhanLoai As String
      sohieu As String
      vat As Integer
      plcha As Long
      plcon As Integer
      cap As Integer
      MaTK As Long
End Type
Dim PhanLoai As tpPhanLoai
Dim tmpSoHieu As String                       ' L­u c¸c th«ng tin cña cÊp trªn (®Ó thªm míi)
Dim tmpMaSo As Long
Dim tmpCap As Integer
Dim sql As String
Dim flag As Integer

Private Sub Form_Activate()
 Dim sh As String, m As Long
   m = SelectSQL("SELECT MaSo AS F1 FROM PhanLoaiKhachHang")

    If Me.tag > 0 Then
        flag = Me.tag
        Select Case Me.tag
            Case 1:
                If pNN = 0 Then Me.Caption = "Ph©n lo¹i vËt t­ vµ tµi kho¶n xuÊt kho tÝnh gia vèn"
                tbl = "PhanLoaiVattu"
                tbl1 = "Vattu"
                Text(2).Visible = True
            Case 2:
                If pNN = 0 Then Me.Caption = "Ph©n lo¹i kh¸ch hµng"
                tbl = "PhanLoaiKhachHang"
                tbl1 = "KhachHang"
            Case 3:
                If pNN = 0 Then Me.Caption = "Ph©n lo¹i C«ng tr×nh, s¶n phÈm"
                tbl = "PhanLoai154"
                tbl1 = "TP154"
            Case 4:
                If pNN = 0 Then Me.Caption = "Ph©n lo¹i Nh©n viªn b¸n hµng"
                tbl = "PhanLoaiNhanVien"
                tbl1 = "NhanVien"
        End Select
        If pNN = 0 And pKhongDau > 0 Then Me.Caption = ABCtoKDau(Me.Caption)
        Me.tag = 0
        LayDanhSachPhanLoai
    End If
End Sub

' Key Down
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyInsert Then
            Outline.ListIndex = -1
            tmpMaSo = 0
            tmpCap = 0
            tmpSoHieu = ""
            RFocus Command(0)
            DoEvents
            Command_Click (0)
      End If
      If (Shift And vbAltMask) > 0 Then
            Select Case KeyCode
                  Case vbKeyT: RFocus Command(0): DoEvents: Command_Click (0)
                  Case vbKeyG:  RFocus Command(1): DoEvents: Command_Click (1)
                  Case vbKeyX: RFocus Command(2): DoEvents: Command_Click (2)
                  Case vbKeyV:  RFocus Command(3): DoEvents: Command_Click (3)
                  Case vbKeyI:  RFocus Command(4): DoEvents: Command_Click (4)
            End Select
      End If
      If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    SetFont Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Select Case flag
        Case 1:    Unload FrmVattu
        Case 2:    Unload FrmKhachHang
        Case 3:    Unload FrmTP
        Case 4:     Unload FrmNhanVien
    End Select
    On Error GoTo 0
End Sub

'======================================================================================
' OUTLINE
'======================================================================================
' Click
Private Sub Outline_Click()
      ' L­u d÷ liÖu cho yªu cÇu thªm míi hay xo¸
      tmpMaSo = Outline.ItemData(Outline.ListIndex)
      tmpCap = Outline.indent(Outline.ListIndex)
      sql = "SELECT SoHieu AS F1 FROM " + tbl + " WHERE MaSo = " + CStr(tmpMaSo)
      tmpSoHieu = CStr(SelectSQL(sql))
End Sub
' DblClick
Private Sub Outline_DblClick()
      ' ChØ ®Þnh ®èi t­îng theo m· sè
      ChiDinh Outline.ItemData(Outline.ListIndex)
      ' KiÓm tra cÊp cña ph©n lo¹i ®­îc chän (cÊp trªn cïng kh«ng thÓ söa ®æi hay xo¸)
'      If PhanLoai.Cap = 1 Then
'            KhoiTao False
'      Else
            Text(0).Text = PhanLoai.TenPhanLoai
            Text(1).Text = PhanLoai.sohieu
            If flag = 1 Then Text(2).Text = MaSo2SoHieu(PhanLoai.MaTK, "HethongTK")
            RFocus Text(0)
 '     End If
End Sub
' GotFocus
Private Sub Outline_GotFocus()
      KhoiTao False
End Sub
' KeyDown
Private Sub Outline_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyReturn Then Outline.ListIndex = -1
End Sub
'======================================================================================
' command
'======================================================================================
Private Sub Command_Click(Index As Integer)
    Dim mpl As Long
    
      Select Case Index
            Case 0      ' Míi
                  KhoiTao Outline.ListIndex >= 0
                  If Outline.ListIndex < 0 Then RFocus Text(1)
            Case 1      ' Ghi
                  If HopLe = 0 Then
                        If PhanLoai.MaSo = 0 Then
                              If ThemMoi = 0 Then KhoiTao True
                        Else
                              Dim vi_tri As Integer
                              If SuaDoi = 0 Then
                                    CapNhatSoHieu PhanLoai.cap, PhanLoai.MaSo, PhanLoai.sohieu
                                    vi_tri = Outline.ListIndex
                                    LayDanhSachPhanLoai
                                    Outline.ListIndex = vi_tri
                                    KhoiTao False
                              End If
                        End If
                  End If
            Case 2      ' Xo¸
                If Outline.ListIndex < 0 Then Exit Sub
                If vbNo = MsgBox("Xo¸ ph©n lo¹i hiÖn t¹i", vbYesNo + vbQuestion, App.ProductName) Then Exit Sub
                If Outline.ListIndex + 1 < Outline.ListCount Then
                      If Outline.indent(Outline.ListIndex + 1) > Outline.indent(Outline.ListIndex) Then
                            Beep
                            MsgBox "VÉn cßn c¸c ph©n lo¹i cÊp d­íi", vbCritical, App.ProductName
                            Exit Sub
                      End If
                End If
                If xoa = 0 Then
                    KhoiTao False
                    tmpMaSo = 0
                    tmpSoHieu = ""
                    tmpCap = 0
                End If
            Case 3      ' Trë vÒ
                Unload Me
            Case 4:
                If Outline.ListIndex < 0 Then mpl = 0 Else mpl = Outline.ItemData(Outline.ListIndex)
                SetRptInfo
                Select Case flag
                    Case 1:                    DanhDiemVT mpl
                    Case 2:                    DanhDiemCN mpl
                End Select
                InBaoCaoRPT
      End Select
End Sub
'======================================================================================
' TEXT
'======================================================================================
' Got Focus
Private Sub Text_GotFocus(Index As Integer)
      AutoSelect Text(Index)
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 1:
            If KeyAscii = 32 Or KeyAscii = 35 Or KeyAscii = 39 Or KeyAscii = 42 Then KeyAscii = 0
        Case 2:
            If KeyAscii = 13 Then
                Text(2).Text = FrmTaikhoan.ChonTk(Text(2).Text)
                RFocus Text(2)
            End If
    End Select
End Sub

' Lost Focus
Private Sub Text_LostFocus(Index As Integer)
      If Len(Text(Index).Text) = 0 Then Text(Index).Text = "(...)"
      Select Case Index
            Case 0:  PhanLoai.TenPhanLoai = Text(0).Text
            Case 1:  PhanLoai.sohieu = Text(1).Text
            Case 2:
                            If flag = 1 Then
                                PhanLoai.MaTK = SoHieu2MaSo(Text(2).Text, "HethongTK")
                            End If
      End Select
End Sub
'======================================================================================
' FUNCTION HopLe
'======================================================================================
Private Function HopLe()
Dim thong_bao As String
Dim TK As New ClsTaikhoan

      With PhanLoai
            If Len(.TenPhanLoai) = 0 Then Text_LostFocus (0)
            If Len(.sohieu) = 0 Then Text_LostFocus (1)
            If .TenPhanLoai = "(...)" Then
                RFocus Text(0)
                thong_bao = "ThiÕu tªn ph©n lo¹i"
                GoTo Err_InValidate
            End If
            If .sohieu = "(...)" Then
                RFocus Text(1)
                thong_bao = "ThiÕu sè hiÖu ph©n lo¹i"
                GoTo Err_InValidate
            End If
            If flag = 1 And PhanLoai.MaTK = 0 Then
                RFocus Text(2)
                thong_bao = "ThiÕu sè hiÖu tµi kho¶n"
                GoTo Err_InValidate
            End If
            If flag = 1 Then
                TK.InitTaikhoanMaSo PhanLoai.MaTK
                If TK.tk_id <> TKVT_ID Then thong_bao = "Chän tµi kho¶n theo dâi chi tiÕt vËt t­": GoTo Err_InValidate
                If TK.tkcon > 0 Then thong_bao = "Chän tµi kho¶n chi tiÕt": GoTo Err_InValidate
            End If
            ' NÕu lµ thªm míi th× nhËn c¸c thuéc tÝnh cña ph©n lo¹i cÊp trªn
            If .MaSo = 0 Then
                  .cap = tmpCap + 1
                  .plcha = tmpMaSo
                  ' KiÓm tra cÊp vµ sè hiÖu
                  If (.cap > 3) Then _
                                                                                    thong_bao = "Sè cÊp v­ît qu¸ quy ®Þnh": GoTo Err_InValidate
                  If Not Left(.sohieu, Len(tmpSoHieu)) = tmpSoHieu Then thong_bao = "Sè hiÖu kh«ng ®óng quy ®Þnh": GoTo Err_InValidate
            Else
                If .plcha > 0 Then
                    Dim shieu_ctren As String
                    ' KiÓm tra sè hiÖu
                    sql = "SELECT SoHieu AS F1 FROM " + tbl + " WHERE MaSo = " + CStr(.plcha)
                    shieu_ctren = CStr(SelectSQL(sql))
                    If shieu_ctren <> "0" Then
                        If Not Left(.sohieu, Len(shieu_ctren)) = shieu_ctren Then thong_bao = "Sè hiÖu kh«ng ®óng quy ®Þnh": GoTo Err_InValidate
                    End If
                End If
            End If
      End With
      HopLe = 0
      Exit Function
Err_InValidate:
      Beep
      MsgBox thong_bao, vbCritical, App.ProductName
      HopLe = -1
End Function
'======================================================================================
' SUB ChiDinh
'======================================================================================
Private Sub ChiDinh(ma_pl As Long)
Dim rs_phanloai As Recordset
      sql = "SELECT * FROM " + tbl + " WHERE MaSo = " + CStr(ma_pl)
      Set rs_phanloai = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
            PhanLoai.MaSo = rs_phanloai!MaSo
            PhanLoai.TenPhanLoai = rs_phanloai!TenPhanLoai
            PhanLoai.sohieu = rs_phanloai!sohieu
            PhanLoai.plcha = rs_phanloai!plcha
            PhanLoai.plcon = rs_phanloai!plcon
            PhanLoai.cap = rs_phanloai!cap
            If flag = 1 Then PhanLoai.MaTK = rs_phanloai!MaTK Else PhanLoai.MaTK = 0
      rs_phanloai.Close
      Set rs_phanloai = Nothing
End Sub
'======================================================================================
' FUNCTION ThemMoi
'======================================================================================
Private Function ThemMoi() As Integer
Dim vi_tri As Integer
      Me.MousePointer = 11
      If ExecuteSQL5("INSERT INTO " + tbl + " (MaSo,TenPhanLoai, SoHieu, PLCha,PLCon, Cap" + IIf(flag = 1, ",MaTK", "") + ") VALUES(" + CStr(Lng_MaxValue("MaSo", tbl) + 1) + ",'" _
            + PhanLoai.TenPhanLoai + "','" + PhanLoai.sohieu + "'," + CStr(PhanLoai.plcha) + "," + CStr(PhanLoai.plcon) + "," + CStr(PhanLoai.cap) + IIf(flag = 1, "," + CStr(PhanLoai.MaTK), "") + ")") = 0 Then
            PhanLoai.MaSo = Lng_MaxValue("MaSo", tbl)
            
            ExecuteSQL5 "UPDATE " + tbl1 + " SET MaPhanLoai=" + CStr(PhanLoai.MaSo) + " WHERE MaPhanLoai=" + CStr(PhanLoai.plcha)
            ExecuteSQL5 "UPDATE " + tbl + " SET PLCon=1 WHERE MaSo=" + CStr(PhanLoai.plcha)
            Do                                                                                ' Thªm vµo vÞ trÝ cuèi cïng trong cÊp
                  vi_tri = vi_tri + 1
                  If Outline.ListIndex + vi_tri = Outline.ListCount Then Exit Do
            Loop Until Outline.indent(Outline.ListIndex + vi_tri) < PhanLoai.cap
            Outline.AddItem PhanLoai.sohieu + "  " + PhanLoai.TenPhanLoai + IIf(flag = 1, " - " + Text(2).Text, ""), Outline.ListIndex + vi_tri
            Outline.indent(Outline.ListIndex + vi_tri) = PhanLoai.cap
            Outline.ItemData(Outline.ListIndex + vi_tri) = PhanLoai.MaSo
            
            If Outline.ListIndex >= 0 Then Outline.Expand(Outline.ListIndex) = True      ' Expand cÊp trªn
            ThemMoi = 0
      Else
            PhanLoai.MaSo = 0
            ThemMoi = -1
      End If
      Me.MousePointer = 0
End Function
'======================================================================================
' FUNCTION SuaDoi
'======================================================================================
Private Function SuaDoi()
      Me.MousePointer = 11
      
      sql = "UPDATE " + tbl + " SET TenPhanLoai = '" + PhanLoai.TenPhanLoai _
                                                                                      + "', SoHieu = '" + PhanLoai.sohieu _
                                                                         + "'" + IIf(flag = 1, ",MaTK=" + CStr(PhanLoai.MaTK), "") + " WHERE MaSo = " + CStr(PhanLoai.MaSo)
      If ExecuteSQL5(sql) = 0 Then
            Outline.List(Outline.ListIndex) = PhanLoai.sohieu + "  " + PhanLoai.TenPhanLoai + IIf(flag = 1, " - " + Text(2).Text, "")
            SuaDoi = 0
      Else
            SuaDoi = -1
      End If
      Me.MousePointer = 0
End Function
'======================================================================================
' FUNCTION Xoa
'======================================================================================
Private Function xoa() As Integer
      Dim mc As Long
      
      Me.MousePointer = 11
      If SelectSQL("SELECT DISTINCTROW Count(MaSo) AS F1 FROM " + tbl1 + " WHERE MaPhanLoai = " + CStr(Outline.ItemData(Outline.ListIndex))) > 0 Then
            MsgBox "Ph©n lo¹i ®· ®¨ng ký chi tiÕt, kh«ng xo¸!", vbInformation, App.ProductName
            xoa = -1
      Else
            mc = SelectSQL("SELECT DISTINCTROW PLCha AS F1 FROM " + tbl + " WHERE MaSo=" + CStr(Outline.ItemData(Outline.ListIndex)))
            sql = "DELETE * FROM " + tbl + " WHERE MaSo = " + CStr(Outline.ItemData(Outline.ListIndex))
            If ExecuteSQL5(sql) = 0 Then
                  ExecuteSQL5 "UPDATE " + tbl1 + " SET MaPhanLoai = " + CStr(mc) + " WHERE MaPhanLoai = " + CStr(Outline.ItemData(Outline.ListIndex))
                  Outline.RemoveItem Outline.ListIndex
                  xoa = 0
            Else
                  xoa = -1
            End If
      End If
      Me.MousePointer = 0
End Function
'======================================================================================
' SUB CapNhatSoHieu
'======================================================================================
Private Sub CapNhatSoHieu(cap_ct As Integer, maso_ct As Long, sohieu_moi As String)
Dim do_dai As Integer
      Me.MousePointer = 11
      do_dai = Len(tmpSoHieu)
      If cap_ct = 1 Then
            ExecuteSQL5 "UPDATE " + tbl + " SET SoHieu = '" + sohieu_moi _
                  + "' + Right(SoHieu,Len(SoHieu) - " + CStr(do_dai) + ") WHERE PLCha = " + CStr(maso_ct)
      End If
      tmpSoHieu = sohieu_moi
      Me.MousePointer = 1
End Sub
'======================================================================================
' SUB LayDanhSachPhanLoai
'======================================================================================
Private Sub LayDanhSachPhanLoai()
Dim rs_danhsach As Recordset
Dim chi_so As Integer
      Me.MousePointer = 11
      If Outline.ListCount > 0 Then Outline.Clear
      If flag = 1 Then
        sql = "SELECT PhanLoaiVattu.*,HethongTK.SoHieu AS SHTK FROM PhanLoaiVattu INNER JOIN HethongTK ON PhanLoaiVattu.MaTK=HethongTK.MaSo WHERE PhanLoaiVattu.Cap > 0 ORDER BY PhanLoaiVattu.SoHieu"
      Else
        sql = "SELECT * FROM " + tbl + " WHERE Cap > 0 AND LEFT(SoHieu,1)<>'#' ORDER BY SoHieu"
       End If
      Set rs_danhsach = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      Do Until rs_danhsach.EOF
            If flag = 1 Then
                Outline.AddItem rs_danhsach!sohieu + Chr(9) + rs_danhsach!TenPhanLoai + Chr(9) + rs_danhsach!shtk
            Else
                Outline.AddItem rs_danhsach!sohieu + Chr(9) + rs_danhsach!TenPhanLoai
            End If
            On Error Resume Next
            Outline.indent(chi_so) = rs_danhsach!cap
            On Error GoTo 0
            Outline.ItemData(chi_so) = rs_danhsach!MaSo
            chi_so = chi_so + 1
            rs_danhsach.MoveNext
      Loop
      rs_danhsach.Close
      Set rs_danhsach = Nothing
      For chi_so = 0 To (Outline.ListCount - 1)
            If Outline.HasSubItems(chi_so) Then Outline.Expand(chi_so) = True
      Next
      Me.MousePointer = 0
End Sub
'======================================================================================
' SUB KhoiTao
'======================================================================================
Private Sub KhoiTao(tiep_tuc As Boolean)
      PhanLoai.MaSo = 0
      PhanLoai.TenPhanLoai = ""
      PhanLoai.sohieu = ""
      PhanLoai.plcha = 0
      PhanLoai.plcon = 0
      PhanLoai.cap = 0
      Text(0).Text = ""
      If tiep_tuc = True Then
            Text(1).Text = tmpSoHieu
            If flag = 1 And Outline.ListIndex >= 0 Then
                Text(2).Text = SelectSQL("SELECT HethongTK.SoHieu AS F1, MaTK AS F2 FROM PhanLoaiVattu INNER JOIN HethongTK ON PhanLoaiVattu.MaTK=HethongTK.MaSo WHERE PhanLoaiVattu.MaSo=" + CStr(Outline.ItemData(Outline.ListIndex)), PhanLoai.MaTK)
            End If
            RFocus Text(1)
            'SendKeys "{END}"
      Else
            Text(1).Text = ""
      End If
End Sub

