VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Begin VB.Form frmKhauHao 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ghi chøng tõ trÝch khÊu hao"
   ClientHeight    =   4200
   ClientLeft      =   5040
   ClientTop       =   2580
   ClientWidth     =   5400
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
   Icon            =   "Khauhao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4200
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Tag             =   "Depreciation"
   Begin VB.OptionButton Option 
      BackColor       =   &H00FFFFC0&
      Caption         =   "C«ng cô dông cô"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Tag             =   "Fixed Asset from Finacial Credit"
      Top             =   2280
      Width           =   2535
   End
   Begin VB.OptionButton Option 
      BackColor       =   &H00FFFFC0&
      Caption         =   "BÊt ®éng s¶n ®Çu t­"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Tag             =   "Fixed Asset from Finacial Credit"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   1
      ItemData        =   "Khauhao.frx":57E2
      Left            =   3240
      List            =   "Khauhao.frx":57E4
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar gauge 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.OptionButton Option 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tµi s¶n cè ®Þnh thuª tµi chÝnh"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Tag             =   "Fixed Asset from Finacial Credit"
      Top             =   1320
      Width           =   2535
   End
   Begin VB.OptionButton Option 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tµi s¶n cè ®Þnh v« h×nh"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Tag             =   "Intangible Fixed Asset"
      Top             =   840
      Width           =   2535
   End
   Begin VB.OptionButton Option 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tµi s¶n cè ®Þnh h÷u h×nh"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Tag             =   "Fixed Asset"
      Top             =   360
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   1
      Left            =   4080
      Picture         =   "Khauhao.frx":57E6
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "&Return"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   0
      Left            =   4080
      Picture         =   "Khauhao.frx":6C08
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "&Save"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   0
      ItemData        =   "Khauhao.frx":8036
      Left            =   1080
      List            =   "Khauhao.frx":8038
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "TrÝch theo TK chi phÝ"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Tag             =   "By Expense Account"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "®Õn th¸ng"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Tag             =   "to"
      Top             =   3480
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   3840
      X2              =   3840
      Y1              =   3240
      Y2              =   120
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   3840
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   3840
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tõ th¸ng"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Tag             =   "From"
      Top             =   3480
      Width           =   735
   End
End
Attribute VB_Name = "frmKhauHao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LoaiKhauHao As Integer

' Key Down
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      If (Shift And vbAltMask) > 0 Then
            Select Case KeyCode
                  Case vbKeyG: RFocus Command(0): Command_Click (0)
                  Case vbKeyV:  RFocus Command(1): Command_Click (1)
            End Select
      End If
      If KeyCode = vbKeyEscape Then
            Unload frmKhauHao
            Set frmKhauHao = Nothing
      End If
End Sub
' Load
Private Sub Form_Load()
Dim chi_so As Integer
      Caption = Caption + " - " + CStr(pNamTC)
      AddMonthToCbo Combo(0)
      AddMonthToCbo Combo(1)
      Option_Click 0
      
      SetFont Me
End Sub
'======================================================================================
' FUNCTION TrichKhauHao
'======================================================================================
Private Function TrichKhauHao() As Integer
Dim thg As Integer, sql As String, thgcuoi As Integer, shct As String
Dim chi_so As Integer
Dim tong_ps As Double
Dim rs_khauhao As Recordset

      Me.MousePointer = 11
      thg = CThangDB(Combo(0).ItemData(Combo(0).ListIndex))
      thgcuoi = CThangDB(Combo(1).ItemData(Combo(1).ListIndex))
      Select Case LoaiKhauHao
            Case 0                  ' KhÊu hao tµi s¶n cè ®Þnh h÷u h×nh
                  sql = "SELECT DISTINCTROW TKCha0,Sum(ThongSo.KH_NS+ThongSo.KH_TBS+ThongSo.KH_CNK+ ThongSo.KH_TD) AS TKH, HethongTK.SoHieu " _
                        & "FROM HethongTK RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON HethongTK.MaSo = ThongSo.MaDTSD " _
                        & "WHERE HethongTK.SoHieu LIKE '" + txt.Text + "*' AND ThongSo.Thang >= " + CStr(thg) + " AND ThongSo.Thang <= " + CStr(thgcuoi) + " AND Mid(LoaiTaiSan.SoHieu,3,1) = '1' GROUP BY HethongTK.SoHieu,TKCha0"
                  thg = 1
            Case 1                  ' KhÊu hao tµi s¶n cè ®Þnh v« h×nh
                  sql = "SELECT DISTINCTROW TKCha0,Sum(ThongSo.KH_NS+ThongSo.KH_TBS+ThongSo.KH_CNK+ ThongSo.KH_TD) AS TKH, HethongTK.SoHieu " _
                        & "FROM HethongTK RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON HethongTK.MaSo = ThongSo.MaDTSD " _
                        & "WHERE HethongTK.SoHieu LIKE '" + txt.Text + "*' AND ThongSo.Thang >=" + CStr(thg) + " AND ThongSo.Thang <= " + CStr(thgcuoi) + " AND Mid(LoaiTaiSan.SoHieu,3,1) = '3' GROUP BY HethongTK.SoHieu,TKCha0"
                  thg = 3
            Case 2                  ' KhÊu hao tµi s¶n cè ®Þnh thuª tµi chÝnh
                  sql = "SELECT DISTINCTROW TKCha0,Sum(ThongSo.KH_NS+ThongSo.KH_TBS+ThongSo.KH_CNK+ ThongSo.KH_TD) AS TKH, HethongTK.SoHieu " _
                        & "FROM HethongTK RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON HethongTK.MaSo = ThongSo.MaDTSD " _
                        & "WHERE HethongTK.SoHieu LIKE '" + txt.Text + "*' AND ThongSo.Thang >=" + CStr(thg) + " AND ThongSo.Thang <= " + CStr(thgcuoi) + " AND Mid(LoaiTaiSan.SoHieu,3,1) = '2' GROUP BY HethongTK.SoHieu,TKCha0"
                  thg = 2
            Case 3:
                sql = "SELECT DISTINCTROW TKCha0,Sum(ThongSo.KH_NS+ThongSo.KH_TBS+ThongSo.KH_CNK+ ThongSo.KH_TD) AS TKH, HethongTK.SoHieu " _
                        & "FROM HethongTK RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON HethongTK.MaSo = ThongSo.MaDTSD " _
                        & "WHERE HethongTK.SoHieu LIKE '" + txt.Text + "*' AND ThongSo.Thang >= " + CStr(thg) + " AND ThongSo.Thang <= " + CStr(thgcuoi) + " AND Mid(LoaiTaiSan.SoHieu,3,1) = '7' GROUP BY HethongTK.SoHieu,TKCha0"
               thg = 7
           Case 4                  ' KhÊu hao tµi s¶n cè ®Þnh h÷u h×nh
                  sql = "SELECT DISTINCTROW TKCha0,Sum(ThongSo.KH_NS+ThongSo.KH_TBS+ThongSo.KH_CNK+ ThongSo.KH_TD) AS TKH, HethongTK.SoHieu " _
                        & "FROM HethongTK RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON HethongTK.MaSo = ThongSo.MaDTSD " _
                        & "WHERE HethongTK.SoHieu LIKE '" + txt.Text + "*' AND ThongSo.Thang >= " + CStr(thg) + " AND ThongSo.Thang <= " + CStr(thgcuoi) + " AND Mid(LoaiTaiSan.SoHieu,1,3) = '242' GROUP BY HethongTK.SoHieu,TKCha0"
                  thg = 1
      End Select
      Set rs_khauhao = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      On Error GoTo Err_NoCurrentRecord
            rs_khauhao.MoveLast
      On Error GoTo 0
      parSoPS = rs_khauhao.recordCount
      ReDim arPhatSinh(0 To parSoPS) As tpPhatSinh
      rs_khauhao.MoveFirst
      ' C¸c dßng ph¸t sinh ®èi øng cña tµi kho¶n chi phÝ khÊu hao
      Do Until rs_khauhao.EOF
            If Not rs_khauhao!TKH = 0 Then
                  arPhatSinh(chi_so).TK_SoHieu = rs_khauhao!sohieu
                  arPhatSinh(chi_so).PS_Loai = -1
                  arPhatSinh(chi_so).PS_SoLg = rs_khauhao!TKH
                  If pDTTP <> 0 Then
                   Dim so_ttttt
                  so_ttttt = Len(SelectSQL("SELECT SoHieu AS F1 FROM HethongTK WHERE MaSo=" + CStr(rs_khauhao!TkCha0)))
                  
                       ' shct = Right(rs_khauhao!sohieu, Len(rs_khauhao!sohieu) - Len(SelectSQL("SELECT SoHieu AS F1 FROM HethongTK WHERE MaSo=" + CStr(rs_khauhao!TkCha0))))
                         shct = Right(rs_khauhao!sohieu, IIf(Len(rs_khauhao!sohieu) - so_ttttt < 0, 0, so_ttttt))
                       ' If SoHieu2MaSo(shct, "TP154") > 0 Then arPhatSinh(chi_so).ShTP = shct
                  End If
                  tong_ps = tong_ps + rs_khauhao!TKH
                  chi_so = chi_so + 1
            End If
            rs_khauhao.MoveNext
      Loop
      ' Dßng ph¸t sinh tæng céng cña tµi s¶n
    '  arPhatSinh(chi_so).TK_SoHieu = "214" + CStr(thg)
      If frmKhauHao.Option(0).Value = True Then
         arPhatSinh(chi_so).TK_SoHieu = "214" + CStr(thg)
      Else
         arPhatSinh(chi_so).TK_SoHieu = "242"
      End If
      arPhatSinh(chi_so).PS_Loai = 1
      arPhatSinh(chi_so).PS_SoLg = tong_ps
      
      TrichKhauHao = parSoPS
Err_NoCurrentRecord:
      rs_khauhao.Close
      Set rs_khauhao = Nothing
      Me.MousePointer = 0
End Function
'======================================================================================
' command
'======================================================================================
Private Sub Command_Click(Index As Integer)
      Dim i As Integer, tdau As Integer, tcuoi As Integer
      
      Me.MousePointer = 11
      Select Case Index
            Case 0
                  If Combo(1).ListIndex < Combo(0).ListIndex Then Combo(1).ListIndex = Combo(0).ListIndex
                  tdau = Combo(0).ItemData(Combo(0).ListIndex)
                  tcuoi = Combo(1).ItemData(Combo(1).ListIndex)
                  If ThangDaKhauHao(tdau, tcuoi, FrmChungtu.CboNguon(0).ItemData(LoaiKhauHao), txt.Text) Then
                        MsgBox "Chøng tõ khÊu hao cò sÏ ®­îc thay b»ng chøng tõ míi!", vbInformation, App.ProductName
                  End If
                  ' ChuÈn bÞ c¸c biÕn trao ®æi d÷ liÖu
                  HienThongBao " CËp nhËt gi¸ trÞ tµi s¶n ...", 1
                  For i = CThangDB(tdau) To CThangDB(tcuoi)
                        ' CËp nhËt gi¸ trÞ tµi s¶n cho th¸ng cÇn trÝch khÊu hao
                        CapNhatGiaTriTaiSan i, GauGe
                        GauGe.Value = 0
                  Next
                  ' Thµnh lËp c¸c dßng ph¸t sinh
                  If Not TrichKhauHao = 0 Then
                        FrmChungtu.CboNguon(0).ListIndex = LoaiKhauHao ' IIf(LoaiKhauHao <> 0, 3 - LoaiKhauHao, 0)
                        FrmChungtu.CboNguon(0).tag = txt.Text
                        SetListIndex FrmChungtu.CboThang, CLng(tcuoi)
                        FrmChungtu.tag = tdau
                        pGhichungtu = 1
                        Unload Me
                  Else
                        Beep
                        MsgBox "Kh«ng cã tµi s¶n nµo !", vbExclamation
                  End If
                  pThangTacDong = 0
            Case 1
                  SendKeys "{Escape}", False
      End Select
      Me.MousePointer = 0
End Sub
'======================================================================================
' Option
'======================================================================================
Private Sub Option_Click(Index As Integer)
      LoaiKhauHao = Index
End Sub
