VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Begin VB.Form FrmBCQT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "B¸o c¸o qu¶n trÞ"
   ClientHeight    =   7395
   ClientLeft      =   1935
   ClientTop       =   870
   ClientWidth     =   9045
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
   Icon            =   "FrmBCQT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   0
      ItemData        =   "FrmBCQT.frx":57E2
      Left            =   5520
      List            =   "FrmBCQT.frx":57E4
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   1
      ItemData        =   "FrmBCQT.frx":57E6
      Left            =   7800
      List            =   "FrmBCQT.frx":57E8
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TiÒn mÆt, tiÒn göi, tiÒn ®ang chuyÓn"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "§Çu t­ ng¾n h¹n vµ dù phßng gi¶m gi¸ ®Çu t­ ng¾n h¹n"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   4455
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ph¶i thu vµ øng tr­íc cña kh¸ch hµng"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   4455
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ph¶i tr¶ vµ øng tr­íc cho ng­êi b¸n"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   4455
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ph¶i thu kh¸c vµ dù phßng thu khã ®ßi"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   3255
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chi tiÕt t¹m øng"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   3255
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chi phÝ tr¶ tr­íc"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   3255
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ThÕ chÊp, ký quü"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   3255
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KiÓm kª vËt t­, c«ng cô dông cô, thµnh phÈm, hµng ho¸"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   4335
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KiÓm kª gi¸ trÞ khèi l­îng x©y l¾p thi c«ng dë dang"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   3975
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chi tiÕt liªn doanh, ®Çu t­ dµi h¹n, dù phßng gi¶m gi¸ ®Çu t­ dµi h¹n"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   5775
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "T×nh h×nh thùc hiÖn ®Çu t­ x©y dùng c¬ b¶n"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   4335
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vay ng¾n h¹n, nî dµi h¹n ®Õn h¹n tr¶, vay dµi h¹n, nî dµi h¹n, ký quü dµi h¹n"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   6615
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "T×nh h×nh thùc hiÖn nghÜa vô víi Nhµ n­íc, c¸c kho¶n nép Ng©n s¸ch, nghÜa vô víi cÊp trªn"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   7215
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "L­¬ng vµ b¶o hiÔm x· héi"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   4455
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Thanh to¸n theo tiÕn ®é, ph¶i tr¶ ph¶i nép kh¸c"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   4455
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nguån vèn kinh doanh, c¸c quü"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   7
      Top             =   6240
      Width           =   4455
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chi phÝ vµ s¶n l­îng c«ng tr×nh, s¶n phÈm"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   4455
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chi phÝ qu¶n lý doanh nghiÖp"
      Height          =   255
      Index           =   18
      Left            =   4800
      TabIndex        =   5
      Top             =   480
      Width           =   3255
   End
   Begin VB.OptionButton OptBC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ph©n tÝch chØ tiªu tµi chÝnh"
      Height          =   255
      Index           =   19
      Left            =   4800
      TabIndex        =   4
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   0
      Left            =   5400
      Picture         =   "FrmBCQT.frx":57EA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   1
      Left            =   6600
      Picture         =   "FrmBCQT.frx":695C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   2
      Left            =   7800
      Picture         =   "FrmBCQT.frx":7DBE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar GauGe 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   7080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tõ th¸ng"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   27
      Tag             =   "From"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "®Õn th¸ng"
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   26
      Tag             =   "to"
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmBCQT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim baocao As Integer
Dim xlapp As Excel.Application, xlsheet As Worksheet

Private Sub Command_Click(Index As Integer)
    Dim i As Integer, tdau As Integer, tcuoi As Integer
    
    Select Case Index
        Case 0, 1:
            Me.MousePointer = 11
            GauGe.Value = 0
            GauGe.Max = 3
            If Combo(1).ListIndex < Combo(0).ListIndex Then Combo(1).ListIndex = Combo(0).ListIndex
            tdau = Combo(0).ItemData(Combo(0).ListIndex)
            tcuoi = Combo(1).ItemData(Combo(1).ListIndex)
            Select Case baocao
                Case 0:
                    Baocao111 tcuoi
                Case 1:
                    Baocao128 tcuoi
                Case 2:
                    Baocao131 tcuoi
                Case 12:
                    Baocao331 tcuoi
                Case 3:
                    Baocao138 tcuoi
                Case 4:
                    Baocao141 tcuoi
                Case 5:
                    Baocao142 tdau, tcuoi
                Case 6:
                    Baocao144 tcuoi
                Case 7:
                    Baocao152 tcuoi
                Case 8:
                    Baocao154 tcuoi
                Case 9:
                    Baocao222 tcuoi
                Case 10:
                    Baocao241 tdau, tcuoi
                Case 11:
                    Baocao311 tcuoi
                Case 13:
                    Baocao333 tdau, tcuoi
                Case 14:
                    Baocao334 tdau, tcuoi
                Case 15:
                    Baocao338 tdau, tcuoi
                Case 16:
                    Baocao411 tdau, tcuoi
                Case 17:
                    Baocao632 tdau, tcuoi
                Case 18:
                    Baocao642 tdau, tcuoi
                Case 19:
                    Baocao911 tdau, tcuoi
            End Select
        Case 2:
            Unload Me
    End Select
    GauGe.Value = GauGe.Max
    HienThongBao Me.Caption, 1
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    AddMonthToCbo Combo(0)
    AddMonthToCbo Combo(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HienThongBao "", 1
    Set xlapp = Nothing
    Set xlsheet = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyX:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyI:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyV:
                RFocus Command(2)
                Command_Click 2
        End Select
    End If
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyP Then frmMain.mnHT_Click 8
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub OptBc_Click(Index As Integer)
    baocao = Index
End Sub

Private Sub Baocao111(tcuoi As Integer)
    Dim duno As Double, duco As Double, rs As Recordset, i As Integer, k As Integer, ncuoi As Date
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "111.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\111.XLS", pCurDir + "111.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "111.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SoDuTK SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE Sohieu='1111'"), tcuoi, duno, duco
    xlsheet.Cells(7, 4) = duno - duco
    
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    xlsheet.Cells(9, 7) = duno - duco
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo, SoHieu, Ten, GhiChu, DuNo_" + CStr(CThangDB(tcuoi)) + "-DuCo_" + CStr(CThangDB(tcuoi)) + " AS SoDu FROM HethongTK WHERE SoHieu LIKE '1121*' AND TKCon=0 AND Cap>2 ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If i > 1 Then
            xlapp.Worksheets.Add count:=1, before:=xlapp.Sheets(2 + i)
            Set xlsheet = xlapp.Worksheets(2 + i)
            xlsheet.Range("A1", "H20").PasteSpecial
        Else
            Set xlsheet = xlapp.Worksheets(3)
        End If
        xlsheet.Cells(5, 3) = rs!Ten
        xlsheet.Cells(6, 3) = rs!GhiChu
        xlsheet.Cells(8, 5) = rs!sodu
        xlsheet.Range("A1", "H20").Copy
        rs.MoveNext
    Loop
    k = 3 + i
    i = 0
    Set rs = DBKetoan.OpenRecordset("SELECT HethongTK.MaSo, SoHieu, Ten, GhiChu, KyHieu, DuNo_" + CStr(CThangDB(tcuoi)) + "-DuCo_" + CStr(CThangDB(tcuoi)) + " AS SoDu,DuNT_" + CStr(CThangDB(tcuoi)) + " AS DuNT FROM HethongTK INNER JOIN NguyenTe ON HethongTK.MaNT=NguyenTe.MaSo " _
        & " WHERE SoHieu LIKE '1122*' AND TKCon=0 AND Cap>2 ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If i > 1 Then
            xlapp.Worksheets.Add count:=1, before:=xlapp.Sheets(k - 1 + i)
            Set xlsheet = xlapp.Worksheets(k - 1 + i)
            xlsheet.Range("A1", "H23").PasteSpecial
            xlsheet.Cells(4, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
        Else
            Set xlsheet = xlapp.Worksheets(k)
        End If
        xlsheet.Cells(6, 3) = rs!Ten
        xlsheet.Cells(7, 3) = rs!GhiChu
        xlsheet.Cells(12, 3) = rs!KyHieu
        xlsheet.Cells(12, 7) = rs!sodu
        xlsheet.Cells(12, 4) = rs!dunt
        xlsheet.Range("A1", "H23").Copy
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "111.XLS"
KetThuc:
End Sub

Private Sub Baocao128(tcuoi As Integer)
    Dim ncuoi As Date
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "128.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\128.XLS", pCurDir + "128.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "128.XLS"
    On Error GoTo 0
    
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "128", 2, tcuoi, 1, 2, 4, 7, 7, 7, ncuoi
    
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "129", 2, tcuoi, 1, 2, 3, 4, 7, 4, ncuoi
        
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "128.XLS"
KetThuc:
End Sub

Private Sub Baocao131(tcuoi As Integer)
    Dim ncuoi As Date
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "131.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\131.XLS", pCurDir + "131.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "131.XLS"
    On Error GoTo 0
    
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "1311", 3, tcuoi, 1, 2, 4, 0, 7, 8, ncuoi
        
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "1312", 3, tcuoi, 1, 2, 4, 0, 7, 5, ncuoi
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "131.XLS"
KetThuc:
End Sub

Private Sub Baocao331(tcuoi As Integer)
    Dim ncuoi As Date
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "331.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\331.XLS", pCurDir + "331.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "331.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "3311", 3, tcuoi, 1, 2, 4, 0, 7, 5, ncuoi
    
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "3312", 3, tcuoi, 1, 2, 4, 0, 7, 7, ncuoi
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "331.XLS"
KetThuc:
End Sub

Private Sub Baocao138(tcuoi As Integer)
    Dim ncuoi As Date
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "138.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\138.XLS", pCurDir + "138.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "138.XLS"
    On Error GoTo 0
    
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "1381", 3, tcuoi, 1, 2, 4, 3, 7, 5, ncuoi
    
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "1388", 3, tcuoi, 1, 2, 4, 5, 7, 8, ncuoi
        
    Set xlsheet = xlapp.Worksheets(3)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "139", 2, tcuoi, 1, 2, 3, 4, 7, 4, ncuoi
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "138.XLS"
KetThuc:
End Sub

Private Sub Baocao141(tcuoi As Integer)
    Dim ncuoi As Date
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "141.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\141.XLS", pCurDir + "141.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "141.XLS"
    On Error GoTo 0
    
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "141", 2, tcuoi, 1, 2, 4, 3, 7, 8, ncuoi
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "141.XLS"
KetThuc:
End Sub

Private Sub Baocao142(tdau As Integer, tcuoi As Integer)
    Dim rs As Recordset, i As Integer, ncuoi As Date, TK As New ClsTaikhoan
    Dim st As String, sqln As String, sqlc As String, sqlnx As String, sqlcx As String
    Dim duno As Double, duco As Double, dunt As Double
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "142.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\142.XLS", pCurDir + "142.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "142.XLS"
    On Error GoTo 0
    
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + No_" + st
        sqlc = sqlc + " + Co_" + st
    Next
    For i = 1 To tcuoi
        st = CStr(i)
        sqlnx = sqlnx + " + No_" + st
        sqlcx = sqlcx + " + Co_" + st
    Next
    i = 0
    Set rs = DBKetoan.OpenRecordset("SELECT SoHieu, Cap, Ten, GhiChu, Sum(DuNo_" + CStr(CThangDB(ThangTruoc(tcuoi))) + "-DuCo_" + CStr(CThangDB(ThangTruoc(tcuoi))) + ") AS DuDK, Sum(" + sqln + ") AS PSN, Sum(" + sqlc + ") AS PSC, Sum(" + sqlnx + ") AS PSNx, Sum(" + sqlcx + ") AS PSCx, Sum(DuNo_" + CStr(CThangDB(tcuoi)) + "-DuCo_" + CStr(CThangDB(tcuoi)) + ") AS DuCK FROM HethongTK WHERE SoHieu LIKE '142*' AND Cap>1 GROUP BY SoHieu, Cap, Ten, GhiChu ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        xlsheet.Cells(7 + i, 1) = CStr(i)
        xlsheet.Cells(7 + i, 2) = Space(2 * (rs!cap - 1)) + rs!Ten
        xlsheet.Cells(7 + i, 3) = rs!DuDk
        xlsheet.Cells(7 + i, 4) = rs!psn
        xlsheet.Cells(7 + i, 5) = rs!psc
        xlsheet.Cells(7 + i, 6) = rs!psnx
        xlsheet.Cells(7 + i, 7) = rs!pscx
        xlsheet.Cells(7 + i, 8) = rs!DuCK
        rs.MoveNext
    Loop
    If i > 0 Then
        TK.InitTaikhoanSohieu 142
        TK.SoDuTK ThangTruoc(tdau), duno, duco, dunt
        xlsheet.Cells(8 + i, 2) = "Tæng céng"
        xlsheet.Cells(8 + i, 3) = duno - duco
        TK.SoPhatSinh tdau, tcuoi, duno, duco
        xlsheet.Cells(8 + i, 4) = duno
        xlsheet.Cells(8 + i, 5) = duco
        TK.SoPhatSinh pThangDauKy, tcuoi, duno, duco
        xlsheet.Cells(8 + i, 6) = duno
        xlsheet.Cells(8 + i, 7) = duco
        TK.SoDuTK tcuoi, duno, duco, dunt
        xlsheet.Cells(8 + i, 8) = duno - duco
        xlsheet.Range("A" + CStr(8 + i), "H" + CStr(8 + i)).Font.Bold = True
    End If
    rs.Close
    Set rs = Nothing
    With xlsheet.Range("A6", "H" + CStr(8 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    XLSFooter xlsheet, 10 + i, 8, NgayCuoiThang(pNamTC, tcuoi)
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "142.XLS"
KetThuc:
    Set TK = Nothing
End Sub

Private Sub Baocao144(tcuoi As Integer)
    Dim ncuoi As Date
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "144.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\144.XLS", pCurDir + "144.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "144.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "144", 2, tcuoi, 1, 2, 3, 4, 8, 4, ncuoi
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "144.XLS"
KetThuc:
End Sub

Private Sub Baocao152(tcuoi As Integer)
    Dim rs As Recordset, i As Integer, ncuoi As Date, shtk As String, shpl As String, kh As String, rs2 As Recordset, j As Integer, k As Integer
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "152.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\152.XLS", pCurDir + "152.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "152.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(3, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW HethongTK.SoHieu AS SHTK, PhanLoaiVattu.SoHieu AS SHPL, HethongTK.Ten, PhanLoaiVattu.TenPhanLoai, KhoHang.TenKho, Sum(Tien_" + CStr(CThangDB(tcuoi)) + ") AS SoDu FROM (((TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo) INNER JOIN KhoHang ON TonKho.MaSoKho=KhoHang.MaSo) INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo) INNER JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo WHERE (HethongTK.SoHieu LIKE '152*' OR HethongTK.SoHieu LIKE '153*' OR HethongTK.SoHieu LIKE '155*' OR HethongTK.SoHieu LIKE '156*') GROUP BY HethongTK.SoHieu, PhanLoaiVattu.SoHieu, HethongTK.Ten, PhanLoaiVattu.TenPhanLoai, KhoHang.TenKho HAVING Sum(Tien_" + CStr(CThangDB(tcuoi)) + ")<>0", dbOpenSnapshot)
    Do While Not rs.EOF
a:
        i = i + 1
        If shtk <> rs!shtk Then
            shtk = rs!shtk
            xlsheet.Cells(7 + i, 2) = rs!Ten
            xlsheet.Range("A" + CStr(7 + i), "N" + CStr(7 + i)).Font.Bold = True
            GoTo a
        End If
        If shpl <> rs!shpl Then
            shpl = rs!shpl
            xlsheet.Cells(7 + i, 2) = rs!TenPhanLoai
            xlsheet.Range("A" + CStr(7 + i), "N" + CStr(7 + i)).Font.Bold = True
            xlsheet.Range("A" + CStr(7 + i), "N" + CStr(7 + i)).Font.Italic = True
            GoTo a
        End If
        xlsheet.Cells(7 + i, 1) = CStr(i)
        xlsheet.Cells(7 + i, 2) = rs!tenkho
        xlsheet.Cells(7 + i, 3) = rs!sodu
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    With xlsheet.Range("A5", "N" + CStr(7 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
        .Font.Name = pFontName
    End With
    XLSFooter xlsheet, 10 + i, 14, NgayCuoiThang(pNamTC, tcuoi)
    
    i = 0
    shtk = ""
    shpl = ""
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(3, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW HethongTK.SoHieu AS SHTK, PhanLoaiVattu.SoHieu AS SHPL, HethongTK.Ten, TenPhanLoai, KhoHang.TenKho, Vattu.SoHieu, TenVattu, DonVi, (Tien_" + CStr(CThangDB(tcuoi)) + ") AS Tien, (Luong_" + CStr(CThangDB(tcuoi)) + ") AS Luong FROM (((TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo) INNER JOIN KhoHang ON TonKho.MaSoKho=KhoHang.MaSo) INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo) INNER JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo WHERE (HethongTK.SoHieu LIKE '152*' OR HethongTK.SoHieu LIKE '153*' OR HethongTK.SoHieu LIKE '155*' OR HethongTK.SoHieu LIKE '156*') AND (Tien_" + CStr(CThangDB(tcuoi)) + "<>0 OR Luong_" + CStr(CThangDB(tcuoi)) + "<>0) ORDER BY HethongTK.SoHieu, PhanLoaiVattu.SoHieu, KhoHang.TenKho, Vattu.SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
B:
        i = i + 1
        If shtk <> rs!shtk Then
            shtk = rs!shtk
            xlsheet.Cells(7 + i, 2) = rs!Ten
            xlsheet.Range("A" + CStr(7 + i), "N" + CStr(7 + i)).Font.Bold = True
            GoTo B
        End If
        If shpl <> rs!shpl Then
            shpl = rs!shpl
            xlsheet.Cells(7 + i, 2) = rs!TenPhanLoai
            xlsheet.Range("A" + CStr(7 + i), "N" + CStr(7 + i)).Font.Bold = True
            xlsheet.Range("A" + CStr(7 + i), "N" + CStr(7 + i)).Font.Italic = True
            GoTo B
        End If
        If kh <> rs!tenkho Then
            kh = rs!tenkho
            xlsheet.Cells(7 + i, 2) = rs!tenkho
            xlsheet.Range("A" + CStr(7 + i), "N" + CStr(7 + i)).Font.Italic = True
            GoTo B
        End If
        xlsheet.Cells(7 + i, 1) = CStr(i)
        xlsheet.Cells(7 + i, 2) = rs!TenVattu
        xlsheet.Cells(7 + i, 3) = rs!sohieu
        xlsheet.Cells(7 + i, 4) = rs!DonVi
        xlsheet.Cells(7 + i, 5) = rs!luong
        xlsheet.Cells(7 + i, 6) = rs!tien
        rs.MoveNext
    Loop
    With xlsheet.Range("A5", "X" + CStr(7 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
        .Font.Name = pFontName
    End With
    XLSFooter xlsheet, 10 + i, 24, NgayCuoiThang(pNamTC, tcuoi)
    i = 0
    Set xlsheet = xlapp.Worksheets(3)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(3, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW KhoHang.MaSo, TenKho FROM (TonKho INNER JOIN KhoHang ON TonKho.MaSoKho=KhoHang.MaSo) INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo WHERE (HethongTK.SoHieu LIKE '152*' OR HethongTK.SoHieu LIKE '153*' OR HethongTK.SoHieu LIKE '155*' OR HethongTK.SoHieu LIKE '156*') AND (Tien_" + CStr(CThangDB(tcuoi)) + "<>0 OR Luong_" + CStr(CThangDB(tcuoi)) + "<>0)ORDER BY TenKho", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        j = 0
        shtk = ""
        If i > 1 Then
            xlapp.Worksheets.Add count:=1, before:=xlapp.Sheets(2 + i)
            Set xlsheet = xlapp.Worksheets(2 + i)
            xlsheet.Range("A1", "T7").PasteSpecial
        Else
            Set xlsheet = xlapp.Worksheets(3)
        End If
        xlsheet.Cells(4, 1) = "Kho :" + rs!tenkho
        Set rs2 = DBKetoan.OpenRecordset("SELECT DISTINCTROW HethongTK.SoHieu AS SHTK, HethongTK.Ten, Vattu.SoHieu, TenVattu, DonVi, (Tien_" + CStr(CThangDB(tcuoi)) + ") AS Tien, (Luong_" + CStr(CThangDB(tcuoi)) + ") AS Luong FROM (TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo) INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo WHERE (HethongTK.SoHieu LIKE '152*' OR HethongTK.SoHieu LIKE '153*' OR HethongTK.SoHieu LIKE '155*' OR HethongTK.SoHieu LIKE '156*') AND (Tien_" + CStr(CThangDB(tcuoi)) + "<>0 OR Luong_" + CStr(CThangDB(tcuoi)) + "<>0) AND MaSoKho=" + CStr(rs!MaSo) + " ORDER BY HethongTK.SoHieu, Vattu.SoHieu", dbOpenSnapshot)
        Do While Not rs2.EOF
c:
            j = j + 1
            If shtk <> rs2!shtk Then
                shtk = rs2!shtk
                xlsheet.Cells(7 + j, 2) = rs2!Ten
                xlsheet.Range("A" + CStr(7 + j), "T" + CStr(7 + j)).Font.Bold = True
                GoTo c
            End If
            xlsheet.Cells(7 + j, 1) = CStr(j)
            xlsheet.Cells(7 + j, 2) = rs2!TenVattu
            xlsheet.Cells(7 + j, 3) = rs2!sohieu
            xlsheet.Cells(7 + j, 4) = rs2!DonVi
            xlsheet.Cells(7 + j, 5) = rs2!luong
            xlsheet.Cells(7 + j, 6) = rs2!tien
            rs2.MoveNext
        Loop
        rs2.Close
        With xlsheet.Range("A5", "T" + CStr(7 + j))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(0, 0, 0)
            .RowHeight = 18
            .VerticalAlignment = xlVAlignCenter
        End With
        XLSFooter xlsheet, 10 + j, 20, NgayCuoiThang(pNamTC, tcuoi)
        xlsheet.Range("A1", "T7").Copy
        rs.MoveNext
    Loop
    rs.Close
    
    k = 3 + i
    i = 0
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW KhoHang.MaSo AS MKho,TenKho,HethongTK.MaSo,HethongTK.SoHieu,Ten FROM (TonKho INNER JOIN KhoHang ON TonKho.MaSoKho=KhoHang.MaSo) INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo WHERE (HethongTK.SoHieu LIKE '152*' OR HethongTK.SoHieu LIKE '153*' OR HethongTK.SoHieu LIKE '155*' OR HethongTK.SoHieu LIKE '156*') AND (Tien_" + CStr(CThangDB(tcuoi)) + "<>0 OR Luong_" + CStr(CThangDB(tcuoi)) + "<>0) GROUP BY TenKho,HethongTK.SoHieu,Ten,KhoHang.MaSo,HethongTK.MaSo ORDER BY TenKho,HethongTK.SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If i > 1 Then
            xlapp.Worksheets.Add count:=1, before:=xlapp.Sheets(k - 1 + i)
            Set xlsheet = xlapp.Worksheets(k - 1 + i)
            xlsheet.Range("A1", "Y8").PasteSpecial
        Else
            Set xlsheet = xlapp.Worksheets(k)
        End If
        xlsheet.Cells(4, 1) = "Kho: " + rs!tenkho
        xlsheet.Cells(5, 1) = "Sè hiÖu TK: " + rs!sohieu + " - Tªn tµi kho¶n: " + rs!Ten
        j = 0
        Set rs2 = DBKetoan.OpenRecordset("SELECT DISTINCTROW Vattu.SoHieu, TenVattu, DonVi, (Tien_" + CStr(CThangDB(tcuoi)) + ") AS Tien, (Luong_" + CStr(CThangDB(tcuoi)) + ") AS Luong FROM TonKho INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo WHERE (Tien_" + CStr(CThangDB(tcuoi)) + "<>0 OR Luong_" + CStr(CThangDB(tcuoi)) + "<>0) AND MaSoKho=" + CStr(rs!mkho) + " AND MaTaiKhoan=" + CStr(rs!MaSo) + " ORDER BY Vattu.SoHieu", dbOpenSnapshot)
        Do While Not rs2.EOF
            j = j + 1
            xlsheet.Cells(8 + j, 1) = CStr(j)
            xlsheet.Cells(8 + j, 2) = rs2!TenVattu
            xlsheet.Cells(8 + j, 3) = rs2!sohieu
            xlsheet.Cells(8 + j, 4) = rs2!DonVi
            xlsheet.Cells(8 + j, 6) = rs2!luong
            xlsheet.Cells(8 + j, 7) = rs2!tien
            rs2.MoveNext
        Loop
        rs2.Close
        With xlsheet.Range("A6", "Y" + CStr(8 + j))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(0, 0, 0)
            .RowHeight = 18
            .VerticalAlignment = xlVAlignCenter
            .Font.Name = pFontName
        End With
        rs.MoveNext
        XLSFooter xlsheet, 10 + j, 25, NgayCuoiThang(pNamTC, tcuoi)
        xlsheet.Range("A1", "Y8").Copy
    Loop
    rs.Close
    Set rs = Nothing
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "152.XLS"
KetThuc:
End Sub

Private Sub Baocao154(tcuoi As Integer)
    Dim rs As Recordset, i As Integer, ncuoi As Date, shpl As String, j As Integer
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "154.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\154.XLS", pCurDir + "154.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "154.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(3, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    
    InChiPhi3 pThangDauKy, tcuoi, 0
    Set rs = DBKetoan.OpenRecordset("QChitiet", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If shpl = rs!shpl Then
            j = j + 1
            xlsheet.Cells(7 + i, 1) = CStr(j)
            xlsheet.Cells(7 + i, 2) = rs!TenVattu
            'xlsheet.Cells(7 + i, 6) = rs!nvl
            'xlsheet.Cells(7 + i, 7) = rs!nc
            'xlsheet.Cells(7 + i, 8) = rs!m
            'xlsheet.Cells(7 + i, 9) = rs!cpsxc
            xlsheet.Cells(7 + i, 12) = rs!CK1
            rs.MoveNext
        Else
            shpl = rs!shpl
            xlsheet.Cells(7 + i, 2) = rs!TenPhanLoai
            xlsheet.Range("B" + CStr(7 + i)).Font.Bold = True
        End If
    Loop
    With xlsheet.Range("A6", "L" + CStr(7 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    XLSFooter xlsheet, 10 + i, 8, NgayCuoiThang(pNamTC, tcuoi)
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "154.XLS"
KetThuc:
End Sub

Private Sub Baocao222(tcuoi As Integer)
    Dim ncuoi As Date
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "222.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\222.XLS", pCurDir + "222.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "222.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "222", 2, tcuoi, 1, 2, 5, 6, 7, 6, ncuoi
    
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "228", 2, tcuoi, 1, 2, 4, 7, 7, 7, ncuoi
        
    Set xlsheet = xlapp.Worksheets(3)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "229", 2, tcuoi, 1, 2, 3, 4, 7, 4, ncuoi
    
    Set xlsheet = xlapp.Worksheets(4)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "244", 2, tcuoi, 1, 2, 3, 4, 8, 4, ncuoi
        
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "222.XLS"
KetThuc:
End Sub

Private Sub Baocao241(tdau As Integer, tcuoi As Integer)
    Dim rs As Recordset, i As Integer, ncuoi As Date, j As Integer
    Dim sqln As String, sqlc As String, sqlnx As String, sqlcx As String, TK As New ClsTaikhoan
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "241.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\241.XLS", pCurDir + "241.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "241.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqln = sqln + " + No_" + CStr(i)
        sqlc = sqlc + " + Co_" + CStr(i)
    Next
    For i = 1 To tcuoi
        sqlnx = sqlnx + " + No_" + CStr(i)
        sqlcx = sqlcx + " + Co_" + CStr(i)
    Next
    
    i = 0
    Set rs = DBKetoan.OpenRecordset("SELECT SoHieu, Cap, Ten, GhiChu, Sum(DuNo_0-DuCo_0) AS DuDK,Sum(" + sqln + ") AS PSN,Sum(" + sqlc + ") AS PSC,Sum(" + sqlnx + ") AS PSNx,Sum(" + sqlcx + ") AS PSCx,Sum(DuNo_" + CStr(CThangDB(tcuoi)) + "-DuCo_" + CStr(CThangDB(tcuoi)) + ") AS DuCK " _
        & " FROM HethongTK WHERE SoHieu LIKE '241*' AND Cap>1 AND (DuNo_0-DuCo_0<>0 OR " + sqlnx + "<>0 OR " + sqlcx + "<>0) GROUP BY SoHieu, Cap, Ten, GhiChu ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        TK.InitTaikhoanSohieu rs!sohieu
        
        xlsheet.Cells(10 + i, 2) = Space(2 * (rs!cap - 2)) + rs!Ten
        xlsheet.Cells(10 + i, 3) = rs!DuDk
        xlsheet.Cells(10 + i, 4) = rs!psn
        xlsheet.Cells(10 + i, 5) = rs!psnx
        xlsheet.Cells(10 + i, 6) = TK.PSTuKhoiCong(ncuoi, -1)
        xlsheet.Cells(10 + i, 7) = rs!psc
        xlsheet.Cells(10 + i, 8) = rs!pscx
        xlsheet.Cells(10 + i, 9) = TK.PSTuKhoiCong(ncuoi, 1)
        xlsheet.Cells(10 + i, 11) = rs!DuCK
        Select Case rs!cap
            Case 2, 3:
                xlsheet.Range("A" + CStr(10 + i), "K" + CStr(10 + i)).Font.Bold = True
                j = 0
            Case 4:
                xlsheet.Range("A" + CStr(10 + i), "K" + CStr(10 + i)).Font.Bold = True
                xlsheet.Range("A" + CStr(10 + i), "K" + CStr(10 + i)).Font.Italic = True
                j = j + 1
                xlsheet.Cells(10 + i, 1) = CStr(j)
        End Select
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If i > 0 Then
        With xlsheet.Range("A10", "K" + CStr(10 + i))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(0, 0, 0)
            .RowHeight = 18
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
    XLSFooter xlsheet, 12 + i, 11, NgayCuoiThang(pNamTC, tcuoi)
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "241.XLS"
KetThuc:
End Sub

Private Sub Baocao311(tcuoi As Integer)
    Dim rs As Recordset, i As Integer, ncuoi As Date, j As Integer
    Dim sqln As String, sqlc As String
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "311.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\311.XLS", pCurDir + "311.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "311.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "311", 2, tcuoi, 1, 2, 3, 6, 7, 6, ncuoi
        
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "315", 2, tcuoi, 1, 2, 3, 5, 7, 5, ncuoi
    
    For i = 1 To CThangDB(tcuoi)
        sqln = sqln + " + No_" + CStr(i)
        sqlc = sqlc + " + Co_" + CStr(i)
    Next
    i = 0
    j = 0
    Set xlsheet = xlapp.Worksheets(3)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    Set rs = DBKetoan.OpenRecordset("SELECT SoHieu, Cap, Ten, GhiChu, Sum(DuCo_0-DuNo_0) AS DN, Sum(" + sqln + ") AS PSN, Sum(" + sqlc + ") AS PSC, Sum(DuCo_" + CStr(CThangDB(tcuoi)) + "-DuNo_" + CStr(CThangDB(tcuoi)) + ") AS SoDu FROM HethongTK WHERE SoHieu LIKE '341*' AND Cap>1 GROUP BY SoHieu, Cap, Ten, GhiChu ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If rs!cap = 2 Then
            xlsheet.Range("A" + CStr(9 + i), "O" + CStr(9 + i)).Font.Bold = True
            j = 0
        Else
            j = j + 1
            xlsheet.Cells(9 + i, 1) = CStr(j)
        End If
        xlsheet.Cells(9 + i, 2) = Space(2 * (rs!cap - 1)) + rs!Ten
        xlsheet.Cells(9 + i, 6) = rs!dn
        xlsheet.Cells(9 + i, 7) = rs!psc
        xlsheet.Cells(9 + i, 10) = rs!psn
        xlsheet.Cells(9 + i, 15) = rs!sodu
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    With xlsheet.Range("A9", "O" + CStr(10 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    XLSFooter xlsheet, 12 + i, 15, NgayCuoiThang(pNamTC, tcuoi)
    
    Set xlsheet = xlapp.Worksheets(4)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "342", 2, tcuoi, 1, 2, 5, 7, 7, 7, ncuoi
        
    Set xlsheet = xlapp.Worksheets(5)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "344", 2, tcuoi, 1, 2, 4, 5, 7, 5, ncuoi
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "311.XLS"
KetThuc:
End Sub

Private Sub Baocao333(tdau As Integer, tcuoi As Integer)
    Dim rs As Recordset, i As Integer, ncuoi As Date, j As Integer, k As Integer
    Dim sqln As String, sqlc As String, sqlnx As String, sqlcx As String, kq As Double
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "333.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\333.XLS", pCurDir + "333.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "333.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi)
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqln = sqln + " + No_" + CStr(i)
        sqlc = sqlc + " + Co_" + CStr(i)
    Next
    For i = 1 To tcuoi
        sqlnx = sqlnx + " + No_" + CStr(i)
        sqlcx = sqlcx + " + Co_" + CStr(i)
    Next
    
    i = 0
    Set rs = DBKetoan.OpenRecordset("SELECT SoHieu, Cap, Ten, GhiChu, Sum(DuCo_0-DuNo_0) AS DuDK,Sum(" + sqln + ") AS PSN,Sum(" + sqlc + ") AS PSC,Sum(" + sqlnx + ") AS PSNx,Sum(" + sqlcx + ") AS PSCx,Sum(DuCo_" + CStr(CThangDB(tcuoi)) + "-DuNo_" + CStr(CThangDB(tcuoi)) + ") AS DuCK " _
        & " FROM HethongTK WHERE ((SoHieu LIKE '333*' AND Cap=2) OR (SoHieu LIKE '333111*' AND Cap=5 AND TKCon=0)) AND (DuCo_0-DuNo_0<>0 OR " + sqlnx + "<>0 OR " + sqlcx + "<>0) GROUP BY SoHieu, Cap, Ten, GhiChu ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If rs!cap = 2 Then
            j = j + 1
            xlsheet.Cells(9 + i, 1) = CStr(j)
            xlsheet.Range("A" + CStr(9 + i), "H" + CStr(9 + i)).Font.Bold = True
            For k = 3 To 8
                If xlsheet.Range(XLSCol(k) + "9").Formula = "" Then
                    xlsheet.Range(XLSCol(k) + "9").Formula = "=" + XLSCol(k) + CStr(9 + i)
                Else
                    xlsheet.Range(XLSCol(k) + "9").Formula = xlsheet.Range(XLSCol(k) + "9").Formula + "+" + XLSCol(k) + CStr(9 + i)
                End If
            Next
        End If
        xlsheet.Cells(9 + i, 2) = Space(2 * (rs!cap - 1)) + rs!Ten
        xlsheet.Cells(9 + i, 3) = rs!DuDk
        xlsheet.Cells(9 + i, 8) = rs!DuCK
        If Left(rs!sohieu, 4) = "3331" Then
            xlsheet.Cells(9 + i, 4) = PhaiNopVAT(tdau, tcuoi, rs!sohieu, kq)
            xlsheet.Cells(9 + i, 6) = kq
            xlsheet.Cells(9 + i, 5) = PhaiNopVAT(pThangDauKy, tcuoi, rs!sohieu, kq)
            xlsheet.Cells(9 + i, 7) = kq
        Else
            xlsheet.Cells(9 + i, 4) = rs!psc
            xlsheet.Cells(9 + i, 5) = rs!pscx
            xlsheet.Cells(9 + i, 6) = rs!psn
            xlsheet.Cells(9 + i, 7) = rs!psnx
        End If
        rs.MoveNext
    Loop
    
    i = 0
    j = 0
    Set rs = DBKetoan.OpenRecordset("SELECT SoHieu, Cap, Ten, GhiChu, Sum(DuCo_0-DuNo_0) AS DuDK,Sum(" + sqln + ") AS PSN,Sum(" + sqlc + ") AS PSC,Sum(" + sqlnx + ") AS PSNx,Sum(" + sqlcx + ") AS PSCx,Sum(DuCo_" + CStr(CThangDB(tcuoi)) + "-DuNo_" + CStr(CThangDB(tcuoi)) + ") AS DuCK " _
        & " FROM HethongTK WHERE ((SoHieu LIKE '3383*') OR (SoHieu LIKE '3384*')) AND (DuCo_0-DuNo_0<>0 OR " + sqlnx + "<>0 OR " + sqlcx + "<>0) GROUP BY SoHieu, Cap, Ten, GhiChu ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If rs!cap = 2 Then
            j = j + 1
            xlsheet.Cells(30 + i, 1) = CStr(j)
            xlsheet.Range("A" + CStr(30 + i), "H" + CStr(30 + i)).Font.Bold = True
            For k = 3 To 8
                If xlsheet.Range(XLSCol(k) + "30").Formula = "" Then
                    xlsheet.Range(XLSCol(k) + "30").Formula = "=" + XLSCol(k) + CStr(30 + i)
                Else
                    xlsheet.Range(XLSCol(k) + "30").Formula = xlsheet.Range(XLSCol(k) + "30").Formula + "+" + XLSCol(k) + CStr(30 + i)
                End If
            Next
        End If
        xlsheet.Cells(30 + i, 2) = Space(2 * (rs!cap - 1)) + rs!Ten
        xlsheet.Cells(30 + i, 3) = rs!DuDk
        xlsheet.Cells(30 + i, 4) = rs!psc
        xlsheet.Cells(30 + i, 5) = rs!pscx
        xlsheet.Cells(30 + i, 6) = rs!psn
        xlsheet.Cells(30 + i, 7) = rs!psnx
        xlsheet.Cells(30 + i, 8) = rs!DuCK
        rs.MoveNext
    Loop
    rs.Close
    
    xlsheet.Cells(50, 3) = SoPSTK("3331", tdau, tcuoi, 1)
    kq = SoPSTK(pVATV, tdau, tcuoi, -1)
    xlsheet.Cells(51, 3) = kq
    If PSDu("3331", pVATV, tdau, tcuoi) > 0 Then
        xlsheet.Cells(52, 3) = PSDu("3331", pVATV, tdau, tcuoi)
    Else
        kq = kq - PSDu("6", pVATV, tdau, tcuoi)
        xlsheet.Cells(52, 3) = kq
    End If
    xlsheet.Cells(50, 4) = SoPSTK("3331", pThangDauKy, tcuoi, 1)
    kq = SoPSTK(pVATV, pThangDauKy, tcuoi, -1)
    xlsheet.Cells(51, 4) = kq
    If PSDu("3331", pVATV, pThangDauKy, tcuoi) > 0 Then
        xlsheet.Cells(52, 4) = PSDu("3331", pVATV, pThangDauKy, tcuoi)
    Else
        kq = kq - PSDu("6", pVATV, pThangDauKy, tcuoi)
        xlsheet.Cells(52, 4) = kq
    End If
    
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi)
    i = 0
    j = 0
    sqln = ""
    Set rs = DBKetoan.OpenRecordset("SELECT HethongTK.SoHieu AS SH, HethongTK.Ten,ChungTu.SoHieu,ChungTu.NgayCT, SoPS FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKTCNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo WHERE " _
        + WThang("ThangCT", tdau, tcuoi) + " AND HethongTK.SoHieu LIKE '333*' AND (TK.SoHieu LIKE '111*' OR TK.SoHieu LIKE '112*') ORDER BY HethongTK.SoHieu, ChungTu.NgayCT, ChungTu.SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If sqln <> rs!sh Then
            j = j + 1
            sqln = rs!sh
            xlsheet.Cells(7 + i, 1) = CStr(j)
        End If
        xlsheet.Cells(7 + i, 2) = rs!Ten
        xlsheet.Cells(7 + i, 3) = rs!sohieu
        xlsheet.Cells(7 + i, 4) = rs!NgayCT
        xlsheet.Cells(7 + i, 5) = rs!sops
        rs.MoveNext
    Loop
    xlsheet.Range("A" + CStr(9 + i), "AE" + CStr(9 + i)).Font.Bold = True
    xlsheet.Cells(9 + i, 2) = "Tæng céng"
    xlsheet.Range("E" + CStr(9 + i)).Formula = "=SUM(E7:E" + CStr(8 + i) + ")"
    With xlsheet.Range("A5", "F" + CStr(10 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    XLSFooter xlsheet, 12 + i, 6, NgayCuoiThang(pNamTC, tcuoi)
    
    Set rs = Nothing
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "333.XLS"
KetThuc:
End Sub

Private Sub Baocao334(tdau As Integer, tcuoi As Integer)
    Dim TK As New ClsTaikhoan, dn As Double, DC As Double, nt As Double
    
    Recycle pCurDir + "334.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\334.XLS", pCurDir + "334.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "334.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi)
    
    xlsheet.Cells(16, 4) = PSDu("", "334", tdau, tcuoi)
    xlsheet.Cells(16, 5) = PSDu("", "334", pThangDauKy, tcuoi)
    
    xlsheet.Cells(17, 4) = PSDu("", "3383", tdau, tcuoi)
    xlsheet.Cells(17, 5) = PSDu("", "3383", pThangDauKy, tcuoi)
        
    Set xlsheet = xlapp.Worksheets(2)
    TK.InitTaikhoanSohieu "3383"
    TK.SoDuTK ThangTruoc(tdau), dn, DC, nt
    xlsheet.Cells(8, 3) = DC - dn
    TK.SoPhatSinh tdau, tcuoi, dn, DC
    xlsheet.Cells(10, 3) = DC
    xlsheet.Cells(17, 3) = dn
    
    TK.InitTaikhoanSohieu "3384"
    TK.SoDuTK ThangTruoc(tdau), dn, DC, nt
    xlsheet.Cells(8, 4) = DC - dn
    TK.SoPhatSinh tdau, tcuoi, dn, DC
    xlsheet.Cells(10, 4) = DC
    xlsheet.Cells(17, 4) = dn
    
    TK.InitTaikhoanSohieu "3382"
    TK.SoDuTK ThangTruoc(tdau), dn, DC, nt
    xlsheet.Cells(8, 5) = DC - dn
    TK.SoPhatSinh tdau, tcuoi, dn, DC
    xlsheet.Cells(10, 5) = DC
    xlsheet.Cells(17, 5) = dn
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "334.XLS"
KetThuc:
    Set TK = Nothing
End Sub

Private Sub Baocao338(tdau As Integer, tcuoi As Integer)
    Dim rs As Recordset, i As Integer, ncuoi As Date, j As Integer, kq As Double
    Dim sqln As String, sqlc As String, sqlnx As String, sqlcx As String
    Dim T(1 To 6) As Double
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "338.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\338.XLS", pCurDir + "338.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "338.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "337", 2, tcuoi, 1, 2, 3, 4, 7, 4, ncuoi
        
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "338", 2, tcuoi, 1, 2, 3, 4, 7, 4, ncuoi
    
    Set xlsheet = xlapp.Worksheets(4)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    SheetSoDu xlsheet, "3381", 3, tcuoi, 1, 2, 4, 5, 7, 5, ncuoi
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqln = sqln + " + No_" + CStr(i)
        sqlc = sqlc + " + Co_" + CStr(i)
    Next
    For i = 1 To tcuoi
        sqlnx = sqlnx + " + No_" + CStr(i)
        sqlcx = sqlcx + " + Co_" + CStr(i)
    Next
    
    i = 0
    j = 0
    Set xlsheet = xlapp.Worksheets(3)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi)
    Set rs = DBKetoan.OpenRecordset("SELECT SoHieu, Cap, Ten, GhiChu, Sum(DuCo_0-DuNo_0) AS DuDK, Sum(" + sqln + ") AS PSN, Sum(" + sqlc + ") AS PSC, Sum(" + sqlnx + ") AS PSNx, Sum(" + sqlcx + ") AS PSCx, Sum(DuCo_" + CStr(CThangDB(tcuoi)) + "-DuNo_" + CStr(CThangDB(tcuoi)) + ") AS DuCK " _
        & " FROM HethongTK WHERE (SoHieu LIKE '335*' AND Cap>1) AND (DuCo_0-DuNo_0<>0 OR " + sqlnx + "<>0 OR " + sqlcx + "<>0) GROUP BY SoHieu, Cap, Ten, GhiChu ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If rs!cap = 2 Then
            j = j + 1
            xlsheet.Cells(9 + i, 1) = CStr(j)
            xlsheet.Range("A" + CStr(9 + i), "H" + CStr(9 + i)).Font.Bold = True
            T(1) = T(1) + rs!DuDk
            T(2) = T(2) + rs!psn
            T(3) = T(3) + rs!psc
            T(4) = T(4) + rs!psnx
            T(5) = T(5) + rs!pscx
            T(6) = T(6) + rs!DuCK
        End If
        xlsheet.Cells(9 + i, 2) = Space(2 * (rs!cap - 1)) + rs!Ten
        xlsheet.Cells(9 + i, 3) = rs!DuDk
        xlsheet.Cells(9 + i, 8) = rs!DuCK
        xlsheet.Cells(9 + i, 4) = rs!psn
        xlsheet.Cells(9 + i, 5) = rs!psc
        xlsheet.Cells(9 + i, 6) = rs!psnx
        xlsheet.Cells(9 + i, 7) = rs!pscx
        rs.MoveNext
    Loop
    xlsheet.Range("A" + CStr(9 + i), "H" + CStr(9 + i)).Font.Bold = True
    xlsheet.Cells(10 + i, 2) = "Tæng céng"
    For j = 1 To 6
        xlsheet.Cells(10 + i, j + 2) = T(j)
    Next
    With xlsheet.Range("A6", "H" + CStr(10 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    XLSFooter xlsheet, 12 + i, 8, NgayCuoiThang(pNamTC, tcuoi)
    
    rs.Close
    Set rs = Nothing
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "338.XLS"
KetThuc:
End Sub

Private Sub Baocao411(tdau As Integer, tcuoi As Integer)
    Dim i As Integer, ncuoi As Date, TK As New ClsTaikhoan
    Dim ckn As Double, ckc As Double, cknt As Double
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "411.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\411.XLS", pCurDir + "411.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "411.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(2, 1) = ThoiGian(tdau, tcuoi)
    
    For i = 1 To 2
        TK.InitTaikhoanSohieu "41110" + CStr(i)
        TK.SoDuTK ThangTruoc(tdau), ckn, ckc, cknt
        xlsheet.Cells(7 + i, 4) = ckc - ckn
        TK.SoPhatSinh tdau, tcuoi, ckn, ckc
        xlsheet.Cells(11 + i, 4) = ckc
        xlsheet.Cells(15 + i, 4) = ckn
        
        TK.SoDuTK 0, ckn, ckc, cknt
        xlsheet.Cells(7 + i, 9) = ckc - ckn
        TK.SoPhatSinh pThangDauKy, tcuoi, ckn, ckc
        xlsheet.Cells(11 + i, 9) = ckc
        xlsheet.Cells(15 + i, 9) = ckn
    Next
    
    For i = 1 To 2
        TK.InitTaikhoanSohieu "41120" + CStr(i)
        TK.SoDuTK ThangTruoc(tdau), ckn, ckc, cknt
        xlsheet.Cells(7 + i, 5) = ckc - ckn
        TK.SoPhatSinh tdau, tcuoi, ckn, ckc
        xlsheet.Cells(11 + i, 5) = ckc
        xlsheet.Cells(15 + i, 5) = ckn
        
        TK.SoDuTK 0, ckn, ckc, cknt
        xlsheet.Cells(7 + i, 10) = ckc - ckn
        TK.SoPhatSinh pThangDauKy, tcuoi, ckn, ckc
        xlsheet.Cells(11 + i, 10) = ckc
        xlsheet.Cells(15 + i, 10) = ckn
    Next
    
    Set xlsheet = xlapp.Worksheets(4)
    xlsheet.Cells(2, 1) = ThoiGian(tdau, tcuoi)
    
    TK.InitTaikhoanSohieu "441"
    TK.SoDuTK ThangTruoc(tdau), ckn, ckc, cknt
    xlsheet.Cells(7, 3) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(8, 3) = ckc
    xlsheet.Cells(16, 3) = ckn

    TK.SoDuTK 0, ckn, ckc, cknt
    xlsheet.Cells(7, 8) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(8, 8) = ckc
    xlsheet.Cells(16, 8) = ckn
    
    Set xlsheet = xlapp.Worksheets(5)
    xlsheet.Cells(5, 1) = ThoiGian(tdau, tcuoi)
    TK.InitTaikhoanSohieu "414"
    TK.SoDuTK ThangTruoc(tdau), ckn, ckc, cknt
    xlsheet.Cells(10, 4) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(12, 4) = ckc
    xlsheet.Cells(18, 4) = ckn

    TK.SoDuTK 0, ckn, ckc, cknt
    xlsheet.Cells(10, 5) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(12, 5) = ckc
    xlsheet.Cells(18, 5) = ckn
    
    TK.InitTaikhoanSohieu "415"
    TK.SoDuTK ThangTruoc(tdau), ckn, ckc, cknt
    xlsheet.Cells(27, 4) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(29, 4) = ckc
    xlsheet.Cells(35, 4) = ckn

    TK.SoDuTK 0, ckn, ckc, cknt
    xlsheet.Cells(27, 5) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(29, 5) = ckc
    xlsheet.Cells(35, 5) = ckn
    
    TK.InitTaikhoanSohieu "416"
    TK.SoDuTK ThangTruoc(tdau), ckn, ckc, cknt
    xlsheet.Cells(46, 4) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(48, 4) = ckc
    xlsheet.Cells(53, 4) = ckn

    TK.SoDuTK 0, ckn, ckc, cknt
    xlsheet.Cells(46, 5) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(48, 5) = ckc
    xlsheet.Cells(53, 5) = ckn
    
    TK.InitTaikhoanSohieu "4311"
    TK.SoDuTK ThangTruoc(tdau), ckn, ckc, cknt
    xlsheet.Cells(62, 4) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(64, 4) = ckc
    xlsheet.Cells(70, 4) = ckn

    TK.SoDuTK 0, ckn, ckc, cknt
    xlsheet.Cells(62, 5) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(64, 5) = ckc
    xlsheet.Cells(70, 5) = ckn
    
    TK.InitTaikhoanSohieu "4312"
    TK.SoDuTK ThangTruoc(tdau), ckn, ckc, cknt
    xlsheet.Cells(79, 4) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(81, 4) = ckc
    xlsheet.Cells(86, 4) = ckn

    TK.SoDuTK 0, ckn, ckc, cknt
    xlsheet.Cells(79, 5) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(81, 5) = ckc
    xlsheet.Cells(86, 5) = ckn
    
    TK.InitTaikhoanSohieu "451"
    TK.SoDuTK ThangTruoc(tdau), ckn, ckc, cknt
    xlsheet.Cells(100, 4) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(102, 4) = ckc
    xlsheet.Cells(107, 4) = ckn

    TK.SoDuTK 0, ckn, ckc, cknt
    xlsheet.Cells(100, 5) = ckc - ckn
    TK.SoPhatSinh tdau, tcuoi, ckn, ckc
    xlsheet.Cells(102, 5) = ckc
    xlsheet.Cells(107, 5) = ckn
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "411.XLS"
KetThuc:
    Set TK = Nothing
End Sub

Private Sub Baocao632(tdau As Integer, tcuoi As Integer)
    Recycle pCurDir + "632.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\632.XLS", pCurDir + "632.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "632.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = ThoiGian(tdau, tcuoi)
    
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = ThoiGian(tdau, tcuoi)
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "632.XLS"
KetThuc:
End Sub

Private Sub Baocao642(tdau As Integer, tcuoi As Integer)
    Dim rs As Recordset, i As Integer, ncuoi As Date, j As Integer
    Dim sqln As String, sqlnx As String, kq1 As Double, kq2 As Double
    
    Recycle pCurDir + "642.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\642.XLS", pCurDir + "642.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "642.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = ThoiGian(tdau, tcuoi)
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        sqln = sqln + " + No_" + CStr(i)
    Next
    For i = 1 To tcuoi
        sqlnx = sqlnx + " + No_" + CStr(i)
    Next
    
    i = 0
    Set rs = DBKetoan.OpenRecordset("SELECT SoHieu, Cap, Ten, GhiChu, Sum(" + sqln + ") AS PSN,Sum(" + sqlnx + ") AS PSNx " _
        & " FROM HethongTK WHERE SoHieu LIKE '642*' AND Cap>1 AND (" + sqlnx + "<>0) GROUP BY SoHieu, Cap, Ten, GhiChu ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If rs!cap = 2 Then
            xlsheet.Range("A" + CStr(11 + i), "I" + CStr(11 + i)).Font.Bold = True
            j = j + 1
            xlsheet.Cells(11 + i, 1) = CStr(j)
            kq1 = kq1 + rs!psn
            kq2 = kq2 + rs!psnx
        End If
        xlsheet.Cells(11 + i, 2) = Space(2 * (rs!cap - 1)) + rs!Ten
        xlsheet.Cells(11 + i, 6) = rs!psn
        xlsheet.Cells(11 + i, 8) = rs!psnx
        rs.MoveNext
    Loop
    xlsheet.Range("A" + CStr(12 + i), "I" + CStr(12 + i)).Font.Bold = True
    xlsheet.Cells(12 + i, 2) = "Tæng céng"
    xlsheet.Cells(12 + i, 6) = kq1
    xlsheet.Cells(12 + i, 8) = kq2
    If i > 0 Then
        With xlsheet.Range("A9", "I" + CStr(12 + i))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(0, 0, 0)
            .RowHeight = 18
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "642.XLS"
KetThuc:
End Sub

Private Sub Baocao911(tdau As Integer, tcuoi As Integer)
    Dim ps As Double
    
    Recycle pCurDir + "911.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\911.XLS", pCurDir + "911.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "911.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(5, 1) = ThoiGian(tdau, tcuoi)
    
    InCdts pThangDauKy, tcuoi, 0
    ps = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=100 OR MaSo=200")
    If ps > 0 Then
        xlsheet.Cells(10, 4) = 100 * SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=200") / ps
        xlsheet.Cells(10, 5) = xlsheet.Cells(10, 4)
    End If
        
    InKqkd tdau, tcuoi, 0, 0
    ps = SoPSTK("511", tdau, tcuoi, 1)
    ps = ps + SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=21") + SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=31")
    If ps <> 0 Then xlsheet.Cells(13, 4) = 100 * SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=50") / ps
        
    InKqkd pThangDauKy, tcuoi, 0, 0
    ps = SoPSTK("511", pThangDauKy, tcuoi, 1)
    ps = ps + SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=21") + SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=31")
    If ps <> 0 Then xlsheet.Cells(13, 5) = 100 * SelectSQL("SELECT Sum(KyNay) AS F1 FROM Kqkd WHERE MaSo=50") / ps
    
    ps = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=300")
    If ps > 0 Then xlsheet.Cells(19, 4) = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=100 OR MaSo=200") / ps
        
    ps = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=310")
    If ps > 0 Then xlsheet.Cells(22, 4) = SelectSQL("SELECT Sum(CuoiKy) AS F1 FROM Cdts WHERE MaSo=110 OR MaSo=120") / ps
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "911.XLS"
KetThuc:
End Sub

Private Sub SheetSoDu(xlsheet As Worksheet, shtk As String, captk As Integer, tcuoi As Integer, tt_col As Integer, ten_col As Integer, sd_col As Integer, gc_col As Integer, fix_row As Integer, fix_col As Integer, n As Date)
    Dim i As Integer, j As Integer, T As Double, rs As Recordset, id As Long, k As Integer, mpl As Long, rs2 As Recordset
    
    id = GetTK_ID(shtk, 0, k)
    If id <> TKCNKH_ID And id <> TKCNPT_ID Then
        Set rs = DBKetoan.OpenRecordset("SELECT MaSo, SoHieu, Cap, Ten, TKCon, TK_ID, GhiChu, (DuNo_" + CStr(CThangDB(tcuoi)) + "-DuCo_" + CStr(CThangDB(tcuoi)) + ") AS SoDu FROM HethongTK WHERE SoHieu LIKE '" + shtk + "*' AND Cap>=" + CStr(captk) + " ORDER BY SoHieu", dbOpenSnapshot)
        Do While Not rs.EOF
            i = i + 1
            If rs!cap = captk Then
                j = j + 1
                xlsheet.Cells(fix_row + i, 1) = CStr(j)
                xlsheet.Range("A" + CStr(fix_row + i), "G" + CStr(fix_row + i)).Font.Bold = True
                T = T + IIf(k <= 0, rs!sodu, -rs!sodu)
            End If
            If ten_col > 0 Then xlsheet.Cells(fix_row + i, ten_col) = Space(2 * (rs!cap - 1)) + rs!Ten
            If gc_col > 0 Then xlsheet.Cells(fix_row + i, gc_col) = rs!GhiChu
            
            If sd_col > 0 Then xlsheet.Cells(fix_row + i, sd_col) = IIf(k <= 0, rs!sodu, -rs!sodu)
            
            If (rs!tk_id = TKCNKH_ID Or rs!tk_id = TKCNPT_ID) And rs!tkcon = 0 Then
                Set rs2 = DBKetoan.OpenRecordset("SELECT SoHieu, Ten, (DuNo_" + CStr(CThangDB(tcuoi)) + "-DuCo_" + CStr(CThangDB(tcuoi)) + ") AS SoDu FROM KhachHang INNER JOIN SoDuKhachHang ON KhachHang.MaSo=SoDuKhachHang.MaKhachHang WHERE MaTaiKhoan=" + CStr(rs!MaSo) + " AND (DuNo_" + CStr(CThangDB(tcuoi)) + "-DuCo_" + CStr(CThangDB(tcuoi)) + ")<>0 ORDER BY SoHieu", dbOpenSnapshot)
                Do While Not rs2.EOF
                    i = i + 1
                    If ten_col > 0 Then xlsheet.Cells(fix_row + i, ten_col) = Space(2 * (rs!cap - 1)) + rs2!Ten
                    If sd_col > 0 Then xlsheet.Cells(fix_row + i, sd_col) = IIf(k <= 0, rs2!sodu, -rs2!sodu)
                    rs2.MoveNext
                Loop
                rs2.Close
            End If
            
            rs.MoveNext
        Loop
        xlsheet.Cells(fix_row + 1 + i, sd_col) = T
    Else
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaPhanLoai, PhanLoaiKhachHang.TenPhanLoai, KhachHang.SoHieu, KhachHang.Ten, SoDuKhachHang.DuNo_" + CStr(CThangDB(tcuoi)) + "-SoDuKhachHang.DuCo_" + CStr(CThangDB(tcuoi)) + " AS SoDu,KhachHang.GhiChu FROM ((SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo) INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo) INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo WHERE HethongTK.SoHieu LIKE '" + shtk + "*' ORDER BY PhanLoaiKhachHang.SoHieu,KhachHang.SoHieu", dbOpenSnapshot)
        Do While Not rs.EOF
            i = i + 1
            If mpl = rs!MaPhanLoai Then
                j = j + 1
                If tt_col > 0 Then xlsheet.Cells(fix_row + i, tt_col) = CStr(j)
                If ten_col > 0 Then xlsheet.Cells(fix_row + i, ten_col) = rs!Ten
                If gc_col > 0 Then xlsheet.Cells(fix_row + i, gc_col) = rs!GhiChu
                If sd_col > 0 Then xlsheet.Cells(fix_row + i, sd_col) = IIf(k <= 0, rs!sodu, -rs!sodu)
                rs.MoveNext
            Else
                mpl = rs!MaPhanLoai
                xlsheet.Cells(fix_row + i, 2) = rs!TenPhanLoai
                xlsheet.Range("B" + CStr(fix_row + i)).Font.Bold = True
            End If
        Loop
        If i > 0 Then xlsheet.Range(XLSCol(sd_col) + CStr(fix_row + 1 + i)).Formula = "=SUM(" + XLSCol(sd_col) + CStr(fix_row + 1) + ":" + XLSCol(sd_col) + CStr(fix_row + i) + ")"
    End If
    If i > 0 Then
        xlsheet.Cells(fix_row + 1 + i, 2) = "Tæng céng"
        xlsheet.Range("A" + CStr(fix_row + 1 + i), XLSCol(fix_col) + CStr(fix_row + 1 + i)).Font.Bold = True
        With xlsheet.Range("A" + CStr(fix_row + 1), XLSCol(fix_col) + CStr(fix_row + 1 + i))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(0, 0, 0)
            .RowHeight = 18
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
    XLSFooter xlsheet, fix_row + 2 + i, fix_col, n
    rs.Close
    Set rs = Nothing
    Set rs2 = Nothing
End Sub

Private Sub SheetSoDu2(xlsheet As Worksheet, shtk As String, captk As Integer, tcuoi As Integer, tt_col As Integer, ten_col As Integer, sd_col As Integer, gc_col As Integer, fix_row As Integer, fix_col As Integer, n As Date)
    Dim i As Integer, j As Integer, T As Double, rs As Recordset, id As Long, k As Integer, mpl As Long, K1 As Integer
    
    id = GetTK_ID(shtk, 0, K1)
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaKhachHang, KhachHang.SoHieu, KhachHang.Ten, HethongTK.Ten AS TenTK, (SoDuKhachHang.DuNo_" + CStr(CThangDB(tcuoi)) + "-SoDuKhachHang.DuCo_" + CStr(CThangDB(tcuoi)) + ") AS SoDu FROM (SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo) INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo WHERE HethongTK.SoHieu LIKE '" + shtk + "*' ORDER BY KhachHang.SoHieu,HethongTK.SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If mpl = rs!MaKhachHang Then
            j = j + 1
            If tt_col > 0 Then xlsheet.Cells(fix_row + i, tt_col) = CStr(j)
            If ten_col > 0 Then xlsheet.Cells(fix_row + i, ten_col) = rs!tentk
            If sd_col > 0 Then xlsheet.Cells(fix_row + i, sd_col) = IIf(K1 <= 0, rs!sodu, -rs!sodu)
            T = T + IIf(K1 <= 0, rs!sodu, -rs!sodu)
            rs.MoveNext
        Else
            If k > 0 Then
                xlsheet.Range(XLSCol(sd_col) + CStr(k)).Formula = "=SUM(" + XLSCol(sd_col) + CStr(k + 1) + ":" + XLSCol(sd_col) + CStr(fix_row + i - 1) + ")"
                xlsheet.Range(XLSCol(sd_col) + CStr(k)).Font.Bold = True
            End If
            k = fix_row + i
            j = 0
            mpl = rs!MaKhachHang
            xlsheet.Cells(fix_row + i, 2) = rs!Ten
            xlsheet.Range("B" + CStr(fix_row + i)).Font.Bold = True
        End If
    Loop
    
    i = i + 1
    j = 0
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo, SoHieu, Cap, Ten, TKCon, TK_ID, GhiChu, (DuNo_" + CStr(CThangDB(tcuoi)) + "-DuCo_" + CStr(CThangDB(tcuoi)) + ") AS SoDu FROM HethongTK WHERE SoHieu LIKE '" + shtk + "*' AND Cap>=" + CStr(captk) + " AND (TK_ID<>" + CStr(TKCNKH_ID) + " AND TK_ID<>" + CStr(TKCNPT_ID) + ") ORDER BY SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        i = i + 1
        If rs!cap = captk Then
            j = j + 1
            xlsheet.Cells(fix_row + i, 1) = CStr(j)
            xlsheet.Range("A" + CStr(fix_row + i), "G" + CStr(fix_row + i)).Font.Bold = True
            T = T + IIf(K1 <= 0, rs!sodu, -rs!sodu)
        End If
        If ten_col > 0 Then xlsheet.Cells(fix_row + i, ten_col) = Space(2 * (rs!cap - 1)) + rs!Ten
        If gc_col > 0 Then xlsheet.Cells(fix_row + i, gc_col) = rs!GhiChu
        
        If sd_col > 0 Then xlsheet.Cells(fix_row + i, sd_col) = IIf(K1 <= 0, rs!sodu, -rs!sodu)
                
        rs.MoveNext
    Loop
    
    If i > 0 Then
        'xlsheet.Range(XLSCol(sd_col) + CStr(fix_row + 1 + i)).Formula = "=SUM(" + XLSCol(sd_col) + CStr(fix_row + 1) + ":" + XLSCol(sd_col) + CStr(fix_row + i) + ")"
        xlsheet.Cells(fix_row + 1 + i, sd_col) = T
        xlsheet.Cells(fix_row + 1 + i, 2) = "Tæng céng"
        xlsheet.Range("A" + CStr(fix_row + 1 + i), XLSCol(fix_col) + CStr(fix_row + 1 + i)).Font.Bold = True
        With xlsheet.Range("A" + CStr(fix_row + 1), XLSCol(fix_col) + CStr(fix_row + 1 + i))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(0, 0, 0)
            .RowHeight = 18
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
    XLSFooter xlsheet, fix_row + 2 + i, fix_col, n
    rs.Close
    Set rs = Nothing
End Sub
