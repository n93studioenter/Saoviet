VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmLocImport 
   Caption         =   "Import Data"
   ClientHeight    =   4725
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   12480
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   4725
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12000
      Top             =   720
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Text            =   "Combo2"
      Top             =   120
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin MSGrid.Grid Grid1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   12015
      _Version        =   65536
      _ExtentX        =   21193
      _ExtentY        =   6588
      _StockProps     =   77
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   20
      Cols            =   8
   End
   Begin VB.Label Label1 
      Caption         =   "Den"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Tu"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmLocImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public colSelect As String
Public rowSelect As String




Private Sub LocData(fromdate As Integer, todate As Integer)

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim folderPath As String
    Dim FilePath As String
    ' Ðu?ng d?n t?i thu m?c c?n l?y t?p
    folderPath = "C:\TCP\Saoviet\Hoadonchungtu"    ' Thay d?i du?ng d?n này theo thu m?c c?a b?n

    ' T?o d?i tu?ng FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ki?m tra xem thu m?c có t?n t?i không
    If fso.FolderExists(folderPath) Then
        Set folder = fso.GetFolder(folderPath)

        ' Duy?t qua t?ng c?t
        maxWidth = 2000
        'For i = 0 To Grid1.Cols - 1
        Grid1.ColWidth(0) = 1000
        Grid1.ColWidth(1) = 1000
        Grid1.ColWidth(2) = 3000
        Grid1.ColWidth(3) = 2000
        Grid1.ColWidth(4) = 2000
        Grid1.ColWidth(5) = 2000
        ' C?u hình MSFlexGrid
        With Grid1
            .Rows = 1    ' Ð?t l?i s? hàng v? 1, ch? còn l?i tiêu d? c?t
            .Cols = 7    ' S? c?t

            .AddItem "Ngay" & vbTab & "SoHD" & vbTab & "Ten Cty" & vbTab & "Dien giai" & vbTab & "Tong Tien" & vbTab & "No TK" & vbTab & "Co TK" & vbTab & "Ghi chu"      ' Thêm tiêu d? c?t
            'Clear List import
            FrmChungtu.ListReset
            ' Duy?t qua t?ng t?p trong thu m?c
            For Each file In folder.Files
                'Doc de lay ngay ra

                ' Kh?i t?o MSXML
                Dim xmlDoc As Object
                Dim ttChungNode As Object
                Dim shNLapNode As Object
                Dim TTNode As Object
                Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
                xmlDoc.async = False
                xmlDoc.validateOnParse = False
                FilePath = file.path
                If xmlDoc.Load(FilePath) Then
                    ' L?y các node
                    Dim shDonNode As Object
                    Dim shKHHDNode As Object
                    Dim ttNguoiBan As Object
                    Dim getMst As Object

                    Set ttNguoiBan = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/Ten")
                    Set ttChungNode = xmlDoc.selectSingleNode("/HDon/DLHDon/TTChung")
                    Set getMst = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/MST")
                    Set shNLapNode = ttChungNode.getElementsByTagName("NLap")(0)
                    Set shDonNode = ttChungNode.getElementsByTagName("SHDon")(0)
                    Set shKHHDNode = ttChungNode.getElementsByTagName("KHHDon")(0)
                    Set TTNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/TToan/TgTCThue")
                    convertedDate = CDate(shNLapNode.Text)
                    ' Ki?m tra xem tháng c?a convertedDate có n?m trong kho?ng t? fromMonth d?n toMonth không
                    If Month(convertedDate) <= todate Then
                        'Them du lieu cho list frmChungtu
                        Dim getMaTKCo As String
                        Dim splitResult() As String
                        getMaTKCo = GetCusByMST(getMst.Text)
                        splitResult = Split(getMaTKCo, ",")

                        FrmChungtu.AddImportData ttNguoiBan.Text, shDonNode.Text, Format(convertedDate, "dd/mm/yy"), "1", file.path, splitResult(0), splitResult(1), splitResult(2), splitResult(3)
                        .AddItem Format(convertedDate, "dd/mm/yy") & vbTab & shDonNode.Text & vbTab & ttNguoiBan.Text & vbTab & splitResult(3) & vbTab & Format(TTNode.Text, "#,##") & vbTab & splitResult(0) & vbTab & splitResult(1)     ' Thêm d? li?u
                    End If
                End If
            Next file
        End With
    Else
        MsgBox "Thu m?c không t?n t?i!", vbExclamation
    End If

    ' Gi?i phóng b? nh?
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Sub

Private Sub btnImport_Click()
    Me.Hide
    
    FrmChungtu.AutoCLickLoai
End Sub

Private Sub Command1_Click()
Dim fromdate As Integer
Dim todate As Integer
fromdate = Combo1.Text
todate = Combo2.Text
LocData fromdate, todate
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'cbb from
    Combo1.Clear
    ' Vòng l?p d? thêm tháng t? 1 d?n 12
    For i = 1 To 12
        Combo1.AddItem i
    Next i
    Combo1.ListIndex = 0
    'cbb to
     Combo2.Clear
    ' Vòng l?p d? thêm tháng t? 1 d?n 12
    For i = 1 To 12
        Combo2.AddItem i
    Next i
    Combo2.ListIndex = 11
    
    Command1_Click
End Sub

Function GetCusByMST(ByVal MaST As String) As String
    Dim numbers(1) As Integer    ' M?ng 2 ph?n t?

    Dim rs_ktra As Recordset
    Dim Query As String
    Dim rs As DAO.Recordset
    Dim fieldCount As Integer
    Dim i As Integer
    Dim rst As String

    'Lay ra ma kh
    Query = "select * from KhachHang where MST = '" & MaST & "'"
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? các b?n ghi
        Do While Not rs_ktra.EOF
            ' L?y s? lu?ng tru?ng
            rst = rs_ktra.Fields("MaSo").Value

            rs_ktra.MoveNext
        Loop
    Else
        rst = ""
    End If

    If rst = "" Then
        GetCusByMST = ""
    End If

    ' '''''''''''''''''''
    Query = "select * from HoaDon    where MaKhachHang  = " & CInt(rst) & " "
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? các b?n ghi
        Do While Not rs_ktra.EOF
            ' L?y s? lu?ng tru?ng
            rst = rs_ktra.Fields("SoHD").Value
            ' Di chuy?n d?n b?n ghi ti?p theo
            rs_ktra.MoveNext
        Loop
    Else
        rst = ""
    End If

    If rst = "" Then
        GetCusByMST = ""
    End If
    ' ''''''''''''''''


    ' Lay MaTC tu bang chung tu
    Query = "SELECT TOP 2 MaTKNo,MaTKCo,Diengiai FROM ChungTu WHERE SoHieu =  '" & rst & "' ORDER BY MaSo DESC"
    'Query = "SELECT * from  ChungTu"
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    Dim index As Integer
    Dim tkco As Integer
    Dim tkno As Integer
    Dim tkthue As Integer
    Dim diengiai As String
    index = 0
    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? các b?n ghi
        Do While Not rs_ktra.EOF
            ' L?y s? lu?ng tru?ng
            If index = 0 Then
                rst = rs_ktra.Fields("MaTKNo").Value
                tkthue = rst
            Else
                tkno = rs_ktra.Fields("MaTKNo").Value
                tkco = rs_ktra.Fields("MaTKCo").Value
                diengiai = rs_ktra.Fields("Diengiai").Value
            End If

            ' Di chuy?n d?n b?n ghi ti?p theo
            ' Di chuy?n d?n b?n ghi ti?p theo
            index = index + 1
            rs_ktra.MoveNext
        Loop
    Else

    End If

    If rst <> "" Then

    Else
        GetCusByMST = rst
        Exit Function  ' Thoát hàm

    End If

    ' '''''''''''''''''''''''''''''''''''''
    For i = 1 To 3

        ' T?o truy v?n SQL d? l?y thông tin khách hàng theo MST
        If i = 1 Then
            Query = "SELECT TOP 1 * FROM HeThongTK WHERE MaTC = " & tkthue & " ORDER BY NgayKC DESC"
        End If
        If i = 2 Then
            Query = "SELECT TOP 1 * FROM HeThongTK WHERE MaTC = " & tkno & " ORDER BY NgayKC DESC"
        End If
        If i = 3 Then
            Query = "SELECT TOP 1 * FROM HeThongTK WHERE MaTC = " & tkco & " ORDER BY NgayKC DESC"
        End If


        'Query = "SELECT * from  ChungTu"

        ' M? Recordset d? l?y thông tin khách hàng
        Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

        If Not rs_ktra.EOF Then
            ' Duy?t qua t?t c? các b?n ghi
            Do While Not rs_ktra.EOF
                ' L?y s? lu?ng tru?ng
                rst = rs_ktra.Fields("SoHieu").Value
                If i = 1 Then
                    tkthue = rst
                End If
                If i = 2 Then
                    tkno = rst
                End If
                If i = 3 Then
                    tkco = rst
                End If
                ' Di chuy?n d?n b?n ghi ti?p theo
                rs_ktra.MoveNext
            Loop
        Else
            GetCusByMST = rst
            Exit Function  ' Thoát hàm
        End If

    Next i

    ' Ðóng Recordset khi không còn s? d?ng
    rs_ktra.Close
    Set rs_ktra = Nothing
    Dim result As String
    result = tkno & "," & tkco & "," & tkthue & "," & diengiai
    
    GetCusByMST = result
End Function
Public Sub ChangeValueInpput(ByVal values As String)
    Grid1.Row = rowSelect
    Grid1.col = colSelect
    Grid1.Text = values
End Sub
Private Sub Grid1_DblClick()
' L?y giá tr? ô hi?n t?i
    Dim Value As String
    ' Gi? s? VBGrid1 là tên c?a Grid Control
    rowSelect = Grid1.Row
    colSelect = Grid1.col
    ftmInput.Show vbModal
    
    
    
    ' Hi?n th? giá tr? ô
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub
