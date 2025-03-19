VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmLocImport 
   Caption         =   "Import Data"
   ClientHeight    =   3615
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   10035
   LinkTopic       =   "Form4"
   ScaleHeight     =   3615
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   9615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      Text            =   "Combo2"
      Top             =   240
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin MSGrid.Grid Grid1 
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   9615
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   3201
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmLocImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

        ' C?u hình MSFlexGrid
        With Grid1
            .Rows = 1    ' Ð?t l?i s? hàng v? 1, ch? còn l?i tiêu d? c?t
            .Cols = 5  ' S? c?t
            .ColWidth(0) = 2000
            .ColWidth(1) = 2000    ' Ð?t d? r?ng c?t Path (3000 twips)
            .ColWidth(2) = 2000    ' Ð?t d? r?ng c?t Time (2000 twips)
            .ColWidth(3) = 2000
            ' Thêm tiêu d? cho c?t
            .AddItem "NBan" & vbTab & "SoHD" & vbTab & "Time" & vbTab & "MaTK" & vbTab & "Noi dung"     ' Thêm tiêu d? c?t
            'Clear List import
            FrmChungtu.ListReset
            ' Duy?t qua t?ng t?p trong thu m?c
            For Each file In folder.Files
                'Doc de lay ngay ra

                ' Kh?i t?o MSXML
                Dim xmlDoc As Object
                Dim ttChungNode As Object
                Dim shNLapNode As Object
                Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
                xmlDoc.async = False
                xmlDoc.validateOnParse = False
                FilePath = file.path
                If xmlDoc.Load(FilePath) Then
                    ' L?y các node
                    Dim shDonNode As Object
                    Dim shKHHDNode As Object
                    Dim ttNguoiBan As Object

                    Set ttNguoiBan = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/Ten")
                    Set ttChungNode = xmlDoc.selectSingleNode("/HDon/DLHDon/TTChung")

                    Set shNLapNode = ttChungNode.getElementsByTagName("NLap")(0)
                    Set shDonNode = ttChungNode.getElementsByTagName("SHDon")(0)
                    Set shKHHDNode = ttChungNode.getElementsByTagName("KHHDon")(0)

                    convertedDate = CDate(shNLapNode.Text)
                    ' Ki?m tra xem tháng c?a convertedDate có n?m trong kho?ng t? fromMonth d?n toMonth không
                    If Month(convertedDate) >= fromdate And Month(convertedDate) <= todate Then
                        'Them du lieu cho list frmChungtu
                        FrmChungtu.AddImportData ttNguoiBan.Text, shDonNode.Text, "6621", Format(convertedDate, "dd/mm/yy"), "1", file.path
                        .AddItem shDonNode.Text & vbTab & Format(convertedDate, "dd/mm/yy") & vbTab & ttNguoiBan.Text & vbTab & "6422" & vbTab & "asdasd"     ' Thêm d? li?u
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

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub
