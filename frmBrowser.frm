VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBrowser 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Xem HD"
   ClientHeight    =   7305
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   13440
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      ExtentX         =   23098
      ExtentY         =   11456
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim mypath As String
    mypath = App.path & "\Hoadon"
    Dim LoaiHD As String
    If FrmChungtu.txtPhanloaichungtu.Text = 1 Or FrmChungtu.txtPhanloaichungtu.Text = 0 Then
        LoaiHD = "\HDVaoChonLoc"
    Else
        LoaiHD = "\HDRaChonLoc"
    End If
    mypath = mypath & LoaiHD & "\" & Month(CDate(FrmChungtu.CboThang.Text)) & "\HD__" & Month(CDate(FrmChungtu.CboThang.Text)) & "_" & FrmChungtu.txt(0).Text & "_" & FrmChungtu.txtVT(1).Text & ".html"
    'MsgBox FrmChungtu.txt(0).Text
    Dim FilePath As String
    FilePath = mypath
    WebBrowser1.Navigate FilePath
End Sub
