VERSION 5.00
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmTaiLieu 
   Caption         =   "Tai liÖu l­u tr÷"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   9555
   Icon            =   "frmTaiLieu.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   8835
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin MSGrid.Grid Grid 
      Height          =   4365
      Left            =   9900
      TabIndex        =   0
      Top             =   180
      Width           =   600
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   7699
      _StockProps     =   77
      BackColor       =   16777215
      FixedRows       =   0
      HighLight       =   0   'False
   End
   Begin MSOutl.Outline Outline 
      Height          =   8775
      Left            =   -120
      TabIndex        =   1
      Top             =   -120
      Width           =   9495
      _Version        =   65536
      _ExtentX        =   16748
      _ExtentY        =   15478
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNI-Times"
         Size            =   9.89
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      MouseIcon       =   "frmTaiLieu.frx":57E2
      PicturePlus     =   "frmTaiLieu.frx":57FE
      PictureMinus    =   "frmTaiLieu.frx":58F8
      PictureLeaf     =   "frmTaiLieu.frx":59F2
      PictureOpen     =   "frmTaiLieu.frx":5AEC
      PictureClosed   =   "frmTaiLieu.frx":5BE6
   End
End
Attribute VB_Name = "frmTaiLieu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mang(100) As String


Sub LayFile(ByVal ThuMuc As String, ByVal i As Integer)
Dim f As String
Dim thoat
thoat = 0


If Right(ThuMuc, 1) <> "\" Then
ThuMuc = ThuMuc & "\"
ElseIf Right(ThuMuc, 1) = "\" Then

End If

f = Dir$(ThuMuc & "*.*")
'List1.Clear
While Len(f)
 Outline.AddItem f
 Outline.indent(i) = 2
 Outline.ItemData(i) = 0

  ' Grid.AddItem Str(i) + Chr(9) + f, 0
   ' List1.AddItem F
    f = Dir$
    i = i + 1
Wend

End Sub
Public Sub ColumnSetUp(Grid_control As Grid, col_index As Integer, col_Width As Integer, col_alignment As Integer)
      Grid_control.Row = 0
      Grid_control.col = col_index
      Grid_control.ColWidth(col_index) = col_Width
      Grid_control.FixedAlignment(col_index) = col_alignment
      If col_index >= Grid_control.FixedCols Then Grid_control.ColAlignment(col_index) = col_alignment
End Sub

Private Sub Form_Load()

Dim fso As New FileSystemObject
Dim fil As Folder ' File
Dim fil1 As file
Dim file
Dim j As Integer
j = 0
Dim ThuMuc As String
ThuMuc = pCurDir + "tailieu\"
'If Right(ThuMuc, 1) <> "\" Then
'ThuMuc = ThuMuc & "\"
'ElseIf Right(ThuMuc, 1) = "\" Then
'End If

f = Dir$(ThuMuc & "*.*")
While Len(f)
 Outline.AddItem f
 Outline.indent(j) = 1
 Outline.ItemData(j) = 0
  f = Dir$
    mang(j) = ThuMuc + f
  j = j + 1
Wend
 
For Each fil In fso.GetFolder(pCurDir + "tailieu\").SubFolders
 Outline.AddItem fil.Name
 Outline.indent(j) = 1
 Outline.ItemData(j) = 0
  mang(j) = fil
For Each fil1 In fso.GetFolder(fil).Files
 j = j + 1
 Outline.AddItem fil1.Name
 Outline.indent(j) = 2
 Outline.ItemData(j) = 0
  mang(j) = fil1
Next
j = j + 1

Next


End Sub

Private Sub Grid_Click()
'Shell "EXPLORER.EXE " & "D:\hinhanh\bill\" + Grid.Text
End Sub



Private Sub Outline_Click()
'MsgBox mang(Outline.ListIndex)
Shell "Explorer.exe " & mang(Outline.ListIndex), vbMaximizedFocus

End Sub
