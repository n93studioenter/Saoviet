VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form2"
   ClientHeight    =   3075
   ClientLeft      =   3900
   ClientTop       =   1980
   ClientWidth     =   4560
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3075
   ScaleWidth      =   4560
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin MSMask.MaskEdBox MaskEdBox 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSMask.MaskEdBox MaskEdBox 
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSForms.Label Label1 
      Height          =   1020
      Left            =   1290
      TabIndex        =   0
      Top             =   375
      Width           =   2340
      BackColor       =   14737632
      Size            =   "4128;1799"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LoadFile
End Sub
Sub LoadFile()
' Read the Text file and display its content in TextBox1(0)
    'If fromEncoding.Text = "Unicode" Then
        'TextBox1(0).Font.Name = "Tahoma"  ' Use Font Tahoma for Unicode
        ' Instantiate an object of class clsUnicodeText
        
        Set MyUnicodeText = New Class1
        ' Read Unicode Text from file txtFileName and display in TextBox1(0)
       ' TextBox1(0).Text = MyUnicodeText.ReadUnicode(txtFileName)
        Label1.Caption = MyUnicodeText.ReadUnicode("D:\vb6\1\ThoiHo.xml")
End Sub
Public Function mang_chuoi(st As String) As String()
Dim mang(100) As String
For i = 1 To 12
If (Len(st) <= 40) And Len(st) > 0 Then
    mang(i) = st
    st = ""
    Exit For
ElseIf Len(st) = 0 Then
    mang(i) = ""
Else
    mang(i) = SUMArr(st, 40)
End If
If Len(mang(i)) > 0 And Len(st) > 0 Then
    st = RTrim(LTrim(Mid(st, Len(mang(i)) + 1, Len(st))))
End If
Next

mang_chuoi = mang
End Function


Function SUMArr(st As String, sokytu As Integer) As String
On Error Resume Next ' Skip Arr "Empty"
    Dim trunggian As String
    Dim i
    Dim nn
    Dim ketqua, chuoi1, chuoi2 As String
    If (Len(st) > 0) Then
    trunggian = Mid(st, 1, sokytu)
    i = Len(Mid(st, 1, sokytu))
    
    'MsgBox Mid(st, sokytu, sokytu + 1)
    If (Mid(st, sokytu + 1, 1) = " ") Then
      ketqua = trunggian
      i = -1
    End If
    
   Do While i > 0
    chuoi1 = Mid(trunggian, i, 1)
     If (chuoi1 = " ") Then
        ketqua = Mid(trunggian, 1, i - 1)
        i = -1
     End If
     i = i - 1
    Loop
    End If
    SUMArr = ketqua
   
End Function
