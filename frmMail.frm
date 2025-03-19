VERSION 5.00
Begin VB.Form frmMail 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Göi tÖp d÷ liÖu b»ng Email"
   ClientHeight    =   3990
   ClientLeft      =   1755
   ClientTop       =   1710
   ClientWidth     =   6495
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
   Icon            =   "frmMail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Send data file by email"
   Begin VB.Frame Frame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   3600
      Width           =   615
      Begin VB.Label Lb 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   4
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   2325
      Index           =   2
      Left            =   1200
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   2
      Top             =   840
      Width           =   3975
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   5280
      Picture         =   "frmMail.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "&Return"
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   5280
      Picture         =   "frmMail.frx":6C04
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "&Send"
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "§· göi (%)"
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   12
      Tag             =   "Sending..."
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TÖp d÷ liÖu"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Tag             =   "Data file"
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Néi dung"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Tag             =   "Content"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tãm t¾t"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Tag             =   "Description"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "§ång göi"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Tag             =   "Cc"
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Göi tíi"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Tag             =   "To"
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SMTPPort = 25
Dim intStatus As Integer
Dim Femail As String
Dim Fsmtp As String

Private Sub Command_Click(Index As Integer)
    Dim lngSMTPSocket As Long
    Dim lngServerAddress As Long
    Dim strServerName As String
    Dim strServerResponse As String
    Dim mmsgCurrentMessage As New MIMEMessage
    Dim strCurrentLine, i As Long, L As Long
    
    
    If Index = 1 Then GoTo KT
    If MsgBox("Tù ®éng cµi ®Æt Dail-up connection ?", vbYesNo + vbInformation, App.ProductName) = vbYes Then StartDUN Me.hwnd
    Me.Refresh
    lngSMTPSocket = 0
    Me.MousePointer = 11
    
    strServerName = Fsmtp
    intStatus = GetIPAddress(lngServerAddress, strServerName)
    intStatus = CreateSocket(lngSMTPSocket, 0)
    intStatus = ConnectSocket(lngSMTPSocket, lngServerAddress, SMTPPort)
    
    intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
    intStatus = SendSocket(lngSMTPSocket, "mail from:<" + Femail + ">")
    
    intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
    intStatus = SendSocket(lngSMTPSocket, "rcpt to:<" + txt(0).Text + ">")
    intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
    intStatus = SendSocket(lngSMTPSocket, "data")
    intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
    
    mmsgCurrentMessage.AddHeader "From", Femail
    mmsgCurrentMessage.AddHeader "To", txt(0).Text
    mmsgCurrentMessage.AddHeader "Subject", txt(1).Text
    
    mmsgCurrentMessage.BodyLines.Add txt(2).Text
    
    If Len(txt(3).Text) > 0 And Len(Dir(txt(3).Text)) > 0 Then
      mmsgCurrentMessage.AddAttachment txt(3).Text, "Base64"
    End If
            
    mmsgCurrentMessage.Compose
    L = mmsgCurrentMessage.Lines.count
    For Each strCurrentLine In mmsgCurrentMessage.Lines
      i = i + 1
      Lb.Caption = CStr(Fix(100 * i / L))
      Frame.Refresh
      intStatus = SendSocket(lngSMTPSocket, CStr(strCurrentLine))
    Next
    
    intStatus = SendSocket(lngSMTPSocket, ".")
    intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
    intStatus = SendSocket(lngSMTPSocket, "quit")
    intStatus = ReceiveSocket(lngSMTPSocket, strServerResponse)
    intStatus = ReleaseSocket(lngSMTPSocket)
  
KT:
    Set mmsgCurrentMessage = Nothing
    Me.MousePointer = 0
    Unload Me
End Sub

Private Sub Form_Load()
    intStatus = StartWinSock()
        
    SetFont Me
    Me.MousePointer = 0
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    intStatus = WSACleanup()
    HangUp
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyT:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyV:
                RFocus Command(1)
                Command_Click 1
        End Select
    End If
    If KeyCode = vbKeyReturn Then Command_Click 0
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txt_GotFocus(Index As Integer)
    AutoSelect txt(Index)
End Sub

Public Sub SendMail(email As String, smtp As String, FileName As String)
    Femail = email
    Fsmtp = smtp
    
    txt(0).Text = SelectSQL("SELECT EmailDB AS F1 FROM License")
    txt(1).Text = pTenCty + IIf(Left(pTenCn, 1) <> ".", " - " + pTenCn, "") + " - Tep du lieu ngay " + Format(Date, Mask_D)
    txt(2).Text = "Chon muc Mo tep du lieu nen trong menu He thong de mo tep du lieu nhan duoc"
    txt(3).Text = FileName
    
    Me.Show vbModal
End Sub
