VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Ch­¬ng tr×nh kÕ to¸n Sao ViÖt"
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   1785
   ClientWidth     =   8250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   ScaleHeight     =   246
   ScaleMode       =   2  'Point
   ScaleWidth      =   412.5
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "3"
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   7800
      TabIndex        =   0
      Top             =   6960
      Width           =   3855
      Begin VB.Label LbAbout 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "Warning:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   6
         Tag             =   "Warning"
         Top             =   0
         Width           =   675
      End
      Begin VB.Label LbAbout 
         BackColor       =   &H80000013&
         Caption         =   $"frmSplash.frx":57E2
         Height          =   735
         Index           =   7
         Left            =   750
         TabIndex        =   5
         Tag             =   "Copyright"
         Top             =   2955
         Width           =   5895
      End
      Begin VB.Label LbAbout 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Accounting Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Index           =   2
         Left            =   3075
         TabIndex        =   4
         Tag             =   "Product"
         Top             =   435
         Width           =   3585
      End
      Begin VB.Label LbAbout 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "SAO VIET Software Center"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3195
         TabIndex        =   3
         Tag             =   "CompanyProduct"
         Top             =   0
         Width           =   3750
      End
      Begin VB.Label LbAbout 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Caption         =   "Version 6.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   5295
         TabIndex        =   2
         Tag             =   "Version"
         Top             =   1335
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "WWW.SAOVIET.COM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   0
         TabIndex        =   1
         Top             =   1545
         Width           =   2010
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "dlgCommonDialog"
   End
   Begin VB.Image Image1 
      Height          =   15015
      Left            =   -9600
      Picture         =   "frmSplash.frx":58B4
      Top             =   -2000
      Width           =   34920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   366
      X2              =   366
      Y1              =   -1000
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   2
      X1              =   0
      X2              =   366
      Y1              =   -1000
      Y2              =   -1000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   228
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   20
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Public Sub StartSAS()


'SaveSetting "MyApp", "Settings", "FirstRun", "True"
    pDataPath = GetSetting(IniPath, "Environment", "Path", pCurDir + "DATA\KETOAN.MDB")
    ' §Æt c¸c format
    'pDataPath = "C:\Tao moi\DATA\KT.MDB"

    Mask_0 = GetSetting(IniPath, "Environment", "IntMask", "###,###,###,###")
    If Cdbl5("1,5") <> 1.5 Then sDecimal = "." Else sDecimal = ","
    pThang = GetSetting(IniPath, "Environment", "NDecimal", 2)
    Select Case pThang
    Case 0: Mask_2 = Mask_0
    Case 1: Mask_2 = "###0.0"
    Case 2: Mask_2 = "Standard"
    Case 3: Mask_2 = "###0.000"
    Case 4: Mask_2 = "###0.0000"
    Case Else: Mask_2 = GetSetting(IniPath, "Environment", "DblMask", "Standard")
    End Select
    Mask_N = 10 ^ pThang
    Mask_D = GetShortDateFormat
    ' §äc File INI
    pThangDauKy = GetSetting(IniPath, "Environment", "StartMonth", "1")
    pTKTrunggian = GetSetting(IniPath, "Environment", "TmpAccount", "5")
    SHCT_Len = GetSetting(IniPath, "Environment", "InvCodeLen", 2)

    ShTkSPDo = GetSetting(IniPath, "ProductCost", "TK_SPDO", "154")
    ShTkTP = GetSetting(IniPath, "ProductCost", "TK_SPDO", "155")
    ShTkKQ = GetSetting(IniPath, "ProductCost", "TK_XDKQ", "911")

    Set WSpace = DBEngine.CreateWorkspace(CStr(Time), "Admin", "", dbUseJet)
    Workspaces.Append WSpace

    

    
    Do While OpenDB(pDataPath) <> 0
        dlgCommonDialog.Flags = &H4&
        dlgCommonDialog.fileName = "*.MDB"
        On Error GoTo QuitSAS
        dlgCommonDialog.ShowOpen
        On Error GoTo 0
        pDataPath = dlgCommonDialog.fileName
    Loop

    pThang = Month(Date)
    pWinDir = GetWinDir

    Select Case UCase(App.EXEName)
    Case "SERVER": pProcessMode = 2
    Case "CLIENT": pProcessMode = 1
    Case Else: pProcessMode = 0
    End Select

    Exit Sub
QuitSAS:
    WSpace.Close
    End
End Sub

Private Sub Form_Load()
    Me.Height = 5280
    Me.Width = 8250
    dlgCommonDialog.InitDir = pCurDir + "DATA"
    If DEMO = 1 Then LbAbout(4).Caption = "Training Version"
End Sub

Private Sub Label2_Click()

End Sub

