VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FrmPBCT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chi tiÕt ph©n bæ chi phÝ"
   ClientHeight    =   7575
   ClientLeft      =   255
   ClientTop       =   690
   ClientWidth     =   11310
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPBCT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Detail Expenses Allocation"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7575
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.TextBox txtTon 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   10800
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6360
      Width           =   255
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8520
      Picture         =   "FrmPBCT.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   41
      Tag             =   "&View"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   9720
      MaxLength       =   20
      TabIndex        =   9
      Text            =   "0"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   9720
      Picture         =   "FrmPBCT.frx":6954
      Style           =   1  'Graphical
      TabIndex        =   37
      Tag             =   "Return"
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   8640
      MaxLength       =   20
      TabIndex        =   8
      Text            =   "0"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   7560
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "0"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   6480
      MaxLength       =   20
      TabIndex        =   6
      Text            =   "0"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5640
      MaxLength       =   20
      TabIndex        =   5
      Text            =   "0"
      Top             =   6360
      Width           =   855
   End
   Begin MSGrid.Grid GrdVT 
      Height          =   5655
      Left            =   0
      TabIndex        =   12
      Tag             =   "30"
      Top             =   720
      Width           =   11295
      _Version        =   65536
      _ExtentX        =   19923
      _ExtentY        =   9975
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   30
      Cols            =   11
      FixedRows       =   0
      FixedCols       =   3
      ScrollBars      =   2
      HighLight       =   0   'False
   End
   Begin VB.CommandButton cmdct 
      Caption         =   "­"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11040
      TabIndex        =   11
      Tag             =   "-1"
      Top             =   6360
      Width           =   255
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "0"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox txtTon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PB"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   10800
      TabIndex        =   42
      Tag             =   "Unit"
      Top             =   480
      Width           =   255
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   13
      Left            =   9840
      TabIndex        =   40
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   9795
      TabIndex        =   39
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CPTC"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   9720
      TabIndex        =   38
      Tag             =   "Administration"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   12
      Left            =   8760
      TabIndex        =   36
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   11
      Left            =   7680
      TabIndex        =   35
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   8715
      TabIndex        =   34
      Top             =   120
      Width           =   975
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   7635
      TabIndex        =   33
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CPQL"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   8640
      TabIndex        =   32
      Tag             =   "Administration"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CPBH"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   7560
      TabIndex        =   31
      Tag             =   "Sale Exp."
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   6555
      TabIndex        =   30
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   5595
      TabIndex        =   29
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   4635
      TabIndex        =   28
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   3555
      TabIndex        =   27
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "§· ph©n bæ"
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
      Index           =   9
      Left            =   1920
      TabIndex        =   26
      Tag             =   "Allocated"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   6555
      TabIndex        =   25
      Top             =   120
      Width           =   975
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   5715
      TabIndex        =   24
      Top             =   120
      Width           =   780
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   4635
      TabIndex        =   23
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CP SXC"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   6480
      TabIndex        =   22
      Tag             =   "Other Exp."
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KH TSC§"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   21
      Tag             =   "Depreciation"
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CP NVL"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   20
      Tag             =   "Material"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tæng chi phÝ"
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
      Index           =   6
      Left            =   1680
      TabIndex        =   19
      Tag             =   "Total Expenses"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label LbTien 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   3555
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CP M¸y"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   17
      Tag             =   "Machine"
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CP Nh©n c«ng"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   16
      Tag             =   "Labour"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "§.vÞ"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   15
      Tag             =   "Unit"
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tªn Thµnh phÈm"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   14
      Tag             =   "Description"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sè hiÖu TP"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Tag             =   "Code"
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "FrmPBCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tp As New Cls154
Dim tdau As Integer, tcuoi As Integer
Dim tcpa As Double, tcpb As Double, tcp1 As Double, tcp2 As Double, tcp3 As Double, tcp4 As Double, tcp1x As Double, tcp2x As Double, tcp3x As Double, tcp4x As Double, tcpax As Double, tcpbx As Double

Private Sub CmdCt_Click()
    Dim pba As Double, pbb As Double, pb1 As Double, pb2 As Double, pb3 As Double, pb4 As Double
    
    If tp.MaSo = 0 Then Exit Sub
    
    pba = Cdbl5(txtTon(4).Text)
    pbb = Cdbl5(txtTon(5).Text)
    pb1 = Cdbl5(txtTon(6).Text)
    pb2 = Cdbl5(txtTon(7).Text)
    pb3 = Cdbl5(txtTon(8).Text)
    pb4 = Cdbl5(txtTon(9).Text)
    
    If (pba > tcpa - tcpax) Then
        RFocus txtTon(4)
        Exit Sub
    End If
    
    If (pbb > tcpb - tcpbx) Then
        RFocus txtTon(5)
        Exit Sub
    End If
    
    If (pb1 > tcp1 - tcp1x) Then
        RFocus txtTon(6)
        Exit Sub
    End If
    
    If (pb2 > tcp2 - tcp2x) Then
        RFocus txtTon(7)
        Exit Sub
    End If
    
    If (pb3 > tcp3 - tcp3x) Then
        RFocus txtTon(8)
        Exit Sub
    End If
    
    If (pb4 > tcp4 - tcp4x) Then
        RFocus txtTon(9)
        Exit Sub
    End If
       
    Me.MousePointer = 11
    With GrdVT
        .AddItem tp.sohieu + Chr(9) + tp.TenVattu + Chr(9) + tp.DonVi + Chr(9) + txtTon(3).Text + Chr(9) + Format(pba, Mask_0) + Chr(9) + Format(pbb, Mask_0) + Chr(9) + Format(pb1, Mask_0) + Chr(9) + Format(pb2, Mask_0) + Chr(9) + Format(pb3, Mask_0) + Chr(9) + Format(pb4, Mask_0) + Chr(9) + txtTon(10).Text, IIf(cmdct.tag < 0, 0, cmdct.tag)
        tp.GhiCPPB tdau, tcuoi, 0, pb1 - tp.SoCPSXCTT(tdau, tcuoi)
        tp.GhiCPPB2 tdau, tcuoi, pb2 - tp.SoCPBH(tdau, tcuoi), pb3, pb4
        tp.GhiCPPB3 tdau, tcuoi, pba - tp.SoCPNC(tdau, tcuoi), pbb - tp.SoCPM(tdau, tcuoi)
        .Row = .Rows - 1
        .col = 0
        If Len(.Text) = 0 Then .RemoveItem .Row
        .Row = 0
    End With
    cmdct.tag = -1
    TongTien
    FrmPBCP.tag = 1
XongDK:
    Me.MousePointer = 0
End Sub

Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 0:
            InBaoCao
        Case 1:
            Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    If Me.tag > 0 Then
        Me.MousePointer = 11
        tdau = Me.tag \ 100
        tcuoi = Me.tag Mod 100
        Me.tag = 0
        LietKeCP
        Me.MousePointer = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((Shift And vbAltMask) > 0 And KeyCode = vbKeyV) Or KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    ColumnSetUp GrdVT, 0, 820, 0
    ColumnSetUp GrdVT, 1, 1900, 0
    ColumnSetUp GrdVT, 2, 700, 0
    ColumnSetUp GrdVT, 3, 1060, 1
    ColumnSetUp GrdVT, 4, 1060, 1
    ColumnSetUp GrdVT, 5, 820, 1
    ColumnSetUp GrdVT, 6, 1060, 1
    ColumnSetUp GrdVT, 7, 1060, 1
    ColumnSetUp GrdVT, 8, 1060, 1
    ColumnSetUp GrdVT, 9, 1060, 1
    ColumnSetUp GrdVT, 10, 220, 2
    SetFont Me
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    Set tp = Nothing
End Sub

Private Sub GrdVT_Click()
    Dim i As Integer
    
    With GrdVT
        If .col = 10 Then
            .Text = IIf(.Text = "X", "-", "X")
            i = IIf(.Text = "X", 0, 1)
            .col = 0
            ExecuteSQL5 "UPDATE TP154 SET KPB=" + CStr(i) + " WHERE SoHieu='" + .Text + "'"
        End If
    End With
End Sub

Private Sub GrdVT_DblClick()
    Dim i As Integer
    
    With GrdVT
        .col = 0
        If Len(.Text) = 0 Then Exit Sub
        For i = 0 To 10
            .col = i
            txtTon(i).Text = .Text
        Next
        cmdct.tag = .Row
        .RemoveItem .Row
        TongTien
        tp.InitTPSohieu txtTon(0).Text
        RFocus txtTon(3)
    End With
End Sub

Private Sub GrdVT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then GrdVT_DblClick
End Sub

Private Sub GrdVT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SearchObj 1, , GrdVT, GrdVT.col
    End If
End Sub

Private Sub txtTon_GotFocus(Index As Integer)
    AutoSelect txtTon(Index)
End Sub

Private Sub txtTon_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0:
            If KeyAscii = 13 Then
                txtTon(0).Text = FrmTP.ChonTP(txtTon(0).Text)
            End If
        Case 1, 2, 3:
            KeyAscii = 0
        Case 4, 5, 6, 7:
            If KeyAscii = 13 Then
                CmdCt_Click
            Else
                KeyProcess txtTon(Index), KeyAscii, True
            End If
    End Select
End Sub

Private Sub txtTon_LostFocus(Index As Integer)
    Select Case Index
        Case 0:
            If Len(txtTon(0).Text) > 0 Then
                tp.InitTPSohieu txtTon(0).Text
            Else
                tp.InitTPMaSo 0
            End If
        Case 6, 7:
            txtTon(Index).Text = Format(txtTon(Index).Text, Mask_0)
    End Select
End Sub

Private Sub TongTien()
    Dim i As Integer
    
    tcpax = 0
    tcpbx = 0
    tcp1x = 0
    tcp2x = 0
    tcp3x = 0
    tcp4x = 0

    With GrdVT
        For i = 0 To .Rows - 1
            .Row = i
            .col = 0
            If Len(.Text) = 0 Then Exit For
            .col = 4
            tcpax = tcpax + Cdbl5(.Text)
            .col = 5
            tcpbx = tcpbx + Cdbl5(.Text)
            .col = 6
            tcp1x = tcp1x + Cdbl5(.Text)
            .col = 7
            tcp2x = tcp2x + Cdbl5(.Text)
            .col = 8
            tcp3x = tcp3x + Cdbl5(.Text)
            .col = 9
            tcp4x = tcp4x + Cdbl5(.Text)
        Next
    End With
    
    LbTien(8).Caption = Format(tcpax, Mask_0)
    LbTien(9).Caption = Format(tcpbx, Mask_0)
    LbTien(10).Caption = Format(tcp1x, Mask_0)
    LbTien(11).Caption = Format(tcp2x, Mask_0)
    LbTien(12).Caption = Format(tcp3x, Mask_0)
    LbTien(13).Caption = Format(tcp4x, Mask_0)
End Sub

Private Sub LietKeCP()
    Dim rs As Recordset, cp2 As String, i As Integer
    Dim tcpc As Double, cp3 As String, cp4 As String, cp5 As String, c621 As Double
    Dim s621 As String, s622 As String, s623 As String
    
    Set rs = DBKetoan.OpenRecordset("SELECT MaSo FROM TP154", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        tp.InitTPMaSo rs!MaSo
        tp.GhiCPTT tdau, tcuoi
        tp.XDDauKy tdau
        rs.MoveNext
    Loop
    rs.Close
        
    tcp1 = PSTKCP("627*", tdau, tcuoi)
    tcp2 = PSTKCP("641*", tdau, tcuoi)
    tcp3 = PSTKCP("642*", tdau, tcuoi)
    tcp4 = PSTKCP("635*", tdau, tcuoi)
    c621 = PSTKCP("621*", tdau, tcuoi)
    tcpa = PSTKCP("622*", tdau, tcuoi)
    tcpb = PSTKCP("623*", tdau, tcuoi)
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        cp2 = cp2 + "+CPBH" + CStr(i) + "+CPBHTT" + CStr(i)
        cp3 = cp3 + "+CPQL" + CStr(i) + "+CPQLTT" + CStr(i)
        cp4 = cp4 + "+CPTC" + CStr(i) + "+CPTCTT" + CStr(i)
        cp5 = cp5 + "+CPSXC_" + CStr(i) + "+CPSXCTT" + CStr(i)
        s621 = s621 + "+CPNVLPB" + CStr(i)
        s622 = s622 + "+CPNCPB" + CStr(i)
        s623 = s623 + "+CPMPB" + CStr(i)
    Next
    tcpax = 0
    tcpbx = 0
    tcp1x = 0
    tcp2x = 0
    tcp3x = 0
    tcp4x = 0
    
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW SoHieu, TenVattu,DonVi, CPNVL+" + s621 + " AS NVL, CPNC+" + s622 + " AS NC, CPM+" + s623 + " AS M, (" + cp5 + ") AS CPSXC, (" + cp2 + ") AS CPBH, (" + cp3 + ") AS CPQL, (" + cp4 + ") AS CPTC, KPB" _
        & " FROM TP154  WHERE DK1<>0 OR CPNVL+" + s621 + "<>0 OR CPNC+" + s622 + "<>0 OR CPM+" + s623 + "<>0 OR (" + cp2 + ")<>0 OR (" + cp3 + ")<>0 OR (" + cp4 + ")<>0 OR (" + cp5 + ")<>0 OR DT<>0 ORDER BY SoHieu DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        tcp1x = tcp1x + rs!cpsxc
        tcp2x = tcp2x + rs!cpbh
        tcp3x = tcp3x + rs!cpql
        tcp4x = tcp4x + rs!cptc
        tcpax = tcpax + rs!nvl
        tcpbx = tcpbx + rs!nC
        tcpc = tcpc + rs!m
        
        GrdVT.AddItem rs!sohieu + Chr(9) + rs!TenVattu + Chr(9) + rs!DonVi + Chr(9) + Format(rs!nvl, Mask_0) + Chr(9) + Format(rs!nC, Mask_0) + Chr(9) + Format(rs!m, Mask_0) _
             + Chr(9) + Format(rs!cpsxc, Mask_0) + Chr(9) + Format(rs!cpbh, Mask_0) + Chr(9) + Format(rs!cpql, Mask_0) + Chr(9) + Format(rs!cptc, Mask_0) + Chr(9) + IIf(rs!KPB = 0, "X", "-"), 0
        rs.MoveNext
    Loop
    
    GrdVT.Rows = IIf(rs.RecordCount > GrdVT.tag, rs.RecordCount, GrdVT.tag)
    
    rs.Close
    Set rs = Nothing
    
    LbTien(0).Caption = Format(c621, Mask_0)
    LbTien(1).Caption = Format(tcpa, Mask_0)
    LbTien(2).Caption = Format(tcpb, Mask_0)
    LbTien(3).Caption = Format(tcp1, Mask_0)
    LbTien(4).Caption = Format(tcp2, Mask_0)
    LbTien(5).Caption = Format(tcp3, Mask_0)
    LbTien(6).Caption = Format(tcp4, Mask_0)
    
    LbTien(7).Caption = Format(tcpax, Mask_0)
    LbTien(8).Caption = Format(tcpbx, Mask_0)
    LbTien(9).Caption = Format(tcpc, Mask_0)
    LbTien(10).Caption = Format(tcp1x, Mask_0)
    LbTien(11).Caption = Format(tcp2x, Mask_0)
    LbTien(12).Caption = Format(tcp3x, Mask_0)
    LbTien(13).Caption = Format(tcp4x, Mask_0)

    tcp1x = tcp1
    tcp2x = tcp2
    tp.InitTPMaSo 0
    txtTon(0).Text = ""
    txtTon(1).Text = ""
    txtTon(2).Text = ""
End Sub

Private Sub InBaoCao()
    Dim i As Integer, sh As String, Ten As String, dv As String, j As Integer
    Dim so(3 To 10) As Double
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    
    With GrdVT
        For i = 0 To .Rows - 1
            .Row = i
            .col = 0
            If Len(.Text) = 0 Then Exit For
            sh = .Text
            .col = 1
            Ten = .Text
            .col = 2
            dv = .Text
            For j = 3 To 9
                .col = j
                so(j) = Cdbl5(.Text)
            Next
            ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu, Ten, Kq1, Kq2, Kq3, Kq5, Kq6, Kq7, Kq8) VALUES (" + CStr(i) + ",'" + sh + "', '" + Ten + "'," + DoiDau(so(3)) + "," + DoiDau(so(4)) + "," + DoiDau(so(5)) + "," + DoiDau(so(6)) + "," + DoiDau(so(7)) + "," + DoiDau(so(8)) + "," + DoiDau(so(9)) + ")"
        Next
    End With
    SetRptInfo
    
    For i = 0 To 6
        frmMain.Rpt.Formulas(5 + i) = "CP" + CStr(i + 1) + "=" + DoiDau(Cdbl5(LbTien(i).Caption))
    Next
    
    frmMain.Rpt.ReportFileName = "CHIPHI4.RPT"
    frmMain.Rpt.Formulas(4) = "ThoiGian='" + Me.Caption + "'"
    InBaoCaoRPT
End Sub
