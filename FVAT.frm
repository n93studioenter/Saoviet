VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FVAT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Th«ng tin ho¸ ®¬n"
   ClientHeight    =   6015
   ClientLeft      =   4680
   ClientTop       =   3165
   ClientWidth     =   6120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FVAT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Detail Invoice Informaion"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Tag             =   "0"
   Begin VB.TextBox T 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   13
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "0"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox ChkV 
      BackColor       =   &H00E0E0E0&
      Caption         =   "§iÒu chØnh"
      Enabled         =   0   'False
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
      Index           =   5
      Left            =   2520
      TabIndex        =   16
      Tag             =   "Adjustment"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox ChkV 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TSC§"
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   4560
      TabIndex        =   19
      Tag             =   "Fixed Assets"
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CheckBox ChkV 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NhËp khÈu"
      Enabled         =   0   'False
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
      Index           =   3
      Left            =   4560
      TabIndex        =   18
      Tag             =   "Import"
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   12
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   11
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox ChkV 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B¶ng kª thu mua"
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
      Index           =   2
      Left            =   3720
      TabIndex        =   17
      Tag             =   "Retail Bill"
      Top             =   3360
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.TextBox T 
      Alignment       =   1  'Right Justify
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   195
      Index           =   10
      Left            =   5280
      MaxLength       =   20
      TabIndex        =   33
      Text            =   "0"
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox ChkV 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Kh«ng chÞu thuÕ"
      Enabled         =   0   'False
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
      Left            =   4560
      TabIndex        =   15
      Tag             =   "Non-Taxable"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   9
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   8
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   7
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   4095
   End
   Begin VB.CheckBox ChkV 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hãa ®¬n GTGT"
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
      Left            =   3720
      TabIndex        =   14
      Tag             =   "VAT Bill"
      Top             =   3000
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   6
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox T 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   1800
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "FVAT.frx":57E2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox T 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   1800
      MaxLength       =   20
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "FVAT.frx":57E4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   3
      Left            =   1800
      MaxLength       =   500
      TabIndex        =   9
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   2
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   1
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   0
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Height          =   375
      Left            =   4800
      Picture         =   "FVAT.frx":57E6
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "&Save"
      Top             =   5400
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSForms.OptionButton OptChon 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   42
      Top             =   5520
      Width           =   4575
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "8070;661"
      Value           =   "0"
      Caption         =   "Hµng hãa dv kh«ng tæng hîp trªn tê khai 01/GTGT"
      FontName        =   "VK Sans Serif"
      FontEffects     =   1073741828
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton OptChon 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   41
      Top             =   5160
      Width           =   5895
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "10398;661"
      Value           =   "0"
      Caption         =   "Hµng hãa, dv dïng cho dù ¸n  ®Çu t­ ®ñ dk khÊu trõ"
      FontName        =   "VK Sans Serif"
      FontEffects     =   1073741828
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton OptChon 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   40
      Top             =   4800
      Width           =   5895
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "10398;661"
      Value           =   "0"
      Caption         =   "Hµng hãa, dv dïng chung cho KD chÞu thuÕ vµ ko chÞu thuÕ ®ñ ®k khÊu trõ"
      FontName        =   "VK Sans Serif"
      FontEffects     =   1073741828
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton OptChon 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   39
      Top             =   4440
      Width           =   5535
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "9763;661"
      Value           =   "0"
      Caption         =   "Hµng hãa, dÞch vô kh«ng ®ñ ®k khÊu trõ/ (®Çu ra ko chÞu thuÕ)"
      FontName        =   "VK Sans Serif"
      FontEffects     =   1073741828
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton OptChon 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   4080
      Width           =   5775
      BackColor       =   14737632
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "10186;661"
      Value           =   "1"
      Caption         =   "Hµng hãa, dÞch vô dïng riªng cho SXKD chÞu thuÕ GTGT ®ñ ®k khÊu trõ"
      FontName        =   "VK Sans Serif"
      FontEffects     =   1073741828
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ký hiÖu mÉu h®"
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
      Index           =   15
      Left            =   3240
      TabIndex        =   37
      Tag             =   "Ex. Rate"
      Top             =   1920
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "M· h®"
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
      Index           =   14
      Left            =   3600
      TabIndex        =   36
      Tag             =   "Bill Type"
      Top             =   1560
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "H×nh thøc thanh to¸n"
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
      Index           =   13
      Left            =   2880
      TabIndex        =   35
      Tag             =   "Payment Type"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gi¸ tÝnh thuÕ"
      Enabled         =   0   'False
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
      Index           =   12
      Left            =   4800
      TabIndex        =   34
      Tag             =   "Taxable Amount"
      Top             =   3360
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "(NhËp dÊu # cho kh¸ch v·ng lai)"
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
      Index           =   11
      Left            =   3240
      TabIndex        =   32
      Tag             =   "(# for un-frequent liability)"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tû lÖ thuÕ (%)"
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
      Index           =   10
      Left            =   120
      TabIndex        =   31
      Tag             =   "Tax Rate"
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Thµnh tiÒn tr­íc thuÕ"
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
      Left            =   120
      TabIndex        =   30
      Tag             =   "Amount before Tax"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sè l­îng"
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
      Index           =   8
      Left            =   120
      TabIndex        =   29
      Tag             =   "Quantity"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MÆt hµng"
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
      Index           =   7
      Left            =   120
      TabIndex        =   28
      Tag             =   "Items"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ngµy ph¸t hµnh"
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
      Left            =   120
      TabIndex        =   27
      Tag             =   "Bill Date"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ký hiÖu ho¸ ®¬n"
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
      Index           =   5
      Left            =   120
      TabIndex        =   26
      Tag             =   "Bill Code"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sè ho¸ ®¬n"
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
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Tag             =   "Bill Number"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "M· sè thuÕ"
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
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Tag             =   "Tax Code"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "§Þa chØ"
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
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Tag             =   "Address"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tªn kh¸ch hµng"
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
      Left            =   120
      TabIndex        =   22
      Tag             =   "Description"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "M· sè ®¬n vÞ"
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
      Left            =   120
      TabIndex        =   21
      Tag             =   "Liability Code"
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FVAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ngay As Date
Dim ckh As New ClsKhachHang

 Public Sub Command_Click()
    Dim i As Integer
       
    If KHDetail And ckh.MaSo = 0 And Left(T(0).Text, 1) = "#" Then 'neu khach hang la dau thang
        ckh.Ten = T(7).Text
        ckh.DiaChi = T(8).Text
        ckh.mst = T(9).Text
       ' ckh.Tel = "000-00"
       ' ckh.Fax = "000-00"
        ckh.sohieu = "#" + CStr(Year(Date) - 2000) + CStr(Month(Date)) + CStr(Day(Date)) + CStr(Hour(Now)) + CStr(Minute(Now)) + CStr(Second(Now))
        ckh.MaPhanLoai = SelectSQL("SELECT MaSo AS F1 FROM PhanLoaiKhachHang WHERE LEFT(SoHieu,1)='#'")
        If ckh.GhiKhachHang <> 0 Then GoTo Er 'Luu khach hang moi la #
    Else
    Dim maso_khachhang
        ' kiem tra khach hang do da co chua
        maso_khachhang = SelectSQL("SELECT TOP 1 Maso AS F1 FROM KhachHang WHERE sohieu = '" + Trim(T(0).Text) + "'")
        If (maso_khachhang = 0) Then
            ckh.sohieu = T(0).Text
            ckh.Ten = T(7).Text
            ckh.DiaChi = T(8).Text
            ckh.mst = T(9).Text
            
            ckh.Tel = "000"
            ckh.Fax = "0000"
            ckh.email = "000"
            ckh.MaPhanLoai = FrmChungtu.CboLoai.ItemData(FrmChungtu.CboLoai.ListIndex)
            If ckh.GhiKhachHang <> 0 Then GoTo Er 'Luu khach hang moi la #
            Else
             ckh.MaSo = maso_khachhang
        End If
    End If
    If KHDetail And ckh.MaSo = 0 Then
Er:
        RFocus T(0)
        Exit Sub
    End If
    With h ' neu khach hang da co roi , thi lay thong tin dua vao class khach hang h
        .MaKhachHang = ckh.MaSo
        .KyHieu = IIf(Len(T(1).Text) > 0, T(1).Text, "...")
        .soHD = IIf(Len(T(2).Text) > 0, T(2).Text, "...")
        .NgayPH = MedNgay.Text
        .MatHang = IIf(Len(T(3).Text) > 0, T(3).Text, "...")
        .SoLuong = Cdbl5(T(4).Text)
        .ThanhTien = Cdbl5(T(5).Text)
        .TyLe = CInt5(T(6).Text)
        .HD = ChkV(0).Value
        .KCT = ChkV(1).Value
        .HDBL = ChkV(2).Value
        .NK = ChkV(3).Value
        .ts = ChkV(4).Value
        .DC = ChkV(5).Value
        .TenKH = T(7).Text
        .DiaChiKH = T(8).Text
        .MSTKH = T(9).Text
        .HTTT = IIf(Len(T(11).Text) > 0, T(11).Text, "...")
        .MauSo = IIf(Len(T(12).Text) > 0, T(12).Text, "...")
        .tygia = Cdbl5(T(13).Text)
    End With
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyG:
                RFocus Command
                Command_Click
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    If Not KHDetail Then
        T(0).Enabled = False
        For i = 7 To 9
            T(i).Locked = False
            T(i).TabStop = True
        Next
    End If
    
    SetFont Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ckh = Nothing
End Sub

Private Sub MedNgay_GotFocus()
    AutoSelect MedNgay
End Sub

Private Sub MedNgay_LostFocus()
    If IsDate(MedNgay.Text) Then
        ngay = CDate(MedNgay.Text)
    Else
        RFocus MedNgay
    End If
End Sub

Private Sub OptChon_Click(Index As Integer)
    Select Case Index
        Case 0:
        T(11).Text = "1"
        Case 1:
        T(11).Text = "2"
        Case 2:
        T(11).Text = "3"
        Case 3:
        T(11).Text = "4"
        Case 4:
        T(11).Text = "5"
    End Select
End Sub

Private Sub T_Change(Index As Integer)
 If Index = 11 Then
   If T(11).Text = "2" Then
     OptChon(1).Value = True
   ElseIf T(11).Text = "3" Then
     OptChon(2).Value = True
   ElseIf T(11).Text = "4" Then
     OptChon(3).Value = True
  ElseIf T(11).Text = "5" Then
     OptChon(4).Value = True
     Else
     OptChon(0).Value = True
   End If
 End If
End Sub

Public Sub T_GotFocus(Index As Integer)
    AutoSelect T(Index)
End Sub

Private Sub T_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0:
            If KeyAscii = 13 And KHDetail Then
                T(0).Text = FrmKhachHang.ChonKhachHang(T(0).Text)
                RFocus T(0)
            End If
        Case 4, 5, 6, 13:
            KeyProcess T(Index), KeyAscii
    End Select
End Sub

Private Sub T_LostFocus(Index As Integer)
    Dim i As Integer
    If Index = 0 Or Index = 1 Or Index = 12 Or Index = 14 Then T(Index).Text = UCase(T(Index).Text)
    Select Case Index
        Case 0:
            If KHDetail And Left(T(0).Text, 1) = "#" Then
                For i = 7 To 9
                    T(i).Locked = False
                    T(i).TabStop = True
                Next
                ckh.InitKhachHangMaSo 0
                RFocus T(7)
            Else
                If KHDetail Then
                    ckh.InitKhachHangSohieu T(0).Text
                    T(7).Text = ckh.Ten
                    T(8).Text = ckh.DiaChi
                    T(9).Text = ckh.mst
                    If Len(T(1).Text) = 0 And ckh.MaSo > 0 Then T(1).Text = CStr(SelectSQL("SELECT Top 1 KyHieu AS F1 FROM HoaDon WHERE MaKhachHang=" + CStr(ckh.MaSo) + " ORDER BY MaSo DESC"))
                End If
            End If
        Case 4:
            T(Index).Text = Format(T(Index).Text, Mask_2)
        Case 5, 6, 13:
            T(Index).Text = Format(T(Index).Text, Mask_0)
        Case 7, 8, 9, 11, 12:
            If Len(T(Index).Text) = 0 Then T(Index).Text = "..."
    End Select
End Sub

Public Sub GetPhieu(ttdb As Boolean)
'    'If KHDetail Then
'        ckh.InitKhachHangMaSo h.MaKhachHang
'        T(0).Text = ckh.SoHieu
'        T(7).Text = ckh.Ten
'        T(8).Text = ckh.DiaChi
'        T(9).Text = ckh.mst
'    'Else
'    '    T(7).Text = h.TenKH
'    '    T(8).Text = h.DiaChiKH
'    '    T(9).Text = h.MSTKH
'    'End If
 If h.MaKhachHang > 0 Then
        ckh.InitKhachHangMaSo h.MaKhachHang
        T(0).Text = ckh.sohieu
        T(7).Text = ckh.Ten
        T(8).Text = ckh.DiaChi
        T(9).Text = ckh.mst
  Else
      ckh.InitKhachHangMaSo h.MaKhachHang
      T(0).Text = FrmChungtu.txtVT(0).Text
      T(7).Text = FrmChungtu.txtVT(7).Text
      T(8).Text = FrmChungtu.txtVT(8).Text
      T(9).Text = FrmChungtu.txtVT(9).Text
       ckh.sohieu = T(0).Text
       ckh.Ten = T(7).Text
       ckh.DiaChi = T(8).Text
       ckh.mst = T(9).Text
  End If
    
    T(1).Text = h.KyHieu
    T(2).Text = h.soHD
    If h.ThanhTien <> 0 Then T(5).Text = Format(h.ThanhTien, Mask_0)
    T(3).Text = h.MatHang
    T(4).Text = Format(h.SoLuong, Mask_2)
    T(6).Text = CStr(Abs(h.TyLe))
    T(11).Text = h.HTTT
    T(12).Text = FrmChungtu.txtVT(2).Text ' h.MauSo
    T(13).Text = FrmChungtu.txtVT(3).Text ' Format(h.tygia, Mask_0)
    ngay = FrmChungtu.MedNgay(0).Text 'h.NgayPH
    MedNgay.Text = FrmChungtu.MedNgay(0).Text 'h.NgayPH 'FrmChungtu.MedNgay(0).Text ' Format(ngay, Mask_D)
    ChkV(0).Value = h.HD
    ChkV(2).Value = h.HDBL
    ChkV(1).Value = h.KCT
    ChkV(1).Enabled = (h.TyLe = 0)
    ChkV(3).Visible = (h.loai = -1)
    ChkV(4).Visible = (h.loai = -1)
    ChkV(3).Value = h.NK
    ChkV(4).Value = h.ts
    ChkV(5).Value = h.DC
    
    If ttdb Then
        If h.ThanhTien <> 0 Then
            T(10).Text = Format(h.ThanhTien, Mask_0)
         '   T(5).Text = Format(Me.tag, Mask_0)
        End If
    Else
        If h.ThanhTien <> 0 Then T(5).Text = Format(h.ThanhTien, Mask_0)
    End If
    
    Label1(12).Visible = ttdb
    T(10).Visible = ttdb
    'ChkV(0).Enabled = Not ttdb
    'ChkV(2).Enabled = Not ttdb
    
    ChkV(0).Enabled = True
    ChkV(2).Enabled = True
    
    
   If FrmChungtu.cho_hien_vat Then
        Me.Show vbModal   ' cho hien form vat len
'        FrmChungtu.cho_hien_vat = False ' tra ve trang thai khong cho hien
   End If
End Sub


