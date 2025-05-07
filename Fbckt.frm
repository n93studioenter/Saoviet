VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FBcKt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sæ chi tiÕt "
   ClientHeight    =   6855
   ClientLeft      =   3090
   ClientTop       =   3810
   ClientWidth     =   10185
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
   Icon            =   "Fbckt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6855
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Detail Administrative Report"
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   0
      TabIndex        =   238
      Top             =   0
      Width           =   2055
      Begin VB.OptionButton OptLoai 
         BackColor       =   &H0080FF80&
         Caption         =   "Tµi s¶n cè ®Þnh"
         Height          =   495
         HelpContextID   =   800
         Index           =   3
         Left            =   0
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   242
         Tag             =   "Common"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton OptLoai 
         BackColor       =   &H0080FF80&
         Caption         =   "C«ng nî - Doanh thu"
         ForeColor       =   &H00400000&
         Height          =   495
         HelpContextID   =   800
         Index           =   2
         Left            =   0
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   241
         Tag             =   "Common"
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton OptLoai 
         BackColor       =   &H0080FF80&
         Caption         =   "Hµng hãa - C«ng tr×nh"
         ForeColor       =   &H00400000&
         Height          =   495
         HelpContextID   =   800
         Index           =   1
         Left            =   0
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   240
         Tag             =   "Common"
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton OptLoai 
         BackColor       =   &H0080FF80&
         Caption         =   "Sæ tæng hîp"
         ForeColor       =   &H00400000&
         Height          =   495
         HelpContextID   =   800
         Index           =   0
         Left            =   0
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   239
         Tag             =   "Common"
         Top             =   0
         Width           =   2055
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6855
      Left            =   2040
      TabIndex        =   0
      Tag             =   "&Common#&Product and Cost#&Liabilities#&Fixed Assets"
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   2
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tæng hîp"
      TabPicture(0)   =   "Fbckt.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LbTenTk(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LbTenTk(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LbTenTk(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LbTenTk(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LbTenTk(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LbTenTk(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label(8)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line(8)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CboNT"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtShTk(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtShTk(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "OptBC(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "OptBC(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "OptBC(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtShTk(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "OptBC(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "OptBC(5)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "OptBC(6)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "OptBC(7)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "OptBC(9)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdtk(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdtk(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdtk(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdtk(1)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdtk(6)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtShTk(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtShTk(3)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtShCT(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "OptBC(32)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtShCT(1)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "ChkDu(6)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdtk(8)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "OptBC(56)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtShTk(8)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtLS(1)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "OptBC(69)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtLS(0)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "ChkDu(1)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "OptBC(75)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "OptBC(76)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "OptBC(42)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "OptBC(3)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "OptBC(105)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "OptBC(104)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "OptBC(103)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "OptBC(102)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "OptBC(101)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "OptBC(100)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cmdtk(5)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "cmdtk(12)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "ChkDu(0)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtShTk(5)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtShTk(12)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "ChkDu(15)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "OptBC(106)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "CboThang(2)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Picture1"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).ControlCount=   61
      TabCaption(1)   =   "VËt t­, thµnh phÈm"
      TabPicture(1)   =   "Fbckt.frx":57FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CboHD"
      Tab(1).Control(1)=   "txtCT(0)"
      Tab(1).Control(2)=   "txtCT(1)"
      Tab(1).Control(3)=   "cmdCT(1)"
      Tab(1).Control(4)=   "OptBC(77)"
      Tab(1).Control(5)=   "OptBC(74)"
      Tab(1).Control(6)=   "OptBC(43)"
      Tab(1).Control(7)=   "txtShTk(10)"
      Tab(1).Control(8)=   "Frame(5)"
      Tab(1).Control(9)=   "OptBC(54)"
      Tab(1).Control(10)=   "cmdvt(0)"
      Tab(1).Control(11)=   "OptBC(65)"
      Tab(1).Control(12)=   "OptBC(64)"
      Tab(1).Control(13)=   "txtShVT(0)"
      Tab(1).Control(14)=   "OptBC(60)"
      Tab(1).Control(15)=   "cmdCT(0)"
      Tab(1).Control(16)=   "OptBC(59)"
      Tab(1).Control(17)=   "OptBC(52)"
      Tab(1).Control(18)=   "OptBC(51)"
      Tab(1).Control(19)=   "OptBC(50)"
      Tab(1).Control(20)=   "OptBC(46)"
      Tab(1).Control(21)=   "txtShTk(11)"
      Tab(1).Control(22)=   "cmdtk(11)"
      Tab(1).Control(23)=   "OptBC(44)"
      Tab(1).Control(24)=   "ChkDu(7)"
      Tab(1).Control(25)=   "cmdtk(10)"
      Tab(1).Control(26)=   "CboLoai(1)"
      Tab(1).Control(27)=   "ChkDu(4)"
      Tab(1).Control(28)=   "CboVV(0)"
      Tab(1).Control(29)=   "OptBC(33)"
      Tab(1).Control(30)=   "cmdtk(4)"
      Tab(1).Control(31)=   "OptBC(16)"
      Tab(1).Control(32)=   "OptBC(15)"
      Tab(1).Control(33)=   "OptBC(14)"
      Tab(1).Control(34)=   "OptBC(13)"
      Tab(1).Control(35)=   "OptBC(12)"
      Tab(1).Control(36)=   "OptBC(11)"
      Tab(1).Control(37)=   "OptBC(10)"
      Tab(1).Control(38)=   "ChkDu(2)"
      Tab(1).Control(39)=   "CboLoai(0)"
      Tab(1).Control(40)=   "ChkDu(3)"
      Tab(1).Control(41)=   "txtShTk(4)"
      Tab(1).Control(42)=   "Line(6)"
      Tab(1).Control(43)=   "Line(4)"
      Tab(1).Control(44)=   "LbCT(1)"
      Tab(1).Control(45)=   "LbCT(0)"
      Tab(1).Control(46)=   "LbTenTk(11)"
      Tab(1).Control(47)=   "LbTenTk(10)"
      Tab(1).Control(48)=   "LbTenTk(4)"
      Tab(1).Control(49)=   "LbTenVT(0)"
      Tab(1).ControlCount=   50
      TabCaption(2)   =   "Doanh thu - C«ng nî"
      TabPicture(2)   =   "Fbckt.frx":581A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame(4)"
      Tab(2).Control(1)=   "cmdkh(0)"
      Tab(2).Control(2)=   "txtshkh(1)"
      Tab(2).Control(3)=   "txtshkh(0)"
      Tab(2).Control(4)=   "cmdkh(1)"
      Tab(2).Control(5)=   "ChkDu(14)"
      Tab(2).Control(6)=   "OptBC(84)"
      Tab(2).Control(7)=   "OptBC(82)"
      Tab(2).Control(8)=   "OptBC(81)"
      Tab(2).Control(9)=   "OptBC(80)"
      Tab(2).Control(10)=   "OptBC(78)"
      Tab(2).Control(11)=   "MedNgay(0)"
      Tab(2).Control(12)=   "OptBC(70)"
      Tab(2).Control(13)=   "OptBC(53)"
      Tab(2).Control(14)=   "OptBC(67)"
      Tab(2).Control(15)=   "OptBC(66)"
      Tab(2).Control(16)=   "OptBC(63)"
      Tab(2).Control(17)=   "OptBC(62)"
      Tab(2).Control(18)=   "OptBC(58)"
      Tab(2).Control(19)=   "OptBC(57)"
      Tab(2).Control(20)=   "OptBC(49)"
      Tab(2).Control(21)=   "OptBC(120)"
      Tab(2).Control(22)=   "txtShVT(1)"
      Tab(2).Control(23)=   "cmdvt(1)"
      Tab(2).Control(24)=   "OptBC(38)"
      Tab(2).Control(25)=   "OptBC(37)"
      Tab(2).Control(26)=   "OptBC(36)"
      Tab(2).Control(27)=   "Frame(3)"
      Tab(2).Control(28)=   "OptBC(35)"
      Tab(2).Control(29)=   "OptBC(34)"
      Tab(2).Control(30)=   "OptBC(17)"
      Tab(2).Control(31)=   "txtHan"
      Tab(2).Control(32)=   "Line(10)"
      Tab(2).Control(33)=   "Line(9)"
      Tab(2).Control(34)=   "Line(1)"
      Tab(2).Control(35)=   "lbkh(0)"
      Tab(2).Control(36)=   "lbkh(1)"
      Tab(2).Control(37)=   "LbTenVT(1)"
      Tab(2).Control(38)=   "Label(3)"
      Tab(2).ControlCount=   39
      TabCaption(3)   =   "Tµi s¶n cè ®Þnh"
      TabPicture(3)   =   "Fbckt.frx":5836
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Line(11)"
      Tab(3).Control(1)=   "OptBC(18)"
      Tab(3).Control(2)=   "OptBC(19)"
      Tab(3).Control(3)=   "OptBC(20)"
      Tab(3).Control(4)=   "OptBC(21)"
      Tab(3).Control(5)=   "OptBC(22)"
      Tab(3).Control(6)=   "OptBC(23)"
      Tab(3).Control(7)=   "OptBC(24)"
      Tab(3).Control(8)=   "OptBC(25)"
      Tab(3).Control(9)=   "OptBC(26)"
      Tab(3).Control(10)=   "OptBC(27)"
      Tab(3).Control(11)=   "OptBC(28)"
      Tab(3).Control(12)=   "OptBC(29)"
      Tab(3).Control(13)=   "OptBC(30)"
      Tab(3).Control(14)=   "OptBC(31)"
      Tab(3).Control(15)=   "Frame(2)"
      Tab(3).Control(16)=   "OptBC(45)"
      Tab(3).Control(17)=   "OptBC(61)"
      Tab(3).Control(18)=   "OptBC(68)"
      Tab(3).Control(19)=   "OptBC(71)"
      Tab(3).Control(20)=   "OptBC(79)"
      Tab(3).ControlCount=   21
      Begin VB.PictureBox Picture1 
         Height          =   15
         Left            =   4860
         ScaleHeight     =   15
         ScaleWidth      =   135
         TabIndex        =   233
         Top             =   1620
         Width           =   135
      End
      Begin VB.ComboBox CboThang 
         Height          =   315
         Index           =   2
         ItemData        =   "Fbckt.frx":5852
         Left            =   3300
         List            =   "Fbckt.frx":5868
         Style           =   2  'Dropdown List
         TabIndex        =   232
         Top             =   1860
         Width           =   735
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ quü tiÒn mÆt"
         Height          =   255
         Index           =   106
         Left            =   180
         TabIndex        =   231
         Tag             =   "Account Balance Report 2"
         Top             =   1260
         Width           =   2775
      End
      Begin VB.CheckBox ChkDu 
         Caption         =   "Theo ®èi øng"
         Height          =   255
         Index           =   15
         Left            =   3780
         TabIndex        =   230
         Tag             =   "by Ref. Acc."
         ToolTipText     =   "by Ref. Acc."
         Top             =   900
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   12
         Left            =   4500
         LinkItem        =   "Sè hiÖu tµi kho¶n ®èi øng"
         MaxLength       =   12
         TabIndex        =   229
         Tag             =   "0"
         Top             =   2700
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   5
         Left            =   5700
         LinkItem        =   "Sè hiÖu tµi kho¶n cÇn xem"
         MaxLength       =   12
         TabIndex        =   228
         Tag             =   "0"
         Top             =   540
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox ChkDu 
         Caption         =   "Nhãm theo ®èi øng"
         Height          =   255
         Index           =   0
         Left            =   3780
         TabIndex        =   227
         Tag             =   "Group by Reference Account"
         Top             =   540
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   12
         Left            =   7200
         Picture         =   "Fbckt.frx":587E
         Style           =   1  'Graphical
         TabIndex        =   226
         Top             =   2820
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   5
         Left            =   6540
         Picture         =   "Fbckt.frx":5CF8
         Style           =   1  'Graphical
         TabIndex        =   225
         Top             =   540
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ nhËt ký chung"
         Height          =   255
         Index           =   100
         Left            =   180
         TabIndex        =   224
         Tag             =   "Journal Ledger"
         Top             =   180
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ c¸i tµi kho¶n"
         Height          =   255
         Index           =   101
         Left            =   180
         TabIndex        =   223
         Tag             =   "Ledger"
         Top             =   540
         Width           =   1575
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ c¸i tµi kho¶n (chøng tõ ghi sæ)"
         Height          =   255
         Index           =   102
         Left            =   180
         TabIndex        =   222
         Tag             =   "Ledger 2 (group by V. Book)"
         Top             =   3420
         Width           =   2775
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ ®¨ng ký chøng tõ ghi sæ"
         Height          =   255
         Index           =   103
         Left            =   180
         TabIndex        =   221
         Tag             =   "Voucher Book Register"
         Top             =   3780
         UseMaskColor    =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chøng tõ ghi sæ sè"
         Height          =   255
         Index           =   104
         Left            =   180
         TabIndex        =   220
         Tag             =   "Voucher Book"
         Top             =   4140
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ tæng hîp chøng tõ ghi sæ"
         Height          =   255
         Index           =   105
         Left            =   180
         TabIndex        =   219
         Tag             =   "Summary Voucher Book"
         Top             =   4500
         Width           =   2775
      End
      Begin VB.Frame Frame 
         Height          =   3015
         Index           =   4
         Left            =   -71280
         TabIndex        =   170
         Top             =   3480
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CheckBox ChkDu 
            Caption         =   "Doanh thu - Gi¸ vèn sau chiÕt khÊu"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   198
            Tag             =   "Salesman"
            Top             =   4440
            Width           =   3495
         End
         Begin VB.TextBox txtShTk 
            Height          =   285
            Index           =   9
            Left            =   1560
            LinkItem        =   "Sè hiÖu chi tiÕt cÇn xem"
            MaxLength       =   12
            TabIndex        =   186
            Tag             =   "0"
            Top             =   1965
            Width           =   1335
         End
         Begin VB.CheckBox ChkDu 
            Caption         =   "Tµi kho¶n"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   184
            Tag             =   "Salesman"
            Top             =   1920
            Width           =   1335
         End
         Begin VB.CommandButton cmdtk 
            Height          =   375
            Index           =   9
            Left            =   3000
            Picture         =   "Fbckt.frx":6172
            Style           =   1  'Graphical
            TabIndex        =   187
            Top             =   1920
            Width           =   375
         End
         Begin VB.ComboBox CboTT 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   197
            ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
            Top             =   4080
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox CboTT 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   195
            ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
            Top             =   3720
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox CboTT 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   193
            ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
            Top             =   2640
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CheckBox ChkTT 
            Caption         =   "Th«ng tin 3"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   196
            Tag             =   "Customer"
            Top             =   4080
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox ChkTT 
            Caption         =   "Th«ng tin 2"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   194
            Tag             =   "Customer"
            Top             =   3720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox ChkTT 
            Caption         =   "Th«ng tin 1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   192
            Tag             =   "Customer"
            Top             =   2640
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox CboVV 
            Height          =   315
            Index           =   1
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   191
            Top             =   2640
            Width           =   2175
         End
         Begin VB.OptionButton OptDT 
            Caption         =   "Theo ®èi t­îng"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   190
            Tag             =   "VAT Rate"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox txtsh 
            Height          =   300
            Index           =   0
            Left            =   1560
            LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
            MaxLength       =   20
            TabIndex        =   182
            Tag             =   "0"
            Top             =   1605
            Width           =   1335
         End
         Begin VB.OptionButton OptDT 
            Caption         =   "Kh¸ch hµng"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   176
            Tag             =   "Customer"
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox ChkDu 
            Caption         =   "NV b¸n hµng"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   181
            Tag             =   "Salesman"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CheckBox ChkDu 
            Caption         =   "Kªnh ph©n phèi"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   180
            Tag             =   "Source"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox ChkDu 
            Caption         =   "Kho hµng"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   177
            Tag             =   "Store"
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox ChkDu 
            Caption         =   "Kh¸ch hµng"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   172
            Tag             =   "Customer"
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox CboKho 
            Height          =   315
            Index           =   2
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   179
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CommandButton cmd 
            Height          =   375
            Index           =   0
            Left            =   3000
            Picture         =   "Fbckt.frx":65EC
            Style           =   1  'Graphical
            TabIndex        =   183
            Tag             =   "0"
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton cmdkh 
            Height          =   375
            Index           =   2
            Left            =   3000
            Picture         =   "Fbckt.frx":6A66
            Style           =   1  'Graphical
            TabIndex        =   174
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtshkh 
            Height          =   285
            Index           =   2
            Left            =   1560
            LinkItem        =   "Sè hiÖu chi tiÕt cÇn xem"
            MaxLength       =   12
            TabIndex        =   173
            Tag             =   "0"
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox CboKho 
            Height          =   315
            Index           =   1
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   178
            Top             =   840
            Width           =   2295
         End
         Begin VB.OptionButton OptDT 
            Caption         =   "ThuÕ VAT"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   189
            Tag             =   "VAT Rate"
            Top             =   2280
            Width           =   1095
         End
         Begin VB.OptionButton OptDT 
            Caption         =   "MÆt hµng"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   171
            Tag             =   "Product"
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label LbTenTk 
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   9
            Left            =   2520
            TabIndex        =   188
            Tag             =   "1"
            Top             =   2640
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lb 
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   185
            Tag             =   "1"
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lbkh 
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   4920
            TabIndex        =   175
            Tag             =   "1"
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.ComboBox CboHD 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -69540
         Style           =   2  'Dropdown List
         TabIndex        =   161
         ToolTipText     =   "NhÊn chuét ph¶i ®Ó ®¨ng ký"
         Top             =   5460
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdkh 
         Height          =   375
         Index           =   0
         Left            =   -71880
         Picture         =   "Fbckt.frx":6EE0
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtshkh 
         Height          =   300
         Index           =   1
         Left            =   -72780
         LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
         MaxLength       =   20
         TabIndex        =   140
         Tag             =   "0"
         Top             =   2220
         Width           =   855
      End
      Begin VB.TextBox txtshkh 
         Height          =   285
         Index           =   0
         Left            =   -72900
         LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
         MaxLength       =   20
         TabIndex        =   124
         Tag             =   "0"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdkh 
         Height          =   375
         Index           =   1
         Left            =   -71880
         Picture         =   "Fbckt.frx":735A
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtCT 
         Height          =   285
         Index           =   0
         Left            =   -71520
         LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
         MaxLength       =   20
         TabIndex        =   30
         Tag             =   "0"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtCT 
         Height          =   300
         Index           =   1
         Left            =   -71340
         LinkItem        =   "Sè hiÖu tµi kho¶n ®èi øng"
         MaxLength       =   12
         TabIndex        =   151
         Tag             =   "0"
         Top             =   5460
         Width           =   1335
      End
      Begin VB.CommandButton cmdCT 
         Height          =   375
         Index           =   1
         Left            =   -70020
         Picture         =   "Fbckt.frx":77D4
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   5445
         Width           =   375
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng tæng hîp sè d­ chi tiÕt"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   41
         Tag             =   "Detail Account Balance Report"
         Top             =   60
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª chi tiÕt ph¸t sinh"
         Height          =   255
         Index           =   42
         Left            =   5280
         TabIndex        =   85
         Tag             =   "Account Detail Group by Ref. Acc."
         Top             =   60
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox ChkDu 
         Caption         =   "Theo ngo¹i tÖ"
         Height          =   285
         Index           =   14
         Left            =   -70560
         TabIndex        =   63
         Tag             =   "View foreign currency"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Tæng hîp sè d­ c«ng nî theo tµi kho¶n"
         Height          =   255
         Index           =   84
         Left            =   -70320
         TabIndex        =   105
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª tÝch sè"
         Height          =   255
         Index           =   83
         Left            =   -74760
         TabIndex        =   104
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Vni 14 AlexBrush"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67080
         TabIndex        =   160
         Top             =   6840
         Width           =   495
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶n ®èi chiÕu c«ng nî theo ho¸ ®¬n"
         Height          =   255
         Index           =   82
         Left            =   -74820
         TabIndex        =   146
         Tag             =   "Liability Detail"
         Top             =   1860
         Width           =   3375
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Tæng hîp hµng khuyÕn m¹i"
         Height          =   255
         Index           =   81
         Left            =   -74820
         TabIndex        =   199
         Top             =   5715
         Width           =   2655
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi tiÕt hµng khuyÕn m¹i"
         Height          =   255
         Index           =   80
         Left            =   -74820
         TabIndex        =   128
         Top             =   6060
         Width           =   2295
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng ph©n bæ khÊu hao"
         Height          =   255
         Index           =   79
         Left            =   -71460
         TabIndex        =   145
         Top             =   3300
         Width           =   2295
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ chi tiÕt sè 3"
         Height          =   255
         Index           =   78
         Left            =   -70320
         TabIndex        =   54
         Top             =   1680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª nhËp xuÊt th¼ng"
         Height          =   255
         Index           =   77
         Left            =   -74700
         TabIndex        =   169
         Tag             =   "Export table by product"
         Top             =   5940
         Width           =   2175
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng c©n ®èi sè ph¸t sinh luü kÕ "
         Height          =   255
         Index           =   76
         Left            =   180
         TabIndex        =   168
         Tag             =   "Account Balance Report 3"
         Top             =   2340
         Width           =   3735
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Tæng hîp ph¸t sinh ch÷ T"
         Height          =   255
         Index           =   75
         Left            =   4500
         TabIndex        =   122
         Top             =   4500
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª s¶n l­îng thµnh phÈm hoµn thµnh"
         Height          =   255
         Index           =   74
         Left            =   -70920
         TabIndex        =   127
         Tag             =   "Inventory"
         Top             =   2520
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Tæng hîp sè d­ c«ng nî theo tuæi nî"
         Height          =   255
         Index           =   72
         Left            =   -71760
         TabIndex        =   167
         Top             =   6960
         Width           =   3375
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi tiÕt c«ng nî theo tuæi nî"
         Height          =   255
         Index           =   73
         Left            =   -69720
         TabIndex        =   103
         Top             =   2160
         Width           =   2655
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "KiÓm kª tµi s¶n cè ®Þnh"
         Height          =   255
         Index           =   71
         Left            =   -71460
         TabIndex        =   143
         Top             =   2940
         Width           =   2295
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¸o c¸o kiÓm kª"
         Height          =   255
         Index           =   43
         Left            =   -69600
         TabIndex        =   86
         Tag             =   "Inventory"
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSMask.MaskEdBox MedNgay 
         Height          =   315
         Index           =   0
         Left            =   -70140
         TabIndex        =   23
         Top             =   2520
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "In toµn bé sæ ct ®èi t­îng thanh to¸n"
         Height          =   255
         Index           =   70
         Left            =   -70425
         TabIndex        =   102
         Tag             =   "Print all Detail Report of Payable and Receivable Form 2"
         Top             =   2160
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CheckBox ChkDu 
         Caption         =   "Theo tµi kho¶n"
         Height          =   255
         Index           =   1
         Left            =   2940
         TabIndex        =   55
         Tag             =   "By Account"
         Top             =   3420
         Width           =   1455
      End
      Begin VB.TextBox txtShTk 
         Height          =   300
         Index           =   10
         Left            =   -70260
         LinkItem        =   "Sè hiÖu tµi kho¶n ®èi øng"
         MaxLength       =   12
         TabIndex        =   113
         Tag             =   "0"
         Top             =   4140
         Width           =   975
      End
      Begin VB.TextBox txtLS 
         Height          =   285
         Index           =   0
         Left            =   7500
         LinkItem        =   "B¸o c¸o sÏ liÖt kª c¸c kho¶n ph¶i thu cña kh¸ch hµng qu¸ thêi h¹n nµy"
         MaxLength       =   5
         TabIndex        =   119
         Text            =   "0"
         Top             =   2820
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª tÝch sè theo l·i suÊt (% th¸ng):                    L·i d­ nî"
         Height          =   255
         Index           =   69
         Left            =   5160
         TabIndex        =   101
         Tag             =   "Detail Interest Table                           Debit Interest Rate"
         Top             =   60
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtLS 
         Height          =   285
         Index           =   1
         Left            =   8460
         LinkItem        =   "B¸o c¸o sÏ liÖt kª c¸c kho¶n ph¶i thu cña kh¸ch hµng qu¸ thêi h¹n nµy"
         MaxLength       =   5
         TabIndex        =   121
         Text            =   "0"
         Top             =   60
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ tµi s¶n cè ®Þnh 2"
         Height          =   255
         Index           =   68
         Left            =   -74820
         TabIndex        =   100
         Tag             =   "Fixed Asset Table 2"
         Top             =   1380
         Width           =   2295
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   5
         Left            =   -72600
         TabIndex        =   37
         Top             =   60
         Width           =   4335
         Begin VB.OptionButton OptKho 
            Caption         =   "Tæng c¸c kho"
            Height          =   375
            Index           =   1
            Left            =   2760
            TabIndex        =   40
            Tag             =   "All store"
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton OptKho 
            Caption         =   "Kho"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   38
            Tag             =   "Store"
            Top             =   120
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.ComboBox CboKho 
            Height          =   315
            Index           =   0
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "§Þnh møc thµnh phÈm"
         Height          =   255
         Index           =   54
         Left            =   -64620
         TabIndex        =   164
         Tag             =   "Production Norm Table"
         Top             =   6420
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi tiÕt doanh thu b¸n hµng "
         Height          =   255
         Index           =   53
         Left            =   -74820
         TabIndex        =   92
         Tag             =   "Detail Turnover 2"
         Top             =   3660
         Width           =   2415
      End
      Begin VB.CommandButton cmdvt 
         Height          =   375
         Index           =   0
         Left            =   -70200
         Picture         =   "Fbckt.frx":7C4E
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   680
         Width           =   375
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Danh s¸ch nî qu¸ h¹n møc"
         Height          =   255
         Index           =   67
         Left            =   -70320
         TabIndex        =   53
         Tag             =   "Over Maximum Credit Table"
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ chi tiÕt b¸n hµng"
         Height          =   255
         Index           =   66
         Left            =   -74820
         TabIndex        =   99
         Tag             =   "Sales Invoice List"
         Top             =   4020
         Width           =   2055
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng tæng hîp chi tiÕt vËt liÖu, sp, hh"
         Height          =   255
         Index           =   65
         Left            =   -74700
         TabIndex        =   98
         Tag             =   "Material and Product Summary Report"
         Top             =   4860
         Width           =   3135
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "ThÎ kho"
         Height          =   255
         Index           =   64
         Left            =   -72600
         TabIndex        =   97
         Tag             =   "Form 2"
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   8
         Left            =   2100
         LinkItem        =   "Sè hiÖu chi tiÕt cÇn xem"
         MaxLength       =   12
         TabIndex        =   111
         Tag             =   "0"
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox txtShVT 
         Height          =   285
         Index           =   0
         Left            =   -71520
         LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
         MaxLength       =   20
         TabIndex        =   106
         Tag             =   "0"
         Top             =   680
         Width           =   1335
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "In toµn bé biªn b¶n ®èi chiÕu c«ng nî"
         Height          =   255
         Index           =   63
         Left            =   -70380
         TabIndex        =   96
         Tag             =   "Print all Liability Details"
         Top             =   1920
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "In toµn bé sæ chi tiÕt c«ng nî"
         Height          =   255
         Index           =   62
         Left            =   -70500
         TabIndex        =   95
         Tag             =   "Print all Detail Report of Payable and Receivable"
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ khÊu hao tæng hîp"
         Height          =   255
         Index           =   61
         Left            =   -74820
         TabIndex        =   52
         Tag             =   "Depreciation Summary"
         Top             =   660
         Width           =   2175
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª xuÊt theo s¶n phÈm"
         Height          =   255
         Index           =   60
         Left            =   -74700
         TabIndex        =   94
         Tag             =   "Export table by product"
         Top             =   5220
         Width           =   2415
      End
      Begin VB.CommandButton cmdCT 
         Height          =   375
         Index           =   0
         Left            =   -70320
         Picture         =   "Fbckt.frx":80C8
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2160
         Width           =   375
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi tiÕt chi phÝ theo c«ng tr×nh"
         Height          =   255
         Index           =   59
         Left            =   -74700
         TabIndex        =   51
         Tag             =   "Detail expense in product"
         Top             =   1860
         Width           =   2535
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng c©n ®èi c«ng nî"
         Height          =   255
         Index           =   58
         Left            =   -74760
         TabIndex        =   166
         Tag             =   "Liabities Balance sheet"
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ tæng hîp c«ng nî theo kh¸ch hµng"
         Height          =   255
         Index           =   57
         Left            =   -70320
         TabIndex        =   50
         Tag             =   "Detail Report of Payable and Receivable 2"
         Top             =   1200
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Tæng hîp ph¸t sinh theo nhãm"
         Height          =   255
         Index           =   56
         Left            =   4500
         TabIndex        =   93
         Tag             =   "Detail Report of Sub-Account"
         Top             =   4260
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng tæng hîp tån kho"
         Height          =   255
         Index           =   55
         Left            =   -74760
         TabIndex        =   165
         Tag             =   "Stock table "
         Top             =   6240
         Width           =   2175
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª l­u chuyÓn néi bé"
         Height          =   255
         Index           =   52
         Left            =   -74700
         TabIndex        =   91
         Tag             =   "InternaI flows Report"
         Top             =   5580
         Width           =   2535
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi tiÕt gi¸ thµnh thµnh phÈm"
         Height          =   255
         Index           =   51
         Left            =   -70140
         TabIndex        =   22
         Tag             =   "Cost of Product Detail Report"
         Top             =   6060
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi phÝ vµ kÕt qu¶ theo thµnh phÈm hoµn thµnh"
         Height          =   255
         Index           =   50
         Left            =   -70200
         TabIndex        =   49
         Tag             =   "Cost and Profit of Finished Goods"
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶n ®èi chiÕu c«ng nî"
         Height          =   255
         Index           =   49
         Left            =   -74820
         TabIndex        =   90
         Tag             =   "Liability Detail"
         Top             =   2220
         Width           =   2655
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª tån kho trªn møc tèi ®a"
         Height          =   255
         Index           =   48
         Left            =   -70680
         TabIndex        =   163
         Tag             =   "Over maximum quantity table "
         Top             =   6600
         Width           =   2655
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª tån kho d­íi møc tèi thiÓu"
         Height          =   255
         Index           =   47
         Left            =   -70680
         TabIndex        =   162
         Tag             =   "Under minimum quantity"
         Top             =   6960
         Width           =   2895
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª nhËp xuÊt tån theo nguån"
         Height          =   255
         Index           =   46
         Left            =   -74700
         TabIndex        =   89
         Tag             =   "Fluctuation following resource"
         Top             =   4140
         Width           =   2775
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ tµi s¶n cè ®Þnh 1"
         Height          =   255
         Index           =   45
         Left            =   -74820
         TabIndex        =   88
         Tag             =   "Fixed Asset Table 1"
         Top             =   1020
         Width           =   2295
      End
      Begin VB.TextBox txtShTk 
         Height          =   300
         Index           =   11
         Left            =   -71340
         LinkItem        =   "Sè hiÖu tµi kho¶n ®èi øng"
         MaxLength       =   12
         TabIndex        =   148
         Tag             =   "0"
         Top             =   5070
         Width           =   1335
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   11
         Left            =   -70020
         Picture         =   "Fbckt.frx":8542
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   5040
         Width           =   375
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª xuÊt theo tµi kho¶n "
         Height          =   255
         Index           =   44
         Left            =   -74700
         TabIndex        =   87
         Tag             =   "Export table by account "
         Top             =   4500
         Width           =   2415
      End
      Begin VB.CheckBox ChkDu 
         Caption         =   "§èi øng"
         Height          =   255
         Index           =   7
         Left            =   -71340
         TabIndex        =   112
         Tag             =   "by Ref. Acc."
         Top             =   4140
         Width           =   1215
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   10
         Left            =   -69180
         Picture         =   "Fbckt.frx":89BC
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   4140
         Width           =   375
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   8
         Left            =   3180
         Picture         =   "Fbckt.frx":8E36
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   540
         Width           =   375
      End
      Begin VB.CheckBox ChkDu 
         Caption         =   "In chi tiÕt kh«ng cã sè liÖu"
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   4500
         TabIndex        =   154
         Tag             =   "Print all account"
         Top             =   5820
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi tiÕt ph¶i thu theo ho¸ ®¬n"
         Height          =   255
         Index           =   120
         Left            =   -74820
         TabIndex        =   123
         Tag             =   "Detail Turnover"
         Top             =   3300
         Width           =   2535
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª theo hîp ®ång ®Çu ra"
         Height          =   255
         Index           =   41
         Left            =   -71760
         TabIndex        =   159
         Tag             =   "List of output contract"
         Top             =   6600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª theo hîp ®ång ®Çu vµo"
         Height          =   255
         Index           =   40
         Left            =   -71760
         TabIndex        =   158
         Tag             =   "List of input contract"
         Top             =   6240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox ChkDu 
         Caption         =   "Theo ngo¹i tÖ"
         Height          =   255
         Index           =   5
         Left            =   -71400
         TabIndex        =   59
         Tag             =   "View foreign currency"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtShVT 
         Height          =   285
         Index           =   1
         Left            =   -72780
         LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
         MaxLength       =   20
         TabIndex        =   132
         Tag             =   "0"
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdvt 
         Height          =   375
         Index           =   1
         Left            =   -71940
         Picture         =   "Fbckt.frx":92B0
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   3960
         Width           =   375
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi tiÕt Doanh thu - Gi¸ vèn"
         Height          =   255
         Index           =   38
         Left            =   -74820
         TabIndex        =   84
         Tag             =   "Detail Sale and cost of good sold report"
         Top             =   4980
         Width           =   2415
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Tæng hîp Doanh thu - Gi¸ vèn"
         Height          =   255
         Index           =   37
         Left            =   -74820
         TabIndex        =   157
         Tag             =   "Summary Sale and Cost of good sold report"
         Top             =   4620
         Width           =   2775
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ chi tiÕt c«ng nî "
         Height          =   255
         Index           =   36
         Left            =   -74760
         TabIndex        =   48
         Tag             =   "Detail Report of Payable and Receivable"
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Frame Frame 
         Height          =   495
         Index           =   3
         Left            =   -71400
         TabIndex        =   12
         Top             =   555
         Width           =   1575
         Begin VB.TextBox txtShTk 
            Height          =   285
            Index           =   7
            Left            =   80
            LinkItem        =   "Sè hiÖu chi tiÕt cÇn xem"
            MaxLength       =   12
            TabIndex        =   15
            Tag             =   "0"
            Top             =   150
            Width           =   1095
         End
         Begin VB.CommandButton cmdtk 
            Height          =   375
            Index           =   7
            Left            =   1200
            Picture         =   "Fbckt.frx":972A
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Tµi kho¶n"
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   17
            Tag             =   "Account"
            Top             =   240
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Ph©n lo¹i"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   16
            Tag             =   "Class"
            ToolTipText     =   "by Group"
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox CboKH 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   180
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label LbTenTk 
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   7
            Left            =   3120
            TabIndex        =   13
            Tag             =   "1"
            Top             =   165
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ tæng hîp c«ng nî theo nhãm tµi kho¶n"
         Height          =   255
         Index           =   35
         Left            =   -74760
         TabIndex        =   83
         Tag             =   "Liabilities balance"
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox CboLoai 
         Height          =   315
         Index           =   1
         Left            =   -70260
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   3420
         Width           =   2295
      End
      Begin VB.CheckBox ChkDu 
         Caption         =   "Nguån"
         Height          =   255
         Index           =   4
         Left            =   -71355
         TabIndex        =   58
         Tag             =   "Resource"
         Top             =   3420
         Width           =   1095
      End
      Begin VB.TextBox txtShCT 
         Height          =   285
         Index           =   1
         Left            =   5220
         LinkItem        =   "Sè hiÖu tµi kho¶n cÇn in toµn bé chi tiÕt"
         MaxLength       =   12
         TabIndex        =   36
         Tag             =   "0"
         Top             =   2460
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame 
         Height          =   855
         Index           =   2
         Left            =   -71520
         TabIndex        =   1
         Top             =   3720
         Width           =   3735
         Begin VB.ComboBox Cbo 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   435
            Width           =   1935
         End
         Begin VB.TextBox txtShTS 
            Height          =   300
            Left            =   1560
            LinkItem        =   "Sè hiÖu chi tiÕt cÇn xem"
            MaxLength       =   12
            TabIndex        =   3
            Tag             =   "0"
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton Opt 
            Caption         =   "§¬n vÞ qu¶n lý"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Tag             =   "Managed by"
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Tµi s¶n"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Tag             =   "Sort by Code"
            Top             =   165
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Tæng hîp doanh thu theo"
         Height          =   255
         Index           =   34
         Left            =   -74820
         TabIndex        =   156
         Tag             =   "Summary Turnover"
         Top             =   5340
         Width           =   2295
      End
      Begin VB.ComboBox CboVV 
         Height          =   315
         Index           =   0
         Left            =   -71340
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1500
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Ph¸t sinh theo vô viÖc"
         Height          =   255
         Index           =   33
         Left            =   -70800
         TabIndex        =   8
         Tag             =   "Invoices Recorded by Index"
         Top             =   2520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "HÖ thèng tµi kho¶n"
         Height          =   255
         Index           =   32
         Left            =   180
         TabIndex        =   155
         Tag             =   "Chart of Account"
         Top             =   2700
         Width           =   2775
      End
      Begin VB.TextBox txtShCT 
         Height          =   285
         Index           =   0
         Left            =   5580
         LinkItem        =   "Sè hiÖu tµi kho¶n cÇn in toµn bé chi tiÕt"
         MaxLength       =   12
         TabIndex        =   33
         Tag             =   "0"
         Top             =   2820
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   3
         Left            =   2820
         LinkItem        =   "Sè hiÖu chi tiÕt cÇn xem toµn bé c¸c chi tiÕt nguyªn tÖ"
         MaxLength       =   12
         TabIndex        =   136
         Tag             =   "0"
         Top             =   5700
         Width           =   1335
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   2
         Left            =   2820
         LinkItem        =   "Sè hiÖu chi tiÕt nguyªn tÖ cÇn xem"
         MaxLength       =   12
         TabIndex        =   129
         Tag             =   "0"
         Top             =   5340
         Width           =   1335
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   6
         Left            =   6540
         Picture         =   "Fbckt.frx":9BA4
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   2700
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   1
         Left            =   6540
         Picture         =   "Fbckt.frx":A01E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   900
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   3
         Left            =   4140
         Picture         =   "Fbckt.frx":A498
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   5700
         Width           =   375
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   2
         Left            =   4140
         Picture         =   "Fbckt.frx":A912
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   5340
         Width           =   375
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   0
         Left            =   3180
         Picture         =   "Fbckt.frx":AD8C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   900
         Width           =   375
      End
      Begin VB.CommandButton cmdtk 
         Height          =   375
         Index           =   4
         Left            =   -69180
         Picture         =   "Fbckt.frx":B206
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   3780
         Width           =   375
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B×nh qu©n tµi s¶n cÇn tÝnh khÊu hao"
         Height          =   255
         Index           =   31
         Left            =   -71460
         TabIndex        =   82
         Tag             =   "Avarage of Depriciated Asset"
         Top             =   4500
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi tiÕt tµi s¶n kh«ng ph¶i tÝnh khÊu hao"
         Height          =   255
         Index           =   30
         Left            =   -71460
         TabIndex        =   81
         Tag             =   "Detail of Asset not Depriciation "
         Top             =   4140
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Tæng hîp biÕn ®éng luü kÕ"
         Height          =   255
         Index           =   29
         Left            =   -71460
         TabIndex        =   80
         Tag             =   "Accured Fluctuation of Fixed Assets"
         Top             =   2340
         Width           =   2415
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Tæng hîp biÕn ®éng trong kú"
         Height          =   255
         Index           =   28
         Left            =   -71460
         TabIndex        =   79
         Tag             =   "Fluctuation of Fixed Asset "
         Top             =   1980
         Width           =   2535
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "N¨ng lùc tµi s¶n"
         Height          =   255
         Index           =   27
         Left            =   -71460
         TabIndex        =   47
         Tag             =   "Liability of fixed asset"
         Top             =   4860
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Tæng hîp ph©n bæ khÊu hao"
         Height          =   255
         Index           =   26
         Left            =   -71460
         TabIndex        =   46
         Tag             =   "Summary Apportion of Depreciation"
         Top             =   660
         Width           =   2775
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi tiÕt ph©n bæ khÊu hao"
         Height          =   255
         Index           =   25
         Left            =   -71460
         TabIndex        =   45
         Tag             =   "Detail Apportion of Depreciation"
         Top             =   300
         Width           =   2415
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "§¸nh gi¸ l¹i gi¶m"
         Height          =   255
         Index           =   24
         Left            =   -74820
         TabIndex        =   78
         Tag             =   "Decreased by Revaluation"
         Top             =   3300
         Width           =   2895
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Gi¶m tµi s¶n"
         Height          =   255
         Index           =   23
         Left            =   -74820
         TabIndex        =   77
         Tag             =   "Decrease Fixed Asset Report"
         Top             =   2940
         Width           =   2895
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "§¸nh gi¸ l¹i t¨ng"
         Height          =   255
         Index           =   22
         Left            =   -74820
         TabIndex        =   76
         Tag             =   "Increased by Revaluation "
         Top             =   2340
         Width           =   2895
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "T¨ng tµi s¶n"
         Height          =   255
         Index           =   21
         Left            =   -74820
         TabIndex        =   75
         Tag             =   "Increasing Fixed Asset Report"
         Top             =   1980
         Width           =   2895
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sè d­ cuèi kú"
         Height          =   255
         Index           =   20
         Left            =   -71460
         TabIndex        =   74
         Tag             =   "Closing Balance"
         Top             =   1380
         Width           =   2895
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sè d­ ®Çu kú"
         Height          =   255
         Index           =   19
         Left            =   -71460
         TabIndex        =   44
         Tag             =   "Opening Balance "
         Top             =   1020
         Width           =   2895
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ khÊu hao chi tiÕt"
         Height          =   255
         Index           =   18
         Left            =   -74820
         TabIndex        =   43
         Tag             =   "Depriciation Detail"
         Top             =   300
         Width           =   2895
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi tiÕt nî ph¶i thu víi thêi h¹n"
         Height          =   255
         Index           =   17
         Left            =   -74820
         TabIndex        =   42
         Tag             =   "Time of liabilities"
         Top             =   2580
         Width           =   2535
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¸o c¸o xuÊt theo kªnh ph©n phèi"
         Height          =   255
         Index           =   16
         Left            =   -74700
         TabIndex        =   73
         Tag             =   "Export table by customer"
         Top             =   3780
         Width           =   3015
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¸o c¸o nhËp theo kªnh ph©n phèi"
         Height          =   255
         Index           =   15
         Left            =   -74700
         TabIndex        =   72
         Tag             =   "Import table by resource"
         Top             =   3420
         Width           =   3255
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng kª nhËp xuÊt tån"
         Height          =   375
         Index           =   14
         Left            =   -74700
         TabIndex        =   71
         Tag             =   "Fluctuation of inventories"
         Top             =   180
         Width           =   1935
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "In toµn bé sæ chi tiÕt vËt t­ hoÆc thÎ kho"
         Height          =   255
         Index           =   13
         Left            =   -70140
         TabIndex        =   70
         Tag             =   "Print all Inventory detail report"
         Top             =   5820
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ chi tiÕt vËt t­"
         Height          =   255
         Index           =   12
         Left            =   -74700
         TabIndex        =   69
         Tag             =   "Inventory detail report"
         Top             =   660
         Width           =   1575
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Kinh doanh theo c«ng tr×nh, s¶n phÈm"
         Height          =   255
         Index           =   11
         Left            =   -74700
         TabIndex        =   20
         Tag             =   "Profits on constructions and products"
         Top             =   2220
         Width           =   3735
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Chi phÝ s¶n xuÊt vµ gi¸ thµnh s¶n phÈm"
         Height          =   255
         Index           =   10
         Left            =   -74700
         TabIndex        =   7
         Tag             =   "Expense and cost of products"
         Top             =   1500
         Width           =   3735
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng c©n ®èi sè ph¸t sinh, in ®Õn cÊp "
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   68
         Tag             =   "Account Balance Report 2"
         Top             =   1980
         Width           =   3135
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "B¶ng tæng hîp sè d­ chi tiÕt nguyªn tÖ"
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   67
         Tag             =   "Balance of foreign currency report"
         Top             =   6060
         Width           =   3135
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "In toµn bé c¸c sæ chi tiÕt nguyªn tÖ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   1051.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   23490
         Index           =   6
         Left            =   -180
         MaskColor       =   &H00FFC0C0&
         TabIndex        =   66
         Tag             =   "Print all Foreign currency detail report"
         Top             =   7020
         Width           =   9285
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "In chi tiÕt nguyªn tÖ tµi kho¶n"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   65
         Tag             =   "Print all detail of account"
         Top             =   5700
         Width           =   2415
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ chi tiÕt nguyªn tÖ"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   64
         Tag             =   "Foreign currency detail report"
         Top             =   5340
         Width           =   1815
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   6
         Left            =   7560
         LinkItem        =   "Sè hiÖu tµi kho¶n ®èi øng"
         MaxLength       =   12
         TabIndex        =   60
         Tag             =   "0"
         Top             =   2940
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "In toµn bé c¸c sæ chi tiÕt tõ"
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   34
         Tag             =   "Print all Detail Report from"
         Top             =   60
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "In toµn bé chi tiÕt tµi kho¶n"
         Height          =   255
         Index           =   1
         Left            =   5400
         TabIndex        =   26
         Tag             =   "Print all Detail Report of Account"
         Top             =   60
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton OptBC 
         Caption         =   "Sæ chi tiÕt tµi kho¶n"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Tag             =   "Account Detail Report"
         Top             =   900
         Width           =   1935
      End
      Begin VB.TextBox txtHan 
         Height          =   285
         Left            =   -72060
         LinkItem        =   "B¸o c¸o sÏ liÖt kª c¸c kho¶n ph¶i thu cña kh¸ch hµng qu¸ thêi h¹n nµy"
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "0"
         Top             =   2580
         Width           =   495
      End
      Begin VB.CheckBox ChkDu 
         Caption         =   "Tµi kho¶n"
         Height          =   255
         Index           =   2
         Left            =   -71355
         TabIndex        =   56
         Tag             =   "Account"
         Top             =   3780
         Width           =   1095
      End
      Begin VB.ComboBox CboLoai 
         Height          =   315
         Index           =   0
         Left            =   -70215
         Style           =   2  'Dropdown List
         TabIndex        =   144
         Top             =   4605
         Width           =   2295
      End
      Begin VB.CheckBox ChkDu 
         Caption         =   "Lo¹i vËt t­"
         Height          =   255
         Index           =   3
         Left            =   -71355
         TabIndex        =   57
         Tag             =   "Class"
         Top             =   4665
         Width           =   1095
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   4
         Left            =   -70260
         LinkItem        =   "Sè hiÖu tµi kho¶n cÇn theo dâi"
         MaxLength       =   12
         TabIndex        =   138
         Tag             =   "0"
         Top             =   3780
         Width           =   975
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   0
         Left            =   2100
         LinkItem        =   "Sè hiÖu chi tiÕt cÇn xem"
         MaxLength       =   12
         TabIndex        =   9
         Tag             =   "0"
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox txtShTk 
         Height          =   285
         Index           =   1
         Left            =   4140
         LinkItem        =   "Sè hiÖu tµi kho¶n cÇn in toµn bé chi tiÕt"
         MaxLength       =   12
         TabIndex        =   27
         Tag             =   "0"
         Top             =   2460
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox CboNT 
         Height          =   315
         ItemData        =   "Fbckt.frx":B680
         Left            =   4500
         List            =   "Fbckt.frx":B682
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   5340
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Line Line 
         Index           =   11
         X1              =   -74940
         X2              =   -67980
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Line 
         Index           =   10
         X1              =   -74880
         X2              =   -71400
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line 
         Index           =   9
         X1              =   -74880
         X2              =   -66960
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line 
         Index           =   8
         X1              =   240
         X2              =   6900
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line 
         Index           =   7
         X1              =   120
         X2              =   6840
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line 
         Index           =   0
         X1              =   300
         X2              =   3780
         Y1              =   -300
         Y2              =   -300
      End
      Begin VB.Line Line 
         Index           =   6
         X1              =   -74760
         X2              =   -67920
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line 
         Index           =   5
         X1              =   300
         X2              =   3780
         Y1              =   -300
         Y2              =   -300
      End
      Begin VB.Line Line 
         Index           =   1
         X1              =   -74880
         X2              =   -71520
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line 
         Index           =   4
         X1              =   -74760
         X2              =   -67920
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lbkh 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   -71400
         TabIndex        =   126
         Tag             =   "1"
         Top             =   1140
         Width           =   3495
      End
      Begin VB.Label lbkh 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   -71460
         TabIndex        =   142
         Tag             =   "1"
         Top             =   2220
         Width           =   3495
      End
      Begin VB.Label LbCT 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   -69840
         TabIndex        =   153
         Tag             =   "1"
         Top             =   5880
         Width           =   2775
      End
      Begin VB.Line Line1 
         X1              =   -74760
         X2              =   -66840
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "L·i d­ cã"
         Height          =   255
         Index           =   8
         Left            =   8940
         TabIndex        =   120
         Tag             =   "Credit Interest Rate"
         Top             =   60
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LbCT 
         ForeColor       =   &H00FF0000&
         Height          =   615
         Index           =   0
         Left            =   -69960
         TabIndex        =   32
         Tag             =   "1"
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   11
         Left            =   -69600
         TabIndex        =   149
         Tag             =   "1"
         Top             =   5040
         Width           =   2535
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   405
         Index           =   10
         Left            =   -68760
         TabIndex        =   118
         Tag             =   "1"
         Top             =   4140
         Width           =   1815
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   -4635
         TabIndex        =   110
         Tag             =   "1"
         Top             =   7140
         Width           =   3255
      End
      Begin VB.Label LbTenVT 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   -71460
         TabIndex        =   134
         Tag             =   "1"
         Top             =   4020
         Width           =   3495
      End
      Begin VB.Label Label 
         Caption         =   "In ®Õn tµi kho¶n cÊp"
         Height          =   255
         Index           =   7
         Left            =   5340
         TabIndex        =   152
         Tag             =   "Level of Account"
         Top             =   1860
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label 
         Caption         =   "®Õn"
         Height          =   255
         Index           =   6
         Left            =   7140
         TabIndex        =   35
         Tag             =   "to"
         Top             =   2700
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   115
         Tag             =   "1"
         Top             =   3780
         Width           =   2055
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   6
         Left            =   7320
         TabIndex        =   62
         Tag             =   "1"
         Top             =   4380
         Width           =   615
      End
      Begin VB.Label Label 
         Caption         =   "ngµy, tÝnh ®Õn"
         Height          =   255
         Index           =   3
         Left            =   -71340
         TabIndex        =   25
         Tag             =   "date, from"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   5940
         TabIndex        =   21
         Tag             =   "1"
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   -68760
         TabIndex        =   117
         Tag             =   "1"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label LbTenVT 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   -69840
         TabIndex        =   108
         Tag             =   "1"
         Top             =   720
         Width           =   2895
      End
      Begin VB.Line Line 
         Index           =   3
         X1              =   -74760
         X2              =   -66720
         Y1              =   1960
         Y2              =   1960
      End
      Begin VB.Line Line 
         Index           =   2
         X1              =   180
         X2              =   6900
         Y1              =   4980
         Y2              =   4980
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   4380
         TabIndex        =   11
         Tag             =   "1"
         Top             =   780
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   1
         Left            =   7500
         TabIndex        =   29
         Tag             =   "1"
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label LbTenTk 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   116
         Tag             =   "1"
         Top             =   4140
         Width           =   3375
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   212
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
      Begin VB.OptionButton OptNN 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ViÖt"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   213
         Tag             =   "VN"
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton OptNN 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Anh"
         Height          =   255
         Index           =   1
         Left            =   1080
         MaskColor       =   &H00EC5928&
         TabIndex        =   214
         Tag             =   "English"
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   1
      Left            =   240
      TabIndex        =   207
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
      Begin MSMask.MaskEdBox MedNgay 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   209
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MedNgay 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   211
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptChar      =   "_"
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "®Õn ngµy"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   210
         Tag             =   "to"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tõ ngµy"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   208
         Tag             =   "From"
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.OptionButton OptTG 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Theo ngµy"
      Height          =   255
      Index           =   1
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   201
      Tag             =   "By date"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton OptTG 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Theo th¸ng"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   200
      Tag             =   "By month"
      Top             =   2160
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   0
      Left            =   240
      TabIndex        =   202
      Top             =   3120
      Width           =   1455
      Begin VB.ComboBox CboThang 
         Height          =   315
         Index           =   0
         ItemData        =   "Fbckt.frx":B684
         Left            =   120
         List            =   "Fbckt.frx":B686
         Style           =   2  'Dropdown List
         TabIndex        =   204
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox CboThang 
         Height          =   315
         Index           =   1
         ItemData        =   "Fbckt.frx":B688
         Left            =   120
         List            =   "Fbckt.frx":B68A
         Style           =   2  'Dropdown List
         TabIndex        =   206
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Tõ th¸ng"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   203
         Tag             =   "From"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "®Õn th¸ng"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   205
         Tag             =   "to"
         Top             =   840
         Width           =   735
      End
   End
   Begin ComctlLib.ProgressBar GauGe 
      Height          =   255
      Left            =   840
      TabIndex        =   218
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   2
      Left            =   60
      Picture         =   "Fbckt.frx":B68C
      Style           =   1  'Graphical
      TabIndex        =   217
      Tag             =   "&Return"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   1
      Left            =   360
      Picture         =   "Fbckt.frx":CAAE
      Style           =   1  'Graphical
      TabIndex        =   216
      Tag             =   "&Print"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Index           =   0
      Left            =   1050
      Picture         =   "Fbckt.frx":DF10
      Style           =   1  'Graphical
      TabIndex        =   215
      Tag             =   "&View"
      Top             =   6360
      Width           =   975
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Index           =   3
      Left            =   10800
      TabIndex        =   237
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
      ForeColor       =   4194304
      BackColor       =   8454016
      Caption         =   "Tµi s¶n cè ®Þnh"
      Size            =   "3625;873"
      FontName        =   "VK Sans Serif"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Index           =   2
      Left            =   10800
      TabIndex        =   236
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
      ForeColor       =   4194304
      BackColor       =   8454016
      Caption         =   "C«ng nî - Doanh thu"
      Size            =   "3625;873"
      FontName        =   "VK Sans Serif"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Index           =   1
      Left            =   10800
      TabIndex        =   235
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
      ForeColor       =   4194304
      BackColor       =   8454016
      Caption         =   "Hµng hãa - C«ng tr×nh"
      Size            =   "3625;873"
      FontName        =   "VK Sans Serif"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Index           =   0
      Left            =   10800
      TabIndex        =   234
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
      ForeColor       =   4194304
      BackColor       =   8454016
      Caption         =   "Sæ tæng hîp"
      Size            =   "3625;873"
      FontName        =   "VK Sans Serif"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FBcKt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim baocao As Integer                 ' ChØ sè b¸o c¸o ®­îc chän
Dim StopPrint As Boolean              ' = True nÕu ngõng in c¸c danh s¸ch b¸o c¸o

Dim Pdelay As Integer

Dim ngay(0 To 2) As Date

Private Const DAUKY = 100
Private Const TANG = 200
Private Const GIAM = 300
Private Const BINHQUAN = 400
Dim tp As New Cls154
Dim pSelectSQL As String
Dim TenFile As String
Dim TenBaoCao As String
Dim SoLieu(1 To 12) As Boolean

Dim xlapp As Excel.Application, xlsheet As Worksheet

Dim NLB As String, KTT As String, GD As String
Dim nn As Integer


Private Sub cmd_Click(Index As Integer)
    txtsh(Index).Text = FrmNhanVien.ChonNV(txtsh(Index).Text)
    RFocus txtsh(Index)
End Sub

Private Sub CmdAdd_Click()
    FrmBCVT.Show 1
End Sub

Private Sub CmdCt_Click(Index As Integer)
    TxtCT(Index).Text = FrmTP.ChonTP(TxtCT(Index).Text)
    TxtCT_LostFocus Index
End Sub

Private Sub cmdkh_Click(Index As Integer)
    Me.MousePointer = 11
    txtshkh(Index).Text = FrmKhachHang.ChonKhachHang(txtshkh(Index).Text)
    Me.MousePointer = 0
    RFocus txtshkh(Index)
End Sub
'====================================================================================================
' Chän sè hiÖu tµi kho¶n
'====================================================================================================
Private Sub cmdtk_Click(Index As Integer)
    Me.MousePointer = 11
    txtShTk(Index).Text = FrmTaikhoan.ChonTk(txtShTk(Index).Text)
    RFocus txtShTk(Index)
    Me.MousePointer = 0
End Sub

'====================================================================================================
' Chän sè hiÖu chi tiÕt vËt t­
'====================================================================================================
Private Sub cmdvt_Click(Index As Integer)
    Me.MousePointer = 11
    txtShVT(Index).Text = FrmVattu.ChonVattu(txtShVT(Index).Text)
    Me.MousePointer = 0
    RFocus txtShVT(Index)
End Sub
'====================================================================================================
' Xem, In hoÆc kÕt thóc
'====================================================================================================

'====================================================================================================
Public Sub Command_Click(Index As Integer)
If OptTG(0).Value = False Then
  If IsDate(MedNgay(1).Text) And IsDate(MedNgay(2).Text) Then
        If CDate(MedNgay(2).Text) < CDate(MedNgay(1).Text) Then
            MedNgay(2).Text = MedNgay(1).Text
            
        End If
  End If
  Else
   If IsDate(CboThang(0).Text) And IsDate(CboThang(1).Text) Then
        If CDate(CboThang(1).Text) < CDate(CboThang(0).Text) Then
            CboThang(1).Text = CboThang(0).Text
          
        End If
  End If
  End If
   If IsDate(MedNgay(1).Text) And IsDate(MedNgay(2).Text) Then
        If CDate(MedNgay(2).Text) < CDate(MedNgay(1).Text) Then
            MedNgay(1).Text = MedNgay(2).Text
         '   MsgBox "Thong bao"
        '   Exit Sub59
        
        End If
  End If
  ' ke thuc kiem tra ngay
    Dim taikhoan As New ClsTaikhoan, doiung As New ClsTaikhoan
    Dim rs_tk As Recordset, tdau As Integer, tcuoi As Integer, n1 As Date
    Dim mkho As Long, mtk As Long, mloai As Long, mn As Long, i As Integer
    Dim mdt1 As Long, mdt2 As Long, mdt3 As Long, mdt As Long, k As Integer
    '///////////////////////////////////////////////
    If Index = 0 Or Index = 1 Then
    If OptBC(100).Value = True Or OptBC(101).Value = True Or OptBC(102).Value = True Or OptBC(103).Value = True Or OptBC(104).Value = True Or OptBC(105).Value = True Then
    
    
    
    FBcTC.MedNgay(0).Text = MedNgay(0).Text
    FBcTC.MedNgay(1).Text = MedNgay(1).Text
    FBcTC.CboThang(0).Text = CboThang(0).Text
    FBcTC.CboThang(1).Text = CboThang(1).Text
    FBcTC.OptTG(0).Value = OptTG(0).Value
    
    FBcTC.OptTG(1).Value = OptTG(1).Value
    FBcTC.OptBC(0).Value = OptBC(100).Value
    If OptBC(100).Value = False Then
            If Len(Trim(txtShTk(8).Text)) <= 0 Then
            RFocus txtShTk(8)
            Exit Sub
            End If
    End If
    
    FBcTC.OptBC(1).Value = OptBC(101).Value
    FBcTC.OptBC(15).Value = OptBC(102).Value
    FBcTC.OptBC(18).Value = OptBC(103).Value
    FBcTC.OptBC(14).Value = OptBC(104).Value
    FBcTC.OptBC(16).Value = OptBC(105).Value
    FBcTC.OptBC(6).Value = OptBC(106).Value
    FBcTC.txtShTk(0).Text = txtShTk(8).Text
     FBcTC.txtShTk(3).Text = txtShTk(8).Text
     
    txtShTk(0).tag = 39
     FBcTC.txtShTk_LostFocus 0
    FBcTC.Command_Click 0
    
    Exit Sub
    End If
    End If
    '///////////////////////////////////////////////////////
    Select Case Index
        Case 0, 1:
            If Not SSTab.TabEnabled(SSTab.Tab) Then Exit Sub
            If Not OptBC(baocao).TabStop Then Exit Sub
            If pRpt = 1 And (Not ChoXemBC(baocao, "C")) Then
                ErrMsg er_KoSD
                Exit Sub
            End If
            Me.MousePointer = 11
            GauGe.Value = 0
            If CboThang(0).Enabled And CboThang(1).ListIndex < CboThang(0).ListIndex Then CboThang(1).ListIndex = CboThang(0).ListIndex
            If baocao = 61 Then
                CboThang(0).ListIndex = 0
                CboThang(1).ListIndex = 11
            End If
            tdau = CboThang(0).ItemData(CboThang(0).ListIndex)
            tcuoi = CboThang(1).ItemData(CboThang(1).ListIndex)
            If OptDT(3).Value And CboVV(1).ListIndex >= 0 Then mdt = CboVV(1).ItemData(CboVV(1).ListIndex) Else mdt = 0
            If pSoVV > 0 And ChkTT(0).Value = 1 And CboTT(0).ListIndex >= 0 And mdt > 0 Then mdt1 = CboTT(0).ItemData(CboTT(0).ListIndex) Else mdt1 = 0
            If pSoVV > 1 And ChkTT(1).Value = 1 And CboTT(1).ListIndex >= 0 And mdt > 0 Then mdt2 = CboTT(1).ItemData(CboTT(1).ListIndex) Else mdt2 = 0
            If pSoVV > 2 And ChkTT(2).Value = 1 And CboTT(2).ListIndex >= 0 And mdt > 0 Then mdt3 = CboTT(2).ItemData(CboTT(2).ListIndex) Else mdt3 = 0
            SetRptInfo
            Select Case baocao
                Case 83:
                    If txtshkh(0).tag = 0 Then
                        MsgBox "H·y nhËp sè hiÖu chi tiÕt c«ng nî!", vbExclamation, App.ProductName
                        RFocus txtshkh(0)
                        GoTo LoiBC
                    End If
                    If OptTG(0).Value Then
                        BangKeTichSo2 txtShTk(7).Text, txtshkh(0).tag, NgayDauThang(pNamTC, tdau), NgayCuoiThang(pNamTC, tcuoi)
                    Else
                        BangKeTichSo2 txtShTk(7).Text, txtshkh(0).tag, ngay(1), ngay(2)
                    End If
                Case 84:
                    If CboKH.ListIndex < 0 Then GoTo LoiBC
                    If OptTG(0).Value Then
                        InBangThopCN tdau, tcuoi, IIf(Opt(3).Value And txtShTk(7).tag > 0, txtShTk(7).Text, ""), IIf(Opt(2).Value, CboKH.ItemData(CboKH.ListIndex), 0), ChkDu(14).Value, 1
                    Else
                        InBangThopCNN ngay(1), ngay(2), IIf(Opt(3).Value And txtShTk(7).tag > 0, txtShTk(7).Text, ""), IIf(Opt(2).Value, CboKH.ItemData(CboKH.ListIndex), 0), 1
                    End If
                Case 81:
                    InTHChietKhau IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2)
                Case 80:
                    InCTKhuyenMai IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), txtshkh(0).tag, txtShVT(1).tag
                Case 78:
                    If txtShTk(7).Text = "" Then txtShTk(7).Text = "511"
                    taikhoan.InitTaikhoanSohieu txtShTk(7).Text
                    If taikhoan.tk_id <> TKDT_ID Then
                        RFocus txtShTk(7)
                        GoTo LoiBC
                    End If
                    If Not SoChiTiet3(txtShTk(7).Text, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2)) Then GoTo LoiBC
                Case 76:
                    InCdt2 tdau, tcuoi, CboThang(2).ListIndex + 1, ChkDu(6).Value
                Case 75:
                    If txtShTk(8).tag = 0 Then
                        MsgBox "H·y nhËp sè hiÖu chi tiÕt !", vbExclamation, App.ProductName
                        RFocus txtShTk(8)
                        GoTo LoiBC
                    End If
                    T txtShTk(8).tag, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), CboThang(2).ListIndex + 1
                Case 74:
                    If CboKho(0).ListIndex < 0 Then
                        ErrMsg er_KhoHang
                        GoTo LoiBC
                    End If

                    mkho = IIf(OptKho(0).Value, CboKho(0).ItemData(CboKho(0).ListIndex), 0)
                    mloai = IIf(ChkDu(3).Value = 1, CboLoai(0).ItemData(CboLoai(0).ListIndex), 0)
                    
                    BKSanLuong mkho, mloai, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), IIf(ChkDu(2).Value = 1 And Len(txtShTk(4).Text) > 0, txtShTk(4).Text, "155")
                
                    GoTo LoiBC
                Case 72:
                    If Opt(3).Value And txtShTk(7).tag > 0 Then taikhoan.InitTaikhoanMaSo txtShTk(7).tag Else taikhoan.InitTaikhoanMaSo 0
                    n1 = IIf(OptTG(0).Value, NgayCuoiThang(pNamTC, tcuoi), ngay(2))
                    InBangThopCN_HN NgayDauThang(pNamTC, pThangDauKy), n1, taikhoan, IIf(Opt(2).Value, CboKH.ItemData(CboKH.ListIndex), 0)
                Case 73:
                    If txtshkh(0).tag = 0 Then
                        MsgBox "H·y nhËp sè hiÖu chi tiÕt c«ng nî!", vbExclamation, App.ProductName
                        RFocus txtshkh(0)
                        GoTo LoiBC
                    End If
                    If Opt(3).Value And txtShTk(7).tag > 0 Then taikhoan.InitTaikhoanMaSo txtShTk(7).tag Else taikhoan.InitTaikhoanMaSo 0
                    n1 = IIf(OptTG(0).Value, NgayCuoiThang(pNamTC, tcuoi), ngay(2))
                    InChitietCN_HN txtshkh(0).tag, NgayDauThang(pNamTC, pThangDauKy), n1, taikhoan
                Case 71:
                    Baocao211 tcuoi
                    GoTo LoiBC
                Case 43:
                    If CboKho(0).ListCount > 0 Then Baocao152 IIf(OptKho(0).Value, CboKho(0).ItemData(CboKho(0).ListIndex), 0), tcuoi
                    GoTo LoiBC
                Case 69:
                    If txtShTk(8).tag = 0 Then
                        If txtShTk(2).tag = 0 Then
                            RFocus txtShTk(8)
                            GoTo LoiBC
                        Else
                            taikhoan.InitTaikhoanMaSo txtShTk(2).tag
                            If CboNT.ListIndex >= 0 Then
                                mn = CboNT.ItemData(CboNT.ListIndex)
                                If mn = 0 Then mn = -1
                            Else
                                mn = 0
                            End If
                        End If
                    Else
                        taikhoan.InitTaikhoanMaSo txtShTk(8).tag
                        mn = 0
                    End If
                    BangKeTichSo taikhoan, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), Cdbl5(txtLS(0).Text), Cdbl5(txtLS(1).Text), mn
                Case 54:
                    BKDinhMuc tcuoi
                Case 67:
                    If CboKH.ListIndex < 0 Then GoTo LoiBC
                    InCongNoQuaHanMuc tdau, IIf(Opt(3).Value And txtShTk(7).tag > 0, txtShTk(7).Text, ""), IIf(Opt(2).Value, CboKH.ItemData(CboKH.ListIndex), 0)
                Case 59:
                    If TxtCT(0).tag = 0 Then
                        MsgBox "H·y nhËp sè hiÖu Hîp ®ång hoÆc chi tiÕt Hîp ®ång !", vbExclamation, App.ProductName
                        RFocus TxtCT(0)
                        GoTo LoiBC
                    End If
                    ChiPhiCT TxtCT(0).tag, tdau, tcuoi
                    'ChiPhiGiaThanh TxtCT(0).tag, tdau, tcuoi
                Case 106:
                    If txtShTk(0).tag = 0 Then
                        ErrMsg er_SHTaiKhoan
                        RFocus txtShTk(0)
                        GoTo LoiBC
                    End If
                    If ChkDu(0).Value = 1 And txtShTk(5).tag = 0 Then
                        RFocus txtShTk(5)
                        GoTo LoiBC
                    End If
                    If ChkDu(0).Value = 1 Then
                        doiung.InitTaikhoanMaSo txtShTk(5).tag
                    Else
                        doiung.InitTaikhoanMaSo 0
                    End If
                
                    taikhoan.InitTaikhoanMaSo txtShTk(0).tag
                    'them
                     If Not InSocaiTk_them(taikhoan, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), True, doiung.sohieu, 0, 2, nn) Then GoTo LoiBC
                   
                      Case 0:
                    If txtShTk(0).tag = 0 Then
                        ErrMsg er_SHTaiKhoan
                        RFocus txtShTk(0)
                        GoTo LoiBC
                    End If
                    If ChkDu(0).Value = 1 And txtShTk(5).tag = 0 Then
                        RFocus txtShTk(5)
                        GoTo LoiBC
                    End If
                    If ChkDu(0).Value = 1 Then
                        doiung.InitTaikhoanMaSo txtShTk(5).tag
                    Else
                        doiung.InitTaikhoanMaSo 0
                    End If
                
                    taikhoan.InitTaikhoanMaSo txtShTk(0).tag
                    'them
                    If Index = 0 Then
                     If Not InSocaiTk(taikhoan, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), True, doiung.sohieu, 0, 1, nn) Then GoTo LoiBC
                   ElseIf taikhoan.MaTC = taikhoan.MaSo Or taikhoan.MaTC = 0 Then
                        If Not InSocaiTk(taikhoan, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), True, doiung.sohieu, 0, 0, nn) Then GoTo LoiBC
                    Else
                        If Not InSoChitiet(taikhoan, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), True, doiung.sohieu, 0, 0, nn) Then GoTo LoiBC
                    End If
                Case 1:
                    If Index = 0 Or OptTG(1).Value Then GoTo KhongIn
                    If txtShTk(1).tag = 0 Then
                        ErrMsg er_SHTaiKhoan
                        RFocus txtShTk(1)
                        GoTo LoiBC
                    End If
                    taikhoan.InitTaikhoanMaSo txtShTk(1).tag
                    If taikhoan.tkcon = 0 Or taikhoan.MaTC <> taikhoan.MaSo Then
                        MsgBox "H·y chän tµi kho¶n Bé Tµi chÝnh ban hµnh ®­îc më chi tiÕt !", vbExclamation, App.ProductName
                        GoTo LoiBC
                    End If
                    Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo, SoHieu, Ten FROM HethongTK WHERE SoHieu LIKE '" _
                        + taikhoan.sohieu + "*' AND TkCon=0 AND MaNT<=0 ORDER BY SoHieu DESC", dbOpenSnapshot, dbForwardOnly)
                    StopPrint = False
                    frmMain.Rpt.Destination = 1
                    GauGe.Max = 10
                    Do While Not rs_tk.EOF And (Not StopPrint)
                        taikhoan.InitTaikhoanMaSo rs_tk!MaSo
                        SetRptInfo
                        If InSoChitiet(taikhoan, tdau, tcuoi, ngay(0), ngay(1), False, "", 0, 0, nn) Then
                            HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
                            InBaoCaoRPT
                            AppIdle Pdelay * 100
                        End If
                        If GauGe.Value < GauGe.Max Then GauGe.Value = GauGe.Value + 1
                        rs_tk.MoveNext
                    Loop
                    rs_tk.Close
                    GoTo LoiBC
                Case 2:
                    If Index = 0 Then GoTo KhongIn
                    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo, SoHieu, Ten FROM HethongTK" _
                        & " WHERE Loai > 0 AND MaTC > 0 AND MaTC <> MaSo AND TkCon=0 AND MaNT<=0" _
                        + IIf(Len(txtShCT(0).Text) > 0, " AND SoHieu >= '" + txtShCT(0).Text + "'", "") + IIf(Len(txtShCT(1).Text) > 0, " AND SoHieu <= '" + txtShCT(1).Text + "'", "") + " ORDER BY SoHieu", dbOpenSnapshot, dbForwardOnly)
                    StopPrint = False
                    frmMain.Rpt.Destination = 1
                    GauGe.Max = 300
                    Do While Not rs_tk.EOF And (Not StopPrint)
                        taikhoan.InitTaikhoanMaSo rs_tk!MaSo
                        SetRptInfo
                        If InSoChitiet(taikhoan, tdau, tcuoi, ngay(0), ngay(1), False, "", 0, 0, nn) Then
                            HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
                            InBaoCaoRPT
                            AppIdle Pdelay * 100
                        End If
                        If GauGe.Value < GauGe.Max Then GauGe.Value = GauGe.Value + 1
                        rs_tk.MoveNext
                    Loop
                    rs_tk.Close
                    GoTo LoiBC
                Case 3:
                    If ChkDu(1).Value = 1 Then
                        If txtShTk(6).tag = 0 Then
                            ErrMsg er_SHTaiKhoan
                            RFocus txtShTk(6)
                            GoTo LoiBC
                        End If
                        taikhoan.InitTaikhoanMaSo txtShTk(6).tag
                    Else
                        taikhoan.InitTaikhoanMaSo 0
                    End If
                    
                    InBangThop tdau, tcuoi, taikhoan, CboThang(2).ListIndex + 1, ChkDu(6).Value
                Case 4:
                    If txtShTk(2).tag = 0 Then
                        ErrMsg er_SHTaiKhoan
                        RFocus txtShTk(2)
                        GoTo LoiBC
                    End If
                    If CboNT.ListIndex < 0 Then
                        MsgBox "H·y chän chi tiÕt thanh to¸n theo nguyªn tÖ !", vbExclamation, App.ProductName
                        RFocus txtShTk(2)
                        GoTo LoiBC
                    End If
                    If CboNT.ItemData(CboNT.ListIndex) < 0 Then mtk = CboNT.ItemData(CboNT.ListIndex - 1) Else mtk = CboNT.ItemData(CboNT.ListIndex)
                    taikhoan.InitTaikhoanMaSo MaTKNguyenTe(txtShTk(2).Text, mtk)
                    If Not InSoChitietNT(taikhoan, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), CboNT.ItemData(CboNT.ListIndex), True) Then GoTo LoiBC
                Case 5:
                    If txtShTk(3).tag = 0 Then
                        ErrMsg er_SHTaiKhoan
                        RFocus txtShTk(3)
                        GoTo LoiBC
                    End If
                    If Index = 0 Then GoTo KhongIn
                    Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo, SoHieu, Ten, KyHieu FROM HethongTK INNER JOIN NguyenTe ON HethongTK.MaNT=NguyenTe.MaSo WHERE SoHieu LIKE '" + txtShTk(3).Text + "*' AND TkCon=0 ORDER BY SoHieu", dbOpenSnapshot, dbForwardOnly)
                    StopPrint = False
                    frmMain.Rpt.Destination = 1
                    GauGe.Max = 10
                    Do While Not rs_tk.EOF And (Not StopPrint)
                        HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten + " - " + rs_tk!KyHieu), 1
                        taikhoan.InitTaikhoanMaSo rs_tk!MaSo
                        SetRptInfo
                        If InSoChitietNT(taikhoan, tdau, tcuoi, ngay(1), ngay(2), CboNT.ItemData(CboNT.ListIndex), False) Then
                            HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
                            InBaoCaoRPT
                            AppIdle Pdelay * 100
                        End If
                        If GauGe.Value < GauGe.Max Then GauGe.Value = GauGe.Value + 1
                        rs_tk.MoveNext
                    Loop
                    rs_tk.Close
                    GoTo LoiBC
                Case 6:
                    If Index = 0 Then GoTo KhongIn
                    Set rs_tk = DBKetoan.OpenRecordset("SELECT MaSo, SoHieu, Ten FROM HethongTK WHERE MaNT>0 ORDER BY SoHieu", dbOpenSnapshot, dbForwardOnly)
                    StopPrint = False
                    frmMain.Rpt.Destination = 1
                    GauGe.Max = 10
                    Do While Not rs_tk.EOF And (Not StopPrint)
                        taikhoan.InitTaikhoanMaSo rs_tk!MaSo
                        SetRptInfo
                        If InSoChitietNT(taikhoan, tdau, tcuoi, ngay(1), ngay(2), CboNT.ItemData(CboNT.ListIndex), False) Then
                            HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
                            InBaoCaoRPT
                            AppIdle Pdelay * 100
                        End If
                        If GauGe.Value < GauGe.Max Then GauGe.Value = GauGe.Value + 1
                        rs_tk.MoveNext
                    Loop
                    rs_tk.Close
                    GoTo LoiBC
                Case 7:
                    InBangThopNT tdau, tcuoi
                Case 8:
                    InCdt False, tdau, tcuoi, CboThang(2).ListIndex + 1, ChkDu(6).Value
                Case 9:
                
                    InCdt True, tdau, tcuoi, CboThang(2).ListIndex + 1, ChkDu(6).Value
                Case 10:
                    If pDTTP <> 0 Then
                        If Not InChiPhi3(tdau, tcuoi, nn) Then GoTo LoiBC
                    Else
                        If Not InChiPhi(tdau, tcuoi, True, nn) Then GoTo LoiBC
                    End If
                Case 11:
                    If pDTTP <> 0 Then
                        taikhoan.InitTaikhoanSohieu "911"
                        If taikhoan.tkcon > 0 Then
                            If InKetQua3(tdau, tcuoi, 0, nn, 1) Then InBaoCaoRPT nn
                            If Not InKetQua(tdau, tcuoi, 0, True, nn) Then GoTo LoiBC
                        Else
                            If Not InKetQua3(tdau, tcuoi, 0, nn) Then GoTo LoiBC
                        End If
                    Else
                        If Not InKetQua(tdau, tcuoi, 0, True, nn) Then GoTo LoiBC
                    End If
                Case 50:
                    If Not InChiPhi2(tdau, tcuoi, nn) Then GoTo LoiBC
                Case 51:
                    If pDTTP <> 0 Then
                        InBKGiaThanh2 tdau, tcuoi
                    Else
                        InBKGiaThanh tcuoi
                    End If
                    GoTo LoiBC
                Case 12, 64:
                    If CboKho(0).ListIndex < 0 Then
                        ErrMsg er_KhoHang
                        GoTo LoiBC
                    End If
                    If CboLoai(1).ListIndex < 0 Then
                        ErrMsg er_NguonNX
                        GoTo LoiBC
                    End If
                    If txtShVT(0).tag = 0 Then
                        ErrMsg er_SHVattu
                        RFocus txtShVT(0)
                        GoTo LoiBC
                    End If
                    If ChkDu(7).Value = 1 And txtShTk(10).tag = 0 Then
                        ErrMsg er_SHTaiKhoan
                        RFocus txtShVT(10)
                        GoTo LoiBC
                    End If
                    If OptTG(0).Value Then
                        If Not InTheKho2(IIf(OptKho(0).Value, CboKho(0).ItemData(CboKho(0).ListIndex), 0), _
                            txtShVT(0).tag, tdau, tcuoi, True, IIf(ChkDu(4).Value = 1, CboLoai(1).ItemData(CboLoai(1).ListIndex), 0), IIf(ChkDu(7).Value = 1 And txtShTk(10).tag > 0, txtShTk(10).Text, ""), IIf(baocao = 64, 1, 0)) Then GoTo LoiBC
                    Else
                        If Not InTheKho2N(IIf(OptKho(0).Value, CboKho(0).ItemData(CboKho(0).ListIndex), 0), _
                            txtShVT(0).tag, ngay(1), ngay(2), True, IIf(ChkDu(4).Value = 1, CboLoai(1).ItemData(CboLoai(1).ListIndex), 0), IIf(ChkDu(7).Value = 1 And txtShTk(10).tag > 0, txtShTk(10).Text, ""), IIf(baocao = 64, 1, 0)) Then GoTo LoiBC
                    End If
                Case 13:
                    If CboKho(0).ListIndex < 0 Then
                        ErrMsg er_KhoHang
                        GoTo LoiBC
                    End If
                    If CboLoai(1).ListIndex < 0 Then
                        ErrMsg er_NguonNX
                        GoTo LoiBC
                    End If
                    StopPrint = False
                    mkho = CboKho(0).ItemData(CboKho(0).ListIndex)
                    If Index = 0 Then
                        InTheKho3 mkho, tdau, tcuoi
                        InBaoCaoRPT
                    Else
                        Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW TonKho.MaVattu, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi FROM" _
                            & " TonKho INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo WHERE TonKho.MaSoKho=" _
                            + CStr(mkho) + " ORDER BY Vattu.SoHieu", dbOpenSnapshot, dbForwardOnly)
                            
                        k = CInt5(FrmGetStr.GetString("NhËp sè 0 ®Ó in sæ chi tiÕt vËt t­ hµng ho¸, sè 2 ®Ó in thÎ kho", "In toµn bé", "0"))
                        
                        frmMain.Rpt.Destination = 1
                        GauGe.Max = 100
                        Do While (Not rs_tk.EOF) And (Not StopPrint)
                            txtShVT(0).Text = rs_tk!sohieu
                            LbTenVT(0).Caption = rs_tk!TenVattu
                            txtShVT(0).tag = rs_tk!MaVattu
                            SetRptInfo
                            If InTheKho2(mkho, rs_tk!MaVattu, tdau, tcuoi, False, IIf(ChkDu(4).Value = 1, CboLoai(1).ItemData(CboLoai(1).ListIndex), 0), , k) Then
                                HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!TenVattu), 1
                                InBaoCaoRPT
                                AppIdle Pdelay * 100
                            End If
                            If GauGe.Value < GauGe.Max Then GauGe.Value = GauGe.Value + 1
                            rs_tk.MoveNext
                        Loop
                        rs_tk.Close
                    End If
                    GoTo LoiBC
                Case 14, 65:
                    If CboKho(0).ListIndex < 0 Then
                        ErrMsg er_KhoHang
                        GoTo LoiBC
                    End If
                    If ChkDu(2).Value = 1 And txtShTk(4).tag = 0 Then
                        ErrMsg er_SHTKVT
                        RFocus txtShTk(4)
                        GoTo LoiBC
                    End If
                    mkho = IIf(OptKho(0).Value, CboKho(0).ItemData(CboKho(0).ListIndex), 0)
                    mloai = IIf(ChkDu(3).Value = 1, CboLoai(0).ItemData(CboLoai(0).ListIndex), 0)
                    If OptTG(0).Value Then
                        InLuyKe mkho, IIf(ChkDu(2).Value = 0, "", txtShTk(4).Text), mloai, tdau, tcuoi, True, IIf(baocao = 65, 1, 0)
                    Else
                        InLuyKeN mkho, IIf(ChkDu(2).Value = 0, "", txtShTk(4).Text), mloai, ngay(1), ngay(2), True, IIf(baocao = 65, 1, 0)
                    End If
                Case 46:
                    If CboKho(0).ListIndex < 0 Then
                        ErrMsg er_KhoHang
                        GoTo LoiBC
                    End If
                    If ChkDu(2).Value = 1 And txtShTk(4).tag = 0 Then
                        ErrMsg er_SHTKVT
                        RFocus txtShTk(4)
                        GoTo LoiBC
                    End If
                    mkho = IIf(OptKho(0).Value, CboKho(0).ItemData(CboKho(0).ListIndex), 0)
                    mtk = IIf(ChkDu(2).Value = 0, 0, txtShTk(4).tag)
                    mloai = IIf(ChkDu(3).Value = 1, CboLoai(0).ItemData(CboLoai(0).ListIndex), 0)
                    InLuyKeX mkho, mtk, mloai, tdau, tcuoi
                    GoTo LoiBC
                Case 52:
                    If CboKho(0).ListIndex < 0 Then
                        ErrMsg er_KhoHang
                        GoTo LoiBC
                    End If
                    InLCNoiBo CboKho(0).ItemData(CboKho(0).ListIndex), IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2)
                    GoTo LoiBC
                Case 47, 48:
                    If CboKho(0).ListIndex < 0 Then
                        ErrMsg er_KhoHang
                        GoTo LoiBC
                    End If
                    If ChkDu(2).Value = 1 And txtShTk(4).tag = 0 Then
                        ErrMsg er_SHTKVT
                        RFocus txtShTk(4)
                        GoTo LoiBC
                    End If
                    mkho = IIf(OptKho(0).Value, CboKho(0).ItemData(CboKho(0).ListIndex), 0)
                    mtk = IIf(ChkDu(2).Value = 0, 0, txtShTk(4).tag)
                    mloai = IIf(ChkDu(3).Value = 1, CboLoai(0).ItemData(CboLoai(0).ListIndex), 0)
                    If OptTG(0).Value Then
                        InLuyKeHM mkho, mtk, mloai, NgayCuoiThang(pNamTC, tcuoi), IIf(baocao = 47, -1, 1)
                    Else
                        InLuyKeHM mkho, mtk, mloai, ngay(2), IIf(baocao = 47, -1, 1)
                    End If
                Case 15, 16:
                    If CboKho(0).ListIndex < 0 Then
                        ErrMsg er_KhoHang
                        GoTo LoiBC
                    End If
                    If ChkDu(2).Value = 1 And txtShTk(4).tag = 0 Then
                        ErrMsg er_SHTKVT
                        RFocus txtShTk(4)
                        GoTo LoiBC
                    End If
                    mkho = IIf(OptKho(0).Value, CboKho(0).ItemData(CboKho(0).ListIndex), 0)
                    mtk = IIf(ChkDu(2).Value = 0, 0, txtShTk(4).tag)
                    mloai = IIf(ChkDu(3).Value = 1, CboLoai(0).ItemData(CboLoai(0).ListIndex), 0)
                    mn = IIf(ChkDu(4).Value = 1, CboLoai(1).ItemData(CboLoai(1).ListIndex), 0)
                    
                    InNguonNhapXuat mkho, mtk, mloai, mn, tdau, tcuoi, IIf(baocao = 15, -1, 1), IIf(OptTG(1).Value, 1, 0), ngay(1), ngay(2)
               Case 17:
                    If CboKH.ListIndex < 0 Then GoTo LoiBC
                    If Not InCTNoQuaHan(ngay(0), CInt5(txtHan.Text), IIf(Opt(2).Value, CboKH.ItemData(CboKH.ListIndex), 0)) Then GoTo LoiBC
                Case 18 To 31, 45, 61, 68, 79:
                    Dim tg As String
                    If CboThang(0).Enabled And CboThang(1).Enabled Then
                        tg = ThoiGian(tdau, tcuoi)
                    Else
                      If CboThang(0).Enabled Then
                            tcuoi = tdau
                      Else
                            tdau = tcuoi
                      End If
                      tg = ThoiGian(tdau, tdau)
                    End If
                    ' ChuÈn bÞ sè liÖu
                    TinhSoLieuBaoCao
                    ' T¹o b¸o c¸o
                    TaoBaoCao
                    ' Tªn File d÷ liÖu vµ b¸o c¸o
                    frmMain.Rpt.ReportFileName = TenFile
                    RptSetDate NgayCuoiThang(pNamTC, tcuoi)
                    ' C«ng thøc
                    frmMain.Rpt.Formulas(2) = "TenBaoCao = IF PageNumber() = 1 THEN '" + TenBaoCao + "'"
                    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + tg + "'"
                    If Opt(1).Value Then
                        frmMain.Rpt.Formulas(0) = "TenCty='" + pTenCn + "'"
                        frmMain.Rpt.Formulas(1) = "TenCn='§¬n vÞ: " + Cbo.Text + "'"
                    End If
                Case 32:
                    InTK CboThang(2).ListIndex + 1, tdau, tcuoi
                Case 33:
                    InPSVuViec CboVV(0).ItemData(CboVV(0).ListIndex), tdau, tcuoi, CboVV(0).Text
                Case 34:
                    If CboKho(1).ListIndex < 0 Or CboKho(2).ListIndex < 0 Then Exit Sub
                    If OptDT(2).Value Then
                            InTHDoanhThuVAT IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2)
                    Else
                            InTHDoanhThu IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), IIf(OptDT(0).Value, 0, 1), IIf(ChkDu(9).Value = 1, CboKho(1).ItemData(CboKho(1).ListIndex), 0), IIf(ChkDu(8).Value = 1, txtshkh(2).Text, ""), IIf(ChkDu(11).Value = 1, txtsh(0).tag, 0), IIf(ChkDu(10).Value = 1, CboKho(2).ItemData(CboKho(2).ListIndex), 0), mdt, mdt1, mdt2, mdt3, IIf(ChkDu(12).Value = 1, txtShTk(9).Text, "")
                    End If
                Case 35:
                    If CboKH.ListIndex < 0 Then GoTo LoiBC
                    If OptTG(0).Value Then
                        InBangThopCN tdau, tcuoi, IIf(Opt(3).Value And txtShTk(7).tag > 0, txtShTk(7).Text, ""), IIf(Opt(2).Value, CboKH.ItemData(CboKH.ListIndex), 0), ChkDu(5).Value
                    Else
                        InBangThopCNN ngay(1), ngay(2), IIf(Opt(3).Value And txtShTk(7).tag > 0, txtShTk(7).Text, ""), IIf(Opt(2).Value, CboKH.ItemData(CboKH.ListIndex), 0)
                    End If
                Case 36:
                    If txtshkh(0).tag = 0 Then
                        ErrMsg er_SHKhachHang
                        RFocus txtshkh(0)
                        GoTo LoiBC
                    End If
                    If txtShTk(7).tag = 0 Then
                        taikhoan.InitTaikhoanMaSo SelectSQL("SELECT MaTaiKhoan AS F1 FROM SoDuKhachHang WHERE MaKhachHang=" + CStr(txtshkh(0).tag))
                        If taikhoan.MaSo > 0 Then
                            txtShTk(7).tag = taikhoan.MaSo
                            txtShTk(7).Text = taikhoan.sohieu
                            LbTenTk(7).Caption = taikhoan.Ten
                        Else
                            ErrMsg er_KoPS1
                            RFocus txtShTk(7)
                            GoTo LoiBC
                        End If
                    Else
                        taikhoan.InitTaikhoanMaSo txtShTk(7).tag
                    End If
                    If Not InSoChitietCN(taikhoan, txtshkh(0).tag, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2)) Then GoTo LoiBC
                Case 62:
                    If Index = 0 Then GoTo KhongIn
                    If txtShTk(7).tag = 0 Then
                        ErrMsg er_SHTaiKhoan
                        RFocus txtShTk(7)
                        GoTo LoiBC
                    Else
                        taikhoan.InitTaikhoanMaSo txtShTk(7).tag
                    End If
                    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW KhachHang.MaSo,KhachHang.SoHieu,KhachHang.Ten FROM (SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo) INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo WHERE HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*' GROUP BY KhachHang.SoHieu,KhachHang.Ten,KhachHang.MaSo", dbOpenSnapshot)
                    StopPrint = False
                    GauGe.Max = 50
                    frmMain.Rpt.Destination = 1
                    Do While Not rs_tk.EOF And (Not StopPrint)
                        If InSoChitietCN(taikhoan, rs_tk!MaSo, tdau, tcuoi, ngay(1), ngay(2)) Then
                            HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
                            InBaoCaoRPT
                            For i = 1 To Pdelay * 100
                                DoEvents
                            Next
                        End If
                        If GauGe.Value < GauGe.Max Then GauGe.Value = GauGe.Value + 1
                        rs_tk.MoveNext
                    Loop
                    rs_tk.Close
                    GoTo LoiBC
                Case 70:
                    If Index = 0 Then GoTo KhongIn
                    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW KhachHang.MaSo,KhachHang.SoHieu,KhachHang.Ten FROM (SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo) INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo WHERE HethongTK.SoHieu LIKE '" + txtShTk(7).Text + "*' GROUP BY KhachHang.MaSo,KhachHang.SoHieu,KhachHang.Ten", dbOpenSnapshot)
                    StopPrint = False
                    If rs_tk.recordCount > 0 Then
                        rs_tk.MoveLast
                        GauGe.Max = rs_tk.recordCount
                        GauGe.Value = 0
                        rs_tk.MoveFirst
                    End If
                    frmMain.Rpt.Destination = 1
                    Do While Not rs_tk.EOF And (Not StopPrint)
                        If THPSNhomCN2(txtshkh(0).tag, tdau, tcuoi) Then
                            HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
                            InBaoCaoRPT
                            For i = 1 To Pdelay * 100
                                DoEvents
                            Next
                        End If
                        GauGe.Value = GauGe.Value + 1
                        rs_tk.MoveNext
                    Loop
                    rs_tk.Close
                    Set rs_tk = Nothing
                    GoTo LoiBC
                Case 63:
                    If Index = 0 Then GoTo KhongIn
                    taikhoan.InitTaikhoanMaSo txtShTk(7).tag
                    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW KhachHang.MaSo,SoHieu,Ten FROM SoDuKhachHang INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo GROUP BY SoHieu,Ten,KhachHang.MaSo", dbOpenSnapshot)
                    StopPrint = False
                    GauGe.Max = 100
                    Do While Not rs_tk.EOF And (Not StopPrint)
                        If DoiChieuCN(IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), rs_tk!MaSo, taikhoan, 1) Then
                            HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
                            InBaoCaoRPT
                            AppIdle Pdelay * 100
                        End If
                        If GauGe.Value < GauGe.Max Then GauGe.Value = GauGe.Value + 1
                        rs_tk.MoveNext
                    Loop
                    rs_tk.Close
                    GoTo LoiBC
                Case 57:
                    If txtshkh(0).tag = 0 Then
                        ErrMsg er_SHKhachHang
                        RFocus txtshkh(0)
                        GoTo LoiBC
                    End If
                    If Not THPSNhomCN2(txtshkh(0).tag, tdau, tcuoi) Then GoTo LoiBC
                Case 49, 82:
                    If txtshkh(1).tag = 0 Then
                        ErrMsg er_SHKhachHang
                        RFocus txtshkh(1)
                        GoTo LoiBC
                    End If
                    taikhoan.InitTaikhoanMaSo txtShTk(7).tag
                    If Not DoiChieuCN(IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), txtshkh(1).tag, taikhoan, IIf(baocao = 49, 1, 0)) Then GoTo LoiBC
                Case 37:
                    THDoanhThuGiaVon IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), IIf(ChkDu(9).Value = 1, CboKho(1).ItemData(CboKho(1).ListIndex), 0), IIf(ChkDu(12).Value = 1, txtShTk(9).Text, ""), ChkDu(13).Value
                Case 38:
                    If txtShVT(1).tag = 0 Then
                        ErrMsg er_SHVattu
                        RFocus txtShVT(1)
                        GoTo LoiBC
                    End If
                    CTDoanhThuGiaVon IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), txtShVT(1).tag
                Case 39:
                    If CboKH.ListIndex < 0 Then Exit Sub
                    InCTDoanhThu IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), txtshkh(0).tag, txtShVT(1).tag, IIf(Opt(2).Value, CboKH.ItemData(CboKH.ListIndex), 0)
                Case 53:
                    If CboKH.ListIndex < 0 Then Exit Sub
                    InCTDoanhThu4 IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), txtshkh(0).tag, txtShVT(1).tag, IIf(Opt(2).Value, CboKH.ItemData(CboKH.ListIndex), 0)
                Case 66:
                    If txtShVT(1).tag = 0 Then
                        ErrMsg er_SHVattu
                        RFocus txtShVT(1)
                        GoTo LoiBC
                    End If
                    InCTDoanhThu2 IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), txtShVT(1).tag
                Case 40:
                    KeHopDongV tdau, tcuoi, txtshkh(0).tag, CboHD.ItemData(CboHD.ListIndex)
                Case 41:
                    KeHopDongR tdau, tcuoi, txtshkh(0).tag, CboHD.ItemData(CboHD.ListIndex)
                Case 42:
                    If txtShTk(8).tag = 0 Then
                        ErrMsg er_SHTaiKhoan
                        RFocus txtShTk(8)
                        GoTo LoiBC
                    End If
                    If Not BKChiTiet(xlapp, xlsheet, txtShTk(8).Text, IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2)) Then GoTo LoiBC
                Case 44:
                    If txtShTk(11).tag = 0 Then
                        ErrMsg er_SHTaiKhoan
                        RFocus txtShTk(11)
                        GoTo LoiBC
                    End If
                    mkho = IIf(OptKho(0).Value, CboKho(0).ItemData(CboKho(0).ListIndex), 0)
                    mtk = IIf(ChkDu(2).Value = 0, 0, txtShTk(4).tag)
                    mloai = IIf(ChkDu(3).Value = 1, CboLoai(0).ItemData(CboLoai(0).ListIndex), 0)
                    InTHXuatKho IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), mkho, mtk, mloai, txtShTk(11).Text
                Case 60, 77:
                    If TxtCT(1).tag = 0 Then
                        ErrMsg er_SHThanhPham
                        RFocus TxtCT(1)
                        GoTo LoiBC
                    End If
                    mkho = IIf(OptKho(0).Value, CboKho(0).ItemData(CboKho(0).ListIndex), 0)
                    mtk = IIf(ChkDu(2).Value = 0, 0, txtShTk(4).tag)
                    mloai = IIf(ChkDu(3).Value = 1, CboLoai(0).ItemData(CboLoai(0).ListIndex), 0)
                    InTHXuatKhoDT IIf(OptTG(0).Value, tdau, 0), IIf(OptTG(0).Value, tcuoi, 0), ngay(1), ngay(2), mkho, mtk, mloai, TxtCT(1).tag, IIf(baocao = 60, 0, 1)
                Case 55:
                    mtk = IIf(ChkDu(2).Value = 0, 0, txtShTk(4).tag)
                    mloai = IIf(ChkDu(3).Value = 1, CboLoai(0).ItemData(CboLoai(0).ListIndex), 0)
                    InBangKeTon IIf(OptTG(0).Value, tcuoi, 0), ngay(2), mtk, mloai
                    GoTo LoiBC
                Case 56:
                    taikhoan.InitTaikhoanMaSo txtShTk(8).tag
                    If taikhoan.MaSo = 0 Or (taikhoan.tk_id <> TKCNKH_ID And taikhoan.tk_id <> TKCNPT_ID) Then
                        THPSNhomTK txtShTk(8).Text, tdau, tcuoi
                    Else
                        THPSNhomCN taikhoan, tdau, tcuoi
                    End If
                Case 58:
                    If CboKH.ListIndex < 0 Then GoTo LoiBC
                    BangCDCno tdau, tcuoi, IIf(Opt(2).Value, CboKH.ItemData(CboKH.ListIndex), 0)
            End Select
            frmMain.Rpt.Destination = Index
            If Len(NLB) > 0 And NLB <> "..." Then frmMain.Rpt.Formulas(52) = "NLB='" + NLB + "'"
            If Len(KTT) > 0 And KTT <> "..." Then frmMain.Rpt.Formulas(53) = "KTT='" + KTT + "'"
            If Len(GD) > 0 And GD <> "..." Then frmMain.Rpt.Formulas(54) = "GD='" + GD + "'"
a:
            If Not RptOK(frmMain.Rpt.ReportFileName, nn) Then
                MsgBox "MÉu b¸o c¸o ®· bÞ thay ®æi!", vbCritical, App.ProductName
                GoTo LoiBC
            End If
            frmMain.Rpt.WindowTitle = OptBC(baocao).Caption
            On Error GoTo Bad
            frmMain.Rpt.Action = 1
            On Error GoTo 0
            GoTo LoiBC
        Case 2:
            Unload Me
    End Select
    GoTo LoiBC
KhongIn:
    ErrMsg er_KoXem
    GoTo LoiBC
Bad:
    If Err.number = 20510 Then
        frmMain.Rpt.Formulas(52) = ""
        frmMain.Rpt.Formulas(53) = ""
        frmMain.Rpt.Formulas(54) = ""
        GoTo a
    End If
    MsgBox "Error " + CStr(Err.number) + ": " + Err.Description, vbExclamation, App.ProductName
LoiBC:
    HienThongBao Me.Caption, 1
    GauGe.Value = GauGe.Max
    Set taikhoan = Nothing
    Set doiung = Nothing
    Set rs_tk = Nothing
    Me.MousePointer = 0
End Sub

Private Sub CommandButton1_Click(Index As Integer)

If Index = 0 Then
SSTab.Tab = 0
End If

If Index = 1 Then
SSTab.Tab = 1
End If

If Index = 2 Then
SSTab.Tab = 2
End If

If Index = 3 Then
SSTab.Tab = 3
End If
  CommandButton1(Index).BackColor = 8438015
End Sub

Private Sub CommandButton1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i
 For i = 0 To 3
        CommandButton1(i).BackColor = &H80FF80    '&HC0FFC0    '&H80000003
 Next
' MsgBox CStr(SSTab.Tabs)
If SSTab.Tabs = 0 Then
     CommandButton1(0).BackColor = 8438015
End If
If SSTab.Tabs = 1 Then
     CommandButton1(1).BackColor = 8438015
End If
If SSTab.Tabs = 2 Then
     CommandButton1(2).BackColor = 8438015
End If
If SSTab.Tabs = 3 Then
     CommandButton1(3).BackColor = 8438015
End If
      CommandButton1(Index).BackColor = 8438015
End Sub

' Xö lý phÝm nãng
'====================================================================================================
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

'====================================================================================================
' Khëi t¹o cöa sæ b¸o c¸o
'====================================================================================================
Private Sub Form_Load()
    Dim chi_so As Integer
    
    AddMonthToCbo CboThang(0)
    AddMonthToCbo CboThang(1)
    
    CboThang(2).ListIndex = CboThang(2).ListCount - 5
    
    ChkDu(11).Visible = (pNVBH > 0)
    txtsh(0).Visible = (pNVBH > 0)
    cmd(0).Visible = (pNVBH > 0)
    
    Int_RecsetToCbo "SELECT MaSo As F2,TenKho As F1 FROM KhoHang ORDER BY TenKho", CboKho(0)
    CboCopy CboKho(0), CboKho(1)
    Int_RecsetToCbo "SELECT PhanLoaiVattu.MaSo As F2, (PhanLoaiVattu.SoHieu + ' - '+ PhanLoaiVattu.TenPhanLoai) As F1" _
        & " FROM PhanLoaiVattu ORDER BY PhanLoaiVattu.SoHieu", CboLoai(0)
    Int_RecsetToCbo "SELECT MaSo As F2,SoHieu + ' - ' + DienGiai As F1 FROM NguonNhapXuat ORDER BY SoHieu", CboLoai(1)
    Int_RecsetToCbo "SELECT DoituongCT.MaSo As F2,(IIF(DoituongCT.MaKhachHang>0,KhachHang.Ten+' - '+DoituongCT.Sohieu+' - ','')+DienGiai+IIF(DoituongCT.MaKhachHang>0,' - '+ Format(NgayKy,'dd/mm/yy'),'')) As F1 FROM DoituongCT LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo ORDER BY  KhachHang.Ten,DoituongCT.SoHieu,DienGiai", CboVV(0)
    CboCopy CboVV(0), CboVV(1)
    Int_RecsetToCbo "SELECT Ten AS F1, MaSo as F2 FROM DTQly ORDER BY Ten", Cbo
    Opt(1).Enabled = Cbo.ListCount > 0
    Int_RecsetToCbo "SELECT DISTINCTROW MaSo As F2,SoHieu + ' - '  + TenPhanLoai As F1 FROM PhanLoaiKhachHang WHERE PLCon=0 AND LEFT(SoHieu,1)<>'#' ORDER BY SoHieu", CboKH
    Int_RecsetToCbo "SELECT MaSo As F2,SoHieu + ' - ' + DienGiai As F1 FROM NguonNhapXuat ORDER BY SoHieu", CboKho(2)
    
    baocao = 0
    
    Pdelay = GetSetting(IniPath, "Environment", "PrinterDelay", 10)
    
    For chi_so = 0 To 2
        InitDateVars MedNgay(chi_so), ngay(chi_so)
    Next
            
    For chi_so = 1 To 12
        SoLieu(chi_so) = False
    Next
        
    Caption = Caption + " - " + CStr(pNamTC)
    
    OptBC(35).Enabled = KHDetail
    OptBC(36).Enabled = KHDetail
    txtshkh(0).Enabled = KHDetail
    cmdkh(0).Enabled = KHDetail
    txtshkh(1).Enabled = KHDetail
    cmdkh(1).Enabled = KHDetail
    Int_RecsetToCbo "SELECT DoituongCT.MaSo As F2,(IIF(DoituongCT.MaKhachHang>0,KhachHang.Ten+' - '+DoituongCT.Sohieu+' - ','')+DienGiai+IIF(DoituongCT.MaKhachHang>0,' - Kyù ngaøy: '+ Format(NgayKy,'dd/mm/yy'),'')) As F1 FROM DoituongCT LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo ORDER BY  KhachHang.Ten,DoituongCT.SoHieu,DienGiai", CboHD
    
    KiemTraUser
    
    NLB = GetSetting(IniPath, "Environment", "SReport1", "")
    KTT = GetSetting(IniPath, "Environment", "SReport2", "")
    GD = GetSetting(IniPath, "Environment", "SReport3", "")
    
    Frame(6).Visible = pSongNgu
    
    StopPrint = True
    If pNN = 1 Then OptNN(1).Value = True
    StopPrint = False
    
    OptBC(72).Visible = (pCongNoHD > 0)
    OptBC(73).Visible = (pCongNoHD > 0)
    
    For chi_so = 1 To pSoVV
        ChkTT(chi_so - 1).Visible = True
        CboTT(chi_so - 1).Visible = True
        Int_RecsetToCbo "SELECT MaSo As F2,DienGiai As F1 FROM DoituongCT" + CStr(chi_so) + " ORDER BY DoituongCT" + CStr(chi_so) + ".DienGiai", CboTT(chi_so - 1)
    Next
    
    If pDTTP = 0 Then
        For chi_so = 0 To 1
            OptBC(59 + chi_so).Visible = False
            TxtCT(chi_so).Visible = False
            cmdct(chi_so).Visible = False
        Next
        OptBC(77).Visible = False
    End If
    
    If pVersion = 3 Then
        OptBC(10).Visible = False
        OptBC(11).Visible = False
        OptBC(50).Visible = False
        OptBC(51).Visible = False
        OptBC(54).Visible = False
    End If
        
    SetFont Me
      Dim ngaytt As Date
 ngaytt = "01/01/" + Mid(str(pNamTC), 4, 5)
 MedNgay(1).Text = ngaytt
 MedNgay(2).Text = ngaytt

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i
For i = 0 To 3
   If (OptLoai(i)) Then
           OptLoai(i).BackColor = 8438015
        Else
        OptLoai(i).BackColor = &H80FF80
        End If
 Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopPrint = True
    
    Set xlapp = Nothing
    Set xlsheet = Nothing
End Sub

Private Sub MedNgay_GotFocus(Index As Integer)
   ' AutoSelect MedNgay(Index)
End Sub

Private Sub MedNgay_LostFocus(Index As Integer)
    If IsDate(MedNgay(Index).Text) Then
        ngay(Index) = CDate(MedNgay(Index).Text)
    Else
        RFocus MedNgay(Index)
    End If
    If Index = 1 Then
    MedNgay(2).Text = MedNgay(1).Text
    End If
End Sub

Private Sub OptBc_Click(Index As Integer)
Dim thg_dau As Boolean
Dim thg_cuoi As Boolean
    
    Select Case Index
        Case 0, 4, 12, 14, 15, 16, 34, 35, 36, 37, 38, 39, 42, 43, 44, 46, 47, 48, 49, 52, 55, 60, 63, 69, 72, 73, 74, 75, 77, 78, 80, 81, 82, 83:
            OptTG(1).Enabled = True
        Case Else
            OptTG_Click 0
            OptTG(0).Value = True
            OptTG(1).Enabled = False
    End Select
    
    Frame(4).Visible = OptBC(34).Value Or OptBC(37).Value
    If OptBC(34).Value Or OptBC(37).Value Then
        OptDT(1).Enabled = OptBC(34).Value
        OptDT(2).Enabled = OptBC(34).Value
        ChkDu(13).Visible = (pChietKhau > 0 And OptBC(37).Value)
    End If
    
    baocao = Index
    ChkDu(0).Visible = (Index = 0)
    txtShTk(5).Visible = (Index = 0)
    CmdTK(5).Visible = (Index = 0)
    LbTenTk(5).Visible = (Index = 0)
    
    If Index = 8 Or Index = 9 Then ChkDu(6).Value = 0
    If Index = 3 Then ChkDu(6).Value = 0
    
    Label(6).Visible = (Index = 2)
    txtShCT(1).Visible = (Index = 2)
    
    Select Case Index
          Case 18:
                TenFile = "KHAUHAO.RPT"
                TenBaoCao = "Sæ khÊu hao tµi s¶n cè ®Þnh"
                thg_dau = True: thg_cuoi = True
          Case 61:
                TenFile = "KHAUHAO2.RPT"
                TenBaoCao = "Sæ khÊu hao tæng hîp"
                thg_dau = True: thg_cuoi = True
          Case 19:
                TenFile = "SODU.RPT"
                TenBaoCao = "Sè d­ ®Çu kú"
                thg_dau = True: thg_cuoi = False
          Case 20:
                TenFile = "SODU.RPT"
                TenBaoCao = "Sè d­ cuèi kú"
                thg_dau = True: thg_cuoi = False
          Case 21:
                TenFile = "TANGGIAM.RPT"
                TenBaoCao = "T¨ng tµi s¶n cè ®Þnh"
                thg_dau = True: thg_cuoi = True
          Case 22:
                TenFile = "DANHGIA.RPT"
                TenBaoCao = "§¸nh gi¸ l¹i t¨ng"
                thg_dau = True: thg_cuoi = True
          Case 25:
                TenFile = "CHITIET.RPT"
                TenBaoCao = "Chi tiÕt ph©n bæ khÊu hao"
                thg_dau = True: thg_cuoi = True
          Case 26:
                TenFile = "TONGHOP.RPT"
                TenBaoCao = "Tæng hîp ph©n bæ khÊu hao"
                thg_dau = True: thg_cuoi = True
          Case 27:
                TenFile = "NANGLUC.RPT"
                TenBaoCao = "N¨ng lùc tµi s¶n cè ®Þnh"
                thg_dau = True: thg_cuoi = False
          Case 28:
                TenFile = "BIENDONG.RPT"
                TenBaoCao = "Tæng hîp biÕn ®éng trong kú"
                thg_dau = True: thg_cuoi = True
          Case 29:
                TenFile = "BIENDONG.RPT"
                TenBaoCao = "Tæng hîp biÕn ®éng luü kÕ"
                thg_dau = False: thg_cuoi = True
          Case 30:
                TenFile = "MIENTRU.RPT"
                TenBaoCao = "Chi tiÕt tµi s¶n kh«ng ph¶i tÝnh khÊu hao"
                thg_dau = True: thg_cuoi = False
          Case 31:
                TenFile = "BINHQUAN.RPT"
                TenBaoCao = "B×nh qu©n tµi s¶n cè ®Þnh cÇn tÝnh khÊu hao"
                thg_dau = True: thg_cuoi = True
          Case 23:
                TenFile = "TANGGIAM.RPT"
                TenBaoCao = "Gi¶m tµi s¶n"
                thg_dau = True: thg_cuoi = True
          Case 24:
                TenFile = "DANHGIA.RPT"
                TenBaoCao = "§¸nh gi¸ l¹i gi¶m"
                thg_dau = True: thg_cuoi = True
          Case 19, 20, 23, 24:
                thg_dau = False: thg_cuoi = False
          Case 22:
                thg_dau = False: thg_cuoi = True
          Case 45:
                TenFile = "SOTSCD.RPT"
                TenBaoCao = "Sæ tµi s¶n cè ®Þnh"
                thg_dau = True: thg_cuoi = False
          Case 68:
                TenFile = "SOTSCD2.RPT"
                TenBaoCao = "Sæ tµi s¶n cè ®Þnh"
                thg_dau = True: thg_cuoi = False
          Case 79:
                TenFile = "PBKH.RPT"
                TenBaoCao = "B¶ng TÝnh vµ Ph©n bæ khÊu hao tµi s¶n cè ®Þnh"
                thg_dau = True: thg_cuoi = False
          Case Else
                thg_dau = True: thg_cuoi = True
    End Select
    If thg_dau = True Then
          Label(0).Enabled = True
          CboThang(0).Enabled = True
    Else
          Label(0).Enabled = False
          CboThang(0).Enabled = False
    End If
    If thg_cuoi = True Then
          Label(1).Enabled = True
          CboThang(1).Enabled = True
    Else
          Label(1).Enabled = False
          CboThang(1).Enabled = False
    End If
      ' in cac muc them
    If Index = 106 Then
    ' OptTG(1).Value = True
     OptTG(1).Enabled = True
    ' ChkDu(0).Visible = True
    ChkDu(0).Visible = False
    End If
    If Index = 0 Then
     ChkDu(0).Visible = False
     txtShTk(5).Visible = False
     CmdTK(5).Visible = False
    End If
    
End Sub

Private Sub OptBC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If pRpt = 1 And Button = 2 And User_Right = 0 Then
        FU1.QuyenBCChitiet Index, OptBC(Index).Caption
    End If
End Sub

Private Sub OptLoai_Click(Index As Integer)
Dim i

For i = 0 To 3
        OptLoai(i).BackColor = &H80FF80
 Next
      
      If (OptLoai(Index) = True) Then
      SSTab.Tab = Index
      OptLoai(Index).BackColor = 8438015
      End If
    
End Sub

Private Sub OptLoai_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i
For i = 0 To 3
   If (OptLoai(i)) Then
      '     OptLoai(i).BackColor = 8438015
        Else
        OptLoai(i).BackColor = &H80FF80
        End If
 Next
  OptLoai(Index).BackColor = 8438015
End Sub

Private Sub OptNN_Click(Index As Integer)
    nn = Index
    If StopPrint Then Exit Sub
    SetFont Me, 1
End Sub

Private Sub OptTG_Click(Index As Integer)
    Frame(1 - Index).Visible = False
    Frame(Index).Visible = True
End Sub


Private Sub SSTab_Click(PreviousTab As Integer)
If SSTab.Tab = 0 Then
OptBC(100).Value = True
End If
If SSTab.Tab = 1 Then
OptBC(14).Value = True
End If
If SSTab.Tab = 2 Then
OptBC(58).Value = True
End If
If SSTab.Tab = 3 Then
OptBC(18).Value = True
End If
'  If PreviousTab = 1 Then
'
'  ElseIf PreviousTab = 0 Then
'    OptBC(14).Value = True
' ElseIf PreviousTab = 2 Then
'    OptBC(18).Value = True
'    Else
'     OptBC(58).Value = True
'  End If
'
End Sub

Private Sub TxtCT_GotFocus(Index As Integer)
    AutoSelect TxtCT(Index)
End Sub

Private Sub TxtCT_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtCT(Index).Text = FrmTP.ChonTP(TxtCT(Index).Text)
        RFocus TxtCT(Index)
    End If
End Sub

Private Sub TxtCT_LostFocus(Index As Integer)
    Dim m As Long
    
    TxtCT(Index).Text = UCase(TxtCT(Index).Text)
    LbCT(Index).Caption = SelectSQL("SELECT TenVattu AS F1,MaSo AS F2 FROM TP154 WHERE SoHieu='" + TxtCT(Index).Text + "'", m)
    TxtCT(Index).tag = m
End Sub

Private Sub txtHan_GotFocus()
    AutoSelect txtHan
End Sub

Private Sub txtHan_KeyPress(KeyAscii As Integer)
    KeyProcess txtHan, KeyAscii
End Sub

Private Sub txtLS_GotFocus(Index As Integer)
    AutoSelect txtLS(Index)
End Sub

Private Sub txtLS_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyProcess txtLS(Index), KeyAscii
End Sub

Private Sub txtsh_GotFocus(Index As Integer)
    AutoSelect txtsh(Index)
End Sub

Private Sub txtsh_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmd_Click 0
End Sub

Private Sub txtsh_LostFocus(Index As Integer)
    Dim ms As Long
    
    Lb(Index).Caption = TenNV(txtsh(Index).Text, ms)
    txtsh(Index).tag = ms
End Sub

Private Sub txtshkh_GotFocus(Index As Integer)
    AutoSelect txtshkh(Index)
End Sub

Private Sub txtshkh_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdkh_Click Index
End Sub

Public Sub txtshkh_LostFocus(Index As Integer)
    Dim mvt As Long
    
    txtshkh(Index).Text = UCase(txtshkh(Index).Text)
    lbkh(Index).Caption = TenKH(txtshkh(Index).Text, mvt)
    txtshkh(Index).tag = mvt
    HienThongBao Me.Caption, 1
End Sub

Private Sub txtShTk_GotFocus(Index As Integer)
    AutoSelect txtShTk(Index)
End Sub

Private Sub txtShTk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdtk_Click Index
End Sub
'==============================================================================================================
' KiÓm tra sè hiÖu tµi kho¶n ®· nhËp
'==============================================================================================================
Public Sub txtShTk_LostFocus(Index As Integer)
    Dim mtk As Long
        
    txtShTk(Index).Text = UCase(txtShTk(Index).Text)
    
    mtk = 0
    LbTenTk(Index).Caption = tentk(txtShTk(Index).Text, mtk)
    txtShTk(Index).tag = mtk
            
    If Index = 2 Then
        Int_RecsetToCbo "SELECT HethongTK.MaNT As F2,NguyenTe.KyHieu As F1 FROM HethongTK INNER JOIN" _
            & " NguyenTe ON HethongTK.MaNT = NguyenTe.MaSo WHERE HethongTK.SoHieu = '" + txtShTk(2).Text _
            + "' ORDER BY NguyenTe.KyHieu", CboNT

        If CboNT.ListCount > 0 Then
            CboNT.AddItem pTienStr, 0
            CboNT.ItemData(0) = 0
        End If
        
        If CboNT.ListCount = 2 Then
            CboNT.AddItem CboNT.List(0) + "+" + CboNT.List(1)
            CboNT.ItemData(CboNT.NewIndex) = -1
        End If
    End If
    
    HienThongBao Me.Caption, 1
End Sub
'============================================================================================================
' Thu tuc in bang tong hop so du chi tiet
' Thang:  Thang in bao cao
' Dest:   In ra m¤y in/Man hinh
' RptObj: Report CrystalReport
'============================================================================================================
Private Sub InBangThop(tdau As Integer, tcuoi As Integer, taikhoan As ClsTaikhoan, cap As Integer, ct As Integer)
    Dim st As String, sqln As String, sqlc As String, i As Integer
    Dim sqlnx As String, sqlcx As String
    
    GauGe.Max = 3
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + No_" + st
        sqlc = sqlc + " + Co_" + st
    Next
    For i = 1 To CThangDB(tcuoi)
        st = CStr(i)
        sqlnx = sqlnx + " + No_" + st
        sqlcx = sqlcx + " + Co_" + st
    Next
    GauGe.Value = 1
    st = CStr(CThangDB(tcuoi))
    If taikhoan.MaSo > 0 Then
        SetSQL "QTongHopCT", "SELECT DISTINCTROW SoHieu, First(Cap) AS FirstOfCap, First(Ten" + IIf(nn > 0, "E", "") + ") AS FirstOfTen, First(Kieu) AS FirstOfKieu, First(Loai) AS FirstOfLoai, First(MaTC = MaSo OR MaTC = 0) AS FirstOfChuanLa, First(TKCon) AS FirstOfTKCon, First(CapDuoi) AS CD," _
            & " Sum(DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkNo, Sum(DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkCo, Sum(" + sqln + ") AS PsNo, Sum(" + sqlc + ") AS PsCo, Sum(" + sqlnx + ") AS PsNoX, Sum(" + sqlcx + ") AS PsCoX, Sum(DuNo_" + st + ") AS CkNo, Sum(DuCo_" + st + ") AS CkCo" _
            & " FROM HeThongTK WHERE (Cap > " + CStr(taikhoan.cap) + " AND Cap<=" + CStr(cap) + ") GROUP BY SoHieu Having ((First(Loai)>0) AND (SoHieu LIKE '" + IIf(Left(taikhoan.sohieu, 1) <> "L", taikhoan.sohieu, Right(taikhoan.sohieu, 1)) + "*')) ORDER BY SoHieu"
    Else
        SetSQL "QTongHopCT", "SELECT DISTINCTROW SoHieu, First(Cap) AS FirstOfCap, First(Ten" + IIf(nn > 0, "E", "") + ") AS FirstOfTen, First(Kieu) AS FirstOfKieu, First(Loai) AS FirstOfLoai, First(MaTC = MaSo OR MaTC = 0) AS FirstOfChuanLa, First(TKCon) AS FirstOfTKCon, First(CapDuoi) AS CD," _
            & " Sum(DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkNo, Sum(DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkCo, Sum(" + sqln + ") AS PsNo, Sum(" + sqlc + ") AS PsCo, Sum(" + sqlnx + ") AS PsNoX, Sum(" + sqlcx + ") AS PsCoX, Sum(HeThongTK.DuNo_" + st + ") AS CkNo, Sum(HeThongTK.DuCo_" + st + ") AS CkCo" _
            & " From HeThongTK WHERE  Cap<=" + CStr(cap) + " GROUP BY SoHieu Having First(Loai)>0 AND First(Cap > 0)" _
            + IIf(ct = 0, " AND (Sum(DuNo_" + st + ")<>0 OR Sum(DuCo_" + st + ")<>0 OR Sum(" + sqln + ")<>0 OR Sum(" + sqlc + ")<>0 OR Sum(" + sqlnx + ")<>0 OR Sum(" + sqlcx + ")<>0)", "") + " ORDER BY SoHieu"
    End If
    GauGe.Value = 2
    XDCapTK IIf(taikhoan.MaSo > 0, taikhoan.sohieu, "")
    frmMain.Rpt.ReportFileName = "THSDCT.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    
    If taikhoan.MaSo > 0 Then
        frmMain.Rpt.Formulas(5) = "TenBaoCao = IF PageNumber() = 1 THEN '" + ABCtoVNI("Tµi kho¶n: ") + taikhoan.sohieu + " - " + IIf(nn > 0, taikhoan.TenE, taikhoan.Ten) + "'"
        frmMain.Rpt.Formulas(6) = "Cap=" + CStr(taikhoan.cap + 1)
    End If
    GauGe.Value = 3
End Sub
'=============================================================================================================
' Thu tuc in bang can doi thu
' Thang:  Thang in bao cao
' Dest:   In ra m¤y in/M¡n hÒnh
' RptObj: Report CrystalReport
'=============================================================================================================
Private Sub InCdt(saukc As Boolean, tdau As Integer, tcuoi As Integer, cap As Integer, ct As Integer)
    Dim st As String, sqln As String, sqlc As String, i As Integer
    
    GauGe.Max = 3
    SoDuTKCN2 ThangTruoc(tdau)
    SoDuTKCN2 tcuoi
    If Not saukc Then XacdinhKC tdau, tcuoi
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + HethongTk.No_" + st
        sqlc = sqlc + " + HethongTk.Co_" + st
    Next
    st = CStr(CThangDB(tcuoi))
    GauGe.Value = 1
    SetSQL "QCdt", "SELECT DISTINCTROW HeThongTK.SoHieu, HeThongTK.Cap, HeThongTK.Ten" + IIf(nn > 0, "E", "") + "," _
        & " HeThongTK.Kieu, HeThongTK.Loai, HeThongTK.DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + " As DkNo, HeThongTK.DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + " As DkCo, " + sqln + " As PsNo, " + sqlc + " As PsCo," _
        & " KC_N, KC_C, " + IIf(saukc, "HeThongTK.DuNo_" + st, "HethongTk.DuNo") + " As CkNo, " + IIf(saukc, "HeThongTK.DuCo_" + st, "HethongTk.DuCo") _
        + " As CkCo From HeThongTK Where ((HeThongTK.MaTC = 0 Or HethongTK.MaTC = HethongTK.MaSo) OR (TK_ID3 Mod 10 >= 1)) And (HeThongTK.Loai > 0)  AND Cap<=" + CStr(cap) _
        + IIf(ct = 0, " AND (DuNo_" + st + "<>0 OR DuCo_" + st + "<>0 OR (" + sqln + ")<>0 OR (" + sqlc + ")<>0)", "")
    GauGe.Value = 2
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    If saukc Then
        frmMain.Rpt.Formulas(5) = "SauKc = TRUE"
    Else
        frmMain.Rpt.Formulas(5) = "SauKc = FALSE"
    End If
        
    frmMain.Rpt.ReportFileName = "CDT.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    GauGe.Value = 3
End Sub

Private Function InSoChitietNT(taikhoan As ClsTaikhoan, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, loai As Long, thongbao As Boolean) As Boolean
    Dim dkn As Double, dkc As Double, dknt As Double, psn As Double, psc As Double, sqlw As String, ps As Boolean, m2 As Long, psn2 As Double, psc2 As Double
        
With taikhoan
    If tcuoi > 0 Then
        ps = .TkCoPS(tdau, tcuoi)
        If loai >= 0 Then
            .SoDuTK ThangTruoc(tdau), dkn, dkc, dknt, 1
        Else
            sqlw = CStr(CThangDB(ThangTruoc(tdau)))
            dkn = SelectSQL("SELECT Sum(DuNo_" + sqlw + ") AS F1,Sum(DuCo_" + sqlw + ") AS F2,Sum(DuNT_" + sqlw + ") AS F3 FROM HethongTK WHERE SoHieu='" + .sohieu + "'", dkc, dknt)
            m2 = MaTKNguyenTe(.sohieu, -1)
            ps = ps Or (SelectSQL("SELECT MaCT AS F1 FROM ChungTu WHERE (MaTKNo=" + CStr(m2) + " OR MaTKCo=" + CStr(m2) + ") AND " + WThang("ThangCT", tdau, tcuoi)) > 0)
        End If
        sqlw = WThang("ThangCT", tdau, tcuoi)
    Else
        ps = .TkCoPSN(ndau, ncuoi)
        If loai >= 0 Then
            .SoDuNgay ndau - 1, dkn, dkc, dknt
        Else
            SoDuTKNgaySH .sohieu, ndau - 1, dkn, dkc, dknt
            m2 = MaTKNguyenTe(.sohieu, -1)
            ps = ps Or (SelectSQL("SELECT MaCT AS F1 FROM ChungTu WHERE (MaTKNo=" + CStr(m2) + " OR MaTKCo=" + CStr(m2) + ") AND " + WNgay("ThangCT", ndau, ncuoi)) > 0)
        End If
        sqlw = WNgay("NgayGS", ndau, ncuoi)
    End If
    ' In bao cao
    If dkn <> 0 Or dkc <> 0 Or dknt <> 0 Or ps Then
        GauGe.Max = 3
        GauGe.Value = 1
        
        If ps Then
            SetSQL "QChitiet", "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.SoPS2No, ChungTu.SoPS2Co, ChungTu.GhiChu, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.MaTKNo, ChungTu.MaTKCo" _
                & " FROM HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo" _
                & " WHERE ((chungtu.MaTkNo = " + CStr(taikhoan.MaSo) + ") Or (chungtu.MaTkCo = " + CStr(taikhoan.MaSo) + ")" + IIf(m2 > 0, " OR (chungtu.MaTkNo = " + CStr(m2) + ") Or (chungtu.MaTkCo = " + CStr(m2) + ")", "") + ") And " + sqlw _
                & " ORDER BY Chungtu.NgayGS, Chungtu.MaCT"
        Else
            SetSQL "QChitiet", "SELECT DISTINCTROW 0 AS MaCT, '' AS SoHieu, Null AS NgayCT, Null AS NgayGS, '' AS DienGiai, 0 AS SoPS, 0 AS SoPS2No, 0 AS SoPS2Co, '' AS GhiChu, '' AS SoHieu1, '' AS SoHieu2, 0 AS MaTKNo, 0 AS MaTKCo" _
                & " FROM ChungTu WHERE (chungtu.MaCT = 0)"
        End If
        GauGe.Value = 2
        
        frmMain.Rpt.Formulas(3) = "MaTk=" + CStr(.MaSo)
        frmMain.Rpt.Formulas(30) = "MaTk2=" + CStr(m2)
        frmMain.Rpt.Formulas(4) = "NoDk=" + DoiDau(dkn - dkc)
        frmMain.Rpt.Formulas(5) = "NoDkNT=" + DoiDau(dknt)
        frmMain.Rpt.Formulas(6) = "Kieu=" + CStr(.kieu)
        
        InSoChitietNT = True
        
        If CboNT.ListIndex < CboNT.ListCount - 1 Then
            psn = SoPSTKMaSo(taikhoan.MaSo, pThangDauKy, tcuoi, -1, psn2)
            psc = SoPSTKMaSo(taikhoan.MaSo, pThangDauKy, tcuoi, 1, psc2)
        Else
            psn = SoPSTK(taikhoan.sohieu, pThangDauKy, tcuoi, -1, psn2)
            psc = SoPSTK(taikhoan.sohieu, pThangDauKy, tcuoi, 1, psc2)
        End If
        
        frmMain.Rpt.Formulas(12) = "NoLK = " + DoiDau(psn)
        frmMain.Rpt.Formulas(13) = "CoLK = " + DoiDau(psc)
        frmMain.Rpt.Formulas(14) = "NoLK2 = " + DoiDau(psn2)
        frmMain.Rpt.Formulas(15) = "CoLK2 = " + DoiDau(psc2)
        
        frmMain.Rpt.Formulas(7) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
        RptSetDate IIf(tdau > 0, NgayCuoiThang(pNamTC, tcuoi), ncuoi), nn
        If Left(taikhoan.sohieu, 3) = "112" And tcuoi = 0 Then
            frmMain.Rpt.Formulas(9) = "TenTk = IF PageNumber() = 1 THEN '" + taikhoan.GhiChu + "'"
            frmMain.Rpt.Formulas(11) = "LTG = IF PageNumber() = 1 THEN ': " + ABCtoVNI("Lo¹i tiÒn: ") + CboNT.Text + "'"
            frmMain.Rpt.ReportFileName = "SOTGNT.RPT"
        Else
            frmMain.Rpt.Formulas(9) = "TenTk = IF PageNumber() = 1 THEN '" + .sohieu + " - " + .Ten + " - " + TenNT(taikhoan.MaNT) + "'"
            frmMain.Rpt.ReportFileName = "CTNGTE.RPT"
        End If
        GauGe.Value = 3
    Else
        If thongbao Then ErrMsg er_KoPS1
        InSoChitietNT = False
    End If
End With
End Function
'===================================================================================================
' Thu tuc in bang tong hop so du chi tiet nguyªn tÖ
' Thang:  Thang in bao cao
' Dest:   In ra m¤y in/Man hinh
' RptObj: Report CrystalReport
'===================================================================================================
Private Sub InBangThopNT(tdau As Integer, tcuoi As Integer)
    Dim st As String, sqln As String, sqlc As String, sqlnnt As String, sqlcnt As String, i As Integer
    
    GauGe.Max = 3
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + HethongTk.No_" + st
        sqlc = sqlc + " + HethongTk.Co_" + st
        sqlnnt = sqlnnt + " + HethongTk.No_" + st + "_NT"
        sqlcnt = sqlcnt + " + HethongTk.Co_" + st + "_NT"
    Next
    
    GauGe.Value = 1
    st = CStr(CThangDB(tcuoi))
    SetSQL "QTongHopCTNT", "SELECT DISTINCTROW HeThongTK.SoHieu, HeThongTK.Ten" + IIf(nn > 0, "E", "") + ", HeThongTK.MaNT, NguyenTe.KyHieu," _
    & " HeThongTK.DuNo_" + CStr(CThangDB(ThangTruoc(tdau))) + " AS DkNo, HeThongTK.DuCo_" + CStr(CThangDB(ThangTruoc(tdau))) + " AS DkCo, HeThongTK.DuNT_" + CStr(CThangDB(ThangTruoc(tdau))) + " AS DkNT," _
    + sqln + " AS PsNo, " + sqlnnt + " AS PsNoNT, " + sqlc + " AS PsCo, " + sqlcnt + " AS PsCoNT," _
    & " HeThongTK.DuNo_" + st + " AS CkNo, HeThongTK.DuCo_" + st + " AS CkCo, HeThongTK.DuNT_" + st + " AS CkNT" _
    & " FROM NguyenTe RIGHT JOIN HeThongTK ON NguyenTe.MaSo = HeThongTK.MaNT" _
    & " Where ((HeThongTK.Loai > 0) And (HeThongTK.mant <> 0) And (" + sqln + "<>0 OR " + sqlc + "<>0 OR HeThongTK.DuNo_" _
    + st + "<>0 OR HeThongTK.DuCo_" + st + "<>0)) ORDER BY HeThongTK.SoHieu"
    GauGe.Value = 2
    frmMain.Rpt.ReportFileName = "THSDCTNT.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
End Sub

Private Sub txtShTS_GotFocus()
    AutoSelect txtShTS
End Sub

Private Sub txtShVT_GotFocus(Index As Integer)
    AutoSelect txtShVT(Index)
End Sub

Private Sub txtShVT_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then cmdvt_Click Index
End Sub

Public Sub txtShVT_LostFocus(Index As Integer)
    Dim mvt As Long, dv As String
    
    txtShVT(Index).Text = UCase(txtShVT(Index).Text)
    LbTenVT(Index).Caption = TenVT(txtShVT(Index).Text, mvt, dv)
    txtShVT(Index).tag = mvt
    HienThongBao Me.Caption, 1
End Sub
'===================================================================================================
' Thñ tôc lËp b¶ng kª lòy kÕ nhËp xuÊt tån
'===================================================================================================
Private Sub InLuyKe(mkho As Long, shtk As String, mloai As Long, tdau As Integer, tcuoi As Integer, Optional T As Boolean = True, Optional loailk As Integer = 0)
    Dim sql As String, st As String, sqlln As String, sqllx As String, sqltn As String, sqltx As String, i As Integer, sqltn2 As String, sqltx2 As String
    
    GauGe.Max = 3
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqlln = sqlln + " + TonKho.Luong_Nhap_" + st
        sqllx = sqllx + " + TonKho.Luong_Xuat_" + st
        sqltn = sqltn + " + TonKho.Tien_Nhap_" + st
        sqltx = sqltx + " + TonKho.Tien_Xuat_" + st
        sqltn2 = sqltn2 + " + TonKho.USDTien_Nhap_" + st
        sqltx2 = sqltx2 + " + TonKho.USDTien_Xuat_" + st
    Next
    GauGe.Value = 1
    st = CStr(CThangDB(tcuoi))
    XDTyLeQD 0
    If mkho > 0 Then
    'neu dua , HethongTK.SoHieu  vao group no se tinh rieng tung tai khoan
        sql = "SELECT DISTINCTROW  Sum(TonKho.Tien_" + st + ")  AS SHTK, First(HeThongTK.Ten" + IIf(nn > 0, "E", "") + ") AS TenTK, First(PhanLoaiVattu.PLCha) AS PLCha3," _
            & " First(PhanLoaiVattu.SoHieu) AS SHPL3, First(PhanLoaiVattu.TenPhanLoai) AS TenPL3, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, Sum(TonKho.Tien_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkTien," _
            & " Sum(TonKho.Luong_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkLuong, Sum(" + sqltn + ") AS TienNhap, Sum(" + sqlln + ") AS LuongNhap, Sum(" + sqltx + ") AS TienXuat," _
            & " Sum(" + sqllx + ") AS LuongXuat, Sum(TonKho.Luong_" + st + ") AS CkLuong, Sum(TonKho.Tien_" + st + ") AS CkTien" _
            + IIf(pGiaUSD > 0, ", Sum(TonKho.USDTien_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkUSD, Sum(" + sqltn2 + ") AS TienNhapUSD, Sum(" + sqltx2 + ") AS TienXuatUSD, Sum(TonKho.USDTien_" + st + ") AS CkUSD", "") _
            + " FROM HeThongTK RIGHT JOIN ((PhanLoaiVattu RIGHT JOIN Vattu ON PhanLoaiVattu.MaSo = Vattu.MaPhanLoai) RIGHT JOIN TonKho ON Vattu.MaSo = TonKho.MaVatTu) ON HeThongTK.MaSo = TonKho.MaTaiKhoan" _
            & " GROUP BY Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi " _
            + IIf(mkho > 0, ", TonKho.MaSoKho", "") + IIf(mloai > 0, ", PhanLoaiVattu.SoHieu", "") _
            + " Having (Sum(TonKho.Luong_" + st + ")<>0 OR Sum(TonKho.Tien_" + st + ")<>0 OR SUM(" + sqlln + ")<>0 OR SUM(" + sqltn + ")<>0 OR SUM(" + sqllx + ")<>0 OR SUM(" + sqltx + ")<>0" + IIf(pGiaUSD > 0, " OR Sum(TonKho.USDTien_" + st + ")<>0 OR SUM(" + sqltn2 + ")<>0 OR SUM(" + sqltx2 + ")<>0)", ")") _
            + IIf(mkho <> 0, " AND TonKho.MaSoKho = " + CStr(mkho), "") + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") _
            + IIf(mloai <> 0, " AND PhanLoaiVattu.SoHieu LIKE '" + MaSo2SoHieu(mloai, "PhanLoaiVattu") + "*'", "") + " ORDER BY Vattu.SoHieu"
    Else
        ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
        ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,MaCha,BC_ID,MK,Kq1,Kq2) SELECT First(MaSo),First(Cstr(MaSo)),MaTKNo,MaVattu,MaKho,Sum(SoPS2No),Sum(SoPS) " _
            & "FROM ChungTu WHERE MaLoai=4 AND MaTKNo=MaTKCo AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY MaKho,MaTKNo,MaVattu"
        sql = "SELECT DISTINCTROW First(HeThongTK.SoHieu) AS SHTK, First(HeThongTK.Ten" + IIf(nn > 0, "E", "") + ") AS TenTK, First(PhanLoaiVattu.PLCha) AS PLCha3," _
            & " First(PhanLoaiVattu.SoHieu) AS SHPL3, First(PhanLoaiVattu.TenPhanLoai) AS TenPL3, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, Sum(TonKho.Tien_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkTien," _
            & " Sum(TonKho.Luong_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkLuong, Sum(" + sqltn + "-iif(isnull(Kq2),0,Kq2)) AS TienNhap, Sum(" + sqlln + "-iif(isnull(Kq1),0,Kq1)) AS LuongNhap, Sum(" + sqltx + "-iif(isnull(Kq2),0,Kq2)) AS TienXuat," _
            & " Sum(" + sqllx + "-iif(isnull(Kq1),0,Kq1)) AS LuongXuat, Sum(TonKho.Luong_" + st + ") AS CkLuong, Sum(TonKho.Tien_" + st + ") AS CkTien" _
            + IIf(pGiaUSD > 0, ", Sum(TonKho.USDTien_" + CStr(CThangDB(ThangTruoc(tdau))) + ") AS DkUSD, Sum(" + sqltn2 + ") AS TienNhapUSD, Sum(" + sqltx2 + ") AS TienXuatUSD, Sum(TonKho.USDTien_" + st + ") AS CkUSD", "") _
            & " FROM (HeThongTK RIGHT JOIN ((PhanLoaiVattu RIGHT JOIN Vattu ON PhanLoaiVattu.MaSo = Vattu.MaPhanLoai) RIGHT JOIN TonKho ON Vattu.MaSo = TonKho.MaVatTu) ON HeThongTK.MaSo = TonKho.MaTaiKhoan) LEFT JOIN BaoCaoCP2 ON TonKho.MaSoKho=BaoCaoCP2.MK AND TonKho.MaTaiKhoan=BaoCaoCP2.MaCha AND TonKho.MaVattu=BaoCaoCP2.BC_ID" _
            & " GROUP BY Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, HethongTK.SoHieu " _
            + IIf(mloai > 0, ", PhanLoaiVattu.SoHieu", "") _
            + " Having Sum(TonKho.Luong_" + st + ")<>0 OR Sum(TonKho.Tien_" + st + ")<>0 OR SUM(" + sqlln + ")<>0 OR SUM(" + sqltn + ")<>0 OR SUM(" + sqllx + ")<>0 OR SUM(" + sqltx + ")<>0" + IIf(pGiaUSD > 0, " OR Sum(TonKho.USDTien_" + st + ")<>0 OR SUM(" + sqltn2 + ")<>0 OR SUM(" + sqltx2 + ")<>0", "") _
            + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") _
            + IIf(mloai <> 0, " AND PhanLoaiVattu.SoHieu LIKE '" + MaSo2SoHieu(mloai, "PhanLoaiVattu") + "*'", "") + " ORDER BY Vattu.SoHieu"
    End If
    SetSQL "QChitiet", sql
    sql = "SELECT QChitiet.*,PhanLoaiVattu.SoHieu AS SHPL2,TenPhanLoai AS TenPL2,PhanLoaiVattu.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoaiVattu ON QChitiet.PLCha3=PhanLoaiVattu.MaSo"
    SetSQL "QDuPhong", sql
    sql = "SELECT QDuPhong.*,PhanLoaiVattu.SoHieu AS SHPL1,TenPhanLoai AS TenPL1 FROM QDuPhong LEFT JOIN PhanLoaiVattu ON QDuPhong.PLCha2=PhanLoaiVattu.MaSo ORDER BY QDuPhong.Sohieu"
    SetSQL "QLuyKe", sql
    
   Dim thong_bao
   thong_bao = ""
   Dim rs_chungtu As Recordset
   sql = "select * from QLuyKe where (ckluong < 0 or cktien < 0) or (ckluong <= 0 and  cktien > 0) "
   Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs_chungtu.EOF
        thong_bao = thong_bao + rs_chungtu!sohieu + " : SL " + str(rs_chungtu!ckluong) + "  Tong tien: " + Format(rs_chungtu!cktien, Mask_2) + Chr(13)
     rs_chungtu.MoveNext
    Loop
       rs_chungtu.Close
    If (Len(thong_bao) > 0) Then
    thong_bao = "B¹n cÇn kiÓm tra l¹i s¶n phÈm: " + Chr(13) + thong_bao
        MsgBox (thong_bao)
    End If

    
    GauGe.Value = 2
    frmMain.Rpt.Formulas(3) = "TenKho='" + IIf(mkho > 0, "Kho: " + CboKho(0).Text, "") + "'"
    frmMain.Rpt.Formulas(4) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    If T And loailk <> 1 Then
        frmMain.Rpt.ReportFileName = "LUYKE" + IIf(pGiaUSD > 0, "X", "") + ".RPT"
    Else
        frmMain.Rpt.ReportFileName = "LUYKEL" + IIf(pGiaUSD > 0, "X", "") + ".RPT"
    End If
'    ExecuteSQL5 "update QLuyKe set SHTK = '111' "
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    GauGe.Value = 3
End Sub
'====================================================================================================
' Thñ tôc in vËt t­ nhËp xuÊt theo nguån
'====================================================================================================
Private Sub InNguonNhapXuat(mkho As Long, mtk As Long, mloai As Long, mn As Long, tdau As Integer, tcuoi As Integer, loai As Integer, dktg As Integer, ndau As Date, ncuoi As Date)
    Dim sql As String
    
    GauGe.Max = 2
    sql = "SELECT DISTINCTROW Chungtu.MaLoai,MaKho,MaNguon,NguonNhapXuat.MaSo, Vattu.MaSo, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, NguonNhapXuat.SoHieu, NguonNhapXuat.DienGiai, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.DienGiai, ChungTu.SoPS2No, ChungTu.SoPS2Co, ChungTu.SoPS" _
        & " FROM PhanLoaiVattu RIGHT JOIN (NguonNhapXuat RIGHT JOIN (Vattu INNER JOIN ChungTu ON Vattu.MaSo = ChungTu.MaVattu) ON NguonNhapXuat.MaSo = ChungTu.MaNguon) ON PhanLoaiVattu.MaSo = Vattu.MaPhanLoai" _
        & " Where ((" + IIf(loai = -1, "ChungTu.MaLoai =1 OR ChungTu.MaLoai =4", "ChungTu.MaLoai =2 OR ChungTu.MaLoai =4") + ") AND " + IIf(loai = -1, "MaTKNo>0", "MaTKCo>0") _
        + IIf(mkho <> 0, IIf(loai = -1, " AND ((Chungtu.MaKho = " + CStr(mkho) + " AND ChungTu.MaLoai=1) OR (MaNguon=" + CStr(mkho) + " AND Chungtu.MaLoai=4)", " AND (Chungtu.MaKho = " + CStr(mkho)) + ")", "") + IIf(mtk <> 0, " AND (" + IIf(loai = -1, "MaTKNo", "MaTKCo") + " = " + CStr(mtk) + ")", "") _
        + IIf(mn > 0, " AND (MaNguon=" + CStr(mn) + ")", "") + IIf(mloai > 0, " AND (PhanLoaiVattu.SoHieu LIKE '" + MaSo2SoHieu(mloai, "PhanLoaiVattu") + "*')", "") + ")"
    If dktg = 0 Then
        sql = sql + " AND " + WThang("ThangCT", tdau, tcuoi)
        frmMain.Rpt.Formulas(4) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    Else
        sql = sql + " AND " + WNgay("NgayGS", ndau, ncuoi)
        frmMain.Rpt.Formulas(4) = "ThoiGian='" + ThoiGianN(ndau, ncuoi, nn) + "'"
    End If
    sql = sql + " ORDER BY NguonNhapXuat.SoHieu, Vattu.SoHieu"
    ' And (chungtu.MaLoai = " + IIf(loai = -1, "1", "2") + "))"
    SetSQL "QNguonNhapXuat", sql
    GauGe.Value = 1
    frmMain.Rpt.Formulas(3) = "TenKho='" + IIf(mkho > 0, "Kho: " + CboKho(0).Text, "") + "'"
    frmMain.Rpt.Formulas(6) = "MaLoai=" + CStr(loai)
    frmMain.Rpt.ReportFileName = "NguonNX.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
End Sub
'====================================================================================================
' Thñ tôc in sæ chi tiÕt c¸c chi tiÕt nî qu¸ h¹n
'====================================================================================================
Private Function InCTNoQuaHan(ngayh As Date, songay As Integer, mpl As Long) As Boolean
    Dim rs_tk As Recordset, rs_ct As Recordset, duno As Double, duco As Double, st As String, i As Long, X As Double, Y As Double
    Dim dk As Boolean
    
    ExecuteSQL5 "DELETE FROM BaoCaoCP"
    i = CThangDB(Month(ngayh))
    Set rs_tk = DBKetoan.OpenRecordset("SELECT DISTINCTROW KhachHang.MaSo,KhachHang.SoHieu,KhachHang.Ten " _
        & " FROM ((KhachHang INNER JOIN SoDuKhachHang ON KhachHang.MaSo=SoDuKhachHang.MaKhachHang) INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo) INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo " _
        & " WHERE TK_ID=" + CStr(TKCNKH_ID) + " AND (SoDuKhachHang.DuNo_" + CStr(i - 1) + ">0 OR SoDuKhachHang.No_" + CStr(i) + ">0) " + IIf(mpl > 0, " AND PhanLoaiKhachHang.SoHieu LIKE '" + MaSo2SoHieu(mpl, "PhanLoaiKhachHang") + "*'", "") + " ORDER BY KhachHang.SoHieu DESC", dbOpenSnapshot)
    If rs_tk.recordCount > 0 Then
        rs_tk.MoveLast
        GauGe.Max = rs_tk.recordCount
    End If
    i = 0
    Do While (Not rs_tk.BOF) And (Not StopPrint)
        HienThongBao VString(rs_tk!sohieu + " - " + rs_tk!Ten), 1
        SoDuKHNgay rs_tk!MaSo, ngayh, duno, duco, X
        duco = duno - duco
        SoDuKHNgay rs_tk!MaSo, ngayh - songay - 1, duno, X, Y
        duno = duno - X
        Set rs_ct = DBKetoan.OpenRecordset("SELECT DISTINCTROW Sum(ChungTu.SoPS) As SoThu FROM " + ChungTu2TKNC(1) _
            & "WHERE (TK_ID=" + CStr(TKCNKH_ID) + ") AND (MaKHC = " + CStr(rs_tk!MaSo) + ") AND " + WNgay("NgayGS", ngayh - songay, ngayh), dbOpenSnapshot, dbForwardOnly)
        If Not IsNull(rs_ct!SoThu) Then duno = duno - rs_ct!SoThu
        X = 0
        If duno > 0 Or duco > 0 Then
            dk = True
            Set rs_ct = DBKetoan.OpenRecordset("SELECT DISTINCTROW ChungTu.* FROM " + ChungTu2TKNC(-1) _
                & "WHERE (TK_ID=" + CStr(TKCNKH_ID) + ") AND (MaKH = " + CStr(rs_tk!MaSo) + ") AND (NgayCT <= #" + Format(ngayh, Mask_DB) + "#) ORDER BY NgayCT DESC", dbOpenSnapshot)
            If rs_ct.recordCount > 0 Then
                Do While (Not rs_ct.EOF) And (duno > 0 Or (dk And duco > 0))
                    If rs_ct!NgayCT + songay < ngayh Then
                        i = i + 1
                        If ExecuteSQL5("INSERT INTO BaoCaoCP (MaSo, SoHieu, MaCha, BC_ID, Kq1, Kq2) VALUES (" _
                            + CStr(i) + ",'" + CStr(i) + "', " + CStr(rs_tk!MaSo) + "," + CStr(rs_ct!MaSo) + "," + DoiDau(IIf(duno > rs_ct!sops, rs_ct!sops, duno)) _
                            + "," + DoiDau(IIf(dk, duco, 0)) + ")") Then GoTo KhongIn
                        dk = False
                        duno = duno - rs_ct!sops
                    Else
                        If dk Or duco - X > 0 Then
                            i = i + 1
                            If ExecuteSQL5("INSERT INTO BaoCaoCP (MaSo, SoHieu, MaCha, BC_ID, Kq1, Kq2) VALUES (" + CStr(i) + ",'" + CStr(i) + "', " + CStr(rs_tk!MaSo) + "," + CStr(rs_ct!MaSo) + ", 0," + DoiDau(IIf(dk, duco, 0)) + ")") Then GoTo KhongIn
                            dk = False
                            X = X + rs_ct!sops
                        End If
                    End If
                    rs_ct.MoveNext
                Loop
                If duno > 0 Or (dk And duco > 0) Then
                    i = i + 1
                    If ExecuteSQL5("INSERT INTO BaoCaoCP (MaSo, SoHieu, MaCha, BC_ID, Kq1, Kq2) VALUES (" + CStr(i) + ",'" + CStr(i) + "', " + CStr(rs_tk!MaSo) + ", -1," + DoiDau(duno) + "," + DoiDau(IIf(dk, duco, 0)) + ")") Then GoTo KhongIn
                    dk = False
                End If
            Else
                i = i + 1
                If ExecuteSQL5("INSERT INTO BaoCaoCP (MaSo, SoHieu, MaCha, BC_ID, Kq1, Kq2) VALUES (" + CStr(i) + ",'" + CStr(i) + "', " + CStr(rs_tk!MaSo) + ", -1," + DoiDau(duno) + "," + DoiDau(IIf(dk, duco, 0)) + ")") Then GoTo KhongIn
                dk = False
            End If
        End If
        rs_tk.MovePrevious
        GauGe.Value = GauGe.Value + 1
    Loop
    HienThongBao Me.Caption, 1
    If i > 0 Then
        InCTNoQuaHan = True
        SetSQL "QChitiet", "SELECT DISTINCTROW KhachHang.MaSo, KhachHang.SoHieu, KhachHang.Ten, ChungTu.ThangCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai, ChungTu.SoPS, ChungTu.SoPS2Co, BaoCaoCP.Kq1, BaoCaoCP.Kq2, HeThongTK_1.SoHieu, BaoCaoCP.BC_ID " _
            & "FROM ((KhachHang INNER JOIN BaoCaoCP ON KhachHang.MaSo = BaoCaoCP.MaCha) LEFT JOIN ChungTu ON BaoCaoCP.BC_ID = ChungTu.MaSo) LEFT JOIN HeThongTK AS HeThongTK_1 ON ChungTu.MaTKTCCo = HeThongTK_1.MaSo ORDER BY KhachHang.SoHieu, NgayCT"
        frmMain.Rpt.ReportFileName = "CNQH.RPT"
        RptSetDate ngayh
        frmMain.Rpt.Formulas(5) = "ThoiGian = 'TÝnh ®Õn ngµy " + Format(ngayh, "dd/mm/yy") + " víi thêi h¹n " + CStr(songay) + " ngµy " + IIf(mpl > 0, " - " + CboKH.Text, "") + "'"
        frmMain.Rpt.Formulas(6) = "NgayHan = Date(" + CStr(Year(ngayh)) + "," + CStr(Month(ngayh)) + "," + CStr(Day(ngayh)) + ")"
    Else
        InCTNoQuaHan = False
        MsgBox "Kh«ng cã nî qu¸ h¹n !", vbInformation, App.ProductName
    End If
    rs_tk.Close
    Set rs_tk = Nothing
    On Error Resume Next
    rs_ct.Close
    Set rs_ct = Nothing
    On Error GoTo 0
    Exit Function
KhongIn:
    InCTNoQuaHan = False
    HienThongBao Me.Caption, 1
    ErrMsg er_RWait
End Function
'==========================================================================================================
' Sub TaoBaoCao
'==========================================================================================================
Private Sub TaoBaoCao()
Dim thg_dau As Integer
Dim thg_cuoi As Integer
Dim mql As Long, wsql As String
       
      wsql = ""
      If Opt(0).Value And baocao <> 28 And baocao <> 29 And Len(txtShTS.Text) > 0 Then wsql = " AND (TaiSan.SoHieu LIKE '" + txtShTS.Text + "*') "
      If Opt(1).Value Then wsql = " AND (ThongSo.MaDTQL = " + CStr(Cbo.ItemData(Cbo.ListIndex)) + ") "
       
      If CboThang(0).Enabled = True Then thg_dau = CboThang(0).ItemData(CboThang(0).ListIndex)
      If CboThang(1).Enabled = True Then thg_cuoi = CboThang(1).ItemData(CboThang(1).ListIndex)
      
      Select Case baocao
            Case 61:
                SetSQL "SoKhauHao", "SELECT DISTINCTROW TaiSan.SoHieu AS SoHieuTS, Last(TaiSan.Ten) AS TenTS, Last(LoaiTaiSan.SoHieu) AS SoHieuTK, Last(LoaiTaiSan.MaSo) AS MaTK, Last(LoaiTaiSan.Ten) AS TenTK, Last(TaiSan.NamKH) AS SoNamKH," _
                        & "Last(NG_NS+NG_TBS+NG_CNK+NG_TD) AS NG, Last(LoaiTaiSan_1.SoHieu) AS SHLoai,Last(LoaiTaiSan_1.Ten) AS TenLoai, " _
                        & "Sum(IIF(Thang=1 AND ThangTang<=" + CStr(CThangFR(1)) + " AND ThangGiam>=" + CStr(CThangFR(1)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH1, Sum(IIF(Thang=2 AND ThangTang<=" + CStr(CThangFR(2)) + " AND ThangGiam>=" + CStr(CThangFR(2)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH2, " _
                        & "Sum(IIF(Thang=3 AND ThangTang<=" + CStr(CThangFR(3)) + " AND ThangGiam>=" + CStr(CThangFR(3)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH3, Sum(IIF(Thang=4 AND ThangTang<=" + CStr(CThangFR(4)) + " AND ThangGiam>=" + CStr(CThangFR(4)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH4, " _
                        & "Sum(IIF(Thang=5 AND ThangTang<=" + CStr(CThangFR(5)) + " AND ThangGiam>=" + CStr(CThangFR(5)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH5, Sum(IIF(Thang=6 AND ThangTang<=" + CStr(CThangFR(6)) + " AND ThangGiam>=" + CStr(CThangFR(6)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH6, " _
                        & "Sum(IIF(Thang=7 AND ThangTang<=" + CStr(CThangFR(7)) + " AND ThangGiam>=" + CStr(CThangFR(7)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH7, Sum(IIF(Thang=8 AND ThangTang<=" + CStr(CThangFR(8)) + " AND ThangGiam>=" + CStr(CThangFR(8)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH8, " _
                        & "Sum(IIF(Thang=9 AND ThangTang<=" + CStr(CThangFR(9)) + " AND ThangGiam>=" + CStr(CThangFR(9)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH9, Sum(IIF(Thang=10 AND ThangTang<=" + CStr(CThangFR(10)) + " AND ThangGiam>=" + CStr(CThangFR(10)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH10, " _
                        & "Sum(IIF(Thang=11 AND ThangTang<=" + CStr(CThangFR(11)) + " AND ThangGiam>=" + CStr(CThangFR(11)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH11, Sum(IIF(Thang=12 AND ThangTang<=" + CStr(CThangFR(12)) + " AND ThangGiam>=" + CStr(CThangFR(12)) + ",KH_NS+KH_TBS+KH_CNK+KH_TD,0)) AS KH12 " _
                        & "FROM LoaiTaiSan AS LoaiTaiSan_1 RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON LoaiTaiSan_1.MaSo=TaiSan.MaLoai " _
                        & "WHERE (True) " + wsql + " GROUP BY TaiSan.SoHieu"
            Case 18:                                          ' Sæ khÊu hao
                  SetSQL "SoKhauHao", "SELECT DISTINCTROW TaiSan.SoHieu AS SoHieuTS, Last(TaiSan.Ten) AS TenTS, Last(LoaiTaiSan.SoHieu) AS SoHieuTK, Last(LoaiTaiSan.MaSo) AS MaTK, Last(LoaiTaiSan.Ten) AS TenTK, " _
                        & "Last(ThongSo.NG_NS) AS NG_NS, Last(ThongSo.NG_TBS) AS NG_TBS, Last(ThongSo.NG_CNK) AS NG_CNK, Last(ThongSo.NG_TD) AS NG_TD, Sum(ThongSo.KH_NS) AS KH_NS, Sum(ThongSo.KH_TBS) AS KH_TBS, Sum(ThongSo.KH_CNK) AS KH_CNK, Sum(ThongSo.KH_TD) AS KH_TD,Last(LoaiTaiSan_1.SoHieu) AS SHLoai,Last(LoaiTaiSan_1.Ten) AS TenLoai " _
                        & "FROM LoaiTaiSan AS LoaiTaiSan_1 RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON LoaiTaiSan_1.MaSo=TaiSan.MaLoai " _
                        & "WHERE " + WThang("ThangTang", 0, thg_cuoi) + " AND " + WThang("ThangGiam", thg_dau, 0) + " AND ThongSo.Thang >= " + CStr(CThangDB(thg_dau)) + " AND ThongSo.Thang <= " + CStr(CThangDB(thg_cuoi)) + wsql + " GROUP BY TaiSan.SoHieu"
            Case 19:                                         ' Sè d­ ®Çu kú
                  SetSQL "SoDuDauCuoiKy", "SELECT DISTINCTROW TaiSan.Ten AS TenTS, TaiSan.SoHieu AS SoHieuTS, LoaiTaiSan.SoHieu AS SoHieuTK, LoaiTaiSan.MaSo AS MaTK, LoaiTaiSan.Ten AS TenTK, " _
                        & "ThongSo.NG_NS, ThongSo.NG_TBS, ThongSo.NG_CNK, ThongSo.NG_TD, ThongSo.CL_NS, ThongSo.CL_TBS, ThongSo.CL_CNK, ThongSo.CL_TD,LoaiTaiSan_1.SoHieu,LoaiTaiSan_1.Ten " _
                        & "FROM LoaiTaiSan AS LoaiTaiSan_1 RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON LoaiTaiSan_1.MaSo=TaiSan.MaLoai " _
                        & "WHERE " + WThang("ThangTang", 0, ThangTruoc(thg_dau)) + " AND " + WThang("ThangGiam", thg_dau, 0) + " AND ThongSo.Thang = " + CStr(CThangDB(ThangTruoc(thg_dau))) + wsql + " ORDER BY TaiSan.SoHieu"
            Case 20:                                         ' Sè d­ cuèi kú
                  SetSQL "SoDuDauCuoiKy", "SELECT DISTINCTROW TaiSan.Ten AS TenTS, TaiSan.SoHieu AS SoHieuTS, LoaiTaiSan.SoHieu AS SoHieuTK, LoaiTaiSan.MaSo AS MaTK, LoaiTaiSan.Ten AS TenTK, " _
                        & "ThongSo.NG_NS, ThongSo.NG_TBS, ThongSo.NG_CNK, ThongSo.NG_TD, ThongSo.CL_NS, ThongSo.CL_TBS, ThongSo.CL_CNK, ThongSo.CL_TD,LoaiTaiSan_1.SoHieu,LoaiTaiSan_1.Ten " _
                        & "FROM LoaiTaiSan AS LoaiTaiSan_1 RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON LoaiTaiSan_1.MaSo=TaiSan.MaLoai " _
                        & "WHERE " + WThang("ThangTang", 0, thg_dau) + " AND " + WThang2("ThangGiam", thg_dau, 0) + " AND ThongSo.Thang = " + CStr(CThangDB(thg_dau)) + wsql + " ORDER BY TaiSan.SoHieu"
            Case 45:                                         ' Sæ TSC§
                  SetSQL "MienTru", "SELECT * FROM CTTaiSan WHERE (Thang=0 AND MaLoai=30)"
                  SetSQL "SoDuDauCuoiKy", "SELECT DISTINCTROW TaiSan.Ten AS TenTS, TaiSan.SoHieu AS SoHieuTS,NangLuc,TaiSan.GhiChu,ThangTang,NamSX,NamSD, LoaiTaiSan.SoHieu AS SoHieuTK, LoaiTaiSan.MaSo AS MaTK, LoaiTaiSan.Ten AS TenTK, HethongTK.SoHieu AS DTCP,DTQLy.Ten AS DTQL," _
                        & "ThongSo.NG_NS, ThongSo.NG_TBS, ThongSo.NG_CNK, ThongSo.NG_TD, ThongSo.CL_NS, ThongSo.CL_TBS, ThongSo.CL_CNK, ThongSo.CL_TD,ThongSo.KH_NS, ThongSo.KH_TBS, ThongSo.KH_CNK, ThongSo.KH_TD,LoaiTaiSan_1.SoHieu,LoaiTaiSan_1.Ten, " _
                        & "(MienTru.NG_NS+MienTru.NG_TBS+MienTru.NG_TD+MienTru.NG_CNK-MienTru.CL_NS-MienTru.CL_TBS-MienTru.CL_TD-MienTru.CL_CNK) AS CLDK " _
                        & "FROM (((LoaiTaiSan AS LoaiTaiSan_1 RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON LoaiTaiSan_1.MaSo=TaiSan.MaLoai) INNER JOIN DTQLy ON ThongSo.MaDTQL=DTQLy.MaSo) INNER JOIN HethongTK ON ThongSo.MaDTSD=HethongTK.MaSo) LEFT JOIN MienTru ON TaiSan.MaSo=MienTru.MaTS " _
                        & "WHERE " + WThang("ThangTang", 0, thg_dau) + " AND " + WThang2("ThangGiam", thg_dau, 0) + " AND ThongSo.Thang = " + CStr(CThangDB(thg_dau)) + wsql + " ORDER BY TaiSan.SoHieu"
            Case 68:                                         ' Sæ TSC§
                  SoTSCD thg_dau
            Case 21, 23:                                    ' T¨ng gi¶m tµi s¶n
                  SetSQL "TangGiam", "SELECT DISTINCTROW TaiSan.MaSo AS MaTS, TaiSan.Ten AS TenTS, TaiSan.SoHieu AS SoHieuTS, LoaiTaiSan.MaSo AS MaTK, LoaiTaiSan.SoHieu AS SoHieuTK, LoaiChungTu.SoHieu AS SoHieuLoai, LoaiChungTu.Ten AS TenLoai, " _
                        & "CTTaiSan.SoHieu AS SoHieuCT, CTTaiSan.VaoSo, CTTaiSan.DienGiai" + IIf(nn > 0, "E", "") + ", CTTaiSan.NG_NS, CTTaiSan.NG_TBS, CTTaiSan.NG_CNK, CTTaiSan.NG_TD, CTTaiSan.CL_NS, CTTaiSan.CL_TBS, CTTaiSan.CL_CNK, CTTaiSan.CL_TD " _
                        & "FROM (LoaiChungTu RIGHT JOIN (LoaiTaiSan RIGHT JOIN (TaiSan INNER JOIN CTTaiSan ON TaiSan.MaSo = CTTaiSan.MaTS) ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) ON LoaiChungTu.MaSo = CTTaiSan.MaNhom) INNER JOIN ThongSo ON (TaiSan.MaSo = ThongSo.MaTS AND TaiSan.ThangTang = ThongSo.Thang) " _
                        & "WHERE CTTaiSan.MaLoai = " + CStr(IIf(baocao = 21, NV_TANG, NV_GIAM)) + " AND " + WThang("CTTaiSan.Thang", thg_dau, 0) + " AND " + WThang("CTTaiSan.Thang", 0, thg_cuoi) + wsql + " ORDER BY TaiSan.SoHieu"
            Case 22, 24:                                    ' §¸nh gi¸ l¹i t¨ng gi¶m
                  SetSQL "DanhGiaLai", "SELECT DISTINCTROW TaiSan.MaSo AS MaTS, TaiSan.Ten AS TenTS, TaiSan.SoHieu AS SoHieuTS, LoaiTaiSan.MaSo AS MaTK, LoaiTaiSan.SoHieu AS SoHieuTK, LoaiChungTu.SoHieu AS SoHieuLoai, LoaiChungTu.Ten AS TenLoai, " _
                        & "CTTaiSan.SoHieu AS SoHieuCT, CTTaiSan.VaoSo, CTTaiSan.DienGiai" + IIf(nn > 0, "E", "") + ", CTTaiSan.NG_NS, CTTaiSan.NG_TBS, CTTaiSan.NG_CNK, CTTaiSan.NG_TD, CTTaiSan.CL_NS, CTTaiSan.CL_TBS, CTTaiSan.CL_CNK, CTTaiSan.CL_TD " _
                        & "FROM ThongSo RIGHT JOIN (LoaiChungTu RIGHT JOIN (LoaiTaiSan RIGHT JOIN (TaiSan RIGHT JOIN CTTaiSan ON TaiSan.MaSo = CTTaiSan.MaTS) ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) ON LoaiChungTu.MaSo = CTTaiSan.MaNhom) ON (TaiSan.MaSo = ThongSo.MaTS AND TaiSan.ThangTang = ThongSo.Thang) " _
                        & "WHERE CTTaiSan.MaLoai = " + CStr(NV_DGLAI) + " AND " + WThang("CTTaiSan.Thang", thg_dau, 0) + " AND " + WThang("CTTaiSan.Thang", 0, thg_cuoi) _
                        & " AND NOT (CTTaiSan.NG_NS + CTTaiSan.NG_TBS + CTTaiSan.NG_CNK + CTTaiSan.NG_TD) " + IIf(baocao = 22, "< 0", "> 0") _
                        & " AND NOT (CTTaiSan.CL_NS + CTTaiSan.CL_TBS + CTTaiSan.CL_CNK + CTTaiSan.CL_TD) " + IIf(baocao = 22, "< 0", "> 0") + wsql + " ORDER BY TaiSan.SoHieu"
            Case 25, 26:                                      ' Ph©n bæ khÊu hao
                  SetSQL "PhanBoKhauHao", "SELECT DISTINCTROW TaiSan.Ten AS TenTS, TaiSan.SoHieu AS SoHieuTS, LoaiTaiSan.SoHieu AS SoHieuTKTS, Sum((ThongSo.KH_NS+ThongSo.KH_TBS+ThongSo.KH_CNK+ThongSo.KH_TD)) AS TKH, HethongTK.MaSo AS MaDT, HethongTK.SoHieu + ' - ' + HethongTK.Ten AS TenDT, HethongTK.SoHieu AS SoHieuTKCP " _
                        & "FROM HethongTK RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON (HethongTK.MaSo = ThongSo.MaDTSD) AND (HethongTK.MaSo = ThongSo.MaDTSD) " _
                        & "WHERE " + WThang("ThangTang", 0, thg_cuoi) + " AND " + WThang("ThangGiam", thg_dau, 0) + " AND " + WThang("ThongSo.Thang", thg_dau, 0) + " AND " + WThang("ThongSo.Thang", 0, thg_cuoi) + wsql _
                        & " GROUP BY TaiSan.Ten, TaiSan.SoHieu, LoaiTaiSan.SoHieu, HethongTK.MaSo, HethongTK.Ten, HethongTK.SoHieu ORDER BY TaiSan.SoHieu"
            Case 27:                                          ' N¨ng lùc tµi s¶n
                  SetSQL "NangLucTaiSan", "SELECT DISTINCTROW TaiSan.SoHieu, TaiSan.Ten AS TenTS, TaiSan.NangLuc, TaiSan.NamSX, TaiSan.NamSD, LoaiTaiSan.MaSo AS MaTK, LoaiTaiSan.Ten AS TenTK,LoaiTaiSan_1.MaSo AS MaLoai, LoaiTaiSan_1.Ten AS TenLoai, LoaiTaiSan_2.MaSo AS MaNhom, LoaiTaiSan_2.Ten AS TenNhom,  TinhTrang.MaSo AS MaTT, TinhTrang.Ten AS TenTT, HethongTK.SoHieu + ' - '+ HethongTK.Ten AS TenDT " _
                        & "FROM LoaiTaiSan AS LoaiTaiSan_2 RIGHT JOIN (LoaiTaiSan AS LoaiTaiSan_1 RIGHT JOIN ((LoaiTaiSan RIGHT JOIN (QuocGia RIGHT JOIN TaiSan ON QuocGia.MaSo = TaiSan.MaNuoc) ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN (TinhTrang RIGHT JOIN (HethongTK RIGHT JOIN ThongSo ON (HethongTK.MaSo = ThongSo.MaDTSD) AND (HethongTK.MaSo = ThongSo.MaDTSD)) ON TinhTrang.MaSo = ThongSo.MaTTSD) " _
                        & "ON TaiSan.MaSo = ThongSo.MaTS) ON LoaiTaiSan_1.MaSo = TaiSan.MaLoai) ON LoaiTaiSan_2.MaSo = TaiSan.MaNhom " _
                        & "WHERE " + WThang("ThangTang", 0, thg_dau) + " AND " + WThang2("ThangGiam", thg_dau, 0) + " AND ThongSo.Thang = " + CStr(CThangDB(thg_dau)) + wsql + " ORDER BY TaiSan.SoHieu"
            Case 28, 29:                                       ' Tæng hîp biÕn ®éng trong kú vµ luü kÕ
                  ' L­îng khÊu hao trong kú - 3
                  SetSQL "BienDong", "SELECT DISTINCTROW 'KhÊu hao' AS TenLoai, 'KH' AS SoHieuLoai, 'KhÊu hao trong kú' AS TenNhom, 'KH' AS SoHieuNhom, 0 AS TNG_NS, 0 AS TNG_TBS, 0 AS TNG_CNK, 0 AS TNG_TD, Sum(ThongSo.KH_NS) AS TCL_NS, Sum(ThongSo.KH_TBS) AS TCL_TBS, Sum(ThongSo.KH_CNK) AS TCL_CNK, Sum(ThongSo.KH_TD) AS TCL_TD, 3 AS TangGiam FROM ThongSo " _
                        & "WHERE Thang >= " + CStr(CThangDB(IIf(baocao = 29, pThangDauKy, thg_dau))) + " AND Thang <= " + CStr(CThangDB(thg_cuoi)) + wsql
                  ' L­îng t¨ng/gi¶m trong kú - 2/3
                  DBKetoan.QueryDefs("BienDong").sql = DBKetoan.QueryDefs("BienDong").sql _
                        & " UNION SELECT DISTINCTROW LoaiChungTu.Ten AS TenLoai, LoaiChungTu.SoHieu AS SoHieuLoai, LoaiChungTu_1.Ten AS TenNhom, LoaiChungTu_1.SoHieu AS SoHieuNhom, Sum(CTTaiSan.NG_NS) AS TNG_NS, Sum(CTTaiSan.NG_TBS) AS TNG_TBS, Sum(CTTaiSan.NG_CNK) AS TNG_CNK, Sum(CTTaiSan.NG_TD) AS TNG_TD, Sum(CTTaiSan.CL_NS) AS TCL_NS, Sum(CTTaiSan.CL_TBS) AS TCL_TBS, Sum(CTTaiSan.CL_CNK) AS TCL_CNK, Sum(CTTaiSan.CL_TD) AS TCL_TD, " _
                        & "IIf(Sum(CTTaiSan.NG_NS)+Sum(CTTaiSan.NG_TBS)+Sum(CTTaiSan.NG_CNK)+Sum(CTTaiSan.NG_TD) > 0, 2, 3) AS TangGiam FROM ThongSo RIGHT JOIN (LoaiChungTu RIGHT JOIN (LoaiChungTu AS LoaiChungTu_1 RIGHT JOIN CTTaiSan ON LoaiChungTu_1.MaSo = CTTaiSan.MaNhom) ON LoaiChungTu.MaSo = CTTaiSan.MaLoai) ON (CTTaiSan.MaTS = ThongSo.MaTS AND CTTaiSan.Thang = ThongSo.Thang) " _
                        & "WHERE " + WThang("CTTaiSan.Thang", IIf(baocao = 29, pThangDauKy, thg_dau), 0) + " AND " + WThang("CTTaiSan.Thang", 0, thg_cuoi) + " AND LoaiChungTu.MaSo <> " + CStr(NV_TKHAO) + wsql + " GROUP BY LoaiChungTu.Ten, LoaiChungTu.SoHieu, LoaiChungTu_1.Ten, LoaiChungTu_1.SoHieu"
                  ' Sè ®Çu kú - 1
                  DBKetoan.QueryDefs("BienDong").sql = DBKetoan.QueryDefs("BienDong").sql _
                        & " UNION SELECT DISTINCTROW '§Çu kú' AS TenLoai, 'DK' AS SoHieuLoai, 'Sè d­ ®Çu kú' AS TenNhom, 'DK' AS SoHieuNhom, Sum(ThongSo.NG_NS) AS TNG_NS, Sum(ThongSo.NG_TBS) AS TNG_TBS, Sum(ThongSo.NG_CNK) AS TNG_CNK, Sum(ThongSo.NG_TD) AS TNG_TD, Sum(ThongSo.CL_NS) AS TCL_NS, Sum(ThongSo.CL_TBS) AS TCL_TBS, Sum(ThongSo.CL_CNK) AS TCL_CNK, Sum(ThongSo.CL_TD) AS TCL_TD, 1 AS TangGiam " _
                        & "FROM TaiSan RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS " _
                        & "WHERE ThongSo.Thang = " + CStr(CThangDB(IIf(baocao = 29, 0, ThangTruoc(thg_dau))))
                  ' Sè cuèi kú - 4
                  DBKetoan.QueryDefs("BienDong").sql = DBKetoan.QueryDefs("BienDong").sql _
                        & " UNION SELECT DISTINCTROW 'Cuèi kú' AS TenLoai, 'CK' AS SoHieuLoai, '' AS TenNhom, 'CK' AS SoHieuNhom, Sum(ThongSo.NG_NS) AS TNG_NS, Sum(ThongSo.NG_TBS) AS TNG_TBS, Sum(ThongSo.NG_CNK) AS TNG_CNK, Sum(ThongSo.NG_TD) AS TNG_TD, Sum(ThongSo.CL_NS) AS TCL_NS, Sum(ThongSo.CL_TBS) AS TCL_TBS, Sum(ThongSo.CL_CNK) AS TCL_CNK, Sum(ThongSo.CL_TD) AS TCL_TD, 4 AS TangGiam " _
                        & "FROM TaiSan RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS " _
                        & "WHERE ThongSo.Thang = " + CStr(CThangDB(thg_cuoi)) + " AND " + WThang("ThangTang", 0, thg_cuoi) + " AND " + WThang2("ThangGiam", thg_cuoi, 0) + wsql
            Case 30:                                          ' Tµi s¶n kh«ng ph¶i tÝnh khÊu hao
                  SetSQL "MienTru", "SELECT DISTINCTROW TaiSan.SoHieu AS SoHieuTS, TaiSan.Ten AS TenTS, LoaiTaiSan.SoHieu AS SoHieuTK, LoaiTaiSan.Ten AS TenTK, ThongSo.NG_NS, ThongSo.NG_TBS, ThongSo.NG_CNK, ThongSo.NG_TD " _
                        & "FROM (LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS " _
                        & "WHERE ThongSo.Thang = " + CStr(CThangDB(thg_dau)) + " AND (ThongSo.CL_NS+ThongSo.CL_TBS+ThongSo.CL_CNK+ThongSo.CL_TD) = 0 AND TaiSan.ThangTang <= " + CStr(thg_dau) + " AND " + WThang2("ThangGiam", thg_dau, 0) + wsql
            Case 31:                                          ' B×nh qu©n tµi s¶n cÇn tÝnh khÊu hao
            Case 79:
                   TinhPB thg_dau
      End Select
      DBEngine.Idle
End Sub
'======================================================================================
' Sub TinhSoLieuBaoCao
'======================================================================================
Private Sub TinhSoLieuBaoCao()
Dim thg_dau As Integer
Dim thg_cuoi As Integer
Dim chi_so As Integer
          
      ' X¸c ®Þnh c¸c th¸ng cÇn tÝnh sè liÖu
      Select Case baocao
            Case 18, 45, 61:
                  thg_dau = CboThang(0).ItemData(CboThang(0).ListIndex)
                  thg_cuoi = CboThang(1).ItemData(CboThang(1).ListIndex)
            Case 20, 30: ' Sæ khÊu hao, sè d­ cuèi kú  vµ tµi s¶n kh«ng tÝnh khÊu hao
                  thg_dau = CboThang(0).ItemData(CboThang(0).ListIndex)
                  thg_cuoi = thg_dau
            Case 19:            ' Sè d­ ®Çu kú
                  thg_dau = ThangTruoc(CboThang(0).ItemData(CboThang(0).ListIndex))
                  thg_cuoi = thg_dau
            Case 28, 31:      ' BiÕn ®éng trong kú vµ b×nh qu©n tµi s¶n
                  thg_dau = ThangTruoc(CboThang(0).ItemData(CboThang(0).ListIndex))
                  thg_cuoi = CboThang(1).ItemData(CboThang(1).ListIndex)
            Case 29:            ' BiÕn ®éng luü kÕ
                  thg_dau = pThangDauKy
                  thg_cuoi = CboThang(1).ItemData(CboThang(1).ListIndex)
            Case 61:
                  thg_dau = pThangDauKy
                  thg_cuoi = ThangCuoiNamTC
            Case Else: Exit Sub
      End Select

      ' TÝnh gi¸ trÞ tµi s¶n cho c¸c th¸ng cÇn lÊy sè liÖu
      HienThongBao "CËp nhËt gi¸ trÞ tµi s¶n ...", 1
      For chi_so = CThangDB(thg_dau) To CThangDB(thg_cuoi)
            If chi_so > 0 Then
                  If SoLieu(chi_so) = False Then
                        Me.MousePointer = 11
                        CapNhatGiaTriTaiSan chi_so, GauGe
                        SoLieu(chi_so) = True
                        Me.MousePointer = 0
                  End If
            End If
      Next
      
      ' TÝnh b×nh qu©n tµi s¶n
      If baocao = 31 Then
          Dim mql As Long, wsql As String
       
            mql = IIf(Opt(0).Value, 0, Cbo.ItemData(Cbo.ListIndex))
            wsql = IIf(mql > 0, " AND (ThongSo.MaDTQL = " + CStr(mql) + ") ", "")

            Me.MousePointer = 11
            HienThongBao "TÝnh b×nh qu©n tµi s¶n", 1
            TinhTongNguyenGia DAUKY, thg_dau, wsql                                ' Gi¸ trÞ ®Çu kú
            TinhLuongTangGiam TANG, thg_dau + 1, thg_cuoi, wsql              ' L­îng t¨ng
            TinhLuongTangGiam GIAM, thg_dau + 1, thg_cuoi, wsql                ' L­îng gi¶m
            TinhBinhQuan thg_dau, thg_cuoi, wsql                                                  ' B×nh qu©n
            Me.MousePointer = 1
      End If
      HienThongBao "", 1
End Sub
'======================================================================================
' Sub TinhTongNguyenGia
'======================================================================================
Private Sub TinhTongNguyenGia(phan_loai As Integer, thg As Integer, wsql As String)
      SetSQL "TrangThai", "SELECT DISTINCTROW ThongSo.MaTS, " _
            & "IIf((ThongSo.CL_NS+ThongSo.CL_TBS+ThongSo.CL_CNK+ThongSo.CL_TD=0),0,1) AS TrangThai, " _
            & "TaiSan.ThangTang, TaiSan.ThangGiam " _
            & "FROM TaiSan RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS " _
            & "WHERE ThongSo.Thang = " + CStr(CThangDB(thg)) _
            + IIf(thg > 0, " AND " + WThang2("ThangGiam", thg, 0), "") + " AND " + WThang("ThangTang", 0, thg) + wsql
      pSelectSQL = "SELECT DISTINCTROW Sum(ThongSo.NG_NS) AS NS, Sum(ThongSo.NG_TBS) AS TBS, Sum(ThongSo.NG_CNK) AS CNK, Sum(ThongSo.NG_TD) AS TD " _
            & "FROM TrangThai INNER JOIN ThongSo ON TrangThai.MaTS = ThongSo.MaTS " _
            & "WHERE ThongSo.Thang = " + CStr(CThangDB(thg)) + " AND TrangThai.TrangThai = 1" + wsql
      CapNhatTongNguyenGia phan_loai, 1
      pSelectSQL = "SELECT DISTINCTROW Sum(ThongSo.NG_NS) AS NS, Sum(ThongSo.NG_TBS) AS TBS, Sum(ThongSo.NG_CNK) AS CNK, Sum(ThongSo.NG_TD) AS TD " _
            & "FROM TrangThai INNER JOIN ThongSo ON TrangThai.MaTS = ThongSo.MaTS " _
            & "WHERE ThongSo.Thang = " + CStr(CThangDB(thg)) + " AND TrangThai.TrangThai = 0" + wsql
      CapNhatTongNguyenGia phan_loai, 0
End Sub
'======================================================================================
' Sub CapNhatTongNguyenGia
'======================================================================================
Private Sub CapNhatTongNguyenGia(phan_loai As Integer, khau_hao As Integer)
Dim rs_tongso As Recordset
      Set rs_tongso = DBKetoan.OpenRecordset(pSelectSQL, dbOpenSnapshot)
      ExecuteSQL5 "UPDATE BinhQuan SET HT_NS = " + DoiDau(IIf(IsNull(rs_tongso!NS), 0, rs_tongso!NS)) _
                                                                                   + ", HT_TBS = " + DoiDau(IIf(IsNull(rs_tongso!TBS), 0, rs_tongso!TBS)) _
                                                                                   + ", HT_CNK = " + DoiDau(IIf(IsNull(rs_tongso!CNK), 0, rs_tongso!CNK)) _
                                                                                   + ", HT_TD = " + DoiDau(IIf(IsNull(rs_tongso!TD), 0, rs_tongso!TD)) _
                                                                                   + ", LK_NS = " + DoiDau(IIf(IsNull(rs_tongso!NS), 0, rs_tongso!NS)) _
                                                                                   + ", LK_TBS = " + DoiDau(IIf(IsNull(rs_tongso!TBS), 0, rs_tongso!TBS)) _
                                                                                   + ", LK_CNK = " + DoiDau(IIf(IsNull(rs_tongso!CNK), 0, rs_tongso!CNK)) _
                                                                                   + ", LK_TD = " + DoiDau(IIf(IsNull(rs_tongso!TD), 0, rs_tongso!TD)) _
            + " WHERE PhanLoai = " + CStr(phan_loai) + " AND KhauHao = " + CStr(khau_hao)
      rs_tongso.Close
      Set rs_tongso = Nothing
End Sub
'======================================================================================
' Sub TinhLuongTangGiam
'======================================================================================
Private Sub TinhLuongTangGiam(phan_loai As Integer, thg_dau As Integer, thg_cuoi As Integer, wsql As String)
      ' Tµi s¶n t¨ng/gi¶m trong kú ph¶i tÝnh khÊu hao
      pSelectSQL = "SELECT DISTINCTROW Sum(CTTaiSan.NG_NS) AS NS, Sum(CTTaiSan.NG_TBS) AS TBS, Sum(CTTaiSan.NG_CNK) AS CNK, Sum(CTTaiSan.NG_TD) AS TD " _
            & "From CTTaiSan INNER JOIN ThongSo ON (CTTaiSan.MaTS = ThongSo.MaTS AND CTTaiSan.Thang = ThongSo.Thang) WHERE " + WThang("CTTaiSan.Thang", thg_dau, 0) + " AND " + WThang("CTTaiSan.Thang", 0, thg_cuoi) _
            + " AND (CTTaiSan.NG_NS + CTTaiSan.NG_TBS + CTTaiSan.NG_CNK + CTTaiSan.NG_TD) " + IIf(phan_loai = TANG, "> 0", "< 0") _
            + " AND (CTTaiSan.CL_NS + CTTaiSan.CL_TBS + CTTaiSan.CL_CNK + CTTaiSan.CL_TD) <> 0 AND CTTaiSan.MaTS>0" + wsql
      CapNhatLuongTangGiam "HT", phan_loai, 1
      ' Tµi s¶n t¨ng/gi¶m trong kú kh«ng ph¶i tÝnh khÊu hao
      pSelectSQL = "SELECT DISTINCTROW Sum(CTTaiSan.NG_NS) AS NS, Sum(CTTaiSan.NG_TBS) AS TBS, Sum(CTTaiSan.NG_CNK) AS CNK, Sum(CTTaiSan.NG_TD) AS TD " _
            & "From CTTaiSan INNER JOIN ThongSo ON (CTTaiSan.MaTS = ThongSo.MaTS AND CTTaiSan.Thang = ThongSo.Thang) WHERE " + WThang("CTTaiSan.Thang", thg_dau, 0) + " AND " + WThang("CTTaiSan.Thang", 0, thg_cuoi) _
            + " AND (CTTaiSan.NG_NS + CTTaiSan.NG_TBS + CTTaiSan.NG_CNK + CTTaiSan.NG_TD) " + IIf(phan_loai = TANG, "> 0", "< 0") _
            + " AND (CTTaiSan.CL_NS + CTTaiSan.CL_TBS + CTTaiSan.CL_CNK + CTTaiSan.CL_TD) = 0 AND CTTaiSan.MaTS>0" + wsql
      CapNhatLuongTangGiam "HT", phan_loai, 0
      ' Tµi s¶n t¨ng/gi¶m lòy kÕ ph¶i tÝnh khÊu hao
      pSelectSQL = "SELECT DISTINCTROW Sum(CTTaiSan.NG_NS) AS NS, Sum(CTTaiSan.NG_TBS) AS TBS, Sum(CTTaiSan.NG_CNK) AS CNK, Sum(CTTaiSan.NG_TD) AS TD " _
            & "From CTTaiSan INNER JOIN ThongSo ON (CTTaiSan.MaTS = ThongSo.MaTS AND CTTaiSan.Thang = ThongSo.Thang) WHERE " + WThang("CTTaiSan.Thang", pThangDauKy, 0) + " AND " + WThang("CTTaiSan.Thang", 0, thg_cuoi) _
            + " AND (CTTaiSan.NG_NS + CTTaiSan.NG_TBS + CTTaiSan.NG_CNK + CTTaiSan.NG_TD) " + IIf(phan_loai = TANG, "> 0", "< 0") _
            + " AND (CTTaiSan.CL_NS + CTTaiSan.CL_TBS + CTTaiSan.CL_CNK + CTTaiSan.CL_TD) <> 0 AND CTTaiSan.MaTS>0" + wsql
      CapNhatLuongTangGiam "LK", phan_loai, 1
      ' Tµi s¶n t¨ng/gi¶m lòy kÕ kh«ng ph¶i tÝnh khÊu hao
      pSelectSQL = "SELECT DISTINCTROW Sum(CTTaiSan.NG_NS) AS NS, Sum(CTTaiSan.NG_TBS) AS TBS, Sum(CTTaiSan.NG_CNK) AS CNK, Sum(CTTaiSan.NG_TD) AS TD " _
            & "From CTTaiSan INNER JOIN ThongSo ON (CTTaiSan.MaTS = ThongSo.MaTS AND CTTaiSan.Thang = ThongSo.Thang) WHERE " + WThang("CTTaiSan.Thang", pThangDauKy, 0) + " AND " + WThang("CTTaiSan.Thang", 0, thg_cuoi) _
            + " AND (CTTaiSan.NG_NS + CTTaiSan.NG_TBS + CTTaiSan.NG_CNK + CTTaiSan.NG_TD) " + IIf(phan_loai = TANG, "> 0", "< 0") _
            + " AND (CTTaiSan.CL_NS + CTTaiSan.CL_TBS + CTTaiSan.CL_CNK + CTTaiSan.CL_TD) = 0 AND CTTaiSan.MaTS>0" + wsql
      CapNhatLuongTangGiam "LK", phan_loai, 0
End Sub
'======================================================================================
' Sub CapNhatLuongTangGiam
'======================================================================================
Private Sub CapNhatLuongTangGiam(thoi_diem As String, phan_loai As Integer, khau_hao As Integer)
Dim rs_tongso As Recordset
      Set rs_tongso = DBKetoan.OpenRecordset(pSelectSQL, dbOpenSnapshot)
      ExecuteSQL5 "UPDATE BinhQuan SET " + thoi_diem + "_NS = " + DoiDau(IIf(IsNull(rs_tongso!NS), 0, rs_tongso!NS)) _
                                                                                   + ", " + thoi_diem + "_TBS = " + DoiDau(IIf(IsNull(rs_tongso!TBS), 0, rs_tongso!TBS)) _
                                                                                   + ", " + thoi_diem + "_CNK = " + DoiDau(IIf(IsNull(rs_tongso!CNK), 0, rs_tongso!CNK)) _
                                                                                   + ", " + thoi_diem + "_TD = " + DoiDau(IIf(IsNull(rs_tongso!TD), 0, rs_tongso!TD)) _
                                    + " WHERE PhanLoai = " + CStr(phan_loai) + " AND KhauHao = " + CStr(khau_hao)
'      ExecuteSQL5 False
      rs_tongso.Close
      Set rs_tongso = Nothing
End Sub
'======================================================================================
' Sub TinhBinhQuan
'======================================================================================
Private Sub TinhBinhQuan(thg_dau As Integer, thg_cuoi As Integer, wsql As String)
Dim chi_so As Integer
      ' TÝnh tæng nguyªn gi¸
      For chi_so = 0 To CThangDB(thg_cuoi)
            TinhTongNguyenGia CThangFR(chi_so), CThangFR(chi_so), wsql
      Next
      CapNhatLuongBinhQuan CThangDB(thg_dau), CThangDB(thg_cuoi), "HT", 1
      CapNhatLuongBinhQuan 0, CThangDB(thg_cuoi), "LK", 1
End Sub
'======================================================================================
' Sub CapNhatLuongBinhQuan
'======================================================================================
Private Sub CapNhatLuongBinhQuan(thg_dau As Integer, thg_cuoi As Integer, _
                                                                                                                thoi_diem As String, khau_hao As Integer)
Dim rs_binhquan As Recordset
Dim NS As Double
Dim TBS As Double
Dim CNK As Double
Dim TD As Double
      pSelectSQL = "SELECT * FROM BinhQuan WHERE PhanLoai >= " + CStr(thg_dau) _
            + " AND PhanLoai <= " + CStr(thg_cuoi) + " ORDER BY PhanLoai"
            
      Set rs_binhquan = DBKetoan.OpenRecordset(pSelectSQL, dbOpenSnapshot)
      Do Until rs_binhquan.EOF
            If rs_binhquan!PhanLoai = thg_dau Or rs_binhquan!PhanLoai = thg_cuoi Then
                  NS = NS + rs_binhquan!HT_NS / 2
                  TBS = TBS + rs_binhquan!HT_TBS / 2
                  CNK = CNK + rs_binhquan!HT_CNK / 2
                  TD = TD + rs_binhquan!HT_TD / 2
            Else
                  NS = NS + rs_binhquan!HT_NS
                  TBS = TBS + rs_binhquan!HT_TBS
                  CNK = CNK + rs_binhquan!HT_CNK
                  TD = TD + rs_binhquan!HT_TD
            End If
            rs_binhquan.MoveNext
      Loop
      If rs_binhquan.recordCount > 0 Then
            If thg_cuoi > thg_dau + 1 Then
                  NS = NS / (thg_cuoi - thg_dau)
                  TBS = TBS / (thg_cuoi - thg_dau)
                  CNK = CNK / (thg_cuoi - thg_dau)
                  TD = TD / (thg_cuoi - thg_dau)
            End If
      Else
            NS = 0
            TBS = 0
            CNK = 0
            TD = 0
      End If
      ExecuteSQL5 "UPDATE BinhQuan SET " + thoi_diem + "_NS = " + DoiDau(RoundMoney(NS)) _
                                                                                   + ", " + thoi_diem + "_TBS = " + DoiDau(RoundMoney(TBS)) _
                                                                                   + ", " + thoi_diem + "_CNK = " + DoiDau(RoundMoney(CNK)) _
                                                                                   + ", " + thoi_diem + "_TD = " + DoiDau(RoundMoney(TD)) _
            + " WHERE PhanLoai = " + CStr(BINHQUAN) + " AND KhauHao = " + CStr(khau_hao)
      rs_binhquan.Close
      Set rs_binhquan = Nothing
End Sub
'======================================================================================
' Sub InPSVuViec:   Liet ke phat sinh theo vu viec
'======================================================================================
Private Sub InPSVuViec(VV As Long, tdau As Integer, tcuoi As Integer, dg As String)
Dim sql As String
Dim rs As Recordset

    sql = "DELETE FROM BaoCaoCP"
    ExecuteSQL5 sql

    sql = "INSERT INTO BaoCaoCP(MaSo, SoHieu, Cap, MaCha) SELECT DISTINCTROW MaSo, SoHieu, Cap, TKCha0 FROM HethongTK WHERE Cap<2 AND Loai>0 ORDER BY IIF(LEFT(SoHieu,1) = 'L',Right(SoHieu,1),SoHieu)"
    ExecuteSQL5 sql
        
    sql = "SELECT DISTINCTROW First(BaoCaoCP.MaSo) AS MS, BaoCaoCP.SoHieu, Sum(ChungTu.SoPS) AS PS" _
        & " FROM (" + ChungTu2TKNC(-2) + ") INNER JOIN BaoCaoCP ON HeThongTK.SoHieu LIKE BaocaoCP.SoHieu" _
        & " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaDT = " + CStr(VV) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1 = " + DoiDau(rs!ps) + " WHERE MaSo = " + CStr(rs!ms)
        rs.MoveNext
    Loop
    
    sql = "SELECT DISTINCTROW First(BaoCaoCP.MaSo) AS MS, BaoCaoCP.SoHieu, Sum(ChungTu.SoPS) AS PS" _
        & " FROM (" + ChungTu2TKNC(2) + ") INNER JOIN BaoCaoCP ON HeThongTK.SoHieu LIKE BaocaoCP.SoHieu" _
        & " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaDT = " + CStr(VV) + " GROUP BY BaoCaoCP.SoHieu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        ExecuteSQL5 "UPDATE BaoCaoCP SET Kq2 = " + DoiDau(rs!ps) + " WHERE MaSo = " + CStr(rs!ms)
        rs.MoveNext
    Loop
    
    sql = "DELETE FROM BaoCaoCP WHERE Kq1 = 0 AND Kq2 = 0 AND Cap > 0"
    ExecuteSQL5 sql
    
    sql = "SELECT DISTINCTROW BaoCaoCP.SoHieu, BaocaoCP.Cap, HethongTK.Ten" + IIf(nn > 0, "E", "") + ", Kq1, Kq2" _
        & " FROM BaoCaoCP INNER JOIN HethongTK ON BaoCaoCP.MaSo = HethongTK.MaSo"
    SetSQL "QTongHopCT", sql
    
    frmMain.Rpt.ReportFileName = "PSVV.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(5) = "TaiKhoan = '" + dg + "'"
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub InCTDoanhThu(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, mkh As Long, mvt As Long, mpl As Long)
    Dim sql As String, dn As Double, DC As Double, nt As Double, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,MaCha,CoCon) SELECT ChungTu.MaSo,Cstr(ChungTu.MaSo),MaCT,IIF(TK_ID=" + CStr(TKDT_ID) + ",0,1) FROM " + ChungTu2TKNC(1) + " WHERE (TK_ID=" + CStr(TKDT_ID) + " OR (TK_ID=" + CStr(GTGTPN_ID) + " AND SoPS<>0)) AND " + wsql + IIf(mvt > 0, " AND MaVattu=" + CStr(mvt), "")
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN ChungTu ON BaoCaoCP.MaCha=ChungTu.MaCT SET BC_ID=MaKH WHERE MaKH>0"
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN (" + ChungTu2TKHD(0) + ") ON BaoCaoCP.MaCha=ChungTu.MaCT SET BC_ID=MaKhachHang WHERE BC_ID=0"
    
    If mpl > 0 Then
        ExecuteSQL5 "DELETE BaoCaoCP.* FROM BaoCaoCP INNER JOIN KhachHang ON BaoCaoCP.BC_ID=KhachHang.MaSo WHERE MaPhanLoai<>" + CStr(mpl)
        frmMain.Rpt.Formulas(4) = "PLKH='" + CboKH.Text + "'"
    End If
        
    If mkh > 0 Then
        ExecuteSQL5 "DELETE * FROM BaoCaoCP WHERE BC_ID<>" + CStr(mkh)
        sql = ""
        frmMain.Rpt.Formulas(4) = "TenKH='" + TenKH(sql, mkh) + "'"
        frmMain.Rpt.Formulas(5) = "MaSoKH='" + sql + "'"
        If tdau > 0 Then
            SoDuKH mkh, ThangTruoc(tdau), dn, DC, nt, , pSHPT
        Else
            SoDuKHNgay mkh, ndau - 1, dn, DC, nt, , pSHPT
        End If
        frmMain.Rpt.Formulas(6) = "NoDK=" + DoiDau(dn - DC)
    End If
    
    sql = "SELECT DISTINCTROW ChungTu.MaSo,ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.SoPS2Co, HeThongTK.SoHieu AS SHTK, Vattu.SoHieu AS SHVT, Vattu.TenVattu, Vattu.DonVi,ChungTu.MaKH, KhachHang.SoHieu AS SHKH,KhachHang.Ten,BaoCaoCP.CoCon, 0 AS TTN, ChungTu.MaCT" _
        & " FROM (((ChungTu LEFT JOIN HethongTK ON ChungTu.MaTKNo = HethongTK.MaSo) LEFT JOIN Vattu ON ChungTu.MaVattu = Vattu.MaSo) INNER JOIN BaoCaoCP ON ChungTu.MaSo=BaoCaoCP.MaSo) LEFT JOIN KhachHang ON BaoCaoCP.BC_ID=KhachHang.MaSo" _
        + " WHERE TK_ID=" + CStr(TKCNKH_ID) + IIf(mkh > 0, " AND BC_ID=" + CStr(mkh), "")
    
    sql = sql + " UNION SELECT DISTINCTROW ChungTu.MaSo,ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.SoPS2Co, HeThongTK.SoHieu AS SHTK, Vattu.SoHieu AS SHVT, Vattu.TenVattu, Vattu.DonVi,ChungTu.MaKH, '' AS SHKH,'' AS Ten,IIF(TK.TK_ID=" + CStr(TKDT_ID) + ",0,1) AS CoCon, 1 AS TTN, ChungTu.MaCT" _
        & " FROM (((" + ChungTu2TKNC(-1) + ") INNER JOIN BaoCaoCP ON ChungTu.MaCT=BaoCaoCP.MaCha) LEFT JOIN Vattu ON ChungTu.MaVattu = Vattu.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo" _
        & " WHERE (HethongTK.SoHieu LIKE '11*' OR HethongTK.SoHieu LIKE '53*') " + IIf(mkh > 0, " AND BC_ID=" + CStr(mkh), "")
        
    sql = sql + " UNION SELECT DISTINCTROW ChungTu.MaSo,ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.SoPS2Co, HeThongTK.SoHieu AS SHTK, '' AS SHVT, '' AS TenVattu, '' AS DonVi,ChungTu.MaKH, KhachHang.SoHieu AS SHKH,KhachHang.Ten,2 AS CoCon, 1 AS TTN, ChungTu.MaCT" _
        & " FROM (((" + ChungTu2TKNC(-1) + ") INNER JOIN KhachHang ON ChungTu.MaKHC=KhachHang.MaSo) INNER JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) LEFT JOIN BaoCaoCP ON ChungTu.MaCT=BaoCaoCP.MaCha" _
        & " WHERE " + wsql + " AND IsNull(BaoCaoCP.MaCha) AND (HethongTK.SoHieu LIKE '11*' OR HethongTK.SoHieu LIKE '331*' OR HethongTK.SoHieu LIKE '531*') " + IIf(mkh > 0, " AND MaKHC=" + CStr(mkh), "") + " AND TK.SoHieu LIKE '" + pSHPT + "*'"
    
    'sql = sql + " UNION SELECT DISTINCTROW ChungTu.MaSo,ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.SoPS2Co, HeThongTK.SoHieu AS SHTK, '', '', '',ChungTu.MaKH, '' AS SHKH,'' AS Ten,1 AS CoCon, 1 AS TTN, ChungTu.MaCT" _
        & " FROM " + ChungTu2TKNC(0) + " WHERE (HethongTK.SoHieu LIKE '11*') AND TK.TK_ID=" + CStr(TKCNKH_ID) + IIf(mkh > 0, " AND MaKHC=" + CStr(mkh), "")
    
    SetSQL "MienTru", sql
    SetSQL "QNhatKy", "SELECT ChungTu.MaCT, HoaDon.SoHD FROM " + ChungTu2TKHD(0) + " WHERE " + wsql + " GROUP BY ChungTu.MaCT, HoaDon.SoHD"
    SetSQL "QChitiet", "SELECT MienTru.*,QNhatKy.SoHD FROM MienTru LEFT JOIN QNhatKy ON MienTru.MaCT=QNhatKy.MaCT ORDER BY NgayCT,SoHieu,CoCon"
    
    frmMain.Rpt.ReportFileName = "CTDT.RPT"
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    
    RptSetDate ncuoi, nn
End Sub

Private Sub InCTDoanhThu2(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, mvt As Long)
    Dim sql As String, st As String, dv As String, wsql As String
        
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    sql = "SELECT DISTINCTROW ChungTu.MaSo,ChungTu.SoHieu, NgayCT, NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.SoPS2Co, HethongTK.SoHieu AS TKDU " _
        & " FROM (ChungTu INNER JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo = HethongTK.MaSo" _
        + " WHERE " + wsql + " AND TK.TK_ID=" + CStr(TKDT_ID) + " AND MaVattu=" + CStr(mvt)
    sql = sql + " UNION SELECT DISTINCTROW ChungTu.MaSo,ChungTu.SoHieu, NgayCT, NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.SoPS2No, TK.SoHieu AS TKDU " _
        & " FROM (ChungTu INNER JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo = HethongTK.MaSo" _
        + " WHERE " + wsql + " AND HethongTK.TK_ID=" + CStr(TKDT_ID) + " AND TK.TK_ID=" + CStr(TKGT_ID) + " AND MaVattu=" + CStr(mvt)
    SetSQL "QChitiet", sql
    
    frmMain.Rpt.ReportFileName = "CTDT3.RPT"
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    sql = "SELECT Sum(SoPS) AS F1 FROM " + ChungTu2TKNC(-2) + " WHERE MaLoai=2 AND HethongTK.SoHieu LIKE '632*' AND MaVattu=" + CStr(mvt) + " AND " + wsql
    frmMain.Rpt.Formulas(4) = "GV = " + DoiDau(SelectSQL(sql))
    sql = TenVT(st, mvt, dv)
    frmMain.Rpt.Formulas(5) = "TenVt = IF PageNumber() = 1 THEN '" + st + " - " + sql + ABCtoVNI(" - §¬n vÞ tÝnh: ") + dv + "'"
    RptSetDate ncuoi, nn
End Sub

Private Sub InTHDoanhThu(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, loai As Integer, mkho As Long, shkh As String, mnv As Long, mn As Long, Optional mdt0 As Long = 0, Optional mdt1 As Long = 0, Optional mdt2 As Long = 0, Optional mdt3 As Long = 0, Optional shtk As String = "")
    Dim sql As String, wsql As String
        
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    If mdt0 > 0 Then
        wsql = wsql + " AND (MaDT=" + CStr(mdt0) + ")"
        frmMain.Rpt.Formulas(4) = "DG1='" + SelectSQL("SELECT DienGiai AS F1 FROM DoiTuongCT WHERE MaSo=" + CStr(mdt0)) + "'"
    End If
    If mdt1 > 0 Then
        wsql = wsql + " AND (MaDT1=" + CStr(mdt1) + ")"
        frmMain.Rpt.Formulas(5) = "DG2='" + SelectSQL("SELECT DienGiai AS F1 FROM DoiTuongCT1 WHERE MaSo=" + CStr(mdt1)) + "'"
    End If
    If mdt2 > 0 Then
        wsql = wsql + " AND (MaDT2=" + CStr(mdt2) + ")"
        frmMain.Rpt.Formulas(6) = "DG3='" + SelectSQL("SELECT DienGiai AS F1 FROM DoiTuongCT2 WHERE MaSo=" + CStr(mdt2)) + "'"
    End If
    If mdt3 > 0 Then
        wsql = wsql + " AND (MaDT3=" + CStr(mdt3) + ")"
        frmMain.Rpt.Formulas(7) = "DG4='" + SelectSQL("SELECT DienGiai AS F1 FROM DoiTuongCT3 WHERE MaSo=" + CStr(mdt3)) + "'"
    End If
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,MaCha) SELECT ChungTu.MaSo,Cstr(ChungTu.MaSo),MaCT FROM " + ChungTu2TKNC(1) + " WHERE (TK_ID=" + CStr(TKDT_ID) + ") AND " + wsql + IIf(mkho > 0, " AND MaKho=" + CStr(mkho), "") + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + " GROUP BY ChungTu.MaSo, MaCT"
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN ChungTu ON BaoCaoCP.MaCha=ChungTu.MaCT SET BC_ID=MaKH WHERE MaKH>0"
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN (" + ChungTu2TKHD(0) + ") ON BaoCaoCP.MaCha=ChungTu.MaCT SET BC_ID=MaKhachHang WHERE BC_ID=0"
    
    sql = "SELECT DISTINCTROW Sum(ChungTu.SoPS) AS Tien, Sum(ChungTu.SoPS2Co) AS Luong, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, First(KhachHang.SoHieu) AS SHKH,First(KhachHang.Ten) AS Ten, First(PhanLoaiVattu.PLCha) AS PLCha3," _
            & " First(PhanLoaiVattu.SoHieu) AS SHPL3, First(PhanLoaiVattu.TenPhanLoai) AS TenPL3" _
        & " FROM ((ChungTu LEFT JOIN (PhanLoaiVattu RIGHT JOIN Vattu ON PhanLoaiVattu.MaSo = Vattu.MaPhanLoai) ON ChungTu.MaVattu = Vattu.MaSo) INNER JOIN BaoCaoCP ON ChungTu.MaSo=BaoCaoCP.MaSo) LEFT JOIN KhachHang ON BaoCaoCP.BC_ID=KhachHang.MaSo" _
        & " WHERE (TRUE) " + IIf(mnv > 0, "AND ChungTu.MaNV=" + CStr(mnv), "") + IIf(mn > 0, "AND ChungTu.MaNguon=" + CStr(mn), "") + IIf(Len(shkh) > 0, "AND KhachHang.SoHieu='" + shkh + "'", "") _
        & " GROUP BY " + IIf(loai <> 0, "KhachHang.SoHieu,KhachHang.Ten,", "") + "Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi"
    
    SetSQL "QChitiet", sql
        
    sql = "SELECT QChitiet.*,PhanLoaiVattu.SoHieu AS SHPL2,TenPhanLoai AS TenPL2,PhanLoaiVattu.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoaiVattu ON QChitiet.PLCha3=PhanLoaiVattu.MaSo"
    SetSQL "QDuPhong", sql
    sql = "SELECT QDuPhong.*,PhanLoaiVattu.SoHieu AS SHPL1,TenPhanLoai AS TenPL1 FROM QDuPhong LEFT JOIN PhanLoaiVattu ON QDuPhong.PLCha2=PhanLoaiVattu.MaSo ORDER BY QDuPhong.Sohieu"
    SetSQL "QLuyKe", sql
        
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    If loai = 0 Then
        frmMain.Rpt.ReportFileName = "THDT1.RPT"
    Else
        If Len(shkh) > 0 Then
            DBKetoan.QueryDefs("MienTru").sql = DBKetoan.QueryDefs("QLuyKe").sql
            SetSQL "QLuyKe", "SELECT * FROM MienTru WHERE SHKH = '" + shkh + "'"
        End If
        frmMain.Rpt.ReportFileName = "THDT2.RPT"
    End If
    If mkho > 0 Then frmMain.Rpt.Formulas(4) = "DG1='" + tenkho(mkho) + "'"
    If Len(shkh) > 0 Then frmMain.Rpt.Formulas(5) = "DG2='" + TenKH(shkh, 0) + "'"
    If mnv > 0 Then frmMain.Rpt.Formulas(6) = "DG3='" + TenNV("", mnv) + "'"
    If mn > 0 Then frmMain.Rpt.Formulas(7) = "DG4='" + TenNguon(mn) + "'"
    If Len(shtk) > 0 Then frmMain.Rpt.Formulas(6) = "DG3='TK: " + shtk + "'"
    RptSetDate ncuoi, nn
End Sub

Private Sub InTK(cap As Integer, tdau As Integer, tcuoi As Integer)
    Dim st As String, sqln As String, sqlc As String, i As Integer
    Dim sqlnx As String, sqlcx As String
    
    GauGe.Max = 3
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + No_" + st
        sqlc = sqlc + " + Co_" + st
    Next
    For i = 1 To CThangDB(tcuoi)
        st = CStr(i)
        sqlnx = sqlnx + " + No_" + st
        sqlcx = sqlcx + " + Co_" + st
    Next
    GauGe.Value = 1
    st = CStr(CThangDB(ThangTruoc(tdau)))
    SetSQL "QTongHopCT", "SELECT DISTINCTROW SoHieu, First(Cap) AS FirstOfCap, First(Ten" + IIf(nn > 0, "E", "") + ") AS FirstOfTen, First(Kieu) AS FirstOfKieu, First(Loai) AS FirstOfLoai, First(MaTC = MaSo OR MaTC = 0) AS FirstOfChuanLa, First(TKCon) AS FirstOfTKCon, First(CapDuoi) AS CD," _
        & " Sum(DuNo_" + st + ") AS DkNo, Sum(DuCo_" + st + ") AS DkCo, Sum(" + sqln + ") AS PsNo, Sum(" + sqlc + ") AS PsCo, Sum(" + sqlnx + ") AS PsNoX, Sum(" + sqlcx + ") AS PsCoX, Sum(HeThongTK.DuNo_" + CStr(CThangDB(tcuoi)) + ") AS CkNo, Sum(HeThongTK.DuCo_" + CStr(CThangDB(tcuoi)) + ") AS CkCo" _
        & " From HeThongTK WHERE Cap<=" + CStr(cap) + " GROUP BY SoHieu Having (((First(Loai))>0) AND (((First(Cap > 0))) OR ((First(MaTC > 0 AND MaTC <> MaSo))))) ORDER BY SoHieu"
    GauGe.Value = 2
    XDCapTK ""
    frmMain.Rpt.ReportFileName = "THSDCT.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
End Sub

Private Sub InBangThopCN(tdau As Integer, tcuoi As Integer, shtk As String, mpl As Long, loai As Integer, Optional th As Integer = 0)
    Dim st As String, sqln As String, sqlc As String, sqln2 As String, sqlc2 As String, i As Integer
    Dim sqlnx As String, sqlcx As String, sqlnx2 As String, sqlcx2 As String
    
    GauGe.Max = 3
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + SoDuKhachHang.No_" + st
        sqlc = sqlc + " + SoDuKhachHang.Co_" + st
        sqln2 = sqln2 + " + SoDuKhachHang.No_" + st + "_NT"
        sqlc2 = sqlc2 + " + SoDuKhachHang.Co_" + st + "_NT"
    Next
    For i = 1 To tcuoi
        st = CStr(i)
        sqlnx = sqlnx + " + SoDuKhachHang.No_" + st
        sqlcx = sqlcx + " + SoDuKhachHang.Co_" + st
        sqlnx2 = sqlnx2 + " + SoDuKhachHang.No_" + st + "_NT"
        sqlcx2 = sqlcx2 + " + SoDuKhachHang.Co_" + st + "_NT"
    Next
    GauGe.Value = 1
    st = CStr(CThangDB(ThangTruoc(tdau)))
    SetSQL "QChitiet", "SELECT DISTINCTROW KhachHang.SoHieu, KhachHang.Ten, PhanLoaiKhachHang.SoHieu AS SHPL, TenPhanLoai,KyHieu," _
        & " Sum(SoDuKhachHang.DuNo_" + st + ") AS DkNo, Sum(SoDuKhachHang.DuCo_" + st + ") AS DkCo, Sum(SoDuKhachHang.DuNT_" + st + ") AS DkNT, Sum(" + sqln + ") AS PsNo, Sum(" + sqln2 + ") AS PsNo2, Sum(" + sqlc + ") AS PsCo, Sum(" + sqlc2 + ") AS PsCo2, Sum(" + sqlnx + ") AS PsNoX, Sum(" + sqlnx2 + ") AS PsNoX2, Sum(" + sqlcx + ") AS PsCoX, Sum(" + sqlcx2 + ") AS PsCoX2, Sum(SoDuKhachHang.DuNo_" + CStr(tcuoi) + ") AS CkNo, Sum(SoDuKhachHang.DuCo_" + CStr(tcuoi) + ") AS CkCo, Sum(SoDuKhachHang.DuNT_" + CStr(tcuoi) + ") AS CkNT,PhanLoaiKhachHang.PLCha" _
        & " FROM (((KhachHang INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo) INNER JOIN SoDuKhachHang ON KhachHang.MaSo=SoDuKhachHang.MaKhachHang) INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo) LEFT JOIN NguyenTe ON KhachHang.MaNT=NguyenTe.MaSo " _
        & " WHERE (SoDuKhachHang.DuNo_" + st + "<>0 OR SoDuKhachHang.DuCo_" + st + "<>0 OR (" + sqlnx + ") <>0 OR (" + sqlcx + ") <>0)" + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + IIf(mpl > 0, " AND MaPhanLoai=" + CStr(mpl), "") + IIf(loai > 0, " AND KhachHang.MaNT>0", "") + " AND LEFT(KhachHang.SoHieu,1)<>'' GROUP BY PhanLoaiKhachHang.Sohieu,KhachHang.SoHieu,KhachHang.Ten,TenPhanLoai,KyHieu,PLCha"
    SetSQL "QSoCai", "SELECT QChitiet.*, PhanLoaiKhachHang.SoHieu AS SH2,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai2,PhanLoaiKhachHang.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoaiKhachHang ON QChitiet.PLCha=PhanLoaiKhachHang.MaSo"
    SetSQL "QTongHopCT", "SELECT QSoCai.*, PhanLoaiKhachHang.SoHieu AS SH1,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai1 FROM QSoCai LEFT JOIN PhanLoaiKhachHang ON QSoCai.PLCha2=PhanLoaiKhachHang.MaSo"
    
    GauGe.Value = 2
    If loai = 0 Then
        frmMain.Rpt.ReportFileName = "THSDCN" + IIf(th <> 0, "1", "") + ".RPT"
    Else
        frmMain.Rpt.ReportFileName = "THSDCN2.RPT"
    End If
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    If Len(shtk) > 0 Then
        frmMain.Rpt.Formulas(5) = "TenBaoCao = IF PageNumber() = 1 THEN '" + ABCtoVNI("Tµi kho¶n: ") + shtk + " - " + LbTenTk(7).Caption + "'"
    End If
End Sub

Private Sub THDoanhThuGiaVon(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, mkho As Long, shtk As String, CK As Integer)
    Dim wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    GauGe.Max = 3
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,Kq1,Kq2) SELECT DISTINCTROW MaVattu,Cstr(MaVattu),Sum(SoPS2Co) AS SL,Sum(SoPS-IIF(" + CStr(CK) + ">0,CK,0)) AS TT FROM " + ChungTu2TKNC(1) _
        & " WHERE TK_ID=" + CStr(TKDT_ID) + " AND " + wsql + IIf(mkho > 0, " AND MaKho=" + CStr(mkho), "") + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + " GROUP BY MaVattu"
    
    GauGe.Value = 2
    SetSQL "MienTru", "SELECT ChungTu.SoHieu FROM " + ChungTu2TKNC(1) + " WHERE TK_ID=" + CStr(TKDT_ID) + " AND " + wsql + IIf(mkho > 0, " AND MaKho=" + CStr(mkho), "") + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + " GROUP BY ChungTu.SoHieu"
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Kq3) SELECT DISTINCTROW MaVattu,Cstr(MaVattu),Sum(SoPS) FROM (" + ChungTu2TKNC(-2) _
        & ") INNER JOIN MienTru ON ChungTu.SoHieu=MienTru.SoHieu+'GV' WHERE HethongTK.SoHieu LIKE '632*' AND " + wsql + IIf(mkho > 0, " AND MaKho=" + CStr(mkho), "") + " GROUP BY MaVattu"
    
    GauGe.Value = 1
    SetSQL "QLuyKe", "SELECT Vattu.SoHieu,TenVattu,DonVi, PhanLoaiVattu.SoHieu AS SHPL,TenPhanLoai,BaoCaoCP.Kq1,BaoCaoCP.Kq2,BaoCaoCP2.Kq3 FROM ((BaoCaoCP LEFT JOIN BaoCaoCP2 ON BaoCaoCP.MaSo=BaoCaoCP2.MaSo) INNER JOIN Vattu ON BaoCaoCP.MaSo=Vattu.MaSo) INNER JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo"
    
    RptSetDate IIf(tdau > 0, NgayCuoiThang(pNamTC, tcuoi), ncuoi), nn
    frmMain.Rpt.Formulas(4) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    If mkho > 0 Then frmMain.Rpt.Formulas(5) = "TenKho='Kho hµng: " + tenkho(mkho) + "'"
    If Len(shtk) > 0 Then frmMain.Rpt.Formulas(6) = "DG='TK: " + shtk + "'"
    frmMain.Rpt.ReportFileName = "GiaVon1.RPT"
End Sub

Private Sub CTDoanhThuGiaVon(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, mvt As Long)
    Dim rs As Recordset, wsql As String
    Dim vt As New ClsVattu
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    vt.InitVattuMaSo mvt
    GauGe.Max = 3
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,MaCha,SoHieu,Ten, Cap) SELECT DISTINCTROW ChungTu.MaSo,MaCT,Cstr(ChungTu.MaSo),ChungTu.SoHieu, ThangCT FROM " + ChungTu2TKNC(1) _
        & " WHERE TK_ID=" + CStr(TKDT_ID) + " AND " + wsql + " AND MaVattu=" + CStr(mvt)
    GauGe.Value = 1
    Set rs = DBKetoan.OpenRecordset("SELECT * FROM BaoCaoCP", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1=" + DoiDau(SelectSQL("SELECT SoPS AS F1 FROM " + ChungTu2TKNC(-2) + " WHERE HethongTK.SoHieu LIKE '632*' AND MaVattu=" + CStr(mvt) + " AND ChungTu.SoHieu = '" + rs!Ten + "GV' AND ThangCT=" + CStr(rs!cap) + " AND " + wsql)) + " WHERE MaSo=" + CStr(rs!MaSo)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    GauGe.Value = 2
    
    SetSQL "QLuyKe", "SELECT ChungTu.SoHieu,NgayCT,DienGiai,SoPS2Co,SoPS,Kq1 FROM ChungTu INNER JOIN BaoCaoCP ON ChungTu.MaSo=BaoCaoCP.MaSo"
    
    frmMain.Rpt.Formulas(4) = "ThoiGian='" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    frmMain.Rpt.Formulas(6) = "Vattu='" + vt.sohieu + " - " + vt.TenVattu + " - " + ABCtoVNI("§¬n vÞ tÝnh: ") + vt.DonVi + "'"
    frmMain.Rpt.ReportFileName = "GiaVon2.RPT"
    RptSetDate IIf(tdau > 0, NgayCuoiThang(pNamTC, tcuoi), ncuoi), nn
    Set vt = Nothing
End Sub

Private Sub KeHopDongV(tdau As Integer, tcuoi As Integer, mkh As Long, mhd As Long)
    Dim sql As String
    
    If mkh = 0 Then
        sql = "SELECT DISTINCTROW ChungTu.*,DoituongCT.*,HethongTK.SoHieu AS TKNo,TK.Sohieu AS TKCo,-1 AS ID,KhachHang.MaNT FROM (((ChungTu INNER JOIN DoituongCT ON ChungTu.MaDT=DoituongCT.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo" _
            & " WHERE Not((TK.SoHieu LIKE '3331*' OR TK.TK_ID=" + CStr(TKDT_ID) + ")) AND " + WThang("ThangCT", tdau, tcuoi) + " AND ChungTu.MaDT>1" + IIf(mhd > 1, " AND ChungTu.MaDT=" + CStr(mhd), "") ' " OR HethongTK.SoHieu LIKE '111*' OR HethongTK.SoHieu LIKE '112*'
    Else
        ExecuteSQL5 "DELETE * FROM BaoCaoCP"
        ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,MaCha) SELECT ChungTu.MaSo,Cstr(ChungTu.MaSo),MaCT FROM (ChungTu LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo " _
            & " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND Not(TK.SoHieu LIKE '3331*' OR TK.TK_ID=" + CStr(TKDT_ID) + ") AND MaDT>1" + IIf(mhd > 1, " AND MaDT=" + CStr(mhd), "")     '" OR HethongTK.SoHieu LIKE '111*' OR HethongTK.SoHieu LIKE '112*'
        ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN ChungTu ON BaoCaoCP.MaCha=ChungTu.MaCT SET BC_ID=IIF(MaKH>0,MaKH,MaKHC) WHERE MaKH>0 OR MaKHC>0"
        ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN (" + ChungTu2TKHD(0) + ") ON BaoCaoCP.MaCha=ChungTu.MaCT SET BC_ID=MaKhachHang WHERE BC_ID=0"
        
        sql = "SELECT DISTINCTROW ChungTu.*,DoituongCT.*,HethongTK.SoHieu AS TKNo,TK.Sohieu AS TKCo,-1 AS ID,KhachHang.MaNT FROM ((((ChungTu INNER JOIN DoituongCT ON ChungTu.MaDT=DoituongCT.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) INNER JOIN BaoCaoCP ON ChungTu.MaCT=BaoCaoCP.MaCha) LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo" _
            & " WHERE BC_ID=" + CStr(mkh) + " AND ChungTu.MaDT>1" + IIf(mhd > 1, " AND ChungTu.MaDT=" + CStr(mhd), "")
        
        frmMain.Rpt.Formulas(5) = "TenKH = '" + txtshkh(0).Text + " - " + lbkh(0).Caption + "'"
    End If
    
    SetSQL "QNhatKy", sql

    frmMain.Rpt.ReportFileName = "DSHD.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
End Sub

Private Sub KeHopDongR(tdau As Integer, tcuoi As Integer, mkh As Long, mhd As Long)
    Dim sql As String
        
    If mkh = 0 Then
        sql = "SELECT DISTINCTROW ChungTu.*,DoituongCT.*,HethongTK.SoHieu AS TKNo,TK.Sohieu AS TKCo,1 AS ID,KhachHang.MaNT FROM (((ChungTu INNER JOIN DoituongCT ON ChungTu.MaDT=DoituongCT.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo" _
            & " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND ((TK.SoHieu LIKE '3331*' OR TK.TK_ID=" + CStr(TKDT_ID) + " OR ((HethongTK.SoHieu LIKE '11*') AND (Not TK.SoHieu LIKE '331*')))) AND MaDT>1" + IIf(mhd > 1, " AND MaDT=" + CStr(mhd), "")
    Else
        ExecuteSQL5 "DELETE * FROM BaoCaoCP"
        ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,MaCha) SELECT ChungTu.MaSo,Cstr(ChungTu.MaSo),MaCT FROM (ChungTu LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo " _
            & " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND (TK.SoHieu LIKE '3331*' OR TK.TK_ID=" + CStr(TKDT_ID) + " OR ((HethongTK.SoHieu LIKE '11*')  AND (Not TK.SoHieu LIKE '331*'))) AND MaDT>1" + IIf(mhd > 1, " AND MaDT=" + CStr(mhd), "")
        ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN ChungTu ON BaoCaoCP.MaCha=ChungTu.MaCT SET BC_ID=BC_ID=IIF(MaKH>0,MaKH,MaKHC) WHERE MaKH>0 OR MaKHC>0"
        ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN (" + ChungTu2TKHD(0) + ") ON BaoCaoCP.MaCha=ChungTu.MaCT SET BC_ID=MaKhachHang WHERE BC_ID=0"
                
        sql = "SELECT DISTINCTROW ChungTu.*,DoituongCT.*,HethongTK.SoHieu AS TKNo,TK.Sohieu AS TKCo,1 AS ID,KhachHang.MaNT FROM ((((ChungTu INNER JOIN DoituongCT ON ChungTu.MaDT=DoituongCT.MaSo) LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) LEFT JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) INNER JOIN BaoCaoCP ON ChungTu.MaCT=BaoCaoCP.MaCha) LEFT JOIN KhachHang ON DoituongCT.MaKhachHang=KhachHang.MaSo" _
            & " WHERE (BC_ID=" + CStr(mkh) + ") AND (TK.SoHieu LIKE '3331*' OR TK.TK_ID=" + CStr(TKDT_ID) + " OR HethongTK.SoHieu LIKE '11*') AND MaDT>1" + IIf(mhd > 1, " AND MaDT=" + CStr(mhd), "")
        frmMain.Rpt.Formulas(5) = "TenKH = '" + txtshkh(0).Text + " - " + lbkh(0).Caption + "'"
    End If
    SetSQL "QNhatKy", sql
    
    frmMain.Rpt.ReportFileName = "DSHD.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
End Sub

Public Function InSoChitietCN(taikhoan As ClsTaikhoan, mcn As Long, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date) As Boolean
    Dim dkn As Double, dkc As Double, dknt As Double, sqlw As String, X As Double
    Dim sql As String, cn As New ClsKhachHang, i As Integer
    
    cn.InitKhachHangMaSo mcn
    If tcuoi > 0 Then
        cn.SoDuKH ThangTruoc(tdau), dkn, dkc, dknt, IIf(taikhoan.tkcon = 0, taikhoan.MaSo, 0), IIf(taikhoan.tkcon > 0, taikhoan.sohieu, "")
        sqlw = WThang("ThangCT", tdau, tcuoi)
    Else
        SoDuKHNgay cn.MaSo, ndau - 1, dkn, dkc, dknt, IIf(taikhoan.tkcon = 0, taikhoan.MaSo, 0), IIf(taikhoan.tkcon > 0, taikhoan.sohieu, "")
        sqlw = WNgay("NgayGS", ndau, ncuoi)
    End If
    
    sql = "((HethongTK_2.SoHieu LIKE '" + taikhoan.sohieu + "*') Or (HethongTK_3.SoHieu LIKE '" + taikhoan.sohieu + "*')) AND (MaKH=" + CStr(mcn) + " OR MaKHC=" + CStr(mcn) + ")"
    If cn.MaNT = 0 Then
        SetSQL "QChitiet", "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.SoPS2No, ChungTu.SoPS2Co, ChungTu.GhiChu, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.MaTKNo, ChungTu.MaTKCo, HethongTK_2.SoHieu As SHNo, HethongTK_3.SoHieu As SHCo, MaKH,MaKHC, ThangCT" _
            & " FROM HeThongTK AS HeThongTK_3 RIGHT JOIN (HeThongTK AS HeThongTK_2 RIGHT JOIN (HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo) ON HeThongTK_2.MaSo = ChungTu.MaTKNo) ON HeThongTK_3.MaSo = ChungTu.MaTKCo" _
            & " Where (SoPS<>0 OR SoPS2No<>0 OR SoPS2Co<>0) AND (" + sqlw + " And (" + sql + ")) AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) ORDER BY ChungTu.NgayGS, ChungTu.MaCT"
        i = GetRowNumber("QChitiet")
        If i = 0 And dkn = 0 And dkc = 0 Then GoTo KhongIn
        If i = 0 Then
            SetSQL "QChitiet", "SELECT DISTINCTROW 0 AS MaCT, '' AS SoHieu, Null AS NgayCT, Null AS NgayGS, '' AS DienGiai, 0 AS SoPS, 0 AS SoPS2No, 0 AS SoPS2Co, '' AS GhiChu, '' AS SoHieu0, '' AS SoHieu1, 0 AS MaTKNo, 0 AS MaTKCo, '' As SHNo, '' As SHCo,0 AS MaKH,0 AS MaKHC, 0 AS ThangCT" _
                & " FROM ChungTu Where (chungtu.MaCT = 0)"
        End If
        frmMain.Rpt.Formulas(3) = "SoHieuTk='" + taikhoan.sohieu + "'"
        frmMain.Rpt.Formulas(9) = "MaCN=" + CStr(cn.MaSo)
        frmMain.Rpt.ReportFileName = "CTTKCN.RPT"
    Else
        If (taikhoan.kieu < 0 And dkc > dkn) Or (taikhoan.kieu > 0 And dkc < dkn) Then dknt = -dknt
        SetSQL "QChitiet", "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.NgayGS, ChungTu.DienGiai" + IIf(nn > 0, "E", "") + ", ChungTu.SoPS, ChungTu.SoPS2No, ChungTu.SoPS2Co, ChungTu.GhiChu, HeThongTK.SoHieu, HeThongTK_1.SoHieu, ChungTu.MaTKNo, ChungTu.MaTKCo, MaKH, MaKHC " _
            & " FROM HeThongTK AS HeThongTK_3 RIGHT JOIN (HeThongTK AS HeThongTK_2 RIGHT JOIN (HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKTCNo) ON HeThongTK_1.MaSo = ChungTu.MaTKTCCo) ON HeThongTK_2.MaSo = ChungTu.MaTKNo) ON HeThongTK_3.MaSo = ChungTu.MaTKCo" _
            & " Where (SoPS<>0 OR SoPS2No<>0 OR SoPS2Co<>0) AND (MaKH=" + CStr(mcn) + " OR MaKHC=" + CStr(mcn) + ") AND " + sqlw + " AND " + sql + " ORDER BY Chungtu.NgayGS, Chungtu.MaCT"
        i = GetRowNumber("QChitiet")
        If i = 0 And dkn = 0 And dkc = 0 Then GoTo KhongIn
        If i = 0 Then
            SetSQL "QChitiet", "SELECT DISTINCTROW 0 AS MaCT, '' AS SoHieu, Null AS NgayCT, Null AS NgayGS, '' AS DienGiai, 0 AS SoPS, 0 AS SoPS2No, 0 AS SoPS2Co, '' AS GhiChu, '' AS SoHieu1, '' AS SoHieu2, 0 AS MaTKNo, 0 AS MaTKCo, 0 AS MaKH,0 AS MaKHC" _
                & " FROM ChungTu Where (chungtu.MaCT = 0)"
        End If
        GauGe.Value = 2
        frmMain.Rpt.Formulas(3) = "MaTk=" + CStr(taikhoan.MaSo)
        frmMain.Rpt.Formulas(9) = "NoDkNT=" + DoiDau(dknt)
        frmMain.Rpt.Formulas(10) = "MaCN=" + CStr(cn.MaSo)
        frmMain.Rpt.ReportFileName = "CTTKCNNT.RPT"
    End If
    
    frmMain.Rpt.Formulas(4) = "NoDk=" + DoiDau(dkn - dkc)
    frmMain.Rpt.Formulas(5) = "Kieu=" + CStr(taikhoan.kieu)
    If tcuoi > 0 Then
        RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
        cn.SoPhatSinh pThangDauKy, tcuoi, dkn, dkc, dknt, dknt, taikhoan.MaSo
    Else
        RptSetDate ncuoi, nn
        SoPhatSinhN cn.MaSo, NgayDauThang(pNamTC, pThangDauKy), ncuoi, dkn, dkc, dknt, X, taikhoan.MaSo
    End If
    frmMain.Rpt.Formulas(6) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    frmMain.Rpt.Formulas(11) = "NoLK = " + DoiDau(dkn)
    frmMain.Rpt.Formulas(12) = "CoLK = " + DoiDau(dkc)
    frmMain.Rpt.Formulas(8) = "TenTk = IF PageNumber() = 1 THEN 'TK: " + taikhoan.sohieu + " - " + cn.sohieu + ": " + IIf(nn > 0, cn.TenE, cn.Ten) + "'"
    
    If pCongNoHD > 0 And taikhoan.tk_id = TKCNKH_ID Then
        sql = cn.HDChuaTT(IIf(tcuoi > 0, tcuoi, 0), ncuoi, taikhoan.sohieu, cn.MaSo)
        If Len(sql) > 0 Then frmMain.Rpt.Formulas(13) = "HDTT = '" + sql + "'"
    End If
    
    InSoChitietCN = True
    GoTo KT
KhongIn:
    InSoChitietCN = False
KT:
    Set cn = Nothing
End Function

Private Sub InTHXuatKho(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, mkho As Long, mtk As Long, mloai As Long, tkdu As String)
    Dim sql As String, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    
    sql = "SELECT DISTINCTROW Vattu.SoHieu,TenVattu,DonVi,Sum(SoPS2Co) AS Luong,Sum(SoPS) AS Tien FROM (" + ChungTu2TKNC(-1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo" _
        & " WHERE (ChungTu.MaLoai=2 OR (ChungTu.MaLoai=4 AND MaTKNo<>MaTKCo)) AND " + wsql + " AND HethongTK.SoHieu LIKE '" + tkdu _
        + "*'" + IIf(mkho > 0, " AND MaKho=" + CStr(mkho), "") + IIf(mtk <> 0, " AND MaTKCo = " + CStr(mtk), "") + IIf(mloai <> 0, " AND MaPhanLoai = " + CStr(mloai), "") _
        + " GROUP BY Vattu.SoHieu,TenVattu,DonVi"
    sql = sql + " UNION SELECT DISTINCTROW Vattu.SoHieu,TenVattu,DonVi,-Sum(SoPS2No) AS Luong,-Sum(SoPS) AS Tien FROM (" + ChungTu2TKNC(1) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo" _
        & " WHERE (ChungTu.MaLoai=1 OR (ChungTu.MaLoai=4 AND MaTKNo<>MaTKCo)) AND " + wsql + " AND HethongTK.SoHieu LIKE '" + tkdu _
        + "*'" + IIf(mkho > 0, " AND ((MaKho=" + CStr(mkho) + " AND MaLoai=1) OR (MaNguon=" + CStr(mkho) + " AND MaLoai=4))", "") + IIf(mtk <> 0, " AND MaTKNo = " + CStr(mtk), "") + IIf(mloai <> 0, " AND MaPhanLoai = " + CStr(mloai), "") _
        + " GROUP BY Vattu.SoHieu,TenVattu,DonVi"
    SetSQL "QLuyKe", sql
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    frmMain.Rpt.Formulas(4) = "TKDU = IF PageNumber() = 1 THEN '" + ABCtoVNI("Tµi kho¶n ®èi øng: ") + tkdu + " - " + LbTenTk(11).Caption + "'"
    frmMain.Rpt.ReportFileName = "LUYKE2.RPT"
    RptSetDate ncuoi, nn
End Sub

Private Sub InLuyKeN(mkho As Long, shtk As String, mloai As Long, ndau As Date, ncuoi As Date, Optional T As Boolean = True, Optional loailk As Integer = 0)
    Dim sql As String
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    
    GauGe.Max = 3
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,MaCha,BC_ID,MK,Kq1,Kq2) SELECT First(TonKho.MaSo),Cstr(First(TonKho.MaSo)),MaTaiKhoan,MaVattu,MaSoKho,Sum(Luong_0),Sum(Tien_0) FROM (TonKho INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo) INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo WHERE (TRUE)" _
        + IIf(mkho > 0, " AND MaSoKho=" + CStr(mkho), "") + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + IIf(mloai > 0, " AND MaPhanLoai=" + CStr(mloai), "") + " GROUP BY MaSoKho,MaTaiKhoan,MaVattu"
        
    ' Tinh dau ky
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,MaCha,BC_ID,MK,Kq1,Kq2) SELECT First(ChungTu.MaSo),Cstr(First(ChungTu.MaSo)),MaTKNo,MaVattu,IIF(MaLoai=1,MaKho,MaNguon),Sum(SoPS2No),Sum(SoPS) FROM (ChungTu INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo " _
        & " WHERE (MaLoai=1" + IIf(mkho > 0, " OR MaLoai=4)", ")") + " AND NgayGS<#" + Format(ndau, Mask_DB) + "#" _
        + IIf(mkho > 0, " AND ((MaKho=" + CStr(mkho) + " AND MaLoai=1) OR (MaNguon=" + CStr(mkho) + " AND MaLoai=4))", "") + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + IIf(mloai > 0, " AND MaPhanLoai=" + CStr(mloai), "") + " GROUP BY IIF(MaLoai=1,MaKho,MaNguon),MaTKNo,MaVattu"
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN BaoCaoCP2 ON BaoCaoCP.MaCha=BaoCaoCP2.MaCha AND BaoCaoCP.BC_ID=BaoCaoCP2.BC_ID AND BaoCaoCP.MK=BaoCaoCP2.MK SET BaoCaoCP.Kq1=BaoCaoCP.Kq1+BaoCaoCP2.Kq1,BaoCaoCP.Kq2=BaoCaoCP.Kq2+BaoCaoCP2.Kq2"
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,MaCha,BC_ID,MK,Kq1,Kq2) SELECT First(ChungTu.MaSo),Cstr(First(ChungTu.MaSo)),MaTKCo,MaVattu,MaKho,Sum(SoPS2Co),Sum(SoPS) FROM (ChungTu INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo " _
        & " WHERE (MaLoai=2" + IIf(mkho > 0, " OR MaLoai=4)", ")") + " AND NgayGS<#" + Format(ndau, Mask_DB) + "#" _
        + IIf(mkho > 0, " AND (MaKho=" + CStr(mkho) + ")", "") + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + IIf(mloai > 0, " AND MaPhanLoai=" + CStr(mloai), "") + " GROUP BY MaKho,MaTKCo,MaVattu"
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN BaoCaoCP2 ON BaoCaoCP.MaCha=BaoCaoCP2.MaCha AND BaoCaoCP.BC_ID=BaoCaoCP2.BC_ID AND BaoCaoCP.MK=BaoCaoCP2.MK SET BaoCaoCP.Kq1=BaoCaoCP.Kq1-BaoCaoCP2.Kq1,BaoCaoCP.Kq2=BaoCaoCP.Kq2-BaoCaoCP2.Kq2"
        
    ' Tinh tong so nhap xuat
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,MaCha,BC_ID,MK,Kq1,Kq2) SELECT First(ChungTu.MaSo),Cstr(First(ChungTu.MaSo)),MaTKNo,MaVattu,IIF(MaLoai=1,MaKho,MaNguon),Sum(SoPS2No),Sum(SoPS) FROM (ChungTu INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo " _
        & " WHERE (MaLoai=1" + IIf(mkho > 0, " OR MaLoai=4)", " OR (MaLoai=4 AND MaTKNo<>MaTKCo))") + " AND " + WNgay("NgayGS", ndau, ncuoi) _
        + IIf(mkho > 0, " AND ((MaKho=" + CStr(mkho) + " AND MaLoai=1) OR (MaNguon=" + CStr(mkho) + " AND MaLoai=4))", "") + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + IIf(mloai > 0, " AND MaPhanLoai=" + CStr(mloai), "") + " GROUP BY IIF(MaLoai=1,MaKho,MaNguon),MaTKNo,MaVattu"
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN BaoCaoCP2 ON BaoCaoCP.MaCha=BaoCaoCP2.MaCha AND BaoCaoCP.BC_ID=BaoCaoCP2.BC_ID AND BaoCaoCP.MK=BaoCaoCP2.MK SET BaoCaoCP.Kq3=BaoCaoCP2.Kq1,BaoCaoCP.Kq4=BaoCaoCP2.Kq2"
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,MaCha,BC_ID,MK,Kq1,Kq2) SELECT First(ChungTu.MaSo),Cstr(First(ChungTu.MaSo)),MaTKCo,MaVattu,MaKho,Sum(SoPS2Co),Sum(SoPS) FROM (ChungTu INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo " _
        & " WHERE (MaLoai=2" + IIf(mkho > 0, " OR MaLoai=4", " OR (MaLoai=4 AND MaTKNo<>MaTKCo)") + ") AND " + WNgay("NgayGS", ndau, ncuoi) _
        + IIf(mkho > 0, " AND (MaKho=" + CStr(mkho) + ")", "") + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + IIf(mloai > 0, " AND MaPhanLoai=" + CStr(mloai), "") + " GROUP BY MaKho,MaTKCo,MaVattu"
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN BaoCaoCP2 ON BaoCaoCP.MaCha=BaoCaoCP2.MaCha AND BaoCaoCP.BC_ID=BaoCaoCP2.BC_ID AND BaoCaoCP.MK=BaoCaoCP2.MK SET BaoCaoCP.Kq5=BaoCaoCP2.Kq1,BaoCaoCP.Kq6=BaoCaoCP2.Kq2"
    
     If GetSetting(IniPath, "Stock", "DailySummary", 0) = 0 Then
        ExecuteSQL5 "DELETE * FROM BaoCaoCP WHERE Kq1=0 AND Kq2=0 AND Kq3=0 AND Kq4=0 AND Kq5=0 AND Kq6=0"
    Else
        ExecuteSQL5 "DELETE * FROM BaoCaoCP WHERE Kq3=0 AND Kq4=0 AND Kq5=0 AND Kq6=0"
    End If
    ExecuteSQL5 "UPDATE BaoCaoCP SET Kq7=Kq1+Kq3-Kq5,Kq8=Kq2+Kq4-Kq6"
    XDTyLeQD 0
    GauGe.Value = 1
    sql = "SELECT DISTINCTROW First(HeThongTK.SoHieu) AS SHTK, First(HeThongTK.Ten" + IIf(nn > 0, "E", "") + ") AS TenTK, First(PhanLoaiVattu.PLCha) AS PLCha3," _
        & " First(PhanLoaiVattu.SoHieu) AS SHPL3, First(PhanLoaiVattu.TenPhanLoai) AS TenPL3, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, Sum(Kq2) AS DkTien," _
        & " Sum(Kq1) AS DkLuong, Sum(Kq4) AS TienNhap, Sum(Kq3) AS LuongNhap, Sum(Kq6) AS TienXuat," _
        & " Sum(Kq5) AS LuongXuat, Sum(Kq7) AS CkLuong, Sum(Kq8) AS CkTien, First(Vattu.MaSo) AS MVT " _
        & " FROM (HeThongTK RIGHT JOIN ((PhanLoaiVattu RIGHT JOIN Vattu ON PhanLoaiVattu.MaSo = Vattu.MaPhanLoai) RIGHT JOIN TonKho ON Vattu.MaSo = TonKho.MaVatTu) ON HeThongTK.MaSo = TonKho.MaTaiKhoan) INNER JOIN BaoCaoCP ON TonKho.MaSoKho=BaoCaoCP.MK AND TonKho.MaTaiKhoan=BaoCaoCP.MaCha AND TonKho.MaVattu=BaoCaoCP.BC_ID" _
        & " GROUP BY HeThongTK.MaSo, HeThongTK.SoHieu, HeThongTK.Ten, PhanLoaiVattu.MaSo, PhanLoaiVattu.SoHieu, PhanLoaiVattu.TenPhanLoai, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi " _
        + IIf(mkho > 0, ", TonKho.MaSoKho", "") + " Having (TRUE)" _
        + IIf(mkho <> 0, " AND TonKho.MaSoKho = " + CStr(mkho), "") + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") _
        + IIf(mloai <> 0, " AND PhanLoaiVattu.SoHieu LIKE '" + MaSo2SoHieu(mloai, "PhanLoaiVattu") + "*'", "") _
        + " ORDER BY HeThongTK.SoHieu, PhanLoaiVattu.SoHieu, Vattu.SoHieu"
    SetSQL "QChitiet", sql
    sql = "SELECT QChitiet.*,PhanLoaiVattu.SoHieu AS SHPL2,TenPhanLoai AS TenPL2,PhanLoaiVattu.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoaiVattu ON QChitiet.PLCha3=PhanLoaiVattu.MaSo"
    SetSQL "QDuPhong", sql
    sql = "SELECT QDuPhong.*,PhanLoaiVattu.SoHieu AS SHPL1,TenPhanLoai AS TenPL1 FROM QDuPhong LEFT JOIN PhanLoaiVattu ON QDuPhong.PLCha2=PhanLoaiVattu.MaSo ORDER BY QDuPhong.Sohieu"
    SetSQL "QLuyKe", sql
    GauGe.Value = 2
    frmMain.Rpt.Formulas(3) = "TenKho='" + IIf(mkho > 0, "Kho: " + CboKho(0).Text, "") + "'"
    frmMain.Rpt.Formulas(5) = "ThoiGian='" + ThoiGianN(ndau, ncuoi, nn) + "'"
    If T And loailk <> 1 Then
        frmMain.Rpt.ReportFileName = "LUYKE.RPT"
    Else
        frmMain.Rpt.ReportFileName = "LUYKEL.RPT"
    End If
    RptSetDate ncuoi, nn
End Sub

Private Sub InLuyKeHM(mkho As Long, mtk As Long, mloai As Long, ncuoi As Date, loaihm As Integer)
    Dim sql As String, rs As Recordset, dkt As Double, dkl As Double, ln As Double, lx As Double, tn As Double, tx As Double, dkt2 As Double
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    
    GauGe.Max = 3
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,MaCha,BC_ID,MK,Kq7,Kq8) SELECT TonKho.MaSo,Cstr(TonKho.MaSo),MaTaiKhoan,MaVattu,MaSoKho,TonMin,TonMax FROM TonKho INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo WHERE " + IIf(loaihm < 0, " TonMin>0", " TonMax>0") _
        + IIf(mkho > 0, " AND MaSoKho=" + CStr(mkho), "") + IIf(mtk > 0, " AND MaTaiKhoan=" + CStr(mtk), "") + IIf(mloai > 0, " AND MaPhanLoai=" + CStr(mloai), "")
    Set rs = DBKetoan.OpenRecordset("SELECT * FROM BaoCaoCP2", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        dkl = SoTonKhoN(ncuoi, rs!mk, rs!macha, rs!BC_ID, dkt, dkt2)
        If loaihm < 0 And dkl < rs!Kq7 Then ExecuteSQL5 "UPDATE BaoCaoCP2 SET CoCon=1 WHERE MaSo=" + CStr(rs!MaSo)
        If loaihm > 0 And dkl > rs!Kq8 Then ExecuteSQL5 "UPDATE BaoCaoCP2 SET CoCon=1 WHERE MaSo=" + CStr(rs!MaSo)
        ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq1=" + DoiDau(dkl) + " WHERE MaSo=" + CStr(rs!MaSo)
        rs.MoveNext
    Loop
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2 WHERE CoCon=0"
    
    GauGe.Value = 1
    sql = "SELECT DISTINCTROW PhanLoaiVattu.SoHieu, PhanLoaiVattu.TenPhanLoai, Vattu.SoHieu, Vattu.TenVattu,Vattu.TonMin,Vattu.TonMax, Kq1, Kq7,Kq8" _
        & " FROM (HeThongTK RIGHT JOIN ((PhanLoaiVattu RIGHT JOIN Vattu ON PhanLoaiVattu.MaSo = Vattu.MaPhanLoai) RIGHT JOIN TonKho ON Vattu.MaSo = TonKho.MaVatTu) ON HeThongTK.MaSo = TonKho.MaTaiKhoan) INNER JOIN BaoCaoCP2 ON TonKho.MaSoKho=BaoCaoCP2.MK AND TonKho.MaTaiKhoan=BaoCaoCP2.MaCha AND TonKho.MaVattu=BaoCaoCP2.BC_ID" _
        & " WHERE (TRUE)" + IIf(mkho <> 0, " And TonKho.MaSoKho = " + CStr(mkho), "") + IIf(mtk <> 0, " And HethongTK.MaSo = " + CStr(mtk), "") _
        + IIf(mloai <> 0, " AND PhanLoaiVattu.SoHieu LIKE '" + MaSo2SoHieu(mloai, "PhanLoaiVattu") + "*'", "") _
        + " ORDER BY HeThongTK.SoHieu, PhanLoaiVattu.SoHieu, Vattu.SoHieu"
           
    SetSQL "QLuyKe", sql
    GauGe.Value = 2
    frmMain.Rpt.Formulas(3) = "TenKho='" + IIf(mkho > 0, "Kho: " + CboKho(0).Text, "") + "'"
    frmMain.Rpt.Formulas(5) = "ThoiGian='" + ThoiGianN(ncuoi, ncuoi, nn) + "'"
    frmMain.Rpt.ReportFileName = "LUYKEL1.RPT"
    RptSetDate ncuoi, nn
End Sub

Private Function DoiChieuCN(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, mkh As Long, TK As ClsTaikhoan, ct As Integer) As Boolean
    Dim sql As String, dn As Double, DC As Double, X As Double, loai As Integer, sqlw1 As String, sqlw2 As String, Y As Double, wsql As String, kh As New ClsKhachHang, shtk As String
    
    kh.InitKhachHangMaSo mkh
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    
    DoiChieuCN = True
    If TK.MaSo = 0 Then
        loai = SelectSQL("SELECT Kieu AS F1,SoHieu AS F2 FROM HethongTK INNER JOIN SoDuKhachHang On HethongTK.MaSo=SoDuKhachHang.MaTaiKhoan WHERE MaKhachHang=" + CStr(mkh), shtk)
        If shtk Like "141*" Then
            TK.InitTaikhoanSohieu shtk
            GoTo AAA
        End If
        If tdau > 0 Then
            kh.SoDuKH ThangTruoc(tdau), dn, DC, X
        Else
            SoDuKHNgay mkh, ndau - 1, dn, DC, X, 0
        End If
    Else
AAA:
        loai = TK.kieu
        If tdau > 0 Then
            kh.SoDuKH ThangTruoc(tdau), dn, DC, X, 0, TK.sohieu
        Else
            SoDuKHNgay mkh, ndau - 1, dn, DC, X, IIf(TK.tkcon = 0, TK.MaSo, 0), IIf(TK.tkcon > 0, TK.sohieu, 0)
        End If
        sqlw1 = " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'"
        sqlw2 = " AND TK.SoHieu LIKE '" + TK.sohieu + "*'"
    End If
    
    SetSQL "NhatKy", "SELECT MaCT,Sum(IIF(MaTKNo>0,SoPS,0)) AS TPS FROM ChungTu WHERE " + wsql + " GROUP BY MaCT"
    If loai > 0 Or TK.sohieu Like "141*" Then
        SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKNC(1) + " WHERE MaKHC=" + CStr(mkh) + " AND " + wsql + sqlw1 + " GROUP BY MaCT"
        SetSQL "QDuPhong", "SELECT MaCT FROM " + ChungTu2TKNC(-1) + " WHERE MaKH=" + CStr(mkh) + " AND " + wsql + sqlw1 + " GROUP BY MaCT"
        sql = "SELECT DISTINCTROW 1 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,IIF(MaVattu>0,Vattu.SoHieu,HethongTK.SoHieu) AS SoHieu,TenVattu,DonVi,Sum(SoPS2No) AS Luong,Sum(SoPS) AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",First(ChungTu.MaSo) AS MS,IIF((HethongTK.Loai=1 OR HethongTK.Loai=6) AND HethongTK.TK_ID<>" + CStr(GTGTKT_ID) + "," + CStr(TKVT_ID) + ",TK_ID) FROM ((" + ChungTu2TKNC(-1) + ") LEFT JOIN  Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT" _
            & " WHERE (TK_ID=" + CStr(TKVT_ID) + " OR TK_ID=" + CStr(GTGTKT_ID) + " OR HethongTK.SoHieu LIKE '138*' OR HethongTK.SoHieu LIKE '338*' OR ((HethongTK.Loai=1 OR HethongTK.Loai=6) AND MaKHC=" + CStr(mkh) + ")) GROUP BY NgayCT,ChungTu.SoHieu,IIF(MaVattu>0,Vattu.SoHieu,HethongTK.SoHieu),TenVattu,DonVi,DienGiai" + IIf(nn > 0, "E", "") + ",IIF((HethongTK.Loai=1 OR HethongTK.Loai=6) AND HethongTK.TK_ID<>" + CStr(GTGTKT_ID) + "," + CStr(TKVT_ID) + ",TK_ID)"
        sql = sql + " UNION SELECT DISTINCTROW 1 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,'TSCD' AS SH,ChungTu.DienGiai,'...' AS Dvt,1 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM (" + ChungTu2TKNC(-1) + ") INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT" _
            & " WHERE (TK_ID=" + CStr(TSCD_ID) + ")"
        sql = sql + " UNION SELECT DISTINCTROW 1 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,'' AS SH,ChungTu.DienGiai,'...' AS Dvt,0 AS Luong,SoPS-TPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,TK_ID FROM (" + ChungTu2TKNC(1) + ") INNER JOIN NhatKy ON ChungTu.MaCT=NhatKy.MaCT" _
            & " WHERE MaTKNo=0 AND MaKHC=" + CStr(mkh) + " AND SoPS<>TPS AND MaLoai=1 AND InStr(ChungTu.GhiChu,'138')>0 AND " + wsql + sqlw1
        
        sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM " + ChungTu2TKNC(0) _
            & " WHERE (TK.SoHieu LIKE '11*' OR TK.SoHieu LIKE '141*' OR TK.SoHieu LIKE '" + pSHPT + "*' OR TK.SoHieu LIKE '311*' OR TK.SoHieu LIKE '336*' OR TK.SoHieu LIKE '15*'  OR TK.SoHieu LIKE '515*'  OR TK.SoHieu LIKE '511*'  OR TK.SoHieu LIKE '711*' OR TK.SoHieu LIKE '3331*' OR TK.SoHieu LIKE '" + pVATV + "*' OR TK.SoHieu LIKE '413*' OR TK.Loai=6) AND MaKH=" + CStr(mkh) + " AND " + wsql + sqlw1
        
        sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai,ChungTu.MaSo,0 AS TK_ID FROM (" + ChungTu2TKNC(1) + ") INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT" _
            & " WHERE (HethongTK.SoHieu LIKE '11*' OR HethongTK.SoHieu LIKE '136*') AND MaTKNo=0"
        sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM " + ChungTu2TKNC(-1) _
            & " WHERE MaKHC=MaKH AND MaKH=" + CStr(mkh) + " AND TK_ID=" + CStr(TKCNPT_ID) + " AND " + wsql + sqlw1
        sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM (" + ChungTu2TKNC(0) + ") INNER JOIN QDuPhong ON ChungTu.MaCT=QDuPhong.MaCT" _
            & " WHERE (HethongTK.SoHieu LIKE '635*' AND TK.SoHieu LIKE '11*')" + sqlw1
        sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,TK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM (" + ChungTu2TKNC(0) + ") INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT" _
            & " WHERE (HethongTK.SoHieu LIKE '152*' AND TK.SoHieu LIKE '11*')" + sqlw1
        SetSQL "QChitiet", sql
        frmMain.Rpt.Formulas(4) = "TenTK = IF PageNumber() = 1 THEN '" + ABCtoVNI("Bªn cung cÊp: ") + kh.sohieu + " - " + kh.Ten + "'"
        frmMain.Rpt.Formulas(5) = "SoDu=" + DoiDau(DC - dn)
        X = SelectSQL("SELECT Sum(IIF(Loai=1,Tien,-Tien)) AS F1,Sum(ABS(Tien)) AS F2 FROM QChitiet", Y)
        frmMain.Rpt.Formulas(6) = "ThanhTienX='" + ToVNText(DC - dn + X) + " ®ång'"
        If DC = 0 And dn = 0 And Y = 0 Then GoTo Z
    Else
Z:
        SetSQL "MienTru", "SELECT MaCT FROM " + ChungTu2TKNC(-1) + " WHERE (MaLoai=8 OR MaLoai=0) AND MaKH=" + CStr(mkh) + " AND " + wsql + sqlw1 + " GROUP BY MaCT"
        SetSQL "QDuPhong", "SELECT MaCT FROM " + ChungTu2TKNC(1) + " WHERE MaKHC=" + CStr(mkh) + " AND " + wsql + sqlw1 + " GROUP BY MaCT"
        sql = "SELECT DISTINCTROW 1 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,IIF(MaVattu>0,Vattu.SoHieu,HethongTK.SoHieu) AS SoHieu,TenVattu,DonVi,SoPS2Co AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,TK_ID FROM ((" + ChungTu2TKNC(1) + ") LEFT JOIN  Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT" _
            & " WHERE TK_ID=" + CStr(TKDT_ID) + " OR TK_ID=" + CStr(TKTNBT_ID) + " OR TK_ID=" + CStr(GTGTPN_ID) + " OR HethongTK.Loai=7 OR HethongTK.SoHieu LIKE '138*' OR HethongTK.SoHieu LIKE '338*'"
        sql = sql + " UNION SELECT 1 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM " + ChungTu2TKNC(-1) _
            & " WHERE MaKHC<>MaKH AND MaKH=" + CStr(mkh) + " AND TK_ID=" + CStr(TKCNKH_ID) + " AND MaTKTCNo=MaTKTCCo AND " + wsql + sqlw1
        sql = sql + " UNION SELECT 1 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,'~' AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,-SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,TK_ID FROM (" + ChungTu2TKNC(-1) + ") INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT" _
            & " WHERE HethongTK.SoHieu LIKE '521*' AND " + wsql + sqlw1
        
        sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM " + ChungTu2TKNC(0) _
            & " WHERE (HethongTK.SoHieu LIKE '11*' OR HethongTK.SoHieu LIKE '15*' OR HethongTK.SoHieu LIKE '" + pVATV + "*' OR HethongTK.SoHieu LIKE '3331*' OR HethongTK.SoHieu LIKE '136*' OR HethongTK.SoHieu LIKE '138*' OR HethongTK.SoHieu LIKE '336*' OR HethongTK.SoHieu LIKE '338*' OR HethongTK.SoHieu LIKE '331*' OR HethongTK.SoHieu LIKE '413*' OR HethongTK.Loai=5 OR HethongTK.Loai=6) AND MaKHC=" + CStr(mkh) + " AND " + wsql + sqlw2
        
        sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM (" + ChungTu2TKNC(-1) + ") INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT" _
            & " WHERE (HethongTK.SoHieu LIKE '11*' OR HethongTK.SoHieu LIKE '15*' OR HethongTK.SoHieu LIKE '136*' OR HethongTK.SoHieu LIKE '138*' OR HethongTK.SoHieu LIKE '336*' OR HethongTK.SoHieu LIKE '338*') AND MaTKCo=0"
        sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM " + ChungTu2TKNC(0) _
            & " WHERE MaKHC=MaKH AND MaKH=" + CStr(mkh) + " AND HethongTK.TK_ID=" + CStr(TKCNPT_ID) + " AND " + wsql + sqlw2
        sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM " + ChungTu2TKNC(-1) _
            & " WHERE MaKHC<>MaKH AND MaKHC=" + CStr(mkh) + " AND TK_ID=" + CStr(TKCNKH_ID) + " AND MaTKTCNo=MaTKTCCo AND " + wsql + sqlw1
        'bo vi khong lay tai khoan 515 chenh lech ti gia
        'sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM (" + ChungTu2TKNC(0) + ") INNER JOIN QDuPhong ON ChungTu.MaCT=QDuPhong.MaCT" _
        '    & " WHERE (HethongTK.SoHieu LIKE '11*' AND TK.SoHieu LIKE '515*')" + sqlw1
        
        sql = sql + " UNION SELECT 2 AS Loai,NgayCT,ChungTu.SoHieu AS SHCT,HethongTK.SoHieu AS SH,'...' AS TenVT,'...' AS Dvt,0 AS Luong,-SoPS AS Tien,DienGiai" + IIf(nn > 0, "E", "") + ",ChungTu.MaSo,0 AS TK_ID FROM (" + ChungTu2TKNC(0) + ") INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT" _
            & " WHERE TK.SoHieu LIKE '11*' AND MaKH=" + CStr(mkh) + " AND HethongTK.TK_ID=" + CStr(TKCNKH_ID) + sqlw1
        SetSQL "QChitiet", sql
        frmMain.Rpt.Formulas(4) = "TenTK = IF PageNumber() = 1 THEN '" + ABCtoVNI("Bªn nhËn hµng: ") + kh.sohieu + " - " + kh.Ten + "'"
        frmMain.Rpt.Formulas(5) = "SoDu=" + DoiDau(dn - DC)
        X = SelectSQL("SELECT Sum(IIF(Loai=1,Tien,-Tien)) AS F1,Sum(ABS(Tien)) AS F2 FROM QChitiet", Y)
        frmMain.Rpt.Formulas(6) = "ThanhTienX='" + ToVNText(dn - DC + X) + " ®ång'"
        If DC = 0 And dn = 0 And Y = 0 Then GoTo KhongIn
    End If
    If X = 0 And DC = 0 And dn = 0 And Y = 0 Then
        SetSQL "QChitiet", "SELECT DISTINCTROW 1 AS Loai,NgayCT,'...' AS SHCT,'...' AS SH,'...' AS TenVattu,'...' AS DonVi,0 AS Luong,0 AS Tien,'...',0,0 FROM ChungTu WHERE MaCT=0"
    End If
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    frmMain.Rpt.ReportFileName = "DCCN" + IIf(ct = 0, "2", "") + ".RPT"
    RptSetDate ncuoi, nn
    Exit Function
KhongIn:
    DoiChieuCN = False
End Function
'======================================================================================
' Thñ tôc lËp b¶ng kª lòy kÕ nhËp xuÊt tån
'======================================================================================
Private Sub InLuyKeX(mkho As Long, mtk As Long, mloai As Long, tdau As Integer, tcuoi As Integer)
    Dim sql As String, st As String, sqlln As String, sqllx As String, sqltn As String, sqltx As String, i As Integer
    Dim rs As Recordset, T As Double, l1 As Double, tx As Double, lx As Double, j As Integer
    
    Recycle pCurDir + "LUYKE.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\LUYKE.XLS", pCurDir + "LUYKE.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "LUYKE.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    If mkho > 0 Then xlsheet.Cells(3, 1) = CboKho(0).Text
    xlsheet.Cells(5, 1) = ThoiGian(tdau, tcuoi, nn)
    
    GauGe.Max = 3
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqlln = sqlln + " + TonKho.Luong_Nhap_" + st
        sqllx = sqllx + " + TonKho.Luong_Xuat_" + st
        sqltn = sqltn + " + TonKho.Tien_Nhap_" + st
        sqltx = sqltx + " + TonKho.Tien_Xuat_" + st
    Next
    GauGe.Value = 1
    st = CStr(CThangDB(ThangTruoc(tdau)))
    sql = "SELECT DISTINCTROW HeThongTK.MaSo AS MSTK, HeThongTK.SoHieu, HeThongTK.Ten, PhanLoaiVattu.MaSo," _
        & " PhanLoaiVattu.SoHieu, PhanLoaiVattu.TenPhanLoai, Vattu.MaSo AS MSVT, Vattu.SoHieu AS SHVT, Vattu.TenVattu, Vattu.DonVi, Vattu.DonVi2, Vattu.Dvt2, Sum(TonKho.Tien_" + st + ") AS DkTien," _
        & " Sum(TonKho.Luong_" + st + ") AS DkLuong, Sum(" + sqltn + ") AS TienNhap, Sum(" + sqlln + ") AS LuongNhap, Sum(" + sqltx + ") AS TienXuat," _
        & " Sum(" + sqllx + ") AS LuongXuat, Sum(TonKho.Luong_" + CStr(CThangDB(tcuoi)) + ") AS CkLuong, Sum(TonKho.Tien_" + CStr(CThangDB(tcuoi)) + ") AS CkTien" _
        & " FROM HeThongTK RIGHT JOIN ((PhanLoaiVattu RIGHT JOIN Vattu ON PhanLoaiVattu.MaSo = Vattu.MaPhanLoai) RIGHT JOIN TonKho ON Vattu.MaSo = TonKho.MaVatTu) ON HeThongTK.MaSo = TonKho.MaTaiKhoan" _
        & " GROUP BY HeThongTK.MaSo, HeThongTK.SoHieu, HeThongTK.Ten, PhanLoaiVattu.MaSo, PhanLoaiVattu.SoHieu, PhanLoaiVattu.TenPhanLoai, Vattu.MaSo, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi2,Vattu.DonVi,Vattu.Dvt2 " _
        + IIf(mkho > 0, ", TonKho.MaSoKho", "") + " Having (Sum(TonKho.Luong_" + CStr(CThangDB(ThangTruoc(tdau))) + ")<>0 OR Sum(TonKho.Tien_" + CStr(CThangDB(ThangTruoc(tdau))) + ")<>0 OR SUM(" + sqlln + ")<>0 OR SUM(" + sqltn + ")<>0 OR SUM(" + sqllx + ")<>0 OR SUM(" + sqltx + ")<>0)" _
        + IIf(mkho <> 0, " AND TonKho.MaSoKho = " + CStr(mkho), "") + IIf(mtk <> 0, " AND HethongTK.MaSo = " + CStr(mtk), "") _
        + IIf(mloai <> 0, " AND PhanLoaiVattu.SoHieu LIKE '" + MaSo2SoHieu(mloai, "PhanLoaiVattu") + "*'", "") + " ORDER BY HeThongTK.SoHieu, PhanLoaiVattu.SoHieu, Vattu.SoHieu"
    i = 0
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
            i = i + 1
            xlsheet.Cells(i + 9, 1) = CStr(i)
            xlsheet.Cells(i + 9, 2) = rs!TenVattu
            xlsheet.Cells(i + 9, 3) = rs!shvt
            xlsheet.Cells(i + 9, 4) = rs!DonVi
            
            If rs!DkLuong <> 0 Then xlsheet.Cells(i + 9, 5) = rs!DkLuong
            If rs!Dktien <> 0 Then xlsheet.Cells(i + 9, 6) = rs!Dktien
            
            If rs!LuongNhap <> 0 Then xlsheet.Cells(i + 9, 13) = rs!LuongNhap
            If rs!TienNhap <> 0 Then xlsheet.Cells(i + 9, 14) = rs!TienNhap
            
            If rs!LuongXuat <> 0 Then xlsheet.Cells(i + 9, 23) = rs!LuongXuat
            If rs!TienXuat <> 0 Then xlsheet.Cells(i + 9, 24) = rs!TienXuat
            
            If rs!ckluong <> 0 Then xlsheet.Cells(i + 9, 25) = rs!ckluong
            If rs!cktien <> 0 Then xlsheet.Cells(i + 9, 26) = rs!cktien
            
            NhapTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "331", lx, tx, mkho
            NhapTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "11", l1, T, mkho
            lx = lx + l1
            tx = tx + T
            If lx <> 0 Then xlsheet.Cells(i + 9, 7) = lx
            If tx <> 0 Then xlsheet.Cells(i + 9, 8) = tx
            
            NhapTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "136", lx, tx, mkho
            NhapTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "336", l1, T, mkho
            lx = lx + l1
            tx = tx + T
            
            NhapTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "138", l1, T, mkho
            lx = lx + l1
            tx = tx + T
            NhapTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "338", l1, T, mkho
            lx = lx + l1
            tx = tx + T
            NhapTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "154", l1, T, mkho
            lx = lx + l1
            tx = tx + T
            NhapTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "152", l1, T, mkho
            lx = lx + l1
            tx = tx + T
            NhapTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "155", l1, T, mkho
            lx = lx + l1
            tx = tx + T
            If lx <> 0 Then xlsheet.Cells(i + 9, 9) = lx
            If tx <> 0 Then xlsheet.Cells(i + 9, 10) = tx
            
            XuatTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "621", lx, tx, mkho
            XuatTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "627", l1, T, mkho
            lx = lx + l1
            tx = tx + T
            If lx <> 0 Then xlsheet.Cells(i + 9, 15) = lx
            If tx <> 0 Then xlsheet.Cells(i + 9, 16) = tx
            
            XuatTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "632", lx, tx, mkho
            If lx <> 0 Then xlsheet.Cells(i + 9, 17) = lx
            If tx <> 0 Then xlsheet.Cells(i + 9, 18) = tx
            
            XuatTheoNguon tdau, tcuoi, rs!MSVT, rs!MSTK, "154", lx, tx, mkho
            If lx <> 0 Then xlsheet.Cells(i + 9, 19) = lx
            If tx <> 0 Then xlsheet.Cells(i + 9, 20) = tx
            
            xlsheet.Range("K" + CStr(i + 9)).Formula = "=M" + CStr(i + 9) + "-G" + CStr(i + 9) + "-I" + CStr(i + 9)
            xlsheet.Range("L" + CStr(i + 9)).Formula = "=N" + CStr(i + 9) + "-H" + CStr(i + 9) + "-J" + CStr(i + 9)
            
            xlsheet.Range("U" + CStr(i + 9)).Formula = "=W" + CStr(i + 9) + "-O" + CStr(i + 9) + "-Q" + CStr(i + 9) + "-S" + CStr(i + 9)
            xlsheet.Range("V" + CStr(i + 9)).Formula = "=X" + CStr(i + 9) + "-P" + CStr(i + 9) + "-R" + CStr(i + 9) + "-T" + CStr(i + 9)
            
            rs.MoveNext
    Loop
    GauGe.Value = 2
    rs.Close
    
    xlsheet.Cells(i + 10, 2) = ABCtoVNI("Tæng céng")
    xlsheet.Range("F" + CStr(i + 10)).Formula = "=SUM(F10:F" + CStr(i + 9) + ")"
    xlsheet.Range("G" + CStr(i + 10)).Formula = "=SUM(G10:G" + CStr(i + 9) + ")"
    xlsheet.Range("H" + CStr(i + 10)).Formula = "=SUM(H10:H" + CStr(i + 9) + ")"
    xlsheet.Range("I" + CStr(i + 10)).Formula = "=SUM(I10:I" + CStr(i + 9) + ")"
    xlsheet.Range("J" + CStr(i + 10)).Formula = "=SUM(J10:J" + CStr(i + 9) + ")"
    xlsheet.Range("K" + CStr(i + 10)).Formula = "=SUM(K10:K" + CStr(i + 9) + ")"
    xlsheet.Range("L" + CStr(i + 10)).Formula = "=SUM(L10:L" + CStr(i + 9) + ")"
    xlsheet.Range("M" + CStr(i + 10)).Formula = "=SUM(M10:M" + CStr(i + 9) + ")"
    xlsheet.Range("N" + CStr(i + 10)).Formula = "=SUM(N10:N" + CStr(i + 9) + ")"
    xlsheet.Range("O" + CStr(i + 10)).Formula = "=SUM(O10:O" + CStr(i + 9) + ")"
    xlsheet.Range("P" + CStr(i + 10)).Formula = "=SUM(P10:P" + CStr(i + 9) + ")"
    xlsheet.Range("Q" + CStr(i + 10)).Formula = "=SUM(Q10:Q" + CStr(i + 9) + ")"
    xlsheet.Range("R" + CStr(i + 10)).Formula = "=SUM(R10:R" + CStr(i + 9) + ")"
    xlsheet.Range("S" + CStr(i + 10)).Formula = "=SUM(S10:S" + CStr(i + 9) + ")"
    xlsheet.Range("T" + CStr(i + 10)).Formula = "=SUM(T10:T" + CStr(i + 9) + ")"
    xlsheet.Range("U" + CStr(i + 10)).Formula = "=SUM(U10:U" + CStr(i + 9) + ")"
    xlsheet.Range("V" + CStr(i + 10)).Formula = "=SUM(V10:V" + CStr(i + 9) + ")"
    xlsheet.Range("W" + CStr(i + 10)).Formula = "=SUM(W10:W" + CStr(i + 9) + ")"
    xlsheet.Range("X" + CStr(i + 10)).Formula = "=SUM(X10:X" + CStr(i + 9) + ")"
    xlsheet.Range("Y" + CStr(i + 10)).Formula = "=SUM(Y10:Y" + CStr(i + 9) + ")"
    xlsheet.Range("Z" + CStr(i + 10)).Formula = "=SUM(Z10:Z" + CStr(i + 9) + ")"
    
    With xlsheet.Range("A7", "Z" + CStr(i + 10))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    GauGe.Value = 3
    CallExcel "LUYKE.XLS"
KetThuc:
    
    Set rs = Nothing
End Sub

Private Sub InBangThopCNN(ndau As Date, ncuoi As Date, shtk As String, mpl As Long, Optional th As Integer = 0)
    Dim rs As Recordset, dkno As Double, dkco As Double, psn As Double, psnx As Double, psc As Double, pscx As Double, ckn As Double, ckc As Double, X As Double, Y As Double
        
    GauGe.Max = 3
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,MaCha,BC_ID) SELECT MaTaiKhoan*100000+MaKhachHang AS MS,Cstr(MaTaiKhoan*100000+MaKhachHang) AS SH,MaTaiKhoan,MaKhachHang FROM (SoDuKhachHang INNER JOIN KhachHang ON SoDuKhachHang.MaKhachHang=KhachHang.MaSo) INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo " _
        & " WHERE LEFT(KhachHang.SoHieu,1)<>'#'" + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + IIf(mpl > 0, " AND MaPhanLoai=" + CStr(mpl), "")
    Set rs = DBKetoan.OpenRecordset("BaoCaoCP", dbOpenSnapshot, dbForwardOnly)
    GauGe.Value = 1
    Do While Not rs.EOF
        SoDuKHNgay rs!BC_ID, ndau - 1, dkno, dkco, X, rs!macha
        SoPhatSinhN rs!BC_ID, ndau, ncuoi, psn, psc, X, Y, rs!macha
        SoPhatSinhN rs!BC_ID, NgayDauThang(pNamTC, pThangDauKy), ncuoi, psnx, pscx, X, Y, rs!macha
        SoDuKHNgay rs!BC_ID, ncuoi, ckn, ckc, X, rs!macha
        ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1=" + DoiDau(dkno) + ",Kq2=" + DoiDau(dkco) + ",Kq3=" + DoiDau(psn) + ",Kq4=" + DoiDau(psnx) + ",Kq5=" + DoiDau(psc) + ",Kq6=" + DoiDau(pscx) + ",Kq7=" + DoiDau(ckn) + ",Kq8=" + DoiDau(ckc) + " WHERE MaSo=" + CStr(rs!MaSo)
        rs.MoveNext
    Loop
    ExecuteSQL5 "DELETE * FROM BaoCaoCP WHERE Kq1=0 AND Kq2=0 AND Kq4=0 AND Kq6=0"
    
    'QTongHopCT
    SetSQL "QChitiet", "SELECT DISTINCTROW KhachHang.SoHieu, KhachHang.Ten, PhanLoaiKhachHang.SoHieu AS SHPL, TenPhanLoai,KyHieu," _
        & " Kq1 AS DkNo, Kq2 AS DkCo, 0 AS DkNT, Kq3 AS PsNo, 0 AS PsNo2, Kq5 AS PsCo, 0 AS PsCo2, Kq4 AS PsNoX, 0 AS PsNoX2, Kq6 AS PsCoX, 0 AS PsCoX2, Kq7 AS CkNo, Kq8 AS CkCo, 0 AS CkNT,PhanLoaiKhachHang.PLCha" _
        & " FROM ((KhachHang INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo) INNER JOIN BaoCaoCP ON KhachHang.MaSo=BaoCaoCP.BC_ID) LEFT JOIN NguyenTe ON KhachHang.MaNT=NguyenTe.MaSo " _
        & " ORDER BY PhanLoaiKhachHang.Sohieu,KhachHang.SoHieu"
    SetSQL "QSoCai", "SELECT QChitiet.*, PhanLoaiKhachHang.SoHieu AS SH2,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai2,PhanLoaiKhachHang.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoaiKhachHang ON QChitiet.PLCha=PhanLoaiKhachHang.MaSo"
    SetSQL "QTongHopCT", "SELECT QSoCai.*, PhanLoaiKhachHang.SoHieu AS SH1,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai1 FROM QSoCai LEFT JOIN PhanLoaiKhachHang ON QSoCai.PLCha2=PhanLoaiKhachHang.MaSo"
    
    GauGe.Value = 2
    
    frmMain.Rpt.ReportFileName = "THSDCN" + IIf(th <> 0, "1", "") + ".RPT"
    RptSetDate ncuoi
    If Len(shtk) > 0 Then
        frmMain.Rpt.Formulas(5) = "TenBaoCao = IF PageNumber() = 1 THEN '" + ABCtoVNI("Tµi kho¶n: ") + shtk + " - " + LbTenTk(7).Caption + "'"
    End If
    frmMain.Rpt.Formulas(6) = "ThoiGian = IF PageNumber() = 1 THEN '" + ThoiGianN(ndau, ncuoi, nn) + "'"
    GauGe.Value = 3
End Sub

Private Sub InLCNoiBo(mkho As Long, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date)
    Dim sql As String, i As Integer, sokho As Integer, sovt As Integer, sh1 As String
    Dim rs As Recordset, T As Double, L As Double, j As Integer, st As String, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
        
    sql = "SELECT DISTINCTROW MaNguon FROM ChungTu WHERE MaLoai=4 AND MaKho=" + CStr(mkho) + " AND " + wsql + " GROUP BY MaNguon"
    i = 0
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.recordCount = 0 Then GoTo KetThuc
    
    GauGe.Max = 5
    Recycle pCurDir + "LCNoiBo.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\LCNoiBo.XLS", pCurDir + "LCNoiBo.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "LCNoiBo.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = CboKho(0).Text
    xlsheet.Cells(5, 1) = ThoiGianN(ndau, ncuoi, nn)
    
    Do While Not rs.EOF
            i = i + 1
            xlsheet.Cells(7, 5 + 2 * i) = CStr(rs!MaNguon)
            xlsheet.Cells(8, 5 + 2 * i) = "L­îng"
            xlsheet.Cells(8, 6 + 2 * i) = "TiÒn"
            xlsheet.Range(XLSCol(5 + 2 * i) + "7", XLSCol(6 + 2 * i) + "7").MergeCells = True
            rs.MoveNext
    Loop
    sokho = i
    GauGe.Value = 2
    
    sql = "SELECT DISTINCTROW MaVattu FROM ChungTu WHERE MaLoai=4 AND MaKho=" + CStr(mkho) + " AND " + wsql + " GROUP BY MaVattu"
    i = 0
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
            i = i + 1
            xlsheet.Cells(8 + i, 2) = rs!MaVattu
            rs.MoveNext
    Loop
    sovt = i
    xlsheet.Cells(sovt + 9, 2) = ABCtoVNI("Tæng céng")
    GauGe.Value = 3
    
    For i = 1 To sovt
        sh1 = ""
        st = ""
        For j = 1 To sokho
            xlsheet.Cells(8 + i, 5 + 2 * j) = Format(SelectSQL("SELECT SUM(SoPS2Co) AS F1,Sum(SoPS) AS F2 FROM ChungTu WHERE MaLoai=4 AND MaKho=" + CStr(mkho) _
                + " AND " + wsql + " AND MaNguon=" + CStr(xlsheet.Cells(7, 5 + 2 * j)) + " AND MaVattu=" + CStr(xlsheet.Cells(8 + i, 2)), T), Mask_2)
            xlsheet.Cells(8 + i, 6 + 2 * j) = DoiDau(T)
            sh1 = sh1 + "+" + XLSCol(5 + 2 * j) + CStr(8 + i)
            st = st + "+" + XLSCol(6 + 2 * j) + CStr(8 + i)
        Next
        xlsheet.Range("E" + CStr(8 + i)).Formula = "=" + sh1
        xlsheet.Range("F" + CStr(8 + i)).Formula = "=" + st
        sh1 = ""
        xlsheet.Cells(8 + i, 1) = CStr(i)
        xlsheet.Cells(8 + i, 2) = TenVT(sh1, CLng5(xlsheet.Cells(8 + i, 2)), st)
        xlsheet.Cells(8 + i, 3) = sh1
        xlsheet.Cells(8 + i, 4) = st
    Next
    GauGe.Value = 4
    
    For i = 1 To sokho
        xlsheet.Range(XLSCol(6 + 2 * i) + CStr(sovt + 9)).Formula = "=SUM(" + XLSCol(6 + 2 * i) + "9:" + XLSCol(6 + 2 * i) + CStr(sovt + 8) + ")"
        xlsheet.Cells(7, 5 + 2 * i) = tenkho(CLng5(xlsheet.Cells(7, 5 + 2 * i)))
    Next
    xlsheet.Range("F" + CStr(sovt + 9)).Formula = "=SUM(F9:F" + CStr(sovt + 8) + ")"
    With xlsheet.Range("A7", XLSCol(2 * sokho + 6) + CStr(sovt + 9))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    GauGe.Value = 3
    CallExcel "LCNoiBo.XLS"
KetThuc:
    rs.Close
    Set rs = Nothing
End Sub

Private Sub InBKGiaThanh(thang As Integer)
    Dim sql As String, i As Integer, sonvl As Integer, sotp As Integer, sh1 As String
    Dim rs As Recordset, T As Double, L As Double, j As Integer, st As String, dgia As Double
        
    sql = "SELECT DISTINCTROW MaNVL FROM ThanhPham INNER JOIN DinhMuc ON ThanhPham.MaTP=DinhMuc.MaTP WHERE ThanhPham.Thang=" + CStr(thang) + " AND MaNVL>0 GROUP BY MaNVL"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.recordCount = 0 Then GoTo KetThuc
    
    GauGe.Max = 5
    Recycle pCurDir + "GiaThanh.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\GiaThanh.XLS", pCurDir + "GiaThanh.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "GiaThanh.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = CboKho(0).Text
    xlsheet.Cells(5, 1) = ThoiGian(thang, thang, nn)
    
    i = 0
    Do While Not rs.EOF
            i = i + 1
            xlsheet.Cells(7, 4 + 2 * i) = CStr(rs!MaNVL)
            xlsheet.Cells(8, 4 + 2 * i) = "L­îng"
            xlsheet.Cells(8, 5 + 2 * i) = "TiÒn"
            xlsheet.Range(XLSCol(4 + 2 * i) + "7", XLSCol(5 + 2 * i) + "7").MergeCells = True
            rs.MoveNext
    Loop
    sonvl = i
    xlsheet.Cells(7, 6 + 2 * sonvl) = "Tæng NVL"
    xlsheet.Cells(7, 7 + 2 * sonvl) = "Nh©n c«ng"
    xlsheet.Cells(7, 8 + 2 * sonvl) = "KH TSC§"
    xlsheet.Cells(7, 9 + 2 * sonvl) = "CP SX kh¸c"
    xlsheet.Cells(7, 10 + 2 * sonvl) = "Tæng CP"
    xlsheet.Cells(7, 11 + 2 * sonvl) = "§¬n gi¸"
    For i = 0 To 5
        xlsheet.Range(XLSCol(6 + i + 2 * sonvl) + "7", XLSCol(6 + i + 2 * sonvl) + "8").MergeCells = True
    Next
    GauGe.Value = 2
    
    
    sql = "SELECT DISTINCTROW MaTP,Sum(SoLuong) AS Luong FROM ThanhPham WHERE Thang=" + CStr(thang) + " GROUP BY MaTP"
    i = 0
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
            i = i + 1
            xlsheet.Cells(8 + i, 2) = rs!MaTP
            xlsheet.Cells(8 + i, 5) = Format(rs!luong, Mask_2)
            rs.MoveNext
    Loop
    sotp = i
    xlsheet.Cells(sotp + 9, 2) = ABCtoVNI("Tæng céng")
    GauGe.Value = 3
    
    For i = 1 To sonvl
        dgia = SelectSQL("SELECT SoPS/SoPS2Co AS F1 FROM ChungTu WHERE MaLoai=2 AND ThangCT=" + CStr(thang) + " AND CT_ID>610000000 AND MaVattu=" + CStr(xlsheet.Cells(7, 4 + 2 * i)))
        For j = 1 To sotp
            L = SelectSQL("SELECT Sum(ThanhPham.SoLuong*DinhMuc.SoLuong) AS F1 FROM ThanhPham INNER JOIN DinhMuc ON ThanhPham.MaTP=DinhMuc.MaTP " _
                & "WHERE ThanhPham.Thang=" + CStr(thang) + " AND ThanhPham.MaTP=" + CStr(xlsheet.Cells(8 + j, 2)) + " AND MaNVL=" + CStr(xlsheet.Cells(7, 4 + 2 * i)))
            xlsheet.Cells(8 + j, 4 + 2 * i) = Format(L, Mask_2)
            T = RoundMoney(L * dgia)
            xlsheet.Cells(8 + j, 5 + 2 * i) = Format(T, Mask_0)
        Next
        xlsheet.Range(XLSCol(5 + 2 * i) + CStr(sotp + 9)).Formula = "=SUM(" + XLSCol(5 + 2 * i) + "9:" + XLSCol(5 + 2 * i) + CStr(sotp + 8) + ")"
        xlsheet.Cells(7, 4 + 2 * i) = CStr(SelectSQL("SELECT SoHieu AS F1 FROM Vattu WHERE MaSo=" + CStr(xlsheet.Cells(7, 4 + 2 * i))))
    Next
    For i = 1 To 5
        xlsheet.Range(XLSCol(5 + 2 * sonvl + i) + CStr(sotp + 9)).Formula = "=SUM(" + XLSCol(5 + 2 * sonvl + i) + "9:" + XLSCol(5 + 2 * sonvl + i) + CStr(sotp + 8) + ")"
    Next
    
    GauGe.Value = 4
        
    For i = 1 To sotp
        sh1 = ""
        For j = 1 To sonvl
            sh1 = sh1 + " + " + XLSCol(5 + 2 * j) + CStr(i + 8)
        Next
        xlsheet.Range(XLSCol(6 + 2 * sonvl) + CStr(i + 8)).Formula = "=" + sh1
        
        xlsheet.Cells(8 + i, 7 + 2 * sonvl) = Format(SelectSQL("SELECT SUM(CPNC) AS F1 FROM ThanhPham WHERE Thang=" + CStr(thang) + " AND MaTP=" + CStr(xlsheet.Cells(8 + i, 2))), Mask_0)
        xlsheet.Cells(8 + i, 8 + 2 * sonvl) = Format(SelectSQL("SELECT SUM(CPKH) AS F1 FROM ThanhPham WHERE Thang=" + CStr(thang) + " AND MaTP=" + CStr(xlsheet.Cells(8 + i, 2))), Mask_0)
        xlsheet.Cells(8 + i, 9 + 2 * sonvl) = Format(SelectSQL("SELECT SUM(CPSXC) AS F1 FROM ThanhPham WHERE Thang=" + CStr(thang) + " AND MaTP=" + CStr(xlsheet.Cells(8 + i, 2))), Mask_0)
        xlsheet.Range(XLSCol(10 + 2 * sonvl) + CStr(i + 8)).Formula = "=" + XLSCol(6 + 2 * sonvl) + CStr(i + 8) + "+" + XLSCol(7 + 2 * sonvl) + CStr(i + 8) + "+" + XLSCol(8 + 2 * sonvl) + CStr(i + 8) + "+" + XLSCol(9 + 2 * sonvl) + CStr(i + 8)
        xlsheet.Range(XLSCol(11 + 2 * sonvl) + CStr(i + 8)).Formula = "=" + XLSCol(10 + 2 * sonvl) + CStr(i + 8) + "/" + "E" + CStr(i + 8)
        sh1 = ""
        xlsheet.Cells(8 + i, 1) = CStr(i)
        xlsheet.Cells(8 + i, 2) = TenVT(sh1, CLng5(xlsheet.Cells(8 + i, 2)), st)
        xlsheet.Cells(8 + i, 3) = sh1
        xlsheet.Cells(8 + i, 4) = st
    Next
    
    With xlsheet.Range("A7", XLSCol(2 * sonvl + 11) + CStr(sotp + 9))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    GauGe.Value = 3
    CallExcel "GiaThanh.XLS"
KetThuc:
    rs.Close
    Set rs = Nothing
End Sub

Private Sub InBKGiaThanh2(tdau As Integer, tcuoi As Integer)
    Dim sql As String, i As Integer, sonvl As Integer, sotp As Integer, sh1 As String, tp As New Cls154
    Dim rs As Recordset, T As Double, L As Double, j As Integer, st As String, dgia As Double
        
    sql = "SELECT DISTINCTROW MaVattu FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaVattu>0 AND MaTP>0 GROUP BY MaVattu"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.recordCount = 0 Then GoTo KetThuc
    
    GauGe.Max = 5
    Recycle pCurDir + "GiaThanh.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\GiaThanh.XLS", pCurDir + "GiaThanh.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "GiaThanh.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    xlsheet.Cells(3, 1) = CboKho(0).Text
    xlsheet.Cells(5, 1) = ThoiGian(tdau, tcuoi, nn)
    
    i = 0
    Do While Not rs.EOF
            i = i + 1
            xlsheet.Cells(7, 4 + 2 * i) = CStr(rs!MaVattu)
            xlsheet.Cells(8, 4 + 2 * i) = "L­îng"
            xlsheet.Cells(8, 5 + 2 * i) = "TiÒn"
            xlsheet.Range(XLSCol(4 + 2 * i) + "7", XLSCol(5 + 2 * i) + "7").MergeCells = True
            rs.MoveNext
    Loop
    sonvl = i
    xlsheet.Cells(7, 6 + 2 * sonvl) = "Tæng NVL"
    xlsheet.Cells(7, 7 + 2 * sonvl) = "Nh©n c«ng"
    xlsheet.Cells(7, 8 + 2 * sonvl) = "KH TSC§"
    xlsheet.Cells(7, 9 + 2 * sonvl) = "CP SX kh¸c"
    xlsheet.Cells(7, 10 + 2 * sonvl) = "Tæng CP"
    xlsheet.Cells(7, 11 + 2 * sonvl) = "§¬n gi¸"
    For i = 0 To 5
        xlsheet.Range(XLSCol(6 + i + 2 * sonvl) + "7", XLSCol(6 + i + 2 * sonvl) + "8").MergeCells = True
    Next
    GauGe.Value = 2
    
    
    sql = "SELECT DISTINCTROW MaTP FROM ChungTu WHERE " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY MaTP"
    i = 0
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
            i = i + 1
            xlsheet.Cells(8 + i, 2) = rs!MaTP
            xlsheet.Cells(8 + i, 5) = 1
            rs.MoveNext
    Loop
    sotp = i
    xlsheet.Cells(sotp + 9, 2) = ABCtoVNI("Tæng céng")
    GauGe.Value = 3
    
    For i = 1 To sonvl
        For j = 1 To sotp
            L = SelectSQL("SELECT Sum(SoPS2Co) AS F1,Sum(SoPS) AS F2 FROM ChungTu " _
                & " WHERE " + WThang("ThangCT", tdau, tcuoi) + " AND MaTP=" + CStr(xlsheet.Cells(8 + j, 2)) + " AND MaVattu=" + CStr(xlsheet.Cells(7, 4 + 2 * i)), T)
            xlsheet.Cells(8 + j, 4 + 2 * i) = Format(L, Mask_2)
            xlsheet.Cells(8 + j, 5 + 2 * i) = Format(T, Mask_0)
        Next
        xlsheet.Range(XLSCol(5 + 2 * i) + CStr(sotp + 9)).Formula = "=SUM(" + XLSCol(5 + 2 * i) + "9:" + XLSCol(5 + 2 * i) + CStr(sotp + 8) + ")"
        xlsheet.Cells(7, 4 + 2 * i) = CStr(SelectSQL("SELECT SoHieu AS F1 FROM Vattu WHERE MaSo=" + CStr(xlsheet.Cells(7, 4 + 2 * i))))
    Next
    For i = 1 To 5
        xlsheet.Range(XLSCol(5 + 2 * sonvl + i) + CStr(sotp + 9)).Formula = "=SUM(" + XLSCol(5 + 2 * sonvl + i) + "9:" + XLSCol(5 + 2 * sonvl + i) + CStr(sotp + 8) + ")"
    Next
    
    GauGe.Value = 4
    
    For i = 1 To sotp
        sh1 = ""
        For j = 1 To sonvl
            sh1 = sh1 + " + " + XLSCol(5 + 2 * j) + CStr(i + 8)
        Next
        xlsheet.Range(XLSCol(6 + 2 * sonvl) + CStr(i + 8)).Formula = "=" + sh1
        tp.InitTPMaSo CLng5(xlsheet.Cells(8 + i, 2))
        xlsheet.Cells(8 + i, 7 + 2 * sonvl) = Format(tp.SoCPNC(tdau, tcuoi), Mask_0)
        T = tp.SoCPKH(tdau, tcuoi)
        xlsheet.Cells(8 + i, 8 + 2 * sonvl) = Format(T, Mask_0)
        xlsheet.Cells(8 + i, 9 + 2 * sonvl) = Format(tp.SoCPSXC(tdau, tcuoi) - T, Mask_0)
        xlsheet.Range(XLSCol(10 + 2 * sonvl) + CStr(i + 8)).Formula = "=" + XLSCol(6 + 2 * sonvl) + CStr(i + 8) + "+" + XLSCol(7 + 2 * sonvl) + CStr(i + 8) + "+" + XLSCol(8 + 2 * sonvl) + CStr(i + 8) + "+" + XLSCol(9 + 2 * sonvl) + CStr(i + 8)
        xlsheet.Range(XLSCol(11 + 2 * sonvl) + CStr(i + 8)).Formula = "=" + XLSCol(10 + 2 * sonvl) + CStr(i + 8) + "/" + "E" + CStr(i + 8)
        sh1 = ""
        xlsheet.Cells(8 + i, 1) = CStr(i)
        xlsheet.Cells(8 + i, 2) = tp.TenVattu
        xlsheet.Cells(8 + i, 3) = tp.sohieu
        xlsheet.Cells(8 + i, 4) = tp.DonVi
    Next
    
    xlsheet.Range("A4", XLSCol(2 * sonvl + 11) + "4").MergeCells = True
    xlsheet.Range("A5", XLSCol(2 * sonvl + 11) + "5").MergeCells = True
    With xlsheet.Range("A7", XLSCol(2 * sonvl + 11) + CStr(sotp + 9))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    GauGe.Value = 3
    CallExcel "GiaThanh.XLS"
KetThuc:
    rs.Close
    Set rs = Nothing
    Set tp = Nothing
End Sub

Private Sub InBangKeTon(tcuoi As Integer, ngay As Date, mtk As Long, mloai As Long)
    Dim sql As String, i As Integer, sokho As Integer, sovt As Integer, sh1 As String
    Dim rs As Recordset, T As Double, L As Double, j As Integer, st As String, wsql As String
    
    wsql = IIf(tcuoi > 0, WThang("ThangCT", 0, tcuoi), WNgay("NgayGS", CVDate("01/01/80"), ngay))
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    GauGe.Max = 3
    
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,MaCha,BC_ID,Kq1,Kq2,Ten) SELECT First(TonKho.MaSo),Cstr(First(TonKho.MaSo)),MaSoKho,MaVattu,Sum(Luong_0),Sum(Tien_0),Vattu.SoHieu FROM (TonKho INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo) INNER JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo WHERE (True) " _
        + IIf(mtk > 0, " AND MaTaiKhoan=" + CStr(mtk), "") + IIf(mloai > 0, " AND PhanLoaiVattu.SoHieu LIKE '" + MaSo2SoHieu(mloai, "PhanLoaiVattu") + "*'", "") + " GROUP BY MaSoKho,Vattu.SoHieu,MaVattu"
    ' Tinh dau ky
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,MaCha,BC_ID,Kq1,Kq2) SELECT First(ChungTu.MaSo),Cstr(First(ChungTu.MaSo)),IIF(MaLoai=1,MaKho,MaNguon),MaVattu,Sum(SoPS2No),Sum(SoPS) FROM (" + ChungTu2TKNC(-1) + ") INNER JOIN BaoCaoCP ON ChungTu.MaVattu=BaoCaoCP.BC_ID AND ((ChungTu.MaKho=BaoCaoCP.MaCha AND MaLoai=1) OR (ChungTu.MaNguon=BaoCaoCP.MaCha AND MaLoai=4)) " _
        & " WHERE (MaLoai=1 OR MaLoai=4) AND TK_ID=" + CStr(TKVT_ID) + " AND " + wsql + IIf(mtk > 0, " AND MaTKNo=" + CStr(mtk), "") + " GROUP BY IIF(MaLoai=1,MaKho,MaNguon),MaVattu"
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN BaoCaoCP2 ON BaoCaoCP.MaCha=BaoCaoCP2.MaCha AND BaoCaoCP.BC_ID=BaoCaoCP2.BC_ID SET BaoCaoCP.Kq1=BaoCaoCP.Kq1+BaoCaoCP2.Kq1,BaoCaoCP.Kq2=BaoCaoCP.Kq2+BaoCaoCP2.Kq2"
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,MaCha,BC_ID,Kq1,Kq2) SELECT First(ChungTu.MaSo),Cstr(First(ChungTu.MaSo)),MaKho,MaVattu,Sum(SoPS2Co),Sum(SoPS) FROM (" + ChungTu2TKNC(1) + ") INNER JOIN BaoCaoCP ON ChungTu.MaKho=BaoCaoCP.MaCha AND ChungTu.MaVattu=BaoCaoCP.BC_ID " _
        & " WHERE (MaLoai=2 OR MaLoai=4) AND TK_ID=" + CStr(TKVT_ID) + " AND " + wsql + IIf(mtk > 0, " AND MaTKCo=" + CStr(mtk), "") + " GROUP BY MaKho,MaVattu"
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN BaoCaoCP2 ON BaoCaoCP.MaCha=BaoCaoCP2.MaCha AND BaoCaoCP.BC_ID=BaoCaoCP2.BC_ID SET BaoCaoCP.Kq1=BaoCaoCP.Kq1-BaoCaoCP2.Kq1,BaoCaoCP.Kq2=BaoCaoCP.Kq2-BaoCaoCP2.Kq2"
            
    ExecuteSQL5 "DELETE * FROM BaoCaoCP WHERE Kq1=0 AND Kq2=0"
        
    sql = "SELECT DISTINCTROW MaCha FROM BaoCaoCP GROUP BY MaCha"
    i = 0
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.recordCount = 0 Then GoTo KetThuc
    
    GauGe.Max = 5
    Recycle pCurDir + "LCNoiBo.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\LCNoiBo.XLS", pCurDir + "LCNoiBo.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "LCNoiBo.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
       
    xlsheet.Cells(4, 1) = "B¶ng tæng hîp tån kho"
    xlsheet.Cells(5, 1) = IIf(tcuoi > 0, ThoiGianN(NgayCuoiThang(pNamTC, tcuoi), NgayCuoiThang(pNamTC, tcuoi), nn), ThoiGianN(ngay, ngay, nn))
    xlsheet.Cells(7, 5) = "Tæng tån kho"
    
    Do While Not rs.EOF
            i = i + 1
            xlsheet.Cells(7, 5 + 2 * i) = CStr(rs!macha)
            xlsheet.Cells(8, 5 + 2 * i) = "L­îng"
            xlsheet.Cells(8, 6 + 2 * i) = "TiÒn"
            xlsheet.Range(XLSCol(5 + 2 * i) + "7", XLSCol(6 + 2 * i) + "7").MergeCells = True
            rs.MoveNext
    Loop
    sokho = i
    GauGe.Value = 2
    
    sql = "SELECT DISTINCTROW BC_ID,Ten FROM BaoCaoCP GROUP BY BC_ID,Ten ORDER BY Ten"
    i = 0
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
            i = i + 1
            xlsheet.Cells(8 + i, 2) = rs!BC_ID
            rs.MoveNext
    Loop
    sovt = i
    xlsheet.Cells(sovt + 9, 2) = ABCtoVNI("Tæng céng")
    GauGe.Value = 3
    
    For i = 1 To sovt
        sh1 = ""
        st = ""
        For j = 1 To sokho
            xlsheet.Cells(8 + i, 5 + 2 * j) = Format(SelectSQL("SELECT Kq1 AS F1,Kq2 AS F2 FROM BaoCaoCP WHERE MaCha=" + CStr(xlsheet.Cells(7, 5 + 2 * j)) + " AND BC_ID=" + CStr(xlsheet.Cells(8 + i, 2)), T), Mask_2)
            xlsheet.Cells(8 + i, 6 + 2 * j) = DoiDau(T)
            sh1 = sh1 + "+" + XLSCol(5 + 2 * j) + CStr(8 + i)
            st = st + "+" + XLSCol(6 + 2 * j) + CStr(8 + i)
        Next
        xlsheet.Range("E" + CStr(8 + i)).Formula = "=" + sh1
        xlsheet.Range("F" + CStr(8 + i)).Formula = "=" + st
        sh1 = ""
        xlsheet.Cells(8 + i, 1) = CStr(i)
        xlsheet.Cells(8 + i, 2) = TenVT(sh1, CLng5(xlsheet.Cells(8 + i, 2)), st)
        xlsheet.Cells(8 + i, 3) = sh1
        xlsheet.Cells(8 + i, 4) = st
    Next
    GauGe.Value = 4
    
    For i = 1 To sokho
        xlsheet.Range(XLSCol(6 + 2 * i) + CStr(sovt + 9)).Formula = "=SUM(" + XLSCol(6 + 2 * i) + "9:" + XLSCol(6 + 2 * i) + CStr(sovt + 8) + ")"
        xlsheet.Cells(7, 5 + 2 * i) = tenkho(CLng5(xlsheet.Cells(7, 5 + 2 * i)))
    Next
    xlsheet.Range("F" + CStr(sovt + 9)).Formula = "=SUM(F9:F" + CStr(sovt + 8) + ")"
    
    With xlsheet.Range("A7", XLSCol(2 * sokho + 6) + CStr(sovt + 9))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .Font.Size = 6
        .VerticalAlignment = xlVAlignCenter
    End With
    
    xlapp.Workbooks(1).Save
    
    xlapp.Workbooks.Close
    
    Set xlsheet = Nothing
    Set xlapp = Nothing
    GauGe.Value = 3
    CallExcel "LCNoiBo.XLS"
KetThuc:
    rs.Close
    Set rs = Nothing
End Sub

Private Sub THPSNhomTK(sh As String, tdau As Integer, tcuoi As Integer)
    Dim rs As Recordset, taikhoan As New ClsTaikhoan
    Dim dkn As Double, dkc As Double, dknt As Double, psn As Double, psc As Double, psn2 As Double, psc2 As Double, ckn As Double, ckc As Double, cknt As Double
    Dim rs_tk As Recordset, sql As String, inso As Boolean, i As Integer, st As String, j As Integer, s As String
    
    ExecuteSQL5 "DELETE * FROM BKNhomPS"
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaSo FROM HethongTK WHERE SoHieu LIKE '" + sh + "*' AND TKCon=0", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        taikhoan.InitTaikhoanMaSo rs!MaSo
        taikhoan.SoDuTK ThangTruoc(tdau), dkn, dkc, dknt
        psn = 0
        psc = 0
        psn2 = 0
        psc2 = 0
        ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo, MaCT,ShTK,SoPSNo,SoPSCo,DienGiai,SoPS2No,SoPS2Co) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",-1,'" + taikhoan.sohieu + "'," _
            + DoiDau(IIf(dkn - dkc > 0, dkn - dkc, 0)) + "," + DoiDau(IIf(dkn - dkc < 0, dkc - dkn, 0)) + ",'" + IIf(nn > 0, taikhoan.TenE, taikhoan.Ten) + "'," + DoiDau(IIf(dkn - dkc > 0, dknt, 0)) + "," + DoiDau(IIf(dkn - dkc < 0, dknt, 0)) + ")"
        
        sql = "SELECT DISTINCTROW ChungTu.MaCT,ChungTu.SoHieu,ChungTu.NgayCT,ChungTu.DienGiai" + IIf(nn > 0, "E", "") + " AS DG,ChungTu.SoPS,IIF(HethongTK.MaNT>0,SoPS2No,0) AS SoPS2N,IIF(HethongTK_1.MaNT>0,SoPS2Co,0) AS SoPS2C,HethongTK.SoHieu As SHNo, HethongTK_1.SoHieu As SHCo" _
            & " FROM HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) ON HeThongTK_1.MaSo = ChungTu.MaTKCo" _
            & " Where (SoPS<>0 OR SoPS2No<>0 OR SoPS2Co<>0) AND (HeThongTK.SoHieu = '" + taikhoan.sohieu + "') AND " + WThang("ThangCT", tdau, tcuoi) _
            + " AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) ORDER BY ChungTu.NgayCT, ChungTu.MaCT"
        Set rs_tk = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
        Do While Not rs_tk.EOF
            psn = psn + rs_tk!sops
            psn2 = psn2 + rs_tk!SoPS2N
            If IsNull(rs_tk!shco) Then s = "..." Else s = rs_tk!shco
            ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo, MaCT,ShTK,SoPSNo,Ngay,SoHieu,DienGiai,ShDu,SoPS2No) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + "," + CStr(rs_tk!MaCT) + ",'" + taikhoan.sohieu + "'," + DoiDau(rs_tk!sops) _
                + ",#" + Format(rs_tk!NgayCT, Mask_DB) + "#,'" + rs_tk!sohieu + "','" + rs_tk!dg + "','" + s + "'," + DoiDau(rs_tk!SoPS2N) + ")"
            rs_tk.MoveNext
        Loop
        
        sql = "SELECT DISTINCTROW ChungTu.MaCT,ChungTu.SoHieu,ChungTu.NgayCT,ChungTu.DienGiai" + IIf(nn > 0, "E", "") + " AS DG,ChungTu.SoPS,IIF(HethongTK.MaNT>0,SoPS2No,0) AS SoPS2N,IIF(HethongTK_1.MaNT>0,SoPS2Co,0) AS SoPS2C,HethongTK.SoHieu As SHNo, HethongTK_1.SoHieu As SHCo" _
            & " FROM HeThongTK AS HeThongTK_1 INNER JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) ON HeThongTK_1.MaSo = ChungTu.MaTKCo" _
            & " Where (SoPS<>0 OR SoPS2No<>0 OR SoPS2Co<>0) AND (HethongTK_1.SoHieu = '" + taikhoan.sohieu + "') AND " + WThang("ThangCT", tdau, tcuoi) _
            + " AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) ORDER BY ChungTu.NgayCT, ChungTu.MaCT"
        Set rs_tk = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
        Do While Not rs_tk.EOF
            psc = psc + rs_tk!sops
            psc2 = psc2 + rs_tk!SoPS2C
            If IsNull(rs_tk!shno) Then s = "..." Else s = rs_tk!shno
            ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,ShTK,SoPSCo,Ngay,SoHieu,DienGiai,ShDu,SoPS2Co) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + "," + CStr(rs_tk!MaCT) + ",'" + taikhoan.sohieu + "'," + DoiDau(rs_tk!sops) _
                + ",#" + Format(rs_tk!NgayCT, Mask_DB) + "#,'" + rs_tk!sohieu + "','" + rs_tk!dg + "','" + s + "'," + DoiDau(rs_tk!SoPS2C) + ")"
            rs_tk.MoveNext
        Loop
        
        ckn = dkn - dkc + psn - psc
        If ckn > 0 Then
            ckc = 0
        Else
            ckc = -ckn
            ckn = 0
        End If
        cknt = IIf(ckn - ckc > 0, dknt + psn2 - psc2, dknt - psn2 + psc2)

        ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,ShTK,SoPSNo,SoPSCo,DienGiai,SoPS2No,SoPS2Co) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",0,'" + taikhoan.sohieu + "'," _
            + DoiDau(IIf(ckn - ckc > 0, ckn - dkc, 0)) + "," + DoiDau(IIf(ckn - ckc < 0, ckc - ckn, 0)) + ",'" + IIf(nn > 0, taikhoan.TenE, taikhoan.Ten) + "'," + DoiDau(IIf(ckn - ckc > 0, cknt, 0)) + "," + DoiDau(IIf(ckn - ckc < 0, cknt, 0)) + ")"

        If dkn = 0 And dkc = 0 And psn = 0 And psc = 0 And psn2 = 0 And psc2 = 0 Then ExecuteSQL5 "DELETE * FROM BKNhomPS WHERE ShTK='" + taikhoan.sohieu + "'"
KhongIn:
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set taikhoan = Nothing
    frmMain.Rpt.ReportFileName = "THPSCT.RPT"

    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
End Sub

Public Sub BangCDCno(tdau As Integer, tcuoi As Integer, mpl As Long)
    Dim st As String, sqln As String, sqlc As String, sqln2 As String, sqlc2 As String, i As Integer
    Dim sqlnx As String, sqlcx As String, sqlnx2 As String, sqlcx2 As String
    
    GauGe.Max = 3
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + SoDuKhachHang.No_" + st
        sqlc = sqlc + " + SoDuKhachHang.Co_" + st
        sqln2 = sqln2 + " + SoDuKhachHang.No_" + st + "_NT"
        sqlc2 = sqlc2 + " + SoDuKhachHang.Co_" + st + "_NT"
    Next
    For i = 1 To CThangDB(tcuoi)
        st = CStr(i)
        sqlnx = sqlnx + " + SoDuKhachHang.No_" + st
        sqlcx = sqlcx + " + SoDuKhachHang.Co_" + st
        sqlnx2 = sqlnx2 + " + SoDuKhachHang.No_" + st + "_NT"
        sqlcx2 = sqlcx2 + " + SoDuKhachHang.Co_" + st + "_NT"
    Next
    GauGe.Value = 1
    st = CStr(CThangDB(ThangTruoc(tdau)))
    SetSQL "QChitiet", "SELECT DISTINCTROW KhachHang.SoHieu, KhachHang.Ten, PhanLoaiKhachHang.SoHieu AS SHPL, TenPhanLoai,KyHieu," _
        & " SoDuKhachHang.DuNo_" + st + " AS DkNo, SoDuKhachHang.DuCo_" + st + " AS DkCo, SoDuKhachHang.DuNT_" + st + " AS DkNT, (" + sqln + ") AS PsNo, (" + sqln2 + ") AS PsNo2, (" + sqlc + ") AS PsCo, (" + sqlc2 + ") AS PsCo2, (" + sqlnx + ") AS PsNoX, (" + sqlnx2 + ") AS PsNoX2, (" + sqlcx + ") AS PsCoX, (" + sqlcx2 + ") AS PsCoX2, SoDuKhachHang.DuNo_" _
        + CStr(CThangDB(tcuoi)) + " AS CkNo, SoDuKhachHang.DuCo_" + CStr(CThangDB(tcuoi)) + " AS CkCo, SoDuKhachHang.DuNT_" + CStr(CThangDB(tcuoi)) + " AS CkNT,PhanLoaiKhachHang.PLCha,HethongTK.SoHieu AS SHTK,HethongTK.Ten AS TenTK" _
        & " FROM (((KhachHang INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo) INNER JOIN SoDuKhachHang ON KhachHang.MaSo=SoDuKhachHang.MaKhachHang) LEFT JOIN NguyenTe ON KhachHang.MaNT=NguyenTe.MaSo) INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo " _
        & " WHERE (SoDuKhachHang.DuNo_" + st + "<>0 OR SoDuKhachHang.DuCo_" + st + "<>0 OR (" + sqlnx + ") <>0 OR (" + sqlcx + ") <>0)" + IIf(mpl > 0, " AND MaPhanLoai=" + CStr(mpl), "") + " AND LEFT(KhachHang.SoHieu,1)<>'#' ORDER BY PhanLoaiKhachHang.Sohieu,KhachHang.SoHieu"
    SetSQL "QSoCai", "SELECT QChitiet.*, PhanLoaiKhachHang.SoHieu AS SH2,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai2,PhanLoaiKhachHang.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoaiKhachHang ON QChitiet.PLCha=PhanLoaiKhachHang.MaSo"
    SetSQL "QTongHopCT", "SELECT QSoCai.*, PhanLoaiKhachHang.SoHieu AS SH1,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai1 FROM QSoCai LEFT JOIN PhanLoaiKhachHang ON QSoCai.PLCha2=PhanLoaiKhachHang.MaSo"
    
    GauGe.Value = 2
    frmMain.Rpt.ReportFileName = "THSDCNT.RPT"

    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
End Sub

Private Sub THPSNhomCN(taikhoan As ClsTaikhoan, tdau As Integer, tcuoi As Integer)
    Dim rs As Recordset, kh As New ClsKhachHang, psn2 As Double, psc2 As Double
    Dim dkn As Double, dkc As Double, dknt As Double, psn As Double, psc As Double, ckn As Double, ckc As Double, cknt As Double
    Dim rs_tk As Recordset, sql As String, inso As Boolean, i As Integer, st As String, j As Integer, s As String
    
    ExecuteSQL5 "DELETE * FROM BKNhomPS"
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaKhachHang,MaTaiKhoan FROM SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*' GROUP BY MaTaiKhoan,MaKhachHang", dbOpenSnapshot, dbForwardOnly)
    
    Do While Not rs.EOF
        kh.InitKhachHangMaSo rs!MaKhachHang
        kh.SoDuKH ThangTruoc(tdau), dkn, dkc, dknt, rs!MaTaiKhoan

        psn = 0
        psc = 0
        psn2 = 0
        psc2 = 0
                
        ExecuteSQL5 "UPDATE BKNhomPS SET SoPSNo=SoPSNo+" + DoiDau(IIf(dkn - dkc > 0, dkn - dkc, 0)) + ",SoPSCo=SoPSCo+" + DoiDau(IIf(dkc - dkn > 0, dkc - dkn, 0)) _
            + ",SoPS2No=SoPS2No+" + DoiDau(IIf(dkn - dkc > 0, dknt, 0)) + ",SoPS2Co=SoPS2Co+" + DoiDau(IIf(dkc - dkn > 0, dknt, 0)) + " WHERE ShTK='" + kh.sohieu + "' AND MaCT=-1"
        If DBKetoan.RecordsAffected = 0 Then
            ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,ShTK,SoPSNo,SoPSCo,DienGiai,SoPS2No,SoPS2Co) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",-1,'" + kh.sohieu + "'," _
                + DoiDau(IIf(dkn - dkc > 0, dkn - dkc, 0)) + "," + DoiDau(IIf(dkc - dkn > 0, dkc - dkn, 0)) + ",'" + kh.Ten + "'," + DoiDau(IIf(dkn - dkc > 0, dknt, 0)) + "," + DoiDau(IIf(dkc - dkn > 0, dknt, 0)) + ")"
        End If
        sql = "SELECT DISTINCTROW ChungTu.MaCT,ChungTu.SoHieu,ChungTu.NgayCT,ChungTu.DienGiai" + IIf(nn > 0, "E", "") + " AS DG,ChungTu.SoPS,IIF(" + CStr(kh.MaNT) + ">0,SoPS2No,0) AS SoPS2N,IIF(" + CStr(kh.MaNT) + ">0,SoPS2Co,0) AS SoPS2C,HethongTK.SoHieu As SHNo, HethongTK_1.SoHieu As SHCo,MaKH" _
            & " FROM (HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) LEFT JOIN HethongTK AS HethongTK_1 ON ChungTu.MaTKCo=HethongTK_1.MaSo" _
            & " Where (MaTKNo=" + CStr(rs!MaTaiKhoan) + ") AND (MaKH=" + CStr(rs!MaKhachHang) + ") AND " + WThang("ThangCT", tdau, tcuoi) + " AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) ORDER BY ChungTu.NgayCT, ChungTu.MaCT"
        Set rs_tk = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
        Do While Not rs_tk.EOF
            'If Left(rs_tk!shno, Len(taikhoan.SoHieu)) = taikhoan.SoHieu And rs!MaKhachHang = rs_tk!MaKH Then
                psn = psn + rs_tk!sops
                psn2 = psn2 + rs_tk!SoPS2N
                If IsNull(rs_tk!shco) Then s = "..." Else s = rs_tk!shco
                ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,ShTK,SoPSNo,Ngay,SoHieu,DienGiai,ShDu,SoPS2No) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + "," + CStr(rs_tk!MaCT) + ",'" + kh.sohieu + "'," + DoiDau(rs_tk!sops) _
                    + ",#" + Format(rs_tk!NgayCT, Mask_DB) + "#,'" + rs_tk!sohieu + "','" + rs_tk!dg + "','" + s + "'," + DoiDau(rs_tk!SoPS2N) + ")"
            'End If
            rs_tk.MoveNext
        Loop
        
        sql = "SELECT DISTINCTROW ChungTu.MaCT,ChungTu.SoHieu,ChungTu.NgayCT,ChungTu.DienGiai" + IIf(nn > 0, "E", "") + " AS DG,ChungTu.SoPS,IIF(" + CStr(kh.MaNT) + ">0,SoPS2No,0) AS SoPS2N,IIF(" + CStr(kh.MaNT) + ">0,SoPS2Co,0) AS SoPS2C,HethongTK.SoHieu As SHNo, HethongTK_1.SoHieu As SHCo,MaKHC" _
            & " FROM (HeThongTK AS HeThongTK_1 INNER JOIN ChungTu ON HeThongTK_1.MaSo = ChungTu.MaTKCo) LEFT JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo" _
            & " Where (MaTKCo=" + CStr(rs!MaTaiKhoan) + ") AND (MaKHC=" + CStr(rs!MaKhachHang) + ") AND " + WThang("ThangCT", tdau, tcuoi) + " AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) ORDER BY ChungTu.NgayCT, ChungTu.MaCT"
        Set rs_tk = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
        Do While Not rs_tk.EOF
            'If Left(rs_tk!shco, Len(taikhoan.SoHieu)) = taikhoan.SoHieu And rs!MaKhachHang = rs_tk!MaKHC Then
                psc = psc + rs_tk!sops
                psc2 = psc2 + rs_tk!SoPS2C
                If IsNull(rs_tk!shno) Then s = "..." Else s = rs_tk!shno
                ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,ShTK,SoPSCo,Ngay,SoHieu,DienGiai,ShDu,SoPS2Co) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + "," + CStr(rs_tk!MaCT) + ",'" + kh.sohieu + "'," + DoiDau(rs_tk!sops) _
                    + ",#" + Format(rs_tk!NgayCT, Mask_DB) + "#,'" + rs_tk!sohieu + "','" + rs_tk!dg + "','" + s + "'," + DoiDau(rs_tk!SoPS2C) + ")"
            'End If
            rs_tk.MoveNext
        Loop
        
        ckn = dkn - dkc + psn - psc
        If ckn > 0 Then
            ckc = 0
        Else
            ckc = -ckn
            ckn = 0
        End If
        cknt = IIf(ckc - ckn < 0, dknt + psn2 - psc2, dknt - psn2 + psc2)
        
        ExecuteSQL5 "UPDATE BKNhomPS SET SoPSNo=SoPSNo+" + DoiDau(IIf(ckn - ckc > 0, ckn - ckc, 0)) + ",SoPSCo=SoPSCo+" + DoiDau(IIf(ckc - ckn > 0, ckc - ckn, 0)) _
            + ",SoPS2No=SoPS2No+" + DoiDau(IIf(ckn - ckc > 0, cknt, 0)) + ",SoPS2Co=SoPS2Co+" + DoiDau(IIf(ckc - ckn > 0, cknt, 0)) + " WHERE ShTK='" + kh.sohieu + "' AND MaCT=0"
        If DBKetoan.RecordsAffected = 0 Then
            ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,ShTK,SoPSNo,SoPSCo,DienGiai,SoPS2No,SoPS2Co) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",0,'" + kh.sohieu + "'," _
                + DoiDau(IIf(ckn - ckc > 0, ckn - ckc, 0)) + "," + DoiDau(IIf(ckc - ckn > 0, ckc - ckn, 0)) + ",'" + kh.Ten + "'," + DoiDau(IIf(ckn - ckc > 0, cknt, 0)) + "," + DoiDau(IIf(ckc - ckn > 0, cknt, 0)) + ")"
        End If
        If dkn = 0 And dkc = 0 And psn = 0 And psc = 0 And psn2 = 0 And psc2 = 0 Then ExecuteSQL5 "DELETE * FROM BKNhomPS WHERE ShTK='" + kh.sohieu + "'"
KhongIn:
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set kh = Nothing
    
    frmMain.Rpt.ReportFileName = "THPSCT.RPT"

    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(5) = "TK='" + taikhoan.sohieu + " - " + IIf(nn > 0, taikhoan.TenE, taikhoan.Ten) + "'"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
End Sub

Private Function THPSNhomCN2(ms As Long, tdau As Integer, tcuoi As Integer) As Boolean
    Dim rs As Recordset, taikhoan As New ClsTaikhoan, psn2 As Double, psc2 As Double, kh As New ClsKhachHang
    Dim dkn As Double, dkc As Double, dknt As Double, psn As Double, psc As Double, ckn As Double, ckc As Double, cknt As Double
    Dim rs_tk As Recordset, sql As String, inso As Boolean, i As Integer, st As String, j As Integer, s As String
    
    THPSNhomCN2 = False
    kh.InitKhachHangMaSo ms
    ExecuteSQL5 "DELETE * FROM BKNhomPS"
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW MaTaiKhoan FROM SoDuKhachHang WHERE MaKhachHang=" + CStr(ms) + " GROUP BY MaTaiKhoan", dbOpenSnapshot)
    
    Do While Not rs.EOF
        taikhoan.InitTaikhoanMaSo rs!MaTaiKhoan
        kh.SoDuKH ThangTruoc(tdau), dkn, dkc, dknt, rs!MaTaiKhoan

        psn = 0
        psc = 0
        psn2 = 0
        psc2 = 0
        
        ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,ShTK,SoPSNo,SoPSCo,DienGiai,SoPS2No,SoPS2Co) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",-1,'" + taikhoan.sohieu + "'," _
            + DoiDau(IIf(dkn - dkc > 0, dkn - dkc, 0)) + "," + DoiDau(IIf(dkc - dkn > 0, dkc - dkn, 0)) + ",'" + IIf(nn > 0, taikhoan.TenE, taikhoan.Ten) + "'," + DoiDau(IIf(dkn - dkc > 0, dknt, 0)) + "," + DoiDau(IIf(dkc - dkn > 0, dknt, 0)) + ")"
        
        sql = "SELECT DISTINCTROW ChungTu.MaCT,ChungTu.SoHieu,ChungTu.NgayCT,ChungTu.DienGiai" + IIf(nn > 0, "E", "") + " AS DG,ChungTu.SoPS,IIF(" + CStr(kh.MaNT) + ">0,SoPS2No,0) AS SoPS2N,IIF(" + CStr(kh.MaNT) + ">0,SoPS2Co,0) AS SoPS2C,HethongTK.SoHieu As SHNo, HethongTK_1.SoHieu As SHCo,MaKH" _
            & " FROM HeThongTK AS HeThongTK_1 RIGHT JOIN (HeThongTK INNER JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) ON HeThongTK_1.MaSo = ChungTu.MaTKCo" _
            & " Where (MaTKNo=" + CStr(rs!MaTaiKhoan) + ") AND (MaKH=" + CStr(kh.MaSo) + ") AND " + WThang("ThangCT", tdau, tcuoi) + " AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) ORDER BY ChungTu.NgayCT, ChungTu.MaCT"
        Set rs_tk = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
        Do While Not rs_tk.EOF
                psn = psn + rs_tk!sops
                psn2 = psn2 + rs_tk!SoPS2N
                If IsNull(rs_tk!shco) Then s = "..." Else s = rs_tk!shco
                ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,ShTK,SoPSNo,Ngay,SoHieu,DienGiai,ShDu,SoPS2No) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + "," + CStr(rs_tk!MaCT) + ",'" + taikhoan.sohieu + "'," + DoiDau(rs_tk!sops) _
                    + ",#" + Format(rs_tk!NgayCT, Mask_DB) + "#,'" + rs_tk!sohieu + "','" + rs_tk!dg + "','" + s + "'," + DoiDau(rs_tk!SoPS2N) + ")"
            rs_tk.MoveNext
        Loop
        sql = "SELECT DISTINCTROW ChungTu.MaCT,ChungTu.SoHieu,ChungTu.NgayCT,ChungTu.DienGiai" + IIf(nn > 0, "E", "") + " AS DG,ChungTu.SoPS,IIF(" + CStr(kh.MaNT) + ">0,SoPS2No,0) AS SoPS2N,IIF(" + CStr(kh.MaNT) + ">0,SoPS2Co,0) AS SoPS2C,HethongTK.SoHieu As SHNo, HethongTK_1.SoHieu As SHCo,MaKH,MaKHC" _
            & " FROM HeThongTK AS HeThongTK_1 INNER JOIN (HeThongTK RIGHT JOIN ChungTu ON HeThongTK.MaSo = ChungTu.MaTKNo) ON HeThongTK_1.MaSo = ChungTu.MaTKCo" _
            & " Where (MaTKCo=" + CStr(rs!MaTaiKhoan) + ") AND (MaKHC=" + CStr(kh.MaSo) + ") AND " + WThang("ThangCT", tdau, tcuoi) + " AND (Chungtu.MaLoai<>4 OR (Chungtu.MaLoai=4 AND Chungtu.MaTKNo<>Chungtu.MaTkco)) ORDER BY ChungTu.NgayCT, ChungTu.MaCT"
        Set rs_tk = DBKetoan.OpenRecordset(sql, dbOpenSnapshot, dbForwardOnly)
        Do While Not rs_tk.EOF
                psc = psc + rs_tk!sops
                psc2 = psc2 + rs_tk!SoPS2C
                If IsNull(rs_tk!shno) Then s = "..." Else s = rs_tk!shno
                ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,ShTK,SoPSCo,Ngay,SoHieu,DienGiai,ShDu,SoPS2Co) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + "," + CStr(rs_tk!MaCT) + ",'" + taikhoan.sohieu + "'," + DoiDau(rs_tk!sops) _
                    + ",#" + Format(rs_tk!NgayCT, Mask_DB) + "#,'" + rs_tk!sohieu + "','" + rs_tk!dg + "','" + s + "'," + DoiDau(rs_tk!SoPS2C) + ")"
            rs_tk.MoveNext
        Loop
        
        ckn = dkn - dkc + psn - psc
        If ckn > 0 Then
            ckc = 0
        Else
            ckc = -ckn
            ckn = 0
        End If
        cknt = IIf(ckc - ckn < 0, dknt + psn2 - psc2, dknt - psn2 + psc2)

        ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,ShTK,SoPSNo,SoPSCo,DienGiai,SoPS2No,SoPS2Co) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",0,'" + taikhoan.sohieu + "'," _
            + DoiDau(IIf(ckn - ckc > 0, ckn - ckc, 0)) + "," + DoiDau(IIf(ckc - ckn > 0, ckc - ckn, 0)) + ",'" + IIf(nn > 0, taikhoan.TenE, taikhoan.Ten) + "'," + DoiDau(IIf(ckn - ckc > 0, cknt, 0)) + "," + DoiDau(IIf(ckc - ckn > 0, cknt, 0)) + ")"

        If dkn = 0 And dkc = 0 And psn = 0 And psc = 0 And psn2 = 0 And psc2 = 0 Then ExecuteSQL5 "DELETE * FROM BKNhomPS WHERE ShTK='" + kh.sohieu + "'"
KhongIn:
        rs.MoveNext
    Loop
    THPSNhomCN2 = (rs.recordCount > 0)
    rs.Close
    Set rs = Nothing
    frmMain.Rpt.ReportFileName = "THPSCT.RPT"

    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(5) = "TK='" + kh.sohieu + " - " + kh.Ten + "'"
    frmMain.Rpt.Formulas(6) = "TenBaoCao='Sæ chi tiÕt ®èi t­îng thanh to¸n'"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    
    Set taikhoan = Nothing
    Set kh = Nothing
End Function

Private Sub ChiPhiCT(mct As Long, tdau As Integer, tcuoi As Integer)
    Dim sql As String, i As Integer, s621 As Double, s622 As Double, s623 As Double, s627 As Double
    
    s621 = SoCPPB(tdau, tcuoi, "CPNVLPB", mct)
    s622 = SoCPPB(tdau, tcuoi, "CPNCPB", mct)
    s623 = SoCPPB(tdau, tcuoi, "CPMPB", mct)
    s627 = SoCPPB(tdau, tcuoi, "CPSXC_", mct)
    If tp.Kiem_tra_qd15() Then
    
    sql = "SELECT ChungTu.SoHieu AS SHCT,NgayCT,DienGiai" + IIf(nn > 0, "E", "") + ",Sum(SoPS) AS PS,LEFT(IIF(TK.TK_ID=" + CStr(TKDT_ID) + ",TK.SoHieu,HethongTK.SoHieu),3) AS SHTK,IIF(TK.TK_ID=" + CStr(TKDT_ID) + ",HethongTK.SoHieu,TK.SoHieu) AS SHCo FROM (" + ChungTu2TKNC(10) + ") INNER JOIN TP154 ON ChungTu.MaTP=TP154.MaSo " _
        & " WHERE (HethongTK.SoHieu LIKE '62*' OR HethongTK.SoHieu LIKE '64*' OR TK.SoHieu LIKE '51*') AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mct > 0, " AND MaTP=" + CStr(mct), "") _
        + " GROUP BY NgayCT,ChungTu.SoHieu,NgayCT,DienGiai" + IIf(nn > 0, "E", "") + ",LEFT(IIF(TK.TK_ID=" + CStr(TKDT_ID) + ",TK.SoHieu,HethongTK.SoHieu),3),IIF(TK.TK_ID=" + CStr(TKDT_ID) + ",HethongTK.SoHieu,TK.SoHieu)" _
        + " UNION SELECT ChungTu.SoHieu AS SHCT,NgayCT,DienGiai" + IIf(nn > 0, "E", "") + ",Sum(SoPS) AS PS,LEFT(HethongTK.SoHieu,3) AS SHTK,'...' AS SHCo FROM (" + ChungTu2TKNC(-1) + ") INNER JOIN TP154 ON ChungTu.MaTP=TP154.MaSo " _
        & " WHERE MaTKCo=0 AND (HethongTK.SoHieu LIKE '62*' OR HethongTK.SoHieu LIKE '64*') AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mct > 0, " AND MaTP=" + CStr(mct), "") _
        + " GROUP BY NgayCT, ChungTu.SoHieu ,DienGiai" + IIf(nn > 0, "E", "") + ",LEFT(HethongTK.SoHieu,3)" _
        & " UNION SELECT ChungTu.SoHieu AS SHCT,NgayCT,DienGiai" + IIf(nn > 0, "E", "") + ",Sum(IIF(LEFT(HethongTK.SoHieu,3)<>'154',-SoPS,SoPS)) AS PS,LEFT(TK.SoHieu,3) AS SHTK,HethongTK.SoHieu AS SHCo FROM (" + ChungTu2TKNC(10) + ") INNER JOIN TP154 ON ChungTu.MaTP=TP154.MaSo " _
        & " WHERE HethongTK.Loai=1 AND (TK.SoHieu LIKE '62*' OR TK.SoHieu LIKE '64*') AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mct > 0, " AND MaTP=" + CStr(mct), "") _
        + " GROUP BY ChungTu.SoHieu,NgayCT,DienGiai" + IIf(nn > 0, "E", "") + ",LEFT(TK.SoHieu,3),HethongTK.SoHieu ORDER BY NgayCT"
    If s621 <> 0 Then sql = sql + " UNION SELECT 'xxx' AS SHCT, #" + Format(NgayCuoiThang(pNamTC, tcuoi), Mask_DB) + "# AS NgayCT,'" + ABCtoVNI("Ph©n bæ chi phÝ nguyªn vËt liÖu") + "'," + DoiDau(s621) + ",'621','...' FROM License"
    If s622 <> 0 Then sql = sql + " UNION SELECT 'xxx' AS SHCT, #" + Format(NgayCuoiThang(pNamTC, tcuoi), Mask_DB) + "# AS NgayCT,'" + ABCtoVNI("Ph©n bæ chi phÝ nh©n c«ng") + "'," + DoiDau(s622) + ",'622','...' FROM License"
    If s623 <> 0 Then sql = sql + " UNION SELECT 'xxx' AS SHCT, #" + Format(NgayCuoiThang(pNamTC, tcuoi), Mask_DB) + "# AS NgayCT,'" + ABCtoVNI("Ph©n bæ chi phÝ m¸y thi c«ng") + "'," + DoiDau(s623) + ",'623','...' FROM License"
    If s627 <> 0 Then sql = sql + " UNION SELECT 'xxx' AS SHCT, #" + Format(NgayCuoiThang(pNamTC, tcuoi), Mask_DB) + "# AS NgayCT,'" + ABCtoVNI("Ph©n bæ chi phÝ s¶n xuÊt chung") + "'," + DoiDau(s627) + ",'627','...' FROM License"
    Else
    
    
     sql = "SELECT ChungTu.SoHieu AS SHCT,NgayCT,DienGiai" + IIf(nn > 0, "E", "") + ",Sum(SoPS) AS PS,LEFT(IIF(TK.TK_ID=" + CStr(TKDT_ID) + ",TK.SoHieu,HethongTK.SoHieu),3) AS c,IIF(TK.TK_ID=" + CStr(TKDT_ID) + ",HethongTK.SoHieu,TK.SoHieu) AS SHCo FROM (" + ChungTu2TKNC(10) + ") INNER JOIN TP154 ON ChungTu.MaTP=TP154.MaSo " _
        & " WHERE (HethongTK.SoHieu LIKE '154*' OR TK.SoHieu LIKE '51*') AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mct > 0, " AND MaTP=" + CStr(mct), "") _
        + " GROUP BY NgayCT,ChungTu.SoHieu,NgayCT,DienGiai" + IIf(nn > 0, "E", "") + ",LEFT(IIF(TK.TK_ID=" + CStr(TKDT_ID) + ",TK.SoHieu,HethongTK.SoHieu),3),IIF(TK.TK_ID=" + CStr(TKDT_ID) + ",HethongTK.SoHieu,TK.SoHieu)" _
        + " UNION SELECT ChungTu.SoHieu AS SHCT,NgayCT,DienGiai" + IIf(nn > 0, "E", "") + ",Sum(SoPS) AS PS,LEFT(HethongTK.SoHieu,3) AS SHTK,'...' AS SHCo FROM (" + ChungTu2TKNC(-1) + ") INNER JOIN TP154 ON ChungTu.MaTP=TP154.MaSo " _
        & " WHERE MaTKCo=0 AND (HethongTK.SoHieu LIKE '154*') AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mct > 0, " AND MaTP=" + CStr(mct), "") _
        + " GROUP BY NgayCT, ChungTu.SoHieu ,DienGiai" + IIf(nn > 0, "E", "") + ",LEFT(HethongTK.SoHieu,3)" _
        & " UNION SELECT ChungTu.SoHieu AS SHCT,NgayCT,DienGiai" + IIf(nn > 0, "E", "") + ",Sum(IIF(LEFT(HethongTK.SoHieu,3)<>'154*',-SoPS,SoPS)) AS PS,LEFT(TK.SoHieu,3) AS SHTK,HethongTK.SoHieu AS SHCo FROM (" + ChungTu2TKNC(10) + ") INNER JOIN TP154 ON ChungTu.MaTP=TP154.MaSo " _
        & " WHERE HethongTK.Loai=1 AND (TK.SoHieu LIKE '154*') AND " + WThang("ThangCT", tdau, tcuoi) + IIf(mct > 0, " AND MaTP=" + CStr(mct), "") _
        + " GROUP BY ChungTu.SoHieu,NgayCT,DienGiai" + IIf(nn > 0, "E", "") + ",LEFT(TK.SoHieu,3),HethongTK.SoHieu ORDER BY NgayCT"

    If s621 <> 0 Then sql = sql + " UNION SELECT 'xxx' AS SHCT, #" + Format(NgayCuoiThang(pNamTC, tcuoi), Mask_DB) + "# AS NgayCT,'" + ABCtoVNI("Ph©n bæ chi phÝ nguyªn vËt liÖu") + "'," + DoiDau(s621) + ",'621','...' FROM License"
    If s622 <> 0 Then sql = sql + " UNION SELECT 'xxx' AS SHCT, #" + Format(NgayCuoiThang(pNamTC, tcuoi), Mask_DB) + "# AS NgayCT,'" + ABCtoVNI("Ph©n bæ chi phÝ nh©n c«ng") + "'," + DoiDau(s622) + ",'622','...' FROM License"
    If s623 <> 0 Then sql = sql + " UNION SELECT 'xxx' AS SHCT, #" + Format(NgayCuoiThang(pNamTC, tcuoi), Mask_DB) + "# AS NgayCT,'" + ABCtoVNI("Ph©n bæ chi phÝ m¸y thi c«ng") + "'," + DoiDau(s623) + ",'623','...' FROM License"
    If s627 <> 0 Then sql = sql + " UNION SELECT 'xxx' AS SHCT, #" + Format(NgayCuoiThang(pNamTC, tcuoi), Mask_DB) + "# AS NgayCT,'" + ABCtoVNI("Ph©n bæ chi phÝ s¶n xuÊt chung") + "'," + DoiDau(s627) + ",'627','...' FROM License"
  End If
    SetSQL "QChitiet", sql ' " select tong.SHCT,tong.ngayct,tong.diengiai,tong.ps as sops,tong.SHTK,tong.SHCo from (" + sql + ") tong  "
      If tp.Kiem_tra_qd15() Then
    frmMain.Rpt.ReportFileName = "CTCT.RPT"
    Else
    frmMain.Rpt.ReportFileName = "CTCT1.RPT"
    End If
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(5) = "TK='" + TxtCT(0).Text + " - " + LbCT(0).Caption + "'"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
End Sub

Private Sub InTHXuatKhoDT(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, mkho As Long, mtk As Long, mloai As Long, tkdu As Long, xT As Integer)
    Dim sql As String, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    
    sql = "SELECT DISTINCTROW Vattu.SoHieu,TenVattu,DonVi,Sum(SoPS2Co) AS Luong,Sum(SoPS) AS Tien FROM ChungTu INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo" _
        & " WHERE (ChungTu.MaLoai=2 OR (ChungTu.MaLoai=4 AND MaTKNo<>MaTKCo)) AND " + wsql + " AND MaTP =" + CStr(tkdu) _
        + IIf(mkho > 0, " AND MaKho=" + CStr(mkho), "") + IIf(mtk <> 0, " AND MaTKCo = " + CStr(mtk), "") + IIf(mloai <> 0, " AND MaPhanLoai = " + CStr(mloai), "") + IIf(xT > 0, " AND CT_ID>900000000", "") _
        + " GROUP BY Vattu.SoHieu,TenVattu,DonVi"
    sql = sql + " UNION SELECT DISTINCTROW Vattu.SoHieu,TenVattu,DonVi,-Sum(SoPS2No) AS Luong,-Sum(SoPS) AS Tien FROM ChungTu INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo" _
        & " WHERE (ChungTu.MaLoai=1 OR (ChungTu.MaLoai=4 AND MaTKNo<>MaTKCo)) AND " + wsql + " AND MaTP =" + CStr(tkdu) _
        + IIf(mkho > 0, " AND ((MaKho=" + CStr(mkho) + " AND MaLoai=1) OR (MaNguon=" + CStr(mkho) + " AND MaLoai=4))", "") + IIf(mtk <> 0, " AND MaTKNo = " + CStr(mtk), "") + IIf(mloai <> 0, " AND MaPhanLoai = " + CStr(mloai), "") + IIf(xT > 0, " AND CT_ID>900000000", "") _
        + " GROUP BY Vattu.SoHieu,TenVattu,DonVi"
        
    SetSQL "QLuyKe", sql
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    frmMain.Rpt.Formulas(4) = "TKDU = IF PageNumber() = 1 THEN '" + ABCtoVNI("XuÊt cho ®èi t­îng: ") + TxtCT(1).Text + " - " + LbCT(1).Caption + "'"
    If xT > 0 Then frmMain.Rpt.Formulas(5) = "TenBaoCao = IF PageNumber()=1 THEN 'B¶ng kª nhËp xuÊt th¼ng'"
    frmMain.Rpt.ReportFileName = "LUYKE2.RPT"
    RptSetDate ncuoi, nn
End Sub

Private Sub InTHDoanhThuVAT(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date)
    Dim wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    SetSQL "Mientru", "SELECT MaCT,KCT,TyLe FROM ChungTu INNER JOIN HoaDon ON ChungTu.MaSo=HoaDon.MaSo WHERE HoaDon.Loai=1 AND MaLoai=8 AND " + wsql
    
    SetSQL "QChitiet", "SELECT DISTINCTROW Vattu.SoHieu,TenVattu, DonVi,SUM(IIF(TyLe=0 AND KCT=1,SoPS,0)) AS DKT, SUM(IIF(TyLe=0 AND KCT=0,SoPS,0)) AS D0, SUM(IIF(TyLe=5,SoPS,0)) AS D5, SUM(IIF(TyLe=10,SoPS,0)) AS D10, SUM(IIF(TyLe=5,TyLe*SoPS/100,0)) AS T5, SUM(IIF(TyLe=10,TyLe*SoPS/100,0)) AS T10" _
        & " FROM ((ChungTu LEFT JOIN Vattu ON ChungTu.MaVattu = Vattu.MaSo) INNER JOIN MienTru ON ChungTu.MaCT=MienTru.MaCT) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo" _
        & " WHERE TK_ID=" + CStr(TKDT_ID) + " OR HethongTK.Loai=7 GROUP BY Vattu.SoHieu,TenVattu,DonVi"
    
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    frmMain.Rpt.ReportFileName = "THDT4.RPT"
    RptSetDate ncuoi, nn
End Sub

Private Sub KiemTraUser()
    SSTab.TabEnabled(0) = (frmMain.tag Mod 10 >= 1)
    If Not ((frmMain.tag Mod 10 >= 1) Or (User_Right = 0)) Then
        SSTab.TabEnabled(1) = (frmMain.tag Mod 100 >= 10 Or frmMain.tag Mod 1000 >= 100)
        SSTab.TabEnabled(2) = (frmMain.tag Mod 100000 >= 10000)
        SSTab.TabEnabled(3) = (frmMain.tag Mod 10000 >= 1000)
    End If
End Sub

Private Sub SoTSCD(thang As Integer)
    SetSQL "MienTru", "SELECT * FROM CTTaiSan WHERE " + WThang("Thang", 0, thang) + " AND MaLoai=33"
    SetSQL "SoDuDauCuoiKy", "SELECT DISTINCTROW TaiSan.Ten AS TenTS, TaiSan.SoHieu AS SoHieuTS,NangLuc,TaiSan.GhiChu,ThangTang,NamSX,NamSD, LoaiTaiSan.SoHieu AS SoHieuTK, LoaiTaiSan.MaSo AS MaTK, LoaiTaiSan.Ten AS TenTK,QuocGia.Ten AS NuocSX," _
          & "(ThongSo.NG_NS+ThongSo.NG_TBS+ThongSo.NG_CNK+ThongSo.NG_TD) AS NG, (ThongSo.KH_NS+ThongSo.KH_TBS+ThongSo.KH_CNK+ThongSo.KH_TD) AS KH,LoaiTaiSan_1.SoHieu,LoaiTaiSan_1.Ten, " _
          & "TaiSan.SHCT AS SCT1,TaiSan.NCT AS NCT1,TaiSan.NCT AS NGS1, MienTru.SoHieu AS SCT2,MienTru.VaoSo AS NCT2,MienTru.DienGiai AS DG2, (MienTru.NG_NS+MienTru.NG_TBS+MienTru.NG_TD+MienTru.NG_CNK-MienTru.CL_NS-MienTru.CL_TBS-MienTru.CL_TD-MienTru.CL_CNK) AS SoDaKH,TaiSan.NamKH " _
          & "FROM ((LoaiTaiSan AS LoaiTaiSan_1 INNER JOIN ((LoaiTaiSan INNER JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) INNER JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON LoaiTaiSan_1.MaSo=TaiSan.MaLoai) INNER JOIN QuocGia ON TaiSan.MaNuoc=QuocGia.MaSo) LEFT JOIN MienTru ON TaiSan.MaSo=MienTru.MaTS " _
          & "WHERE " + WThang("ThangTang", 0, thang) + " AND ((ThongSo.Thang = " + CStr(CThangDB(thang)) + " AND " + WThang2("ThangGiam", thang, 0) + ") OR (ThongSo.Thang = ThangGiam-1 AND " + WThang("ThangGiam", 0, thang) + ")) ORDER BY TaiSan.SoHieu"
End Sub

Private Sub Form_DblClick()
    FThuChi.tag = 10
    FThuChi.GetPhieu NLB, KTT, GD, 0
    SaveSetting IniPath, "Environment", "SReport1", NLB
    SaveSetting IniPath, "Environment", "SReport2", KTT
    SaveSetting IniPath, "Environment", "SReport3", GD
End Sub

Private Sub InCongNoQuaHanMuc(thang As Integer, shtk As String, mpl As Long)
    Dim st As String
    
    st = CThangDB(thang)
    SetSQL "QChitiet", "SELECT DISTINCTROW KhachHang.SoHieu, KhachHang.Ten, PhanLoaiKhachHang.SoHieu AS SHPL, TenPhanLoai,KyHieu," _
        & " Sum(SoDuKhachHang.DuNo_" + st + ") AS CkNo, Sum(SoDuKhachHang.DuCo_" + st + ") AS CkCo, Sum(SoDuKhachHang.DuNT_" + st + ") AS CkNT,PhanLoaiKhachHang.PLCha, First(HethongTK.Kieu) AS LoaiNo, First(KhachHang.DuMax) AS HM" _
        & " FROM (((KhachHang INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo) INNER JOIN SoDuKhachHang ON KhachHang.MaSo=SoDuKhachHang.MaKhachHang) INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo) LEFT JOIN NguyenTe ON KhachHang.MaNT=NguyenTe.MaSo " _
        & " WHERE  KhachHang.DuMax>0 AND LEFT(KhachHang.SoHieu,1)<>'#' " + IIf(Len(shtk) > 0, " AND HethongTK.SoHieu LIKE '" + shtk + "*'", "") + IIf(mpl > 0, " AND MaPhanLoai=" + CStr(mpl), "") _
        + " AND ((HethongTK.Kieu<0 AND (SoDuKhachHang.DuNo_" + st + "-SoDuKhachHang.DuCo_" + st + ")>KhachHang.DuMax) OR (HethongTK.Kieu>0 AND (SoDuKhachHang.DuCo_" + st + "-SoDuKhachHang.DuNo_" + st + ")>KhachHang.DuMax))" _
        + " GROUP BY HethongTK.MaSo,PhanLoaiKhachHang.Sohieu,KhachHang.SoHieu,KhachHang.Ten,TenPhanLoai,KyHieu,PLCha"
    SetSQL "QSoCai", "SELECT QChitiet.*, PhanLoaiKhachHang.SoHieu AS SH2,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai2,PhanLoaiKhachHang.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoaiKhachHang ON QChitiet.PLCha=PhanLoaiKhachHang.MaSo"
    SetSQL "QTongHopCT", "SELECT QSoCai.*, PhanLoaiKhachHang.SoHieu AS SH1,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai1 FROM QSoCai LEFT JOIN PhanLoaiKhachHang ON QSoCai.PLCha2=PhanLoaiKhachHang.MaSo"
    
    GauGe.Value = 2
    frmMain.Rpt.ReportFileName = "SDCNQH.RPT"
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(thang, thang, nn) + "'"
    RptSetDate NgayCuoiThang(pNamTC, thang), nn
    If Len(shtk) > 0 Then
        frmMain.Rpt.Formulas(5) = "TenBaoCao = IF PageNumber() = 1 THEN '" + ABCtoVNI("Tµi kho¶n: ") + shtk + " - " + LbTenTk(7).Caption + "'"
    End If
End Sub

Private Sub InCTDoanhThu4(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, mkh As Long, mvt As Long, mkho As Long)
    Dim sql As String, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,MaCha,CoCon) SELECT ChungTu.MaSo,Cstr(ChungTu.MaSo),MaCT,IIF(TK_ID=" + CStr(TKDT_ID) + ",0,1) FROM " + ChungTu2TKNC(1) + " WHERE (TK_ID=" + CStr(TKDT_ID) + " OR TK_ID=" + CStr(GTGTPN_ID) + ") AND " + wsql + IIf(mvt > 0, " AND MaVattu=" + CStr(mvt), "") + IIf(mkho > 0, " AND MaKho=" + CStr(mkho), "")
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN ChungTu ON BaoCaoCP.MaCha=ChungTu.MaCT SET BC_ID=MaKH WHERE MaKH>0"
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN (" + ChungTu2TKHD(0) + ") ON BaoCaoCP.MaCha=ChungTu.MaCT SET BC_ID=MaKhachHang WHERE BC_ID=0"
    
    sql = "SELECT DISTINCTROW ChungTu.MaSo,ChungTu.SoHieu, ChungTu.NgayCT, ChungTu.DienGiai, ChungTu.SoPS, ChungTu.SoPS2Co, HeThongTK.SoHieu, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi,ChungTu.MaKH, KhachHang.SoHieu,KhachHang.Ten,BaoCaoCP.CoCon, ChungTu.MaCT" _
        & " FROM (((ChungTu LEFT JOIN HethongTK ON ChungTu.MaTKNo = HethongTK.MaSo) LEFT JOIN Vattu ON ChungTu.MaVattu = Vattu.MaSo) INNER JOIN BaoCaoCP ON ChungTu.MaSo=BaoCaoCP.MaSo) LEFT JOIN KhachHang ON BaoCaoCP.BC_ID=KhachHang.MaSo" _
        + IIf(mkh > 0, " WHERE BC_ID=" + CStr(mkh), "") + " ORDER BY ChungTu.NgayGS, ChungTu.MaCT"
    
    SetSQL "QNhatKy", sql
    SetSQL "MienTru", "SELECT ChungTu.MaCT, HoaDon.SoHD FROM " + ChungTu2TKHD(0) + " WHERE " + wsql + " GROUP BY ChungTu.MaCT, HoaDon.SoHD"
    SetSQL "QChitiet", "SELECT QNhatKy.*,MienTru.SoHD FROM QNhatKy LEFT JOIN MienTru ON QNhatKy.MaCT=MienTru.MaCT ORDER BY NgayCT, QNhatKy.ChungTu.SoHieu"
    
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    frmMain.Rpt.ReportFileName = "CTDT4.RPT"
    RptSetDate ncuoi
End Sub

Private Sub BKDinhMuc(thang As Integer)
    Dim rs As Recordset, thangdm As Integer
    
    ExecuteSQL5 "UPDATE DinhMuc SET ThangHL=IIF(Thang=" + CStr(thang) + ",Thang,0)"
    Set rs = DBKetoan.OpenRecordset("SELECT MaTP FROM DinhMuc WHERE ThangHL=0 GROUP BY MaTP", dbOpenSnapshot, dbForwardOnly)
    Do While Not rs.EOF
        thangdm = SelectSQL("SELECT  TOP 1 Thang AS F1 FROM DinhMuc WHERE MaNVL>0 AND MaTP=" + CStr(rs!MaTP) + " AND " + WThang("Thang", 0, thang) + " ORDER BY " + SetMonthOrder("Thang") + " DESC")
        ExecuteSQL5 "UPDATE DinhMuc SET ThangHL=" + CStr(thang) + " WHERE Thang=" + CStr(thangdm) + " AND MaTP=" + CStr(rs!MaTP)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    SetSQL "QLuyKe", "SELECT Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, VT.SoHieu AS SNVL, VT.TenVattu AS TNVL, VT.DonVi AS DNVL, DinhMuc.SoLuong AS LuongNVL" _
        & " FROM (DinhMuc INNER JOIN Vattu ON DinhMuc.MaTP=Vattu.MaSo) INNER JOIN Vattu AS VT ON DinhMuc.MaNVL=VT.MaSo " _
        & " WHERE ThangHL=" + CStr(thang) + " ORDER BY Vattu.SoHieu, VT.SoHieu"
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(thang, thang, nn) + "'"
    frmMain.Rpt.ReportFileName = "DM.RPT"
    RptSetDate NgayCuoiThang(pNamTC, thang), nn
End Sub

Private Sub BangKeTichSo(taikhoan As ClsTaikhoan, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, lsv As Double, lsg As Double, mnt As Long)
    Dim rs As Recordset, n As Date, sql As String, ms As Long, sdn As Double, sdc As Double, nt As Double, n1 As Date
    Dim TK As New ClsTaikhoan, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    taikhoan.DanXuat TK
    If mnt <> 0 Then TK.LayMaNT mnt
    If mnt = 0 Then
        sql = "SELECT -1 AS LoaiPS,NgayCT,Sum(SoPS) AS PS FROM " + ChungTu2TKNC(-1) + " WHERE " + wsql + " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*' GROUP BY NgayCT" _
            & " UNION SELECT 1 AS LoaiPS,NgayCT,Sum(SoPS) AS PS FROM " + ChungTu2TKNC(1) + " WHERE " + wsql + " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*' GROUP BY NgayCT"
    Else
        sql = "SELECT -1 AS LoaiPS,NgayCT,Sum(IIF(" + CStr(mnt) + ">0,SoPS2No,SoPS)) AS PS FROM " + ChungTu2TKNC(-1) + " WHERE " + wsql + " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*' GROUP BY NgayCT" _
            & " UNION SELECT 1 AS LoaiPS,NgayCT,Sum(IIF(" + CStr(mnt) + ">0,SoPS2Co,SoPS)) AS PS FROM " + ChungTu2TKNC(1) + " WHERE " + wsql + " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*' GROUP BY NgayCT"
    End If
    SetSQL "MienTru", sql
    Set rs = DBKetoan.OpenRecordset("SELECT * FROM MienTru ORDER BY NgayCT", dbOpenSnapshot, dbForwardOnly)
    If tcuoi > 0 Then
        TK.SoDuTK ThangTruoc(tdau), sdn, sdc, nt
    Else
        TK.SoDuNgay ndau - 1, sdn, sdc, nt
    End If
    If mnt > 0 Then
        If sdn > 0 Then sdn = nt
        If sdc > 0 Then sdc = nt
    End If
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    If tcuoi > 0 Then n = NgayDauThang(pNamTC, tdau) Else n = ndau
    ms = 1
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Kq3,Kq4,Ngay1) VALUES (1,'1'," + DoiDau(sdn) + "," + DoiDau(sdc) + ",#" + Format(n, Mask_DB) + "#)"
    Do While Not rs.EOF
        If rs!loaips < 0 Then sdn = sdn + rs!ps Else sdc = sdc + rs!ps
        If sdn - sdc >= 0 Then
            sdn = sdn - sdc
            sdc = 0
        Else
            sdc = sdc - sdn
            sdn = 0
        End If
        If rs!NgayCT <> n Then
            ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=" + CStr(rs!NgayCT - n) + " WHERE MaSo=" + CStr(ms)
            n = rs!NgayCT
            ms = ms + 1
            ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Kq1,Kq2,Kq3,Kq4,Ngay1) VALUES (" + CStr(ms) + ",'" + CStr(ms) + "'," + IIf(rs!loaips < 0, DoiDau(rs!ps), "0") + "," + IIf(rs!loaips > 0, DoiDau(rs!ps), "0") + "," + DoiDau(sdn) + "," + DoiDau(sdc) + ",#" + Format(n, Mask_DB) + "#)"
        Else
            ExecuteSQL5 "UPDATE BaoCaoCP2 SET " + IIf(rs!loaips < 0, "Kq1=Kq1+" + DoiDau(rs!ps), "Kq2=Kq2+" + DoiDau(rs!ps)) + ",Kq3=" + DoiDau(sdn) + ",Kq4=" + DoiDau(sdc) + " WHERE MaSo=" + CStr(ms)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If tcuoi > 0 Then n1 = NgayCuoiThang(pNamTC, tcuoi) Else n1 = ncuoi
    If n <> n1 Then
        ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=" + CStr(n1 - n + 1) + " WHERE MaSo=" + CStr(ms)
    Else
        ExecuteSQL5 "UPDATE BaoCaoCP2 SET Cap=1 WHERE MaSo=" + CStr(ms)
    End If
    ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq5=Kq3*Cap"
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2 WHERE Kq3=0 AND Kq4=0"
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\BKTS" + IIf(mnt > 0, "2", "") + ".RPT"
    frmMain.Rpt.WindowTitle = "B¶ng kª tÝch sè"
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    frmMain.Rpt.Formulas(5) = "TaiKhoan='" + taikhoan.sohieu + " - " + taikhoan.Ten + IIf(mnt > 0, " - " + TenNT(mnt), "") + "'"
    frmMain.Rpt.Formulas(6) = "LS=" + DoiDau(Format(lsv, Mask_2))
    frmMain.Rpt.Formulas(7) = "LSG=" + DoiDau(Format(lsg, Mask_2))
    RptSetDate n1
    Set TK = Nothing
End Sub

Private Sub Baocao152(mkho As Long, tcuoi As Integer)
    Dim rs As Recordset, i As Integer, ncuoi As Date, shtk As String, shpl As String, kh As String, rs2 As Recordset, j As Integer, k As Integer
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "KIEMKE.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\KIEMKE.XLS", pCurDir + "KIEMKE.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "KIEMKE.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(3, 1) = "§Õn ngµy :" + Format(ncuoi, Mask_DR)
    
    xlsheet.Cells(4, 1) = "Kho :" + tenkho(mkho)
    Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW HethongTK.SoHieu AS SHTK, HethongTK.Ten, Vattu.SoHieu, TenVattu, DonVi, (Tien_" + CStr(CThangDB(tcuoi)) + ") AS Tien, (Luong_" + CStr(CThangDB(tcuoi)) + ") AS Luong FROM (TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo) INNER JOIN Vattu ON TonKho.MaVattu=Vattu.MaSo WHERE (HethongTK.SoHieu LIKE '152*' OR HethongTK.SoHieu LIKE '153*' OR HethongTK.SoHieu LIKE '155*' OR HethongTK.SoHieu LIKE '156*') AND (Tien_" + CStr(CThangDB(tcuoi)) + "<>0 OR Luong_" + CStr(CThangDB(tcuoi)) + "<>0) AND MaSoKho=" + CStr(mkho) + " ORDER BY HethongTK.SoHieu, Vattu.SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
c:
        j = j + 1
        If shtk <> rs!shtk Then
            shtk = rs!shtk
            xlsheet.Cells(7 + j, 2) = rs!Ten
            xlsheet.Range("A" + CStr(7 + j), "T" + CStr(7 + j)).Font.Bold = True
            GoTo c
        End If
        xlsheet.Cells(7 + j, 1) = CStr(j - 1)
        xlsheet.Cells(7 + j, 2) = rs!TenVattu
        xlsheet.Cells(7 + j, 3) = rs!sohieu
        xlsheet.Cells(7 + j, 4) = rs!DonVi
        xlsheet.Cells(7 + j, 5) = rs!luong
        xlsheet.Cells(7 + j, 6) = rs!tien
        rs.MoveNext
    Loop
    rs.Close
    With xlsheet.Range("A5", "T" + CStr(7 + j))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    XLSFooter xlsheet, 10 + j, 20, NgayCuoiThang(pNamTC, tcuoi)
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "KIEMKE.XLS"
KetThuc:
End Sub

Private Sub Baocao211(tcuoi As Integer)
    Dim rs As Recordset, i As Integer, ncuoi As Date, shtk As String, shpl As String, j As Integer
    Dim T(1 To 8) As Double
    
    ncuoi = NgayCuoiThang(pNamTC, tcuoi)
    Recycle pCurDir + "211.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\211.XLS", pCurDir + "211.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "211.XLS"
    On Error GoTo 0
    Set xlsheet = xlapp.Worksheets(1)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = "§Õn 0h ngµy : " + Format(ncuoi + 1, Mask_DR)
    
    SetSQL "SoDuDauCuoiKy", "SELECT DISTINCTROW TaiSan.Ten AS TenTS, TaiSan.SoHieu AS SoHieuTS, LoaiTaiSan.SoHieu AS SoHieuTK, LoaiTaiSan.MaSo AS MaTK, LoaiTaiSan.Ten AS TenTK, " _
                        & "ThongSo.NG_NS, ThongSo.NG_TBS, ThongSo.NG_CNK, ThongSo.NG_TD, ThongSo.CL_NS, ThongSo.CL_TBS, ThongSo.CL_CNK, ThongSo.CL_TD,LoaiTaiSan_1.SoHieu,LoaiTaiSan_1.Ten " _
                        & "FROM LoaiTaiSan AS LoaiTaiSan_1 RIGHT JOIN ((LoaiTaiSan RIGHT JOIN TaiSan ON LoaiTaiSan.MaSo = TaiSan.MaTaiKhoan) RIGHT JOIN ThongSo ON TaiSan.MaSo = ThongSo.MaTS) ON LoaiTaiSan_1.MaSo=TaiSan.MaLoai " _
                        & "WHERE " + WThang("ThangTang", 0, tcuoi) + " AND " + WThang2("ThangGiam", tcuoi, 0) + " AND ThongSo.Thang = " + CStr(CThangDB(tcuoi)) + " ORDER BY TaiSan.SoHieu"
                    
    Set rs = DBKetoan.OpenRecordset("SoDuDauCuoiKy", dbOpenSnapshot)
    Do While Not rs.EOF
a:
        If shtk <> rs!SoHieuTK Then
            shtk = rs!SoHieuTK
            xlsheet.Cells(9 + i, 2) = rs!tentk
            xlsheet.Cells(9 + i, 3) = shtk
            xlsheet.Range("A" + CStr(9 + i), "AE" + CStr(9 + i)).Font.Bold = True
            xlsheet.Cells(9 + i, 7) = rs!NG_NS
            xlsheet.Cells(9 + i, 8) = rs!NG_TBS
            xlsheet.Cells(9 + i, 9) = rs!NG_TD
            xlsheet.Cells(9 + i, 10) = rs!NG_CNK
            T(1) = T(1) + rs!NG_NS
            T(2) = T(2) + rs!NG_TBS
            T(3) = T(3) + rs!NG_TD
            T(4) = T(4) + rs!NG_CNK
            xlsheet.Range("F" + CStr(9 + i)).Formula = "=SUM(G" + CStr(9 + i) + ":J" + CStr(9 + i) + ")"
            xlsheet.Cells(9 + i, 12) = rs!CL_NS
            xlsheet.Cells(9 + i, 13) = rs!CL_TBS
            xlsheet.Cells(9 + i, 14) = rs!CL_TD
            xlsheet.Cells(9 + i, 15) = rs!CL_CNK
            T(5) = T(5) + rs!CL_NS
            T(6) = T(6) + rs!CL_TBS
            T(7) = T(7) + rs!CL_TD
            T(8) = T(8) + rs!CL_CNK
            xlsheet.Range("K" + CStr(9 + i)).Formula = "=SUM(L" + CStr(9 + i) + ":O" + CStr(9 + i) + ")"
            i = i + 1
            GoTo a
        End If
        If shpl <> rs!sohieu Then
            shpl = rs!sohieu
            xlsheet.Cells(9 + i, 2) = rs!Ten
            xlsheet.Cells(9 + i, 3) = rs!sohieu
            xlsheet.Range("A" + CStr(9 + i), "AE" + CStr(9 + i)).Font.Bold = True
            xlsheet.Range("A" + CStr(9 + i), "AE" + CStr(9 + i)).Font.Italic = True
            xlsheet.Cells(9 + i, 7) = rs!NG_NS
            xlsheet.Cells(9 + i, 8) = rs!NG_TBS
            xlsheet.Cells(9 + i, 9) = rs!NG_TD
            xlsheet.Cells(9 + i, 10) = rs!NG_CNK
            xlsheet.Range("F" + CStr(9 + i)).Formula = "=SUM(G" + CStr(9 + i) + ":J" + CStr(9 + i) + ")"
            xlsheet.Cells(9 + i, 12) = rs!CL_NS
            xlsheet.Cells(9 + i, 13) = rs!CL_TBS
            xlsheet.Cells(9 + i, 14) = rs!CL_TD
            xlsheet.Cells(9 + i, 15) = rs!CL_CNK
            xlsheet.Range("K" + CStr(9 + i)).Formula = "=SUM(L" + CStr(9 + i) + ":O" + CStr(9 + i) + ")"
            i = i + 1
            GoTo a
        End If
        rs.MoveNext
    Loop
    xlsheet.Range("A" + CStr(10 + i), "AE" + CStr(10 + i)).Font.Bold = True
    xlsheet.Cells(10 + i, 2) = "Tæng céng"
    xlsheet.Range("F" + CStr(10 + i)).Formula = "=SUM(G" + CStr(10 + i) + ":J" + CStr(10 + i) + ")"
    xlsheet.Cells(10 + i, 7) = T(1)
    xlsheet.Cells(10 + i, 8) = T(2)
    xlsheet.Cells(10 + i, 9) = T(3)
    xlsheet.Cells(10 + i, 10) = T(4)
    xlsheet.Range("K" + CStr(10 + i)).Formula = "=SUM(L" + CStr(10 + i) + ":O" + CStr(10 + i) + ")"
    xlsheet.Cells(10 + i, 12) = T(5)
    xlsheet.Cells(10 + i, 13) = T(6)
    xlsheet.Cells(10 + i, 14) = T(7)
    xlsheet.Cells(10 + i, 15) = T(8)
    With xlsheet.Range("A6", "AE" + CStr(10 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    XLSFooter xlsheet, 12 + i, 31, NgayCuoiThang(pNamTC, tcuoi)
    
    For i = 1 To 8
        T(i) = 0
    Next
    i = 0
    j = 0
    shtk = ""
    shpl = ""
    Set xlsheet = xlapp.Worksheets(2)
    xlsheet.Cells(1, 1) = pTenCty + " - " + pTenCn
    xlsheet.Cells(4, 1) = "§Õn 0h ngµy : " + Format(ncuoi + 1, Mask_DR)
    Set rs = DBKetoan.OpenRecordset("SoDuDauCuoiKy", dbOpenSnapshot)
    Do While Not rs.EOF
B:
        i = i + 1
        If shtk <> rs!SoHieuTK Then
            shtk = rs!SoHieuTK
            xlsheet.Cells(9 + i, 2) = rs!tentk
            xlsheet.Cells(9 + i, 3) = shtk
            xlsheet.Range("A" + CStr(9 + i), "AE" + CStr(9 + i)).Font.Bold = True
            xlsheet.Cells(9 + i, 7) = rs!NG_NS
            xlsheet.Cells(9 + i, 8) = rs!NG_TBS
            xlsheet.Cells(9 + i, 9) = rs!NG_TD
            xlsheet.Cells(9 + i, 10) = rs!NG_CNK
            xlsheet.Range("F" + CStr(9 + i)).Formula = "=SUM(G" + CStr(9 + i) + ":J" + CStr(9 + i) + ")"
            xlsheet.Cells(9 + i, 12) = rs!CL_NS
            xlsheet.Cells(9 + i, 13) = rs!CL_TBS
            xlsheet.Cells(9 + i, 14) = rs!CL_TD
            xlsheet.Cells(9 + i, 15) = rs!CL_CNK
            xlsheet.Range("K" + CStr(9 + i)).Formula = "=SUM(L" + CStr(9 + i) + ":O" + CStr(9 + i) + ")"
            T(1) = T(1) + rs!NG_NS
            T(2) = T(2) + rs!NG_TBS
            T(3) = T(3) + rs!NG_TD
            T(4) = T(4) + rs!NG_CNK
            T(5) = T(5) + rs!CL_NS
            T(6) = T(6) + rs!CL_TBS
            T(7) = T(7) + rs!CL_TD
            T(8) = T(8) + rs!CL_CNK
            GoTo B
        End If
        If shpl <> rs!sohieu Then
            shpl = rs!sohieu
            xlsheet.Cells(9 + i, 2) = rs!Ten
            xlsheet.Cells(9 + i, 3) = rs!sohieu
            xlsheet.Range("A" + CStr(9 + i), "AE" + CStr(9 + i)).Font.Bold = True
            xlsheet.Range("A" + CStr(9 + i), "AE" + CStr(9 + i)).Font.Italic = True
            xlsheet.Cells(9 + i, 7) = rs!NG_NS
            xlsheet.Cells(9 + i, 8) = rs!NG_TBS
            xlsheet.Cells(9 + i, 9) = rs!NG_TD
            xlsheet.Cells(9 + i, 10) = rs!NG_CNK
            xlsheet.Range("F" + CStr(9 + i)).Formula = "=SUM(G" + CStr(9 + i) + ":J" + CStr(9 + i) + ")"
            xlsheet.Cells(9 + i, 12) = rs!CL_NS
            xlsheet.Cells(9 + i, 13) = rs!CL_TBS
            xlsheet.Cells(9 + i, 14) = rs!CL_TD
            xlsheet.Cells(9 + i, 15) = rs!CL_CNK
            xlsheet.Range("K" + CStr(9 + i)).Formula = "=SUM(L" + CStr(9 + i) + ":O" + CStr(9 + i) + ")"
            GoTo B
        End If
        j = j + 1
        xlsheet.Cells(9 + i, 1) = CStr(j)
        xlsheet.Cells(9 + i, 2) = rs!TenTS
        xlsheet.Cells(9 + i, 3) = rs!SoHieuTS
        xlsheet.Cells(9 + i, 7) = rs!NG_NS
        xlsheet.Cells(9 + i, 8) = rs!NG_TBS
        xlsheet.Cells(9 + i, 9) = rs!NG_TD
        xlsheet.Cells(9 + i, 10) = rs!NG_CNK
        xlsheet.Range("F" + CStr(9 + i)).Formula = "=SUM(G" + CStr(9 + i) + ":J" + CStr(9 + i) + ")"
        xlsheet.Cells(9 + i, 12) = rs!CL_NS
        xlsheet.Cells(9 + i, 13) = rs!CL_TBS
        xlsheet.Cells(9 + i, 14) = rs!CL_TD
        xlsheet.Cells(9 + i, 15) = rs!CL_CNK
        xlsheet.Range("K" + CStr(9 + i)).Formula = "=SUM(L" + CStr(9 + i) + ":O" + CStr(9 + i) + ")"
        rs.MoveNext
    Loop
    xlsheet.Range("A" + CStr(10 + i), "AE" + CStr(10 + i)).Font.Bold = True
    xlsheet.Cells(10 + i, 2) = "Tæng céng"
    xlsheet.Range("F" + CStr(10 + i)).Formula = "=SUM(G" + CStr(10 + i) + ":J" + CStr(10 + i) + ")"
    xlsheet.Cells(10 + i, 7) = T(1)
    xlsheet.Cells(10 + i, 8) = T(2)
    xlsheet.Cells(10 + i, 9) = T(3)
    xlsheet.Cells(10 + i, 10) = T(4)
    xlsheet.Range("K" + CStr(10 + i)).Formula = "=SUM(L" + CStr(10 + i) + ":O" + CStr(10 + i) + ")"
    xlsheet.Cells(10 + i, 12) = T(5)
    xlsheet.Cells(10 + i, 13) = T(6)
    xlsheet.Cells(10 + i, 14) = T(7)
    xlsheet.Cells(10 + i, 15) = T(8)
    With xlsheet.Range("A6", "AE" + CStr(10 + i))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    XLSFooter xlsheet, 12 + i, 31, NgayCuoiThang(pNamTC, tcuoi)
    
    rs.Close
    Set rs = Nothing
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    CallExcel "211.XLS"
KetThuc:
End Sub

Private Sub InBangThopCN_HN(ndau As Date, ncuoi As Date, taikhoan As ClsTaikhoan, mpl As Long)
    Dim sql As String, TK As New ClsTaikhoan
        
    TinhCongNo ncuoi, taikhoan, 0
    If taikhoan.tk_id = TKCNPT_ID Then
        SetSQL "MienTru", "SELECT MaCT FROM ChungTu2  WHERE MaKHC>0 AND NgayGS>=#" + Format(ndau, Mask_DB) + "# AND NgayGS<=#" + Format(ncuoi, Mask_DB) + "# GROUP BY MaCT"
        SetSQL "QCdt", "SELECT MaKHC AS MaKH, HanTT AS HTT, NgayCT, Sum(SoPS-SoXuat) AS DuNo FROM (ChungTu2 INNER JOIN HethongTK ON ChungTu2.MaTKCo=HethongTK.MaSo) INNER JOIN MienTru ON ChungTu2.MaCT=MienTru.MaCT" _
            & " WHERE HethongTK.Kieu>0 AND HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*' GROUP BY MaKHC, HanTT, NgayCT  HAVING Sum(SoPS-SoXuat)>0 " _
            & " UNION SELECT MaKhachHang,HanTT AS HTT, CVDate('01/01/1980') AS NgayCT,Sum(CNDauNam2.DuCo_0-SoXuat) AS DuNo FROM CNDauNam2 INNER JOIN HethongTK ON CNDauNam2.MaTaiKhoan=HethongTK.MaSo " _
            & " WHERE HethongTK.Kieu>0 AND HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*' AND CNDauNam2.DuCo_0-SoXuat>0 GROUP BY MaKhachHang, HanTT"
    Else
        SetSQL "MienTru", "SELECT MaCT FROM ChungTu2 WHERE MaKH>0 AND NgayGS>=#" + Format(ndau, Mask_DB) + "# AND NgayGS<=#" + Format(ncuoi, Mask_DB) + "# GROUP BY MaCT"
        SetSQL "QCdt", "SELECT MaKH, HanTT AS HTT, NgayCT, Sum(SoPS-SoXuat) AS DuNo FROM (ChungTu2 INNER JOIN HethongTK ON ChungTu2.MaTKNo=HethongTK.MaSo) INNER JOIN MienTru ON ChungTu2.MaCT=MienTru.MaCT" _
            & " WHERE HethongTK.Kieu<0 AND HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*' GROUP BY MaKH, HanTT, NgayCT HAVING Sum(SoPS-SoXuat)>0 " _
            & " UNION SELECT MaKhachHang,HanTT AS HTT, CVDate('01/01/1980') AS NgayCT,Sum(CNDauNam2.DuNo_0-SoXuat) AS DuNo FROM CNDauNam2 INNER JOIN HethongTK ON CNDauNam2.MaTaiKhoan=HethongTK.MaSo " _
            & " WHERE HethongTK.Kieu<0 AND HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*' AND CNDauNam2.DuNo_0-SoXuat>0 GROUP BY MaKhachHang, HanTT"
    End If
    
    SetSQL "QChitiet", "SELECT DISTINCTROW KhachHang.SoHieu, KhachHang.Ten, PhanLoaiKhachHang.SoHieu AS SHPL, TenPhanLoai, QCdt.NgayCT, HTT, DuNo, PhanLoaiKhachHang.PLCha" _
        & " FROM (KhachHang INNER JOIN PhanLoaiKhachHang ON KhachHang.MaPhanLoai=PhanLoaiKhachHang.MaSo) INNER JOIN QCdt ON KhachHang.MaSo=QCdt.MaKH " _
        & " WHERE (True) " + IIf(mpl > 0, " AND MaPhanLoai=" + CStr(mpl), "") + " GROUP BY PhanLoaiKhachHang.Sohieu,KhachHang.SoHieu,KhachHang.Ten,TenPhanLoai,PLCha,NgayCT,HTT,DuNo"
    SetSQL "QSoCai", "SELECT QChitiet.*, PhanLoaiKhachHang.SoHieu AS SH2,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai2,PhanLoaiKhachHang.PLCha AS PLCha2 FROM QChitiet LEFT JOIN PhanLoaiKhachHang ON QChitiet.PLCha=PhanLoaiKhachHang.MaSo"
    SetSQL "QTongHopCT", "SELECT QSoCai.*, PhanLoaiKhachHang.SoHieu AS SH1,PhanLoaiKhachHang.TenPhanLoai AS TenPhanLoai1 FROM QSoCai LEFT JOIN PhanLoaiKhachHang ON QSoCai.PLCha2=PhanLoaiKhachHang.MaSo"
    
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\THSDCN5.RPT"
    
    frmMain.Rpt.Formulas(3) = "ThoiGian='Tõ ngµy " + Format(ndau, Mask_DR) + " ®Õn ngµy " + Format(ncuoi, Mask_DR) + "'"
    frmMain.Rpt.Formulas(4) = "Ngay1=Date(" + CStr(Year(ncuoi)) + "," + CStr(Month(ncuoi)) + "," + CStr(Day(ncuoi)) + ")"
    RptSetDate ncuoi
    If taikhoan.sohieu <> "" Then
        frmMain.Rpt.Formulas(5) = "TenBaoCao = IF PageNumber() = 1 THEN '" + ABCtoVNI("Tµi kho¶n: ") + taikhoan.sohieu + " - " + LbTenTk(7).Caption + "'"
    End If
End Sub

Private Sub InChitietCN_HN(mcn As Long, ndau As Date, ncuoi As Date, taikhoan As ClsTaikhoan)
    Dim sql As String
    
    TinhCongNo ncuoi, taikhoan, mcn
    
    If taikhoan.tk_id = TKCNPT_ID Then
        SetSQL "MienTru", "SELECT MaCT FROM ChungTu2 WHERE MaKHC=" + CStr(mcn) + " AND NgayGS>=#" + Format(ndau, Mask_DB) + "# AND NgayGS<=#" + Format(ncuoi, Mask_DB) + "# GROUP BY MaCT"
                             
        SetSQL "QCdt", "SELECT ChungTu2.SoHieu, NgayCT, DienGiai, First(HanTT) AS HTT, Sum(SoPS-SoXuat) AS DuNo FROM (ChungTu2 INNER JOIN HethongTK ON ChungTu2.MaTKCo=HethongTK.MaSo) INNER JOIN MienTru ON ChungTu2.MaCT=MienTru.MaCT" _
            & " WHERE HethongTK.Kieu>0 AND HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*' GROUP BY ChungTu2.SoHieu, NgayCT, DienGiai HAVING Sum(SoPS-SoXuat)>0" _
            & " UNION SELECT 'DK' AS SH, CVDate('01/01/1980') AS NgayCT, '" + ABCtoVNI("Nî ®Çu n¨m") + "',HanTT AS HTT,(CNDauNam2.DuCo_0-SoXuat) AS DuNo FROM CNDauNam2 INNER JOIN HethongTK ON CNDauNam2.MaTaiKhoan=HethongTK.MaSo " _
            & " WHERE HethongTK.Kieu>0 AND HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*' AND CNDauNam2.DuCo_0-SoXuat>0 AND MaKhachHang=" + CStr(mcn)
    Else
        SetSQL "MienTru", "SELECT MaCT FROM ChungTu2 WHERE MaKH=" + CStr(mcn) + " AND NgayGS>=#" + Format(ndau, Mask_DB) + "# AND NgayGS<=#" + Format(ncuoi, Mask_DB) + "# GROUP BY MaCT"
        SetSQL "QCdt", "SELECT ChungTu2.SoHieu, NgayCT, DienGiai, First(HanTT) AS HTT, Sum(SoPS-SoXuat) AS DuNo FROM (ChungTu2 INNER JOIN HethongTK ON ChungTu2.MaTKNo=HethongTK.MaSo) INNER JOIN MienTru ON ChungTu2.MaCT=MienTru.MaCT" _
            & " WHERE HethongTK.Kieu<0 AND HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*' GROUP BY ChungTu2.SoHieu, NgayCT, DienGiai HAVING Sum(SoPS-SoXuat)>0" _
            & " UNION SELECT 'DK' AS SH, CVDate('01/01/1980') AS NgayCT, '" + ABCtoVNI("Nî ®Çu n¨m") + "',HanTT AS HTT,(CNDauNam2.DuNo_0-SoXuat) AS DuNo FROM CNDauNam2 INNER JOIN HethongTK ON CNDauNam2.MaTaiKhoan=HethongTK.MaSo " _
            & " WHERE HethongTK.Kieu<0 AND HethongTK.SoHieu LIKE '" + taikhoan.sohieu + "*' AND CNDauNam2.DuNo_0-SoXuat>0 AND MaKhachHang=" + CStr(mcn)
    End If
    SetSQL "QTongHopCT", "SELECT * FROM QCdt ORDER BY NgayCT"
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\CTCN5.RPT"
    
    frmMain.Rpt.Formulas(3) = "ThoiGian='Tõ ngµy " + Format(ndau, Mask_DR) + " ®Õn ngµy " + Format(ncuoi, Mask_DR) + "'"
    frmMain.Rpt.Formulas(4) = "Ngay1=Date(" + CStr(Year(ncuoi)) + "," + CStr(Month(ncuoi)) + "," + CStr(Day(ncuoi)) + ")"
    RptSetDate ncuoi
    If taikhoan.sohieu <> "" Then
        frmMain.Rpt.Formulas(5) = "TenBaoCao = IF PageNumber() = 1 THEN '" + ABCtoVNI("Tµi kho¶n: ") + taikhoan.sohieu + " - " + taikhoan.Ten + "'"
    End If
    frmMain.Rpt.Formulas(6) = "DG = IF PageNumber() = 1 THEN '" + MaSo2SoHieu(mcn, "KhachHang") + ": " + TenKH("", mcn) + "'"
End Sub

Private Sub TinhCongNo(ncuoi As Date, TK As ClsTaikhoan, mkh As Long)
    Dim rs_ktra As Recordset, m As Long, sql As String
    
    XoaBang "ChungTu2"
    CopyTable2 "ChungTu", "ChungTu2"
    
    KtraCongNo
    
    ExecuteSQL5 "DELETE * FROM CNDauNam2"
    
    ExecuteSQL5 "INSERT INTO CNDauNam2 SELECT CNDauNam.* FROM CNDauNam INNER JOIN HethongTK ON CNDauNam.MaTaiKhoan=HethongTK.MaSo WHERE True " + IIf(TK.MaSo > 0, " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'", "") + IIf(mkh > 0, " AND MaKhachHang=" + CStr(mkh), "")
    ExecuteSQL5 "INSERT INTO ChungTu2 SELECT ChungTu.* FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE (TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ") AND NgayGS<=#" + Format(ncuoi, Mask_DB) + "#" + IIf(mkh > 0, " AND (MaKH=" + CStr(mkh) + " OR MaKHC=" + CStr(mkh) + ")", "") + IIf(TK.MaSo > 0, " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'", "")
    ExecuteSQL5 "INSERT INTO ChungTu2 SELECT ChungTu.* FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) LEFT JOIN ChungTu2 ON ChungTu.MaSo=ChungTu2.MaSo WHERE IsNULL(ChungTu2.MaSo) AND (TK_ID=" + CStr(TKCNKH_ID) + " OR TK_ID=" + CStr(TKCNPT_ID) + ") AND ChungTu.NgayGS<=#" + Format(ncuoi, Mask_DB) + "#" + IIf(mkh > 0, " AND (ChungTu.MaKH=" + CStr(mkh) + " OR ChungTu.MaKHC=" + CStr(mkh) + ")", "") + IIf(TK.MaSo > 0, " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'", "")
    
    ExecuteSQL5 "UPDATE CNDauNam2 SET SoXuat=0"
    sql = "SELECT ChungTu2.MaSo,MaTKNo,MaKH,SoPS,CT_ID FROM ChungTu2 INNER JOIN HethongTK ON ChungTu2.MaTKNo=HethongTK.MaSo WHERE ABS(CT_ID)>2000000000 AND TK_ID=" + CStr(TKCNPT_ID) + IIf(TK.MaSo > 0, " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'", "") + IIf(mkh > 0, " AND MaKH=" + CStr(mkh), "")
    Set rs_ktra = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs_ktra.EOF
        m = Abs(rs_ktra!CT_ID) - 2000000000
        If SelectSQL("SELECT MaSo AS F1 FROM CNDauNam2 WHERE MaSo=" + CStr(m)) = 0 Then
            m = SelectSQL("SELECT MaSo AS F1 FROM CNDauNam2 WHERE MaTaiKhoan=" + CStr(rs_ktra!MaTkNo) + " AND MaKhachHang=" + CStr(rs_ktra!makh) + " ORDER BY HanTT DESC")
            If m > 0 Then
                ExecuteSQL5 "UPDATE CNDauNam2 SET SoXuat=SoXuat+" + CStr(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
                ExecuteSQL5 "UPDATE ChungTu2 SET CT_ID=" + CStr(-(2000000000 - m)) + " WHERE MaSo=" + CStr(rs_ktra!MaSo)
            End If
        Else
            ExecuteSQL5 "UPDATE CNDauNam2 SET SoXuat=SoXuat+" + CStr(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
        End If
        rs_ktra.MoveNext
    Loop
     
     sql = "SELECT ChungTu2.MaSo,MaTKCo,MaKHC,SoPS,CT_ID FROM ChungTu2 INNER JOIN HethongTK ON ChungTu2.MaTKCo=HethongTK.MaSo WHERE ABS(CT_ID)>2000000000 AND TK_ID=" + CStr(TKCNKH_ID) + IIf(TK.MaSo > 0, " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'", "") + IIf(mkh > 0, " AND MaKHC=" + CStr(mkh), "")
     Set rs_ktra = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
     Do While Not rs_ktra.EOF
        m = Abs(rs_ktra!CT_ID) - 2000000000
        If SelectSQL("SELECT MaSo AS F1 FROM CNDauNam2 WHERE MaSo=" + CStr(m)) = 0 Then
            m = SelectSQL("SELECT MaSo AS F1 FROM CNDauNam2 WHERE MaTaiKhoan=" + CStr(rs_ktra!MaTkCo) + " AND MaKhachHang=" + CStr(rs_ktra!MaKHC) + " ORDER BY HanTT DESC")
            If m > 0 Then
                ExecuteSQL5 "UPDATE CNDauNam2 SET SoXuat=SoXuat+" + CStr(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
                ExecuteSQL5 "UPDATE ChungTu2 SET CT_ID=" + CStr(-(2000000000 - m)) + " WHERE MaSo=" + CStr(rs_ktra!MaSo)
            End If
        Else
            ExecuteSQL5 "UPDATE CNDauNam2 SET SoXuat=SoXuat+" + CStr(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
        End If
        rs_ktra.MoveNext
     Loop
     
     
     ExecuteSQL5 "UPDATE ChungTu2 INNER JOIN HethongTK ON ChungTu2.MaTKNo=HethongTK.MaSo SET ChungTu2.SoXuat=0 WHERE TK_ID=" + CStr(TKCNKH_ID) + IIf(TK.MaSo > 0, " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'", "") + IIf(mkh > 0, " AND MaKH=" + CStr(mkh), "")
     sql = "SELECT ChungTu2.MaSo,SoPS,CT_ID FROM ChungTu2 INNER JOIN HethongTK ON ChungTu2.MaTKCo=HethongTK.MaSo WHERE ABS(CT_ID)>0 AND TK_ID=" + CStr(TKCNKH_ID) + IIf(TK.MaSo > 0, " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'", "") + IIf(mkh > 0, " AND MaKHC=" + CStr(mkh), "")
     Set rs_ktra = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
     Do While Not rs_ktra.EOF
        m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNKH_ID) + " AND MaCT=" + CStr(Abs(rs_ktra!CT_ID)))
        ExecuteSQL5 "UPDATE ChungTu2 SET SoXuat=SoXuat+" + CStr(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
        rs_ktra.MoveNext
     Loop
     
     ExecuteSQL5 "UPDATE ChungTu2 INNER JOIN HethongTK ON ChungTu2.MaTKCo=HethongTK.MaSo SET ChungTu2.SoXuat=0 WHERE TK_ID=" + CStr(TKCNPT_ID) + IIf(TK.MaSo > 0, " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'", "") + IIf(mkh > 0, " AND MaKHC=" + CStr(mkh), "")
     sql = "SELECT ChungTu2.MaSo,SoPS,CT_ID FROM ChungTu2 INNER JOIN HethongTK ON ChungTu2.MaTKNo=HethongTK.MaSo WHERE ABS(CT_ID)>0 AND TK_ID=" + CStr(TKCNPT_ID) + IIf(TK.MaSo > 0, " AND HethongTK.SoHieu LIKE '" + TK.sohieu + "*'", "") + IIf(mkh > 0, " AND MaKH=" + CStr(mkh), "")
     Set rs_ktra = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
     Do While Not rs_ktra.EOF
        m = SelectSQL("SELECT TOP 1 ChungTu.MaSo AS F1 FROM ChungTu INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo WHERE TK_ID=" + CStr(TKCNPT_ID) + " AND MaCT=" + CStr(Abs(rs_ktra!CT_ID)))
        ExecuteSQL5 "UPDATE ChungTu2 SET SoXuat=SoXuat+" + CStr(rs_ktra!sops) + " WHERE MaSo=" + CStr(m)
        rs_ktra.MoveNext
     Loop
     
     rs_ktra.Close
     Set rs_ktra = Nothing
End Sub

Public Sub BKSanLuong(mk As Long, mloai As Long, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, shtk As String)
    Dim rs As Recordset, pl1 As String, pl2 As String, pl3 As String, wsql As String
    Dim r As Integer, tt As Integer, sokho As Integer, somn As Integer, i As Integer, luong As Double, tien As Double, lx As Double, tx As Double
    ReDim mkho(1 To 1) As Long
    ReDim mn(1 To 1) As Long
    Dim xlsheet2 As Worksheet, xlsheet3 As Worksheet
    
    Recycle pCurDir + "SANLUONG.XLS"
    On Error GoTo KetThuc
    FileCopy pCurDir + "REPORTS\SANLUONG.XLS", pCurDir + "SANLUONG.XLS"
    Set xlapp = CreateObject("Excel.Application")
    xlapp.Workbooks.Open pCurDir + "SANLUONG.XLS"
    Set xlsheet = xlapp.Worksheets(1)
    Set xlsheet2 = xlapp.Worksheets(2)
    Set xlsheet3 = xlapp.Worksheets(3)
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    
    On Error GoTo 0
    xlsheet.Cells(1, 1) = pTenCty
    xlsheet.Cells(2, 1) = pTenCn
    If mk > 0 Then
        xlsheet.Cells(3, 1) = tenkho(mk)
        xlsheet2.Cells(3, 1) = xlsheet.Cells(3, 1)
        xlsheet3.Cells(3, 1) = xlsheet.Cells(3, 1)
    End If
    xlsheet.Cells(5, 1) = IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn))
    xlsheet2.Cells(5, 1) = xlsheet.Cells(5, 1)
    xlsheet3.Cells(5, 1) = xlsheet.Cells(5, 1)
    
    If mk > 0 Then
        Set rs = DBKetoan.OpenRecordset("SELECT TenKho, KhoHang.MaSo FROM  (" + ChungTu2TKNC(1) + ") INNER JOIN KhoHang ON ChungTu.MaNguon=KhoHang.MaSo WHERE MaLoai=4 AND HethongTK.SoHieu LIKE '" + shtk + "*' AND " + wsql + IIf(mk > 0, " AND MaKho=" + CStr(mk), "") + " GROUP BY TenKho, KhoHang.MaSo", dbOpenSnapshot)
        Do While Not rs.EOF
            sokho = sokho + 1
            xlsheet.Range(XLSCol(11 + 2 * sokho) + "8", XLSCol(12 + 2 * sokho) + "8").MergeCells = True
            xlsheet.Cells(8, 11 + 2 * sokho) = rs!tenkho
            
            xlsheet2.Cells(8, 8 + sokho) = rs!tenkho
            xlsheet3.Cells(8, 8 + sokho) = rs!tenkho
            
            xlsheet.Cells(9, 11 + 2 * sokho) = "L­îng"
            xlsheet.Cells(9, 12 + 2 * sokho) = "TiÒn"
            ReDim Preserve mkho(1 To sokho) As Long
            mkho(sokho) = rs!MaSo
            rs.MoveNext
        Loop
        rs.Close
        If sokho > 0 Then xlsheet.Range("M7", XLSCol(12 + 2 * sokho) + "7").MergeCells = True
        If sokho > 1 Then
            xlsheet2.Range("I7", XLSCol(8 + sokho) + "7").MergeCells = True
            xlsheet3.Range("I7", XLSCol(8 + sokho) + "7").MergeCells = True
        End If
    End If
    
    Set rs = DBKetoan.OpenRecordset("SELECT NguonNhapXuat.DienGiai, NguonNhapXuat.MaSo FROM NguonNhapXuat INNER JOIN ChungTu ON NguonNhapXuat.MaSo=ChungTu.MaNguon WHERE ChungTu.MaLoai=8 AND " + wsql + IIf(mk > 0, " AND MaKho=" + CStr(mk), "") + " GROUP BY NguonNhapXuat.DienGiai, NguonNhapXuat.MaSo", dbOpenSnapshot)
    Do While Not rs.EOF
        somn = somn + 1
        xlsheet.Range(XLSCol(11 + 2 * sokho + 2 * somn) + "8", XLSCol(12 + 2 * sokho + 2 * somn) + "8").MergeCells = True
        xlsheet.Cells(8, 11 + 2 * sokho + 2 * somn) = rs!diengiai
        xlsheet.Cells(9, 11 + 2 * sokho + 2 * somn) = "L­îng"
        xlsheet.Cells(9, 12 + 2 * sokho + 2 * somn) = "TiÒn"
        
        xlsheet2.Cells(8, 8 + sokho + somn) = rs!diengiai
        xlsheet3.Cells(8, 8 + sokho + somn) = rs!diengiai
        
        ReDim Preserve mn(1 To somn) As Long
        mn(somn) = rs!MaSo
        rs.MoveNext
    Loop
    rs.Close
    If somn > 0 Then xlsheet.Range(XLSCol(13 + 2 * sokho) + "7", XLSCol(12 + 2 * sokho + 2 * somn) + "7").MergeCells = True
    If somn > 1 Then
        xlsheet2.Range(XLSCol(9 + sokho) + "7", XLSCol(8 + sokho + somn) + "7").MergeCells = True
        xlsheet3.Range(XLSCol(9 + sokho) + "7", XLSCol(8 + sokho + somn) + "7").MergeCells = True
    End If
    
    xlsheet.Cells(7, 13 + 2 * sokho) = "XuÊt b¸n theo kªnh ph©n phèi"
    xlsheet2.Cells(7, 9 + sokho) = xlsheet.Cells(7, 13 + 2 * sokho)
    xlsheet3.Cells(7, 9 + sokho) = xlsheet.Cells(7, 13 + 2 * sokho)
    
    'xlsheet.Range(XLSCol(13 + 2 * sokho + 2 * somn) + "8", XLSCol(14 + 2 * sokho + 2 * somn) + "8").MergeCells = True
    'xlsheet.Range(XLSCol(13 + 2 * sokho + 2 * somn) + "7", XLSCol(14 + 2 * sokho + 2 * somn) + "8").MergeCells = True
    'xlsheet.Cells(7, 13 + 2 * sokho + 2 * somn) = "XuÊt huû"
    'xlsheet.Cells(9, 13 + 2 * sokho + 2 * somn) = "L­îng"
    'xlsheet.Cells(9, 14 + 2 * sokho + 2 * somn) = "TiÒn"
    
    'xlsheet.Range(XLSCol(15 + 2 * sokho + 2 * somn) + "8", XLSCol(16 + 2 * sokho + 2 * somn) + "8").MergeCells = True
    'xlsheet.Range(XLSCol(15 + 2 * sokho + 2 * somn) + "7", XLSCol(16 + 2 * sokho + 2 * somn) + "8").MergeCells = True
    'xlsheet.Cells(7, 15 + 2 * sokho + 2 * somn) = "XuÊt kh¸c"
    'xlsheet.Cells(9, 15 + 2 * sokho + 2 * somn) = "L­îng"
    'xlsheet.Cells(9, 16 + 2 * sokho + 2 * somn) = "TiÒn"
    
    r = 10
    InLuyKeN mk, shtk, mloai, ndau, ncuoi
    Set rs = DBKetoan.OpenRecordset("SELECT * FROM QLuyKe ORDER BY SHPL1, SHPL2, SHPL3, SoHieu", dbOpenSnapshot)
    Do While Not rs.EOF
        If Not IsNull(rs!SHPL1) Then
            If pl1 <> rs!SHPL1 Then
                pl1 = rs!SHPL1
                r = r + 1
                xlsheet.Cells(r, 2) = rs!TenPL1
                xlsheet.Cells(r, 3) = rs!SHPL1
                xlsheet.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Bold = True
                
                xlsheet2.Cells(r, 2) = rs!TenPL1
                xlsheet2.Cells(r, 3) = rs!SHPL1
                xlsheet2.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Bold = True
                
                xlsheet3.Cells(r, 2) = rs!TenPL1
                xlsheet3.Cells(r, 3) = rs!SHPL1
                xlsheet3.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Bold = True
            End If
        End If
        
        If Not IsNull(rs!SHPL2) Then
            If pl2 <> rs!SHPL2 Then
                pl2 = rs!SHPL2
                r = r + 1
                xlsheet.Cells(r, 2) = rs!TenPL2
                xlsheet.Cells(r, 3) = rs!SHPL2
                xlsheet.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Bold = True
                xlsheet.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Italic = True
                
                xlsheet2.Cells(r, 2) = rs!TenPL2
                xlsheet2.Cells(r, 3) = rs!SHPL2
                xlsheet2.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Bold = True
                xlsheet2.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Italic = True
                
                xlsheet3.Cells(r, 2) = rs!TenPL2
                xlsheet3.Cells(r, 3) = rs!SHPL2
                xlsheet3.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Bold = True
                xlsheet3.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Italic = True
            End If
        End If
        
        If Not IsNull(rs!SHPL3) Then
            If pl3 <> rs!SHPL3 Then
                pl3 = rs!SHPL3
                r = r + 1
                xlsheet.Cells(r, 2) = rs!TenPL3
                xlsheet.Cells(r, 3) = rs!SHPL3
                xlsheet.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Bold = True
                xlsheet.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Italic = True
                
                xlsheet2.Cells(r, 2) = rs!TenPL3
                xlsheet2.Cells(r, 3) = rs!SHPL3
                xlsheet2.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Bold = True
                xlsheet2.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Italic = True
                
                xlsheet3.Cells(r, 2) = rs!TenPL3
                xlsheet3.Cells(r, 3) = rs!SHPL3
                xlsheet3.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Bold = True
                xlsheet3.Range("A" + CStr(r), XLSCol(18 + 2 * sokho) + CStr(r)).Font.Italic = True
            End If
        End If
                
        tt = tt + 1
        r = r + 1
        
        xlsheet.Cells(r, 1) = CStr(tt)
        xlsheet.Cells(r, 2) = rs!TenVattu
        xlsheet.Cells(r, 3) = rs!sohieu
        xlsheet.Cells(r, 4) = rs!DonVi
        
        xlsheet2.Cells(r, 1) = CStr(tt)
        xlsheet2.Cells(r, 2) = rs!TenVattu
        xlsheet2.Cells(r, 3) = rs!sohieu
        xlsheet2.Cells(r, 4) = rs!DonVi
        
        xlsheet3.Cells(r, 1) = CStr(tt)
        xlsheet3.Cells(r, 2) = rs!TenVattu
        xlsheet3.Cells(r, 3) = rs!sohieu
        xlsheet3.Cells(r, 4) = rs!DonVi
        
        xlsheet.Cells(r, 5) = rs!DkLuong
        xlsheet.Cells(r, 6) = rs!Dktien
        xlsheet.Cells(r, 7) = rs!LuongNhap
        xlsheet.Cells(r, 8) = rs!TienNhap
        xlsheet.Cells(r, 9) = rs!LuongXuat
        xlsheet.Cells(r, 10) = rs!TienXuat
        xlsheet.Cells(r, 11) = rs!ckluong
        xlsheet.Cells(r, 12) = rs!cktien
        
        xlsheet2.Cells(r, 5) = rs!DkLuong
        xlsheet3.Cells(r, 5) = rs!Dktien
        xlsheet2.Cells(r, 6) = rs!LuongNhap
        xlsheet3.Cells(r, 6) = rs!TienNhap
        xlsheet2.Cells(r, 7) = rs!LuongXuat
        xlsheet3.Cells(r, 7) = rs!TienXuat
        xlsheet2.Cells(r, 8) = rs!ckluong
        xlsheet3.Cells(r, 8) = rs!cktien
        
        lx = 0
        tx = 0
        If mk > 0 Then
            For i = 1 To sokho
                luong = XuatTheoKho(mk, mkho(i), rs!mvt, ndau, ncuoi, tien)
                xlsheet.Cells(r, 11 + 2 * i) = luong
                xlsheet.Cells(r, 12 + 2 * i) = tien
                
                xlsheet2.Cells(r, 8 + i) = luong
                xlsheet3.Cells(r, 8 + i) = tien
                
                lx = lx + luong
                tx = tx + tien
            Next
        End If
        
        For i = 1 To somn
            luong = XuatBan(shtk, mk, mn(i), rs!mvt, ndau, ncuoi, tien)
            xlsheet.Cells(r, 11 + 2 * sokho + 2 * i) = luong
            xlsheet.Cells(r, 12 + 2 * sokho + 2 * i) = tien
            
            xlsheet2.Cells(r, 8 + sokho + i) = luong
            xlsheet3.Cells(r, 8 + sokho + i) = tien
            
            lx = lx + luong
            tx = tx + tien
        Next
                    
        'xlsheet.Range(XLSCol(17 + 2 * sokho) + CStr(r)).Formula = "=I" + CStr(r) + "-" + CStr(lx)
        'xlsheet.Range(XLSCol(18 + 2 * sokho) + CStr(r)).Formula = "=J" + CStr(r) + "-" + CStr(tx)
        
        rs.MoveNext
    Loop
    
    If r > 1 Then
        xlsheet.Cells(r + 1, 2) = ABCtoVNI("Tæng céng")
        xlsheet2.Cells(r + 1, 2) = xlsheet.Cells(r + 1, 2)
        xlsheet3.Cells(r + 1, 2) = xlsheet.Cells(r + 1, 2)
        For i = 6 To 12 + 2 * sokho + 2 * somn
            xlsheet.Range(XLSCol(i) + CStr(r + 1)).Formula = "=Sum(" + XLSCol(i) + "10:" + XLSCol(i) + CStr(r) + ")"
        Next
        For i = 5 To 8 + sokho + somn
            xlsheet2.Range(XLSCol(i) + CStr(r + 1)).Formula = "=Sum(" + XLSCol(i) + "10:" + XLSCol(i) + CStr(r) + ")"
            xlsheet3.Range(XLSCol(i) + CStr(r + 1)).Formula = "=Sum(" + XLSCol(i) + "10:" + XLSCol(i) + CStr(r) + ")"
        Next
    End If
    
    With xlsheet.Range("A7", XLSCol(12 + 2 * sokho + 2 * somn) + CStr(r + 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    
    With xlsheet2.Range("A7", XLSCol(8 + sokho + somn) + CStr(r + 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    
    With xlsheet3.Range("A7", XLSCol(8 + sokho + somn) + CStr(r + 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 18
        .VerticalAlignment = xlVAlignCenter
    End With
    
    rs.Close
    Set rs = Nothing
                
    xlsheet.Range("A1", XLSCol(12 + 2 * sokho + 2 * somn) + CStr(r + 1)).Copy
    xlapp.Worksheets.Add count:=1, before:=xlapp.Sheets(4)
    Set xlsheet = xlapp.Worksheets(4)
    xlsheet.Range("A1", XLSCol(12 + 2 * sokho + 2 * somn) + CStr(r + 1)).PasteSpecial
    xlsheet.Cells(4, 1) = "PhiÕu xuÊt thµnh phÈm"
    xlsheet.Range("E1", "H" + CStr(r + 1)).Delete
    xlsheet.Range("K1", "L" + CStr(r + 1)).Delete
    
    xlapp.Workbooks(1).Save
    xlapp.Workbooks.Close
    
    Set xlsheet2 = Nothing
    Set xlsheet3 = Nothing
    
    CallExcel "SANLUONG.XLS"
KetThuc:
    Erase mkho
    Erase mn
End Sub

Private Sub T(mtk As Long, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, cap As Integer)
    Dim dn As Double, DC As Double, rs As Recordset, L As Integer, psn As Double, psc As Double, shtk As String, sqlw As String
    Dim tpsn As Double, tpsc As Double, sh As String, i As Integer
    
    shtk = MaSo2SoHieu(mtk, "HethongTK")
    ExecuteSQL5 "DELETE * FROM BKNhomPS"
    If tdau > 0 Then
        SoDuTK mtk, tdau - 1, dn, DC
        tpsn = SoPSTK(shtk, tdau, tcuoi, -1)
        tpsc = SoPSTK(shtk, tdau, tcuoi, 1)
    Else
        SoDuTKNgay mtk, ndau - 1, dn, DC
        SoPSTKN shtk, ndau, ncuoi, tpsn, tpsc
    End If
    ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,SoHieu,SoPSNo, SoPSCo) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",0,'Dö ñaàu'," + CStr(dn) + "," + CStr(DC) + ")"
    dn = dn - DC
    sqlw = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    
    Set rs = DBKetoan.OpenRecordset("SELECT TK.Sohieu,SUM(SoPS) AS PS FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) INNER JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo WHERE HethongTK.Sohieu LIKE '" + CStr(shtk) + "*' AND " + sqlw + " AND (ChungTu.MaLoai<>4 OR MaTKNo<>MaTKCo) GROUP BY TK.Sohieu", dbOpenSnapshot)
    Do While Not rs.EOF
        For i = 1 To cap
            sh = LaySHTKCapTren(rs!sohieu, i)
            ExecuteSQL5 "UPDATE BKNhomPS SET SoPSNo=SoPSNo+" + DoiDau(rs!ps) + " WHERE MaCT=1 AND Sohieu='" + Space(2 * (i - 1)) + sh + "'"
            If DBKetoan.RecordsAffected = 0 Then ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,Sohieu,SoPSNo) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",1,'" + Space(2 * (i - 1)) + sh + "'," + CStr(rs!ps) + ")"
        Next
        psn = psn + rs!ps
        rs.MoveNext
    Loop
    If tpsn > psn Then ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,Sohieu,SoPSNo) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",1,'xxx'," + CStr(tpsn - psn) + ")"
    
    Set rs = DBKetoan.OpenRecordset("SELECT HethongTK.Sohieu,SUM(SoPS) AS PS FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) INNER JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo WHERE TK.Sohieu LIKE '" + CStr(shtk) + "*' AND " + sqlw + " AND (ChungTu.MaLoai<>4 OR MaTKNo<>MaTKCo) GROUP BY HethongTK.Sohieu", dbOpenSnapshot)
    Do While Not rs.EOF
        For i = 1 To cap
            sh = LaySHTKCapTren(rs!sohieu, i)
            ExecuteSQL5 "UPDATE BKNhomPS SET SHTK=SoHieu,SoPSCo=SoPSCo+" + CStr(rs!ps) + " WHERE Sohieu='" + Space(2 * (i - 1)) + sh + "'"
            If DBKetoan.RecordsAffected = 0 Then ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,SoHieu,SHTK,SoPSCo) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",1,'" + Space(2 * (i - 1)) + sh + "','" + Space(2 * (i - 1)) + sh + "'," + CStr(rs!ps) + ")"
        Next
        psc = psc + rs!ps
        rs.MoveNext
    Loop
    If tpsc > psc Then ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,Sohieu,SoPSNo) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",1,'xxx'," + CStr(tpsc - psc) + ")"
    
    ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,SoHieu,SoPSNo, SoPSCo) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",2,'Coäng PS'," + CStr(tpsn) + "," + CStr(tpsc) + ")"
    dn = dn + tpsn - tpsc
    ExecuteSQL5 "INSERT INTO BKNhomPS (MaSo,MaCT,SoHieu,SoPSNo, SoPSCo) VALUES (" + CStr(Lng_MaxValue("MaSo", "BKNhomPS") + 1) + ",3,'Dö cuoái'," + CStr(IIf(dn >= 0, dn, 0)) + "," + CStr(IIf(dn < 0, -dn, 0)) + ")"
    rs.Close
    Set rs = Nothing
    
    frmMain.Rpt.ReportFileName = "T.RPT"
    
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    frmMain.Rpt.Formulas(4) = "TK='TK: " + shtk + "'"
End Sub

Private Sub InCdt2(tdau As Integer, tcuoi As Integer, cap As Integer, ct As Integer)
    Dim st As String, sqln As String, sqlc As String, i As Integer, sqlnx As String, sqlcx As String
    
    GauGe.Max = 3
    SoDuTKCN2 ThangTruoc(tdau)
    SoDuTKCN2 tcuoi
    
    For i = CThangDB(tdau) To CThangDB(tcuoi)
        st = CStr(i)
        sqln = sqln + " + HethongTk.No_" + st
        sqlc = sqlc + " + HethongTk.Co_" + st
    Next
    
    For i = 1 To CThangDB(tcuoi)
        st = CStr(i)
        sqlnx = sqlnx + " + HethongTk.No_" + st
        sqlcx = sqlcx + " + HethongTk.Co_" + st
    Next
    
    st = CStr(CThangDB(tcuoi))
    GauGe.Value = 1
    SetSQL "QCdt", "SELECT DISTINCTROW HeThongTK.SoHieu, HeThongTK.Cap, HeThongTK.Ten" + IIf(nn > 0, "E", "") + "," _
        & " HeThongTK.Kieu, HeThongTK.Loai, HeThongTK.DuNo_0 As DkNo, HeThongTK.DuCo_0 As DkCo, " + sqln + " As PsNo, " + sqlc + " As PsCo, " + sqlnx + " As PsNoX, " + sqlcx + " As PsCoX," _
        & " HeThongTK.DuNo_" + st + " As CkNo, HeThongTK.DuCo_" + st + " As CkCo From HeThongTK " _
        & " Where ((HeThongTK.MaTC = 0 Or HeThongTK.MaTC = HeThongTK.MaSo) Or (TK_ID3 Mod 10 >= 1)) And (HeThongTK.loai > 0) And cap <= " + CStr(cap) _
        + IIf(ct = 0, " AND (DuNo_" + st + "<>0 OR DuCo_" + st + "<>0 OR (" + sqln + ")<>0 OR (" + sqlc + ")<>0)", "")
    GauGe.Value = 2
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
        
    frmMain.Rpt.ReportFileName = "CDT2.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    GauGe.Value = 3
End Sub

Private Function SoChiTiet3(TK As String, tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date) As Boolean
    Dim rs As Recordset, sql As String, i As Integer, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    sql = "SELECT DISTINCTROW LEFT(HethongTK.SoHieu,3) AS SH FROM (ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) INNER JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo" _
        & " WHERE TK.SoHieu LIKE '" + TK + "*' AND TK.TK_ID=" + CStr(TKDT_ID) + " AND " + wsql + "  GROUP BY LEFT(HethongTK.SoHieu,3)"
    Set rs = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs.recordCount = 0 Then GoTo KgIn
    Do While Not rs.EOF And i < 5
        i = i + 1
        frmMain.Rpt.Formulas(2 + i) = "TK" + CStr(i) + "='" + rs!sh + "'"
        rs.MoveNext
    Loop
    frmMain.Rpt.Formulas(3 + i) = "SoHieuTK='" + TK + "'"
    frmMain.Rpt.Formulas(4 + i) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    sql = "SELECT DISTINCTROW LEFT(HethongTK.SoHieu,3) AS SH,ChungTu.SoHieu,NgayCT,DienGiai,SoPS,SoPS2Co,TenVattu, DonVi,Vattu.SoHieu AS SHVT FROM ((ChungTu INNER JOIN HethongTK ON ChungTu.MaTKNo=HethongTK.MaSo) INNER JOIN HethongTK AS TK ON ChungTu.MaTKCo=TK.MaSo) INNER JOIN Vattu ON ChungTu.Mavattu=Vattu.MaSo" _
        & " WHERE TK.SoHieu LIKE '" + TK + "*' AND TK.TK_ID=" + CStr(TKDT_ID) + " AND " + wsql + "  ORDER BY LEFT(HethongTK.SoHieu,3)"
    SetSQL "QSoCai", sql
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\CHITIET3.RPT"
    RptSetDate ncuoi
    frmMain.Rpt.WindowTitle = "Sæ chi tiÕt sè 3"
    rs.Close
    Set rs = Nothing
    SoChiTiet3 = True
    Exit Function
KgIn:
    SoChiTiet3 = False
    rs.Close
End Function

Private Sub TinhPB(thang As Integer)
    Dim rs As Recordset, K1 As Double, k2 As Double, k3 As Double, k4 As Double, k5 As Double, k6 As Double
    
    ExecuteSQL5 "DELETE * FROM BaoCaoCP"
    
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,CoCon,Ten) VALUES (10000,'10000',1,'I. Soá khaáu hao ñaõ trích thaùng tröôùc')"
    
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,Ten,Cap,Kq1,Kq2,BC_ID,MaCha) SELECT MaTS+10000,Cstr(MaTS+10000),Ten,NamKH,(NG_NS+NG_TBS+NG_TD+NG_CNK),(KH_NS+KH_TBS+KH_TD+KH_CNK),MaDTSD,10000 " _
        & " FROM ThongSo INNER JOIN TaiSan ON ThongSo.MaTS = TaiSan.MaSo WHERE (KH_NS+KH_TBS+KH_TD+KH_CNK)<>0 AND Thang=" + CStr(IIf(thang > 1, thang - 1, 1)) + " AND ThangTang<" + CStr(thang)
    
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,CoCon,Ten) VALUES (20000,'20000',1,'II. Soá khaáu hao TSCÑ taêng trong thaùng')"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,Ten,Cap,Kq1,Kq2,BC_ID,MaCha) SELECT MaTS+20000,Cstr(MaTS+20000),Ten,NamKH,(NG_NS+NG_TBS+NG_TD+NG_CNK),(KH_NS+KH_TBS+KH_TD+KH_CNK),MaDTSD,20000 " _
        & " FROM ThongSo INNER JOIN TaiSan ON ThongSo.MaTS = TaiSan.MaSo WHERE (KH_NS+KH_TBS+KH_TD+KH_CNK)<>0 AND Thang=" + CStr(thang) + " AND ThangTang=" + CStr(thang)
    
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,CoCon,Ten) VALUES (30000,'30000',1,'III. Soá khaáu hao TSCÑ giaûm trong thaùng')"
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,Ten,Cap,Kq1,Kq2,BC_ID,MaCha) SELECT MaTS+30000,Cstr(MaTS+30000),Ten,NamKH,(NG_NS+NG_TBS+NG_TD+NG_CNK),(KH_NS+KH_TBS+KH_TD+KH_CNK),MaDTSD,30000 " _
        & " FROM ThongSo INNER JOIN TaiSan ON ThongSo.MaTS = TaiSan.MaSo WHERE ((KH_NS+KH_TBS+KH_TD+KH_CNK)<>0 AND Thang=" + CStr(thang - 1) + " AND ThangGiam=" + CStr(thang - 1) + ") OR ((KH_NS+KH_TBS+KH_TD+KH_CNK)=0 AND Thang=" + CStr(thang) + " AND ThangGiam=" + CStr(thang) + ")"
            
    ExecuteSQL5 "INSERT INTO BaoCaoCP (MaSo,SoHieu,CoCon,Ten) VALUES (40000,'40000',1,'IV. Soá khaáu hao TSCÑ phaûi trích thaùng naøy')"
    
    ExecuteSQL5 "UPDATE BaoCaoCP INNER JOIN HethongTK ON BaoCaoCP.BC_ID=HethongTK.MaSo SET Kq3=IIF(LEFT(HethongTK.SoHieu,3)='627' ,Kq2,0),Kq5=IIF(LEFT(HethongTK.SoHieu,3)='641' ,Kq2,0),Kq6=IIF(LEFT(HethongTK.SoHieu,3)='642' ,Kq2,0)"
    Set rs = DBKetoan.OpenRecordset("SELECT MaCha,Sum(Kq1) AS K1,Sum(Kq2) AS K2,Sum(Kq3) AS K3,Sum(Kq4) AS K4,Sum(Kq5) AS K5,Sum(Kq6) AS K6 " _
        & " FROM BaoCaoCP WHERE MaCha>0 GROUP BY MaCha", dbOpenSnapshot)
    Do While Not rs.EOF
        ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1=" + CStr(rs!K1) + ",Kq2=" + CStr(rs!k2) + ",Kq3=" + CStr(rs!k3) + ",Kq4=" + CStr(rs!k4) + ",Kq5=" + CStr(rs!k5) + ",Kq6=" + CStr(rs!k6) + " WHERE MaSo=" + CStr(rs!macha)
        If rs!macha < 30000 Then
            K1 = K1 + rs!K1
            k2 = k2 + rs!k2
            k3 = k3 + rs!k3
            k4 = k4 + rs!k4
            k5 = k5 + rs!k5
            k6 = k6 + rs!k6
        Else
            K1 = K1 - rs!K1
            k2 = k2 - rs!k2
            k3 = k3 - rs!k3
            k4 = k4 - rs!k4
            k5 = k5 - rs!k5
            k6 = k6 - rs!k6
        End If
        rs.MoveNext
    Loop
    ExecuteSQL5 "UPDATE BaoCaoCP SET Kq1=" + CStr(K1) + ",Kq2=" + CStr(k2) + ",Kq3=" + CStr(k3) + ",Kq4=" + CStr(k4) + ",Kq5=" + CStr(k5) + ",Kq6=" + CStr(k6) + " WHERE MaSo=40000"
    rs.Close
    Set rs = Nothing
End Sub

Private Sub InCTKhuyenMai(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date, mkh As Long, mvt As Long)
    Dim sql As String, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
    
    sql = "SELECT DISTINCTROW ChungTu.NgayCT, ChungTu.SoHieu, ChungTu.SoPS, ChungTu.SoPS2Co, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, KhachHang.SoHieu AS SHKH, KhachHang.Ten" _
        & " FROM ((ChungTu INNER JOIN Vattu ON ChungTu.MaVattu = Vattu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo) LEFT JOIN KhachHang ON ChungTu.MaKH=KhachHang.MaSo" _
        & " WHERE SoPS=0 AND TK_ID=" + CStr(TKDT_ID) + " AND " + wsql + IIf(mkh > 0, " AND MaKH=" + CStr(mkh), "") + IIf(mvt > 0, " AND MaVattu=" + CStr(mvt), "") + " ORDER BY ChungTu.NgayCT, ChungTu.SoHieu"
    
    SetSQL "QChitiet", sql
    
    frmMain.Rpt.WindowTitle = "Tæng hîp chiÕt khÊu b¸n hµng"
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    If mkh > 0 Then frmMain.Rpt.Formulas(4) = "DG = IF PageNumber() = 1 THEN 'KH: " + MaSo2SoHieu(mkh, "KhachHang") + " - " + TenKH("", mkh) + "'"
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\CTCK.RPT"
    RptSetDate IIf(tdau > 0, NgayCuoiThang(pNamTC, tcuoi), ncuoi)
End Sub

Private Sub InTHChietKhau(tdau As Integer, tcuoi As Integer, ndau As Date, ncuoi As Date)
    Dim sql As String, wsql As String
    
    wsql = IIf(tdau > 0, WThang("ThangCT", tdau, tcuoi), WNgay("NgayGS", ndau, ncuoi))
        
    sql = "SELECT DISTINCTROW ChungTu.MaCT, ChungTu.SoPS, ChungTu.SoPS2Co, Vattu.SoHieu, Vattu.TenVattu, Vattu.DonVi, HethongTK.SoHieu,HethongTK.Ten" _
        & " FROM (ChungTu LEFT JOIN Vattu ON ChungTu.MaVattu = Vattu.MaSo) INNER JOIN HethongTK ON ChungTu.MaTKCo=HethongTK.MaSo" _
        & " WHERE TK_ID=" + CStr(TKDT_ID) + " AND SoPS=0 AND " + wsql + " ORDER BY ChungTu.NgayGS, ChungTu.MaCT"
    
    SetSQL "QChitiet", sql
    
    'frmMain.Rpt.WindowTitle = "Tæng hîp chiÕt khÊu b¸n hµng"
    frmMain.Rpt.Formulas(3) = "ThoiGian = IF PageNumber() = 1 THEN '" + IIf(tdau > 0, ThoiGian(tdau, tcuoi, nn), ThoiGianN(ndau, ncuoi, nn)) + "'"
    frmMain.Rpt.ReportFileName = pCurDir + "REPORTS\THCK.RPT"
    RptSetDate IIf(tdau > 0, NgayCuoiThang(pNamTC, tcuoi), ncuoi)
End Sub

Private Sub CboVV_Click(Index As Integer)
    If Index = 0 Then CboVVClick CboVV(0), CboVV(1)
End Sub

Private Sub ChiPhiGiaThanh(mtp As Long, tdau As Integer, tcuoi As Integer)
    Dim i As Integer, tp As New Cls154, sanluong As Double, CP As Double, rs As Recordset, shpl As String, tien As Double, tienx As Double
    
    tp.InitTPMaSo mtp
    i = 1
    ExecuteSQL5 "DELETE * FROM BaoCaoCP2"
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,Cap,Kq2) VALUES (1,'621','" + ABCtoVNI("Nguyªn vËt liÖu trùc tiÕp") + "',1," + DoiDau(tp.SoCPNVL(tdau, tcuoi)) + ")"
    Set rs = DBKetoan.OpenRecordset("SELECT PhanLoaiVattu.SoHieu,TenPhanLoai,MaPhanLoai,Vattu.Sohieu AS SHVT,TenVattu,Sum(IIF(MaLoai=2,SoPS2Co,-SoPS2No)) AS Luong, Sum(IIF(MaLoai=2,SoPS,-SoPS)) AS Tien " _
        & " FROM ((" + ChungTu2TKNC(0) + ") INNER JOIN Vattu ON ChungTu.MaVattu=Vattu.MaSo) INNER JOIN PhanLoaiVattu ON Vattu.MaPhanLoai=PhanLoaiVattu.MaSo " _
        & " WHERE ((ChungTu.MaLoai=2 AND HethongTK.SoHieu LIKE '621*') OR (ChungTu.MaLoai=1 AND HethongTK.SoHieu LIKE '152' AND TK.SoHieu LIKE '621*')) AND MaTP=" + CStr(mtp) + " AND " + WThang("ThangCT", tdau, tcuoi) + " GROUP BY PhanLoaiVattu.SoHieu,TenPhanLoai,MaPhanLoai,Vattu.Sohieu,TenVattu", dbOpenSnapshot)
    Do While Not rs.EOF
        If shpl <> rs!sohieu Then
            If Len(shpl) > 0 Then ExecuteSQL5 "UPDATE BaoCaoCP2 SET Kq2=" + DoiDau(tien) + " WHERE SoHieu='" + shpl + "'"
            shpl = rs!sohieu
            tien = 0
            i = i + 1
            ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,Cap,MK) VALUES (" + CStr(i) + ",'" + rs!sohieu + "','" + rs!TenPhanLoai + "',2," + CStr(rs!MaPhanLoai) + ")"
        End If
        tien = tien + rs!tien
        tienx = tienx + rs!tien
        i = i + 1
        ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,Cap,MaCha,Kq1,Kq2) VALUES (" + CStr(i) + ",'" + rs!shvt + "','" + rs!TenVattu + "',3," + CStr(rs!MaPhanLoai) + "," + DoiDau(rs!luong) + "," + DoiDau(rs!tien) + ")"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    i = i + 1
    tien = tp.SoCPNC(tdau, tcuoi) + tp.SoCPNCPB(tdau, tcuoi)
    tienx = tienx + tien
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,Cap,Kq2) VALUES (" + CStr(i) + ",'622','" + ABCtoVNI("Chi phÝ nh©n c«ng") + "',1," + DoiDau(tien) + ")"
    i = i + 1
    tien = tp.SoCPSXC(tdau, tcuoi) + tp.SoCPSXCTT(tdau, tcuoi)
    tienx = tienx + tien
    ExecuteSQL5 "INSERT INTO BaoCaoCP2 (MaSo,SoHieu,Ten,Cap,Kq2) VALUES (" + CStr(i) + ",'627','" + ABCtoVNI("Chi phÝ s¶n xuÊt chung") + "',1," + DoiDau(tien) + ")"
        
    frmMain.Rpt.Formulas(3) = "ThoiGian='" + ThoiGian(tdau, tcuoi, nn) + "'"
    frmMain.Rpt.Formulas(4) = "DienGiai='" + tp.sohieu + " - " + tp.TenVattu + " - " + tp.DonVi + "'"
    frmMain.Rpt.Formulas(5) = "SanLuong=" + DoiDau(tp.SoSanLuong(tdau, tcuoi))
    frmMain.Rpt.Formulas(6) = "TCP=" + DoiDau(tienx)
    
    frmMain.Rpt.ReportFileName = "CPGT.RPT"
    RptSetDate NgayCuoiThang(pNamTC, tcuoi), nn
    
    Set tp = Nothing
End Sub
