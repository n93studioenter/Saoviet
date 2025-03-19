VERSION 5.00
Begin VB.Form frmTonDauDaTaBaSE 
   Caption         =   "frmTonDauDaTaBaSE"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   Icon            =   "frmTonDauDaTaBaSE.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5130
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Chuyen ton kho"
      Height          =   855
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Chuyen khach hang"
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Xoa ton dau khach hang"
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "xoa ton kho"
      Height          =   735
      Left            =   5160
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frmTonDauDaTaBaSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
' Dim psw As String, fn As String
'  Dim rs_chungtu As Recordset
'  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
'  'Dim DBKetoan1 As Database
'  'Set DBKetoan1 = WSpace.OpenDatabase(psw, False, False, ";PWD=" + "1234")
'' DBKetoan1.Execute "ALTER TABLE TonKho ALTER COLUMN maso AutoNumber" 'TEXT(10)"
'
' 'DBKetoan1.Execute "ALTER TABLE TonKho ALTER COLUMN maso AutoNumber" 'TEXT(10)"
'
'
''  DBKetoan1.Execute "Insert into [C:\a2.mdb" & ";PWD=1234].license SELECT * From license"
''DBKetoan1.Execute "Insert into license select * from [C:\a2.mdb" & ";PWD=1234].license"
'Dim sql, st
''  sql = "update TonKho set "
''  sql = sql + " TonKho.MaSoKho = a.MaSoKho "
''  sql = sql + " from [C:\a2.mdb" & ";PWD=1234].TonKho as a where a.maso = TonKho.maso "
'
''DBKetoan.Execute "delete from Tonkho"
''sql = " Update [Tonkho] INNER JOIN [C:\a2.mdb;PWD=1234].TonKho a ON [Tonkho].maso = a.maso "
''        sql = sql + " SET [Tonkho].tien_0 = a.tien_12"
''        sql = sql + " ,[Tonkho].Luong_0 = a.Luong_12"
''        For i = 1 To 12
''        st = CStr(i)
''        sql = sql + ",[Tonkho].Luong_" + st + " = a.Luong_12, [Tonkho].Tien_" + st + " = a.Tien_12"
''      Next
'
'' update
'sql = "insert into vattu select * from [" + psw + ";PWD=" + pPSW + "].vattu where maso  not in (select maso from vattu) "
'DBKetoan.Execute sql
'sql = " insert into [Tonkho] select * from  [" + psw + ";PWD=" + pPSW + "].TonKho where maso not in (select maso from tonkho) "
'DBKetoan.Execute sql
'
'
'sql = " Update [Tonkho] b INNER JOIN [" + psw + ";PWD=" + pPSW + "].TonKho a ON b.maso = a.maso "
'        sql = sql + " SET b.tien_0 = a.tien_12"
'        sql = sql + " ,b.Luong_0 = a.Luong_12"
'        For i = 1 To 12
'        st = CStr(i)
'        sql = sql + ",b.Luong_" + st + " = a.Luong_12, b.Tien_" + st + " = a.Tien_12"
'      Next
'' insert
'
' DBKetoan.Execute sql
' MsgBox "thanhcong1"
'' 'sql = " insert into [Tonkho](tien_0,Luong_0) select tien_12,luong_12 from [C:\a2.mdb;PWD=1234].TonKho where maso not in (select maso from tonkho)"
'' sql = " insert into [Tonkho](tien_0,Luong_0) select tien_12 as tien_0,luong_12 as Luong_0 from [" + psw + ";PWD=1234].TonKho where maso not in (select maso from tonkho)"
'' DBKetoan.Execute sql
''  MsgBox "thanhcong2"
''
'
'
'  'connServer.Execute "Insert into [" & strClientDBPath & "\" & strClientDB & ";PWD=123456].MyTableName SELECT * From MyTableName"
'
''
''  Dim sql
''  sql = "select * from license "
''    Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
''
''        Do While Not rs_chungtu.EOF
''             MsgBox rs_chungtu!tenCty
''             rs_chungtu.MoveNext
''        Loop
'
''sql = " SELECT * INTO [";PWD=TestPass;DATABASE=C:\Docs\database2.mdb"].table2 FROM table1"
'
''DBKetoan. " INSERT INTO TonKho IN '" + psw + "' select * from TonKho"
''"INSERT INTO ThongSoCuoiKy IN 'D:\MyPham\Access\db2.mdb'" & Chr(10) _
''& "SELECT " & tbl.Name & ".*" & Chr(10) & _
''"FROM " & tbl.Name & "; "
'
'
''  ExecuteSQL5 "drop table SoDuKhachHang"
'' ExecuteSQL5 "drop table ThongSoCuoiKy"
'' ExecuteSQL5 "drop table TonKho"
'
'
''
''     Dim objAcc As Object
''    Dim objAcc2 As Object
''
''        Set objAcc = New Access.Application
''        objAcc.OpenCurrentDatabase ("C:\a1.mdb")
''
''
''        objAcc.DoCmd.CopyObject "C:\a2.mdb", , acTable, "license"
''
''        objAcc.CloseCurrentDatabase
''
''        Set objAcc = Nothing
'
'End Sub
'
'Private Sub Command2_Click()
'Dim psw As String, fn As String
'Dim sql
'  Dim rs_chungtu As Recordset
'  psw = frmMain.ChonTenTep("Chän tÖp d÷ liÖu", &H4&, "*.MDB", 1)
''Update [Tonkho] INNER JOIN [" + psw + ";PWD=1234].TonKho a ON [Tonkho].maso = a.maso "
''
'sql = "insert into khachhang select * from [" + psw + ";PWD=" + pPSW + "].khachhang where maso not in (select maso from khachhang)"
'DBKetoan.Execute sql
'sql = "insert into sodukhachhang select * from [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang where maso not in (select maso from SoDuKhachHang )"
'DBKetoan.Execute sql
'sql = " UPDATE SoDuKhachHang INNER JOIN [" + psw + ";PWD=" + pPSW + "].SoDuKhachHang a "
'sql = sql + " on SoDuKhachHang.maso = a.maso "
'sql = sql + " set SoDuKhachHang.DuNo_0  = a.DuNo_12 "
'sql = sql + ",SoDuKhachHang.DuCo_0 = a.DuCo_12 "
'sql = sql + ",SoDuKhachHang.DuNT_0 = a.DuNT_12 "
'DBKetoan.Execute sql
'MsgBox "thanh cong"
'
''  sql = "UPDATE SoDuKhachHang SET DuNo_0 = DuNo_12, DuCo_0 = DuCo_12, DuNT_0 = DuNT_12"
''    For i = 1 To 12
''        st = CStr(i)
''        sql = sql + ", No_" + st + " = 0, Co_" + st + " = 0, No_" + st + "_NT = 0, Co_" + st + "_NT = 0, DuNo_" _
''                    + st + " = DuNo_12, DuCo_" + st + " = DuCo_12, DuNT_" + st + " = DuNT_12"
''    Next
'End Sub
'
'Private Sub Command3_Click()
'sql = " UPDATE SoDuKhachHang "
'sql = sql + " set DuNo_0  = 0 "
'sql = sql + ",DuCo_0 = 0 "
'sql = sql + ",DuNT_0 = 0"
'DBKetoan.Execute sql
'MsgBox "xoa thanh cong"
'End Sub
'
'Private Sub Command4_Click()
'sql = " Update [Tonkho] b  "
'        sql = sql + " SET b.tien_0 = 0"
'        sql = sql + " ,b.Luong_0 = 0"
'        DBKetoan.Execute sql
'MsgBox "xoa thanh cong"
'End Sub
'
'Private Sub Command5_Click()
'Dim sql
'  sql = " UPDATE DISTINCTROW ThongSo LEFT JOIN ThongSoCuoiKy ON ThongSo.MaTS = ThongSoCuoiKy.MaTS "
'  sql = sql + " SET ThongSo.KH_NS = ThongSoCuoiKy.KH_NS,"
'  sql = sql + " ThongSo.KH_TBS = ThongSoCuoiKy.KH_TBS,"
'  sql = sql + " ThongSo.KH_CNK = ThongSoCuoiKy.KH_CNK,"
'  sql = sql + " ThongSo.KH_TD = ThongSoCuoiKy.KH_TD,"
'  sql = sql + " ThongSo.MaDTQL = ThongSoCuoiKy.MaDTQL,"
'  sql = sql + " ThongSo.MaDTSD = ThongSoCuoiKy.MaDTSD,"
'  sql = sql + " ThongSo.MaTTSD = ThongSoCuoiKy.MaTTSD "
'  sql = sql + "WHERE (((ThongSo.Thang)<12 And (ThongSo.Thang)>0));"
'     ExecuteSQL5
'End Sub
'
'Private Sub Form_Load()
'
'End Sub
