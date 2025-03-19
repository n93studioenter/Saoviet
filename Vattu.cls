VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsVattu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Kieu vat tu
Public MaSo As Long             ' Ma vat tu
Public MaPhanLoai As Long       ' ma phan loai
Public sohieu As String         ' So hieu vat tu
Public TenVattu As String       ' Ten vat tu
Public DonVi As String          ' Don vi tinh
Public TonMin As Double         ' Luong ton toi thieu (=0 neu bo qua)
Public TonMax As Double         ' Luong ton toi da (=0 neu bo qua)
Public GiaHT As Double
Public Dvt2 As Integer
Public DonVi2 As String
Public TyLeQD As Double
Public GhiChu As String
Public VAT As Integer
Public GiaBan1 As Double
Public GiaBan2 As Double
Public GiaBan3 As Double
Public CK As Double
Public ThueNK As Double
'======================================================================================
' Thu tuc khoi tao object tu ma so vat tu
' mvt:  Ma so vat tu can tham chieu
'       =0 neu vat tu moi
'======================================================================================
Public Sub InitVattuMaSo(mvt As Long)
    Dim rs_vattu As Recordset
    
    If mvt > 0 Then
        Set rs_vattu = DBKetoan.OpenRecordset("SELECT Vattu.*" _
            & " FROM Vattu" _
            & " WHERE Vattu.MaSo=" + CStr(mvt), dbOpenSnapshot)
        If rs_vattu.RecordCount = 0 Then GoTo iv
        MaSo = rs_vattu!MaSo
        MaPhanLoai = rs_vattu!MaPhanLoai
        sohieu = rs_vattu!sohieu
        TenVattu = rs_vattu!TenVattu
        DonVi = rs_vattu!DonVi
        TonMin = rs_vattu!TonMin
        TonMax = rs_vattu!TonMax
        GiaHT = rs_vattu!GiaHT
        Dvt2 = rs_vattu!Dvt2
        DonVi2 = rs_vattu!DonVi2
        TyLeQD = rs_vattu!TyLeQD
        GhiChu = rs_vattu!GhiChu
        VAT = rs_vattu!VAT
        GiaBan1 = rs_vattu!GiaBan1
        GiaBan2 = rs_vattu!GiaBan2
        GiaBan3 = rs_vattu!GiaBan3
        CK = rs_vattu!CK
        'ThueNK = rs_vattu!ThueNK
        rs_vattu.Close
        Set rs_vattu = Nothing
    Else
iv:
        MaSo = 0
        MaPhanLoai = 0
        sohieu = ""
        TenVattu = ""
        DonVi = ""
        TonMin = 0
        TonMax = 0
        GiaHT = 0
        Dvt2 = 0
        DonVi2 = "..."
        TyLeQD = 0
        GhiChu = "..."
        VAT = 0
        GiaBan1 = 0
        GiaBan2 = 0
        GiaBan3 = 0
        CK = 0
        ThueNK = 0
    End If
End Sub
'==========================================================================
' Thu tuc khoi tao object tu so hieu vat tu
' mvt:  So hieu vat tu can tham chieu
'==========================================================================
Public Function InitVattuSohieu(shvt As String) As String
    Dim sql As String
    
    sql = "SELECT DISTINCTROW MaSo AS F1 FROM Vattu WHERE Sohieu='" + shvt + "'"
    
    InitVattuMaSo SelectSQL(sql)
    InitVattuSohieu = TenVattu
End Function
'==========================================================================
' Thu tuc dang ky vat tu
'==========================================================================
Public Function GhiVattu() As Integer
    Dim sql As String
        
    sql = "INSERT INTO Vattu (MaSo, MaPhanLoai,SoHieu,TenVattu,DonVi,TonMin,TonMax,GiaHT,Dvt2,DonVi2,TyLeQD,GhiChu,VAT,GiaBan1,GiaBan2,GiaBan3,CK,ThueNK) VALUES (" _
        + CStr(Lng_MaxValue("MaSo", "Vattu") + 1) + "," + CStr(MaPhanLoai) + ",'" + sohieu + "','" + TenVattu + "','" + DonVi + "'," + DoiDau(TonMin) + "," + DoiDau(TonMax) _
        + "," + DoiDau(RoundMoney(GiaHT)) + "," + CStr(Dvt2) + ",'" + DonVi2 + "'," + DoiDau(CDbl(TyLeQD)) + ",'" + GhiChu + "'," + CStr(VAT) + "," _
        + DoiDau(GiaBan1) + "," + DoiDau(GiaBan2) + "," + DoiDau(GiaBan3) + "," + DoiDau(CK) + "," + DoiDau(ThueNK) + ")"
    If ExecuteSQL5(sql) <> 0 Then
        GhiVattu = -1
    Else
        GhiVattu = 0
        MaSo = Lng_MaxValue("MaSo", "Vattu")
    End If
End Function
'==========================================================================
' Thu tuc tra ve TRUE neu vat tu da co phat sinh
'==========================================================================
Public Function VTCoDuLieu() As Boolean
    Dim sql As String
    
    sql = "SELECT DISTINCTROW Count(MaSo) AS F1 FROM Chungtu WHERE MaVattu=" + CStr(MaSo)
    VTCoDuLieu = SelectSQL(sql) > 0
    If Not VTCoDuLieu Then
        sql = "SELECT DISTINCTROW Count(MaSo) AS F1 FROM DinhMuc WHERE MaNVL=" + CStr(MaSo)
        VTCoDuLieu = SelectSQL(sql) > 0
    End If
End Function
'==========================================================================
' Cap nhat thong tin ve vat tu
'==========================================================================
Public Function SuaVT() As Integer
    Dim sql As String
    sql = "UPDATE Vattu SET MaPhanLoai=" + CStr(MaPhanLoai) + ",SoHieu='" + sohieu + "',TenVattu='" + TenVattu + "',DonVi='" + DonVi + "',TonMin=" _
        + DoiDau(TonMin) + ",TonMax=" + DoiDau(TonMax) + ",GiaHT=" + DoiDau(RoundMoney(GiaHT)) + ",Dvt2=" + CStr(Dvt2) + ",DonVi2='" _
        + DonVi2 + "',TyLeQD=" + DoiDau(CDbl(TyLeQD)) + ",GhiChu='" + GhiChu + "',VAT=" + CStr(VAT) + ",GiaBan1=" + DoiDau(GiaBan1) _
        + ",GiaBan2=" + DoiDau(GiaBan2) + ",GiaBan3=" + DoiDau(GiaBan3) + ",CK=" + DoiDau(CK) + ",ThueNK=" + DoiDau(ThueNK) + " WHERE MaSo=" + CStr(MaSo)
    If ExecuteSQL5(sql) <> 0 Then
        SuaVT = -1
    Else
        SuaVT = 0
    End If
End Function
'==========================================================================
' Thu tuc xoa mot vat tu
'==========================================================================
Public Function XoaVT() As Integer
    Dim i As Integer, sql As String
    
    If VTCoDuLieu() Then
        XoaVT = -1
    Else
'        sql = "DELETE * FROM DVTVattu WHERE MaVattu=" + CStr(MaSo)
'        ExecuteSQL5 sql
    For i = 1 To 12
            sql = sql + " or Luong_Nhap_" + CStr(i) + "<>0 or Tien_Nhap_" + CStr(i) + "<>0 "
        Next
          sql = "select count(*) as f1 FROM TonKho WHERE MaVattu=" + CStr(MaSo) + " AND (Luong_0<>0 or  Tien_0<>0) "
         
        If SelectSQL(sql) <= 0 Then
        sql = "DELETE FROM TonKho WHERE MaVattu=" + CStr(MaSo) + " AND (Luong_0=0 AND Tien_0=0) "
        For i = 1 To 12
            sql = sql + " AND Luong_Nhap_" + CStr(i) + "=0 AND Tien_Nhap_" + CStr(i) + "=0 "
        Next
        ExecuteSQL5 sql
        
        If ExecuteSQL5("DELETE FROM Vattu WHERE MaSo=" + CStr(MaSo)) = 0 Then
            XoaVT = 0
        Else
            XoaVT = -1
        End If
        Else
        XoaVT = -1
        End If
        
    End If
End Function

Public Function XoaDVT(m As Long) As Boolean
    XoaDVT = False
    If SelectSQL("SELECT MaCT AS F1 FROM ChungTu WHERE MaVattu=" + CStr(MaSo) + " AND DVT=" + CStr(m)) = 0 Then
        If ExecuteSQL5("DELETE * FROM DVTVattu WHERE MaSo=" + CStr(m)) = 0 Then XoaDVT = True
    End If
End Function

Public Sub KTraDVT2()
    Dim dv2 As Long
    
    dv2 = SelectSQL("SELECT Top 1 MaSo AS F1 FROM DVTVattu WHERE MaVattu=" + CStr(MaSo))
    ExecuteSQL5 "UPDATE Vattu SET Dvt2=" + IIf(dv2 > 0, "1", "0") + " WHERE MaSo=" + CStr(MaSo)
End Sub

Public Function SoChietKhau(sops As Double) As Double
    SoChietKhau = RoundMoney(sops * CK / 100)
End Function
