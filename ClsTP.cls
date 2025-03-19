VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsThanhPham"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MaSo As Long
Public MaCT As Long
Public SoHieu As String
Public thang As Integer
Public ngay As Date
Public MaKhoTP As Long
Public MaKhoNVL As Long
Public MaTK As Long
Public MaTP As Long
Public SoLuong As Double
Public MaTKCP As Long
Public MaTKNC As Long
Public MaTKSX As Long
Public MaTKKH As Long

Public Sub InitTPMaSo(ms As Long)
    Dim rs As Recordset
    
    If ms > 0 Then
        Set rs = DBKetoan.OpenRecordset("SELECT DISTINCTROW * FROM ThanhPham WHERE MaSo=" + CStr(ms), dbOpenSnapshot)
        If rs.RecordCount > 0 Then
            MaCT = rs!MaCT
            SoHieu = rs!SoHieu
            thang = rs!thang
            ngay = rs!ngay
            MaKhoTP = rs!MaKhoTP
            MaKhoNVL = rs!MaKhoNVL
            MaTK = rs!MaTK
            MaTP = rs!MaTP
            SoLuong = rs!SoLuong
            MaTKCP = rs!MaTKCP
            MaTKNC = rs!MaTKNC
            MaTKSX = rs!MaTKSX
            MaTKKH = rs!MaTKKH
            rs.Close
            Set rs = Nothing
        Else
            rs.Close
            Set rs = Nothing
            GoTo KT
        End If
    Else
KT:
        MaCT = 0
        SoHieu = "."
        thang = 0
        ngay = Date
        MaKhoTP = 0
        MaKhoNVL = 0
        MaTK = 0
        MaTP = 0
        SoLuong = 0
        MaTKCP = 0
        MaTKNC = 0
        MaTKSX = 0
        MaTKKH = 0
    End If
End Sub

Public Function SaveTP() As Integer
    SaveTP = ExecuteSQL5("INSERT INTO ThanhPham (MaSo,MaCT,Thang,Ngay,SoHieu,MaKhoTP,MaKhoNVL,MaTK,MaTP,SoLuong,MaTKCP,MaTKNC,MaTKSX,MaTKKH) VALUES (" + CStr(Lng_MaxValue("MaSo", "ThanhPham") + 1) + "," + CStr(MaCT) + "," _
        + CStr(thang) + ",#" + Format(ngay, Mask_DB) + "#,'" + SoHieu + "'," + CStr(MaKhoTP) + "," + CStr(MaKhoNVL) + "," + CStr(MaTK) + "," + CStr(MaTP) _
        + "," + DoiDau(SoLuong) + "," + CStr(MaTKCP) + "," + CStr(MaTKNC) + "," + CStr(MaTKSX) + "," + CStr(MaTKKH) + ")")
    MaSo = Lng_MaxValue("MaSo", "ThanhPham")
End Function

Public Function DeleteTP() As Integer
    DeleteTP = ExecuteSQL5("DELETE * FROM ThanhPham WHERE MaSo=" + CStr(MaSo))
End Function

