Attribute VB_Name = "NguyenTe"
Option Explicit

'======================================================================================
' Hµm tr¶ vÒ gi¸ trÞ VND tõ l­îng nguyªn tÖ theo tû gi¸ mÆc ®Þnh
'======================================================================================
Public Function TyGiaNT(mnt As Long) As Double
    Dim sql As String
    
    If mnt < 0 Then
        TyGiaNT = 0
    Else
        If mnt = 0 Then
            sql = "SELECT TyGia AS F1 FROM NguyenTe WHERE KyHieu = 'USD'"
            
            TyGiaNT = SelectSQL(sql)
        Else
            sql = "SELECT TyGia AS F1 FROM NguyenTe WHERE MaSo = " + CStr(mnt)
            
            TyGiaNT = SelectSQL(sql)
        End If
    End If
End Function

Public Sub CapNhatTyGia(mnt As Long, tygia As Double)
    If tygia <> 0 Then ExecuteSQL5 "UPDATE NguyenTe SET TyGia = " + DoiDau(tygia) + " WHERE MaSo = " + CStr(mnt)
End Sub

Public Function TyGiaBQ(shtk As String, mnt As Long, ngay As Date) As Double
    Dim mtk As Long, duno As Double, duco As Double, dunt As Double
    
    mtk = SelectSQL("SELECT MaSo AS F1 FROM HethongTK WHERE SoHieu='" + shtk + "' AND MaNT=" + CStr(mnt))
    If mtk > 0 Then
        SoDuTKNgay mtk, ngay, duno, duco, dunt
        If dunt <> 0 Then TyGiaBQ = Abs((duno - duco) / dunt)
    End If
End Function

Public Function TyGiaDK() As Double
    TyGiaDK = SelectSQL("SELECT TyGia AS F1 FROM License")
End Function
