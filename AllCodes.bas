Attribute VB_Name = "AllCodes"
'What you see on the computer screen isn't what you will get when you print,
'the computer screen doesn't have the same resolution as a printer, therefore
'lines might appear to "merge" on the screen.
'The values in varBar1 are the available text in a given Barcode language to be printed
'The values in varBar2 are the Barcode equivalent of the text in varBar1
'sBar is the accumulated Barcode equivalents of the text to be printed
'The Barcode() Function will print one character of sBar at a time in a loop
'To add more Barcode types, just continue to build functions that make the appropriate sBar String
Option Explicit

'Public Const pBCode = 39
Public LO_XXXX As String
Public SL_XXXX As Double
Dim sBar As String, i0 As Integer, i1 As Integer

Public Function Code39(strCode As String)
Dim varBar1, varBar2
    varBar1 = Split("0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,-,., ,$,/,+,%,*", ",")
    varBar2 = Split("111221211,211211112,112211112,212211111,111221112,211221111,112221111,111211212,211211211,112211211,211112112,112112112,212112111,111122112,211122111,112122111,111112212,211112211,112112211,111122211,211111122,112111122,212111121,111121122,211121121,112121121,111111222,211111221,112111221,111121221,221111112,122111112,222111111,121121112,221121111,122121111,121111212,221111211,122111211,121212111,121211121,121112121,111212121,121121211", ",")
sBar = "121121211" & "1"
For i0 = 1 To Len(strCode)
    For i1 = 0 To UBound(varBar1)
        If Mid(strCode, i0, 1) = varBar1(i1) Then sBar = sBar & varBar2(i1) & "1"
    Next
Next
sBar = sBar & "121121211"
End Function

Public Function Code128(strCode As String)
Dim varBar1, varBar2
    varBar1 = Split(" <>!<>" & Chr(34) & "<>#<>$<>%<>&<>'<>(<>)<>*<>+<>,<>-<>.<>/<>0<>1<>2<>3<>4<>5<>6<>7<>8<>9<>:<>;<><<>=<>><>?<>@<>A<>B<>C<>D<>E<>F<>G<>H<>I<>J<>K<>L<>M<>N<>O<>P<>Q<>R<>S<>T<>U<>V<>W<>X<>Y<>Z<>[<>\<>]<>^<>_<>`<>a<>b<>c<>d<>e<>f<>g<>h<>i<>j<>k<>I<>m<>n<>o<>p<>q<>r<>s<>t<>u<>v<>w<>x<>y<>z<>{<>|<>}<>~<>DEL<>FNC 3<>FNC 2<>SHIFT<>CODE C<>FNC 4<>CODE A<>FNC 1<>Start A<>Start B<>Start C<>Stop", "<>")
    varBar2 = Split("212222,222122,222221,121223,121322,131222,122213,122312,132212,221213,221312,231212,112232,122132,122231,113222,123122,123221,223211,221132,221231,213212,223112,312131,311222,321122,321221,312212,322112,322211,212123,212321,232121,111323,131123,131321,112313,132113,132311,211313,231113,231311,112133,112331,132131,113123,113321,133121,313121,211331,231131,213113,213311,213131,311123,311321,331121,312113,312311,332111,314111,221411,431111,111224,111422,121124,121421,141122,141221,112214,112412,122114,122411,142112,142211,241211,221114,413111,241112,134111,111242,121142,121241,114212,124112,124211,411212,421112,421211,212141,214121,412121,111143,111341,131141,114113,114311,411113,411311,113141,114131,311141,411131,211412,211214,211232,2331112", ",")
Dim chksum As Single: chksum = 104
sBar = "211214"
For i0 = 1 To Len(strCode)
    For i1 = 0 To UBound(varBar1)
        If Mid(strCode, i0, 1) = varBar1(i1) Then
            sBar = sBar & varBar2(i1)
            chksum = chksum + (i1 * i0)
            Exit For
        End If
    Next
Next
sBar = sBar & varBar2(chksum - (Int(chksum / 103) * 103)) & "2331112"
End Function

Public Function BarCode(strCode As String, Pic As Object, barscale As Integer, barHeight As Single, StartX As Single, startY As Single)
Dim barWidth As Single, i0 As Integer, barStart As Single

'Select Case pBCode
'    Case 39:    strCode = UCase(strCode): Code39 strCode
'    Case 128:   Code128 strCode
'End Select

Code128 strCode

barStart = StartX
For i0 = 1 To Len(sBar)
    barWidth = Mid(sBar, i0, 1) * barscale
    If i0 Mod 2 > 0 Then Pic.Line (barStart, startY)-Step(barWidth, barHeight), vbBlack, BF
    barStart = barStart + IIf(i0 Mod 2 > 0, barWidth, barWidth * 1.3)
Next

End Function

Public Function PrintBarCode(vt As ClsVattu, sl As Integer) As Double
    Dim i As Integer, j As Integer
    
    i = sl \ 3 + IIf(sl Mod 3 <> 0, 1, 0)
    ExecuteSQL5 "DELETE * FROM BarCode"
    For j = 1 To i
        ExecuteSQL5 "INSERT INTO BarCode (MaSo,BarCode, Ten, GiaBan) VALUES (" + CStr(Lng_MaxValue("MaSo", "BarCode") + 1) + ",'" + vt.sohieu + "','" + vt.TenVattu + "'," + DoiDau(vt.GiaBan1) + ")"
    Next
    
    SetRptInfo
    frmMain.Rpt.ReportFileName = "BARCODE.RPT"
    frmMain.Rpt.PrinterName = "DATAMAX DMX I-4208"
    frmMain.Rpt.Destination = crptToWindow
    InBaoCaoRPT
    
End Function

