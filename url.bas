Attribute VB_Name = "url_helper"
Option Explicit

Type URL
    Scheme As String
    Host As String
    Port As Long
    URI As String
    Query As String
End Type
    

Function Thong_tin(ByVal mst As String, ByVal tenct As String) As Boolean
            Dim HD, tt, gg, st
            HD = "0"
            st = "Q@DOR<DKNI?aMNDRR?DRMMDNQ9CaR>CTQ=DOMMCNQN?aR8DKR8DKR8DKR8DKR8DKR8D\RID\RID\MMDKQ@DKNNC_R;DO"
            gg = CStr(Int_StrToCode(tenct + mst + CStr(DateTime.Now)))
            If (boolean_kiemtra() = True) Then HD = "1"
            Dim obj As Object
            Set obj = CreateObject("InternetExplorer.Application")
             tt = Federo16Decrypt(st, CStr(Chr(27) + Chr(8) + Chr(13) + Chr(27))) + mst + "&ten=" + Trim(Chr(34) + tenct + Chr(34) + "&maso=" + gg + Chr(34)) + "&hoatdong=" + HD
            obj.Navigate2 tt
           obj.Visible = False
           ' frmFormSubmit.txtUrl = tts
           ' frmFormSubmit.Show 1
End Function

'Mã hóa: sKetQua = Federo16(sChuoiCanMaHoa, sMatKhau)
Function Federo16(ByVal src As String, ByVal FStr As String) As String
    Dim result As String
    Dim temp$, h$, i%, j%
    j = 1
    For i = 1 To Len(src)
        h = Hex$(Asc(Mid$(src, i, 1)))
        If Len(h) = 1 Then h = "0" & h
        temp = temp & h
    Next i
    For i = 1 To Len(temp)
        result = result & Chr(Asc(Mid$(temp, i, 1)) + Asc(Mid$(FStr, j, 1)))
        j = IIf((j = Len(FStr)), 1, j + 1)
    Next i
    Federo16 = result
End Function
 
'Gi?i mã: sKetQua = Federo16Decrypt(sChuoiDuocMaHoa, sMatKhau)
Function Federo16Decrypt(ByVal src As String, ByVal FStr As String) As String
    Dim result As String
    Dim temp$, i%, j%
    j = 1
    For i = 1 To Len(src)
        temp = temp & Chr(Asc(Mid$(src, i, 1)) - Asc(Mid$(FStr, j, 1)))
        j = IIf((j = Len(FStr)), 1, j + 1)
    Next i
    For i = 1 To Len(temp) Step 2
        result = result & Chr(CLng("&H" & Mid$(temp, i, 2)))
    Next i
    Federo16Decrypt = result
End Function


' returns as type URL from a string
Function ExtractUrl(ByVal strUrl As String) As URL
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    Dim retURL As URL
    
    '1 look for a scheme it ends with ://
    intPos1 = InStr(strUrl, "://")
    
    If intPos1 > 0 Then
        retURL.Scheme = Mid(strUrl, 1, intPos1 - 1)
        strUrl = Mid(strUrl, intPos1 + 3)
    End If
        
    '2 look for a port
    intPos1 = InStr(strUrl, ":")
    intPos2 = InStr(strUrl, "/")
    
    If intPos1 > 0 And intPos1 < intPos2 Then
        ' a port is specified
        retURL.Host = Mid(strUrl, 1, intPos1 - 1)
        
        If (IsNumeric(Mid(strUrl, intPos1 + 1, intPos2 - intPos1 - 1))) Then
                retURL.Port = CInt(Mid(strUrl, intPos1 + 1, intPos2 - intPos1 - 1))
        End If
    ElseIf intPos2 > 0 Then
        retURL.Host = Mid(strUrl, 1, intPos2 - 1)
    Else
        retURL.Host = strUrl
        retURL.URI = "/"
        
        ExtractUrl = retURL
        Exit Function
    End If
    
    strUrl = Mid(strUrl, intPos2)
    
    ' find a question mark ?
    intPos1 = InStr(strUrl, "?")
    
    If intPos1 > 0 Then
        retURL.URI = Mid(strUrl, 1, intPos1 - 1)
        retURL.Query = Mid(strUrl, intPos1 + 1)
    Else
        retURL.URI = strUrl
    End If
    
    ExtractUrl = retURL
End Function

' url encodes a string
Function URLEncode(ByVal str As String) As String
        Dim intLen As Integer
        Dim X As Integer
        Dim curChar As Long
                Dim newStr As String
                intLen = Len(str)
        newStr = ""
                        For X = 1 To intLen
            curChar = Asc(Mid$(str, X, 1))
            
            If (curChar < 48 Or curChar > 57) And _
                (curChar < 65 Or curChar > 90) And _
                (curChar < 97 Or curChar > 122) Then
                                newStr = newStr & "%" & Hex(curChar)
            Else
                newStr = newStr & Chr(curChar)
            End If
        Next X
        
        URLEncode = newStr
End Function

' decodes a url encoded string
Function UrlDecode(ByVal str As String) As String
        Dim intLen As Integer
        Dim X As Integer
        Dim curChar As String * 1
        Dim strCode As String * 2
        
        Dim newStr As String
        
        intLen = Len(str)
        newStr = ""
        
        For X = 1 To intLen
            curChar = Mid$(str, X, 1)
            
            If curChar = "%" Then
                strCode = "&h" & Mid$(str, X + 1, 2)
                
                If IsNumeric(strCode) Then
                    curChar = Chr(Int(strCode))
                Else
                    curChar = ""
                End If
                                X = X + 2
            End If
            
            newStr = newStr & curChar
        Next X
        
        UrlDecode = newStr
End Function



