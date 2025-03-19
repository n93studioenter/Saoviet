Attribute VB_Name = "Module1"

Option Explicit

      Public Function Join(source() As String, Optional _
            sDelim As String = " ") As String
      Dim sOut As String, iC As Integer
      On Error GoTo errh:
          For iC = LBound(source) To UBound(source) - 1
              sOut = sOut & source(iC) & sDelim
          Next
          sOut = sOut & source(iC)
          Join = sOut
          Exit Function
errh:
          Err.Raise Err.Number
      End Function

      Public Function Split(ByVal sIn As String, Optional sDelim As _
            String, Optional nLimit As Long = -1, Optional bCompare As _
             VbCompareMethod = vbBinaryCompare) As Variant
          Dim sRead As String, sOut() As String, nC As Integer
          If sDelim = "" Then
              Split = sIn
          End If
          sRead = ReadUntil(sIn, sDelim, bCompare)
          Do
              ReDim Preserve sOut(nC)
              sOut(nC) = sRead
              nC = nC + 1
              If nLimit <> -1 And nC >= nLimit Then Exit Do
              sRead = ReadUntil(sIn, sDelim)
          Loop While sRead <> ""
          ReDim Preserve sOut(nC)
          sOut(nC) = sIn
          Split = sOut
      End Function

      Public Function ReadUntil(ByRef sIn As String, _
            sDelim As String, Optional bCompare As VbCompareMethod _
          = vbBinaryCompare) As String
          Dim nPos As String
          nPos = InStr(1, sIn, sDelim, bCompare)
          If nPos > 0 Then
              ReadUntil = Left(sIn, nPos - 1)
              sIn = Mid(sIn, nPos + Len(sDelim))
          End If
      End Function

      Public Function StrReverse(ByVal sIn As String) As String
          Dim nC As Integer, sOut As String
          For nC = Len(sIn) To 1 Step -1
          sOut = sOut & Mid(sIn, nC, 1)
          Next
          StrReverse = sOut
      End Function

      Public Function InStrRev(ByVal sIn As String, sFind As String, _
       Optional nStart As Long = 1, Optional bCompare As _
            VbCompareMethod = vbBinaryCompare) As Long
          Dim nPos As Long
          sIn = StrReverse(sIn)
          sFind = StrReverse(sFind)
          nPos = InStr(nStart, sIn, sFind, bCompare)
          If nPos = 0 Then
              InStrRev = 0
          Else
              InStrRev = Len(sIn) - nPos - Len(sFind) + 2
          End If
      End Function

      Public Function Replace(sIn As String, sFind As String, _
            sReplace As String, Optional nStart As Long = 1, _
            Optional nCount As Long = -1, Optional bCompare As _
            VbCompareMethod = vbBinaryCompare) As String

          Dim nC As Long, nPos As Integer, sOut As String
          sOut = sIn
          nPos = InStr(nStart, sOut, sFind, bCompare)
          If nPos = 0 Then GoTo EndFn:
          Do
              nC = nC + 1
              sOut = Left(sOut, nPos - 1) & sReplace & _
                 Mid(sOut, nPos + Len(sFind))
              If nCount <> -1 And nC >= nCount Then Exit Do
              nPos = InStr(nStart, sOut, sFind, bCompare)
          Loop While nPos > 0
EndFn:
          Replace = sOut
      End Function


