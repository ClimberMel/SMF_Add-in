Attribute VB_Name = "modExtractData"
'@Lang VBA
Public Function RCHExtractData(ByVal pSource As String, _
                                ByVal pElement As String, _
                                ByVal pFind1 As String, _
                                ByVal pFind2 As String, _
                                ByVal pFind3 As String, _
                                ByVal pFind4 As String, _
                                ByVal pRows As Integer, _
                                ByVal pEnd As String, _
                                ByVal pCells As Integer, _
                                ByVal pLook As Integer) As Variant
                                
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.04.28 -- Change cDec() to smfConvertData()
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    '--------------------------------> Find initial position on web page
    iPos1 = 0
    iPos1 = InStr(iPos1 + 1, sData(3), UCase(pFind1))
    If iPos1 = 0 Then GoTo ErrorExit
    If pFind2 > " " Then
       iPos1 = InStr(iPos1 + 1, sData(3), UCase(pFind2))
       If iPos1 = 0 Then GoTo ErrorExit
       End If
    If pFind3 > " " Then
       iPos1 = InStr(iPos1 + 1, sData(3), UCase(pFind3))
       If iPos1 = 0 Then GoTo ErrorExit
       End If
    If pFind4 > " " Then
       aSplit = Split(UCase(pFind4), "|")
       For i1 = 0 To UBound(aSplit, 1)
           iPos2 = InStr(iPos1 + 1, sData(3), aSplit(i1))
           If iPos2 > 0 Then Exit For
           If i1 = UBound(aSplit, 1) Then GoTo ErrorExit
           Next i1
       iPos1 = iPos2
       End If
    '--------------------------------> Skip backward/forward the number of specified table rows
    Select Case True
       Case pRows > 0
            iPos2 = InStr(iPos1, sData(3), UCase(pEnd))
            For i1 = 1 To pRows
                iPos1 = InStr(iPos1 + 1, sData(3), "<TR")
                iPos3 = InStr(iPos1, sData(3), "</TR")
                If iPos3 > iPos2 Then
                   RCHExtractData = vError
                   Exit Function
                   End If
                Next i1
       Case pRows < 0
            iPos2 = InStrRev(sData(3), UCase(IIf(pEnd = "</BODY", "<BODY", pEnd)), iPos1)
            For i1 = 1 To Abs(pRows)
                iPos1 = InStrRev(sData(3), "<TR", iPos1 - 1)
                If iPos1 < iPos2 Then
                   RCHExtractData = vError
                   Exit Function
                   End If
                Next i1
            End Select
    '--------------------------------> Skip forward or backward the number of specified table cells
    iPos2 = iPos1
    iRowBeg = InStrRev(sData(3), "<TR", iPos2)
    If pCells = 0 Then
       iRowEnd = InStr(iPos2, sData(3), "</TR")
       iLoop = 1
    ElseIf pCells < 0 Then
       iRowEnd = InStr(iPos2, sData(3), "</TR")
       iPos2 = iRowEnd
       iLoop = -pCells
    Else
       iLoop = pCells
       If pEnd <> "" Then
          iRowEnd = InStr(iPos2, sData(3), "</TR")
       Else
          iRowEnd = Len(sData(3))
          End If
       End If
    For i1 = 1 To iLoop + pLook
        If pCells > 0 Then
           iPos2 = InStr(iPos2, sData(3), "<TD")
        Else
           iPos2 = InStrRev(sData(3), "<TD", iPos2)
           End If
        If iPos2 = 0 Or iPos2 < iRowBeg Or iPos2 > iRowEnd Then GoTo ErrorExit
        iPos2 = InStr(iPos2, sData(3), ">")
        If i1 >= iLoop Then
           iPos3 = InStr(iPos2, sData(3), "</TD")
           '-------------------------> Extract cell contents and strip out HTML tags
           s1 = Trim(Mid(sData(2), iPos2 + 1, iPos3 - iPos2 - 1))
           s1 = Replace(Trim(s1), "<br>", Chr(10))
           Do
               iPos4 = InStr(s1, "<")
               If iPos4 = 0 Then Exit Do
               iPos5 = InStr(iPos4, s1, ">")
               If iPos5 = 0 Then Exit Do
               s1 = IIf(iPos4 = 1, "", Left(s1, iPos4 - 1)) & Trim(Mid(s1 & " ", iPos5 + 1, 99999))
               Loop
           If s1 <> "" Then Exit For
           End If
        If pCells < 0 Then
           iPos2 = InStrRev(sData(3), "<TD", iPos2) - 1
           End If
        Next i1
    RCHExtractData = smfConvertData(s1)
    Exit Function
ErrorExit: RCHExtractData = vError
    End Function


