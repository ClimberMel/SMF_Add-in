Attribute VB_Name = "smfConvertData_"
Public Function smfConvertData(ByVal pData As String, _
                      Optional ByVal pConv As Integer = 0)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.02.16 -- Add routine
    ' 2012.01.21 -- Trim data coming in to function
    ' 2012.04.07 -- Add "Bill" and "Mill" suffixes conversions
    ' 2014.06.20 -- Add "Billion" and "Million" suffixes conversions
    ' 2014.08.15 -- Add "bil" and "mil" suffixes conversions
    ' 2016.05.18 -- Change CDec() to smfCDec() to ease transition between operating systems
    ' 2016.07.13 -- Add "K" suffix
    '-----------------------------------------------------------------------------------------------------------*
    s1 = Trim(pData)
    On Error GoTo ErrorExit
    If InStr(s1, "/") > 0 Then
    Else
        If s1 = "-" Then s1 = "0"
        If s1 = "--" Then s1 = "0"
        If s1 = "---" Then s1 = "0"
        If s1 = Chr(150) Then s1 = "0"
        If Left(s1, 1) = "$" Then s1 = Mid(s1, 2)
        If Left(s1, 1) = "(" And Right(s1, 1) = ")" Then s1 = "-" & Mid(s1, 2, Len(s1) - 2)
        s2 = s1
        
        Select Case True
            Case UCase(Right(s2, 1)) = "B": s2 = Left(s2, Len(s2) - 1): nMult = 1000000
            Case UCase(Right(s2, 1)) = "K": s2 = Left(s2, Len(s2) - 1): nMult = 1000
            Case UCase(Right(s2, 1)) = "M": s2 = Left(s2, Len(s2) - 1): nMult = 1000
            Case Right(s2, 1) = "%": s2 = Left(s2, Len(s2) - 1): nMult = 0.01
            Case Right(s2, 4) = " Bil": s2 = Left(s2, Len(s2) - 4): nMult = 1000000000
            Case Right(s2, 4) = " Mil": s2 = Left(s2, Len(s2) - 4): nMult = 1000000
            Case Right(s2, 4) = " bil": s2 = Left(s2, Len(s2) - 4): nMult = 1000000000
            Case Right(s2, 4) = " mil": s2 = Left(s2, Len(s2) - 4): nMult = 1000000
            Case Right(s2, 5) = " Bill": s2 = Left(s2, Len(s2) - 5): nMult = 1000000000
            Case Right(s2, 5) = " Mill": s2 = Left(s2, Len(s2) - 5): nMult = 1000000
            Case Right(s2, 8) = " Billion": s2 = Left(s2, Len(s2) - 8): nMult = 1000000000
            Case Right(s2, 8) = " Million": s2 = Left(s2, Len(s2) - 8): nMult = 1000000
            Case Else: nMult = 1
            End Select
       
       On Error Resume Next
       s1 = smfCDec(s2) * nMult
       On Error GoTo ErrorExit
       End If
    
ErrorExit:
    smfConvertData = s1
    
    End Function


