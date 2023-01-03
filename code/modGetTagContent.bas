Attribute VB_Name = "modGetTagContent"
Public Function smfGetTagContent(ByVal pURL As String, _
                                 ByVal pTag As String, _
                        Optional ByVal pTags As Integer = 1, _
                        Optional ByVal pFind1 As String = "<", _
                        Optional ByVal pFind2 As String = " ", _
                        Optional ByVal pFind3 As String = " ", _
                        Optional ByVal pFind4 As String = " ", _
                        Optional ByVal pConv As Integer = 0, _
                        Optional ByVal pError As Variant = "Error", _
                        Optional ByVal pType As Integer = 0, _
                        Optional ByVal pLen As Integer = 32767) As Variant
                        
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to return content from between a paired HTML tags
    '-----------------------------------------------------------------------------------------------> Version 2.0h
    ' 2009.06.14 -- Created function
    ' 2010.10.10 -- Added code to change HTML code &#151; to a normal hyphen
    ' 2010.10.22 -- Added code to change HTML code &mdash; to a normal hyphen
    ' 2011.02.16 -- Convert to use smfGetWebPage() function
    ' 2012.01.27 -- Added "pLen" parm to prevent excessive length of returned data
    ' 2014.04.10 -- Add call to smfStripHTML() for pConv=1
    ' 2017.11.11 -- Allow text string to be passed instead of a URL
    '-----------------------------------------------------------------------------------------------------------*
    ' > Example of an invocation:
    '
    '   =smfGetTagContent("http://finance.google.com/finance?client=ob&q=MUTF:GLRBX", "TD", 2, "Sharpe ratio")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    vError = pError
    
    '--------------------------------> Retrieve web page
    If Left(pURL, 4) = "http" Then sData1 = smfGetWebPage(pURL, pType, 0) Else sData1 = pURL
    sData2 = UCase(sData1)
    
    '--------------------------------> Find initial position on web page
    iPos1 = 0
    iPos1 = InStr(iPos1 + 1, sData2, UCase(pFind1))
    If iPos1 = 0 Then GoTo ErrorExit
    If pFind2 > " " Then
       iPos1 = InStr(iPos1 + 1, sData2, UCase(pFind2))
       If iPos1 = 0 Then GoTo ErrorExit
       End If
    If pFind3 > " " Then
       iPos1 = InStr(iPos1 + 1, sData2, UCase(pFind3))
       If iPos1 = 0 Then GoTo ErrorExit
       End If
    If pFind4 > " " Then
       iPos1 = InStr(iPos1 + 1, sData2, UCase(pFind4))
       If iPos1 = 0 Then GoTo ErrorExit
       End If
    
    '--------------------------------> Skip forward or backward number of HTML tags
    For i1 = 1 To Abs(pTags)
        If pTags > 0 Then
           iPos1 = InStr(iPos1 + 1, sData2, "<" & UCase(pTag))
        Else
           iPos1 = InStrRev(sData2, "<" & UCase(pTag), iPos1)
           End If
        If iPos1 = 0 Then GoTo ErrorExit
        Next i1
    
    '--------------------------------> Extract data between HTML tags
    iPos2 = InStr(iPos1, sData2, ">")
    iPos3 = InStr(iPos2, sData2, "</" & UCase(pTag))
    If UCase(pTag) = "TD" Then
       iPos4 = InStr(iPos2, sData2, "<TD")
       If iPos4 > 0 And (iPos3 = 0 Or iPos3 > iPos4) Then iPos3 = iPos4
       iPos4 = InStr(iPos2, sData2, "</TR")
       If iPos4 > 0 And (iPos3 = 0 Or iPos3 > iPos4) Then iPos3 = iPos4
       iPos4 = InStr(iPos2, sData2, "<TR")
       If iPos4 > 0 And (iPos3 = 0 Or iPos3 > iPos4) Then iPos3 = iPos4
       iPos4 = InStr(iPos2, sData2, "</TABLE")
       If iPos4 > 0 And (iPos3 = 0 Or iPos3 > iPos4) Then iPos3 = iPos4
       End If
    
    s1 = Trim(Mid(sData1, iPos2 + 1, iPos3 - iPos2 - 1))
    If pConv = 1 Then
       s1 = smfStripHTML(s1)
       s1 = smfConvertData(s1, 0)
       End If
    
    If Len(s1) > pLen Then s1 = Left(s1, pLen)
    smfGetTagContent = s1
    
    Exit Function

ErrorExit: smfGetTagContent = vError
   
   End Function


