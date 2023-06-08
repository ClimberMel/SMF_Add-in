Attribute VB_Name = "modGetWebData"
Public Function RCHGetWebData(ByVal pURL As String, _
                         Optional ByVal pPos As Variant = 1, _
                         Optional ByVal pLen As Integer = 32767, _
                         Optional ByVal pOffset As Integer = 0, _
                         Optional ByVal pUseIE As Integer = 0) As Variant
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.02.16 -- Convert to use smfGetWebPage() function
    ' 2023-06-06 -- This is basically a pass through function to maintain old code that called it
    '            -- modules such as smfGetYahooPortfolio_.smfGetYahooPortfolioView call it
    '-----------------------------------------------------------------------------------------------------------*
    On Error GoTo ErrorExit
    '--------------------------------> Retrieve web page, if needed
    s1 = smfGetWebPage(pURL, pUseIE, 0)
    '--------------------------------> Preprocess web page data
    iPos = IIf(IsNumeric(pPos), pPos, InStr(s1, pPos) + pOffset)
    iLen = IIf(iPos + pLen <= Len(s1), pLen, Len(s1) - iPos + 1)
    RCHGetWebData = Mid(s1, iPos, iLen)
    Exit Function
ErrorExit: RCHGetWebData = "Error"
    End Function


