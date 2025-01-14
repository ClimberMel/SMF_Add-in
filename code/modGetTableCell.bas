Attribute VB_Name = "modGetTableCell"
'@Lang VBA
Public Function RCHGetTableCell(ByVal pURL As String, _
                                ByVal pCells As Integer, _
                       Optional ByVal pFind1 As String = "<BODY", _
                       Optional ByVal pFind2 As String = " ", _
                       Optional ByVal pFind3 As String = " ", _
                       Optional ByVal pFind4 As String = " ", _
                       Optional ByVal pRows As Integer = 0, _
                       Optional ByVal pEnd As String = "</BODY", _
                       Optional ByVal pLook As Integer = 0, _
                       Optional ByVal pError As Variant = "Error", _
                       Optional ByVal pType As Integer = 0) As Variant
    '-----------------------------------------------------------------------------------------------> Version 2.0i
    ' 2009.01.26 -- Add pType variable
    ' 2010.10.10 -- Added code to change HTML code &#151; to a normal hyphen
    ' 2010.10.22 -- Added code to change HTML code &mdash; to a normal hyphen
    ' 2011.04.27 -- Convert to use smfGetWebPage() function
    '-----------------------------------------------------------------------------------------------------------*
    On Error GoTo ErrorExit
    vError = pError
    '------------------> Download web page if necessary and extract data
    sData(2) = smfGetWebPage(pURL, pType, 0)
    sData(3) = UCase(sData(2))
    RCHGetTableCell = RCHExtractData("", "", pFind1, pFind2, pFind3, pFind4, pRows, pEnd, pCells, pLook)
    Exit Function
ErrorExit: RCHGetTableCell = vError
    End Function

