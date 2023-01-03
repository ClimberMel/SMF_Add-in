Attribute VB_Name = "modMStarReturns"
Option Explicit
Public Function smfGetMorningstarHistReturns(pTicker As String, _
                              Optional ByVal pFreq As String = "m", _
                              Optional ByVal pYears As Integer = 5, _
                              Optional ByVal pDim1 As Integer = 0, _
                              Optional ByVal pError As Variant = "Error")
                              
    '-------------------------------------------------------------------------------------------------------*
    ' 2014.06.01 -- Created function to get historical returns from Morningstar, monthly ("m") or qtrly ("q")
    '-------------------------------------------------------------------------------------------------------*

    Dim iEnd As Integer, iYear As Integer, iPeriod As Integer, iCount As Integer
    Dim iYearDate As Variant
    Dim kDim1 As Integer, kDim2 As Integer, i1 As Integer, i2 As Integer
    Dim sURL As String
    
    On Error GoTo ErrorExit
    
    '------------------> Null Return Item
    If pTicker = "None" Or pTicker = "" Then
       ReDim vData(1 To 1, 1 To 1) As Variant
       vData(1, 1) = "None"
       smfGetMorningstarHistReturns = vData
       Exit Function
       End If
    
    '------------------> Determine size of array to return
    kDim1 = pDim1  ' Rows
    kDim2 = 2      ' Columns
    If pDim1 = 0 Then
       On Error Resume Next
       kDim1 = Application.Caller.Rows.Count
       kDim2 = Application.Caller.Columns.Count
       On Error GoTo ErrorExit
       End If
    If kDim2 = 1 Then kDim2 = 2
  
    '------------------> Initialize return array
    ReDim vData(1 To kDim1, 1 To kDim2) As Variant
    For i1 = 1 To kDim1
        For i2 = 1 To kDim2
            vData(i1, i2) = ""
            Next i2
        Next i1
    
    '------------------> Create initial values
    sURL = "http://performance.morningstar.com/Performance/fund/historical-returns.action?ndec=3&y=" & pYears & "&freq=" & pFreq & "&t=" & pTicker
    iEnd = IIf(pFreq = "m", 11, 3)
    iYear = 0
    iCount = 0
    iPeriod = smfStrExtr(smfGetTagContent(sURL, "tr", -1, "year0_" & pFreq), "_" & pFreq, """") - 1
    
    '------------------> Extract each period until there are no more
    Do While True
       
       iPeriod = IIf(iPeriod = iEnd, 0, iPeriod + 1)
       If iPeriod = 0 Then iYear = iYear + 1
       iYearDate = RCHGetTableCell(sURL, -1, "year" & iYear)
       If iYearDate = "Error" Then Exit Do
       
       iCount = iCount + 1
       vData(iCount, 1) = DateSerial(iYearDate, 13 - IIf(pFreq = "m", 1, 3) * iPeriod, 0)
       vData(iCount, 2) = RCHGetTableCell(sURL, 1, "year" & iYear & "_" & pFreq & iPeriod) / 100
       
       If iCount = kDim1 Then Exit Do
       
       Loop
    
    smfGetMorningstarHistReturns = vData
    Exit Function

ErrorExit:
    smfGetMorningstarHistReturns = pError
    End Function


