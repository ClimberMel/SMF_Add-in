Attribute VB_Name = "modGetYahooHistory"
Public Function RCHGetYahooHistory(pTicker As String, _
                          Optional pStartYear As Integer = 1970, _
                          Optional pStartMonth As Integer = 1, _
                          Optional pStartDay As Integer = 1, _
                          Optional pEndYear As Integer = 2020, _
                          Optional pEndMonth As Integer = 12, _
                          Optional pEndDay As Integer = 31, _
                          Optional pPeriod As String = "d", _
                          Optional pItems As String = "DOHLCA", _
                          Optional pNames As Integer = 1, _
                          Optional pAdjust As Integer = 1, _
                          Optional pResort As Integer = 0, _
                          Optional pDim1 As Integer = 20000, _
                          Optional pDim2 As Integer = 7) ' As Variant()
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to create backward compatible RCHGetYahooHistory() function
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.05.25 -- Added for backward compatibility
    ' 2017.05.26 -- Check if pDim1 and pDim2 are overridden
    '-----------------------------------------------------------------------------------------------------------*
    ' 2023-01-22 -- Mel Pryor (climbermel@gmail.com)
    ' 2023-01-22 -- Created new version of RCHGetYahooHistory funtion in modGetYahooHistory
    '               Create routine to build URL as needed to then call RCHGetURLData
    '               Possibly add routine to parse returned table for data as requested in pItems
    ' 2023-01-24 -- Fixed pPeriod using default.  Added Process period section
    ' 2023-02-07 -- Restored previous version since issue was with smfGetYahooHistory
    '-----------------------------------------------------------------------------------------------------------*

    Dim sItems As String
    sItems = UCase(pItems)
    ' Adjusted Close is provided by Yahoo so A is now acceptable
    'Select Case True
    '   Case InStr(sItems, "C") > 0: sItems = Replace(sItems, "A", "")
    '   Case Else: sItems = Replace(sItems, "A", "C")
    '   End Select
    
    If pAdjust = 0 Then
       RCHGetYahooHistory = "Error: All data is now adjusted"
       Exit Function
       End If
       
    Dim iDim1 As Integer, iDim2 As Integer
    iDim1 = pDim1
    iDim2 = pDim2
    If pDim1 = 20000 And pDim2 = 7 Then
       On Error Resume Next
       iDim1 = Application.Caller.Rows.Count
       iDim2 = Application.Caller.Columns.Count
       End If
   
    RCHGetYahooHistory = testGetYahooHistory(pTicker, _
                                            pStartMonth & "/" & pStartDay & "/" & pStartYear, _
                                            pEndMonth & "/" & pEndDay & "/" & pEndYear, _
                                            pPeriod, sItems, pNames, pResort, iDim1, iDim2)

    End Function

Public Function RCHGetYahooHistory2(pTicker As String, _
                          Optional pStartYear As Integer = 0, _
                          Optional pStartMonth As Integer = 0, _
                          Optional pStartDay As Integer = 0, _
                          Optional pEndYear As Integer = 0, _
                          Optional pEndMonth As Integer = 0, _
                          Optional pEndDay As Integer = 0, _
                          Optional pPeriod As String = "d", _
                          Optional pItems As String = "DOHLCVA", _
                          Optional pNames As Integer = 1, _
                          Optional pAdjust As Integer = 0, _
                          Optional pResort As Integer = 0, _
                          Optional pDim1 As Integer = 0, _
                          Optional pDim2 As Integer = 0) ' As Variant()
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to download historical quotes from Yahoo!
    '-----------------------------------------------------------------------------------------------------------*
    ' 2005.01.18 -- Written by Randy Harmelink (rharmelink@gmail.com)
    ' 2005.08.14 -- Corrected row dimension to be based on Application.Caller instead of Selection
    ' 2005.08.21 -- Made pPeriod an optional parameter defaulting to daily quotes
    ' 2005.08.21 -- Added optional parameter pItems to allow selection of specific columns of data
    ' 2005.08.21 -- Added ability to add ticker symbol to each output row using the pItems parameter ("T")
    ' 2005.08.21 -- Added optional parameter pNames to allow removal of the first row of data with column names
    ' 2006.04.26 -- Added conversion of pass "pItems" variable into upper case
    ' 2006.04.26 -- Added translation of strings into amounts, where appropriate
    ' 2006.06.11 -- Remove edit of ticker length
    ' 2006.06.15 -- Add "v" (dividends only) possibility to pPeriod
    ' 2006.07.02 -- Fix to all weekly time period (was using "y" instead of "w" for some reason)
    ' 2006.07.10 -- Added ability to adjust data for splits and dividend (i.e. "pAdjust" variable)
    ' 2006.07.10 -- Added ability to resort data in ascending date order (i.e. "pResort" variable)
    ' 2006.07.10 -- Added conversion of date field into an EXCEL serial date number so it can be used as a date
    ' 2006.07.24 -- Changed date parameters to be optional for ease of use
    ' 2006.07.24 -- Added parameters "pDim1" and "pDim2" for VBA function call usage
    ' 2006.08.11 -- Removed adjustment of volume -- apparently Yahoo presents adjusted volume?
    ' 2006.09.11 -- Fix processing if there are more columns of data than can be returned (i.e. kDim2)
    ' 2006.10.06 -- Fix date comparison, using 2-digit months and days
    ' 2007.01.17 -- Change CCur() usage to CDec() because of precision issues
    ' 2007.01.19 -- Change defaults to set the early date because of changes to the Yahoo URL process
    ' 2007.01.22 -- Fix date defaults for weekly/monthly/dividend requests
    ' 2007.01.22 -- Change MsgBox errors to return the error message as the first data cell
    ' 2007.09.18 -- Modify pDim1/pDim2 processing
    ' 2010.02.15 -- Check to see if adjusted closing price is zero before calculating adjustment factor
    ' 2011.04.28 -- Change cDec() to smfConvertData()
    ' 2012.07.06 -- Handle doubling up on current date
    ' 2012.07.06 -- Allow returned dividend data to be resorted
    ' 2012.07.09 -- Fix handling of doubling up on current date
    ' 2017.04.17 -- Change protocol from "http://" to "https://"
    ' 2017.05.31 -- Add call to smfGetYahooHistoryCSV()
    '-----------------------------------------------------------------------------------------------------------*
    ' > Example of an invocation to get daily quotes for 2004 for IBM:
    '
    '   =RCHGetYahooHistory2("IBM",2004,1,1,2004,12,31,"d")
    '-----------------------------------------------------------------------------------------------------------*
    Dim sURL As String
    
    On Error GoTo ErrorExit
    
    '------------------> Null Return Item
    If pTicker = "None" Or pTicker = "" Then
       ReDim vData(1 To 1, 1 To 1) As Variant
       vData(1, 1) = "None"
       RCHGetYahooHistory2 = vData
       Exit Function
       End If
    
    '------------------> Determine size of array to return
    kDim1 = pDim1  ' Rows
    kDim2 = pDim2  ' Columns
    If pDim1 = 0 Or pDim2 = 0 Then
       On Error Resume Next
       kDim1 = Application.Caller.Rows.Count
       kDim2 = Application.Caller.Columns.Count
       On Error GoTo ErrorExit
       End If
  
    '------------------> Initialize return array
    ReDim vData(1 To kDim1, 1 To kDim2) As Variant
    For i1 = 1 To kDim1
        For i2 = 1 To kDim2
            vData(i1, i2) = ""
            Next i2
        Next i1
    
    '------------------> Edit parameters
    If pStartYear = 0 And _
       pStartMonth = 0 And _
       pStartDay = 0 And _
       pEndYear = 0 And _
       pEndMonth = 0 And _
       pEndDay = 0 Then
    Else
       If pStartYear < 1900 Or pStartYear > 2100 Or _
             pStartMonth < 1 Or pStartMonth > 12 Or _
             pStartDay < 1 Or pStartDay > 31 Or _
             pEndYear < 1900 Or pEndYear > 2100 Or _
             pEndMonth < 1 Or pEndMonth > 12 Or _
             pEndDay < 1 Or pEndDay > 31 Or _
                 pStartYear & Right("0" & pStartMonth, 2) & Right("0" & pStartDay, 2) > _
                 pEndYear & Right("0" & pEndMonth, 2) & Right("0" & pEndDay, 2) Then
          vData(1, 1) = "Something wrong with dates -- asked for " & _
                        pStartYear & "/" & pStartMonth & "/" & pStartDay & " thru " & _
                        pEndYear & "/" & pEndMonth & "/" & pEndDay
          GoTo ErrorExit
          End If
       End If
    Select Case pPeriod
       Case "d": iEndYear = Year(Date) - Int(kDim1 / 250) - 1
       Case "w": iEndYear = Year(Date) - Int(kDim1 / 50) - 1
       Case "m": iEndYear = Year(Date) - Int(kDim1 / 12) - 1
       Case "v": iEndYear = Year(Date) - Int(kDim1 / 4) - 1
       Case Else
            vData(1, 1) = "Invalid Period Requested: " & pPeriod
            GoTo ErrorExit
       End Select
       
    '------------------> Create URL and download historical quotes
    
    'sBase = "https://ichart.finance.yahoo.com/table.csv?s="
    'sURL = sBase & pTicker & _
    '       IIf(pStartMonth = 0, "&a=0", "&a=" & (pStartMonth - 1)) & _
    '       IIf(pStartDay = 0, "&b=1", "&b=" & pStartDay) & _
    '       IIf(pStartYear = 0, "&c=" & iEndYear, "&c=" & pStartYear) & _
    '       IIf(pEndMonth = 0, "", "&d=" & (pEndMonth - 1)) & _
    '       IIf(pEndDay = 0, "", "&e=" & pEndDay) & _
    '       IIf(pEndYear = 0, "", "&f=" & pEndYear) & _
    '       "&g=" & pPeriod & _
    '       "&ignore=.csv"
    ' sqData = RCHGetURLData(sURL)
    sURL = "https://query1.finance.yahoo.com/v7/finance/download/MMM?period1=1493610466&period2=1496202466&interval=1d&events=history&crumb="
    sqData = smfGetYahooHistoryCSVData(sURL)
    
    '------------------> Determine items needed
    pItems2 = UCase(pItems)
    If pPeriod = "v" Then
       If InStr(pItems2, "T") > 0 Then iTick = 1
       iDate = 1 + iTick
       iDiv = 2 + iTick
       iOpen = 0
       iHigh = 0
       iLow = 0
       iClos = 0
       iVol = 0
       iAdjC = 0
    Else
       iDiv = 0
       iTick = InStr(pItems2, "T")
       iDate = InStr(pItems2, "D")
       iOpen = InStr(pItems2, "O")
       iHigh = InStr(pItems2, "H")
       iLow = InStr(pItems2, "L")
       iClos = InStr(pItems2, "C")
       iVol = InStr(pItems2, "V")
       iAdjC = InStr(pItems2, "A")
       End If
    If iTick > kDim2 Then iTick = 0
    If iDate > kDim2 Then iDate = 0
    If iDiv > kDim2 Then iDiv = 0
    If iOpen > kDim2 Then iOpen = 0
    If iHigh > kDim2 Then iHigh = 0
    If iLow > kDim2 Then iLow = 0
    If iClos > kDim2 Then iClos = 0
    If iVol > kDim2 Then iVol = 0
    If iAdjC > kDim2 Then iAdjC = 0
    
    '------------------> Parse web quotes
    Dim sPrevDate As String
    vLine = Split(sqData, Chr(10))
    nLines = IIf(kDim1 - pNames < UBound(vLine) + pNames, kDim1 - pNames, UBound(vLine) + pNames)
    iRow = 1 - pNames - 1
    i1 = 1 - pNames - 1
    Do While iRow < nLines And i1 < UBound(vLine) + pNames
        i1 = i1 + 1
        If vLine(i1) = "" Then Exit Do
        vItem = Split(vLine(i1), ",")
        iRow = iRow + 1
        If iRow = 0 Then
           sAdjust = IIf(pAdjust = 1, "Adj. ", "")
           If iTick > 0 Then vData(iRow + pNames, iTick) = "Ticker"
           If iDate > 0 Then vData(iRow + pNames, iDate) = vItem(0)
           If iDiv > 0 Then vData(iRow + pNames, iDiv) = vItem(1)
           If iOpen > 0 Then vData(iRow + pNames, iOpen) = sAdjust & vItem(1)
           If iHigh > 0 Then vData(iRow + pNames, iHigh) = sAdjust & vItem(2)
           If iLow > 0 Then vData(iRow + pNames, iLow) = sAdjust & vItem(3)
           If iClos > 0 Then vData(iRow + pNames, iClos) = sAdjust & vItem(4)
           If iVol > 0 Then vData(iRow + pNames, iVol) = vItem(5)
           If iAdjC > 0 Then vData(iRow + pNames, iAdjC) = vItem(6)
        Else
           If sPrevDate = vItem(0) Then iRow = iRow - 1
           If iTick > 0 Then vData(iRow + pNames, iTick) = pTicker
           If iDate > 0 Then vData(iRow + pNames, iDate) = CDate(vItem(0))
           If iDiv > 0 Then vData(iRow + pNames, iDiv) = smfConvertData(vItem(1))
           sPrevDate = vItem(0)
           If pPeriod <> "v" Then
              If smfConvertData(vItem(4)) = 0 Or pAdjust <> 1 Then
                 nAdjust = 1
              Else
                 nAdjust = smfConvertData(vItem(6)) / smfConvertData(vItem(4))
                 End If
              If iOpen > 0 Then vData(iRow + pNames, iOpen) = smfConvertData(vItem(1)) * nAdjust
              If iHigh > 0 Then vData(iRow + pNames, iHigh) = smfConvertData(vItem(2)) * nAdjust
              If iLow > 0 Then vData(iRow + pNames, iLow) = smfConvertData(vItem(3)) * nAdjust
              If iClos > 0 Then vData(iRow + pNames, iClos) = smfConvertData(vItem(4)) * nAdjust
              If iVol > 0 Then vData(iRow + pNames, iVol) = smfConvertData(vItem(5))
              If iAdjC > 0 Then vData(iRow + pNames, iAdjC) = smfConvertData(vItem(6))
              End If
           End If
        Loop
    
    '------------------> Reverse the sort order of the data if requested
    If pResort = 1 Then
       Dim vTemp As Variant
       i1 = 1 + pNames
       i2 = iRow + pNames
       Do While i1 < i2
          For i3 = 1 To kDim2
              vTemp = vData(i1, i3)
              vData(i1, i3) = vData(i2, i3)
              vData(i2, i3) = vTemp
              Next i3
          i1 = i1 + 1
          i2 = i2 - 1
          Loop
       End If
    
ErrorExit:
    RCHGetYahooHistory2 = vData
    End Function
