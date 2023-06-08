Attribute VB_Name = "smfGetYahooPortfolio_"
Option Explicit
Public Function smfGetYahooPortfolioView(ByVal pTickers As Variant, _
                         Optional ByVal pItems As Variant = "01020304050607080910111213141516171819202122232425262728293031323334353637383940414243444546474849505152535455565758596061626364656667686970717273747576777879808182838485868788899091", _
                         Optional ByVal pMultiple As String = "N", _
                         Optional ByVal pHeader As Integer = 0, _
                         Optional ByVal pDim1 As Integer = 0, _
                         Optional ByVal pDim2 As Integer = 0)
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to download columns from a portfolio view on Yahoo!
    '-----------------------------------------------------------------------------------------------------------*
    ' 2016.08.05 -- Created by Randy Harmelink (rharmelink@gmail.com)
    ' 2017.05.02 -- Obsoleted because portfolio table was replaced by JSON file
    ' 2017.06.21 -- Rewrite to extract fields from JSON file
    ' 2017.10.21 -- Rewrite to extract line by line instead of by field name
    ' 2017.11.02 -- Minor updates
    ' 2017.11.03 -- Fix to handle non-US currency combinations
    ' 2017.11.04 -- Add 52 additional fields
    ' 2017.11.04 -- Create list of request fields instead of asking for everything
    ' 2017.11.04 -- Maintain order of ticker symbol requests
    ' 2017.11.04 -- Add processing for EXCEL serial date/time values
    ' 2017.11.04 -- Divide percentage fields by 100, as needed
    ' 2017.11.04 -- Fix earnings dates
    ' 2017.11.06 -- Allow a ticker of "NONE" in first spot to bypass processing
    ' 2017.11.06 -- Adjust necessary date/time fields by GMT offset
    ' 2017.11.08 -- Fix errors on percentage adjustments when value returned is non-numeric
    ' 2017.11.09 -- Backed out percentage adjustments for fields 58, 61, 65, 68
    ' 2017.11.09 -- Fixed field list adjustment when only default fields are requested
    ' 2017.11.17 -- Allow a ticker symbol to be returned more than once
    ' 2023-05-09 -- Mel Pryor
    ' 2023-05-09 -- In sURL change v7 to v6.  Appears that the Export button in Yahoo Portfolio has changed that.
    ' 2023-06-07 -- Changed sUrl back to v7, modified calls to RCHGetWebData to use new pUseIE = 4
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get portfolio quotes for IBM and MMM:
    '
    '   =smfGetYahooPortfolioView("IBM,MMM")
    '   =smfGetYahooPortfolioView("IBM,MMM","00010203")
    '   =smfGetYahooPortfolioView("IBM,MMM","0001021011",,1)
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim aFieldNeed() As String:  aFieldNeed = Split("0,0,1,1,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1", ",")
    Dim aFieldName() As String:  aFieldName = Split("--,symbol,longName,shortName,exchange,fullExchangeName,market,marketState,sourceInterval,exchangeTimezoneName,exchangeTimezoneShortName,gmtOffSetMilliseconds,language,quoteType,quoteSourceName,regularMarketPrice,regularMarketTime,regularMarketChange,regularMarketOpen,regularMarketDayHigh,regularMarketDayLow,regularMarketVolume,bid,ask,sharesOutstanding,marketCap,averageDailyVolume3Month,targetPriceMean,revenue,priceToSales,trailingPE,epsTrailingTwelveMonths,exDividendDate,dividendsPerShare,dividend" & _
                                                    "Yield,dividendDate,dividendRate,trailingAnnualDividendYield,trailingAnnualDividendRate,earningsTimestamp,priceToBook,bookValue,epsForward,pegRatio,forwardPE,ebitda,shortRatio,shareFloat,currency,bidSize,askSize,regularMarketPreviousClose,regularMarketChangePercent,regularMarketDayRange,averageDailyVolume10Day,exchangeDataDelayedBy,fiftyDayAverage,fiftyDayAverageChange,fiftyDayAverageChangePercent,twoHundredDayAverage,twoHundredDayAverageChange,twoHundredDayAverageChangePercent,fiftyTwoWeekRange,fifty" & _
                                                    "TwoWeekLow,fiftyTwoWeekLowChange,fiftyTwoWeekLowChangePercent,fiftyTwoWeekHigh,fiftyTwoWeekHighChange,fiftyTwoWeekHighChangePercent,postMarketTime,postMarketPrice,postMarketChange,postMarketChangePercent,preMarketTime,preMarketPrice,preMarketChange,preMarketChangePercent,tradeable,regularMarketTime,regularMarketTime,exDividendDate,dividendDate,earningsTimestamp,postMarketTime,postMarketTime,preMarketTime,preMarketTime,regularMarketTime,postMarketTime,preMarketTime,earningsTimestampStart,earningsTimestampEnd", ",")
    Dim aHeading() As String:  aHeading = Split("--,Symbol,Long Name,Short Name,Exchange,Full Exchange Name,Market,Market State,Source Interval,Exchange Timezone Name,Exchange Timezone Short Name,GMT Offset Milliseconds,Language,Quote Type,Quote Source Name,Last Price,Last Traded (UNIX),Change,Open,High,Low,Volume,Bid,Ask,Shares Outstanding,Market Cap,Average 3M Volume,Mean Target Price,Revenue,P/S,P/E,EPS TTM,Ex-Dividend Date (UNIX),Dividends Per Share,Dividend Yield,Dividend Payment Date (UNIX),Forward Annual Div Rate,Trailing Annual Div Yield," & _
                                                "Trailing Annual Div Rate,Earnings Date (UNIX),Price/Book,Book Val,EPS Est Next Year,PEG Ratio (5yr expected),Forward P/E,EBITDA,Short Ratio,Float,Currency,Bid Size,Ask Size,Prev Close,% Chg,Day Range,Avg Vol (10 day),Data Delayed,50-DMA,50-DMA Chg,50-DMA Chg %,200-DMA,200-DMA Chg,200-DMA Chg %,52-Wk Range,52-Wk Low,52-Wk Low Chg,52-Wk Low Chg %,52-Wk High,52-Wk High Chg,52-Wk High Chg %,Post-Mkt Time (UNIX),Post-Mkt Price,Post-Mkt Chg,Post-Mkt % Chg,Pre-Mkt Time (UNIX),Pre-Mkt Price,Pre-Mkt Chg," & _
                                                "Pre-Mkt % Chg,Tradeable,Last Traded Date,Last Traded Time,Ex-Dividend Date,Dividend Payment Date,Earnings Date,Post-Mkt Date,Post-Mkt Time,Pre-Mkt Date,Pre-Mkt Time,Last Traded Date/Time,Post-Mkt Date/Time,Pre-Mkt Date/Time,Earnings Date Start,Earnings Date End", ",")
    
    Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, s1 As String
        
    '------------------> Determine size of array to return
    Dim iRows As Integer, iCols As Integer
    iRows = pDim1  ' Rows
    iCols = pDim2  ' Columns
    If pDim1 = 0 Or pDim2 = 0 Then
       If pDim1 = 0 Then iRows = 200   ' Old default
       If pDim2 = 0 Then iCols = 100   ' Old default
       On Error Resume Next
       iRows = Application.Caller.Rows.Count
       iCols = Application.Caller.Columns.Count
       On Error GoTo ErrorExit
       End If
  
    '------------------> Initialize return array
    ReDim vData(1 To iRows, 1 To iCols) As Variant
    For i1 = 1 To iRows
        For i2 = 1 To iCols
            vData(i1, i2) = "--"
            Next i2
        Next i1
    
    '------------------> Verify item and ticker and view parameters
    Dim oCell As Range, sItems As String, sTickers As String, aCols(1 To 99) As String, sFieldList As String
    Dim iFind As Integer, aTickers As Variant
    Select Case VarType(pItems)
        Case vbString
             sItems = LCase(Replace(pItems, " ", ""))
        Case Is >= 8192
             sItems = ""
             For Each oCell In pItems
                 If oCell.Value > "" Then sItems = sItems & Right(LCase(Format(oCell.Value, "00")), 2)
                 Next oCell
        Case Else
            smfGetYahooPortfolioView = "Invalid items parameter: " & pItems
            Exit Function
        End Select
    i1 = Len(sItems) / 2
    If i1 < iCols Then iCols = i1
    sFieldList = ","
    For i1 = 1 To iCols
        s1 = Mid(sItems & String$(68, "0"), 2 * i1 - 1, 2)
        Select Case s1
           Case "00" To "91"
                aCols(i1) = CInt(s1)
                If aFieldNeed(s1) <> 0 Then
                   iFind = InStr(sFieldList, "," & aFieldName(s1) & ",")
                   If iFind = 0 Then sFieldList = sFieldList & aFieldName(s1) & ","
                   End If
           Case Else: aCols(i1) = 0
           End Select
        Next i1
    If Len(sFieldList) > 2 Then sFieldList = Mid(sFieldList, 2, Len(sFieldList) - 2)  ' Remove leading and trailing comma
           
    Select Case VarType(pTickers)
        Case vbString
             sTickers = UCase(pTickers)
        Case Is >= 8192
             sTickers = ""
             For Each oCell In pTickers
                 sTickers = sTickers & IIf(oCell.Value <> "", UCase(oCell.Value), "XXXXX") & ","
                 Next oCell
             sTickers = Left(sTickers, Len(sTickers) - 1)
        Case Else
            smfGetYahooPortfolioView = "Invalid tickers parameter: " & pTickers
            Exit Function
        End Select
    aTickers = Split(sTickers, ",")
    
    '------------------> Create header if requested
    If pHeader = 1 Then
       For i1 = 1 To iCols
           vData(1, i1) = aHeading(aCols(i1))
           Next i1
       End If
    If aTickers(0) = "NONE" Then GoTo ErrorExit
    
    '------------------> Extract requested data items
    Dim iPtr As Long, iPos1 As Long, sData As String, sLine As String, sURL As String, v1 As Variant, vGMTOffset As Variant
    sURL = "https://query1.finance.yahoo.com/v7/finance/quote?fields=" & sFieldList & "&formatted=false&symbols=" & Replace(sTickers, ",XXXXX", "")
    iPtr = 1
    sData = RCHGetWebData(sURL, iPtr, 6000, , 4)
    iPos1 = InStr(2, sData, "result")
    iPtr = iPtr + iPos1 + 1
    sData = RCHGetWebData(sURL, iPtr, 6000, , 4)
    For i2 = 1 + pHeader To iRows
        iPos1 = InStr(2, sData, "{")
        If iPos1 = 0 Then Exit For
        iPtr = iPtr + iPos1 + 1
        sData = RCHGetWebData(sURL, iPtr, 6000, , 4)
        sLine = """" & smfStrExtr(sData & "}", "~", "}") & ","""
        s1 = smfStrExtr(sLine & ",", """symbol"":""", """")
        vGMTOffset = smfStrExtr(sLine & ",", """gmtOffSetMilliseconds"":", ",", 1) / 86400000
        For i3 = 0 To UBound(aTickers)
            If s1 = aTickers(i3) Then
               i4 = i3 + 1 + pHeader
               For i1 = 1 To iCols
                   v1 = smfStrExtr(sLine & ",", """" & aFieldName(aCols(i1)) & """:", ",""", 1)
                   If v1 = "" Then
                      vData(i4, i1) = "--"
                   Else
                      If Left(v1, 1) = """" Then v1 = smfStrExtr(v1 & """", """", """", 1)
                      Select Case 0 + aCols(i1)
                         Case 34, 52, 72, 76
                              On Error Resume Next
                              v1 = v1 / 100
                              On Error GoTo ErrorExit
                              vData(i4, i1) = v1
                         Case 80, 81, 82, 90, 91
                              vData(i4, i1) = Int(smfUnix2Date(0 + v1))
                         Case 78, 83, 85
                              vData(i4, i1) = Int(smfUnix2Date(0 + v1) + vGMTOffset)
                         Case 79, 84, 86
                              vData(i4, i1) = smfUnix2Date(0 + v1) + vGMTOffset - Int(smfUnix2Date(0 + v1) + vGMTOffset)
                         Case 87, 88, 89
                              vData(i4, i1) = smfUnix2Date(0 + v1) + vGMTOffset
                         Case Else
                              vData(i4, i1) = v1
                         End Select
                      End If ' v1
                   Next i1
               If pMultiple = "N" Then Exit For
               End If ' s1
            Next i3
        Next i2

ErrorExit:
    smfGetYahooPortfolioView = vData
    End Function





