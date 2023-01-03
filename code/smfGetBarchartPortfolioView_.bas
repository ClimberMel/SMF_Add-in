Attribute VB_Name = "smfGetBarchartPortfolioView_"
Option Explicit
Public Function smfGetBarchartPortfolioView(ByVal pTickers As Variant, _
                         Optional ByVal pItems As Variant = "001016009010011006007008012013021022", _
                         Optional ByVal pMultiple As String = "N", _
                         Optional ByVal pHeader As Integer = 0, _
                         Optional ByVal pDim1 As Integer = 0, _
                         Optional ByVal pDim2 As Integer = 0)
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to download columns from a portfolio view on Barchart!
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.11.06 -- Created by Randy Harmelink (rharmelink@gmail.com)
    ' 2017.11.17 -- Allow a ticker symbol to be returned more than once
    ' 2018.06.14 -- Allow a "&list=" item to be pass as pTickers
    ' 2018.06.14 -- Add fields "totalOptionsVolume,percentCallOptions,percentPutOptions"
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get portfolio quotes for IBM and MMM:
    '
    '   =smfGetBarchartPortfolioView("IBM,MMM")
    '   =smfGetBarchartPortfolioView("IBM,MMM","001002")
    '   =smfGetBarchartPortfolioView("IBM,MMM","001002",,1)
    '-----------------------------------------------------------------------------------------------------------*
    
    'On Error GoTo ErrorExit
    Dim aUseWhich() As String:  aUseWhich = Split("--,1,1,1,1,1,1,1,1,1,1,2,1,1,1,2,1,1,1,1,1,1,1,1,1,1,1,1,2,2,2,2,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,1,1", ",")
    Dim aFieldName() As String:  aFieldName = Split("--,symbol,symbolName,symbolShortName,contractName,exchange,openPrice,highPrice,lowPrice,lastPrice,priceChange,percentChange,volume,previousPrice,industry,contractExpirationDate,tradeTime,hasOptions,standardDeviation,weightedAlpha,openInterest,highPrice1y,lowPrice1y,pivotPoint,resistanceLevel1,resistanceLevel2,supportLevel1,supportLevel2,marketCap,sharesOutstanding,annualSales,annualNetIncome,beta,percentInsider,percentInstitutional,growth1y,growth3y,growth5y,revenueGrowth5y,earningsGrowth5y,dividen" & _
                                                    "dGrowth5y,earnings,epsDate,nextEarningsDate,epsAnnual,epsGrowthQuarter,epsGrowthYear,dividendRate,dividendYield,dividend,dividendDate,dividendExDate,paymentDate,dividendPayout,split,splitDate,peRatioTrailing,peRatioForward,pegRatio,returnOnEquity,returnOnAssets,profitMargin,debtEquity,priceSales,priceCashFlow,priceBook,bookValue,interestCoverage,movingAverage1m,movingAverage65d,movingAverage130d,movingAverage9m,movingAverage260d,movingAverageYtd,movingAverage5d,movingAverage20d,movingAverage50d," & _
                                                    "movingAverage100d,movingAverage200d,averageVolume1m,averageVolume3m,averageVolume6m,averageVolume9m,averageVolume1y,averageVolumeYtd,averageVolume5d,averageVolume20d,averageVolume50d,averageVolume100d,averageVolume200d,rawStochastic9d,rawStochastic14d,rawStochastic20d,rawStochastic50d,rawStochastic100d,stochasticK9d,stochasticK14d,stochasticK20d,stochasticK50d,stochasticK100d,stochasticD9d,stochasticD14d,stochasticD20d,stochasticD50d,stochasticD100d,averageTrueRange9d,averageTrueRange14d,average" & _
                                                    "TrueRange20d,averageTrueRange50d,averageTrueRange100d,relativeStrength9d,relativeStrength14d,relativeStrength20d,relativeStrength50d,relativeStrength100d,percentR9d,percentR14d,percentR20d,percentR50d,percentR100d,historicVolatility9d,historicVolatility14d,historicVolatility20d,historicVolatility50d,historicVolatility100d,macdOscillator9d,macdOscillator14d,macdOscillator20d,macdOscillator50d,macdOscillator100d,priceChange1m,priceChange3m,priceChange6m,priceChange9m,priceChange1y,priceChangeYtd,p" & _
                                                    "riceChange5d,priceChange20d,priceChange50d,priceChange100d,priceChange200d,percentChange1m,percentChange3m,percentChange6m,percentChange9m,percentChange1y,percentChangeYtd,percentChange5d,percentChange20d,percentChange50d,percentChange100d,percentChange200d,highPrice5d,highDate5d,lowPrice5d,lowDate5d,highHits5d,highPercent5d,lowHits5d,lowPercent5d,highPrice1m,highDate1m,lowPrice1m,lowDate1m,highHits1m,highPercent1m,lowHits1m,lowPercent1m,highPrice3m,highDate3m,lowPrice3m,lowDate3m,highHits3m,hig" & _
                                                    "hPercent3m,lowHits3m,lowPercent3m,highPrice6m,highDate6m,lowPrice6m,lowDate6m,highHits6m,highPercent6m,lowHits6m,lowPercent6m,highDate1y,lowDate1y,highHits1y,highPercent1y,lowHits1y,lowPercent1y,highPriceYtd,highDateYtd,lowPriceYtd,lowDateYtd,highHitsYtd,highPercentYtd,lowHitsYtd,lowPercentYtd,opinion,opinionStrength,opinionDirection,opinionPrevious,opinionLastWeek,opinionLastMonth,opinionShortTerm,opinionMediumTerm,opinionLongTerm,averageRecommendation,trendSpotterSignal,trendSpotterStrength,tr" & _
                                                    "endSpotterDirection,average7dSignal,average7dStrength,average7dDirection,movingAverage10to8dSignal,movingAverage10to8dStrength,movingAverage10to8dDirection,movingAverage20dSignal,movingAverage20dStrength,movingAverage20dDirection,macd20to50dSignal,macd20to50dStrength,macd20to50dDirection,bollingerBands20dSignal,bollingerBands20dStrength,bollingerBands20dDirection,commodityChannel40dSignal,commodityChannel40dStrength,commodityChannel40dDirection,movingAverage50dSignal,movingAverage50dStrength,mov" & _
                                                    "ingAverage50dDirection,macd20to100dSignal,macd20to100dStrength,macd20to100dDirection,parabolicTimePrice50dSignal,parabolicTimePrice50dStrength,parabolicTimePrice50dDirection,commodityChannel60dSignal,commodityChannel60dStrength,commodityChannel60dDirection,movingAverage100dSignal,movingAverage100dStrength,movingAverage100dDirection,macd50to100dSignal,macd50to100dStrength,macd50to100dDirection,opinion,opinionStrength,opinionDirection,opinionPrevious,opinionLastWeek,opinionLastMonth,opinionShortTe" & _
                                                    "rm,opinionMediumTerm,opinionLongTerm,trendSpotterSignal,trendSpotterStrength,trendSpotterDirection,average7dSignal,average7dStrength,average7dDirection,movingAverage10to8dSignal,movingAverage10to8dStrength,movingAverage10to8dDirection,movingAverage20dSignal,movingAverage20dStrength,movingAverage20dDirection,macd20to50dSignal,macd20to50dStrength,macd20to50dDirection,bollingerBands20dSignal,bollingerBands20dStrength,bollingerBands20dDirection,commodityChannel40dSignal,commodityChannel40dStrength,c" & _
                                                    "ommodityChannel40dDirection,movingAverage50dSignal,movingAverage50dStrength,movingAverage50dDirection,macd20to100dSignal,macd20to100dStrength,macd20to100dDirection,parabolicTimePrice50dSignal,parabolicTimePrice50dStrength,parabolicTimePrice50dDirection,commodityChannel60dSignal,commodityChannel60dStrength,commodityChannel60dDirection,movingAverage100dSignal,movingAverage100dStrength,movingAverage100dDirection,macd50to100dSignal,macd50to100dStrength,macd50to100dDirection,totalOptionsVolume,percent" & _
                                                    "CallOptions,percentPutOptions", ",")
    Dim aHeading() As String:  aHeading = Split("--,Symbol,Name,Short Name,Contract Name,Exchange,Open,High,Low,Last,Change $,Change %,Volume,Previous,Industry,Expiration Date,Trade Time,Has Options,Std Dev,Wtd Alpha,Open Interest,52W High,52W Low,Pivot Point,1st Resistance,2nd Resistance,1st Support,2nd Support,Market Cap,Shares Outstanding,Annual Sales,Net Income,Beta,% Insider,% Institutional,1Y Return%,3Y Return%,5Y Return%,5Y Rev%,5Y Earn%,5Y Div%,Earnings,Earnings Date,Next Earnings Date,Earnings ttm,EPS Growth Prv Qtr,EPS Growth Prv Yr,Ann" & _
                                                    "ual Dividend,Dividend Yield,Dividend,Last Div Date,Ex-Div Date,Div Pymt Date,Div Payout%,Split Amt,Split Date,P/E ttm,Fwd P/E,PEG,ROE%,ROA%,Profit%,Debt/Equity,P/S,P/CF,P/B,Book Value,Int Coverage,1M SMA,3M SMA,6M SMA,9M SMA,12M SMA,YTD SMA,5D SMA,20D SMA,50D SMA,100D SMA,200D SMA,1M Avg Vol,3M Avg Vol,6M Avg Vol,9M Avg Vol,52W Avg Vol,YTD Avg Vol,5D Avg Vol,20D Avg Vol,50D Avg Vol,100D Avg Vol,200D Avg Vol,9D Stoch R,14D Stoch R,20D Stoch R,50D Stoch R,100D Stoch R,9D Stoch %K,14D Stoch %K,20D " & _
                                                    "Stoch %K,50D Stoch %K,100D Stoch %K,9D Stoch %D,14D Stoch %D,20D Stoch %D,50D Stoch %D,100D Stoch %D,9D Range,14D Range,20D Range,50D Range,100D Range,9D Rel Str,14D Rel Str,20D Rel Str,50D Rel Str,100D Rel Str,9D %R,14D %R,20D %R,50D %R,100D %R,9D Hist Volatility,14D Hist Volatility,20D Hist Volatility,50D Hist Volatility,100D Hist Volatility,9D MACD,14D MACD,20D MACD,50D MACD,100D MACD,1M Chg,3M Chg,6M Chg,9M Chg,52W Chg,YTD Chg,5D Chg,20D Chg,50D Chg,100D Chg,200D Chg,1M %Chg,3M %Chg,6M %Chg," & _
                                                    "9M %Chg,52W %Chg,YTD %Chg,5D %Chg,20D %Chg,50D %Chg,100D %Chg,200D %Chg,5D High,5D High Date,5D Low,5D Low Date,5D #Highs,5D %/High,5D #Lows,5D %/Low,1M High,1M High Date,1M Low,1M Low Date,1M #Highs,1M %/High,1M #Lows,1M %/Low,3M High,3M High Date,3M Low,3M Low Date,3M #Highs,3M %/High,3M #Lows,3M %/Low,6M High,6M High Date,6M Low,6M Low Date,6M #Highs,6M %/High,6M #Lows,6M %/Low,52W High Date,52W Low Date,52W #Highs,52W %/High,52W #Lows,52W %/Low,YTD High,YTD High Date,YTD Low,YTD Low Date,YTD" & _
                                                    " #Highs,YTD %/High,YTD #Lows,YTD %/Low,Opinion,Opin Strength,Opin Direction,Opin Previous,Opin Last Week,Opin Last Month,Opin Short Term,Opin Medium Term,Opin Long Term,Avg Recommend,Trendspotter Signal,Trendspotter Strength,Trendspotter Direction,7D ADX Signal,7D ADX Strength,7D ADX Direction,10-8D HiLo MA,10-8D HiLo MA Strength,10-8D HiLo MA Direction,20D MA Signal,20D MA Strength,20D MA Direction,20-50D MACD,20-50D MACD Strength,20-50D MACD Direction,20D BBands Signal,20D BBands Strength,20D " & _
                                                    "BBands Direction,40D CCI Signal,40D CCI Strength,40D CCI Direction,50D MA Signal,50D MA Strength,50D MA Direction,20-100D MACD Signal,20-100D MACD Strength,20-100D MACD Direction,50D Parabolic Signal,50D Parabolic Strength,50D Parabolic Direction,60D CCI Signal,60D CCI Strength,60D CCI Direction,100D MA Signal,100D MA Strength,100D MA Direction,50-100D MACD Signal,50-100D MACD Strength,50-100D MACD Direction,Opinion Score,Opin Strength Score,Opin Direction Score,Opin Previous Score,Opin Last Wee" & _
                                                    "k Score,Opin Last Month Score,Opin Short Term Score,Opin Medium Term Score,Opin Long Term Score,Trendspotter Signal Score,Trendspotter Strength Score,Trendspotter Direction Score,7D ADX Signal Score,7D ADX Strength Score,7D ADX Direction Score,10-8D HiLo MA Score,10-8D HiLo MA Strength Score,10-8D HiLo MA Direction Score,20D MA Signal Score,20D MA Strength Score,20D MA Direction Score,20-50D MACD Score,20-50D MACD Strength Score,20-50D MACD Direction Score,20D BBands Signal Score,20D BBands Stre" & _
                                                    "ngth Score,20D BBands Direction Score,40D CCI Signal Score,40D CCI Strength Score,40D CCI Direction Score,50D MA Signal Score,50D MA Strength Score,50D MA Direction Score,20-100D MACD Signal Score,20-100D MACD Strength Score,20-100D MACD Direction Score,50D Parabolic Signal Score,50D Parabolic Strength Score,50D Parabolic Direction Score,60D CCI Signal Score,60D CCI Strength Score,60D CCI Direction Score,100D MA Signal Score,100D MA Strength Score,100D MA Direction Score,50-100D MACD Signal Scor" & _
                                                    "e,50-100D MACD Strength Score,50-100D MACD Direction Score,Options Volume,% Calls,% Puts", ",")
    
    Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, s1 As String, bList As Boolean
        
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
    Dim oCell As Range, sItems As String, sTickers As String, aCols(1 To 300) As String, sFieldList As String
    Dim iFind As Integer, aTickers As Variant
    Select Case VarType(pItems)
        Case vbString
             sItems = LCase(Replace(pItems, " ", ""))
        Case Is >= 8192
             sItems = ""
             For Each oCell In pItems
                 If oCell.Value > "" Then sItems = sItems & Right(LCase(Format(oCell.Value, "000")), 3)
                 Next oCell
        Case Else
            smfGetBarchartPortfolioView = "Invalid items parameter: " & pItems
            Exit Function
        End Select
    i1 = Len(sItems) / 3
    If i1 < iCols Then iCols = i1
    sFieldList = ",symbol,"
    For i1 = 1 To iCols
        s1 = Mid(sItems & String$(1000, "0"), 3 * i1 - 2, 3)   ' Make sure string is long enough
        Select Case s1
           Case "000" To "297"
                aCols(i1) = CInt(s1)
                iFind = InStr(sFieldList, "," & aFieldName(s1) & ",")
                If iFind = 0 Then sFieldList = sFieldList & aFieldName(s1) & ","
           Case Else: aCols(i1) = 0
           End Select
        Next i1
    sFieldList = Mid(sFieldList, 2, Len(sFieldList) - 2)  ' Remove leading and trailing comma
           
    Select Case VarType(pTickers)
        Case vbString
             If pTickers = "None" Then GoTo ErrorExit
             sTickers = pTickers
        Case Is >= 8192
             sTickers = ""
             For Each oCell In pTickers
                 sTickers = sTickers & IIf(oCell.Value <> "", oCell.Value, "XXXXX") & ","
                 Next oCell
             sTickers = Left(sTickers, Len(sTickers) - 1)
        Case Else
            smfGetBarchartPortfolioView = "Invalid tickers parameter: " & pTickers
            Exit Function
        End Select
    If Left(sTickers, 6) = "&list=" Then
       sTickers = smfWord(sTickers & ",", 1, ",")
       bList = True
    Else
       sTickers = UCase(sTickers)
       bList = False
       End If
    aTickers = Split(sTickers, ",")
        
    '------------------> Create header if requested
    If pHeader = 1 Then
       For i1 = 1 To iCols
           vData(1, i1) = aHeading(aCols(i1))
           Next i1
       End If
    If aTickers(0) = "NONE" Then GoTo ErrorExit
    
    '------------------> Extract requested data items
    Dim iPtr As Long, iPos1 As Long, sData As String, sLine1 As String, sLine2 As String, sURL As String, v1 As Variant
    sURL = "https://core-api.barchart.com/v1/quotes/get?symbols=" & Replace(sTickers, ",XXXXX", "") & "&fields=" & sFieldList & "&raw=1"
    iPtr = 1
    sData = RCHGetWebData(sURL, iPtr, 30000)
    iPos1 = InStr(2, sData, """data"":")
    iPtr = iPtr + iPos1 + 1
    sData = RCHGetWebData(sURL, iPtr, 30000)
    For i2 = 1 + pHeader To iRows
        iPos1 = InStr(2, sData, "{")
        If iPos1 = 0 Then Exit For
        iPtr = iPtr + iPos1 + 1
        sData = RCHGetWebData(sURL, iPtr, 30000)
        'sLine = """" & smfStrExtr(sData & "}", "~", "}") & ","""
        sLine1 = """" & smfStrExtr(sData & "}", "~", """raw"":") & ","""
        sLine2 = """" & smfStrExtr(sData & "}", """raw"":{", "}") & ","""
        iPos1 = InStr(2, sData, "}}")
        If iPos1 = 0 Then Exit For
        iPtr = iPtr + iPos1 + 1
        sData = RCHGetWebData(sURL, iPtr, 30000)
        s1 = smfStrExtr(sLine1 & ",", """symbol"":""", """")
        For i3 = 0 To UBound(aTickers)
            If s1 = aTickers(i3) Or bList Then
               If bList Then i4 = i2 Else i4 = i3 + 1 + pHeader
               For i1 = 1 To iCols
                   If aFieldName(aCols(i1)) <> "--" Then
                      v1 = smfStrExtr(IIf(aUseWhich(aCols(i1)) = 1, sLine1, sLine2) & ",", """" & aFieldName(aCols(i1)) & """:", ",""")
                      v1 = smfConvertData(Replace(v1, "\/", "/"))
                      If v1 = "" Then
                         vData(i4, i1) = "--"
                      Else
                         If Left(v1, 1) = """" Then v1 = smfStrExtr(v1 & """", """", """", 1)
                         Select Case 0 + aCols(i1)
                            Case 16, 42, 43, 50, 51, 52, 55, 153, 155, 161, 163, 169, 171, 177, 179, 184, 185, 191, 193
                                 vData(i4, i1) = v1
                                 On Error Resume Next
                                 vData(i4, i1) = DateValue(v1)
                                 On Error GoTo ErrorExit
                            Case Else
                                 vData(i4, i1) = v1
                            End Select
                         End If ' v1
                      End If ' aFieldName
                   Next i1
               If pMultiple = "N" Then Exit For
               End If ' s1
            Next i3
        Next i2

ErrorExit:
    smfGetBarchartPortfolioView = vData
    End Function









