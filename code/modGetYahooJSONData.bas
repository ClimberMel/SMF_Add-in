Attribute VB_Name = "modGetYahooJSONData"
Public Function smfGetYahooJSONField(ByVal pTicker As String, _
                                     ByVal pModule As String, _
                                     ByVal pField As String, _
                            Optional ByVal pStart As Double = 0)
                    
   '-----------------------------------------------------------------------------------------------------------*
   ' User defined function to extract fields from Yahoo's new JSON feeds for financial statements data
   '-----------------------------------------------------------------------------------------------------------*
   ' 2017.04.27 -- Created function (rharmelink@gmail.com)
   ' 2017.05.16 -- Add "portfolioView" processing
   ' 2017.05.16 -- Add "headlineNews" processing
   ' 2017.05.19 -- Add "barChartMM" processing
   ' 2017.07.05 -- Allow URL to be passed for a module
   ' 2017.10.12 -- Add error exit
   ' 2017.10.21 -- Fix "portfolioView" URL
   ' 2018.12.13 -- Change RCHGetWebData() to smfGetWebPage()
   ' 2023-05-29 -- Will not work with 64bit Excel due to ScriptEngine only having 32bit available.
   ' 2023-08-29 -- Use new function if 64bit Excel.
   '-----------------------------------------------------------------------------------------------------------*
   ' > Example of an invocation:
   '
   '   =smfGetYahooJSONField("MMM", "cashFlowStatementHistory", "quoteSummary.result.0.cashflowStatementHistory.cashflowStatements.0.changeInCash.raw")
   '   =smfGetYahooJSONField("MMM", "financialData", "quoteSummary.result.0.financialData.targetMeanPrice.raw")
   '-----------------------------------------------------------------------------------------------------------*
                                    
   smfGetYahooJSONField = "Not Found"
   On Error GoTo ExitFunction
                                    
   Dim sURL As String, s1 As String
   Select Case True
      Case Left(pModule, 4) = "http"
           sURL = pModule
      Case pModule = "barChartMM"
           sURL = aConstants(1)
      Case pModule = "portfolioView"
           sURL = "https://query1.finance.yahoo.com/v7/finance/quote?fields=symbol,longName,shortName,regularMarketPrice,regularMarketTime,regularMarketChange," & _
                  "regularMarketDayHigh,regularMarketDayLow,regularMarketPrice,regularMarketOpen,regularMarketVolume,averageDailyVolume3Month,marketCap,bid,ask," & _
                  "dividendYield,dividendsPerShare,exDividendDate,trailingPE,priceToSales,targetPriceMean&formatted=false&symbols=" & pTicker
      Case pModule = "headlineNews"
           sURL = "https://query1.finance.yahoo.com/v2/finance/news?count=20&symbols=" & pTicker & "&start=" & pStart
      Case Else
           sURL = "https://query1.finance.yahoo.com/v10/finance/quoteSummary/" & pTicker & "?modules=" & pModule
      End Select
  ' s1 = smfGetWebPage(sURL)
   s1 = smfGetWebPage(sURL, 4)                   ' Get url using crumb
   
   #If Win64 Then
        ' .. Excel x64
        smfGetYahooJSONField = smfJSONExtractField_x64(s1, pField)
   #Else
        ' .. Excel x32
        smfGetYahooJSONField = smfJSONExtractField(s1, pField)
   #End If
   
ExitFunction:
   End Function

Public Function smfGetYahooJSONData(ByVal pTicker As String, _
                                    ByVal pModule As String, _
                                    ByVal pField As String, _
                           Optional ByVal pPeriod As Integer = 1, _
                           Optional ByVal pProcess As String = "raw", _
                           Optional ByVal pEndDate As Integer = 0)
                    
   '-----------------------------------------------------------------------------------------------------------*
   ' User defined function to extract fields from Yahoo's new JSON feeds for financial statements data
   '-----------------------------------------------------------------------------------------------------------*
   ' 2017.04.19 -- Created function (rharmelink@gmail.com)
   ' 2017.04.21 -- Add multi-level pField parameter
   ' 2023-06-05 -- Add "num format type to deal with unquoted numbers in the JSON data
   '            -- see https://github.com/ClimberMel/SMF_Add-in/issues/37 for details
   ' 2023-06-14 -- BS#01 Add pEndDate parameter and code to deal with date "ranges" in the JSON data
   ' 2023-07-16 -- BS#02 Call RCHGetWebData() with pUseIE=4 - JSON url now using "crumb"
   ' 2024-06-20 -- Add processing for new Yahoo Balance Sheet & Income statement "time-series" url.
   '-----------------------------------------------------------------------------------------------------------*
   ' > Example of an invocation:
   '
   '   =smfGetYahooJSONData("MMM", "cashFlowStatementHistory", "changeInCash")
   '   =smfGetYahooJSONData("MMM","financialData","targetMeanPrice")
   '   =smfGetYahooJSONData("AAPL", "quoteType", "gmtOffSetMilliseconds", , "num")
   '   =smfGetYahooJSONData("AAPL", "calendarEvents", "earningsDate", ,"fmt", 1)
   '   =smfGetYahooJSONData("IBM","smfYBSAnnual","annualTotalAssets.reportedValue",4)
'-----------------------------------------------------------------------------------------------------------*
                                      
   Dim sURL As String, s1 As String, aSplit As Variant, i1 As Integer, s2 As String
   smfGetYahooJSONData = "Error"
   Select Case Left(pModule, 4)
      Case "smfY": sURL = smfBuildYahooTimeSeriesURL(pTicker, pModule)      ' pseudo module?
      Case "http": sURL = Replace(pModule, "~~~~~", pTicker)
      Case Else: sURL = "https://query1.finance.yahoo.com/v10/finance/quoteSummary/" & pTicker & "?modules=" & pModule
      End Select
      
   aSplit = Split(pField, ".")
   
'   If pPeriod < 1 Then s1 = "" Else s1 = RCHGetWebData(sURL, """" & aSplit(0) & """:")
   If pPeriod < 1 Then s1 = "" Else s1 = RCHGetWebData(sURL, """" & aSplit(0) & """:", , , 4)    ' BS#02 - get using crumb
   
   For i1 = 1 To UBound(aSplit)
       pField = aSplit(i1)
       s1 = """" & pField & """:" & smfStrExtr(s1, """" & pField & """:", "~")
       Next i1
   s1 = smfWord(s1, pPeriod + 1, """" & pField & """:")
   
   '--------- Special handling for "Earnings Date End" field --- BS#01 ---------------------
    If pEndDate = 1 Then
       s2 = smfWord(s1, 2, "},{")
       If s2 <> "" Then s1 = s2
       End If
   '--------- Special handling for "Earnings Date End" field -------------------------------

   s1 = smfStrExtr(s1, "~", "}")
   If Left(s1, 4) = "null" Then s1 = ""
   
   Select Case LCase(pProcess)
      Case "": s1 = smfStrExtr(s1, """", """")
      Case "fmt": s1 = smfStrExtr(s1, """fmt"":""", """")
      Case "raw": s1 = smfStrExtr(s1, """raw"":", ",")
      Case "num": s1 = smfStrExtr(s1, "~", ",""")
      Case Else: s1 = smfStrExtr(s1, """raw"":", ",")
      End Select
      
   If s1 = "" Then s1 = "Not found"
   smfGetYahooJSONData = smfConvertData(s1)
   End Function
   
Public Function smfBuildYahooTimeSeriesURL(ByVal pTicker As String, _
                                         ByVal pModule As String)
   '----------------------------------------------------------------------------------------------------------------*
   ' UDF to insert fields into Yahoo's new ".../ws/fundamentals-timeSeries..." URL for financial statements data
   '----------------------------------------------------------------------------------------------------------------*
   ' 2024.06.20 -- Created function
   '
   '----------------------------------------------------------------------------------------------------------------*
   ' Note: The "&period2=" field in the time-series url is the ending UNIX UTC time passed to get the JSON
   '       data. The value is the current UNIX UTC time but modified to always use 59 min and 59 seconds at the end of
   '       the hour. This was determined by how Yahoo calculated the "&period2=" value for its own processing.
   '       For example assuming PDT (UTC + 7):
   '            if the current local time is 06/12/24 13:22:44, the modified UTC is 06/12/24 20:59:59
   '            if the current local time is 06/12/24 13:44:13, the modified UTC is 06/12/24 20:59:59
   '            if the current local time is 06/12/24 14:01:31, the modified UTC is 06/12/24 21:59:59
   '            if the current local time is 06/12/24 17:43:14, the modified UTC is 06/13/24 00:59:59
   '
   '       Testing shows that JSON data will still be returned using any future UNIX time for "&period2=" but it
   '       may tip off Yahoo to non-sanctioned calls.
   '----------------------------------------------------------------------------------------------------------------*
   
    Dim sURL As String, sFieldlist As String, iPtr As Integer
    Dim sUtcNow As Variant, sPeriod2 As Variant

    smfBuildYahooTimeSeriesURL = "Error"
    On Error GoTo ExitFunction
    
    sUtcNow = JsonConverter.ConvertToUtc(Now())                  ' Get UTC date/time
    sPeriod2 = Int(sUtcNow) + TimeSerial(Hour(sUtcNow), 59, 59)  ' modify it
    sPeriod2 = smfDate2Unix(sPeriod2)                            ' convert to Unix date/time
   
   '-----------------------------------------------------------------------------------*
   '   Get corresponding JSON fields from "Constants" array. (smf-elements-22.txt file.)
    
    sFieldlist = ""
    For iPtr = 1 To UBound(aConstants)
        If smfStrExtr(aConstants(iPtr), "~", "/") = pModule Then
            sFieldlist = smfStrExtr(aConstants(iPtr), pModule & "/", "~")       ' Get field list for "pseudo" module
            Exit For
        End If
    Next iPtr
   
   If sFieldlist = "" Then
        sURL = "Invalid module" & pModule
   Else
        sURL = "https://query1.finance.yahoo.com/ws/fundamentals-timeseries/v1/finance/timeseries/" & pTicker & _
               "?merge=false&padTimeSeries=true" & "&period1=493590046&period2=" & sPeriod2 & _
               "&type=" & sFieldlist & "&lang=en-US&region=US"
   End If
   
   smfBuildYahooTimeSeriesURL = sURL
       
ExitFunction:
   End Function

