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
   s1 = smfGetWebPage(sURL)
   smfGetYahooJSONField = smfJSONExtractField(s1, pField)
   
ExitFunction:
   End Function


Public Function smfGetYahooJSONData(ByVal pTicker As String, _
                                    ByVal pModule As String, _
                                    ByVal pField As String, _
                           Optional ByVal pPeriod As Integer = 1, _
                           Optional ByVal pProcess As String = "raw")
                    
   '-----------------------------------------------------------------------------------------------------------*
   ' User defined function to extract fields from Yahoo's new JSON feeds for financial statements data
   '-----------------------------------------------------------------------------------------------------------*
   ' 2017.04.19 -- Created function (rharmelink@gmail.com)
   ' 2017.04.21 -- Add multi-level pField parameter
   '-----------------------------------------------------------------------------------------------------------*
   ' > Example of an invocation:
   '
   '   =smfGetYahooJSONData("MMM", "cashFlowStatementHistory", "changeInCash")
   '   =smfGetYahooJSONData("MMM","financialData","targetMeanPrice")
   '-----------------------------------------------------------------------------------------------------------*
                                      
   Dim sURL As String, s1 As String, aSplit As Variant, i1 As Integer
   smfGetYahooJSONData = "Error"
   Select Case Left(pModule, 4)
      Case "http": sURL = Replace(pModule, "~~~~~", pTicker)
      Case Else: sURL = "https://query1.finance.yahoo.com/v10/finance/quoteSummary/" & pTicker & "?modules=" & pModule
      End Select
   aSplit = Split(pField, ".")
   If pPeriod < 1 Then s1 = "" Else s1 = RCHGetWebData(sURL, """" & aSplit(0) & """:")
   For i1 = 1 To UBound(aSplit)
       pField = aSplit(i1)
       s1 = """" & pField & """:" & smfStrExtr(s1, """" & pField & """:", "~")
       Next i1
   s1 = smfWord(s1, pPeriod + 1, """" & pField & """:")
   s1 = smfStrExtr(s1, "~", "}")
   If Left(s1, 4) = "null" Then s1 = ""
   Select Case LCase(pProcess)
      Case "": s1 = smfStrExtr(s1, """", """")
      Case "fmt": s1 = smfStrExtr(s1, """fmt"":""", """")
      Case "raw": s1 = smfStrExtr(s1, """raw"":", ",")
      Case Else: s1 = smfStrExtr(s1, """raw"":", ",")
      End Select
   If s1 = "" Then s1 = "Not found"
   smfGetYahooJSONData = smfConvertData(s1)
   End Function

