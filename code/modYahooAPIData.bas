Attribute VB_Name = "modYahooAPIData"
Option Explicit
Function smfYahooAPIData(pTicker As String, _
         Optional pItem As String = "LastTradePriceOnly", _
         Optional pFeed As String = "a", _
         Optional pOptionSymbol As String = "")
                                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to access Yahoo API feeds
    '-----------------------------------------------------------------------------------------------------------*
    ' 2012.12.11 -- Created by Randy Harmelink (rharmelink@gmail.com)
    '            -- Based on http://www.philadelphia-reflections.com/blog/2392.htm
    '            -- "a" feed is from the CSV data
    '            -- "b" feed is from yahoo.finance.quotes
    '            -- "c" and "d" feeds are from yahoo.finance.quant and quant2
    '            -- "e" feed is from yahoo.finance.stocks
    '            -- "f" feed is from yahoo.finance.options
    ' 2017.04.26 -- Change "http://" protocol to "https://"
    '-----------------------------------------------------------------------------------------------------------*
         
    Const kURLa = "https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20csv%20where%20url%3D'http%3A%2F%2Fdownload.finance.yahoo.com%2Fd%2Fquotes.csv%3Fs%3D~~~~~" _
                & "%26f%3Dsnll1d1t1cc1p2t7va2ibb6aa5pomwj5j6k4k5ers7r1qdyj1t8e7e8e9r6r7r5b4p6p5j4m3m7m8m4m5m6k1b3b2i5x%26e%3D.csv'%20and%20" _
                & "columns%3D'Symbol%2CName%2CLastTradeWithTime%2CLastTradePriceOnly%2CLastTradeDate%2CLastTradeTime%2CChange%20PercentChange%2CChange%2CChangeinPercent%2CTickerTrend%2CVolume%2CAverageDailyVolume%2CMoreInfo%2CBid%2CBidSize%2CAsk%2CAskSize%2CPreviousClose%2COpen%2CDayRange%2CFiftyTwoWeekRange%2CChangeFromFiftyTwoWeekLow%2CPercentChangeFromFiftyTwoWeekLow%2CChangeFromFiftyTwoWeekHigh%2CPercentChangeFromFiftyTwoWeekHigh%2CEarningsPerShare%2CPE%20Ratio%2CShortRatio%2CDividendPayDate%2CExDividendDate%2CDividendPerShare%2CDividend%20Yield%2CMarketCapitalization%2COneYearTargetPrice%2CEPS%20Est%20Current%20Yr%2CEPS%20Est%20Next%20Year%2CEPS%20Est%20Next%20Quarter%2CPrice%20per%20EPS%20Est%20Current%20Yr%2CPrice%20per%20EPS%20Est%20Next%20Yr%2CPEG%20Ra" _
                & "tio%2CBook%20Value%2CPrice%20to%20Book%2CPrice%20to%20Sales%2CEBITDA%2CFiftyDayMovingAverage%2CChangeFromFiftyDayMovingAverage%2CPercentChangeFromFiftyDayMovingAverage%2CTwoHundredDayMovingAverage%2CChangeFromTwoHundredDayMovingAverage%2CPercentChangeFromTwoHundredDayMovingAverage%2CLastTrade%20(Real-time)%20with%20Time%2CBid%20(Real-time)%2CAsk%20(Real-time)%2COrderBook%20(Real-time)%2CStockExchange'"
    Const kURLb = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quotes%20where%20symbol%20in%20%28%22~~~~~%22%29&diagnostics=false&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys"
    Const kURLc = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quant%20where%20symbol%20in%20(%22~~~~~%22)&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys"
    Const kURLd = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quant2%20where%20symbol%20in%20(%22~~~~~%22)&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys"
    Const kURLe = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.stocks%20where%20symbol%20in%20(%22~~~~~%22)&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys"
    Const kURLf = "http://query.yahooapis.com/v1/public/yql?q=SELECT%20*%20FROM%20yahoo.finance.options%20WHERE%20symbol=%22~~~~~%22%20AND%20expiration%20in%20%28SELECT%20contract%20FROM%20yahoo.finance.option_contracts%20WHERE%20symbol=%22~~~~~%22%29&env=http%3A%2F%2Fdatatables.org%2Falltables.env"
         
    Dim sURL As String
    Select Case UCase(pFeed)
       Case "A": sURL = Replace(kURLa, "~~~~~", pTicker)
       Case "B": sURL = Replace(kURLb, "~~~~~", pTicker)
       Case "C": sURL = Replace(kURLc, "~~~~~", pTicker)
       Case "D": sURL = Replace(kURLd, "~~~~~", pTicker)
       Case "E": sURL = Replace(kURLe, "~~~~~", pTicker)
       Case "F": sURL = Replace(kURLf, "~~~~~", pTicker)
       End Select
         
    smfYahooAPIData = smfConvertData(smfGetTagContent(sURL, pItem & ">", 1, pOptionSymbol))

End Function
