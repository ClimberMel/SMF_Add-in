Attribute VB_Name = "modGetElementNumber"
Const kVersion = "2.2.2023.01.22"                   ' Version number of add-in
    
Const kElements = 20000                             ' Number of data elements
Dim aParms(1 To kElements) As String                ' Extraction parameters for each element
Public aConstants(1 To 100) As String               ' Constants for use in RCHGetElementNumber() formulas

Public sElementsLocation As String                  ' Location of element defintion files
Public iInit As Integer                             ' Has data element list been initialized?


Public Function RCHGetElementNumber(ByVal pTicker As String, _
                           Optional ByVal pItem As Integer = 1, _
                           Optional ByVal pError As Variant = "Error", _
                           Optional ByVal pFile As String = "")
                           
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to retrieve individual data elements from various web sites
    '-----------------------------------------------------------------------------------------------------------*
    ' 2005.08.01 -- Fix 10-year summary retrieval of stocks with less than 10 years of history
    ' 2005.08.06 -- Simplify code flow and remove redundant modules, add comments
    ' 2005.08.06 -- Add Yahoo! Finance Analyst Estimates page
    ' 2005.08.09 -- Add several Morningstar items
    ' 2005.08.14 -- Add BusinessWeek Online ProSearch Criteria Report elements
    ' 2005.08.17 -- Add BusinessWeek Online Financial Summary page
    '-----------------------------------------------------------------------------------------------> Version 1.0
    ' 2006.02.01 -- Fix MSN "% Growth Rate -- EPS..." items; now "% Growth Rate -- Net Income..."
    ' 2006.02.01 -- Fix YahooKS Cash Flow table extraction tags
    ' 2006.02.01 -- Fix YahooAE "EPS Trends" table extraction
    ' 2006.02.01 -- Fix YahooAE "Revenue Est" table extraction
    ' 2006.02.01 -- Add BarCharts Technical Report page
    ' 2006.02.01 -- Add "Next Earnings Date" from YahooAE page
    '-----------------------------------------------------------------------------------------------> Version 1.1
    ' 2006.02.04 -- Add Earnings.com statistics
    ' 2006.02.06 -- Add StockCharts.com P&F pattern/price objective
    ' 2006.03.07 -- Fix YahooAE "Earnings Growth Rates" table extraction
    ' 2006.03.10 -- Add several MSN Quarterly Income Statement and Balance Sheet items
    ' 2006.03.10 -- Add several Yahoo Quarterly Income Statement and Balance Sheet items
    ' 2006.03.13 -- Add Reuters Quarterly and Annual Income Statements
    ' 2006.03.13 -- Add process to convert extracted numeric data
    ' 2006.03.14 -- Allow multiple table extract labels to be searched for
    ' 2006.03.15 -- Add pFile parameter to RCHGetElementNumber()
    ' 2006.03.15 -- Fix several BWOpin elements that were out of sync with parsing rules database
    ' 2006.04.24 -- Add process to convert strings like 2.34M or 2.34B into numeric values
    ' 2006.05.25 -- Add Google financial statement elements
    ' 2006.05.26 -- Add MSNMoney financial statement elements
    ' 2006.06.18 -- Fixed "short table" data extraction problems
    ' 2006.06.18 -- Modified MSN FYI Advisor extraction
    ' 2006.06.19 -- Modify numeric conversion process to handle strings like "8.52%"
    ' 2006.06.19 -- Modify BWOpin extractions due to web page changes
    ' 2006.06.19 -- Add new dividend items added to YahooKS web page
    ' 2006.06.19 -- Modify MSN Market Capitalization for web page change
    '-----------------------------------------------------------------------------------------------> Version 1.2
    ' 2006.06.21 -- Major rewrite to create internal tables of extraction definitions
    ' 2006.06.26 -- Add "YahooPM" data source for mutual fund performance elements
    ' 2006.06.26 -- Add "YahooPR" data source for mutual fund profile elements
    ' 2006.06.26 -- Add "YahooRK" data source for mutual fund risk-related elements
    ' 2006.06.26 -- Add "YahooHL" data source for mutual fund holdings-related elements
    ' 2006.06.27 -- Add the ability to create a calculated field
    ' 2006.07.01 -- Add the ADVFN annual and quarterly financial statements elements
    ' 2006.07.05 -- Add the Reuters Ratios Comparison page elements
    ' 2006.07.10 -- Add "YahooMS" data source for the Yahoo Market Statistics page (Advances and Declines)
    ' 2006.07.12 -- Add company name elements for "YahooKS" and "MSN" data sources
    ' 2006.07.12 -- Allow for custom "Error" value
    ' 2006.07.14 -- Modify P&F "Bullish"/"Bearish" decision to look at P&F Pattern before price objective
    ' 2006.07.19 -- Fix usage of pFile parameter
    ' 2006.07.27 -- Add "YahooIN" data source for sector and industry numbers and names
    ' 2006.07.27 -- Fix "Zacks" data parsing due to web page changes
    ' 2006.08.04 -- Add "Web Page" value for pTicker parameter
    ' 2006.08.07 -- Force pTicker parameter to be upper case
    ' 2006.08.09 -- Add calculated fields for Pitroski/Altman/Rule One/Magic Formula
    ' 2006.08.10 -- Add 10-year price history fields from Business Week Online
    ' 2006.08.14 -- Add more Morningstar pages/fields
    ' 2006.08.14 -- Fixed Barchart.com Buy/Hold/Sell 3-table cell parsing process
    ' 2006.08.16 -- Added MsgBox with error message to XMLHTTP failure
    ' 2006.08.16 -- Added Google financial statements currenty type and magnitude
    ' 2006.08.19 -- Changed MSN 10-year summary of financial statement items to MSN10 web page source
    ' 2006.08.23 -- Obsoleted a number of MSN Financial Statement elements (MSN changed to Reuters data)
    ' 2006.08.27 -- Added RCHGetTableCell() function
    ' 2006.08.29 -- Added RCHGetWebData() function
    ' 2006.09.02 -- Misc changes to RCHExtractData function related to pCells=0 and pRows
    ' 2006.09.18 -- Removed MsgBox with error message upon XMLHTTP failure
    ' 2006.10.07 -- Add SMFForceRecalculation() macro
    ' 2006.11.14 -- Fix RCHExtractData processing to convert M/B into Million and Trillions on non-numeric data
    ' 2006.12.11 -- Fix several MSN tags related to 5-year average growth rates (Canadian stocks had different tag)
    ' 2006.12.20 -- Change all HTML table header tags to normal table cell tags on web pages retrieved
    ' 2007.01.04 -- Change CCur() usage to CDec() because of precision issues
    ' 2007.01.07 -- Fix Earnings.com extraction of Dividends and Splits data when no Splits table exists
    ' 2007.01.16 -- Modify TickerReset and ArrayReset processing
    ' 2007.01.17 -- Fix RCHGetTableCell() to return vError instead of blanks when not finding new rows
    ' 2007.01.19 -- Fix conversion of "- " and "-- " to return zero values
    ' 2007.01.23 -- Transfer BWHist and BWTech data elements to pick up Telescan data
    ' 2007.01.23 -- Obsolete remaining BWOpin elements (726 to 727)
    ' 2007.01.23 -- Obsolete all BWPro elements (884 to 940)
    ' 2007.01.23 -- Obsolete all BWSumm elements (994 To 1214)
    ' 2007.01.27 -- Fix HTML table header tag processing for RCHGetElementNumber()
    ' 2007.03.19 -- Add "Too many web page retrievals" error message
    ' 2007.03.20 -- Add ability to retrieve RCHGetElementNumber parameters via formula
    ' 2007.05.22 -- Obsolete all Telescan elements since website is gone (728 to 779, 837 to 847, 13891 to 13930)
    '-----------------------------------------------------------------------------------------------> Version 1.3
    ' 2007.06.03 -- Externalize element definitions for RCHGetElementNumber()
    '-----------------------------------------------------------------------------------------------> Version 2.0a
    ' 2007.08.09 -- Added ability to retrieve data via IE object for items that fail with XMLHTTP
    '-----------------------------------------------------------------------------------------------> Version 2.0b
    ' 2007.08.14 -- Add conversion of "---" to return zero values
    '-----------------------------------------------------------------------------------------------> Version 2.0c
    ' 2007.08.23 -- Removed IIF() function using pUseIE
    ' 2007.08.27 -- Added ability to use "~~~~~" in Find1-Find4 external parameters for RCHGetElementNumber()
    '-----------------------------------------------------------------------------------------------> Version 2.0d
    ' 2007.08.30 -- Added smfGetAData and smfGetAParms functions for debugging purposes
    '-----------------------------------------------------------------------------------------------> Version 2.0f
    ' 2007.09.10 -- Changed smfForceRecalculation for EXCEL 2000 processing
    '-----------------------------------------------------------------------------------------------> Version 2.0g
    ' 2007.09.21 -- Restored code to initialize Morningstar for retrieval of data from their web site
    '-----------------------------------------------------------------------------------------------> Version 2.0h
    ' 2007.09.30 -- Added translation of chr(160) to blank
    ' 2007.11.20 -- Broke out major functions into their own modules
    ' 2008.03.16 -- Add P-TYPE parameter to RCHGetElementNumber()
    '-----------------------------------------------------------------------------------------------> Version 2.0i
    ' 2009.12.19 -- Add smfConvertYahooTicker() processing
    ' 2009.12.30 -- Modified P&F element extractions from StockCharts.com
    ' 2010.07.06 -- Fix smfGetAdvFNElementNumber() parameter numbering
    ' 2010.07.18 -- Changed location of smfConvertYahooTicker() processing to prevent "Undefine" error
    ' 2010.10.10 -- Added code to change HTML code &#151; to a normal hyphen
    ' 2010.10.22 -- Added code to change HTML code &mdash; to a normal hyphen
    ' 2010.12.15 -- Added code to return regional setting
    ' 2011.01.26 -- Obsoleted "FYI Alerts" from MSN
    ' 2011.04.27 -- Convert to use smfGetWebPage() function
    ' 2012.01.11 -- Add ability to define element as SMF formula by using leading "=" on formula
    ' 2012.05.13 -- Change placement of smfConvertYahooTicker() for EVALUATE() functions
    ' 2013.08.14 -- Add sAdvFNPrefix to "Version" output
    ' 2014.03.08 -- Add sElementsLocation to allow element definitions on Internet
    ' 2014.03.15 -- Add vTemp processing for returned array from evaluate() results
    ' 2015.04.29 -- Add "Definition" value for pTicker parameter
    ' 2015.09.15 -- Fix "Rule #1 MOS Price" (used to be based on obsoleted MSN data elements)
    ' 2016.05.18 -- Add Application.PathSeparator to ease transition between operating systems
    ' 2016.06.14 -- Add Operating System and Version of computer to "Version" string
    ' 2017.05.05 -- Add smfSetElementsLocation() sub
    ' 2017.05.19 -- Add aConstants() processing
    ' 2017.05.21 -- Add smfGetaConstants() function
    ' 2017.07.23 -- Remove iMorningStar variable
    ' 2018.08.27 -- ERASE aData() array rather than resetting the individual items
    ' 2020.03.09 -- Replace EVALUATE() function with smfEvaluateTwice(), a fix for Microsoft changes
    ' 2023-01-22 -- Mel Pryor (ClimberMel@gmail.com)
    '               Changed version number to reflect updates to other modules
    '-----------------------------------------------------------------------------------------------> Version 2.0k
    ' > Example of an invocation to get The "Trend Spotter" value for IBM from the BarChart website:
    '
    '   =RCHGetElementNumber("IBM", 701)
    '   =RCHGetTableCell("http://quote.barchart.com/texpert.asp?sym=IBM",1,"Trend Spotter",,,,,1,3)
    '
    '   The first is the preferred method.
    '-----------------------------------------------------------------------------------------------------------*
       
    Dim sURL As String
    Dim sTicker1 As String
    Dim sTicker2 As String
    Dim vTemp As Variant

    On Error GoTo ErrorExit
    vError = pError
    
    '--------------------------------> Special cases to return immediately
    sTicker1 = UCase(pTicker)
    Select Case True
       Case sTicker1 = "": GoTo ErrorExit
       Case sTicker1 = "NONE": GoTo ErrorExit
       Case sTicker1 = "ERROR": GoTo ErrorExit
       Case sTicker1 = "VERSION": RCHGetElementNumber = "Stock Market Functions add-in, Version " & kVersion & _
                       " (" & ThisWorkbook.Path & "; " _
                            & Application.OperatingSystem & "; " _
                            & Application.Version & "; " _
                            & sAdvFNPrefix & "; " _
                            & sElementsLocation & "; " _
                            & Application.International(xlCountrySetting) & ")"
       Case sTicker1 = "COUNTRY": RCHGetElementNumber = Application.International(xlCountrySetting)
       Case pItem > kElements: RCHGetElementNumber = "EOL"
       Case Else: RCHGetElementNumber = ""
       End Select
    If RCHGetElementNumber <> "" Then Exit Function
    
    '--------------------------------> Load extraction definitions if needed
    If iInit = 0 Then
       iInit = 1
       Erase aData, aParms
       'For i1 = 1 To kPages
       '    aData(i1, 1) = ""  ' Reset stored ticker array
       '    Next i1
       Call LoadElementsLocation
       For i1 = 0 To 19
           Select Case sElementsLocation
              Case "Internet": Call LoadElementsFromInternet(i1)
              Case Else: Call LoadElementsFromFile(i1)
              End Select
           Next i1
       Call LoadElementsFromFile(20)  ' User Custom Element Definitions
       Call LoadElementsFromFile(21)  ' User default settings
       Call LoadElementsFromFile(22)  ' Constant values
       End If
    
    '--------------------------------> Additional special cases to return immediately
    aParm = Split(aParms(pItem) & ";N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A", ";")
    sTicker2 = smfConvertYahooTicker(sTicker1, aParm(0))
    Select Case True
       Case aParms(pItem) = "": RCHGetElementNumber = "Undefined"
       Case sTicker1 = "DEFINITION": RCHGetElementNumber = aParms(pItem): Exit Function
       Case sTicker1 = "SOURCE": RCHGetElementNumber = aParm(0): Exit Function
       Case sTicker1 = "ELEMENT": RCHGetElementNumber = aParm(1): Exit Function
       Case sTicker1 = "WEB PAGE": RCHGetElementNumber = aParm(2): Exit Function
       Case sTicker1 = "P-URL": RCHGetElementNumber = aParm(2): Exit Function
       Case sTicker1 = "P-CELLS": RCHGetElementNumber = aParm(3): Exit Function
       Case sTicker1 = "P-FIND1": RCHGetElementNumber = aParm(4): Exit Function
       Case sTicker1 = "P-FIND2": RCHGetElementNumber = aParm(5): Exit Function
       Case sTicker1 = "P-FIND3": RCHGetElementNumber = aParm(6): Exit Function
       Case sTicker1 = "P-FIND4": RCHGetElementNumber = aParm(7): Exit Function
       Case sTicker1 = "P-ROWS": RCHGetElementNumber = aParm(8): Exit Function
       Case sTicker1 = "P-END": RCHGetElementNumber = aParm(9): Exit Function
       Case sTicker1 = "P-LOOK": RCHGetElementNumber = aParm(10): Exit Function
       Case sTicker1 = "P-TYPE": RCHGetElementNumber = aParm(11): Exit Function
       Case aParm(0) = "AdvFN-A"
            RCHGetElementNumber = smfGetADVFNElement(sTicker1, "A", aParm(3), aParm(4), aParm(5), pError)
       Case aParm(0) = "AdvFN-Q"
            RCHGetElementNumber = smfGetADVFNElement(sTicker1, "Q", aParm(3), aParm(4), aParm(5), pError)
       Case aParm(0) = "Evaluate"
#If Mac Then
            RCHGetElementNumber = Evaluate("=" & Replace(aParm(2), "~~~~~", sTicker2))
#Else
            RCHGetElementNumber = smfEvaluateTwice(Replace(aParm(2), "~~~~~", sTicker2))
#End If
       Case Left(aParm(2), 1) = "="
#If Mac Then
            vTemp = Evaluate("=" & Replace(Mid(aParm(2), 2), "~~~~~", sTicker2))
#Else
            vTemp = smfEvaluateTwice(Replace(Mid(aParm(2), 2), "~~~~~", sTicker2))
#End If
            If VarType(vTemp) > 8192 Then
               RCHGetElementNumber = vTemp(1)
            Else
               RCHGetElementNumber = vTemp
               End If
       Case Else: RCHGetElementNumber = ""
       End Select
    If RCHGetElementNumber <> "" Then Exit Function
    '--------------------------------> Preprocess web page data
    sURL = Replace(aParm(2), "~~~~~", sTicker2)
    If aParm(0) = "Calculated" Then
       sData(2) = ""
    Else
       sData(2) = smfGetWebPage(sURL, aParm(11), 0)
       End If
    sData(3) = UCase(sData(2))
    '--------------------------------> Return requested item
    If pFile <> "" Then
       On Error GoTo ErrorExit
       Set wb = Application.Workbooks(pFile).Worksheets(aParm(0))
       iRow = 0
       On Error Resume Next
       iRow = Application.WorksheetFunction.Match(pTicker, wb.Columns("A:A"), 0)
       On Error GoTo ErrorExit
       If iRow > 0 Then
          iCol = 0
          On Error Resume Next
          iCol = Application.WorksheetFunction.Match(pItem, wb.Rows("1:1"), 0)
          On Error GoTo ErrorExit
          If iCol = 0 Then
             Set wb = Application.Workbooks(pFile).Worksheets(aParm(0) & "_2")
             iCol = 0
             On Error Resume Next
             iCol = Application.WorksheetFunction.Match(pItem, wb.Rows("1:1"), 0)
             On Error GoTo ErrorExit
             End If
          If iCol = 0 Then
             Set wb = Application.Workbooks(pFile).Worksheets(aParm(0) & "_3")
             iCol = 0
             On Error Resume Next
             iCol = Application.WorksheetFunction.Match(pItem, wb.Rows("1:1"), 0)
             On Error GoTo ErrorExit
             End If
          RCHGetElementNumber = wb.Cells(iRow, iCol)
          Exit Function
          End If
       End If
    Select Case True
       Case aParm(2) = "?": RCHGetElementNumber = RCHSpecialExtraction("" & aParm(1), sTicker1)
       Case aParm(3) = "?": RCHGetElementNumber = RCHSpecialExtraction("" & aParm(1), sTicker1)
       Case Left(aParm(2), 8) = "Obsolete": RCHGetElementNumber = aParm(2)
       Case Else: RCHGetElementNumber = RCHExtractData(aParm(0), aParm(1), _
                                                       Replace(aParm(4), "~~~~~", sTicker2), _
                                                       Replace(aParm(5), "~~~~~", sTicker2), _
                                                       Replace(aParm(6), "~~~~~", sTicker2), _
                                                       Replace(aParm(7), "~~~~~", sTicker2), _
                                                       aParm(8), aParm(9), aParm(3), aParm(10))
       End Select
    
    Exit Function
ErrorExit: RCHGetElementNumber = vError
    End Function
Private Function RCHSpecialExtraction(pLookup As String, Optional pTicker As String = "")
    On Error GoTo ErrorExit
    Select Case pLookup
        Case "Financial Statements Currency Magnitude" ' Google
             iPos2 = InStr(sData(3), "(EXCEPT FOR PER SHARE ITEMS)")
             iPos2 = InStrRev(sData(3), " OF ", iPos2)
             iPos1 = InStrRev(sData(3), ">IN ", iPos2)
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 4, iPos2 - iPos1 - 4)
        Case "Financial Statements Currency Type" ' Google
             iPos2 = InStr(sData(3), "(EXCEPT FOR PER SHARE ITEMS)")
             iPos1 = InStrRev(sData(3), " OF ", iPos2)
             iPos2 = InStr(iPos1, sData(3), "<")
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 4, iPos2 - iPos1 - 4)
        Case "FYI Alerts" ' MSN
             RCHSpecialExtraction = "No longer available"
             'iPos1 = InStr(sData(3), ">ADVISOR FYI<")
        Case "Company Description" ' MSN
             iPos1 = InStr(sData(3), "<BODY")
             iPos1 = InStr(iPos1, sData(3), "COMPANY REPORT")
             iPos1 = InStr(iPos1, sData(3), "<P>") + 2
             iPos2 = InStr(iPos1, sData(3), "</P>")
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 1, iPos2 - iPos1 - 1)
        Case "Company Name" ' MSN
             iPos1 = InStr(sData(3), "<TITLE>")
             iPos2 = InStr(iPos1, sData(3), " REPORT - ")
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 7, iPos2 - iPos1 - 7)
        Case "Risk Grade"  ' MSN
             iPos1 = InStr(sData(3), "RISK:") + 6
             If iPos1 = 6 Then
                RCHSpecialExtraction = vError
             Else
                i1 = CInt(Mid(sData(2), iPos1, 1))
                RCHSpecialExtraction = Mid("ABCDF", i1, 1)
                End If
        Case "Return Grade"  ' MSN
             iPos1 = InStr(sData(3), "RETURN:") + 8
             If iPos1 = 8 Then
                RCHSpecialExtraction = vError
             Else
                i1 = CInt(Mid(sData(2), iPos1, 1))
                RCHSpecialExtraction = Mid("FDCBA", i1, 1)
                End If
        Case "Quick Summary"  ' MSN
             RCHSpecialExtraction = ""
             iPos1 = InStr(sData(3), "RETURN:") + 8
             iPos2 = InStr(sData(3), "QUICK SUMMARY")
             For i1 = 1 To 20
                 iPos2 = InStr(iPos2, sData(3), "<DD>") + 4
                 If iPos2 > iPos1 Or iPos2 = 4 Then Exit For
                 iPos3 = InStr(iPos2, sData(3), "<B>")
                 iPos4 = InStr(iPos2, sData(3), "</B>")
                 s1 = Mid(sData(2), iPos3 + 3, iPos4 - iPos3 - 3)
                 s2 = Mid(sData(2), iPos2, iPos3 - iPos2)
                 RCHSpecialExtraction = RCHSpecialExtraction & s1 & " -- " & s2 & vbLf
                 Next i1
        Case "StockScouter Rating -- Summary"  ' MSN
             iPos1 = InStr(sData(3), "ALT=""STOCKSCOUTER RATING: ")
             iPos1 = InStr(iPos1, sData(3), "<P>") + 3
             iPos2 = InStr(iPos1, sData(3), "</P>")
             RCHSpecialExtraction = Replace(Replace(Mid(sData(2), iPos1, iPos2 - iPos1), "<b>", ""), "</b>", "")
        Case "Short Term Outlook"  ' MSN
             iPos1 = InStr(sData(3), "SHORT-TERM OUTLOOK")
             iPos1 = InStr(iPos1, sData(3), "<P>") + 3
             iPos2 = InStr(iPos1, sData(3), "</P>")
             RCHSpecialExtraction = Replace(Replace(Mid(sData(2), iPos1, iPos2 - iPos1), "<b>", ""), "</b>", "")
        Case "StockScouter Rating -- Current"  ' MSN
             iPos1 = InStr(sData(3), "ALT=""STOCKSCOUTER RATING: ")
             iPos2 = InStr(iPos1, sData(3), ":") + 2
             iPos3 = InStr(iPos2, sData(3), """")
             RCHSpecialExtraction = CInt(Mid(sData(2), iPos2, iPos3 - iPos2))
        Case "Risk Alert Level" ' Reuters
             iPos1 = InStr(sData(3), "IMAGES/SELLALERT")
             iPos1 = InStr(iPos1, sData(3), "ALT=""") + 5
             iPos2 = InStr(iPos1, sData(3), """")
             RCHSpecialExtraction = Mid(sData(2), iPos1, iPos2 - iPos1)
        Case "P&F -- Pattern" ' Stockcharts
             iPos1 = InStr(sData(3), "P&F PATTERN:")
             If iPos1 = 0 Then
                RCHSpecialExtraction = "No P&F Pattern Found"
                Exit Function
                End If
             iPos2 = InStr(iPos1, sData(3), "</DIV")
             iPos1 = InStrRev(sData(3), ">", iPos2) + 1
             iPos3 = InStrRev(sData(3), "#00AA00", iPos1)
             If iPos1 - iPos3 < 40 Then
                RCHSpecialExtraction = "Bullish -- " & Trim(Mid(sData(2), iPos1, iPos2 - iPos1))
                Exit Function
                End If
             iPos3 = InStrRev(sData(3), "#FF0000", iPos1)
             If iPos1 - iPos3 < 40 Then
                RCHSpecialExtraction = "Bearish -- " & Trim(Mid(sData(2), iPos1, iPos2 - iPos1))
                Exit Function
                End If
             RCHSpecialExtraction = "Unknown -- " & Trim(Mid(sData(2), iPos1, iPos2 - iPos1))
        Case "P&F -- Price Objective"  ' Stockcharts
             iPos1 = InStr(sData(3), " PRICE OBJ. ")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos1 = InStr(iPos1, sData(3), ":") + 2
             iPos2 = InStr(iPos1, sData(3), "<")
             RCHSpecialExtraction = Trim(Mid(sData(2), iPos1, iPos2 - iPos1))
        Case "P&F -- Trend"  ' Stockcharts
             iPos1 = InStr(sData(3), " PRICE OBJ. ")
             If iPos1 > 0 Then
                RCHSpecialExtraction = Mid(sData(2), iPos1 - 7, 7)
             Else
                RCHSpecialExtraction = "Unknown"
                End If
        Case "Next Earnings Date" ' Yahoo
             iPos1 = InStr(sData(3), "NEXT EARNINGS DATE: ") + 20
             If iPos1 = 20 Then
                RCHSpecialExtraction = "N/A"
                Exit Function
                End If
             iPos2 = InStr(iPos1, sData(3), " - ")
             If iPos2 = 0 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1, iPos2 - iPos1)
        Case "Sector Number" ' Yahoo
             iPos1 = InStr(sData(3), "HTTP://BIZ.YAHOO.COM/P/")
             If iPos1 = 0 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 23, 1)
        Case "Industry Number" ' Yahoo
             iPos1 = InStr(sData(3), ">INDUSTRY:<")
             iPos1 = InStr(iPos1, sData(3), "HTTP://BIZ.YAHOO.COM/IC/")
             If iPos1 = 0 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 24, 3)
        Case "Industry Symbol" ' Yahoo
             iPos1 = InStr(sData(3), ">^")
             If iPos1 = 0 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 1, 8)
        Case "Company Name" ' Yahoo
             iPos2 = InStr(sData(3), " (" & pTicker & ")</B>")
             iPos1 = InStrRev(sData(3), "<B>", iPos2)
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 3, iPos2 - iPos1 - 3)
        Case "Fund Profile -- Morningstar Rating" ' Yahoo
             Select Case True
                Case InStr(sData(3), "/STAR1.GIF") > 0: RCHSpecialExtraction = 1
                Case InStr(sData(3), "/STAR2.GIF") > 0: RCHSpecialExtraction = 2
                Case InStr(sData(3), "/STAR3.GIF") > 0: RCHSpecialExtraction = 3
                Case InStr(sData(3), "/STAR4.GIF") > 0: RCHSpecialExtraction = 4
                Case InStr(sData(3), "/STAR5.GIF") > 0: RCHSpecialExtraction = 5
                Case Else: RCHSpecialExtraction = vError
                End Select
        Case "Fund Profile -- Last Dividend -- Date" ' Yahoo
             iPos1 = InStr(sData(3), ">FUND OPERATIONS")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos1 = InStr(iPos1, sData(3), "LAST DIVIDEND")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos1 = InStr(iPos1, sData(3), "(")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos2 = InStr(iPos1, sData(3), ")")
             If iPos2 < iPos1 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 1, iPos2 - iPos1 - 1)
        Case "Fund Profile -- Last Cap Gain -- Date" ' Yahoo
             iPos1 = InStr(sData(3), ">FUND OPERATIONS")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos1 = InStr(iPos1, sData(3), "LAST CAP GAIN")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos1 = InStr(iPos1, sData(3), "(")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos2 = InStr(iPos1, sData(3), ")")
             If iPos2 < iPos1 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 1, iPos2 - iPos1 - 1)
        Case "Piotroski 1 (Positive Net Income)"
             n1 = RCHGetElementNumber(pTicker, 8806)  ' FQ1, Net Income (Continuing Operations)
             n2 = RCHGetElementNumber(pTicker, 8807)  ' FQ2, Net Income (Continuing Operations)
             n3 = RCHGetElementNumber(pTicker, 8808)  ' FQ3, Net Income (Continuing Operations)
             n4 = RCHGetElementNumber(pTicker, 8809)  ' FQ4, Net Income (Continuing Operations)
             RCHSpecialExtraction = IIf((n1 + n2 + n3 + n4) > 0, 1, 0)
        Case "Piotroski 2 (Positive Operating Cash Flow)"
             n1 = RCHGetElementNumber(pTicker, 11326) ' FQ1, YTD Net Cash Flow (Continuing Operations)
             n2 = RCHGetElementNumber(pTicker, 11330) ' FQ5, YTD Net Cash Flow (Continuing Operations)
             n3 = RCHGetElementNumber(pTicker, 6856)  ' FY1, Net Cash Flow (Continuing Operations)
             RCHSpecialExtraction = IIf(n1 - n2 + n3 > 0, 1, 0)
        Case "Piotroski 3 (Increasing Net Income)"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n6 = RCHGetElementNumber(pTicker, 5596)  ' FY1, Net Income (Continuing Operations)
                N7 = RCHGetElementNumber(pTicker, 5597)  ' FY2, Net Income (Continuing Operations)
             Else
                n2 = RCHGetElementNumber(pTicker, 8806)  ' FQ1, Net Income (Continuing Operations)
                n3 = RCHGetElementNumber(pTicker, 8807)  ' FQ2, Net Income (Continuing Operations)
                n4 = RCHGetElementNumber(pTicker, 8808)  ' FQ3, Net Income (Continuing Operations)
                n5 = RCHGetElementNumber(pTicker, 8809)  ' FQ4, Net Income (Continuing Operations)
                n6 = n2 + n3 + n4 + n5
                N7 = RCHGetElementNumber(pTicker, 5596)  ' FY1, Net Income (Continuing Operations)
                End If
             RCHSpecialExtraction = IIf(n6 > N7, 1, 0)
        Case "Piotroski 4 (Operating Cash flow exceeds Net Income)"
             n1 = RCHGetElementNumber(pTicker, 11326) ' FQ1, YTD Net Cash Flow (Continuing Operations)
             n2 = RCHGetElementNumber(pTicker, 11330) ' FQ5, YTD Net Cash Flow (Continuing Operations)
             n3 = RCHGetElementNumber(pTicker, 6856)  ' FY1, Net Cash Flow (Continuing Operations)
             n4 = RCHGetElementNumber(pTicker, 8806)  ' FQ1, Net Income (Continuing Operations)
             n5 = RCHGetElementNumber(pTicker, 8807)  ' FQ2, Net Income (Continuing Operations)
             n6 = RCHGetElementNumber(pTicker, 8808)  ' FQ3, Net Income (Continuing Operations)
             N7 = RCHGetElementNumber(pTicker, 8809)  ' FQ4, Net Income (Continuing Operations)
             RCHSpecialExtraction = IIf(n1 - n2 + n3 > n4 + n5 + n6 + N7, 1, 0)
        Case "Piotroski 5 (Decreasing ratio of long-term debt to assets )"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n2 = RCHGetElementNumber(pTicker, 6376)  ' FY1, Long Term Debt
                n3 = RCHGetElementNumber(pTicker, 6266)  ' FY1, Total Assets
                n4 = RCHGetElementNumber(pTicker, 6377)  ' FY2, Long Term Debt
                n5 = RCHGetElementNumber(pTicker, 6267)  ' FY2, Total Assets
             Else
                n2 = RCHGetElementNumber(pTicker, 10366) ' FQ1, Long Term Debt
                n3 = RCHGetElementNumber(pTicker, 10146) ' FQ1, Total Assets
                n4 = RCHGetElementNumber(pTicker, 6376)  ' FY1, Long Term Debt
                n5 = RCHGetElementNumber(pTicker, 6266)  ' FY1, Total Assets
                End If
             RCHSpecialExtraction = IIf((n2 / n3) < (n4 / n5), 1, 0)
        Case "Piotroski 6 (Increasing Current Ratio)"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n2 = RCHGetElementNumber(pTicker, 6116)  ' FY1, Current Assets
                n3 = RCHGetElementNumber(pTicker, 6366)  ' FY1, Current Liabilities
                n4 = RCHGetElementNumber(pTicker, 6117)  ' FY2, Current Assets
                n5 = RCHGetElementNumber(pTicker, 6367)  ' FY2, Current Liabilities
             Else
                n2 = RCHGetElementNumber(pTicker, 9846)  ' FQ1, Current Assets
                n3 = RCHGetElementNumber(pTicker, 10346) ' FQ1, Current Liabilities
                n4 = RCHGetElementNumber(pTicker, 6116)  ' FY1, Current Assets
                n5 = RCHGetElementNumber(pTicker, 6366)  ' FY1, Current Liabilities
                End If
             RCHSpecialExtraction = IIf((n2 / n3) > (n4 / n5), 1, 0)
        Case "Piotroski 7 (No increase in outstanding shares)"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n2 = RCHGetElementNumber(pTicker, 6646)  ' FY1, Total Common Shares Out
                n3 = RCHGetElementNumber(pTicker, 6647)  ' FY2, Total Common Shares Out
             Else
                n2 = RCHGetElementNumber(pTicker, 10906) ' FQ1, Total Common Shares Out
                n3 = RCHGetElementNumber(pTicker, 6646)  ' FY1, Total Common Shares Out
                End If
             RCHSpecialExtraction = IIf(n2 > n3, 0, 1)
        Case "Piotroski 8 (Increasing Gross Margins)"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n6 = RCHGetElementNumber(pTicker, 5346)  ' FY1, Gross Operating Profit
                N7 = RCHGetElementNumber(pTicker, 5347)  ' FY2, Gross Operating Profit
                n8 = RCHGetElementNumber(pTicker, 5286)  ' FY1, Operating Revenue
                n9 = RCHGetElementNumber(pTicker, 5287)  ' FY2, Operating Revenue
             Else
                n2 = RCHGetElementNumber(pTicker, 8306)  ' FQ1, Gross Operating Profit
                n3 = RCHGetElementNumber(pTicker, 8307)  ' FQ2, Gross Operating Profit
                n4 = RCHGetElementNumber(pTicker, 8308)  ' FQ3, Gross Operating Profit
                n5 = RCHGetElementNumber(pTicker, 8309)  ' FQ4, Gross Operating Profit
                n6 = n2 + n3 + n4 + n5
                N7 = RCHGetElementNumber(pTicker, 5346)  ' FY1, Gross Operating Profit
                n2 = RCHGetElementNumber(pTicker, 8186)  ' FQ1, Operating Revenue
                n3 = RCHGetElementNumber(pTicker, 8187)  ' FQ2, Operating Revenue
                n4 = RCHGetElementNumber(pTicker, 8188)  ' FQ3, Operating Revenue
                n5 = RCHGetElementNumber(pTicker, 8189)  ' FQ4, Operating Revenue
                n8 = n2 + n3 + n4 + n5
                n9 = RCHGetElementNumber(pTicker, 5286)  ' FY1, Operating Revenue
                End If
             RCHSpecialExtraction = IIf((n6 / n8) > (N7 / n9), 1, 0)
        Case "Piotroski 9 (Increasing Asset Turnover)"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n6 = RCHGetElementNumber(pTicker, 5286)  ' FY1, Operating Revenue
                N7 = RCHGetElementNumber(pTicker, 6266)  ' FY1, Total Assets
                n8 = RCHGetElementNumber(pTicker, 5287)  ' FY2, Operating Revenue
                n9 = RCHGetElementNumber(pTicker, 6267)  ' FY2, Total Assets
             Else
                n2 = RCHGetElementNumber(pTicker, 8186)  ' FQ1, Operating Revenue
                n3 = RCHGetElementNumber(pTicker, 8187)  ' FQ2, Operating Revenue
                n4 = RCHGetElementNumber(pTicker, 8188)  ' FQ3, Operating Revenue
                n5 = RCHGetElementNumber(pTicker, 8189)  ' FQ4, Operating Revenue
                n6 = n2 + n3 + n4 + n5
                N7 = RCHGetElementNumber(pTicker, 10146) ' FQ1, Total Assets
                n8 = RCHGetElementNumber(pTicker, 5286)  ' FY1, Operating Revenue
                n9 = RCHGetElementNumber(pTicker, 6266)  ' FY1, Total Assets
                End If
             RCHSpecialExtraction = IIf((n6 / N7) > (n8 / n9), 1, 0)
        Case "Piotroski F-Score"
             n1 = RCHGetElementNumber(pTicker, 15001)
             n2 = RCHGetElementNumber(pTicker, 15002)
             n3 = RCHGetElementNumber(pTicker, 15003)
             n4 = RCHGetElementNumber(pTicker, 15004)
             n5 = RCHGetElementNumber(pTicker, 15005)
             n6 = RCHGetElementNumber(pTicker, 15006)
             N7 = RCHGetElementNumber(pTicker, 15007)
             n8 = RCHGetElementNumber(pTicker, 15008)
             n9 = RCHGetElementNumber(pTicker, 15009)
             RCHSpecialExtraction = n1 + n2 + n3 + n4 + n5 + n6 + N7 + n8 + n9
        Case "Altman Z-Score"
             n1 = RCHGetElementNumber(pTicker, 10786)  ' FQ1, Working Capital
             n2 = RCHGetElementNumber(pTicker, 10146)  ' FQ1, Total Assets
             n3 = RCHGetElementNumber(pTicker, 10646)  ' FQ1, Retained Earnings
             n4 = RCHGetElementNumber(pTicker, 8666)   ' FQ1, EBIT
             n5 = RCHGetElementNumber(pTicker, 8667)   ' FQ2, EBIT
             n6 = RCHGetElementNumber(pTicker, 8668)   ' FQ3, EBIT
             N7 = RCHGetElementNumber(pTicker, 8669)   ' FQ4, EBIT
             n8 = n4 + n5 + n6 + N7
             n9 = RCHGetElementNumber(pTicker, 941)   ' Market Capitalization
             n10 = RCHGetElementNumber(pTicker, 10526) ' Total Liabilities
             n11 = RCHGetElementNumber(pTicker, 8186)  ' FQ1, Operating Revenue
             n12 = RCHGetElementNumber(pTicker, 8187)  ' FQ2, Operating Revenue
             n13 = RCHGetElementNumber(pTicker, 8188)  ' FQ3, Operating Revenue
             n14 = RCHGetElementNumber(pTicker, 8189)  ' FQ4, Operating Revenue
             n15 = n11 + n12 + n13 + n14
             RCHSpecialExtraction = 1.2 * (n1 / n2) + 1.4 * (n3 / n2) + 3.3 * (n8 / n2) + 0.6 * (n9 / n10 / 1000) + (n15 / n2)
        Case "Rule #1 MOS Price"
             n1 = RCHGetElementNumber(pTicker, 13630)  ' 5-Year High P/E from Reuter's
             n2 = RCHGetElementNumber(pTicker, 13634)  ' 5-Year Low P/E from Reuter's
             n3 = RCHGetElementNumber(pTicker, 962)    ' Current EPS from Yahoo
             n4 = RCHGetElementNumber(pTicker, 621)    ' 5-Year Projected Growth Rate from Yahoo
             If n1 > 50 Then n1 = 50
             n5 = FV(n4, 10, 0, -n3)
             n6 = PV(0.15, 10, 0, -n5 * (n1 + n2) / 2) / 2
             RCHSpecialExtraction = n6
        Case "Magic Formula Investing -- Earnings Yield"
             n1 = RCHGetElementNumber(pTicker, 949)    ' Enterprise value to EBITDA
             RCHSpecialExtraction = 1 / n1
        Case "Magic Formula Investing -- Return on Capital"
             n1 = RCHGetElementNumber(pTicker, 960)    ' EBITDA
             n2 = RCHGetElementNumber(pTicker, 964)    ' Cash
             n3 = RCHGetElementNumber(pTicker, 10026)  ' FQ1, Net Fixed Assets (Plant & Equipment)
             RCHSpecialExtraction = n1 / (n2 + 1000 * n3)
        Case Else: RCHSpecialExtraction = vError
        End Select
        Exit Function
ErrorExit: RCHSpecialExtraction = vError
    End Function

Sub LoadElementsFromFile(pSuffix As Variant)
    '------------------------------------------------------------------------------------------------------*
    ' 2017.05.05 -- Add processing for settings
    ' 2017.05.19 -- Add processing for constants
    '------------------------------------------------------------------------------------------------------*
    On Error GoTo ErrorExit
    Open ThisWorkbook.Path & Application.PathSeparator & "smf-elements-" & pSuffix & ".txt" For Input As #1
    On Error Resume Next
    Do Until EOF(1) = True
       Line Input #1, sLine
       Select Case True
          Case Left(sLine, 1) = "'"
          Case Left(UCase(sLine), 7) = "SETTING"
               Evaluate (smfWord(sLine, 3, ";"))
          Case Left(UCase(sLine), 8) = "CONSTANT"
               i1 = CInt(smfWord(sLine, 2, ";"))
               aConstants(i1) = smfWord(sLine, 4, ";")
          Case Else
               iPos1 = InStr(sLine, ";")
               aParms(CInt(Left(sLine, iPos1 - 1))) = Mid(sLine, iPos1 + 1)
          End Select
       Loop
    Close #1
ErrorExit:
    End Sub
Sub LoadElementsFromInternet(pSuffix As Variant)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2016.05.18 -- Change to RCHGetURLData to ease transition between operating systems
    '-----------------------------------------------------------------------------------------------------------*
    On Error GoTo ErrorExit
    s1 = RCHGetURLData("http://ogres-crypt.com/SMF/Elements/smf-elements-" & pSuffix & ".txt")
    If s1 = "Error" Then Exit Sub
    v1 = Split(s1, Chr(13) & Chr(10))
    For i1 = 0 To UBound(v1)
       Select Case Left(Trim(v1(i1)) & " ", 1)
          Case "'"
          Case " "
          Case Else
               iPos1 = InStr(v1(i1), ";")
               aParms(CInt(Left(v1(i1), iPos1 - 1))) = Mid(v1(i1), iPos1 + 1)
          End Select
        Next i1
    x = 1
ErrorExit:
    End Sub
    
Sub smfSetElementsLocation(p1 As String)
    sElementsLocation = p1
    End Sub
    
Sub LoadElementsLocation()
    If sElementsLocation = "" Then sElementsLocation = "Local"
    On Error GoTo ErrorExit
    Open ThisWorkbook.Path & Application.PathSeparator & "smf-Elements-Location.txt" For Input As #1
    Line Input #1, sElementsLocation
    Close #1
ErrorExit:
    End Sub

Public Function smfGetAParms(p1 As Integer)
    smfGetAParms = aParms(p1)
    End Function

Public Function smfGetAConstants(p1 As Integer)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.05.21 -- New function to view constant values
    '-----------------------------------------------------------------------------------------------------------*
    smfGetAConstants = aConstants(p1)
    End Function

