Attribute VB_Name = "modGetYahooQuotes"
Public Function RCHGetYahooQuotes(ByVal pTickers As Variant, _
                         Optional ByVal pItems As Variant = "sl1d1t1c1ohgv", _
                         Optional ByVal pServerID As String = "", _
                         Optional ByVal pRefresh As Variant = 0, _
                         Optional ByVal pHeader As Integer = 0, _
                         Optional ByVal pDim1 As Integer = 0, _
                         Optional ByVal pDim2 As Integer = 0, _
                         Optional ByVal pDelimiter As String = ",") ' As Variant()
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to download historical quotes from Yahoo!
    '-----------------------------------------------------------------------------------------------------------*
    '               Original code written by Randy Harmelink
    ' 2005.02.16 -- New function; Adapted from other VBA modules
    ' 2005.06.18 -- Add code to convert numeric items to values instead of leaving as strings
    ' 2006.08.08 -- Return Yahoo response if no data can be parsed from the response
    ' 2006.08.17 -- Fixed kDim1/kDim2 processing using iDim1/iDim2
    ' 2006.08.18 -- Fixed line parsing to allow for double-quoted fields in comma-delimited data
    ' 2006.09.12 -- Fixed line parsing for truncation of last digit of non-double-quoted fields
    ' 2006.09.12 -- Added pServerPrefix parameter
    ' 2007.01.17 -- Change CCur() usage to CDec() because of precision issues
    ' 2007.01.19 -- Change URL of quotes server
    ' 2007.08.28 -- Change end-of-line (CR+LF) processing to handle changes in data files
    ' 2007.08.28 -- Convert strings with percentage values into actual percents
    ' 2007.08.30 -- Added pRefresh parameter to allow easy recalculation using NOW() as its passed value
    ' 2007.09.18 -- Modify pDim1/pDim2 processing
    ' 2007.09.26 -- Added pHeader parameter to allow insertion of column headings
    ' 2008.03.10 -- Added another double carriage return removal from parsing process
    ' 2008.07.17 -- Added ability to process any passed URL for a CSV file
    ' 2009.09.28 -- Added pDelimiter parameter
    ' 2010.04.21 -- Modify pDim1/pDim2 processing so return size can be overridden
    ' 2010.05.15 -- Make sure returned string is no longer than 255 bytes (causes #VALUE! error)
    ' 2011.04.28 -- Change cDec() to smfConvertData()
    ' 2011.07.03 -- Add LCASE() function to sItems concatenation to URL
    ' 2012.06.11 -- Force pServerID to be the U.S. server
    ' 2014.01.14 -- Use "XXXXXX" as a placeholder symbol where spaces are found in the passed ticker array
    ' 2014.05.23 -- Prevent "XXXXXX" placeholder lines from having data displayed
    ' 2015.06.08 -- Fix for unexpected "," field in GuruFocus CSV file that stopped parsing
    ' 2016.05.18 -- Modify heading creation to ease transition between operating systems
    ' 2017.04.26 -- Change "http://" protocol to "https://"
    ' 2017.05.31 -- Add alternative process to get CSV file if "&crumb=" is in the URL
    ' 2023-01-21 -- Mel Pryor (ClimberMel@gmail.com)
    '               Note that if calling module such as RCHGetYahooHistory provides a URL in pTickers that will get used later in code
    '               as long as pItems is ""
    ' 2023-03-01    NOTE that the RCHGetYahooQuotes function is considered OBSOLETE.  The only part of this function that works is
    '               calling it with a URL and pItems = "" (see example below)
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for IBM and MMM:
    '   =RCHGetYahooQuotes("IBM,MMM")                       OBSOLETE
    '   =RCHGetYahooQuotes("IBM,MMM",,,NOW())               OBSOLETE
    '   =RCHGetYahooQuotes("IBM,MMM","l1d1t1",,NOW(),1)     OBSOLETE
    '
    '   Example calling with URL to return table (pItems needs to be "")
    '   =RCHGetYahooQuotes("https://query1.finance.yahoo.com/v7/finance/download/msft?period1=1262304000&period2=1735689600&interval=1devents=history&includeAdjustedClose=true", "")
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim sURL As String
    Dim sItems As String
    
    '------------------> Determine size of array to return
    kDim1 = pDim1  ' Rows
    kDim2 = pDim2  ' Columns
    If pDim1 = 0 Or pDim2 = 0 Then
       If pDim1 = 0 Then kDim1 = 200   ' Old default
       If pDim2 = 0 Then kDim2 = 100   ' Old default
       On Error Resume Next
       kDim1 = Application.Caller.Rows.Count
       kDim2 = Application.Caller.Columns.Count + 1
       On Error GoTo ErrorExit
       End If
  
    '------------------> Initialize return array
    ReDim vData(1 To kDim1, 1 To kDim2) As Variant
    For i1 = 1 To kDim1
        For i2 = 1 To kDim2
            vData(i1, i2) = ""
            Next i2
        Next i1
    
    '------------------> Create URL
    Select Case VarType(pItems)
        Case vbString
             sItems = Replace(pItems, " ", "")
        Case Is >= 8192
             sItems = ""
             For Each oCell In pItems
                 sItems = sItems & oCell.Value
                 Next oCell
        Case Else
            GoTo ErrorExit
        End Select
    Select Case VarType(pTickers)
        Case vbString
             If pTickers = "None" Then GoTo ErrorExit
             sTickers = Replace(pTickers, ",", "+")
        Case Is >= 8192
             sTickers = ""
             For Each oCell In pTickers
                 If oCell.Value > " " Then
                    sTickers = sTickers & oCell.Value & "+"
                 Else
                    sTickers = sTickers & "XXXXXX" & "+"
                    End If
                 Next oCell
             sTickers = Left(sTickers, Len(sTickers) - 1)
        Case Else
            GoTo ErrorExit
        End Select
        
    '------------------> Set the quotes delimiter based on server prefix
    pServerID = ""     ' Temporary?
    Select Case pServerID
        Case ""
             sURL = "https://download.finance.yahoo.com/d/quotes.csv?s="
        Case "jp"
             sURL = "https://finance.yahoo.com." & pServerPrefix & "/d/quotes.csv?s="
        Case "mx"
             sURL = "https://" & pServerID & ".finance.yahoo.com/d/quotes.csv?s="
             sTickers = Replace(sTickers, "+", ",")
        Case Else
             sURL = "https://" & pServerID & ".finance.yahoo.com/d/quotes.csv?s="
        End Select
    sItems = LCase(sItems)
    sURL = sURL & sTickers & "&f=s" & sItems & "&e=.ignore"
    
    '------------------> Set the quotes delimiter based on server prefix
    Select Case pServerID
        Case "ar", "fr", "de", "it"
             sDel = ";"
        Case Else
             sDel = ","
        End Select
    
    '------------------> Create column headings if requested
    If pHeader = 1 Then
       iPos = 1
       iPtr = 1
       Do While (iPos <= Len(sItems))
          sTemp = smfYahooCodeDesc(Mid(sItems & " ", iPos, 2))
          If sTemp = "--" Then
             sTemp = smfYahooCodeDesc(Mid(sItems, iPos, 1))
             iPos = iPos + 1
          Else
             iPos = iPos + 2
             End If
          If iPtr > kDim2 Then Exit Do
          vData(1, iPtr) = sTemp
          iPtr = iPtr + 1
          Loop
       
       End If
    
    '------------------> Overrides for specified CSV file
    If sItems = "" Then
       sURL = pTickers      'This switches sURL back to URL provided in pTickers from calling module
       pHeader = 0
       sDel = pDelimiter
       iOffset = 0
    Else
       iOffset = -1
       End If
    
    '------------------> Download current quotes
    If InStr(sURL, "&crumb=") > 0 Then
       sqData = smfGetYahooHistoryCSVData(sURL)
    Else
       sqData = RCHGetURLData(sURL) '& Chr(13)
       End If
    vData(1 + pHeader, 1) = sqData
    
    '------------------> Parse returned data
    'sqData = Replace(sqData, Chr(10), Chr(13))
    'sqData = Replace(sqData, Chr(13) & Chr(13), Chr(13))
    'sqData = Replace(sqData, Chr(13) & Chr(13), Chr(13))
    sqData = Replace(sqData, vbCrLf, vbLf)
    sqData = Replace(sqData, ""","","",""", ""","" "",""")  ' Fix for GuruFocus?
    aqData = Split(sqData, vbLf)
    iDim1 = UBound(aqData, 1)
    If iDim1 > kDim1 Then iDim1 = kDim1
    For i1 = 0 To iDim1 - 1
        iPos1 = 1
        For i2 = 0 To 200
            If i2 + 1 > kDim2 Then Exit For
            If iPos1 > Len(aqData(i1)) Then Exit For
            sFind = IIf(Mid(aqData(i1), iPos1, 1) = Chr(34), Chr(34), "") & sDel
            iPos2 = InStr(iPos1, aqData(i1) & sDel, sFind)
            s1 = Left(Mid(aqData(i1), iPos1 + Len(sFind) - 1, iPos2 - iPos1 - Len(sFind) + 1), 255)
            s2 = Trim(s1)
            If iOffset = -1 And s2 = "XXXXXX" Then Exit For
            If Right(s2, 1) = "%" Then
               n1 = 100
               s2 = Left(s2, Len(s2) - 1)
            Else
               n1 = 1
               End If
            On Error Resume Next
            s1 = smfConvertData(s2) / n1
            On Error GoTo ErrorExit
            If iOffset = 0 Or i2 > 0 Then vData(i1 + 1 + pHeader, i2 + 1 + iOffset) = s1
            iPos1 = iPos2 + Len(sFind)
            Next i2
        Next i1

ErrorExit:
    RCHGetYahooQuotes = vData
    End Function

Function smfYahooCodeDesc(pCode As String)
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to give a description to a Yahoo code (to be used for column headings)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2007.09.09 -- Created function
    '-----------------------------------------------------------------------------------------------> Version 2.0h
    ' > Examples of an invocation:
    '
    '   =smfYahooCodeDesc("l1")
    '   =smfYahooCodeDesc(A1)
    '-----------------------------------------------------------------------------------------------------------*

    Select Case pCode
       Case "a": smfYahooCodeDesc = "Ask"
       Case "a2": smfYahooCodeDesc = "Average Daily Volume"
       Case "a5": smfYahooCodeDesc = "Ask Size"
       Case "b": smfYahooCodeDesc = "Bid"
       Case "b2": smfYahooCodeDesc = "Ask (ECN)"
       Case "b3": smfYahooCodeDesc = "Bid (ECN)"
       Case "b4": smfYahooCodeDesc = "Book Value"
       Case "b6": smfYahooCodeDesc = "Bid Size"
       Case "c": smfYahooCodeDesc = "Change & Percent"
       Case "c1": smfYahooCodeDesc = "Change"
       Case "c6": smfYahooCodeDesc = "Change (ECN)"
       Case "c8": smfYahooCodeDesc = "After Hours Change (ECN)"
       Case "d": smfYahooCodeDesc = "Dividend/Share"
       Case "d1": smfYahooCodeDesc = "Date of Last Trade"
       Case "e": smfYahooCodeDesc = "Earnings/Share"
       Case "e3": smfYahooCodeDesc = "Expiration date"
       Case "e7": smfYahooCodeDesc = "EPS Est. Current Yr"
       Case "e8": smfYahooCodeDesc = "EPS Est. Next Year"
       Case "e9": smfYahooCodeDesc = "EPS Est. Next Quarter"
       Case "f6": smfYahooCodeDesc = "Float Shares"
       Case "g": smfYahooCodeDesc = "Low"
       Case "g5": smfYahooCodeDesc = "Holdings Gain & Percent (ECN)"
       Case "g6": smfYahooCodeDesc = "Holdings Gain (ECN)"
       Case "h": smfYahooCodeDesc = "High"
       Case "i5": smfYahooCodeDesc = "Order Book (ECN)"
       Case "j": smfYahooCodeDesc = "52-week Low"
       Case "j": smfYahooCodeDesc = "52-week Low"
       Case "j1": smfYahooCodeDesc = "Market Capitalization"
       Case "j3": smfYahooCodeDesc = "Market Cap (ECN)"
       Case "j4": smfYahooCodeDesc = "EBITDA"
       Case "j5": smfYahooCodeDesc = "Change From 52-week Low"
       Case "j6": smfYahooCodeDesc = "Pct Chg From 52-week Low"
       Case "k": smfYahooCodeDesc = "52-week High"
       Case "k": smfYahooCodeDesc = "52-week High"
       Case "k1": smfYahooCodeDesc = "Last Trade (ECN with Time)"
       Case "k2": smfYahooCodeDesc = "Change & Percent (ECN)"
       Case "k3": smfYahooCodeDesc = "Last Trade Size"
       Case "k4": smfYahooCodeDesc = "Change From 52-week High"
       Case "k5": smfYahooCodeDesc = "Pct Chg From 52-week High"
       Case "l": smfYahooCodeDesc = "Last Trade (With Time)"
       Case "l1": smfYahooCodeDesc = "Last Trade"
       Case "m": smfYahooCodeDesc = "Day's Range"
       Case "m2": smfYahooCodeDesc = "Day's Range (ECN)"
       Case "m3": smfYahooCodeDesc = "50-day Moving Avg"
       Case "m4": smfYahooCodeDesc = "200-day Moving Avg"
       Case "m5": smfYahooCodeDesc = "Change From 200-day Moving Avg"
       Case "m6": smfYahooCodeDesc = "% off 200-day Avg"
       Case "m7": smfYahooCodeDesc = "Change From 50-day Moving Avg"
       Case "m8": smfYahooCodeDesc = "% off 50-day Avg"
       Case "n": smfYahooCodeDesc = "Name"
       Case "n": smfYahooCodeDesc = "Name of option"
       Case "o": smfYahooCodeDesc = "Open"
       Case "o1": smfYahooCodeDesc = "Open interest?"
       Case "p": smfYahooCodeDesc = "Previous Close"
       Case "p2": smfYahooCodeDesc = "Percent Change"
       Case "p3": smfYahooCodeDesc = "Type of option"
       Case "p5": smfYahooCodeDesc = "Price/Sales"
       Case "p6": smfYahooCodeDesc = "Price/Book"
       Case "q": smfYahooCodeDesc = "Ex-Dividend Date"
       Case "r": smfYahooCodeDesc = "P/E Ratio"
       Case "r1": smfYahooCodeDesc = "Dividend Pay Date"
       Case "r2": smfYahooCodeDesc = "P/E (ECN)"
       Case "r5": smfYahooCodeDesc = "PEG Ratio"
       Case "r6": smfYahooCodeDesc = "Price/EPS Est. Current Yr"
       Case "r7": smfYahooCodeDesc = "Price/EPS Est. Next Yr"
       Case "s": smfYahooCodeDesc = "Symbol"
       Case "s3": smfYahooCodeDesc = "Strike price"
       Case "s7": smfYahooCodeDesc = "Short Ratio"
       Case "t1": smfYahooCodeDesc = "Time of Last Trade"
       Case "t7": smfYahooCodeDesc = "Ticker Trend"
       Case "t8": smfYahooCodeDesc = "1yr Target Price"
       Case "v": smfYahooCodeDesc = "Volume"
       Case "v7": smfYahooCodeDesc = "Holdings Value (ECN)"
       Case "w": smfYahooCodeDesc = "52-week Range"
       Case "w4": smfYahooCodeDesc = "Day's Value Change (ECN)"
       Case "x": smfYahooCodeDesc = "Exchange"
       Case "y": smfYahooCodeDesc = "Dividend Yield"
       Case Else: smfYahooCodeDesc = "--"
       End Select

    End Function
