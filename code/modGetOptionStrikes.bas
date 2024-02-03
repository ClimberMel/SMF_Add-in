Attribute VB_Name = "modGetOptionStrikes"
Option Explicit
Function smfGetOptionStrikes(ByVal pTicker As String, _
                    Optional ByVal pExpiry As Variant = 0, _
                    Optional ByVal pPutCall As String = "P", _
                    Optional ByVal pSource As String = "Y", _
                    Optional ByVal pSymbols As Integer = 0, _
                    Optional ByVal pRows As Integer = 0, _
                    Optional ByVal pCols As Integer = 0, _
                    Optional ByVal pType As Integer = 0)
                                                      
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get available option strikes from various data sources
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.04.04 -- Created function
    ' 2011.11.29 -- Added OptionsXpress as possible data source
    ' 2011.11.30 -- Fix URL to get all strike prices for OptionsXpress
    ' 2014.10.21 -- Modified Yahoo processing for new web page structure
    ' 2015.02.21 -- Restored Yahoo processing to old web page structure
    ' 2015.08.13 -- Use Yahoo as the data source if NASDAQ is requested
    ' 2017.03.09 -- Change to use Yahoo as default
    ' 2017.03.09 -- Use Yahoo if Barchart is requested
    ' 2017.03.15 -- Modified Yahoo processing for new JSON call
    ' 2017.04.26 -- Change "http://" protocol to "https://" for Yahoo
    ' 2017.07.05 -- Modified Yahoo processing to use smfGetYahooJSONField() to get data
    ' 2017.07.09 -- Removed smfGetYahooJSONField() to get data, as it was not reliable
    ' 2017.10.10 -- optionsXpress is no longer a valid data source
    ' 2017.10.12 -- Try Yahoo processing using smfGetYahooJSONField() to get data again
    ' 2017.10.14 -- Removed smfGetYahooJSONField() to get data, as it was still not reliable with this JSON file
    ' 2017.10.24 -- Add default value of 0 for pExpiry so first available date is used for Yahoo calls
    ' 2024.01.30 -- Invoke =RCHGetWebData with TYPE(4) param for Yahoo options url. (BS)
    '               Add "?" to end of url if no other params used so "&crumb= ..." can can be processed.
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for SPY:
    '
    '   =smfGetOptionStrikes("SPY", smfGetOptionExpiry(), "P", "Google")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    Dim dDate As Double
    
    '------------------> Determine size of array to return
    Dim kRows As Integer, kCols As Integer
    kRows = pRows
    kCols = pCols
    If pRows = 0 Or pCols = 0 Then
       If kRows = 0 Then kRows = 40
       If kCols = 0 Then kCols = 1
       On Error Resume Next
       kRows = Application.Caller.Rows.Count
       kCols = Application.Caller.Columns.Count
       On Error GoTo ErrorExit
       End If
  
    '------------------> Initialize return array
    Dim iRow As Integer, iCol As Integer
    Dim vData0(1 To 500)
    ReDim vData(1 To kRows, 1 To kCols) As Variant
    For iRow = 1 To kRows: For iCol = 1 To kCols: vData(iRow, iCol) = "": Next iCol: Next iRow
    vData(1, 1) = "None"
    
    '------------------> Verify Put/Call and expiration date parameters
    Dim sPutCall As String
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P"
       Case "C"
       Case Else
            vData(1, 1) = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            GoTo ErrorExit
       End Select
    
    Select Case True
       Case pExpiry = 0 And pSource = "Y"
       Case VarType(pExpiry) = vbDouble
       Case IsDate(pExpiry)
       Case Else
            vData(1, 1) = "Bad expiration date: " & pExpiry
            GoTo ErrorExit
       End Select
    
    '------------------> Determine which data source to use
    Dim s1 As String, s2 As String, s3 As String, sURL As String
    Dim i1 As Integer, iPtr As Integer
    Dim nPrice As Double
    Select Case UCase(pSource)
       Case "8", "888": GoTo Source_888
       Case "B", "BC", "BARCHART": GoTo Source_Yahoo
       Case "G", "GOOGLE": GoTo Source_Google
       Case "N", "NASDAQ": GoTo Source_Yahoo
       Case "Y", "YAHOO": GoTo Source_Yahoo
       Case Else
            vData(1, 1) = "Invalid data source: " & pSource
            GoTo ErrorExit
       End Select

    '------------------> Google processing
Source_888:
    sURL = "http://oic.ivolatility.com/oic_adv_options.j?exp_date=-1&ticker=" & UCase(pTicker)
    nPrice = RCHGetYahooQuotes(pTicker, "l1")(1, 1)
    s1 = Format(pExpiry, "yymmdd") & "C"
    i1 = -1
    iPtr = 0
    For iRow = 1 To 500
        s2 = smfStrExtr(smfGetTagContent(sURL, "tr", i1, s1), "1"">", "<")
        If s2 = "" Then Exit For
        vData0(iRow) = 0 + s2
        If iPtr = 0 And vData0(iRow) > nPrice Then iPtr = iRow   ' First ITM price
        i1 = 2 * iRow
        Next iRow
    GoTo ExitFunction

    '------------------> Google processing
Source_Google:
    Dim sTicker As String
    If InStr(pTicker, ":") > 0 Then sTicker = smfStrExtr(pTicker & "|", ":", "|") Else sTicker = pTicker
    nPrice = RCHGetYahooQuotes(sTicker, "l1")(1, 1)
    sURL = "http://www.google.com/finance/option_chain?output=json&q=" & UCase(pTicker) & _
           "&expd=" & Day(pExpiry) & _
           "&expm=" & Month(pExpiry) & _
           "&expy=" & Year(pExpiry)
    s1 = IIf(sPutCall = "C", "calls:[", "puts:[")
    s2 = smfStrExtr(RCHGetWebData(sURL, s1), s1, "]")
    iPtr = 0
    For iRow = 1 To 500
        s3 = smfStrExtr(smfWord(s2, iRow, "}"), "strike:""", """")
        If s3 = "" Then Exit For
        vData0(iRow) = 0 + s3
        If iPtr = 0 And vData0(iRow) > nPrice Then iPtr = iRow   ' First ITM price
        Next iRow
    GoTo ExitFunction
    
    '------------------> OptionsXpress processing
Source_OptionsXpress:
    sURL = "https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Range=0&lstMarket=0&Symbol=" & UCase(pTicker) & _
           "&lstMonths=" & Format(pExpiry, "mm/dd/yyyy") & ";7"
    s1 = Format(pExpiry, "yymmdd") & "C"
    iPtr = 0
    nPrice = RCHGetYahooQuotes(pTicker, "l1")(1, 1)
    For iRow = 1 To 500
        s2 = smfStrExtr(smfGetTagContent(sURL, "tr", iRow - 1, s1), s1, "&")
        If s2 = "" Then Exit For
        vData0(iRow) = s2 / 1000
        If iPtr = 0 And vData0(iRow) > nPrice Then iPtr = iRow   ' First ITM price
        Next iRow
    GoTo ExitFunction

    '------------------> Yahoo processing from 2017.10.14 on
Source_Yahoo:
    Dim iPos1 As Double, iPos2 As Double, iPos3 As Double
    If pExpiry = 0 Then
       s1 = "?"   ' Use first available expiration date
    Else
       dDate = 86400 * (DateSerial(Year(pExpiry), Month(pExpiry), Day(pExpiry)) - DateSerial(1970, 1, 1))
       s1 = "?date=" & dDate
       End If
    sURL = "https://query1.finance.yahoo.com/v7/finance/options/" & UCase(pTicker) & s1
    nPrice = smfStrExtr(RCHGetWebData(sURL, "regularMarketPrice", 100, , 4), ":", ",", 1)
    If pExpiry = 0 Then pExpiry = smfUnix2Date(smfStrExtr(RCHGetWebData(sURL, """expirationDates"":", 100, , 4), "[", ",", 1))
    iPtr = 0  ' Pointer for first ITM price
    i1 = 0
    iPos1 = 1
    s1 = RCHGetWebData(sURL, iPos1, , , 4)
    For iRow = 1 To 500
        iPos2 = InStr(2, s1, """puts"":")
        iPos3 = InStr(2, s1, """strike"":")
        If iPos3 = 0 Or (iPos2 > 0 And iPos2 < iPos3) Then Exit For
        iPos1 = iPos1 + iPos3
        s1 = RCHGetWebData(sURL, iPos1, , , 4)
        vData0(iRow) = smfStrExtr(s1, ":", ",", 1)
        If iPtr = 0 And vData0(iRow) > nPrice Then iPtr = iRow   ' First ITM price
        Next iRow
    GoTo ExitFunction
    
    '------------------> Yahoo processing from 2017.10.12 to 2017.10.13
Source_Yahoo5:
    dDate = 86400 * (DateSerial(Year(pExpiry), Month(pExpiry), Day(pExpiry)) - DateSerial(1970, 1, 1))
    sURL = "https://query1.finance.yahoo.com/v7/finance/options/" & UCase(pTicker) & "?date=" & dDate
    nPrice = smfGetYahooJSONField(pTicker, sURL, "optionChain.result.0.quote.regularMarketPrice")
    iPtr = 0
    i1 = 0
    For iRow = 1 To 500
        s3 = smfGetYahooJSONField(pTicker, sURL, "optionChain.result.0.options.0.calls." & (iRow - 1) & ".strike")
        If s3 = "Not Found" Then Exit For
        vData0(iRow) = smfConvertData(s3)
        If iPtr = 0 And s3 > nPrice Then iPtr = iRow   ' First ITM price
        Next iRow
    GoTo ExitFunction

    '------------------> Yahoo processing from 2017.03.15 to 2017.10.12
Source_Yahoo4:
    dDate = 86400 * (DateSerial(Year(pExpiry), Month(pExpiry), Day(pExpiry)) - DateSerial(1970, 1, 1))
    sURL = "https://query1.finance.yahoo.com/v7/finance/options/" & UCase(pTicker) & "?date=" & dDate
    nPrice = smfConvertData(smfStrExtr(RCHGetWebData(sURL, "regularMarketPrice"":", 100), ":", ","))
    iPtr = 0
    s1 = smfStrExtr(RCHGetWebData(sURL, """strikes"":[", 1000), "[", "]")
    i1 = 0
    For iRow = 1 To 500
        s3 = smfWord(s1, iRow, ",")
        If s3 = "" Then Exit For
        vData0(iRow) = s3
        If iPtr = 0 And s3 > nPrice Then iPtr = iRow   ' First ITM price
        Next iRow
    GoTo ExitFunction

    '------------------> Yahoo processing prior to 2014.10.21 and from 2015.02.01 to 2017-03-14
Source_Yahoo3:
    sURL = "https://finance.yahoo.com/q/op?s=" & UCase(pTicker) & "&m=" & Format(pExpiry, "yyyy-mm")
    nPrice = RCHGetYahooQuotes(pTicker, "l1")(1, 1)
    iPtr = 0
    s1 = UCase(pTicker) & Format(pExpiry, "yymmdd") & sPutCall
    s2 = IIf(sPutCall = "P", ">Put Options", ">Call Options")
    i1 = 0
    For iRow = 1 To 500
        s3 = smfGetTagContent(sURL, "a", 2 * iRow, s2)
        If Left(Right(s3, 9), 1) <> sPutCall Then Exit For
        If Left(s3, Len(s1)) = s1 Then
           i1 = i1 + 1
           vData0(i1) = Right(s3, 8) / 1000
           If iPtr = 0 And vData0(i1) > nPrice Then iPtr = i1   ' First ITM price
           End If
        Next iRow
    GoTo ExitFunction

    '------------------> Yahoo processing from 2014.10.21 thru 2015.02.21
Source_Yahoo2:
    dDate = 86400 * (DateSerial(Year(pExpiry), Month(pExpiry), Day(pExpiry)) - DateSerial(1970, 1, 1))
    sURL = "http://finance.yahoo.com/q/op?s=" & UCase(pTicker) & "&date=" & dDate
    
    nPrice = RCHGetYahooQuotes(pTicker, "l1")(1, 1)
    iPtr = 0
    s1 = UCase(pTicker) & Format(pExpiry, "yymmdd") & sPutCall
    s2 = IIf(sPutCall = "P", """optionsPuts""", """optionsCalls""")
    i1 = 0
    For iRow = 1 To 500
        s3 = smfGetTagContent(sURL, "a", 2 * iRow + 1, s2)
        If Left(Right(s3, 9), 1) <> sPutCall Then Exit For
        If Left(s3, Len(s1)) = s1 Then
           i1 = i1 + 1
           vData0(i1) = Right(s3, 8) / 1000
           If iPtr = 0 And vData0(i1) > nPrice Then iPtr = i1   ' First ITM price
           End If
        Next iRow
    GoTo ExitFunction

    '------------------> Yahoo processing from 2016.08.07 to 2017-03-14
Source_Yahoo1:
    dDate = 86400 * (DateSerial(Year(pExpiry), Month(pExpiry), Day(pExpiry)) - DateSerial(1970, 1, 1))
    sURL = "http://finance.yahoo.com/quote/" & UCase(pTicker) & "/options?date=" & dDate
    nPrice = RCHGetYahooQuotes(pTicker, "l1")(1, 1)
    iPtr = 0
    s1 = UCase(pTicker) & Format(pExpiry, "yymmdd") & sPutCall
    s2 = IIf(sPutCall = "P", ">Puts", ">Calls")
    i1 = 0
    For iRow = 1 To 500
        s3 = smfGetTagContent(sURL, "a", 2 * iRow, s2)
        If Left(Right(s3, 9), 1) <> sPutCall Then Exit For
        If Left(s3, Len(s1)) = s1 Then
           i1 = i1 + 1
           vData0(i1) = Right(s3, 8) / 1000
           If iPtr = 0 And vData0(i1) > nPrice Then iPtr = i1   ' First ITM price
           End If
        Next iRow
    GoTo ExitFunction

ExitFunction:
    '------------------> Extract requested range of strike prices from full list found
    vData(1, 1) = ""
    For iRow = 1 To kRows
        i1 = iPtr - Int(kRows / 2) + iRow - 1
        Select Case True
           Case i1 < 1
           Case vData0(i1) = Empty: Exit For
           Case pSymbols = 1
                vData(iRow, 1) = UCase(pTicker) & _
                                 Format(pExpiry, " m/d yyyy ") & _
                                 Format(vData0(i1), "$0.00 ") & _
                                 IIf(sPutCall = "P", "Put", "Call")
           Case Else
                vData(iRow, 1) = vData0(i1)
           End Select
        Next iRow

ErrorExit:

    smfGetOptionStrikes = vData
                        
    End Function



