Attribute VB_Name = "modGetOptionExpirations"
Option Explicit
Function smfGetOptionExpirations(ByVal pTicker As String, _
                        Optional ByVal pSource As String = "Yahoo", _
                        Optional ByVal pPutCall As String = "X", _
                        Optional ByVal pStrike As Double = 0, _
                        Optional ByVal pRows As Integer = 0, _
                        Optional ByVal pCols As Integer = 0, _
                        Optional ByVal pType As Integer = 0, _
                        Optional ByVal pPeriod As String = "A" _
                        )
                                                      
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get available option expirations from various data sources
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.04.03 -- Created function
    ' 2011.04.07 -- Added ability to create option ticker symbols for the smfGetOptionQuotes() function
    ' 2011.11.29 -- Added OptionsXpress as possible data source
    ' 2012.01.14 -- Fix overflow on array assignment
    ' 2012.02.14 -- Added 888options.com as possible data source
    ' 2012.01.06 -- Modified Yahoo processing to pick up expiration dates from Yahoo's API feed
    ' 2014.03.15 -- Sort expiration dates
    ' 2014.03.15 -- Add period selection of expiration dates for OptionsXPress
    ' 2014.10.21 -- Modified Yahoo processing for new web page structure
    ' 2015.02.21 -- Drop period selection of expiration dates for OptionsXPress
    ' 2015.02.21 -- Modified OptionsXPress processing for new web page structure
    ' 2015.02.21 -- Backed out Yahoo processing for new web page structure
    ' 2015.08.13 -- Use Yahoo as the data source if NASDAQ is requested
    ' 2016.08.07 -- Update for new Yahoo option quotes page
    ' 2017.03.09 -- Change to use Yahoo as default
    ' 2017.03.09 -- Use Yahoo if Barchart is requested
    ' 2017.03.15 -- Modified Yahoo processing for new JSON call
    ' 2017.06.10 -- Add Vartype() check when creating ticker symbols
    ' 2017.10.10 -- optionsXpress is no longer a valid data source
    ' 2017.12.19 -- Add ability to request Yahoo expirations by type of period
    ' 2017.12.25 -- Add ability to request Yahoo expirations for multiple period types
    ' 2024.01.30 -- Invoke =RCHGetWebData with TYPE(4) param for Yahoo options url. (BS)
    '               Add "?" to end of url if no other params used so "&crumb= ..." can be processed.
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for SPY:
    '
    '   =smfGetOptionExpirations("SPY", "Google")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    
    '------------------> Determine size of array to return
    Dim kRows As Integer, kCols As Integer
    kRows = pRows
    kCols = pCols
    If pRows = 0 Or pCols = 0 Then
       If kRows = 0 Then kRows = 20
       If kCols = 0 Then kCols = 1
       On Error Resume Next
       kRows = Application.Caller.Rows.Count
       kCols = Application.Caller.Columns.Count
       On Error GoTo ErrorExit
       If kRows < 3 Then kRows = 3
       End If
  
    '------------------> Initialize return array
    Dim iRow As Integer, iCol As Integer
    ReDim vData(1 To kRows, 1 To kCols) As Variant
    For iRow = 1 To kRows: For iCol = 1 To kCols: vData(iRow, iCol) = "": Next iCol: Next iRow
    vData(1, 1) = "None"
    
    '------------------> Determine which data source to use
    Dim s1 As String, s2 As String, s3 As String, sURL As String, iPtr As Integer, d1 As Date, i1 As Integer
    Select Case UCase(pSource)
       Case "8", "888": GoTo Source_888
       Case "B", "BC", "BARCHART": GoTo Source_Yahoo
       Case "G", "GOOGLE": GoTo Source_Google
       Case "N", "NASDAQ": GoTo Source_Yahoo
       'Case "OX": GoTo Source_OptionsXpress
       Case "Y", "YAHOO": GoTo Source_Yahoo
       Case Else
            vData(1, 1) = "Invalid Data Source: " & pSource
            GoTo ErrorExit
       End Select

    '------------------> 888options.com processing
Source_888:
    sURL = "http://oic.ivolatility.com/oic_adv_options.j?exp_date=-1&ticker=" & UCase(pTicker)
    s1 = ""
    For iRow = 1 To kRows
        s2 = RCHGetTableCell(sURL, 0, s1, "Days:")
        If s2 = "Error" Then Exit For
        vData(iRow, 1) = DateValue(smfStrExtr(s2, "Expiry:", "Days:"))
        s1 = "Days:" & smfStrExtr(s2 & "|", "Days:", "|")
        Next iRow
    GoTo ExitFunction

    '------------------> Google processing
Source_Google:
    s1 = smfStrExtr(RCHGetWebData("http://www.google.com/finance/option_chain?output=json&q=" & UCase(pTicker)), "expirations:[", "]")
    For iRow = 1 To kRows
        If smfWord(s1, iRow, "}") = "" Then Exit For
        vData(iRow, 1) = DateSerial(smfStrExtr(smfWord(s1, iRow, "}"), "y:", ","), _
                                    smfStrExtr(smfWord(s1, iRow, "}"), "m:", ","), _
                                    smfStrExtr(smfWord(s1, iRow, "}") & "|", "d:", "|"))
        Next iRow
    GoTo ExitFunction


    '------------------> OptionsXpress processing
Source_OptionsXpress:
    sURL = "https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Range=All&lstMarket=0&ChainType=14&lstMonths=0&Symbol=" & UCase(pTicker)
    s1 = smfGetTagContent(sURL, "select", -1, "id=""lstMonths""")
    For iRow = 1 To kRows
        s2 = smfWord(s1, iRow + 1, "value=""")
        If s2 = "" Then Exit For
        If InStr(s2, ">All<") > 0 Then Exit For
        vData(iRow, 1) = DateValue(smfStrExtr(s2, "~", ";"))
        Next iRow
    GoTo ExitFunction

    '------------------> OptionsXpress processing prior to 2014.03.15
Source_OptionsXpress1:
    s1 = RCHGetWebData("https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Symbol=" & UCase(pTicker), ":GetOptionChain", 2000)
    d1 = smfGetOptionExpiry(, , "M")
    i1 = 0
    For iRow = 1 To kRows
        s2 = smfStrExtr(smfWord(s1, iRow, ")"), "','", ";")
        If s2 = "" Then Exit For
        i1 = i1 + 1
        If d1 < DateValue(s2) Then
           vData(i1, 1) = d1
           i1 = i1 + 1
           d1 = #12/31/2099#
           End If
        If i1 > kRows Then Exit For
        vData(i1, 1) = DateValue(s2)
        Next iRow
    GoTo ExitFunction


    '------------------> OptionsXpress processing prior to 2015-02-21
Source_OptionsXpress2:
    pPeriod = UCase(pPeriod)
    Dim iM As Integer, iKeep As Integer
    s1 = RCHGetWebData("https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Symbol=" & UCase(pTicker), ":GetOptionChain", 2000)
    If InStr("AMW", pPeriod) > 0 Then
       vData(1, 1) = smfGetOptionExpiry(, , "M")
       i1 = 1
       iM = 1
    Else
       iM = 0
       i1 = 0
       End If
    For iRow = 1 To 50
        s2 = smfWord(s1, iRow, "|")
        If s2 = "" Then Exit For
        iKeep = 1
        Select Case True
           Case InStr(s2, "strong") > 0
                iKeep = 0
           Case Left(smfStrExtr(s2, ">", "<"), 1) = "Q"
                If InStr("MW", pPeriod) > 0 Then iKeep = 0
           Case Mid(smfStrExtr(s2, ">", "<"), 4, 2) = "Wk"
                If InStr("QM", pPeriod) > 0 Then iKeep = 0
           Case Else
                iM = iM + 1
                If iM > 2 And pPeriod = "W" Then iKeep = 0
                If pPeriod = "Q" Then iKeep = 0
           End Select
        If iKeep = 1 Then
           s3 = smfStrExtr(s2, ",'", ";")
           i1 = i1 + 1
           If i1 > kRows Then Exit For
           vData(i1, 1) = DateValue(s3)
           End If
        Next iRow
    If pPeriod = "W" Then
       If vData(3, 1) = "" Then
          vData(1, 1) = ""
          vData(2, 1) = ""
          End If
       End If
    GoTo ExitFunction
    
    '------------------> Yahoo processing after 2017-03-15
Source_Yahoo:
    Dim vFirst As Variant, vNext As Variant
    sURL = "https://query1.finance.yahoo.com/v7/finance/options/" & UCase(pTicker) & "?"
    s1 = smfStrExtr(RCHGetWebData(sURL, """expirationDates"":[", 500, , 4), "[", "]")
    vFirst = Int(smfUnix2Date(smfWord(s1, 1, ",", 1)))
    iPtr = 0
    For iRow = 1 To 50
        s2 = smfWord(s1, iRow, ",")
        If s2 = "" Then Exit For
        vNext = Int(smfUnix2Date(0 + s2))
        If UCase(pPeriod) = "A" _
              Or (InStr(UCase(pPeriod), "W") > 0 And vNext <= vFirst + 43 And Weekday(vNext) = 6) _
              Or (InStr(UCase(pPeriod), "M") > 0 And Day(vNext) > 14 And Day(vNext) < 23 And Weekday(vNext) = 6) _
              Or (InStr(UCase(pPeriod), "Q") > 0 And Day(vNext) > 27 And Mid("001001001001", Month(vNext), 1) = "1") _
              Or (InStr(UCase(pPeriod), "H") > 0 And vNext <= vFirst + 43 And Weekday(vNext) = 4) _
              Or (InStr(UCase(pPeriod), "V") > 0 And Day(vNext) > 14 And Day(vNext) < 23 And Weekday(vNext) = 4) _
              Then
           iPtr = iPtr + 1
           vData(iPtr, 1) = vNext
           End If
        Next iRow
    
    GoTo ExitFunction

    '------------------> Yahoo processing prior to 2012.01.06 and after 2015-02-21
Source_Yahoo0:
    sURL = "http://finance.yahoo.com/q/op?s=" & UCase(pTicker)
    s1 = smfGetTagContent(sURL, "td", -1, "Expiration:")
    iPtr = 0
    
    For iRow = 1 To 7
        s2 = smfGetOptionExpiry(, , "W" & iRow)
        If RCHGetWebData(sURL & "&m=" & Format(s2, "yyyy-mm"), pTicker & Format(s2, "yymmdd"), 5) <> "Error" Then
           iPtr = iPtr + 1
           vData(iPtr, 1) = DateValue(s2)
           End If
        Next iRow
    
    For iRow = 2 To kRows
        s2 = smfWord(s1, iRow, "&m=")
        If s2 = "" Then Exit For
        d1 = smfGetOptionExpiry(Left(s2, 4), Mid(s2, 6, 2), "M")
        If d1 > vData(iPtr, 1) Then
           iPtr = iPtr + 1
           If iPtr > kRows Then Exit For
           vData(iPtr, 1) = d1
           Select Case Mid(s2, 6, 2)
              Case "03", "06", "09", "12"
                   s3 = RCHGetWebData(sURL & "&m=" & Left(s2, 7), pTicker & Mid(s2, 3, 2) & Mid(s2, 6, 2) & "3", 12)
                   If s3 <> "Error" Then
                      s3 = "20" & Mid(s3, Len(pTicker) + 1, 6)
                      iPtr = iPtr + 1
                      If iPtr > kRows Then Exit For
                      vData(iPtr, 1) = DateSerial(Mid(s3, 1, 4), Mid(s3, 5, 2), Mid(s3, 7, 2))
                      End If
              End Select
           End If
        Next iRow
    
    GoTo ExitFunction
    
    '------------------> Yahoo processing for new format page that was backed out
Source_Yahoo1:
    sURL = "http://finance.yahoo.com/q/op?s=" & UCase(pTicker)
    iRow = 0
    On Error GoTo ExitFunction
    For i1 = 1 To 1500 ' Arbitrary value to cover 4 years?
        s1 = smfGetTagContent(sURL, "option", i1, "class=""SelectBox-Pick""")
        If s1 <> "Error" Then
           d1 = DateValue(s1)
           d1 = DateSerial(Year(d1), Month(d1), Day(d1))
           iRow = iRow + 1
           vData(iRow, 1) = d1
           End If
        Next i1
    GoTo ExitFunction

ExitFunction:
   
    '------------------> Sort expiration dates
    Dim iRow2 As Integer
    Dim vTemp As Variant
    For iRow = 1 To kRows - 1
        If vData(iRow, 1) = "None" Then Exit For
        For iRow2 = iRow + 1 To kRows
            If vData(iRow, 1) > vData(iRow2, 1) Then
               vTemp = vData(iRow, 1)
               vData(iRow, 1) = vData(iRow2, 1)
               vData(iRow2, 1) = vTemp
               End If
            Next iRow2
        Next iRow

    '------------------> Convert expiration dates into option ticker symbols
    If Left(UCase(pPutCall), 1) = "P" Or Left(UCase(pPutCall), 1) = "C" Then
       For iRow = 1 To kRows
           If IsDate(vData(iRow, 1)) Or VarType(vData(iRow, 1)) = vbDouble Then
              vData(iRow, 1) = UCase(pTicker) & _
                               Format(vData(iRow, 1), " m/d yyyy ") & _
                               Format(pStrike, "$0.00 ") & _
                               IIf(Left(UCase(pPutCall), 1) = "P", "Put", "Call")
              End If
           Next iRow
       End If

ErrorExit:

    smfGetOptionExpirations = vData
                        
    End Function

