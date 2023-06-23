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
    ' 2023-01-17 -- Mel Pryor - testing
    ' 2023-06-23 -- Clean up code.  Remove obsolete data sources. Yahoo is the only valid selection now
    '            -- Fix Subscript out of range: https://github.com/ClimberMel/SMF_Add-in/issues/49
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for SPY:
    '
    '   =smfGetOptionExpirations("SPY", "Yahoo")
    '   --> returns next expiry date
    '
    '   Array enter it over 10 rows and it will return the next 10 Expiry Dates
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
       'Case "8", "888": GoTo Source_888
       'Case "B", "BC", "BARCHART": GoTo Source_Yahoo
       'Case "G", "GOOGLE": GoTo Source_Google
       'Case "N", "NASDAQ": GoTo Source_Yahoo
       'Case "OX": GoTo Source_OptionsXpress
       Case "Y", "YAHOO": GoTo Source_Yahoo
       Case Else
            vData(1, 1) = "Invalid Data Source: " & pSource
            GoTo ErrorExit
       End Select

    '------------------> 888options.com processing
    ' Web site has changed
    ' https://www.optionseducation.org/toolsoptionquotes/options-quotes

    
    '------------------> Google processing
    ' Web site has changed


    '------------------> OptionsXpress processing removed
    ' Web site has changed

    
    '------------------> Yahoo processing after 2017-03-15
Source_Yahoo:
    Dim vFirst As Variant, vNext As Variant
    sURL = "https://query1.finance.yahoo.com/v7/finance/options/" & UCase(pTicker)
    s1 = smfStrExtr(RCHGetWebData(sURL, """expirationDates"":[", 500), "[", "]")
    vFirst = Int(smfUnix2Date(smfWord(s1, 1, ",", 1)))
    iPtr = 0
'    For iRow = 1 To 50
    For iRow = 1 To kRows
Debug.Print iRow
        
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

