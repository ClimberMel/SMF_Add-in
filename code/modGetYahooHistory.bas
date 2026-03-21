Attribute VB_Name = "smfGetYahooHistory_"

Public Function smfGetYahooHistory(ByVal pTicker As String, _
                          Optional ByVal pStartDate As Variant = "", _
                          Optional ByVal pEndDate As Variant = "", _
                          Optional ByVal pPeriod As Variant = "d", _
                          Optional ByVal pItems As String = "dohlcav", _
                          Optional ByVal pHeaders As Integer = 1, _
                          Optional ByVal pReSort As Integer = 1, _
                          Optional ByVal pRowNbrs As Integer = 0, _
                          Optional ByVal pColNbrs As Integer = 0)
                         
    '============================================================================================================
    ' User defined function to download historical quotes from Yahoo!
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.05.18 -- Originally written by Randy Harmelink (rharmelink@gmail.com)
    ' 2025.09.30 -- Rewritten by Bruce S.
    '            -- Rewrite due to Yahoo redesigning their financial web pages.
    '               Data is now in JSON structures.
    '-----------------------------------------------------------------------------------------------------------*
    ' --> Example of an invocation to get daily quotes for AAPL:
    '
    '   =smfGetYahooHistory("AAPL","1/1/2025","5/18/2025","d")
    '============================================================================================================
    
    Dim sURL As String, sPeriod As Variant
    Dim vSelectedItems As String
    Dim iHeaderRowOffset As Integer, iColumnPos As Integer
    Dim iRowCount As Integer, iColCount As Integer
    Dim i As Integer, j As Integer
    
    On Error GoTo ErrorExit
    
    ' =======================================================================
    ' ---> Set defaults, if necessary
    ' =======================================================================
    If pPeriod = "" Then pPeriod = "d"
    If pItems = "" Then pItems = "dohlcav"
'    If pHeaders = "" Then pHeaders = 1
'    If pReSort = "" Then pReSort = 1
     
    ' =======================================================================
    ' ---> Determine size of array to return
    ' =======================================================================
    Dim kRows As Integer, kCols As Integer
    kRows = pRowNbrs
    kCols = pColNbrs
    If pRowNbrs = 0 Or pColNbrs = 0 Then
       If kRows = 0 Then kRows = 12000
       If kCols = 0 Then kCols = 8
       On Error Resume Next
       kRows = Application.Caller.Rows.Count
       kCols = Application.Caller.Columns.Count
       On Error GoTo ErrorExit
       If kRows < 3 Then kRows = 3
    End If
    
    ' =======================================================================
    ' ---> Initialize output array
    ' =======================================================================
    
    ReDim vData(1 To kRows, 1 To kCols) As Variant
    For i = 1 To kRows: For j = 1 To kCols: vData(i, j) = "": Next j: Next i
   
    ' =======================================================================
    ' --> Validate "Headers" and "ReSort" values
    ' =======================================================================
   
     vData(1, 1) = "Error: Invalid Headers value - " & pHeaders
     Select Case pHeaders
           Case 0, 1
           Case Else
                GoTo ErrorExit
     End Select
 
     vData(1, 1) = "Error: Invalid ReSort value - " & pReSort
     Select Case pReSort
           Case 0, 1
           Case Else
                GoTo ErrorExit
     End Select
          
    ' =======================================================================
    ' --> Validate and convert dates to Unix UTC
    ' =======================================================================
    Dim dUnixStartDate As Double
    Dim dUnixEndDate As Double
   
    vData(1, 1) = "Error on starting date: " & pStartDate
    Select Case True
          Case VarType(pStartDate) = vbDate Or VarType(pStartDate) = vbDouble   'vbDate = 7 so if VarType(pStartDate) = 7 then pStartDate is a Date)
               dUnixStartDate = smfDate2Unix(pStartDate)
          Case pStartDate = ""
               dUnixStartDate = smfDate2Unix(DateValue("1/1/2000"))
          Case Else
               dUnixStartDate = smfDate2Unix(DateValue(pStartDate))
          End Select
    
    vData(1, 1) = "Error on ending date: " & pEndDate
    Select Case True
          Case VarType(pEndDate) = vbDate Or VarType(pEndDate) = vbDouble
               dUnixEndDate = smfDate2Unix(Int(pEndDate) + 1)               ' Add one to the end date
          Case pEndDate = ""
               dUnixEndDate = smfDate2Unix(Int(Now) + 1)
          Case Else
               dUnixEndDate = smfDate2Unix(Int(DateValue(pEndDate)) + 1)
    End Select
                
    If dUnixStartDate > dUnixEndDate Then
        vData(1, 1) = "Error: Starting date (" & pStartDate & ") is after ending date (" & pEndDate & ")"
        GoTo ErrorExit
    End If
     
    ' =======================================================================
    ' --> Convert Period value for url
    ' =======================================================================
    sPeriod = pPeriod
    Select Case LCase(pPeriod)
        Case "d": sPeriod = "1d"
        Case "w": sPeriod = "1wk"
        Case "m": sPeriod = "1mo"
        Case "q": sPeriod = "3mo"
        Case Else:   vData(1, 1) = "Invalid Period: " & pPeriod: GoTo ErrorExit
    End Select
     
          
    ' =======================================================================
    ' Build the Yahoo URL and get the JSON data page
    ' =======================================================================
    sURL = "https://query2.finance.yahoo.com/v8/finance/chart/" & UCase(pTicker) & _
           "?events=capitalGain%7Cdiv%7Csplit" & _
           "&formatted=true&includeAdjustedClose = true" & _
           "&period1=" & dUnixStartDate & _
           "&period2=" & dUnixEndDate & _
           "&interval=" & sPeriod & "&symbol=" & UCase(pTicker)
    
     s1 = RCHGetURLData(sURL, 4)             ' Get the JSON data page
     vData(1, 1) = "None"

    ' =======================================================================
    ' Parse the JSON data page
    ' =======================================================================
     Dim oJsonPage As Object
     Dim oQuote As Object, oTimestamps As Object, oAdjCloseData As Object
     Dim oColMap As Object
       
     Set oJsonPage = ParseJson(s1)
     Set oQuote = oJsonPage("chart")("result")(1)("indicators")("quote")(1)
     Set oTimestamps = oJsonPage("chart")("result")(1)("timestamp")

    ' =======================================================================
    ' Check if Adjusted Close data is available
    ' =======================================================================
     If Not IsEmpty(oJsonPage("chart")("result")(1)("indicators")("adjclose")) Then
        Set oAdjCloseData = oJsonPage("chart")("result")(1)("indicators")("adjclose")(1)("adjclose")
     End If

     vSelectedItems = LCase(pItems)
     iRowCount = Application.WorksheetFunction.Min(kRows, oQuote("open").Count + IIf(pHeaders = 1, 1, 0))
     iColCount = Len(vSelectedItems)
   
    ' =======================================================================
    ' Add column headings to row 1 if pHeaders is 1
    ' =======================================================================
    If pHeaders = 1 Then
        '-------------------------------------------------------------------
        '   Define the column labels
        '-------------------------------------------------------------------
        Set oColMap = CreateObject("Scripting.Dictionary")
        oColMap.Add "d", "Date"
        oColMap.Add "o", "Open"
        oColMap.Add "h", "High"
        oColMap.Add "l", "Low"
        oColMap.Add "c", "Close"
        oColMap.Add "a", "Adj Close"
        oColMap.Add "v", "Volume"
        oColMap.Add "t", "Ticker"
        
        iColumnPos = 1
        For i = 1 To iColCount
            vData(1, iColumnPos) = oColMap(Mid(vSelectedItems, i, 1))
            iColumnPos = iColumnPos + 1
        Next i
    End If

    ' =======================================================================
    ' --> Build the output array based on items selected
    ' =======================================================================
    Dim iDataRowPtr As Integer
    iDataRowPtr = 0
    
    iHeaderRowOffset = IIf(pHeaders = 1, 1, 0)             ' Start at row 2 if using headers
    
    For i = 1 To iRowCount - iHeaderRowOffset
        iColumnPos = 1
        
        ' -------------------------------------------------------------------------------
        ' --> Need the "If .. " stmt bcuz option tickers could have empty JSON fields
        ' -------------------------------------------------------------------------------
        If Not IsNumeric(oQuote("open")(i)) Then     ' Do nothing if no data
        Else
            iDataRowPtr = iDataRowPtr + 1
        
            For j = 1 To iColCount
                Select Case Mid(vSelectedItems, j, 1)
                    Case "d"                                            ' Convert Unix date to Excel format
                          vData(iDataRowPtr + iHeaderRowOffset, iColumnPos) = smfUnix2Date(oTimestamps(i))
                    Case "o"
                          vData(iDataRowPtr + iHeaderRowOffset, iColumnPos) = oQuote("open")(i) + 0       ' Add zero to convert to number
                    Case "h"
                          vData(iDataRowPtr + iHeaderRowOffset, iColumnPos) = oQuote("high")(i) + 0
                    Case "l"
                          vData(iDataRowPtr + iHeaderRowOffset, iColumnPos) = oQuote("low")(i) + 0
                    Case "c"
                          vData(iDataRowPtr + iHeaderRowOffset, iColumnPos) = oQuote("close")(i) + 0
                    Case "a"
                        If Not IsEmpty(oAdjCloseData) Then
                          vData(iDataRowPtr + iHeaderRowOffset, iColumnPos) = oAdjCloseData(i) + 0
                        End If
                    Case "v"
                          vData(iDataRowPtr + iHeaderRowOffset, iColumnPos) = oQuote("volume")(i) + 0
                    Case "t"
                          vData(iDataRowPtr + iHeaderRowOffset, iColumnPos) = UCase(pTicker)
                 End Select
                iColumnPos = iColumnPos + 1
            Next j
            
        End If      ' ... has data in row
    Next i
    
    ' =======================================================================
    '   If pReSort is 1, reverse the order of data                 =========
    ' =======================================================================

    If pReSort = 1 And IsArray(vData) And Not IsEmpty(vData) Then
        vData = smfGetYHReverse(vData, pHeaders, iDataRowPtr + iHeaderRowOffset, iColCount, kRows, kCols)
    End If

ErrorExit:

    smfGetYahooHistory = vData
    
End Function

    ' =======================================================================
    '   Function to reverse the order of data extracted.
    ' =======================================================================
    
Function smfGetYHReverse(pArray As Variant, _
                        pHeaders As Integer, _
                        pRowCount As Integer, _
                        pColCount As Integer, _
                        pRowsMax As Integer, _
                        pColSMax As Integer) As Variant
    
    Dim vTempArray() As Variant
    Dim vRowCount As Integer, vColCount As Integer
    Dim i As Integer, j As Integer
       
    vRowCount = pRowCount
    vColCount = pColCount
    
    ' =======================================================================
    ' ---> Initialize temp array
    ' =======================================================================
    ReDim vTempArray(1 To pRowsMax, 1 To pColSMax) As Variant
    
    For i = 1 To pRowsMax
        For j = 1 To pColSMax
            vTempArray(i, j) = ""
        Next j
    Next i
        
    ' =======================================================================
    '    Reverse the rows
    ' =======================================================================
    
    If pHeaders = 1 Then
        '
        ' --- Copy the first line (column headings)  ---
        '
        For j = 1 To vColCount
            vTempArray(1, j) = pArray(1, j)       ' Copy the column headings
        Next j
        
        '
        ' --- Reverse list from the second line onward ---
        '
        For i = 2 To vRowCount
            For j = 1 To vColCount
                vTempArray(i, j) = pArray(vRowCount - i + 2, j)
            Next j
        Next i
    Else
        '
        ' If pHeaders is 0, reverse *all* rows.
        '
        For i = 1 To vRowCount
            For j = 1 To vColCount
                vTempArray(i, j) = pArray(vRowCount - i + 1, j)
            Next j
        Next i
    End If
    
    smfGetYHReverse = vTempArray
    
End Function

    '-----------------------------------------------------------------------------------------------------------*
    ' *** OBSOLETE *** 10/01/25
    '-----------------------------------------------------------------------------------------------------------*
Function smfGetYahooHistory1(ByVal pTicker As String, _
                   Optional ByVal pStartDate As Variant = "", _
                   Optional ByVal pEndDate As Variant = "", _
                   Optional ByVal pPeriod As String = "d", _
                   Optional ByVal pItems As String = "dohlcvu", _
                   Optional ByVal pNames As Integer = 1, _
                   Optional ByVal pReSort As Integer = 0, _
                   Optional ByVal pRows As Integer = 0, _
                   Optional ByVal pCols As Integer = 0)
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to download historical quotes from Yahoo!
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.05.18 -- Written by Randy Harmelink (rharmelink@gmail.com)
    ' 2017.05.19 -- Set minimum rows to 2 (for processing, not returning)
    ' 2017.05.19 -- Add varType() function to cumulative dividend processing
    ' 2017.05.20 -- Set defaults on pRows and pCols for VBA calls, with no range involved
    ' 2017.05.21 -- Change default starting date to "1/1/1970"
    ' 2017.05.21 -- Change "null" values to zeroes and backfill zeroed values
    ' 2017.05.29 -- Fix sorting of split or dividend requests
    ' 2017.05.30 -- Change to use smfGetWebPage() instead of RCHGetURLData(), to remove redundant retrievals
    ' 2017.06.09 -- Remove calculated dividend adjustments, as Yahoo appears to be doing them now
    ' 2017.07.12 -- Add back in adjustments for O/H/L amounts, get adjusted close and close
    ' 2023-01-22 -- Mel Pryor (ClimberMel@gmail.com)
    ' 2023-02-08 -- Fixed issues with module trying to scrape json data when Yahoo is now csv data
    '               It now uses RCHGetURLData() as it works well with CSV data
    ' 2023-02-15 -- Returned data seems to default to oldest to newest, resort=1 gives newest to oldest
    '               Returned values are string not numbers.  Should be float.   Issue #10
    '               Changed to call smfCDec to convert them
    '               Changed pStartDate to default to 1950 instead of 1970
    '               Fixed Issue# 16 with data retreived from start date
    ' 2023-02-21 -- The original code for smfGetYahooHistory used C for Adjusted Close and U for Unadjusted close
    '               A is now accepted but changed to C
    '               RCHGetYahooHistory in modGetYahooHistory has been updated to match
    '               with the use of U for unadjusted close, C & A for adjusted close
    '-----------------------------------------------------------------------------------------------------------*

    '-----------------------------------------------------------------------------------------------------------*
    ' > Example of an invocation to get daily quotes for 2017 for IBM:
    '
    '   =smfGetYahooHistory("IBM","1/1/2017","5/18/2017","d")
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim s1 As String, sURL As String, sData As String, sFind As String, sFound As String
    Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer
    Dim iRows As Integer, iCols As Integer, iRow As Integer
    Dim vDivAmt As Variant, vNum As Variant, vDen As Variant
    Dim dAdj As Double, d1 As Double, d2 As Double, dSplitAdj As Double
    Dim vByDay As Variant
    Dim vItem As Variant
    ReDim vData(1 To 1, 1 To 1) As Variant
    
    vData(1, 1) = "Error"
    
    On Error GoTo ErrorExit
    
    '------------------> Set defaults, if necessary
    If pPeriod = "" Then pPeriod = "d"
    If pItems = "" Then pItems = "dohlcvu"
    
    '------------------> Null Return Item
    If pTicker = "None" Or pTicker = "" Then
       vData(1, 1) = "None"
       GoTo ErrorExit
       End If
       
    '------------------> Verify and Process starting and ending dates
    Dim dBegin As Double, dEnd As Double
    vData(1, 1) = "Error on starting date: " & pStartDate
    Select Case True
          Case VarType(pStartDate) = vbDate Or VarType(pStartDate) = vbDouble   'vbDate = 7 so if VarType(pStartDate) = 7 then pStartDate is a Date)
               dBegin = smfDate2Unix(pStartDate)
          Case pStartDate = ""
               dBegin = smfDate2Unix(DateValue("1/1/1950"))
          Case Else
               dBegin = smfDate2Unix(DateValue(pStartDate))
          End Select
    vData(1, 1) = "Error on ending date: " & pEndDate
    Select Case True
          Case VarType(pEndDate) = vbDate Or VarType(pEndDate) = vbDouble
               dEnd = smfDate2Unix(Int(pEndDate) + 1)
          Case pEndDate = ""
               dEnd = smfDate2Unix(Int(Now) + 1)
          Case Else
               dEnd = smfDate2Unix(Int(DateValue(pEndDate)) + 1)
          End Select
     If dBegin > dEnd Then
        vData(1, 1) = "Error: Starting date cannot be after ending date: " & pStartDate & "," & pEndDate
        GoTo ErrorExit
        End If
    
    '------------------> Determine size of array to return
    iRows = pRows  ' Rows
    iCols = pCols  ' Columns
    If pRows = 0 Or pCols = 0 Then
       On Error Resume Next
       iRows = Application.Caller.Rows.Count
       iCols = Application.Caller.Columns.Count
       On Error GoTo ErrorExit
       End If
    If iRows = 0 Then iRows = Int(smfUnix2Date(Int(dEnd)) - smfUnix2Date(Int(dBegin))) + 2
    If iCols = 0 Then iCols = Len(pItems) + 1
    If iRows = 1 Then iRows = 2
  
    '------------------> Initialize return array
    ReDim vData(1 To iRows, 1 To iCols) As Variant
    For i1 = 1 To iRows
        For i2 = 1 To iCols
            vData(i1, i2) = ""
            Next i2
        Next i1
    vData(1, 1) = "Error"
       
    '------------------> Process period
    Dim sPeriod As String, sFreq As String, sfilter As String, sInterval As String
    vData(1, 1) = "Error on period: " & pPeriod
    sPeriod = UCase(pPeriod)
    Select Case sPeriod
       Case "D": sFreq = "1d": sfilter = "history": sInterval = "1d"
       Case "W": sFreq = "1wk": sfilter = "history": sInterval = "1wk"
       Case "A", "Q", "M": sFreq = "1mo": sfilter = "history": sInterval = "1mo"
       Case "V": sFreq = "1d": sfilter = "div": sInterval = "1d"
       Case "S": sFreq = "1d": sfilter = "split": sInterval = "1d"
       Case Else: GoTo ErrorExit
       End Select
    vData(1, 1) = "Error"
    
    '------------------> Verify and Process pItems parameter
    '  aItems will contain a list of the pItems requested and the order for the columns to return the data to
    
    'Const kItemList As String = "Ticker,Date,Open,High,Low,Close,Volume,Unadj,Div Adj,Split Adj,Dividend,Split"
    'Yahoo provides split & div adjusted close, so I'm not going to try to manually calculate it for both splits and dividends
   
   Dim currYear As Integer
   
   Dim sItems As String, aItems(1 To 10) As Integer
   
   sItems = UCase(pItems)
   
   Select Case True
      ' Note that C is for Adjusted Close so if C was requested, there is no point in having A as well
      ' U would be used to request Unadjusted close
      Case InStr(sItems, "C") > 0: sItems = Replace(sItems, "A", "")
      ' If A was requested for adjusted close, replace with C as that is used to call the adjusted close
      Case Else: sItems = Replace(sItems, "A", "C")
      End Select

    Const kItemList As String = "Ticker,Date,Open,High,Low,Close,AdjClose,Volume,Dividend,Split"
    For i1 = 1 To 10: aItems(i1) = 0: Next i1
    sItems = UCase(pItems)
    Select Case sPeriod                                     ' sPeriod is Day, Week, Month for historical data or V for Dividends or S for Splits
       Case "V"                                             ' Just bring back Date and Dividend
            If InStr(sItems, "T") > 0 Then aItems(1) = 1    ' include Ticker
            aItems(2) = 1 + aItems(1)                       ' Date
            aItems(9) = 2 + aItems(1)                       ' Dividends
       Case "S"                                             ' Just bring back Date and Splits
            If InStr(sItems, "T") > 0 Then aItems(1) = 1    ' include Ticker
            aItems(2) = 1 + aItems(1)                       ' Date
            aItems(10) = 2 + aItems(1)                      ' Splits
       Case Else                                            ' Just checks that all sItems are valid
            For i1 = 1 To Len(sItems)
                'i2 = InStr("TDOHLCVUFGXS", Mid(sItems, i1, 1))
                i2 = InStr("TDOHLUCVXS", Mid(sItems, i1, 1))
                If i2 = 0 Then
                   vData(1, 1) = "Invalid data item requested: " & Mid(sItems, i1, 1)
                   GoTo ErrorExit
                   End If
                If i1 <= iCols Then aItems(i2) = i1
                Next i1
        End Select
  
    '------------------> Verify and Process pNames parameter (Headers)
    ' If pNames =1 then insert Headings into vData(1)
    'CALL to smfWord to get Heading name and insert into vData based on order from aItems
    
    Select Case pNames
       Case 0
       Case 1
            For i1 = 1 To 10
                If aItems(i1) > 0 Then vData(1, aItems(i1)) = smfWord(kItemList, i1, ",")
                Next i1
       Case Else
            vData(1, 1) = "Invalid pNames parameter: " & pNames
       End Select
  
    '------------------> Create URL and retrieve data
    sURL = "https://query1.finance.yahoo.com/v7/finance/download/" & pTicker & "?period1=" & dBegin & "&period2=" & dEnd & _
           "&interval=" & sInterval & "&events=" & sfilter & "&includeAdjustedClose=true"
    sData = RCHGetURLData(sURL)
    vByDay = Split(sData, Chr(10))
    
    '------------------> Extract data
    dAdj = 1
    dSplitAdj = 1
    vDivAmt = 0
    vDen = 0            'Denominator used if calculating Split adjusted prices
    iRow = pNames
    
    'For i1 = 0 To UBound(vByDay) Incremented this to skip the header row returned from Yahoo
    'For i1 = 1 To UBound(vByDay)
    
    '------------------> Change data order to start with most recent date see Issue #16

    For i1 = UBound(vByDay) To 1 Step -1
       s1 = vByDay(i1)
       s1 = Replace(s1, "null", 0)
       vItem = Split(s1, ",")
       d1 = DateValue(vItem(0))    ' Get the date from the first field
            
        Select Case True
           Case iRow > pNames And sPeriod = "A" And Month(d1) <> 1
           Case iRow > pNames And sPeriod = "Q" And Month(d1) <> 1 And Month(d1) <> 4 And Month(d1) <> 7 And Month(d1) <> 10
           Case Else
                If iRow = iRows Then Exit For
                iRow = iRow + 1
                If aItems(1) > 0 Then vData(iRow, aItems(1)) = pTicker                  'ticker
                If aItems(2) > 0 Then vData(iRow, aItems(2)) = DateValue(vItem(0))      'date
                If aItems(3) > 0 Then vData(iRow, aItems(3)) = smfCDec(vItem(1))        'open
                If aItems(4) > 0 Then vData(iRow, aItems(4)) = smfCDec(vItem(2))        'high
                If aItems(5) > 0 Then vData(iRow, aItems(5)) = smfCDec(vItem(3))        'low
                If aItems(6) > 0 Then vData(iRow, aItems(6)) = smfCDec(vItem(4))        'close
                If aItems(7) > 0 Then vData(iRow, aItems(7)) = smfCDec(vItem(5))        'adj close
                If aItems(8) > 0 Then vData(iRow, aItems(8)) = smfCDec(vItem(6))        'volume
                If aItems(9) > 0 Then vData(iRow, aItems(9)) = smfCDec(vItem(1))        'dividend
                If aItems(10) > 0 Then vData(iRow, aItems(10)) = vItem(1)               'split
        End Select
        
    Next i1
    
    If sPeriod = "S" Or sPeriod = "V" Then
       For i1 = 1 To iCols
           If iRow + 1 > UBound(vData) Then Exit For
           vData(iRow + 1, i1) = ""
           Next i1
       End If
    
    '------------------> Reverse the sort order of the data if requested
    If pReSort = 1 Then
       Dim vTemp As Variant
       i1 = 1 + pNames
       i2 = iRow
       Do While i1 < i2
          For i3 = 1 To iCols
              vTemp = vData(i1, i3)
              vData(i1, i3) = vData(i2, i3)
              vData(i2, i3) = vTemp
              Next i3
          i1 = i1 + 1
          i2 = i2 - 1
          Loop
       End If
    
ErrorExit:
    smfGetYahooHistory1 = vData
    
    End Function
