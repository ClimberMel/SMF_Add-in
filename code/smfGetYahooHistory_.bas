Attribute VB_Name = "smfGetYahooHistory_"
Option Explicit

Function smfGetYahooHistory(ByVal pTicker As String, _
                   Optional ByVal pStartDate As Variant = "", _
                   Optional ByVal pEndDate As Variant = "", _
                   Optional ByVal pPeriod As String = "d", _
                   Optional ByVal pItems As String = "dohlcvufgxs", _
                   Optional ByVal pNames As Integer = 1, _
                   Optional ByVal pResort As Integer = 0, _
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
    ' 2022-12-30 -- This module is no longer working.  Functions such as RCHGetYahooHistory & smfPrices* modules
    '               call this and therfore are not working either.
    ' 2023-01-22 -- Mel Pryor (ClimberMel@gmail.com)
    '               Created a fix in RCHGetYahooHistory so it doesn't call this module.  Will continue to look for a fix.
    '-----------------------------------------------------------------------------------------------------------*
    ' > Example of an invocation to get daily quotes for 2004 for IBM:
    '
    '   =smfGetYahooHistory("IBM","1/1/2017","5/18/2017","d")
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim s1 As String, sURL As String, sData As String, sFind As String, sFound As String
    Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer
    Dim iRows As Integer, iCols As Integer, iRow As Integer
    Dim vDivAmt As Variant, vNum As Variant, vDen As Variant
    Dim dAdj As Double, d1 As Double, d2 As Double, dSplitAdj As Double
    Dim vByDay As Variant
    ReDim vData(1 To 1, 1 To 1) As Variant
    
    vData(1, 1) = "Error"
    
    On Error GoTo ErrorExit
    
    '------------------> Set defaults, if necessary
    If pPeriod = "" Then pPeriod = "d"
    If pItems = "" Then pItems = "dohlcvufgxs"
    
    '------------------> Null Return Item
    If pTicker = "None" Or pTicker = "" Then
       vData(1, 1) = "None"
       GoTo ErrorExit
       End If
       
    '------------------> Verify and Process starting and ending dates
    Dim dBegin As Double, dEnd As Double
    vData(1, 1) = "Error on starting date: " & pStartDate
    Select Case True
          Case VarType(pStartDate) = vbDate Or VarType(pStartDate) = vbDouble
               dBegin = smfDate2Unix(pStartDate)
          Case pStartDate = ""
               dBegin = smfDate2Unix(DateValue("1/1/1970"))
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
       Case "V": sFreq = "1d": sfilter = "div": sInterval = "div|split"
       Case "S": sFreq = "1d": sfilter = "split": sInterval = "div|split"
       Case Else: GoTo ErrorExit
       End Select
    vData(1, 1) = "Error"
    
    '------------------> Verify and Process pItems parameter
    Const kItemList As String = "Ticker,Date,Open,High,Low,Close,Volume,Unadj,Div Adj,Split Adj,Dividend,Split"
    Dim sItems As String, aItems(1 To 12) As Integer
    For i1 = 1 To 12: aItems(i1) = 0: Next i1
    sItems = UCase(pItems)
    Select Case sPeriod
       Case "V"
            If InStr(sItems, "T") > 0 Then aItems(1) = 1
            aItems(2) = 1 + aItems(1)  ' Date
            aItems(11) = 2 + aItems(1) ' Dividends
       Case "S"
            If InStr(sItems, "T") > 0 Then aItems(1) = 1
            aItems(2) = 1 + aItems(1)  ' Date
            aItems(12) = 2 + aItems(1) ' Splits
       Case Else
            For i1 = 1 To Len(sItems)
                i2 = InStr("TDOHLCVUFGXS", Mid(sItems, i1, 1))
                If i2 = 0 Then
                   vData(1, 1) = "Invalid data item requested: " & Mid(sItems, i1, 1)
                   GoTo ErrorExit
                   End If
                If i1 <= iCols Then aItems(i2) = i1
                Next i1
        End Select
  
    '------------------> Verify and Process pNames parameter
    Select Case pNames
       Case 0
       Case 1
            For i1 = 1 To 12
                If aItems(i1) > 0 Then vData(1, aItems(i1)) = smfWord(kItemList, i1, ",")
                Next i1
       Case Else
            vData(1, 1) = "Invalid pNames parameter: " & pNames
       End Select
  
    '------------------> Create URL and retrieve data
    sURL = "https://finance.yahoo.com/quote/" & pTicker & "/history?period1=" & dBegin & "&period2=" & dEnd & _
           "&interval=" & sInterval & "&filter=" & sfilter & "&frequency=" & sFreq
    'sData = RCHGetURLData(sURL)
    sData = smfGetWebPage(sURL)
    sData = smfStrExtr(sData, "HistoricalPriceStore", "]")   ' Keep only the "HistoricalPriceStore" JSON data
    vByDay = Split(sData, "},{")
  
    '------------------> Extract data
    dAdj = 1
    dSplitAdj = 1
    vDivAmt = 0
    vDen = 0
    iRow = pNames
    For i1 = 0 To UBound(vByDay)
       s1 = vByDay(i1) & "}"
       s1 = Replace(s1, "null", 0)
       Select Case True
          Case InStr(s1, "DIVIDEND") > 0
               vDivAmt = smfStrExtr(s1, """amount"":", ",", 1)
          Case InStr(s1, "SPLIT") > 0
               vDen = smfStrExtr(s1, """denominator"":", ",", 1)
               vNum = smfStrExtr(s1, """numerator"":", ",", 1)
          Case Else
               If iRow > iRows Then Exit For
               Select Case True
                  Case vDivAmt > 0 And sPeriod <> "S"
                       If sPeriod = "V" Then iRow = iRow + 1
                       i2 = aItems(11)
                       If i2 > 0 And iRow > pNames Then
                          If VarType(vData(iRow, i2)) = vbString Then vData(iRow, i2) = 0
                          vData(iRow, i2) = vData(iRow, i2) + vDivAmt
                          End If
                       d1 = smfStrExtr(s1, """close"":", ",", 1)
                       If d1 <> 0 Then dAdj = dAdj * (d1 - vDivAmt) / d1
                       vDivAmt = 0
                  Case vDen > 0 And sPeriod <> "V"
                       If sPeriod = "S" Then iRow = iRow + 1
                       If aItems(12) > 0 Then vData(iRow, aItems(12)) = vDen & " for " & vNum
                       dSplitAdj = dSplitAdj * vNum / vDen
                       vDen = 0
                       vNum = 0
                  End Select
               d1 = Int(smfUnix2Date(smfStrExtr(s1, """date"":", ",")))
               Select Case True
                  Case iRow > pNames And sPeriod = "A" And Month(d1) <> 1
                  Case iRow > pNames And sPeriod = "Q" And Month(d1) <> 1 And Month(d1) <> 4 And Month(d1) <> 7 And Month(d1) <> 10
                  Case Else
                       If iRow = iRows Then Exit For
                       iRow = iRow + 1
                       If aItems(1) > 0 Then vData(iRow, aItems(1)) = pTicker
                       If aItems(2) > 0 Then vData(iRow, aItems(2)) = d1
                       If aItems(3) > 0 Then vData(iRow, aItems(3)) = smfStrExtr(s1, """open"":", ",", 1) * dAdj
                       If aItems(4) > 0 Then vData(iRow, aItems(4)) = smfStrExtr(s1, """high"":", ",", 1) * dAdj
                       If aItems(5) > 0 Then vData(iRow, aItems(5)) = smfStrExtr(s1, """low"":", ",", 1) * dAdj
                       d2 = smfStrExtr(s1, """adjclose"":", "}", 1) ' * dAdj
                       If aItems(6) > 0 Then vData(iRow, aItems(6)) = d2
                       If aItems(7) > 0 Then vData(iRow, aItems(7)) = smfStrExtr(s1, """volume"":", ",", 1)
                       'If aItems(8) > 0 Then vData(iRow, aItems(8)) = smfStrExtr(s1, """unadjclose"":", "}".1)
                       If aItems(8) > 0 Then vData(iRow, aItems(8)) = smfStrExtr(s1, """close"":", ",", 1)
                       If aItems(9) > 0 Then vData(iRow, aItems(9)) = dAdj
                       If aItems(10) > 0 Then vData(iRow, aItems(10)) = dSplitAdj
                       '----------------------------* Forward fill missing data
                       If d2 > 0 Then
                          For i4 = 3 To 6
                              If aItems(i4) > 0 Then
                                 For i2 = iRow - 1 To pNames + 1 Step -1
                                     If vData(i2, aItems(i4)) <> 0 Then Exit For
                                     If aItems(i4) > 0 Then vData(i2, aItems(i4)) = d2
                                     Next i2
                                 End If
                              Next i4
                          If aItems(8) > 0 Then
                             For i2 = iRow - 1 To pNames + 1 Step -1
                                 If vData(i2, aItems(8)) <> 0 Then Exit For
                                 vData(i2, aItems(8)) = vData(iRow, aItems(8))
                                 Next i2
                              End If
                          End If
                       If sPeriod = "S" Or sPeriod = "V" Then iRow = iRow - 1
                   End Select
          End Select
       Next i1
    If sPeriod = "S" Or sPeriod = "V" Then
       For i1 = 1 To iCols
           If iRow + 1 > UBound(vData) Then Exit For
           vData(iRow + 1, i1) = ""
           Next i1
       End If
    
    '------------------> Reverse the sort order of the data if requested
    If pResort = 1 Then
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
    smfGetYahooHistory = vData
    
    End Function
