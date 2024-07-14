Attribute VB_Name = "smfPricesBetween_"
'@Lang VBA
Public Function smfPricesBetween(ByVal pTicker As String, _
                        Optional ByVal pBegDate As Variant, _
                        Optional ByVal pEndDate As Variant, _
                        Optional ByVal pItems As Variant = "01020304050607080910111213")
                        
   '-----------------------------------------------------------------------------------------------------------*
   ' Sample routine to summarize historical data between two dates -- O/H/L/C/V/PC
   '-----------------------------------------------------------------------------------------------------------*
   ' 2007.05.17 -- Created by Randy Harmelink (rharmelink@gmail.com)
   ' 2007.08.08 -- Added return element #10 (previous closing price)
   ' 2007.08.08 -- Added parameter to specify return results
   ' 2007.08.08 -- Customized number of days to return
   ' 2007.08.29 -- Increased number of days to return
   ' 2007.10.03 -- Added ability to generate column headings
   ' 2007.10.03 -- Added ErrorExit
   ' 2017.05.19 -- Change to use smfGetYahooHistory()
   ' 2017.07.23 -- Add total return / max drawdown / CAGR output options
   '-----------------------------------------------------------------------------------------------------------*
   ' Samples of use:
   '
   '    =smfPricesBetween("MMM",DATE(2007,1,1),DATE(2007,3,4))
   '    =TRANSPOSE(smfPricesBetween("MMM",DATE(2007,1,1),DATE(2007,3,4)))
   '
   ' Both would need to be array-entered.  The first would return a 1-row by 10-column range.  The second would
   ' return a 1-column by 10-row range.  The 9 elements of the range would be Date and Value of opening price,
   ' Date and Value of highest price, Date and Value of Lowest price, Data and Value of closing price, total
   ' volume between the two dates, as well as the previous closing price.
   '-----------------------------------------------------------------------------------------------------------*
     
   Const kItems = 13
   Dim vData(1 To 1, 1 To kItems) As Variant
   
   On Error GoTo ErrorExit
   vData(1, 1) = "Error"
   
   If pTicker = "Header" Or pTicker = "Ticker" Or pTicker = "Symbol" Then
      vData(1, 1) = "Open Date"
      vData(1, 2) = "Open Price"
      vData(1, 3) = "High Date"
      vData(1, 4) = "High Price"
      vData(1, 5) = "Low Date"
      vData(1, 6) = "Low Price"
      vData(1, 7) = "Close Date"
      vData(1, 8) = "Close Price"
      vData(1, 9) = "Volume"
      vData(1, 10) = "Previous Close"
      vData(1, 11) = "Total Return"
      vData(1, 12) = "CAGR"
      vData(1, 13) = "Max Drawdown"
      GoTo SkipRetrieval
      End If
   
   'vHQ = RCHGetYahooHistory2(pTicker, , , , , , , "d", "DOHLCV", 0, 1, 0, Int(Now - pBegDate + 3), 6)
   vHQ = smfGetYahooHistory(pTicker, pBegDate - 5, Int(Now) + 1, "d", "DOHLCV", 0, 0, Int(Now - pBegDate + 5), 6)
  
   vData(1, 1) = ""     ' Date of open price
   vData(1, 2) = 0      ' Value of open price
   vData(1, 3) = ""     ' Date of high price
   vData(1, 4) = 0      ' Value of high price
   vData(1, 5) = ""     ' Date of low price
   vData(1, 6) = 0      ' Value of low price
   vData(1, 7) = ""     ' Date of closing price
   vData(1, 8) = 0      ' Value of closing price
   vData(1, 9) = 0      ' Total volume during period
   vData(1, 10) = 0     ' Value of previous closing price
   vData(1, 11) = 0     ' Total return
   vData(1, 12) = 0     ' CAGR
   vData(1, 13) = 0     ' Max drawdown
   
   For i1 = 1 To UBound(vHQ, 1)
       Select Case vHQ(i1, 1)
          Case Is > pEndDate: Exit For
          Case Is < pBegDate
          Case Else
               If vData(1, 8) = 0 Then
                  vData(1, 3) = vHQ(i1, 1) ' Latest date
                  vData(1, 4) = vHQ(i1, 3) ' Latest high
                  vData(1, 5) = vHQ(i1, 1) ' Latest date
                  vData(1, 6) = vHQ(i1, 4) ' Latest low
                  vData(1, 7) = vHQ(i1, 1) ' Latest date
                  vData(1, 8) = vHQ(i1, 5) ' Latest close
                  End If
               vData(1, 1) = vHQ(i1, 1)    ' Earliest date
               vData(1, 2) = vHQ(i1, 2)    ' Earliest open
               vData(1, 10) = vHQ(i1 + 1, 5)    ' Previous closing price
               vData(1, 9) = vData(1, 9) + vHQ(i1, 6)
               If vData(1, 6) > vHQ(i1, 4) Then
                  vData(1, 5) = vHQ(i1, 1) ' Date of lowest
                  vData(1, 6) = vHQ(i1, 4) ' Lower low
                  End If
               If vData(1, 4) < vHQ(i1, 3) Then
                  vData(1, 3) = vHQ(i1, 1) ' Date of highest
                  vData(1, 4) = vHQ(i1, 3) ' Higher high
                  End If
       End Select
       Next i1
       vData(1, 11) = vData(1, 8) / vData(1, 10) - 1
       vData(1, 12) = (vData(1, 8) / vData(1, 10)) ^ (365 / (vData(1, 7) - vData(1, 1) + 1)) - 1
       vData(1, 13) = vData(1, 6) / vData(1, 10) - 1
       
   
SkipRetrieval:
   
   Dim vReturn(1 To 1, 1 To kItems) As Variant
   For i1 = 1 To kItems
       If 2 * i1 > Len(pItems) Then
          vReturn(1, i1) = ""
       Else
          iItem = CInt(Mid(pItems, 2 * i1 - 1, 2))
          vReturn(1, i1) = vData(1, iItem)
          End If
       Next i1
   
ErrorExit:
   
   smfPricesBetween = vReturn
   
   End Function


