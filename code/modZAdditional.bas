Attribute VB_Name = "modZAdditional"
'@Lang VBA


Public Function DateOfHigh(pTicker As String, pDays As Integer)
   Dim vData() As Variant
   vData = RCHGetYahooHistory2(pTicker, , , , , , , "d", "DH", 0, 1, 0, pDays, 2)
   dStart = vData(1, 1)
   dHighDay = dStart
   nHighDay = vData(1, 2)
   For i1 = 2 To UBound(vData, 1)
       If vData(i1, 1) < dStart - pDays Then Exit For
       If vData(i1, 2) > nHighDay Then
          nHighDay = vData(i1, 2)
          dHighDay = vData(i1, 1)
          End If
       Next i1
   DateOfHigh = dHighDay
   End Function
Public Function SMFHighBetween(pTicker As String, pBegDate As Variant, pEndDate As Variant)
   ' Checks for highest price between two dates
   Dim vData(1 To 1, 1 To 4) As Variant
   vHQ = RCHGetYahooHistory2(pTicker, , , , , , , "d", "DHOC", 0, 1, 0, 9999, 4)
   vData(1, 1) = 0    ' Value of high price
   vData(1, 2) = ""   ' Day of high price
   vData(1, 3) = 0    ' Starting price
   vData(1, 4) = 0    ' Ending price
   For i1 = 1 To UBound(vHQ, 1)
       Select Case vHQ(i1, 1)
          Case Is > pEndDate
          Case Is < pBegDate: Exit For
          Case Else
               If vHQ(i1, 1) = pBegDate Then vData(1, 3) = vHQ(i1, 3)
               If vHQ(i1, 1) = pEndDate Then vData(1, 4) = vHQ(i1, 4)
               If vHQ(i1, 2) > vData(1, 1) Then
                  vData(1, 1) = vHQ(i1, 2)
                  vData(1, 2) = vHQ(i1, 1)
                  End If
       End Select
       Next i1
   SMFHighBetween = vData
   End Function
Public Function SMFLowBetween(pTicker As String, pBegDate As Variant, pEndDate As Variant)
   ' Checks for highest price between two dates
   Dim vData(1 To 1, 1 To 4) As Variant
   vHQ = RCHGetYahooHistory2(pTicker, , , , , , , "d", "DLOC", 0, 1, 0, 9999, 4)
   vData(1, 1) = 999999 ' Value of low price
   vData(1, 2) = ""     ' Day of low price
   vData(1, 3) = 0      ' Starting price
   vData(1, 4) = 0      ' Ending price
   For i1 = 1 To UBound(vHQ, 1)
       Select Case vHQ(i1, 1)
          Case Is > pEndDate
          Case Is < pBegDate: Exit For
          Case Else
               If vData(1, 4) = 0 Then vData(1, 4) = vHQ(i1, 4)
               vData(1, 3) = vHQ(i1, 3)
               If vHQ(i1, 2) < vData(1, 1) Then
                  vData(1, 1) = vHQ(i1, 2)
                  vData(1, 2) = vHQ(i1, 1)
                  End If
       End Select
       Next i1
   SMFLowBetween = vData
   End Function
Public Function smfLastPrice(pTicker As String, pEndDate As Variant) As Variant
   '-----------------------------------------------------------------------------------------------------------*
   ' Sample routine to get the last traded price (adjusted) for a given day
   '-----------------------------------------------------------------------------------------------------------*
   ' 2007.07.26 -- Created by Randy Harmelink (rharmelink@gmail.com)
   '-----------------------------------------------------------------------------------------------------------*
   ' Sample of use:
   '
   '    =smfLastPrice("MMM",DATE(2007,1,1))
   '
   '-----------------------------------------------------------------------------------------------------------*
   vHQ = RCHGetYahooHistory2(pTicker, , , , , , , "d", "DA", 0, 1, 0, 9999, 2)
   smfLastPrice = 0
   For i1 = 1 To UBound(vHQ, 1)
       If vHQ(i1, 1) <= pEndDate Then
          smfLastPrice = vHQ(i1, 2)
          Exit Function
          End If
       Next i1

   End Function
Sub Testing2()
    Open ThisWorkbook.Path & "\smf-elements.txt" For Input As #1
    Do Until EOF(1) = True
       Line Input #1, sLine
       Loop
    Close #1
    On Error GoTo ErrorExit
    Open ThisWorkbook.Path & "\smf-elements-1.txt" For Input As #1
    Do Until EOF(1) = True
       Line Input #1, sLine
       Loop
    Close #1
ErrorExit:
    End Sub
