Attribute VB_Name = "smfPricesByDates_"
'@Lang VBA
Public Function smfPricesByDates(ByVal pTicker As String, _
                            ParamArray pDates() As Variant)
                        
   '-----------------------------------------------------------------------------------------------------------*
   ' Function to return prices for multiple historical dates
   '-----------------------------------------------------------------------------------------------------------*
   ' 2008.05.23 -- Created by Randy Harmelink (rharmelink@gmail.com)
   ' 2009.01.01 -- Added vbDouble "Case" for DATE() and DATEVALUE() passed items, but check YEAR() values
   ' 2009.01.02 -- Change date lookup to be a binary search
   ' 2009.01.02 -- Change invalid date returns to be EXCEL #N/A error values
   ' 2009.01.02 -- Add ability to pass string dates (e.g. "12/31/2006")
   ' 2009.01.02 -- Add retrieval of current date for today's date after available history
   ' 2015.02.21 -- Fix retrieval of current date for today's date after available history
   ' 2017.05.18 -- Change to use new smfGetYahooHistory() function
   ' 2017.05.21 -- Allow range of string dates to be passed
   ' 2019.09.09 -- Change RCHGetYahooQuotes() to smfGetYahooPortfolioView()
   '-----------------------------------------------------------------------------------------------> Version 2.0i
   ' Samples of use:
   '
   '    =smfPricesByDates("MMM",DATE(2007,1,1),DATE(2007,3,4))
   '    =smfPricesByDates("MMM","1/1/2007")
   '    =smfPricesByDates("MMM",C4:D4)
   '    =smfPricesByDates("MMM",DATE(2007,1,1),DATE(2007,3,4),C4:D4)
   '
   '-----------------------------------------------------------------------------------------------------------*
      
   '----------------------------------> Extract passed dates from parameters and/or ranges
   ReDim vDates(1 To 1) As Variant
   iCount = 0
   dBegin = Int(Now)
   For i1 = 0 To UBound(pDates)
       Select Case VarType(pDates(i1))
          Case vbDate, vbDouble: Call AddToList(vDates, iCount, dBegin, pDates(i1))
          Case vbString
             If IsDate(pDates(i1)) Then
                Call AddToList(vDates, iCount, dBegin, DateValue(pDates(i1)))
             Else
                Call AddToList(vDates, iCount, dBegin, "")
                End If
          Case Is >= 8192
               For Each oCell In pDates(i1)
                   Select Case True
                      Case VarType(oCell.Value) = vbDate: Call AddToList(vDates, iCount, dBegin, oCell.Value)
                      Case VarType(oCell.Value) = vbString And IsDate(oCell.Value): Call AddToList(vDates, iCount, dBegin, DateValue(oCell.Value))
                      Case Else: Call AddToList(vDates, iCount, dBegin, "")
                      End Select
                   Next oCell
          Case Else: Call AddToList(vDates, iCount, dBegin, "")
          End Select
       Next i1
   
   '----------------------------------> Get historical data and extract requested data
   Dim iDays As Integer
   iDays = Int(Now - dBegin + 3)
   vHQ = smfGetYahooHistory(pTicker, dBegin - 5, Int(Now), "d", "DC", 0, 0, iDays, 2)
   ReDim vReturn(1 To iCount) As Variant
   For i1 = 1 To iCount
       If vDates(i1) = "" Or vDates(i1) > Date Then
          vReturn(i1) = CVErr(xlErrNA)
       ElseIf vDates(i1) > vHQ(1, 1) Then
          If vDates(i1) = Date Then
             vPrice = smfGetYahooPortfolioView(pTicker, "15")
             vReturn(i1) = vPrice(1, 1)
          Else
             vReturn(i1) = vHQ(1, 2)
             End If
       Else
          iLo = 1
          iHi = iDays
          Do
             i2 = Int((iHi + iLo) / 2)
             If vDates(i1) = vHQ(i2, 1) Then
                vReturn(i1) = vHQ(i2, 2)
                Exit Do
             ElseIf iLo = iHi - 1 Then
                If vHQ(iHi, 2) <> "" Then
                   vReturn(i1) = vHQ(iHi, 2)
                Else
                   vReturn(i1) = CVErr(xlErrNA)
                   End If
                Exit Do
             Else
                If vDates(i1) > vHQ(i2, 1) Or vHQ(i2, 1) = "" Then
                   iHi = i2
                Else
                   iLo = i2
                   End If
                End If
             Loop While True
          End If
       Next i1
       
   '----------------------------------> Return data
   smfPricesByDates = vReturn
               
ErrorExit:
               
    End Function

Private Sub AddToList(pList As Variant, pCount As Variant, pBegin As Variant, pDate As Variant)
    pCount = pCount + 1
    ReDim Preserve pList(1 To pCount) As Variant
    If pDate = "" Then
       pList(pCount) = ""
    ElseIf Year(pDate) < 1928 Or Year(pDate) > Year(Date) Then
       pList(pCount) = ""
    Else
       pList(pCount) = pDate
       If pDate < pBegin Then pBegin = pDate
       End If
    End Sub

