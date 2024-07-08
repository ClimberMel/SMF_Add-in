Attribute VB_Name = "modGetOptionExpiry"
'@Lang VBA
Option Explicit
Public Function smfGetOptionExpiry(Optional ByVal pYear As Integer = 0, _
                                   Optional ByVal pMonth As Integer = 0, _
                                   Optional ByVal pType As String = "M")
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option expiration date for a given month from Yahoo
    '-----------------------------------------------------------------------------------------------------------*
    ' 2010.04.08 -- Created function
    ' 2011.03.30 -- Fixed usage of iIncr so it is only used for default date values of year and month
    ' 2011.11.23 -- Fixed calculation of quarter end expiration dates that fall on Saturday and Sunday
    ' 2014.08.18 -- Added W3/W4/W5 values for pType
    ' 2014.08.18 -- Changed monthly expiration dates for 2016 and forward, due to alignment to weeklies
    ' 2014.09.15 -- Changed monthly expiration dates for 2015-02-01 and forward, due to alignment to weeklies
    ' 2015.02.21 -- Added W6/W7 values for pType
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for IBM and MMM:
    '
    '   =smfGetOptionExpiry(2012,12)
    '   =smfGetOptionExpiry(2012,12,"M")
    '   =smfGetOptionExpiry(2012,12,"Q")
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim dTemp As Date, iIncr As Integer, iYear As Integer, iMonth As Integer
    pType = UCase(pType)
    If pYear = 0 Or Left(pType, 1) = "W" Then iYear = Year(Date) Else iYear = pYear
    If pMonth = 0 Or Left(pType, 1) = "W" Then iMonth = Month(Date) Else iMonth = pMonth
    dTemp = DateSerial(iYear, iMonth, 21) - Weekday(DateSerial(iYear, iMonth, 16), 1) + 2
    If pMonth = 0 Or pYear = 0 Or Left(pType, 1) = "M" Then iIncr = -(dTemp < Date) Else iIncr = 0
    
    Select Case True
       Case pType = "W" Or pType = "W1" Or pType = "WEEK" Or pType = "WEEKLY"
            smfGetOptionExpiry = Date - Weekday(Date) + 6
       Case pType = "W2"
            smfGetOptionExpiry = Date - Weekday(Date) + 13
       Case pType = "W3"
            smfGetOptionExpiry = Date - Weekday(Date) + 20
       Case pType = "W4"
            smfGetOptionExpiry = Date - Weekday(Date) + 27
       Case pType = "W5"
            smfGetOptionExpiry = Date - Weekday(Date) + 34
       Case pType = "W6"
            smfGetOptionExpiry = Date - Weekday(Date) + 41
       Case pType = "W7"
            smfGetOptionExpiry = Date - Weekday(Date) + 48
       Case pType = "M" Or pType = "M1" Or pType = "MONTH" Or pType = "MONTHLY"
            smfGetOptionExpiry = DateSerial(iYear, iMonth + iIncr, 21) - Weekday(DateSerial(iYear, iMonth + iIncr, 16), 1) + 2
            If smfGetOptionExpiry > DateSerial(2015, 2, 1) Then smfGetOptionExpiry = smfGetOptionExpiry - 1
       Case pType = "M2"
            smfGetOptionExpiry = DateSerial(iYear, iMonth + iIncr + 1, 21) - Weekday(DateSerial(iYear, iMonth + iIncr + 1, 16), 1) + 2
            If smfGetOptionExpiry > DateSerial(2015, 2, 1) Then smfGetOptionExpiry = smfGetOptionExpiry - 1
       Case pType = "M3"
            smfGetOptionExpiry = DateSerial(iYear, iMonth + iIncr + 2, 21) - Weekday(DateSerial(iYear, iMonth + iIncr + 2, 16), 1) + 2
            If smfGetOptionExpiry > DateSerial(2015, 2, 1) Then smfGetOptionExpiry = smfGetOptionExpiry - 1
       Case pType = "M4"
            smfGetOptionExpiry = DateSerial(iYear, iMonth + iIncr + 3, 21) - Weekday(DateSerial(iYear, iMonth + iIncr + 3, 16), 1) + 2
            If smfGetOptionExpiry > DateSerial(2015, 2, 1) Then smfGetOptionExpiry = smfGetOptionExpiry - 1
       Case pType = "M5"
            smfGetOptionExpiry = DateSerial(iYear, iMonth + iIncr + 4, 21) - Weekday(DateSerial(iYear, iMonth + iIncr + 4, 16), 1) + 2
            If smfGetOptionExpiry > DateSerial(2015, 2, 1) Then smfGetOptionExpiry = smfGetOptionExpiry - 1
       Case (pType = "Q" Or pType = "QTR" Or pType = "QUARTER" Or pType = "QUARTERLY") And (iMonth = 3 Or iMonth = 12)
            smfGetOptionExpiry = DateSerial(iYear, iMonth, 31)
            If Weekday(smfGetOptionExpiry) = vbSaturday Then smfGetOptionExpiry = smfGetOptionExpiry - 1
            If Weekday(smfGetOptionExpiry) = vbSunday Then smfGetOptionExpiry = smfGetOptionExpiry - 2
       Case (pType = "Q" Or pType = "QTR" Or pType = "QUARTER" Or pType = "QUARTERLY") And (iMonth = 6 Or iMonth = 9)
            smfGetOptionExpiry = DateSerial(iYear, iMonth, 30)
            If Weekday(smfGetOptionExpiry) = vbSaturday Then smfGetOptionExpiry = smfGetOptionExpiry - 1
            If Weekday(smfGetOptionExpiry) = vbSunday Then smfGetOptionExpiry = smfGetOptionExpiry - 2
       Case (pType = "Q" Or pType = "QTR" Or pType = "QUARTER" Or pType = "QUARTERLY")
            smfGetOptionExpiry = "Invalid month for quarterly option (3, 6, 9, or 12): " & iMonth
       Case Else
            smfGetOptionExpiry = "Invalid period type (W/W1-W7/M/M1-M5/Q): " & pType
       End Select
       
       '------------* Can be removed in 2016, when monthlies and weeklies align on Fridays
       If Left(pType, 1) = "W" Then
          dTemp = smfGetOptionExpiry(Year(smfGetOptionExpiry), Month(smfGetOptionExpiry), "M")
          If dTemp - 1 = smfGetOptionExpiry Then smfGetOptionExpiry = dTemp
          End If

    End Function


