Attribute VB_Name = "modGetMSNHistory"
Public Function smfGetMSNHistory(pTicker As String, _
                          Optional pStartYear As Integer = 1900, _
                          Optional pStartMonth As Integer = 1, _
                          Optional pEndYear As Integer = 2100, _
                          Optional pEndMonth As Integer = 12, _
                          Optional pItems As String = "DOHLCV", _
                          Optional pNames As Integer = 1, _
                          Optional pResort As Integer = 0, _
                          Optional pDim1 As Integer = 0, _
                          Optional pDim2 As Integer = 0) ' As Variant()
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to download historical quotes from MSN
    '-----------------------------------------------------------------------------------------------------------*
    ' 2008.03.17 -- Written by Randy Harmelink (rharmelink@gmail.com)
    ' 2008.11.07 -- Adjust row processing
    ' 2009.11.07 -- Adjust row processing
    ' 2009.04.09 -- Adjust editing of ending date
    ' 2011.04.28 -- Change cDec() to smfConvertData()
    '-----------------------------------------------------------------------------------------------------------*
    ' > Example of an invocation to get daily quotes for 2004 for IBM:
    '
    '   =smfGetMSNHistory("IBM")
    '   =smfGetMSNHistory("IBM",2004,1,2008,3)
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim sURL As String
    
    On Error GoTo ErrorExit
    
    '------------------> Null Return Item
    If pTicker = "None" Or pTicker = "" Then
       ReDim vData(1 To 1, 1 To 1) As Variant
       vData(1, 1) = "None"
       smfGetMSNHistory = vData
       Exit Function
       End If
    
    '------------------> Determine size of array to return
    kDim1 = pDim1  ' Rows
    kDim2 = pDim2  ' Columns
    If pDim1 = 0 Or pDim2 = 0 Then
       On Error Resume Next
       kDim1 = Application.Caller.Rows.Count
       kDim2 = Application.Caller.Columns.Count
       On Error GoTo ErrorExit
       End If
  
    '------------------> Initialize return array
    ReDim vData(1 To kDim1, 1 To kDim2) As Variant
    For i1 = 1 To kDim1
        For i2 = 1 To kDim2
            vData(i1, i2) = ""
            Next i2
        Next i1
    
    '------------------> Edit parameters
    If DateSerial(pStartYear, pStartMonth, 1) < DateSerial(Year(Date) - 9, Month(Date) + 1, 1) Then
       dBegDate = DateSerial(Year(Date) - 9, Month(Date) + 1, 1)
    Else
       dBegDate = DateSerial(pStartYear, pStartMonth, 1)
       End If
    
    If DateSerial(pEndYear, pEndMonth, 1) > DateSerial(Year(Date), Month(Date) + 1, 0) Then
       dEndDate = DateSerial(Year(Date), Month(Date), 1)
    Else
       dEndDate = DateSerial(pEndYear, pEndMonth, 1)
       End If
       
    '------------------> Create URL and download historical quotes
    
    sBase = "http://data.moneycentral.msn.com/scripts/chrtsrv.dll?C1=2&C2=&FileDownload=&C9=0"
    sURL = sBase & "&Symbol=" & pTicker & _
           "&C1=" & Month(dBegDate) & _
           "&C6=" & Year(dBegDate) & _
           "&C7=" & Month(dEndDate) & _
           "&C8=" & Year(dEndDate)

    sqData = RCHGetURLData(sURL)
    
    '------------------> Determine items needed
    pItems2 = UCase(pItems)
    iTick = InStr(pItems2, "T")
    iDate = InStr(pItems2, "D")
    iOpen = InStr(pItems2, "O")
    iHigh = InStr(pItems2, "H")
    iLow = InStr(pItems2, "L")
    iClos = InStr(pItems2, "C")
    iVol = InStr(pItems2, "V")
    If iTick > kDim2 Then iTick = 0
    If iDate > kDim2 Then iDate = 0
    If iOpen > kDim2 Then iOpen = 0
    If iHigh > kDim2 Then iHigh = 0
    If iLow > kDim2 Then iLow = 0
    If iClos > kDim2 Then iClos = 0
    If iVol > kDim2 Then iVol = 0
    
    '------------------> Parse web quotes
    vLine = Split(sqData, Chr(10))
    nLines = IIf(kDim1 - pNames < UBound(vLine) - 1, kDim1 - pNames, UBound(vLine) - 1)
    For iRow = (6 - pNames) To nLines + 5
        vItem = Split(vLine(iRow), ",")
        If iRow = 5 Then
           If iTick > 0 Then vData(iRow + pNames - 5, iTick) = "Ticker"
           If iDate > 0 Then vData(iRow + pNames - 5, iDate) = vItem(0)
           If iOpen > 0 Then vData(iRow + pNames - 5, iOpen) = vItem(1)
           If iHigh > 0 Then vData(iRow + pNames - 5, iHigh) = vItem(2)
           If iLow > 0 Then vData(iRow + pNames - 5, iLow) = vItem(3)
           If iClos > 0 Then vData(iRow + pNames - 5, iClos) = vItem(4)
           If iVol > 0 Then vData(iRow + pNames - 5, iVol) = vItem(5)
        Else
           If iTick > 0 Then vData(iRow + pNames - 5, iTick) = pTicker
           If iDate > 0 Then vData(iRow + pNames - 5, iDate) = CDate(vItem(0))
           If iOpen > 0 Then vData(iRow + pNames - 5, iOpen) = smfConvertData(vItem(1))
           If iHigh > 0 Then vData(iRow + pNames - 5, iHigh) = smfConvertData(vItem(2))
           If iLow > 0 Then vData(iRow + pNames - 5, iLow) = smfConvertData(vItem(3))
           If iClos > 0 Then vData(iRow + pNames - 5, iClos) = smfConvertData(vItem(4))
           If iVol > 0 Then vData(iRow + pNames - 5, iVol) = smfConvertData(vItem(5))
           End If
        Next iRow
    
    '------------------> Reverse the sort order of the data if requested
    If pResort = 1 Then
       Dim vTemp As Variant
       i1 = 1 + pNames
       i2 = nLines + pNames
       Do While i1 < i2
          For i3 = 1 To kDim2
              vTemp = vData(i1, i3)
              vData(i1, i3) = vData(i2, i3)
              vData(i2, i3) = vTemp
              Next i3
          i1 = i1 + 1
          i2 = i2 - 1
          Loop
       End If
    
ErrorExit:
    smfGetMSNHistory = vData
    End Function


