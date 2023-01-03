Attribute VB_Name = "modRSITargetPrices"
Option Explicit
Public Function smfRSITargetPrices(ByVal pTicker As String, _
                          Optional ByVal pLoTrigger As Integer = 20, _
                          Optional ByVal pHiTrigger As Integer = 80, _
                          Optional ByVal pItems As Variant = "010203040506070809101112")
                        
   '-----------------------------------------------------------------------------------------------------------*
   ' Function to return RSI indicator buy and sell target prices
   '-----------------------------------------------------------------------------------------------------------*
   ' 2012.01.06 -- Created by Randy Harmelink (rharmelink@gmail.com)
   ' 2017.05.19 -- Change to use smfGetYahooHistory()
   ' 2017.05.23 -- Create starting date for smfGetYahooHistory() call
   '-----------------------------------------------------------------------------------------------------------*
   ' Samples of use:
   '
   '    =smfRSITargetPrices("MMM",20,80)
   '    =smfRSITargetPrices("MMM",20,80,"010203")
   '-----------------------------------------------------------------------------------------------------------*
     
   Const kItems = 12
   Dim vData(1 To 1, 1 To kItems) As Variant
   
   On Error GoTo ErrorExit
   vData(1, 1) = "Error"
   
   Dim i1 As Integer
   If pTicker = "Header" Or pTicker = "Ticker" Or pTicker = "Symbol" Then
      For i1 = 1 To kItems
          Select Case Mid(pItems, 2 * i1 - 1, 2)
             Case "01": vData(1, i1) = "Current RSI"
             Case "02": vData(1, i1) = "Buy Target Price"
             Case "03": vData(1, i1) = "Sell Target Price"
             Case "04": vData(1, i1) = "Last Traded Price"
             Case "05": vData(1, i1) = "Bid Price"
             Case "06": vData(1, i1) = "Ask Price"
             Case "07": vData(1, i1) = "Open Price"
             Case "08": vData(1, i1) = "Low Price"
             Case "09": vData(1, i1) = "High Price"
             Case "10": vData(1, i1) = "Volume"
             Case "11": vData(1, i1) = "Previous Close"
             Case "12": vData(1, i1) = "Previous RSI"
             Case Else: vData(1, i1) = "--"
             End Select
          Next i1
      GoTo ErrorExit
      End If
   
   Dim vCQ As Variant, vHQ As Variant, vRSI As Variant
   vCQ = RCHGetYahooQuotes(pTicker, "l1baoghvd1t1")
   'vHQ = RCHGetYahooHistory2(pTicker, , , , , , , , , , 1, 1, 22, 6)
   vHQ = smfGetYahooHistory(pTicker, Int(Now) - 40, , , , , 1, 22, 6)
   vRSI = SMFTech(vHQ, "RSI", 2)
   
   For i1 = 1 To kItems
       Select Case Mid(pItems, 2 * i1 - 1, 2)
          Case "01": vData(1, i1) = 100 - 100 / (1 + (vRSI(22, 5) + Application.WorksheetFunction.Max(vCQ(1, 1) - vHQ(22, 5), 0)) / (vRSI(22, 6) + Application.WorksheetFunction.Max(0, vHQ(22, 5) - vCQ(1, 1))))
          Case "02": vData(1, i1) = IIf(pLoTrigger > vRSI(22, 1), "--", vHQ(22, 5) + vRSI(22, 6) - vRSI(22, 5) * (100 - pLoTrigger) / pLoTrigger)
          Case "03": vData(1, i1) = IIf(pHiTrigger < vRSI(22, 1), "--", vHQ(22, 5) - vRSI(22, 5) + vRSI(22, 6) * pHiTrigger / (100 - pHiTrigger))
          Case "04": vData(1, i1) = vCQ(1, 1)
          Case "05": vData(1, i1) = vCQ(1, 2)
          Case "06": vData(1, i1) = vCQ(1, 3)
          Case "07": vData(1, i1) = vCQ(1, 4)
          Case "08": vData(1, i1) = vCQ(1, 5)
          Case "09": vData(1, i1) = vCQ(1, 6)
          Case "10": vData(1, i1) = vCQ(1, 7)
          Case "11": vData(1, i1) = vHQ(22, 5)
          Case "12": vData(1, i1) = vRSI(22, 1)
          Case Else: vData(1, i1) = "--"
          End Select
       Next i1
  
ErrorExit:
   
   smfRSITargetPrices = vData
   
   End Function




