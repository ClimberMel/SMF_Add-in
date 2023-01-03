Attribute VB_Name = "smfConvertYahooTicker_"
Function smfConvertYahooTicker(ByVal pTicker As String, _
                               ByVal pSource As String) As String
                         
   '-----------------------------------------------------------------------------------------------------------*
   ' User defined function to convert a Yahoo ticker symbol into another data provider's ticker symbol
   '-----------------------------------------------------------------------------------------------------------*
   ' 2009.12.02 -- Created function
   ' 2009.12.24 -- Add ".OB" translations
   ' 2012.05.13 -- Add ".V" translations
   ' 2012.07.13 -- Add "-" conversion to "." for Canadian ticker symbols going from Yahoo to MSN
   ' 2013.10.26 -- Add "-" conversion to "." for going from Yahoo to Zacks
   ' 2017.06.26 -- Correction to length of ".V" translation
   '-----------------------------------------------------------------------------------------------------------*
   ' > Example of an invocation:
   '
   '   =smfConvertYahooTicker("CMG.TO", "MSN2")
   '-----------------------------------------------------------------------------------------------------------*

   Dim sTicker As String
   Const sYahooxNDX = "~^DJI  "
   Const sGoogleNDX = "~.DJI  "
   Const sMSNxxxNDX = "~$INDU "
                               
   sTicker = Trim(UCase(pTicker))
   Select Case True
      Case Left(UCase(pSource), 5) = "ADVFN": GoTo Yahoo2AdvFN
      Case Left(UCase(pSource), 8) = "BARCHART": GoTo Yahoo2BarChart
      Case Left(UCase(pSource), 8) = "EARNINGS": GoTo Yahoo2Earnings
      Case Left(UCase(pSource), 3) = "MSN": GoTo Yahoo2MSN
      Case Left(UCase(pSource), 6) = "GOOGLE": GoTo Yahoo2Google
      Case Left(UCase(pSource), 11) = "MORNINGSTAR": GoTo Yahoo2MorningStar
      Case Left(UCase(pSource), 7) = "REUTERS": GoTo Yahoo2Reuters
      Case Left(UCase(pSource), 11) = "STOCKCHARTS": GoTo Yahoo2StockCharts
      Case Left(UCase(pSource), 10) = "STOCKHOUSE": GoTo Yahoo2StockHouse
      Case Left(UCase(pSource), 5) = "ZACKS": GoTo Yahoo2Zacks
      End Select
   GoTo ExitFunction
      
Yahoo2AdvFN:
   Select Case True
      Case Right(sTicker, 3) = ".OB": sTicker = "USBB:" & Replace(sTicker, ".OB", "")
      Case Right(sTicker, 3) = ".TO": sTicker = "TSE:" & Replace(sTicker, ".TO", "")
      Case Right(sTicker, 2) = ".V": sTicker = "TSE:" & Replace(sTicker, ".V", "")
      End Select
   GoTo ExitFunction
      
Yahoo2BarChart:
   Select Case True
      Case Right(sTicker, 3) = ".OB": sTicker = Replace(sTicker, ".OB", "")
      'Case Right(sTicker, 3) = ".TO"
      End Select
   GoTo ExitFunction
      
Yahoo2Earnings:
   Select Case True
      Case Right(sTicker, 3) = ".OB": sTicker = Replace(sTicker, ".OB", "")
      'Case Right(sTicker, 3) = ".TO"
      End Select
   GoTo ExitFunction
     
Yahoo2Google:
   Select Case True
      Case Right(sTicker, 3) = ".OB": sTicker = "OTC:" & Replace(sTicker, ".OB", "")
      Case Right(sTicker, 3) = ".TO": sTicker = "CVE:" & Replace(sTicker, ".TO", "")
      Case Right(sTicker, 2) = ".V": sTicker = "TSE:" & Replace(sTicker, ".V", "")
      Case InStr(sYahooxNDX, "~" & sTicker) > 0: sTicker = Trim(Mid(sGoogleNDX, InStr(sYahooxNDX, "~" & sTicker) + 1, 5))
      End Select
   GoTo ExitFunction

Yahoo2MorningStar:
   Select Case True
      Case Right(sTicker, 3) = ".OB": sTicker = Replace(sTicker, ".OB", "")
      Case Right(sTicker, 3) = ".TO": sTicker = "XTSE:" & Replace(sTicker, ".TO", "")
      End Select
   GoTo ExitFunction
      
Yahoo2MSN:
   Select Case True
      Case Right(sTicker, 3) = ".OB": sTicker = Replace(sTicker, ".OB", "")
      Case Right(sTicker, 3) = ".TO": sTicker = "CA:" & Replace(Replace(Replace(sTicker, ".TO", ""), "-P", "."), "-", ".")
      Case Right(sTicker, 2) = ".V": sTicker = "CA:" & Replace(sTicker, ".V", "")
      Case Right(sTicker, 2) = ".X": sTicker = "." & Replace(sTicker, ".X", "")
      Case InStr(sTicker, "-P") > 0: sTicker = Replace(sTicker, "-P", "-")
      Case InStr(sTicker, "-") > 0: sTicker = Replace(sTicker, "-", "/")
      Case InStr(sYahooxNDX, "~" & sTicker) > 0: sTicker = Trim(Mid(sMSNxxxNDX, InStr(sYahooxNDX, "~" & sTicker) + 1, 5))
      End Select
   GoTo ExitFunction
      
Yahoo2Reuters:
   Select Case True
      'Case Right(sTicker, 3) = ".OB"
      'Case Right(sTicker, 3) = ".TO"
      End Select
   GoTo ExitFunction
      
Yahoo2StockCharts:
   Select Case True
      Case Right(sTicker, 3) = ".OB": sTicker = Replace(sTicker, ".OB", "")
      'Case Right(sTicker, 3) = ".TO"
      End Select
   GoTo ExitFunction

Yahoo2StockHouse:
   Select Case True
      Case Right(sTicker, 3) = ".OB": sTicker = Replace(sTicker, ".OB", "")
      Case Right(sTicker, 3) = ".TO": sTicker = "T." & Replace(sTicker, ".TO", "")
      End Select
   GoTo ExitFunction

Yahoo2Zacks:
   Select Case True
      Case Right(sTicker, 3) = ".OB": sTicker = Replace(sTicker, ".OB", "")
      Case Right(sTicker, 3) = ".TO": sTicker = "T." & Replace(sTicker, ".TO", "")
      Case InStr(sTicker, "-") > 0: sTicker = Replace(sTicker, "-", ".")
      End Select
   GoTo ExitFunction
                                
ExitFunction:
   smfConvertYahooTicker = sTicker
   End Function


