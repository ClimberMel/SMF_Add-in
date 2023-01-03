Attribute VB_Name = "modGetOptionQuotes"
Option Explicit
Public Function smfGetOptionQuotes(ByVal pTickers As Variant, _
                                   ByVal pItems As Variant, _
                          Optional ByVal pHeader As Integer = 0, _
                          Optional ByVal pSource As String = "Y", _
                          Optional ByVal pDim1 As Integer = 0, _
                          Optional ByVal pDim2 As Integer = 0)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get multiple data items for multiple options from Yahoo or MSN
    '-----------------------------------------------------------------------------------------------------------*
    ' 2010.06.02 -- Created function
    ' 2010.07.16 -- Added ability to send "m/d" as the period designation for the ticker symbol
    ' 2010.07.19 -- Added ability to use 2-digit item codes and retrieve data from multiple sources
    ' 2010.09.10 -- Add "u" data item for last traded price of the underlying equity
    ' 2011.04.01 -- Add Google as a data source
    ' 2011.04.07 -- Change to allow a blank option ticker symbol so output is synchronized with input
    ' 2011.04.07 -- Allow an array to be sent for pTickers (e.g. from another function)
    ' 2011.04.28 -- Change cDec() to smfConvertData()
    ' 2011.11.30 -- Allow cells of input ticker range to contain multiple ticker symbols
    ' 2011.11.30 -- Add OX3 data source
    ' 2012.01.29 -- Return null value for a null pItem and pTicker parameters
    ' 2012.02.14 -- Add "7" data item for alpha
    ' 2015.02.21 -- Add "W2" to "W7" and "M2" to "M5" date choices
    ' 2015.08.13 -- Add NASDAQ as a data source
    ' 2016.08.05 -- Obsolete MSN and MW as data sources
    ' 2016.12.03 -- Add BarChart as a data source
    ' 2017.08.06 -- Restore MW as a data source
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for IBM and MMM:
    '
    '   =smfGetOptionQuotes("SPY Jun 2010 $110 Call+SPY Jun 2010 $120 Call","ba")
    '   =smfGetOptionQuotes("SPY Jun 2010 $110 Call+SPY Jun 2010 $120 Call","ba",1,"MSN")
    '   =smfGetOptionQuotes("SPY Jun 2010 $110 Call+SPY Jun 2010 $120 Call","1b1a32",1,2)
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim iRow As Integer, iCol As Integer
    On Error GoTo ErrorExit
    
    '------------------> Process possible list of option ticker symbols
    Dim sTickers As String, aTickers() As String
    Dim oCell As Object
    Select Case TypeName(pTickers)
        Case "String"
             If pTickers = "None" Then
                smfGetOptionQuotes = ""
                Exit Function
                End If
             sTickers = Replace(pTickers, ",", "+")
        Case "Variant()"
             sTickers = ""
             For iRow = 1 To UBound(pTickers)
                 sTickers = sTickers & pTickers(iRow, 1) & "+"
                 Next iRow
             sTickers = Left(sTickers, Len(sTickers) - 1)
        Case "Range"
             sTickers = ""
             For Each oCell In pTickers
                 sTickers = sTickers & Replace(oCell.Value, ",", "+") & "+"
                 Next oCell
             sTickers = Left(sTickers, Len(sTickers) - 1)
        Case Else
            smfGetOptionQuotes = ""
            Exit Function
        End Select
    aTickers = Split(sTickers, "+")
    
    '------------------> Process possible list of option item codes
    Dim sItems As String, iStep As Integer
    pSource = UCase(pSource)
    If pSource = "2" Then iStep = 2 Else iStep = 1
    Select Case VarType(pItems)
        Case vbString
             sItems = Replace(pItems, " ", "")
        Case Is >= 8192
             sItems = ""
             For Each oCell In pItems
                 sItems = sItems & oCell.Value
                 Next oCell
        Case Else
            smfGetOptionQuotes = ""
            Exit Function
        End Select
    sItems = UCase(sItems)
    
    '------------------> Determine size of array to return
    Dim kDim1 As Integer, kDim2 As Integer
    kDim1 = pDim1  ' Rows
    kDim2 = pDim2  ' Columns
    If kDim1 = 0 Or kDim2 = 0 Then
       If kDim1 = 0 Then kDim1 = 1
       If kDim2 = 0 Then kDim2 = Len(sItems) / iStep
       On Error Resume Next
       kDim1 = Application.Caller.Rows.Count
       kDim2 = Application.Caller.Columns.Count
       On Error GoTo ErrorExit
       End If
  
    '------------------> Initialize return array
    ReDim vData(1 To kDim1, 1 To kDim2) As Variant
    For iRow = 1 To kDim1
        For iCol = 1 To kDim2
            vData(iRow, iCol) = ""
            Next iCol
        Next iRow
    
    '------------------> Create headings if requested
    Dim iPtr As Integer
    If pHeader <> 1 Then
       pHeader = 0
    Else
       iCol = 0
       For iPtr = iStep To Len(sItems) Step iStep
           iCol = iCol + 1
           If iCol > kDim2 Then Exit For
           Select Case Mid(sItems, iPtr, 1)
              Case "%": vData(1, iCol) = "% Change"
              Case "A": vData(1, iCol) = "Ask Price"
              Case "B": vData(1, iCol) = "Bid Price"
              Case "C": vData(1, iCol) = "$ Change"
              Case "E": vData(1, iCol) = "Bid Size"
              Case "F": vData(1, iCol) = "Ask Size"
              Case "G": vData(1, iCol) = "Daily Low"
              Case "H": vData(1, iCol) = "Daily High"
              Case "I": vData(1, iCol) = "Open Interest"
              Case "J": vData(1, iCol) = "Contract Low"
              Case "K": vData(1, iCol) = "Contract High"
              Case "L": vData(1, iCol) = "Last Price"
              Case "O": vData(1, iCol) = "Open"
              Case "P": vData(1, iCol) = "Previous Close"
              Case "S": vData(1, iCol) = "Strike Price"
              Case "T": vData(1, iCol) = "Last Trade Time"
              Case "U": vData(1, iCol) = "Underlying Price"
              Case "V": vData(1, iCol) = "Volume"
              Case "X": vData(1, iCol) = "Expiry"
              Case "Y": vData(1, iCol) = "Time/Theo Value"
              Case "Z": vData(1, iCol) = "Ticker Symbol"
              Case "1": vData(1, iCol) = "Vega"
              Case "2": vData(1, iCol) = "Theta"
              Case "3": vData(1, iCol) = "Rho"
              Case "4": vData(1, iCol) = "Gamma"
              Case "5": vData(1, iCol) = "Delta"
              Case "6": vData(1, iCol) = "Implied Volatility"
              Case "7": vData(1, iCol) = "Alpha"
              Case "8": vData(1, iCol) = "Net"
              Case "9": vData(1, iCol) = "Tick"
              End Select
           Next iPtr
       End If
    
    
    '------------------> Get each individual quote item
    Dim aParts() As String
    Dim vStrike As Variant, vExpiry As Variant
    Dim sType As String, iMonth As Integer, iYear As Integer, sItem As String, sChoice As String
    For iRow = 0 To UBound(aTickers)
        If iRow > kDim1 - 1 Then Exit For
        If aTickers(iRow) = "" Then GoTo NextTicker
        aParts = Split(aTickers(iRow), " ")
        If InStr(aParts(1), "/") > 0 Then
           vExpiry = DateValue(aParts(1) & "/" & aParts(2))
        Else
           Select Case UCase(aParts(1))
              Case "Q1", "Q2", "Q3", "Q4": sType = "Q"
              Case "W", "W1", "WEEK": sType = "W"
              Case "W2" To "W7": sType = UCase(aParts(1))
              Case "M1" To "M5": sType = UCase(aParts(1))
              Case Else: sType = "M"
              End Select
           Select Case UCase(aParts(1))
              Case "JAN": iMonth = 1
              Case "FEB": iMonth = 2
              Case "MAR", "Q1": iMonth = 3
              Case "APR": iMonth = 4
              Case "MAY": iMonth = 5
              Case "JUN", "Q2": iMonth = 6
              Case "JUL": iMonth = 7
              Case "AUG": iMonth = 8
              Case "SEP", "Q3": iMonth = 9
              Case "OCT": iMonth = 10
              Case "NOV": iMonth = 11
              Case "DEC", "Q4": iMonth = 12
              Case 1 To 12: iMonth = aParts(1)
              Case Else: iMonth = Month(Date)
              End Select
           iYear = smfConvertData(aParts(2))
           vExpiry = smfGetOptionExpiry(iYear, iMonth, sType)
           End If
        Select Case True
           Case Left(aParts(3), 3) = "OTM": vStrike = aParts(3)
           Case Left(aParts(3), 3) = "ITM": vStrike = aParts(3)
           Case Else: vStrike = smfConvertData(aParts(3))
           End Select
        iCol = 0
        For iPtr = 1 To Len(sItems) Step iStep
            iCol = iCol + 1
            If iCol > kDim2 Then Exit For
            If pSource = "2" Then
               sChoice = Mid(sItems, iPtr, 1)
               Select Case sChoice
                  Case 1: sChoice = "Y"
                  Case 2: sChoice = "MSN"
                  Case 3: sChoice = "OX"
                  Case 4: sChoice = "MW"
                  Case 5: sChoice = "OX2"
                  Case 6: sChoice = "G"
                  Case 7: sChoice = "OX3"
                  Case 8: sChoice = "8"
                  Case 9: sChoice = "N"
                  End Select
               sItem = Mid(sItems, iPtr + 1, 1)
            Else
               sChoice = pSource
               sItem = Mid(sItems, iPtr, 1)
               End If
            Select Case sChoice
               Case "8"
                    vData(iRow + 1 + pHeader, iCol) = smfGet888OptionQuote(aParts(0), Left(aParts(4), 1), vExpiry, vStrike, sItem)
               Case "Y"
                    vData(iRow + 1 + pHeader, iCol) = smfGetYahooOptionQuote(aParts(0), Left(aParts(4), 1), vExpiry, vStrike, sItem)
               Case "B"
                    vData(iRow + 1 + pHeader, iCol) = smfGetBarChartOptionQuote(aParts(0), Left(aParts(4), 1), vExpiry, vStrike, sItem)
               Case "G"
                    vData(iRow + 1 + pHeader, iCol) = smfGetGoogleOptionQuote(aParts(0), Left(aParts(4), 1), vExpiry, vStrike, sItem)
               Case "MSN"
                    'vData(iRow + 1 + pHeader, iCol) = smfGetMSNOptionQuote(aParts(0), Left(aParts(4), 1), vExpiry, vStrike, sItem)
                    vData(iRow + 1 + pHeader, iCol) = "MSN obsolete"
               Case "MW"
                    vData(iRow + 1 + pHeader, iCol) = smfGetMWOptionQuote(aParts(0), Left(aParts(4), 1), vExpiry, vStrike, sItem)
               Case "N"
                    vData(iRow + 1 + pHeader, iCol) = smfGetNASDAQOptionQuote(aParts(0), Left(aParts(4), 1), vExpiry, vStrike, sItem)
               Case "OX"
                    vData(iRow + 1 + pHeader, iCol) = smfGetOXOptionQuote(aParts(0), Left(aParts(4), 1), vExpiry, vStrike, sItem)
               Case "OX2"
                    vData(iRow + 1 + pHeader, iCol) = smfGetOX2OptionQuote(aParts(0), Left(aParts(4), 1), vExpiry, vStrike, sItem)
               Case "OX3"
                    vData(iRow + 1 + pHeader, iCol) = smfGetOX3OptionQuote(aParts(0), Left(aParts(4), 1), vExpiry, vStrike, sItem)
               Case Else
                    vData(iRow + 1 + pHeader, iCol) = "Bad Source Code: " & pSource
               End Select
            Next iPtr
NextTicker:
        Next iRow
   
ErrorExit:
    smfGetOptionQuotes = vData
    End Function

Public Function smfGet888OptionQuote(ByVal pTicker As Variant, _
                                     ByVal pPutCall As Variant, _
                                     ByVal pExpiry As Variant, _
                                     ByVal pStrike As Variant, _
                                     ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quotes and related data from 888options.com
    '-----------------------------------------------------------------------------------------------------------*
    ' 2012.02.14 -- Created function
    ' 2014.06.12 -- Added ticker symbol to search string because of new "7" options
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
        
    '------------------> Get the special historical volatility data?
    Dim sItem As String, sURL As String
    sURL = "http://oic.ivolatility.com/oic_adv_options.j?exp_date=-1&ticker=" & UCase(pTicker)
    sItem = Trim(UCase(pItem))
    Select Case sItem
       Case "HV10C": smfGet888OptionQuote = RCHGetTableCell(sURL, 1, "1 Wk Ago", "10 days")
       Case "HV10W": smfGet888OptionQuote = RCHGetTableCell(sURL, 2, "1 Wk Ago", "10 days")
       Case "HV10M": smfGet888OptionQuote = RCHGetTableCell(sURL, 3, "1 Wk Ago", "10 days")
       Case "HV10H": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 4, "1 Wk Ago", "10 days"), "|", " "))
       Case "HV10HD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 4, "1 Wk Ago", "10 days") & "|", " - ", "|")
       Case "HV10L": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 5, "1 Wk Ago", "10 days"), "|", " "))
       Case "HV10LD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 5, "1 Wk Ago", "10 days") & "|", " - ", "|")
       Case "HV20C": smfGet888OptionQuote = RCHGetTableCell(sURL, 1, "1 Wk Ago", "20 days")
       Case "HV20W": smfGet888OptionQuote = RCHGetTableCell(sURL, 2, "1 Wk Ago", "20 days")
       Case "HV20M": smfGet888OptionQuote = RCHGetTableCell(sURL, 3, "1 Wk Ago", "20 days")
       Case "HV20H": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 4, "1 Wk Ago", "20 days"), "|", " "))
       Case "HV20HD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 4, "1 Wk Ago", "20 days") & "|", " - ", "|")
       Case "HV20L": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 5, "1 Wk Ago", "20 days"), "|", " "))
       Case "HV20LD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 5, "1 Wk Ago", "20 days") & "|", " - ", "|")
       Case "HV30C": smfGet888OptionQuote = RCHGetTableCell(sURL, 1, "1 Wk Ago", "30 days")
       Case "HV30W": smfGet888OptionQuote = RCHGetTableCell(sURL, 2, "1 Wk Ago", "30 days")
       Case "HV30M": smfGet888OptionQuote = RCHGetTableCell(sURL, 3, "1 Wk Ago", "30 days")
       Case "HV30H": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 4, "1 Wk Ago", "30 days"), "|", " "))
       Case "HV30HD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 4, "1 Wk Ago", "30 days") & "|", " - ", "|")
       Case "HV30L": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 5, "1 Wk Ago", "30 days"), "|", " "))
       Case "HV30LD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 5, "1 Wk Ago", "30 days") & "|", " - ", "|")
       Case "IVICC": smfGet888OptionQuote = RCHGetTableCell(sURL, 1, "1 Wk Ago", "IV Index Call")
       Case "IVICW": smfGet888OptionQuote = RCHGetTableCell(sURL, 2, "1 Wk Ago", "IV Index Call")
       Case "IVICM": smfGet888OptionQuote = RCHGetTableCell(sURL, 3, "1 Wk Ago", "IV Index Call")
       Case "IVICH": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 4, "1 Wk Ago", "IV Index Call"), "|", " "))
       Case "IVICHD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 4, "1 Wk Ago", "IV Index Call") & "|", " - ", "|")
       Case "IVICL": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 5, "1 Wk Ago", "IV Index Call"), "|", " "))
       Case "IVICLD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 5, "1 Wk Ago", "IV Index Call") & "|", " - ", "|")
       Case "IVIPC": smfGet888OptionQuote = RCHGetTableCell(sURL, 1, "1 Wk Ago", "IV Index Put")
       Case "IVIPW": smfGet888OptionQuote = RCHGetTableCell(sURL, 2, "1 Wk Ago", "IV Index Put")
       Case "IVIPM": smfGet888OptionQuote = RCHGetTableCell(sURL, 3, "1 Wk Ago", "IV Index Put")
       Case "IVIPH": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 4, "1 Wk Ago", "IV Index Put"), "|", " "))
       Case "IVIPHD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 4, "1 Wk Ago", "IV Index Put") & "|", " - ", "|")
       Case "IVIPL": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 5, "1 Wk Ago", "IV Index Put"), "|", " "))
       Case "IVIPLD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 5, "1 Wk Ago", "IV Index Put") & "|", " - ", "|")
       Case "IVIMC": smfGet888OptionQuote = RCHGetTableCell(sURL, 1, "1 Wk Ago", "IV Index Mean")
       Case "IVIMW": smfGet888OptionQuote = RCHGetTableCell(sURL, 2, "1 Wk Ago", "IV Index Mean")
       Case "IVIMM": smfGet888OptionQuote = RCHGetTableCell(sURL, 3, "1 Wk Ago", "IV Index Mean")
       Case "IVIMH": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 4, "1 Wk Ago", "IV Index Mean"), "|", " "))
       Case "IVIMHD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 4, "1 Wk Ago", "IV Index Mean") & "|", " - ", "|")
       Case "IVIML": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 5, "1 Wk Ago", "IV Index Mean"), "|", " "))
       Case "IVIMLD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 5, "1 Wk Ago", "IV Index Mean") & "|", " - ", "|")
       Case "HC30C": smfGet888OptionQuote = RCHGetTableCell(sURL, 1, "1 Wk Ago", "30 days", "30 days")
       Case "HC30W": smfGet888OptionQuote = RCHGetTableCell(sURL, 2, "1 Wk Ago", "30 days", "30 days")
       Case "HC30M": smfGet888OptionQuote = RCHGetTableCell(sURL, 3, "1 Wk Ago", "30 days", "30 days")
       Case "HC30H": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 4, "1 Wk Ago", "30 days", "30 days"), "|", " "))
       Case "HC30HD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 4, "1 Wk Ago", "30 days", "30 days") & "|", " - ", "|")
       Case "HC30L": smfGet888OptionQuote = smfConvertData(smfStrExtr("|" & RCHGetTableCell(sURL, 5, "1 Wk Ago", "30 days", "30 days"), "|", " "))
       Case "HC30LD": smfGet888OptionQuote = smfStrExtr(RCHGetTableCell(sURL, 5, "1 Wk Ago", "30 days", "30 days") & "|", " - ", "|")
       Case Else: smfGet888OptionQuote = ""
       End Select
    If smfGet888OptionQuote <> "" Then Exit Function
        
    '------------------> Verify the pPutCall parameter
    Dim sPutCall As String
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P"
       Case "C"
       Case Else
            smfGet888OptionQuote = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
    
    '------------------> Verify pExpiry and pStrike
    Dim sFind1 As String
    pStrike = Trim(UCase(pStrike))
    pTicker = Trim(UCase(pTicker))
    Select Case True
       Case Not (VarType(pExpiry) = vbDouble Or IsDate(pExpiry))
            smfGet888OptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       Case IsNumeric(pStrike)
            sFind1 = Left(pTicker & "      ", 6) & Format(pExpiry, "yymmdd") & sPutCall & Format(1000 * pStrike, "00000000")
       Case Else
            smfGet888OptionQuote = "Bad strike price: " & pStrike
            Exit Function
       End Select
       
    '------------------> Verify the pItem parameter and set the # of cells to skip
    Dim iCells As Integer
    Select Case sItem
       Case "A": iCells = 3       ' Ask price
       Case "B": iCells = 2       ' Bid price
       Case "C": iCells = 4       ' $ Change
       Case "I": iCells = 6       ' Open Interest
       Case "L": iCells = 1       ' Bid/Ask Mean
       Case "V": iCells = 5       ' Volume
       Case "6": iCells = 7       ' Implied Volatility
       Case "%": iCells = 4       ' % Change
       Case "7": iCells = 11      ' Alpha
       Case "5": iCells = 8       ' Delta
       Case "4": iCells = 9       ' Gamma
       Case "3": iCells = 13      ' Rho
       Case "2": iCells = 10      ' Theta
       Case "1": iCells = 12      ' Vega
       Case "Z": iCells = 0       ' 888 ticker symbol
       Case "X": iCells = 0       ' Expiration date
       Case "S": iCells = 0       ' Strike price
       Case "U": iCells = 0       ' Last price of underlying equity
       Case Else
            smfGet888OptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
    
    '------------------> Find data item
    Dim s1 As String
    Select Case sItem
       Case "C"
            s1 = RCHGetTableCell(sURL, iCells, sFind1)
            smfGet888OptionQuote = 0 + smfStrExtr("|" & s1, "|", "(")
       Case "%"
            s1 = RCHGetTableCell(sURL, iCells, sFind1)
            smfGet888OptionQuote = smfStrExtr(s1, "(", ")") / 100
       Case "S": smfGet888OptionQuote = 0 + pStrike
       Case "U": smfGet888OptionQuote = RCHGetTableCell(sURL, 1, ">Price", "<tr")
       Case "X": smfGet888OptionQuote = pExpiry
       Case "Z": smfGet888OptionQuote = UCase(Trim(pTicker)) & " " & sFind1
       Case Else: smfGet888OptionQuote = RCHGetTableCell(sURL, iCells, sFind1)
       End Select
    Exit Function

ErrorExit:
    smfGet888OptionQuote = "Error"
    
    End Function
Public Function smfGetBarChartOptionQuote(ByVal pTicker As Variant, _
                                          ByVal pPutCall As Variant, _
                                          ByVal pExpiry As Variant, _
                                          ByVal pStrike As Variant, _
                                          ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quotes from BarChart
    '-----------------------------------------------------------------------------------------------------------*
    ' 2016.12.03 -- Created function
    ' 2017.08.21 -- Change to allow strike prices of $1000 or more
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for SPY:
    '
    '   =smfGetBarChartOptionQuote("SPY","C",DATE(2012,6,1),110,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
        
    '------------------> Verify the pPutCall parameter
    Dim sPutCall As String
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P": sPutCall = "Put"
       Case "C": sPutCall = "Call"
       Case Else
            smfGetBarChartOptionQuote = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
    
    '------------------> Verify pExpiry and pStrike
    Dim sFind1 As String, s1 As String
    pStrike = Trim(UCase(pStrike))
    pTicker = Trim(UCase(pTicker))
    Select Case True
       Case Not (VarType(pExpiry) = vbDouble Or IsDate(pExpiry))
            smfGetBarChartOptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       Case IsNumeric(pStrike)
            sFind1 = """optionType"":""" & sPutCall & """,""strikePrice"":""" & Format(pStrike, "#,##0.00") & """"
       Case Else
            smfGetBarChartOptionQuote = "Bad strike price: " & pStrike
            Exit Function
       End Select
    
    '------------------> Verify the pItem parameter
    Dim sFind2 As String
    pItem = Trim(UCase(pItem))
    Select Case pItem
       Case "%": sFind2 = "percentChange"
       Case "A": sFind2 = "askPrice"
       Case "B": sFind2 = "bidPrice"
       Case "C": sFind2 = "priceChange"
       Case "I": sFind2 = "openInterest"
       Case "L": sFind2 = "lastPrice"
       Case "S": sFind2 = "strikePrice"
       Case "U": sFind2 = ""       ' Last price of underlying equity
       Case "V": sFind2 = "volume"
       Case "X": sFind2 = "expirationDate"
       Case "Y": sFind2 = "theoretical"
       Case "1": sFind2 = "vega"
       Case "2": sFind2 = "theta"
       Case "3": sFind2 = "rho"
       Case "4": sFind2 = "gamma"
       Case "5": sFind2 = "delta"
       Case "6": sFind2 = "volatility"
       Case Else
            smfGetBarChartOptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
    
    '------------------> Find data item
    Dim sURL As String
    sURL = "https://core-api.barchart.com/v1/options/chain?fields=" & _
           "optionType,strikePrice,lastPrice,percentFromLast,priceChange,percentChange,bidPrice,midpoint,askPrice,theoretical,volatility,delta,gamma,rho,theta,vega,volume,openInterest,daysToExpiration,expirationDate" & _
           "&groupBy=optionType&raw=0&symbol=" & pTicker & "&expirationDate=" & Format(pExpiry, "yyyy-mm-dd")
    Select Case pItem
       Case "U": smfGetBarChartOptionQuote = RCHGetYahooQuotes(pTicker, "l1")(1, 1)
       Case Else
            s1 = smfStrExtr(RCHGetWebData(sURL, sFind1), "~", "}")
            smfGetBarChartOptionQuote = smfConvertData(Replace(smfStrExtr(s1, """" & sFind2 & """:""", """"), "\", ""))
       End Select
    Exit Function

ErrorExit:
    smfGetBarChartOptionQuote = "Error"

    End Function
Public Function smfGetBigChartsOptionQuote(ByVal pTicker As Variant, _
                                           ByVal pPutCall As Variant, _
                                           ByVal pExpiry As Variant, _
                                           ByVal pStrike As Variant, _
                                           ByVal pItem As Variant, _
                                  Optional ByVal pHistory As Variant = "")
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quotes from BigCharts
    '-----------------------------------------------------------------------------------------------------------*
    ' 2012.03.02 -- Created function
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for SPY:
    '
    '   =smfGetBigChartsOptionQuote("SPY","C",DATE(2012,6,1),110,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
        
    '------------------> Verify the pPutCall parameter
    Dim sPutCall As String, iOffset As Integer
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P": iOffset = 12
       Case "C": iOffset = 0
       Case Else
            smfGetBigChartsOptionQuote = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
    
    '------------------> Verify pExpiry and pStrike
    Dim sTicker As String, s1 As String, s2 As String
    pStrike = Trim(UCase(pStrike))
    pTicker = Trim(UCase(pTicker))
    Select Case True
       Case Not (VarType(pExpiry) = vbDouble Or IsDate(pExpiry))
            smfGetBigChartsOptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       Case IsNumeric(pStrike)
            'If InStr(pTicker, ":") > 0 Then s1 = smfStrExtr(pTicker & "|", ":", "|") Else s1 = pTicker
            sTicker = pTicker & Chr(64 + Month(pExpiry) + iOffset) & Format(pExpiry, "ddyy")
            If pStrike < 100 Then
               sTicker = sTicker & Format(10000 * pStrike, "4000000")
            Else
               sTicker = sTicker & Format(1000 * pStrike, "3000000")
               End If
       Case Else
            smfGetBigChartsOptionQuote = "Bad strike price: " & pStrike
            Exit Function
       End Select
    
    '------------------> Historical quote request?
    If pHistory <> "" Then GoTo BC_History
    
    '------------------> Verify the pItem parameter
    Dim iCells As Integer
    pItem = Trim(UCase(pItem))
    Select Case pItem
       Case "A": iCells = 5      ' Ask price
       Case "B": iCells = 4      ' Bid price
       Case "C": iCells = 2      ' $ Change
       Case "I": iCells = 6      ' Open Interest
       Case "L": iCells = 1      ' Last price
       Case "S": iCells = 0      ' Strike price
       Case "U": iCells = 0      ' Last price of underlying equity
       Case "V": iCells = 3      ' Volume
       Case "X": iCells = 0      ' Option expiration date
       Case "Z": iCells = 0      ' Option ticker symbol
       Case Else
            smfGetBigChartsOptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
    
    '------------------> Find data item
    Dim sURL As String
    sURL = "http://bigcharts.marketwatch.com/quickchart/options.asp?showAll=True&symb=" & pTicker
    Select Case pItem
       Case "S": smfGetBigChartsOptionQuote = pStrike
       Case "U": smfGetBigChartsOptionQuote = RCHGetYahooQuotes(pTicker, "l1")(1, 1)
       Case "X": smfGetBigChartsOptionQuote = pExpiry
       Case "Z": smfGetBigChartsOptionQuote = sTicker
       Case Else
            smfGetBigChartsOptionQuote = RCHGetTableCell(sURL, iCells, sTicker)
       End Select
    Exit Function

BC_History:
    
    '------------------> Create URL
    sURL = "http://bigcharts.marketwatch.com/historical/default.asp?symb=" & sTicker & _
           "&closeDate=" & Format(pHistory, "mm/dd/yy")
    
    '------------------> Verify the pItem parameter
    Dim sFindIt As String
    pItem = Trim(UCase(pItem))
    Select Case pItem
       Case "D": smfGetBigChartsOptionQuote = smfGetTagContent(sURL, "div", 2, "enddate=", "<tr")
       Case "G": smfGetBigChartsOptionQuote = RCHGetTableCell(sURL, 1, "Low:")
       Case "H": smfGetBigChartsOptionQuote = RCHGetTableCell(sURL, 1, "High:")
       Case "L": smfGetBigChartsOptionQuote = RCHGetTableCell(sURL, 1, "Closing Price:")
       Case "O": smfGetBigChartsOptionQuote = RCHGetTableCell(sURL, 1, "Open:")
       Case "V": smfGetBigChartsOptionQuote = RCHGetTableCell(sURL, 1, "Volume:")
       Case "Z": smfGetBigChartsOptionQuote = smfGetTagContent(sURL, "div", 1, "enddate=", "<tr")
       Case Else
            smfGetBigChartsOptionQuote = "Unrecognized item ID: " & pItem
       End Select
    Exit Function

ErrorExit:
    smfGetBigChartsOptionQuote = "Error"

    End Function
Public Function smfGetGoogleOptionQuote(ByVal pTicker As Variant, _
                                        ByVal pPutCall As Variant, _
                                        ByVal pExpiry As Variant, _
                                        ByVal pStrike As Variant, _
                                        ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quotes from Google
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.04.01 -- Created function
    ' 2011.04.02 -- Added code to strip exchange from ticker symbol
    ' 2014.08.17 -- Allow "Put" or "Call" for pPutCall parameter
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for SPY:
    '
    '   =smfGetGoogleOptionQuote("SPY","C",DATE(2012,6,1),110,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
        
    '------------------> Verify the pPutCall parameter
    Dim sPutCall As String
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P"
       Case "C"
       Case Else
            smfGetGoogleOptionQuote = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
    
    '------------------> Verify pExpiry and pStrike
    Dim sTicker As String, s1 As String
    pStrike = Trim(UCase(pStrike))
    pTicker = Trim(UCase(pTicker))
    Select Case True
       Case Not (VarType(pExpiry) = vbDouble Or IsDate(pExpiry))
            smfGetGoogleOptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       Case IsNumeric(pStrike)
            If InStr(pTicker, ":") > 0 Then s1 = smfStrExtr(pTicker & "|", ":", "|") Else s1 = pTicker
            sTicker = s1 & Format(pExpiry, "yymmdd") & sPutCall & Format(1000 * pStrike, "00000000")
       Case Else
            smfGetGoogleOptionQuote = "Bad strike price: " & pStrike
            Exit Function
       End Select
    
    '------------------> Verify the pItem parameter
    Dim sFindIt As String
    pItem = Trim(UCase(pItem))
    Select Case pItem
       Case "%": sFindIt = "cp"     ' % change
       Case "A": sFindIt = "a"      ' Ask price
       Case "B": sFindIt = "b"      ' Bid price
       Case "C": sFindIt = "c"      ' $ Change
       Case "I": sFindIt = "oi"     ' Open Interest
       Case "L": sFindIt = "p"      ' Last price
       Case "S": sFindIt = "strike" ' Strike price
       Case "U": sFindIt = ""       ' Last price of underlying equity
       Case "V": sFindIt = "vol"    ' Volume
       Case "X": sFindIt = "expiry" ' Option expiration date
       Case "Z": sFindIt = ""       ' Option ticker symbol
       Case Else
            smfGetGoogleOptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
    
    '------------------> Find data item
    Dim sURL As String
    sURL = "http://www.google.com/finance/option_chain?output=json&q=" & pTicker & _
           "&expd=" & Day(pExpiry) & _
           "&expm=" & Month(pExpiry) & _
           "&expy=" & Year(pExpiry)
    Select Case pItem
       Case "U": smfGetGoogleOptionQuote = RCHGetYahooQuotes(pTicker, "l1")(1, 1)
       Case "Z": smfGetGoogleOptionQuote = sTicker
       Case Else
            s1 = smfStrExtr(RCHGetWebData(sURL, sTicker, 200), "," & sFindIt & ":""", """")
            smfGetGoogleOptionQuote = smfConvertData(s1)
       End Select
    Exit Function

ErrorExit:
    smfGetGoogleOptionQuote = "Error"

    End Function

Public Function smfGetMSNOptionQuote(ByVal pTicker As Variant, _
                                     ByVal pPutCall As Variant, _
                                     ByVal pExpiry As Variant, _
                                     ByVal pStrike As Variant, _
                                     ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quote from Yahoo
    '-----------------------------------------------------------------------------------------------------------*
    ' 2010.04.07 -- Created function
    ' 2010.09.10 -- Add "u" data item for last traded price of the underlying equity
    ' 2011.11.30 -- Change "u" data item to be able to pick up after hours price
    ' 2012.02.17 -- Change URL for option quotes
    ' 2013.01.04 -- Remove day 1, 30, and 31 assumptions for option expiration date
    ' 2013.12.21 -- Fix sPutCall processing to only use 1st byte of parameter
    ' 2014.08.15 -- Attempted to update to their new option ticker symbol structure
    ' 2014.12.23 -- Obsoleted function
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for IBM and MMM:
    '
    '   =smfGetMSNQuotes("SPY","C",DATE(2012,12,22),65,"b")
    '-----------------------------------------------------------------------------------------------------------*
        
    smfGetMSNOptionQuote = "Obsolete"
    Exit Function
    
    On Error GoTo ErrorExit
    Dim sPutCall As String, sStrike As String, sItem As String
    Dim sURL As String, sFind1 As String, sFind2 As String, sYr As String, sMon As String, sDay As String
    Dim sLabel As String, iCells As Integer, iRows As Integer
    Dim iYear As Integer, iMonth As Integer, iExpiry As Date, iOffset As Integer
    
    '------------------> Verify the pPutCall parameter
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P"
            iOffset = Asc("M") - 1
       Case "C"
            iOffset = Asc("A") - 1
       Case Else
            smfGetMSNOptionQuote = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
       
    '------------------> Verify the pItem parameter and set the # of cells to skip
    sItem = Trim(UCase(pItem))
    Select Case sItem
       Case "X"                                          ' Expiration date
       Case "S": iCells = -10: sLabel = ""               ' Strike price
       Case "Z": iCells = 0: sLabel = ""                 ' MSN ticker symbol
       Case "L": iCells = 1: sLabel = ""                 ' Last price
       Case "C": iCells = 2: sLabel = ""                 ' $ Change
       Case "%": iCells = 3: sLabel = ""                 ' % Change
       Case "Y": iCells = 4: sLabel = ""                 ' Time Value
       Case "B": iCells = 5: sLabel = ""                 ' Bid price
       Case "A": iCells = 6: sLabel = ""                 ' Ask price
       Case "U": iCells = 0: sLabel = ""                 ' Last price of underlying equity
       Case "V": iCells = 7: sLabel = ""                 ' Volume
       Case "I": iCells = 8: sLabel = ""                 ' Open Interest
       Case "G": iCells = 0: sLabel = ">Day's Low"       ' Daily low
       Case "H": iCells = 0: sLabel = ">Day's High"      ' Daily high
       Case "E": iCells = 0: sLabel = ">Bid Size"        ' Bid size
       Case "F": iCells = 0: sLabel = ">Ask Size"        ' Ask size
       Case "O": iCells = 0: sLabel = ">Open"            ' Open
       Case "P": iCells = 0: sLabel = ">Previous Close"  ' Previous close
       Case "T": iCells = 0: sLabel = ">Last Trade"      ' Last trade time
       Case Else
            smfGetMSNOptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
       
    '------------------> Handle special strike price strings
    sStrike = Trim(UCase(pStrike))
    Select Case True
       Case sPutCall & Left(sStrike, 3) = "CITM"
            smfGetMSNOptionQuote = "CITM"
            Exit Function
       Case sPutCall & Left(sStrike, 3) = "COTM"
            smfGetMSNOptionQuote = "COTM"
            Exit Function
       Case sPutCall & Left(sStrike, 3) = "POTM"
            smfGetMSNOptionQuote = "POTM"
            Exit Function
       Case sPutCall & Left(sStrike, 3) = "PITM"
            smfGetMSNOptionQuote = "PITM"
            Exit Function
       Case VarType(pExpiry) = vbDouble Or IsDate(pExpiry)
            iRows = 0
            iExpiry = pExpiry
            sYr = Right(Format(pExpiry, "yyyy"), 1)
            If sPutCall = "C" Then
               sMon = Mid("ABCDEFGHIJKL", Month(pExpiry), 1)
            Else
               sMon = Mid("MNOPQRSTUVWX", Month(pExpiry), 1)
               End If
            sDay = Mid("123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", Day(pExpiry), 1)
            sFind1 = "." & Trim(UCase(pTicker)) & sYr & sMon & sDay & "C" & Format(1000 * pStrike, "000000")
            sFind2 = ""
       Case Else
            smfGetMSNOptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       End Select
    
    '------------------> Do primary search
    Select Case sItem
       Case "X": smfGetMSNOptionQuote = iExpiry: Exit Function
       End Select
    sURL = "http://investing.money.msn.com/investments/equity-options/?symbol=" & Trim(UCase(pTicker)) & _
           "&optionsdate=" & Format(pExpiry, "mm/dd/yyyy")
    If sItem = "U" Then
       smfGetMSNOptionQuote = "Error"
       On Error Resume Next
       smfGetMSNOptionQuote = 0 + smfGetTagContent(sURL, "span", -2, "/images/trend")
       On Error GoTo ErrorExit
    Else
       smfGetMSNOptionQuote = RCHGetTableCell(sURL, iCells, sFind1, sFind2, , , iRows, "</table")
       End If
    
    If sLabel = "" Then Exit Function  ' Primary search item already retrieved
    
    '------------------> Do extended search
    sURL = "http://moneycentral.msn.com/detail/market_quote?symbol=" & smfGetMSNOptionQuote
    Select Case sItem
        Case "T"
             smfGetMSNOptionQuote = Mid(smfGetTagContent(sURL, "p", -1, ">Last Trade"), 12, 99)
             smfGetMSNOptionQuote = Trim(Left(smfGetMSNOptionQuote, InStr(smfGetMSNOptionQuote, "<") - 1))
        Case Else
             smfGetMSNOptionQuote = RCHGetTableCell(sURL, 1, sLabel)
        End Select
    
    Exit Function

ErrorExit:
    smfGetMSNOptionQuote = "Error"

    End Function
Public Function smfGetNASDAQOptionQuote(ByVal pTicker As Variant, _
                                        ByVal pPutCall As Variant, _
                                        ByVal pExpiry As Variant, _
                                        ByVal pStrike As Variant, _
                                        ByVal pItem As Variant, _
                               Optional ByVal pHistory As Variant = "")
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quotes from NASDAQ
    '-----------------------------------------------------------------------------------------------------------*
    ' 2015.08.13 -- Created function
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for SPY:
    '
    '   =smfGetNASDAQOptionQuote("SPY","C",DATE(2012,6,1),110,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    
    '------------------> Handle special ticker symbol
    Dim sTicker As String, sMini As String
    sTicker = Trim(LCase(pTicker))
    Select Case Right(sTicker, 1)
       Case "0" To "9"
            sMini = Right(sTicker, 1)
            sTicker = Left(sTicker, Len(sTicker) - 1)
       Case Else
            sMini = ""
       End Select
    
    '------------------> Verify the pPutCall parameter
    Dim sPutCall As String
    sPutCall = Left(Trim(LCase(pPutCall)), 1)
    Select Case sPutCall
       Case "p": sPutCall = "Put"
       Case "c": sPutCall = "Call"
       Case Else
            smfGetNASDAQOptionQuote = "Invalid Put/Call indicator (must start with p or c): " & pPutCall
            Exit Function
       End Select
    
    '------------------> Verify pExpiry and pStrike
    Dim sStrike As String, sURL As String
    sStrike = Trim(UCase(pStrike))
    Select Case True
       Case Not (VarType(pExpiry) = vbDouble Or IsDate(pExpiry))
            smfGetNASDAQOptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       Case IsNumeric(sStrike)
            sURL = "http://www.nasdaq.com/symbol/" & sTicker & "/option-chain/" & _
            Format(pExpiry, "yymmdd") & Left(sPutCall, 1) & Format(1000 * sStrike, "00000000") & _
            "-" & sTicker & sMini & "-" & sPutCall
       Case Else
            smfGetNASDAQOptionQuote = "Bad strike price: " & pStrike
            Exit Function
       End Select
    
    '------------------> Verify the pItem parameter
    Dim sSearch As String, sTag As String
    pItem = Trim(UCase(pItem))
    Select Case pItem
       Case "A": sTag = "b": sSearch = ">Ask"
       Case "B": sTag = "b": sSearch = ">Bid"
       Case "C"  ' Change in $
       Case "%"  ' Change in %
       Case "E": sTag = "b": sSearch = ">Bid Size"
       Case "F": sTag = "b": sSearch = ">Ask Size"
       Case "G": sTag = "b": sSearch = ">Day Low"
       Case "H": sTag = "b": sSearch = ">Day High"
       Case "I": sTag = "b": sSearch = ">Open Interest"
       Case "J": sTag = "b": sSearch = ">Contract Low"
       Case "K": sTag = "b": sSearch = ">Contract High"
       Case "L": sTag = "b": sSearch = ">Last Sale"
       Case "O": sTag = "b": sSearch = ">Open"
       Case "P": sTag = "b": sSearch = ">Prev Close"
       Case "S"  ' Strike price
       Case "T"  ' Time of last trade
       Case "U"  ' Underlying price
       Case "V": sTag = "b": sSearch = ">Volume"
       Case "X"  ' Expiration date
       Case "Z"  ' Ticker symbol
       Case "1": sTag = "span": sSearch = ">Vega"
       Case "2": sTag = "span": sSearch = ">Theta"
       Case "3": sTag = "span": sSearch = ">Rho"
       Case "4": sTag = "span": sSearch = ">Gamma"
       Case "5": sTag = "span": sSearch = ">Delta"
       Case "6": sTag = "span": sSearch = ">ImpVol"
       Case "7": sTag = "b": sSearch = ">Net"
       Case "8": sTag = "b": sSearch = ">Tick"
       Case Else
            smfGetNASDAQOptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
    
    '------------------> Find data item
    Select Case pItem
       Case "C": smfGetNASDAQOptionQuote = smfConvertData(smfGetTagContent(sURL, "b", 1, ">Prev Close")) - smfConvertData(smfGetTagContent(sURL, "b", 1, ">Last Sale"))
       Case "%": smfGetNASDAQOptionQuote = smfConvertData(smfGetTagContent(sURL, "b", 1, ">Prev Close")) / smfConvertData(smfGetTagContent(sURL, "b", 1, ">Last Sale")) - 1
       Case "S": smfGetNASDAQOptionQuote = pStrike
       Case "T": smfGetNASDAQOptionQuote = smfGetTagContent(sURL, "span", -1, "markettime""")
       Case "U": smfGetNASDAQOptionQuote = RCHGetYahooQuotes(sTicker, "l1")(1, 1)
       Case "X": smfGetNASDAQOptionQuote = pExpiry
       Case "Z": smfGetNASDAQOptionQuote = smfStrExtr(sURL, "option-chain/", "~")
       Case "1" To "6"
            If sPutCall = "Call" Then
               smfGetNASDAQOptionQuote = smfConvertData(smfGetTagContent(sURL, sTag, 1, sSearch))
            Else
               smfGetNASDAQOptionQuote = smfConvertData(smfGetTagContent(sURL, sTag, 1, sSearch, sSearch))
               End If
       Case Else
            smfGetNASDAQOptionQuote = smfConvertData(smfGetTagContent(sURL, sTag, 1, sSearch))
       End Select
    Exit Function

ErrorExit:
    smfGetNASDAQOptionQuote = "Error"

    End Function
Public Function smfGetMWOptionQuote(ByVal pTicker As Variant, _
                                    ByVal pPutCall As Variant, _
                                    ByVal pExpiry As Variant, _
                                    ByVal pStrike As Variant, _
                                    ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quotes from MarketWatch
    '-----------------------------------------------------------------------------------------------------------*
    ' 2010.07.16 -- Created function
    ' 2010.08.20 -- Added ability to get index options
    ' 2010.09.10 -- Add "u" data item for last traded price of the underlying equity
    ' 2014.08.17 -- Allow "Put" or "Call" for pPutCall parameter
    ' 2017.08.06 -- Update for new web page format
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for IBM and MMM:
    '
    '   =smfGetMWOptionQuote("SPY","C",DATE(2012,6,1),110,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    
    '------------------> Edit ticker symbol and determine if fund or stock
    pTicker = UCase(Trim(pTicker))
    Dim sURL As String, sURL1 As String, sURL2 As String, sURL3 As String
    sURL1 = "http://www.marketwatch.com/investing/fund/" & pTicker & "/options?countrycode=US&showAll=True"
    sURL2 = "http://www.marketwatch.com/investing/stock/" & pTicker & "/options?countrycode=US&showAll=True"
    sURL3 = "http://www.marketwatch.com/investing/index/" & pTicker & "/options?countrycode=US&showAll=True"
    If RCHGetTableCell(sURL1, 0, "Current price as of") <> "Error" Then
       sURL = sURL1
    ElseIf RCHGetTableCell(sURL2, 0, "Current price as of") <> "Error" Then
       sURL = sURL2
    Else
       sURL = sURL3
       End If
    
    '------------------> Verify the pPutCall parameter
    Dim sPutCall As String
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P"
       Case "C"
       Case Else
            smfGetMWOptionQuote = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
    
    '------------------> Verify the pItem parameter
    Dim iCells As Integer
    pItem = Trim(UCase(pItem))
    Select Case pItem
       Case "S": If sPutCall = "C" Then iCells = 7 Else iCells = -8  ' Strike price
       Case "L": iCells = 1    ' Last price
       Case "C": iCells = 2    ' Change
       Case "V": iCells = 3    ' Volume
       Case "B": iCells = 4    ' Bid price
       Case "A": iCells = 5    ' Ask price
       Case "I": iCells = 6    ' Open Interest
       Case "U": iCells = 0    ' Last price of underlying equity
       Case "X": iCells = 0    ' Option expiration date
       Case "Z": iCells = 0    ' Option ticker symbol
       Case Else
            smfGetMWOptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
       
    '------------------> Handle special strike price strings
    Dim sFind1 As String, sFind2 As String, iRows As Integer
    pStrike = Trim(UCase(pStrike))
    iRows = 0
    Select Case True
       Case Not (VarType(pExpiry) = vbDouble Or IsDate(pExpiry))
            smfGetMWOptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       Case IsNumeric(pStrike)
            sFind1 = pTicker & Chr(Month(pExpiry) - 12 * (sPutCall = "P") + 64) & Format(pExpiry, "ddyy") & _
                     Left(Right(smfStrExtr(smfGetTagContent(sURL, "td", -1, "/option/"), "/option/", """"), 7), 1) & _
                     Format(IIf(pStrike < 100, 10000, 1000) * pStrike, "000000")
            sFind2 = " "
       Case Mid(pStrike, 2, 2) = "TM"
            sFind1 = Format(pExpiry, "yymmdd") & sPutCall
            sFind2 = "Current Price as of"
       Case Else
            smfGetMWOptionQuote = "Bad strike price: " & pStrike
            Exit Function
       End Select
    Select Case sPutCall & Left(pStrike, 3)
       Case "CITM"
            iRows = -CInt(Mid(pStrike, 4, 2)) - 1
            iCells = iCells + 1
       Case "COTM"
            iRows = CInt(Mid(pStrike, 4, 2))
            iCells = iCells + 1
       Case "POTM"
            iRows = -CInt(Mid(pStrike, 4, 2)) - 1
            If iCells = -8 Then iCells = 8 Else iCells = iCells + 9
       Case "PITM"
            iRows = CInt(Mid(pStrike, 4, 2))
            If iCells = -8 Then iCells = 8 Else iCells = iCells + 9
       End Select
    
    '------------------> Find data item
    Dim nTemp As Variant
    Select Case pItem
       Case "U": smfGetMWOptionQuote = RCHGetTableCell(sURL, 2, "Current Price As Of", , , , -1, , , "Error")
       Case "X": smfGetMWOptionQuote = pExpiry
       Case "Z"
          If Len(sFind1) > 7 Then
             smfGetMWOptionQuote = sFind1
          Else
             nTemp = RCHGetTableCell(sURL, 8, sFind1, sFind2, " ", " ", iRows)
             smfGetMWOptionQuote = pTicker & sFind1 & Format(1000 * nTemp, "00000000")
             End If
       Case Else
          smfGetMWOptionQuote = RCHGetTableCell(sURL, iCells, sFind1, sFind2, " ", " ", iRows)
       End Select
    Exit Function

ErrorExit:
    smfGetMWOptionQuote = "Error"

    End Function
Public Function smfGetOXOptionQuote(ByVal pTicker As Variant, _
                                    ByVal pPutCall As Variant, _
                                    ByVal pExpiry As Variant, _
                                    ByVal pStrike As Variant, _
                                    ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quote from Yahoo
    '-----------------------------------------------------------------------------------------------------------*
    ' 2010.06.20 -- Created function
    ' 2010.09.10 -- Add "u" data item for last traded price of the underlying equity
    ' 2013.01.04 -- Remove day 1, 30, and 31 assumptions for option expiration date
    ' 2014.08.17 -- Allow "Put" or "Call" for pPutCall parameter
    ' 2017.01.05 -- Fix "u" data item for last traded price of the underlying equity
    ' 2017.10.10 -- optionsXpress is no longer a valid data source
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for IBM and MMM:
    '
    '   =smfGetOXOptionQuote("SPY","C",DATE(2012,12,22),65,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    Dim sStrike As String, sItem As String, sTicker As String
    Dim sURL As String, sFind1 As String
    Dim sLabel As String, iCells As Integer
    Dim iYear As Integer, iMonth As Integer, iExpiry As Date
    
    smfGetOXOptionQuote = "Obsolete -- OptionsXpress is no longer a valid data source"
    Exit Function
    
    '------------------> Verify the pPutCall parameter
    Dim sPutCall As String
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "C"
       Case Else
            smfGetOXOptionQuote = "Invalid Put/Call indicator (must be a C): " & pPutCall
            Exit Function
       End Select
       
    '------------------> Verify the pItem parameter and set the # of cells to skip
    sItem = Trim(UCase(pItem))
    Select Case sItem
       Case "S": iCells = -12  ' Strike price
       Case "L": iCells = -11  ' Last price
       Case "B": iCells = -10  ' Bid price
       Case "A": iCells = -9   ' Ask price
       Case "Y": iCells = -8   ' Theoretical Value
       Case "I": iCells = -7   ' Open Interest
       Case "5": iCells = -6   ' Delta
       Case "4": iCells = -5   ' Gamma
       Case "3": iCells = -4   ' Rho
       Case "2": iCells = -3   ' Theta
       Case "1": iCells = -2   ' Vega
       Case "U": iCells = 0    ' Last price of underlying equity
       Case "X": iCells = 0    ' Option expiration date
       Case "Z": iCells = 0    ' Option ticker symbol
       Case Else
            smfGetOXOptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
       
    '------------------> Handle special strike price strings
    sStrike = Trim(UCase(pStrike))
    sTicker = Trim(UCase(pTicker))
    Select Case True
       Case VarType(pExpiry) = vbDouble Or IsDate(pExpiry)
            Select Case Day(pExpiry)
               'Case 30, 31 ' Quarterly expiration
               '   iExpiry = smfGetOptionExpiry(Year(pExpiry), Month(pExpiry), "Q")
               '   sFind1 = Left(sTicker & "^^^^^^", 6) & _
               '            Format(iExpiry, "yymmdd") & _
               '            spPutCall & Format(1000 * pStrike, "00000000")
               '   sLabel = Format(iExpiry, "m/d/yyyy") & ";3"
               'Case Is = 1 ' Monthly expiration
               '   iExpiry = smfGetOptionExpiry(Year(pExpiry), Month(pExpiry))
               '   sFind1 = Left(sTicker & "^^^^^^", 6) & _
               '            Format(iExpiry, "yymmdd") & _
               '            sPutCall & Format(1000 * pStrike, "00000000")
               '   sLabel = Format(iExpiry, "m/d/yyyy") & ";1"
               Case Else
                  iExpiry = pExpiry
                  sFind1 = Left(sTicker & "^^^^^^", 6) & _
                           Format(iExpiry, "yymmdd") & _
                           sPutCall & Format(1000 * pStrike, "00000000")
                  sLabel = Format(pExpiry, "m/d/yyyy") & ";1"
               End Select
       Case Else
            smfGetOXOptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       End Select
    
    '------------------> Find data item
    sURL = "https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Range=0&lstMarket=0&ChainType=3&lstMonths=" & _
           sLabel & "&Symbol=" & sTicker
    Select Case sItem
       Case "U": smfGetOXOptionQuote = RCHGetTableCell(sURL, 2, ">Change", ">Change", , , 1, , , "Error")
       Case "X": smfGetOXOptionQuote = iExpiry
       Case "Z": smfGetOXOptionQuote = sFind1
       Case Else: smfGetOXOptionQuote = RCHGetTableCell(sURL, iCells, sFind1)
       End Select
    Exit Function

ErrorExit:
    smfGetOXOptionQuote = "Error"

    End Function

Public Function smfGetOX2OptionQuote(ByVal pTicker As Variant, _
                                     ByVal pPutCall As Variant, _
                                     ByVal pExpiry As Variant, _
                                     ByVal pStrike As Variant, _
                                     ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quotes from OptionsXPress
    '-----------------------------------------------------------------------------------------------------------*
    ' 2010.07.24 -- Created function
    ' 2010.09.10 -- Add "u" data item for last traded price of the underlying equity
    ' 2011.11.30 -- Change URL to one that gets all expiration dates in one Internet access
    ' 2014.08.17 -- Allow "Put" or "Call" for pPutCall parameter
    ' 2017.01.05 -- Fix "u" data item for last traded price of the underlying equity
    ' 2017.10.10 -- optionsXpress is no longer a valid data source
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for SPY:
    '
    '   =smfGetOX2OptionQuote("SPY","C",DATE(2012,6,1),110,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    
    smfGetOX2OptionQuote = "Obsolete -- OptionsXpress is no longer a valid data source"
    Exit Function
        
    '------------------> Verify the pPutCall parameter
    Dim iCall As Integer
    Dim sPutCall As String
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P": iCall = 0
       Case "C": iCall = -9
       Case Else
            smfGetOX2OptionQuote = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
    
    '------------------> Verify pExpiry and pStrike
    Dim sFind1 As String
    pStrike = Trim(UCase(pStrike))
    pTicker = Trim(UCase(pTicker))
    Select Case True
       Case Not (VarType(pExpiry) = vbDouble Or IsDate(pExpiry))
            smfGetOX2OptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       Case IsNumeric(pStrike)
            sFind1 = Left(pTicker & "^^^^^^", 6) & Format(pExpiry, "yymmdd") & sPutCall & Format(1000 * pStrike, "00000000")
       Case Else
            smfGetOX2OptionQuote = "Bad strike price: " & pStrike
            Exit Function
       End Select
    
    '------------------> Verify the pItem parameter
    Dim iCells As Integer
    pItem = Trim(UCase(pItem))
    Select Case pItem
       Case "5": iCells = -3 + iCall   ' Delta
       Case "6": iCells = -4 + iCall   ' Implied Volatility
       Case "A": iCells = -5 + iCall   ' Ask price
       Case "B": iCells = -6 + iCall   ' Bid price
       Case "C": iCells = -7 + iCall   ' Change
       Case "L": iCells = -8 + iCall   ' Last price
       Case "S": iCells = -9           ' Strike price
       Case "U": iCells = 0            ' Last price of underlying equity
       Case "X": iCells = 0            ' Option expiration date
       Case "Z": iCells = 0            ' Option ticker symbol
       Case Else
            smfGetOX2OptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
    
    '------------------> Find data item
    Dim sURL As String
    sURL = "https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Range=All&lstMarket=0&ChainType=14&lstMonths=" & _
           Format(smfGetOptionExpiry(), "mm/dd/yyyy") & ";7&Symbol=" & pTicker
    Select Case pItem
       Case "U": smfGetOX2OptionQuote = RCHGetTableCell(sURL, 2, ">Change", ">Change", , , 1, , , "Error")
       Case "X": smfGetOX2OptionQuote = pExpiry
       Case "Z": smfGetOX2OptionQuote = sFind1
       Case Else: smfGetOX2OptionQuote = RCHGetTableCell(sURL, iCells, sFind1)
       End Select
    Exit Function

ErrorExit:
    smfGetOX2OptionQuote = "Error"

    End Function

Public Function smfGetOX3OptionQuote(ByVal pTicker As Variant, _
                                     ByVal pPutCall As Variant, _
                                     ByVal pExpiry As Variant, _
                                     ByVal pStrike As Variant, _
                                     ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quotes from OptionsXPress
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.11.30 -- Created function
    ' 2011.12.01 -- Fix strike price extraction because of no "a" tags on zero bid prices
    ' 2014.08.17 -- Allow "Put" or "Call" for pPutCall parameter
    ' 2017.01.05 -- Fix "u" data item for last traded price of the underlying equity
    ' 2017.10.10 -- optionsXpress is no longer a valid data source
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for SPY:
    '
    '   =smfGetOX2OptionQuote("SPY","C",DATE(2012,6,1),110,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    
    smfGetOX3OptionQuote = "Obsolete -- OptionsXpress is no longer a valid data source"
    Exit Function
        
    '------------------> Verify the pPutCall parameter
    Dim iCall As Integer
    Dim sPutCall As String
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P": iCall = 0
       Case "C": iCall = -9
       Case Else
            smfGetOX3OptionQuote = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
    
    '------------------> Verify pExpiry and pStrike
    Dim sFind1 As String
    pStrike = Trim(UCase(pStrike))
    pTicker = Trim(UCase(pTicker))
    Select Case True
       Case Not (VarType(pExpiry) = vbDouble Or IsDate(pExpiry))
            smfGetOX3OptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       Case IsNumeric(pStrike)
            sFind1 = Left(pTicker & "^^^^^^", 6) & Format(pExpiry, "yymmdd") & sPutCall & Format(1000 * pStrike, "00000000")
       Case Else
            smfGetOX3OptionQuote = "Bad strike price: " & pStrike
            Exit Function
       End Select
    
    '------------------> Verify the pItem parameter
    Dim iCells As Integer
    pItem = Trim(UCase(pItem))
    Select Case pItem
       Case "I": iCells = -3 + iCall   ' Open Interest
       Case "V": iCells = -4 + iCall   ' Volume
       Case "A": iCells = -5 + iCall   ' Ask price
       Case "B": iCells = -6 + iCall   ' Bid price
       Case "C": iCells = -7 + iCall   ' Change
       Case "L": iCells = -8 + iCall   ' Last price
       Case "S": iCells = -9           ' Strike price
       Case "U": iCells = 0            ' Last price of underlying equity
       Case "X": iCells = 0            ' Option expiration date
       Case "Z": iCells = 0            ' Option ticker symbol
       Case Else
            smfGetOX3OptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
    
    '------------------> Find data item
    Dim sURL As String
    sURL = "https://www.optionsxpress.com/OXNetTools/Chains/index.aspx?Range=All&lstMarket=0&lstMonths=" & _
           Format(smfGetOptionExpiry(), "mm/dd/yyyy") & ";7&Symbol=" & pTicker
    Select Case pItem
       Case "S"
            If iCall = 0 Then iCall = -1 Else iCall = 1
            smfGetOX3OptionQuote = 0 + smfStrExtr(smfGetTagContent(sURL, "div", iCall, sFind1), ">", "<")
       Case "U": smfGetOX3OptionQuote = RCHGetTableCell(sURL, 2, ">Change", ">Change", , , 1, , , "Error")
       Case "X": smfGetOX3OptionQuote = pExpiry
       Case "Z": smfGetOX3OptionQuote = sFind1
       Case Else: smfGetOX3OptionQuote = RCHGetTableCell(sURL, iCells, sFind1)
       End Select
    Exit Function

ErrorExit:
    smfGetOX3OptionQuote = "Error"

    End Function

Public Function smfGetYahooOptionQuote(ByVal pTicker As Variant, _
                                       ByVal pPutCall As Variant, _
                                       ByVal pExpiry As Variant, _
                                       ByVal pStrike As Variant, _
                                       ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quote from Yahoo
    '-----------------------------------------------------------------------------------------------------------*
    ' 2010.04.07 -- Created function
    ' 2010.09.10 -- Add "u" data item for last traded price of the underlying equity
    ' 2012.01.15 -- Update elements c/t/u for Yahoo web page changes (dropped last trade and time)
    ' 2013.01.04 -- Remove day 1, 30, and 31 assumptions for option expiration date
    ' 2013.02.06 -- Add workaround to fix the "^VIX" / "VIX" issue
    ' 2013.06.03 -- Add ability to use mini options by apending "7" to ticker symbol
    ' 2013.06.28 -- Add ability to use adjusted ticker symbols (i.e. rightward numeric)
    ' 2013.12.21 -- Fix sPutCall processing to only use 1st byte of parameter
    ' 2016.07.13 -- Update extractions for updated Yahoo website
    ' 2016.08.04 -- Update quotes page URL because of Yahoo change
    ' 2016.11.26 -- Update extractions because of Yahoo change
    ' 2017.04.26 -- Change protocol from "http://" to "https://
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for IBM and MMM:
    '
    '   =smfGetYahooOptionQuote("SPY","C",DATE(2012,6,1),110,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    Dim sPutCall As String, sStrike As String, sItem As String
    Dim sURL As String, sFind1 As String, sFind2 As String
    Dim sLabel As String, iCells As Integer, iRows As Integer
    Dim iYear As Integer, iMonth As Integer, iDay As Integer
    
    '------------------> Verify the pPutCall parameter
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P"
       Case "C"
       Case Else
            smfGetYahooOptionQuote = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
       
    '------------------> Handle special strike price strings
    Dim sTicker As String, sMini As String
    sTicker = Trim(UCase(pTicker))
    sMini = ""
    Select Case Right(sTicker, 1)
       Case "0" To "9"
            sMini = Right(sTicker, 1)
            sTicker = Left(sTicker, Len(sTicker) - 1)
       End Select
    sStrike = Trim(UCase(pStrike))
    Select Case True
       Case VarType(pExpiry) = vbDouble Or IsDate(pExpiry)
            sFind1 = sTicker & sMini & Format(pExpiry, "yymmdd") & sPutCall & Format(1000 * pStrike, "00000000")
       Case Else
            smfGetYahooOptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       End Select
       
    '------------------> Verify the pItem parameter and return data item
    If sTicker = "VIX" Then sTicker = "^VIX"
    sURL = "https://finance.yahoo.com/quote/" & sFind1
    
    sItem = Trim(UCase(pItem))
    Select Case sItem
       Case "A": smfGetYahooOptionQuote = smfGetTagContent(sURL, "td", 1, ">Ask", , , , 1)
       Case "B": smfGetYahooOptionQuote = smfGetTagContent(sURL, "td", 1, ">Bid", , , , 1)
       Case "C": smfGetYahooOptionQuote = smfConvertData(smfStrExtr(smfGetTagContent(sURL, "span", -1, "quote-market-notice"), "~", " "))
       Case "O": smfGetYahooOptionQuote = smfGetTagContent(sURL, "td", 1, ">Open", , , , 1)
       Case "P": smfGetYahooOptionQuote = smfGetTagContent(sURL, "td", 1, ">Previous Close", , , , 1)
       Case "V": smfGetYahooOptionQuote = smfGetTagContent(sURL, "td", 1, ">Volume", , , , 1)
       Case "I": smfGetYahooOptionQuote = smfGetTagContent(sURL, "td", 1, ">Open Interest", , , , 1)
       Case "L": smfGetYahooOptionQuote = smfGetTagContent(sURL, "span", -2, "quote-market-notice", , , , 1)
       Case "S": smfGetYahooOptionQuote = pStrike
       Case "X": smfGetYahooOptionQuote = smfGetTagContent(sURL, "td", 1, ">Expire Date", , , , 1)
       Case "Z": smfGetYahooOptionQuote = sFind1
       Case "%": smfGetYahooOptionQuote = smfConvertData(smfStrExtr(smfGetTagContent(sURL, "span", -1, "quote-market-notice"), "(", ")"))
       Case Else: smfGetYahooOptionQuote = "Unrecognized item ID: " & pItem
       End Select
    Exit Function

ErrorExit:
    smfGetYahooOptionQuote = "Error"

    End Function
    
Public Function smfGetYahooOptionQuote2(ByVal pTicker As Variant, _
                                       ByVal pPutCall As Variant, _
                                       ByVal pExpiry As Variant, _
                                       ByVal pStrike As Variant, _
                                       ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quote from Yahoo
    '-----------------------------------------------------------------------------------------------------------*
    ' 2010.04.07 -- Created function
    ' 2010.09.10 -- Add "u" data item for last traded price of the underlying equity
    ' 2012.01.15 -- Update elements c/t/u for Yahoo web page changes (dropped last trade and time)
    ' 2013.01.04 -- Remove day 1, 30, and 31 assumptions for option expiration date
    ' 2013.02.06 -- Add workaround to fix the "^VIX" / "VIX" issue
    ' 2013.06.03 -- Add ability to use mini options by apending "7" to ticker symbol
    ' 2013.06.28 -- Add ability to use adjusted ticker symbols (i.e. rightward numeric)
    ' 2013.12.21 -- Fix sPutCall processing to only use 1st byte of parameter
    ' 2014.10.21 -- Update for changes in Yahoo's option quotes pages
    ' 2014.10.21 -- Removed "xITM" strike price choices
    ' 2014.10.23 -- Handle situations where ticker symbols contain a hypen (e.g. BRK-B)
    ' 2014.10.23 -- Fix offset for picking up the strike price
    ' 2014.10.24 -- Handle situations where ticker symbols contain an uptick (e.g. ^VIX)
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for IBM and MMM:
    '
    '   =smfGetYahooOptionQuote2("SPY","C",DATE(2012,6,1),110,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    Dim sPutCall As String, sStrike As String, sItem As String
    Dim sURL As String, sFind1 As String, sFind2 As String
    Dim sLabel As String, iCells As Integer, iRows As Integer
    Dim iYear As Integer, iMonth As Integer, iDay As Integer
    
    '------------------> Verify the pPutCall parameter
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P"
       Case "C"
       Case Else
            smfGetYahooOptionQuote2 = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
       
    '------------------> Verify the pItem parameter and set the # of cells to skip
    sItem = Trim(UCase(pItem))
    Select Case sItem
       Case "S": iCells = -10: sLabel = ""               ' Strike price
       Case "Z": iCells = 0: sLabel = ""                ' Yahoo ticker symbol
       Case "L": iCells = 1: sLabel = ""                ' Last price
       Case "B": iCells = 2: sLabel = ""                ' Bid price
       Case "A": iCells = 3: sLabel = ""                ' Ask price
       Case "C": iCells = 4: sLabel = ""                ' $ Change
       Case "%": iCells = 5: sLabel = ""                ' % Change
       Case "V": iCells = 6: sLabel = ""                ' Volume
       Case "I": iCells = 7: sLabel = ""                ' Open Interest
       Case "6": iCells = 8: sLabel = ""                ' Implied volatility
       Case "G": iCells = 0: sLabel = "Day's Range:"    ' For computing daily low
       Case "H": iCells = 0: sLabel = "Day's Range:"    ' For computing daily high
       Case "J": iCells = 0: sLabel = "Contract Range:" ' For computing contract low
       Case "K": iCells = 0: sLabel = "Contract Range:" ' For computing contract high
       Case "O": iCells = 0: sLabel = "Open:"           ' Open
       Case "P": iCells = 0: sLabel = "Prev Close:"     ' Previous close
       Case "T": iCells = 0: sLabel = "x"               ' Last trade time
       Case "U": iCells = 0: sLabel = ""                ' Last price of underlying equity
       Case "X": iCells = 0: sLabel = ""                ' For computing expiry date
       Case Else
            smfGetYahooOptionQuote2 = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
       
    '------------------> Handle special strike price strings
    Dim sTicker As String, sMini As String, sMini2 As String
    sTicker = Trim(UCase(pTicker))
    sMini = ""
    Select Case Right(sTicker, 1)
       Case "0" To "9"
            sMini = Right(sTicker, 1)
            sTicker = Left(sTicker, Len(sTicker) - 1)
            sMini2 = "&size=mini"
       Case Else
            sMini2 = ""
       End Select
    sStrike = Trim(UCase(pStrike))
    Select Case True
'       Case sPutCall & Left(sStrike, 3) = "CITM"
'            sFind1 = ">Call Options"
'            sFind2 = "yfnc_tabledata1"
'            iRows = -CInt(Mid(sStrike, 4, 2)) - 1
'            iCells = Application.WorksheetFunction.Max(1, iCells + 2)
'       Case sPutCall & Left(sStrike, 3) = "COTM"
'            sFind1 = ">Call Options"
'            sFind2 = "yfnc_tabledata1"
'            iRows = CInt(Mid(sStrike, 4, 2)) - 1
'            If iRows = 0 Then iRows = -1
'            iCells = Application.WorksheetFunction.Max(1, iCells + 2)
'       Case sPutCall & Left(sStrike, 3) = "POTM"
'            sFind1 = ">Put Options"
'            sFind2 = "yfnc_h"
'            iRows = -CInt(Mid(sStrike, 4, 2)) - 1
'            iCells = Application.WorksheetFunction.Max(1, iCells + 2)
'       Case sPutCall & Left(sStrike, 3) = "PITM"
'            sFind1 = ">Put Options"
'            sFind2 = "yfnc_h"
'            iRows = CInt(Mid(sStrike, 4, 2)) - 1
'            If iRows = 0 Then iRows = -1
'            iCells = Application.WorksheetFunction.Max(1, iCells + 2)
       Case VarType(pExpiry) = vbDouble Or IsDate(pExpiry)
            iRows = 0
            Select Case Day(pExpiry)
               'Case 30, 31 ' Quarterly expiration
               '   sFind1 = Format(smfGetOptionExpiry(Year(pExpiry), Month(pExpiry), "Q"), "yymmdd") & _
               '            sPutCall & Format(1000 * pStrike, "00000000")
               'Case Is = 1 ' Monthly expiration
               '   sFind1 = Format(smfGetOptionExpiry(Year(pExpiry), Month(pExpiry)), "yymmdd") & _
               '            sPutCall & Format(1000 * pStrike, "00000000")
               Case Else
                  sFind1 = Replace(Replace(sTicker, "-", ""), "^", "") & sMini & Format(pExpiry, "yymmdd") & sPutCall & Format(1000 * pStrike, "00000000")
               End Select
            sFind2 = ""
       Case Else
            smfGetYahooOptionQuote2 = "Bad expiration date: " & pExpiry
            Exit Function
       End Select
    
    '------------------> Do primary search
    Dim nLast As Variant, dDate As Double
    If sTicker = "VIX" Then sTicker = "^VIX"
    dDate = 86400 * (DateSerial(Year(pExpiry), Month(pExpiry), Day(pExpiry)) - DateSerial(1970, 1, 1))
    sURL = "http://finance.yahoo.com/q/op?s=" & sTicker & sMini2 & "&date=" & dDate

    Select Case sItem
        Case "U"
             smfGetYahooOptionQuote2 = "Error"
             On Error Resume Next
             smfGetYahooOptionQuote2 = smfConvertData(smfGetTagContent(sURL, "span", -1, "yfs_l84_"))
             On Error GoTo ErrorExit
        Case Else
             smfGetYahooOptionQuote2 = RCHGetTableCell(sURL, iCells, sFind1, sFind2, , , iRows, "</table")
        End Select
    
    If sItem = "X" Then
       iYear = Mid(smfGetYahooOptionQuote2, Len(smfGetYahooOptionQuote2) - 14, 2)
       iMonth = Mid(smfGetYahooOptionQuote2, Len(smfGetYahooOptionQuote2) - 12, 2)
       iDay = Mid(smfGetYahooOptionQuote2, Len(smfGetYahooOptionQuote2) - 10, 2)
       smfGetYahooOptionQuote2 = DateSerial(iYear, iMonth, iDay)
       End If
    If sLabel = "" Then Exit Function  ' Primary search item already retrieved
    
    '------------------> Do extended search
    sURL = "http://finance.yahoo.com/q?s=" & smfGetYahooOptionQuote2
    smfGetYahooOptionQuote2 = RCHGetTableCell(sURL, 1, sLabel)
    
    '------------------> Special processing items
    Select Case sItem
        Case "G"
             smfGetYahooOptionQuote2 = smfConvertData(Left(smfGetYahooOptionQuote2, InStr(smfGetYahooOptionQuote2, "-") - 1))
        Case "H"
             smfGetYahooOptionQuote2 = smfConvertData(Mid(smfGetYahooOptionQuote2, InStr(smfGetYahooOptionQuote2, "-") + 1, 99))
        Case "J"
             smfGetYahooOptionQuote2 = smfConvertData(Left(smfGetYahooOptionQuote2, InStr(smfGetYahooOptionQuote2, "-") - 1))
        Case "K"
             smfGetYahooOptionQuote2 = smfConvertData(Mid(smfGetYahooOptionQuote2, InStr(smfGetYahooOptionQuote2, "-") + 1, 99))
        Case "T"
             smfGetYahooOptionQuote2 = "Error"
             On Error Resume Next
             smfGetYahooOptionQuote2 = smfConvertData(smfGetTagContent(sURL, "span", 1, "yfs_t10_"))
             On Error GoTo ErrorExit
        End Select
    
    Exit Function

ErrorExit:
    smfGetYahooOptionQuote2 = "Error"

    End Function


Public Function smfGetYahooOptionQuote3(ByVal pTicker As Variant, _
                                       ByVal pPutCall As Variant, _
                                       ByVal pExpiry As Variant, _
                                       ByVal pStrike As Variant, _
                                       ByVal pItem As Variant)
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get option quote from Yahoo
    '-----------------------------------------------------------------------------------------------------------*
    ' 2010.04.07 -- Created function
    ' 2010.09.10 -- Add "u" data item for last traded price of the underlying equity
    ' 2012.01.15 -- Update elements c/t/u for Yahoo web page changes (dropped last trade and time)
    ' 2013.01.04 -- Remove day 1, 30, and 31 assumptions for option expiration date
    ' 2013.02.06 -- Add workaround to fix the "^VIX" / "VIX" issue
    ' 2013.06.03 -- Add ability to use mini options by apending "7" to ticker symbol
    ' 2013.06.28 -- Add ability to use adjusted ticker symbols (i.e. rightward numeric)
    ' 2013.12.21 -- Fix sPutCall processing to only use 1st byte of parameter
    ' 2016.07.13 -- This processing became obsolete
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get current quotes for IBM and MMM:
    '
    '   =smfGetYahooOptionQuote3("SPY","C",DATE(2012,6,1),110,"b")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    Dim sPutCall As String, sStrike As String, sItem As String
    Dim sURL As String, sFind1 As String, sFind2 As String
    Dim sLabel As String, iCells As Integer, iRows As Integer
    Dim iYear As Integer, iMonth As Integer, iDay As Integer
    
    '------------------> Verify the pPutCall parameter
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P"
       Case "C"
       Case Else
            smfGetYahooOptionQuote3 = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
       
    '------------------> Verify the pItem parameter and set the # of cells to skip
    sItem = Trim(UCase(pItem))
    Select Case sItem
       Case "S": iCells = -8: sLabel = ""               ' Strike price
       Case "Z": iCells = 0: sLabel = ""                ' Yahoo ticker symbol
       Case "L": iCells = 1: sLabel = ""                ' Last price
       Case "B": iCells = 3: sLabel = ""                ' Bid price
       Case "A": iCells = 4: sLabel = ""                ' Ask price
       Case "V": iCells = 5: sLabel = ""                ' Volume
       Case "I": iCells = 6: sLabel = ""                ' Open Interest
       Case "G": iCells = 0: sLabel = "Day's Range:"    ' For computing daily low
       Case "H": iCells = 0: sLabel = "Day's Range:"    ' For computing daily high
       Case "J": iCells = 0: sLabel = "Contract Range:" ' For computing contract low
       Case "K": iCells = 0: sLabel = "Contract Range:" ' For computing contract high
       Case "C": iCells = 0: sLabel = "Prev Close:"     ' Previous close to compute change
       Case "O": iCells = 0: sLabel = "Open:"           ' Open
       Case "P": iCells = 0: sLabel = "Prev Close:"     ' Previous close
       Case "T": iCells = 0: sLabel = "x"               ' Last trade time
       Case "U": iCells = 0: sLabel = ""                ' Last price of underlying equity
       Case "X": iCells = 0: sLabel = ""                ' For computing expiry date
       Case Else
            smfGetYahooOptionQuote3 = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
       
    '------------------> Handle special strike price strings
    Dim sTicker As String, sMini As String
    sTicker = Trim(UCase(pTicker))
    sMini = ""
    Select Case Right(sTicker, 1)
       Case "0" To "9"
            sMini = Right(sTicker, 1)
            sTicker = Left(sTicker, Len(sTicker) - 1)
       End Select
    sStrike = Trim(UCase(pStrike))
    Select Case True
       Case sPutCall & Left(sStrike, 3) = "CITM"
            sFind1 = ">Call Options"
            sFind2 = "yfnc_tabledata1"
            iRows = -CInt(Mid(sStrike, 4, 2)) - 1
            iCells = Application.WorksheetFunction.Max(1, iCells + 2)
       Case sPutCall & Left(sStrike, 3) = "COTM"
            sFind1 = ">Call Options"
            sFind2 = "yfnc_tabledata1"
            iRows = CInt(Mid(sStrike, 4, 2)) - 1
            If iRows = 0 Then iRows = -1
            iCells = Application.WorksheetFunction.Max(1, iCells + 2)
       Case sPutCall & Left(sStrike, 3) = "POTM"
            sFind1 = ">Put Options"
            sFind2 = "yfnc_h"
            iRows = -CInt(Mid(sStrike, 4, 2)) - 1
            iCells = Application.WorksheetFunction.Max(1, iCells + 2)
       Case sPutCall & Left(sStrike, 3) = "PITM"
            sFind1 = ">Put Options"
            sFind2 = "yfnc_h"
            iRows = CInt(Mid(sStrike, 4, 2)) - 1
            If iRows = 0 Then iRows = -1
            iCells = Application.WorksheetFunction.Max(1, iCells + 2)
       Case VarType(pExpiry) = vbDouble Or IsDate(pExpiry)
            iRows = 0
            Select Case Day(pExpiry)
               'Case 30, 31 ' Quarterly expiration
               '   sFind1 = Format(smfGetOptionExpiry(Year(pExpiry), Month(pExpiry), "Q"), "yymmdd") & _
               '            sPutCall & Format(1000 * pStrike, "00000000")
               'Case Is = 1 ' Monthly expiration
               '   sFind1 = Format(smfGetOptionExpiry(Year(pExpiry), Month(pExpiry)), "yymmdd") & _
               '            sPutCall & Format(1000 * pStrike, "00000000")
               Case Else
                  sFind1 = sTicker & sMini & Format(pExpiry, "yymmdd") & sPutCall & Format(1000 * pStrike, "00000000")
               End Select
            sFind2 = ""
       Case Else
            smfGetYahooOptionQuote3 = "Bad expiration date: " & pExpiry
            Exit Function
       End Select
    
    '------------------> Do primary search
    Dim nLast As Variant
    If sTicker = "VIX" Then sTicker = "^VIX"
    sURL = "http://finance.yahoo.com/q/op?s=" & sTicker & "&m=" & Format(pExpiry, "yyyy-mm")

    Select Case sItem
        Case "C"
             nLast = RCHGetTableCell(sURL, 1, sFind1, sFind2, , , iRows, "</table")
             smfGetYahooOptionQuote3 = RCHGetTableCell(sURL, iCells, sFind1, sFind2, , , iRows, "</table")
        Case "U"
             smfGetYahooOptionQuote3 = "Error"
             On Error Resume Next
             smfGetYahooOptionQuote3 = smfConvertData(smfGetTagContent(sURL, "span", -1, "yfs_l84_"))
             On Error GoTo ErrorExit
        Case Else
             smfGetYahooOptionQuote3 = RCHGetTableCell(sURL, iCells, sFind1, sFind2, , , iRows, "</table")
        End Select
    
    If sItem = "X" Then
       iYear = Mid(smfGetYahooOptionQuote3, Len(smfGetYahooOptionQuote3) - 14, 2)
       iMonth = Mid(smfGetYahooOptionQuote3, Len(smfGetYahooOptionQuote3) - 12, 2)
       iDay = Mid(smfGetYahooOptionQuote3, Len(smfGetYahooOptionQuote3) - 10, 2)
       smfGetYahooOptionQuote3 = DateSerial(iYear, iMonth, iDay)
       End If
    If sLabel = "" Then Exit Function  ' Primary search item already retrieved
    
    '------------------> Do extended search
    sURL = "http://finance.yahoo.com/q?s=" & smfGetYahooOptionQuote3
    smfGetYahooOptionQuote3 = RCHGetTableCell(sURL, 1, sLabel)
    
    '------------------> Special processing items
    Select Case sItem
        Case "C"
             smfGetYahooOptionQuote3 = nLast - smfGetYahooOptionQuote3
        Case "G"
             smfGetYahooOptionQuote3 = smfConvertData(Left(smfGetYahooOptionQuote3, InStr(smfGetYahooOptionQuote3, "-") - 1))
        Case "H"
             smfGetYahooOptionQuote3 = smfConvertData(Mid(smfGetYahooOptionQuote3, InStr(smfGetYahooOptionQuote3, "-") + 1, 99))
        Case "J"
             smfGetYahooOptionQuote3 = smfConvertData(Left(smfGetYahooOptionQuote3, InStr(smfGetYahooOptionQuote3, "-") - 1))
        Case "K"
             smfGetYahooOptionQuote3 = smfConvertData(Mid(smfGetYahooOptionQuote3, InStr(smfGetYahooOptionQuote3, "-") + 1, 99))
        Case "T"
             smfGetYahooOptionQuote3 = "Error"
             On Error Resume Next
             smfGetYahooOptionQuote3 = smfConvertData(smfGetTagContent(sURL, "span", 1, "yfs_t10_"))
             On Error GoTo ErrorExit
        End Select
    
    Exit Function

ErrorExit:
    smfGetYahooOptionQuote3 = "Error"

    End Function

