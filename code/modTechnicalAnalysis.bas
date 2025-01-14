Attribute VB_Name = "modTechnicalAnalysis"
'@Lang VBA
Const kD = 1
Const kO = 2
Const kH = 3
Const kL = 4
Const kC = 5
Const kV = 6

Function SMFTech(pDataRange As Variant, _
                 pIndicator As String, _
        Optional pParm1 As Variant = "", _
        Optional pParm2 As Variant = "", _
        Optional pParm3 As Variant = "", _
        Optional pParm4 As Variant = "", _
        Optional pParm5 As Variant = "", _
        Optional pParm6 As Variant = "")
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to create technical analysis indicators
    '-----------------------------------------------------------------------------------------------------------*
    ' 2006.10.07 -- Created by Randy Harmelink (rharmelink@gmail.com)
    ' 2006.10.07 -- Add Simple Moving Average (SMA) / Average True Range (ATR) / Relative Strength Index (RSI)
    ' 2006.10.08 -- Add Commodity Channel Index (CCI) / Stochastic (Sto)
    ' 2006.10.08 -- Add Moving Average Convergence Divergence (MACD)
    ' 2006.10.13 -- Add Accumulation/Distribuion Line (ADL) / On Balance Volume (OBV)
    ' 2007.01.07 -- Fix ADL to account for days when High and Low price are the same
    '-----------------------------------------------------------------------------------------------> Version 1.3
    ' 2007.09.19 -- Change pRange to be of type Variant instead of Range so an array can be passed to function
    ' 2007.09.20 -- Add Rate of Change (ROC) indicator
    '-----------------------------------------------------------------------------------------------> Version 2.0h
    ' Note: The pDataRange parameter is assumed to be a range of historical quotes data (e.g. from Yahoo!) where
    ' the columns are Date/Open/High/Low/Close/Volume, the first row contain column names, and the rows are in
    ' ascending date sequence.
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
  
    '------------------> Initialize return array
    kData = pDataRange
    kDim1 = UBound(kData, 1)
    kDim2 = UBound(kData, 2)
    ReDim vData(1 To kDim1, 1 To 7) As Variant
    For i1 = 1 To kDim1: vData(i1, 1) = "": Next i1

    '------------------> Determine which indicator to return
    Select Case UCase(pIndicator)
       Case "ADL": GoTo TA_ADL
       Case "ATR": GoTo TA_ATR
       Case "CCI": GoTo TA_CCI
       Case "EMA": GoTo TA_EMA
       Case "MACD": GoTo TA_MACD
       Case "OBV": GoTo TA_OBV
       Case "ROC": GoTo TA_ROC
       Case "RSI": GoTo TA_RSI
       Case "SMA": GoTo TA_SMA
       Case "STO": GoTo TA_STO
       Case Else: GoTo ErrorExit
       End Select

'------------------> Accumulation/Distribution Line
TA_ADL:

    vData(1, 1) = "ADL"
    vData(1, 2) = "CLV"
    n1 = 0
    n2 = 0
    For i1 = 2 To kDim1
        If kData(i1, kH) > kData(i1, kL) Then
           n2 = (kData(i1, kC) - kData(i1, kL) + kData(i1, kC) - kData(i1, kH)) / (kData(i1, kH) - kData(i1, kL))
           n1 = n1 + n2 * kData(i1, kV)
           End If
        vData(i1, 1) = n1
        vData(i1, 2) = n2
        Next i1

    GoTo ExitFunction
   
'------------------> Average True Range
TA_ATR:
   
    If pParm1 = "" Then pParm1 = 20
    vData(1, 1) = "ATR" & Chr(10) & pParm1
    vData(1, 2) = "H - L"
    vData(1, 3) = "Abs(H-C1)"
    vData(1, 4) = "Abs(L-C1)"
    vData(1, 5) = "Daily TR"
    nSum = 0
    nTol = IIf(pParm1 = 1, 2, pParm1)
    For i1 = 2 To kDim1
        n2 = kData(i1, kH) - kData(i1, kL)
        If i1 = 2 Then
           n3 = ""
           n4 = ""
           n5 = n2
        Else
           n3 = Abs(kData(i1, kH) - kData(i1 - 1, kC))
           n4 = Abs(kData(i1, kL) - kData(i1 - 1, kC))
           n5 = IIf(n4 > n3, n4, n3)
           n5 = IIf(n5 > n2, n5, n2)
           End If
        If i1 > nTol Then
           n1 = (n5 + (pParm1 - 1) * vData(i1 - 1, 1)) / pParm1
        Else
           nSum = nSum + n5
           n1 = nSum / (i1 - 1)
           End If
        vData(i1, 1) = n1
        vData(i1, 2) = n2
        vData(i1, 3) = n3
        vData(i1, 4) = n4
        vData(i1, 5) = n5
        Next i1

    GoTo ExitFunction

'------------------> Commodity Channel Index
TA_CCI:
   
    If pParm1 = "" Then pParm1 = 20
    vData(1, 1) = "CCI" & Chr(10) & pParm1
    vData(1, 2) = "TP"
    vData(1, 3) = "TPMA"
    vData(1, 4) = "MD"
    For i1 = 2 To pParm1: For i2 = 2 To 4: vData(i1, i2) = "": Next i2: Next i1

    nSum = 0
    For i1 = 2 To kDim1
        n2 = (kData(i1, kC) + kData(i1, kH) + kData(i1, kL)) / 3
        If i1 > pParm1 + 1 Then
           nSum = nSum + n2 - vData(i1 - pParm1, 2)
           n3 = nSum / pParm1
           nSum2 = Abs(n3 - n2)
           For i2 = i1 - pParm1 + 1 To i1 - 1
               nSum2 = nSum2 + Abs(n3 - vData(i2, 2))
               Next i2
           n4 = nSum2 / pParm1
           n1 = (n2 - n3) / (0.015 * n4)
        Else
           nSum = nSum + n2
           n3 = nSum / (i1 - 1)
           n1 = ""
           n4 = ""
           End If
        vData(i1, 1) = n1
        vData(i1, 2) = n2
        vData(i1, 3) = n3
        vData(i1, 4) = n4
        Next i1

    GoTo ExitFunction

'------------------> Exponential Moving Average
TA_EMA:

    If pParm1 = "" Then pParm1 = 50
    nMult = 2 / (pParm1 + 1)
    vData(1, 1) = "EMA" & Chr(10) & pParm1

    nSum = 0
    For i1 = 2 To kDim1
        If i1 > pParm1 + 1 Then
           n1 = nMult * (kData(i1, kC) - n1) + n1
           vData(i1, 1) = n1
        Else
           nSum = nSum + kData(i1, kC)
           n1 = nSum / (i1 - 1)
           vData(i1, 1) = IIf(i1 < pParm1 + 1, "", n1)
           End If
        Next i1

    GoTo ExitFunction

'------------------> Moving Average Convergence Divergence
TA_MACD:
   
    If pParm1 = "" Then pParm1 = 12
    If pParm2 = "" Then pParm2 = 26
    If pParm3 = "" Then pParm3 = 9
    vData(1, 1) = "MACD" & Chr(10) & pParm1 & "-" & pParm2 & "-" & pParm3
    vData(1, 2) = "SMA" & Chr(10) & pParm1
    vData(1, 3) = "SMA" & Chr(10) & pParm2

    nSum1 = 0
    nSum2 = 0
    nSum3 = 0
    For i1 = 2 To kDim1
        If i1 > pParm1 + 1 Then
           nSum2 = nSum2 + kData(i1, kC) - kData(i1 - pParm1, kC)
           n2 = nSum2 / pParm1
        Else
           nSum2 = nSum2 + kData(i1, kC)
           n2 = nSum2 / (i1 - 1)
           End If
        If i1 > pParm2 + 1 Then
           nSum3 = nSum3 + kData(i1, kC) - kData(i1 - pParm2, kC)
           n3 = nSum3 / pParm2
        Else
           nSum3 = nSum3 + kData(i1, kC)
           n3 = nSum3 / (i1 - 1)
           End If
        If i1 > pParm3 + 1 Then
           nSum1 = nSum1 + (n2 - n3) - (vData(i1 - pParm3, 2) - vData(i1 - pParm3, 3))
           n1 = nSum1 / pParm3
        Else
           nSum1 = nSum1 + (n2 - n3)
           n1 = nSum1 / (i1 - 1)
           End If
        vData(i1, 1) = n1
        vData(i1, 2) = n2
        vData(i1, 3) = n3
        Next i1
   
    GoTo ExitFunction

'------------------> On Balance Volume
TA_OBV:

    vData(1, 1) = "OBV"
    n1 = 0
    For i1 = 3 To kDim1
        Select Case (kData(i1, kC) - kData(i1 - 1, kC))
           Case Is > 0: n1 = n1 + kData(i1, kV)
           Case Is < 0: n1 = n1 - kData(i1, kV)
           End Select
        vData(i1, 1) = n1
        Next i1

    GoTo ExitFunction
    
'------------------> Rate of Change
TA_ROC:
    If pParm1 = "" Then pParm1 = 21
    vData(1, 1) = "ROC" & Chr(10) & pParm1
    
    For i1 = 2 To kDim1
        If i1 < pParm1 + 2 Then
           vData(i1, 1) = ""
        Else
           vData(i1, 1) = kData(i1, kC) / kData(i1 - pParm1, kC) - 1
           End If
        Next i1
    
    GoTo ExitFunction
    
'------------------> Relative Strength Index
TA_RSI:
   
    If pParm1 = "" Then pParm1 = 20
    vData(1, 1) = "RSI" & Chr(10) & pParm1
    vData(1, 2) = "Chg"
    vData(1, 3) = "Adva"
    vData(1, 4) = "Decl"
    vData(1, 5) = "AvgGain"
    vData(1, 6) = "AvgLoss"
    vData(1, 7) = "RS"
    For i1 = 2 To pParm1: For i2 = 2 To 7: vData(i1, i2) = "": Next i2: Next i1

    s3 = 0
    s4 = 0
    For i1 = 3 To kDim1
        n2 = kData(i1, kC) - kData(i1 - 1, kC)
        If i1 > pParm1 + 2 Then
            If n2 > 0 Then
              n3 = n2
              n4 = ""
              n5 = ((pParm1 - 1) * n5 + n2) / pParm1
              n6 = ((pParm1 - 1) * n6) / pParm1
           Else
              n3 = ""
              n4 = -n2
              n5 = ((pParm1 - 1) * n5) / pParm1
              n6 = ((pParm1 - 1) * n6 - n2) / pParm1
              End If
        Else
            If n2 > 0 Then
              n3 = n2
              s3 = s3 + n2
              n4 = ""
           Else
              n3 = ""
              n4 = -n2
              s4 = s4 - n2
              End If
           If i1 = pParm1 + 2 Then
              n5 = s3 / pParm1
              n6 = s4 / pParm1
           Else
              n5 = ""
              n6 = ""
              End If
           End If
        If n5 = "" Then
           n1 = ""
           N7 = ""
        Else
           If n6 = 0 Then
              N7 = 0
              n1 = 100
           Else
              N7 = n5 / n6
              n1 = 100 - (100 / (1 + N7))
              End If
           End If
        vData(i1, 1) = n1
        vData(i1, 2) = n2
        vData(i1, 3) = n3
        vData(i1, 4) = n4
        vData(i1, 5) = n5
        vData(i1, 6) = n6
        vData(i1, 7) = N7
        Next i1
    
    GoTo ExitFunction

'------------------> Simple Moving Average
TA_SMA:
   
    If pParm1 = "" Then pParm1 = 50
    vData(1, 1) = "SMA" & Chr(10) & pParm1
    
    nSum = 0
    For i1 = 2 To kDim1
        If i1 > pParm1 + 1 Then
           nSum = nSum + kData(i1, kC) - kData(i1 - pParm1, kC)
           vData(i1, 1) = nSum / pParm1
        Else
           nSum = nSum + kData(i1, kC)
           vData(i1, 1) = nSum / (i1 - 1)
           End If
        Next i1

    GoTo ExitFunction

'------------------> Stochastics
TA_STO:
   
    If pParm1 = "" Then pParm1 = 14
    If pParm2 = "" Then pParm2 = 5
    If pParm3 = "" Then pParm3 = 1
    vData(1, 1) = "Stoch" & Chr(10) & pParm1 & "-" & pParm2 & "-" & pParm3
    vData(1, 2) = "%K"
    vData(1, 3) = "%D"
    For i1 = 2 To pParm1: For i2 = 2 To 4: vData(i1, i2) = "": Next i2: Next i1

    nSum1 = 0
    nSum3 = 0
    For i1 = 2 To kDim1
        nHi = kData(i1, kH)
        nLo = kData(i1, kL)
        For i2 = IIf(i1 - pParm1 + 1 > 1, i1 - pParm1 + 1, 2) To i1 - 1
            If kData(i2, kH) > nHi Then nHi = kData(i2, kH)
            If kData(i2, kL) < nLo Then nLo = kData(i2, kL)
            Next i2
        n2 = 100 * (kData(i1, kC) - nLo) / (nHi - nLo)
        If i1 > pParm2 + 1 Then
           nSum3 = nSum3 + n2 - vData(i1 - pParm2, 2)
           n3 = nSum3 / pParm2
        Else
           nSum3 = nSum3 + n2
           n3 = nSum3 / (i1 - 1)
           End If
        If i1 > pParm3 + 1 Then
           nSum1 = nSum1 + n3 - vData(i1 - pParm3, 3)
           n1 = nSum1 / pParm3
        Else
           nSum1 = nSum1 + n3
           n1 = nSum1 / (i1 - 1)
           End If
        vData(i1, 1) = n1
        vData(i1, 2) = n2
        vData(i1, 3) = n3
        Next i1

    GoTo ExitFunction

ExitFunction:
ErrorExit:

    SMFTech = vData
    End Function
