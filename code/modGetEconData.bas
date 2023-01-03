Attribute VB_Name = "modGetEconData"
Function smfGetEconData(pID As String, _
                        pDate As Variant, _
               Optional pError As Variant = "Error")
    
    '-----------------------------------------------------------------------------------------------------------*
    ' Function to grab economic data from the St. Louis Federal Reserve web site
    '-----------------------------------------------------------------------------------------------------------*
    ' 2007.11.19 -- Created
    ' 2010.03.13 -- Add extraction of title elements
    ' 2011.04.28 -- Change cDec() to smfConvertData()
    ' 2016.06.20 -- Change to new URL of "https://fred.stlouisfed.org/data/"
    '-----------------------------------------------------------------------------------------------------------*
    ' Example of usage:
    '
    '        =smfGetEconData("CURRNS", DATE(2007,3,24))
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim sURL As String, sDataE As String
    On Error GoTo ErrorExit

    '--------------------------------> See if web page has already been retrieved
'   sURL = "https://research.stlouisfed.org/fred2/data/" & UCase(pID) & ".txt"
    sURL = "https://fred.stlouisfed.org/data/" & UCase(pID) & ".txt"       ' New URL as of 2016-06-20
    For iData = 1 To kPages
        Select Case True
           Case aData(iData, 1) = ""
                aData(iData, 1) = "0:" & pID
                s2 = RCHGetURLData(sURL)
                s2 = Replace(s2, Chr(10), Chr(13))
                aData(iData, 2) = s2
                Exit For
           Case aData(iData, 1) = "0:" & pID: Exit For
           Case iData = kPages
                smfGetEconData = "Error -- Too many web page retrievals"
                Exit Function
           End Select
        Next iData
    sDataE = aData(iData, 2)
    
    '--------------------------------> Check for special items
    smfGetEconData = ""
    Select Case True
       Case UCase(pDate) = "TITLE": smfGetEconData = Trim(smfStrExtr(sDataE, "Title:", Chr(13)))
       Case UCase(pDate) = "SERIES ID": smfGetEconData = Trim(smfStrExtr(sDataE, "Series ID:", Chr(13)))
       Case UCase(pDate) = "SOURCE": smfGetEconData = Trim(smfStrExtr(sDataE, "Source:", Chr(13)))
       Case UCase(pDate) = "RELEASE": smfGetEconData = Trim(smfStrExtr(sDataE, "Release:", Chr(13)))
       Case UCase(pDate) = "SEASONAL ADJUSTMENT": smfGetEconData = Trim(smfStrExtr(sDataE, "Seasonal Adjustment:", Chr(13)))
       Case UCase(pDate) = "FREQUENCY": smfGetEconData = Trim(smfStrExtr(sDataE, "Frequency:", Chr(13)))
       Case UCase(pDate) = "UNITS": smfGetEconData = Trim(smfStrExtr(sDataE, "Units:", Chr(13)))
       Case UCase(pDate) = "DATE RANGE": smfGetEconData = Trim(smfStrExtr(sDataE, "Date Range:", Chr(13)))
       Case UCase(pDate) = "LAST UPDATED": smfGetEconData = Trim(smfStrExtr(sDataE, "Last Updated:", Chr(13)))
       Case UCase(pDate) = "NOTES"
            smfGetEconData = Trim(smfStrExtr(sDataE, "Notes:", "DATE"))
            iLen = 0
            While iLen <> Len(smfGetEconData)
               iLen = Len(smfGetEconData)
               smfGetEconData = Replace(smfGetEconData, "  ", " ")
               Wend
       End Select
    If smfGetEconData <> "" Then Exit Function
    
    '--------------------------------> Get title of file
    Dim vData(1 To 3) As Variant
    iPos1 = InStr(sDataE, "Title:")
    iPos2 = InStr(iPos1, sDataE, Chr(13))
    sTitle = Trim(Mid(sDataE, iPos1 + 6, iPos2 - iPos1 - 6))
    
    '--------------------------------> Look for date in file
    iPos0 = InStr(sDataE, "Notes:")
    For n1 = pDate To pDate - 800 Step -1
        sDate = Application.WorksheetFunction.Text(n1, "yyyy-mm-dd")
        iPos1 = InStr(iPos0, sDataE, sDate)
        If iPos1 > 0 Then Exit For
        Next n1
    If iPos1 = 0 Then GoTo ErrorExit
       
    iPos2 = InStr(iPos1, sDataE, Chr(13))
    s3 = Mid(sDataE, iPos1 + 11, iPos2 - iPos1 - 11)
    On Error Resume Next
    s3 = smfConvertData(s3)
    On Error GoTo ErrorExit
    vData(1) = s3
    vData(2) = sDate
    vData(3) = sTitle
    smfGetEconData = vData
    Exit Function

ErrorExit:
    smfGetEconData = pError
    
    End Function

