Attribute VB_Name = "modGetEconData"
'@Lang VBA
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
    ' 2024.07.01 -- Web page change. Extract data from html instead of ".txt" format.
    '-----------------------------------------------------------------------------------------------------------*
    ' Example of usage:
    '
    '        =smfGetEconData("CURRNS", DATE(2007,3,24))
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim sURL As String, sDataE As String
    On Error GoTo ErrorExit

    '--------------------------------> See if web page has already been retrieved
'   sURL = "https://research.stlouisfed.org/fred2/data/" & UCase(pID) & ".txt"
'   sURL = "https://fred.stlouisfed.org/data/" & UCase(pID) & ".txt"       ' New URL as of 2016-06-20
    sURL = "https://fred.stlouisfed.org/data/" & UCase(pID)                ' New URL as of 2024-07-01
    
    If IsDate(pDate) Then
        pDate = Format(pDate, "yyyy-mm-dd")
        End If
        
    smfGetEconData = RCHGetTableCell(sURL, 1, ">" & pDate & "<")
    Exit Function

ErrorExit:
    smfGetEconData = pError
    
    End Function

