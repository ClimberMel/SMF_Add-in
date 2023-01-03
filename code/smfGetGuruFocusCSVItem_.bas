Attribute VB_Name = "smfGetGuruFocusCSVItem_"
Option Explicit
Function smfGetGuruFocusCSVItem(ByVal pTicker As String, _
                                ByVal pItem As String, _
                       Optional ByVal pPeriod As Variant = "TTM", _
                       Optional ByVal pError As Variant = "Error", _
                       Optional ByVal pType As Integer = 0) As Variant
                  
   '-----------------------------------------------------------------------------------------------------------*
   ' Function to return a data item from GuruFocus CSV file
   '-----------------------------------------------------------------------------------------------------------*
   ' 2017.09.26 -- Created by Randy Harmelink (rharmelink@gmail.com)
   ' 2017.10.12 -- Official release
   '-----------------------------------------------------------------------------------------------------------*
   ' Sample of use:
   '
   '    =smfGetGuruFocusCSVItem("MMM","Fiscal Period","TTM")
   '    =smfGetGuruFocusCSVItem("MMM","Fiscal Period","A0")
   '    =smfGetGuruFocusCSVItem("MMM","Fiscal Period","Q0")
   '
   '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    Dim vError As Variant, sURL As String, sLine As String
    
    smfGetGuruFocusCSVItem = pError
    
    If UCase(pTicker) = "NONE" Then
       smfGetGuruFocusCSVItem = "--"
       Exit Function
       End If
        
    '------------------> Extract line of desired item
    sURL = "https://www.gurufocus.com/download_financials_in_CSV.php?symbol=" & UCase(pTicker)
    Select Case UCase(pItem)
       Case "DOWNLOADED"
            smfGetGuruFocusCSVItem = "Downloaded " & smfStrExtr(RCHGetWebData(sURL, "Downloaded ", 3000), "~", ",")
            Exit Function
       Case "HEADER"
            smfGetGuruFocusCSVItem = "30 Year Financials " & smfStrExtr(RCHGetWebData(sURL, "30 Year Financials ", 3000), "~", "All Numbers")
            Exit Function
       Case "ALL NUMBERS"
            smfGetGuruFocusCSVItem = "All Numbers " & smfStrExtr(RCHGetWebData(sURL, "All Numbers ", 3000), "~", "Annual Data:")
            Exit Function
       Case "NAME"
            smfGetGuruFocusCSVItem = smfWord(smfStrExtr(RCHGetWebData(sURL, "Change log", 10000), "~", "Key Statistics:"), 2, ",")
            Exit Function
       Case "COUNTRY"
            smfGetGuruFocusCSVItem = smfWord(smfStrExtr(RCHGetWebData(sURL, "Change log", 10000), "~", "Key Statistics:"), 3, ",")
            Exit Function
       Case "SECTOR"
            smfGetGuruFocusCSVItem = smfWord(smfStrExtr(RCHGetWebData(sURL, "Change log", 10000), "~", "Key Statistics:"), 4, ",")
            Exit Function
       Case "INDUSTRY"
            smfGetGuruFocusCSVItem = smfWord(smfStrExtr(RCHGetWebData(sURL, "Change log", 10000), "~", "Key Statistics:"), 5, ",")
            Exit Function
       Case "KS-PB RATIO", "KS-PE RATIO", "KS-PS RATIO"
            sLine = Replace(RCHGetWebData(sURL, """" & Replace(pItem, "KS-", "") & """", 3000), "Growth Rates:", ",")
       Case "BOOK VALUE GROWTH (%)"
            sLine = smfStrExtr(RCHGetWebData(sURL, """" & pItem & """", 3000), "~", "30 Year Financials ")
       Case "PB RATIO", "PE RATIO", "PS RATIO"
            sLine = """" & pItem & """" & smfStrExtr(RCHGetWebData(sURL, """Free Cash Flow""", 20000), """" & pItem & """", "~")
       Case Else
            sLine = RCHGetWebData(sURL, """" & pItem & """", 3000)
       End Select
    If sLine = "" Then
       smfGetGuruFocusCSVItem = "Error -- pItem not found: " & pItem
       Exit Function
       End If
    sLine = smfStrExtr(Replace(Replace(sLine, ","""",", ","" "","), ","""",", ","" "",") & """""", """", """""")
    
    '------------------> Which data period?
    Select Case UCase(pPeriod)
       Case "TTM"
            smfGetGuruFocusCSVItem = smfWord(sLine, 32, """,""", 1)
       Case "Q0" To "Q9", "Q10" To "Q99", "Q100" To "Q119"
            smfGetGuruFocusCSVItem = smfWord(sLine, 153 - Mid(pPeriod, 2, 6), """,""", 1)
       Case "A0" To "A9", "A10" To "A29"
            smfGetGuruFocusCSVItem = smfWord(sLine, 31 - Mid(pPeriod, 2, 6), """,""", 1)
       Case 1 To 154
            smfGetGuruFocusCSVItem = smfWord(sLine, pPeriod, """,""", 1)
       Case Else
            smfGetGuruFocusCSVItem = "Error -- Invalid pPeriod parameter: " & pPeriod
       End Select
    
ErrorExit:
   End Function

