Attribute VB_Name = "modGetAdvFNElement"
Public sAdvFNPrefix As String                       ' URL prefix to use for AdvFN
Function smfGetADVFNElement(ByVal pTicker As String, _
                            ByVal pPeriod As String, _
                            ByVal pCells As Integer, _
                   Optional ByVal pFind1 As String = "", _
                   Optional ByVal pFind2 As String = "", _
                   Optional ByVal pError As Variant = "Error", _
                   Optional ByVal pType As Integer = 0) As Variant
                        
   '-----------------------------------------------------------------------------------------------------------*
   ' Function to return a financial statements data element from AdvFN
   '-----------------------------------------------------------------------------------------------------------*
   ' 2009.11.04 -- Created by Randy Harmelink
   ' 2009.12.19 -- Add smfConvertYahooTicker() processing
   ' 2012.01.30 -- Change AdvFn URL from "http://www..." to "http://us...."
   ' 2012.02.02 -- Make AdvFN URL prefix a variable
   ' 2012.02.04 -- Default AdvFN URL prefix to "www" unless external file exists to set it to something else
   ' 2013.10.16 -- Update for AdvFN URL changes
   ' 2013.10.17 -- Fix for above update
   ' 2014.12.02 -- Fix error processing when calling external functions
   ' 2015.03.05 -- Change "/exchanges/" search string to "/stock-market/"
   ' 2015.03.21 -- Allow exchange prefix on ticker symbol
   ' 2017.05.05 -- Add smfSetAdvFNPrefix() sub
   ' 2018.01.24 -- Change AdvFN URL from "http://" to "https://"
   ' 2023-01-24 -- Testing by Mel Pryor (ClimberMel@gmail.com)
   '               Currently this is returning all "--"
   '
   '-----------------------------------------------------------------------------------------------------------*
   ' Sample of use:
   '
   '    =smfGetADVFNElement("MMM","A",1,">Year End Date")
   '
   '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    vError = pError
    
    '------------------> Create labels for annual vs quarterly processing
    Dim sTicker As String, sExchange As String
    sTicker = smfConvertYahooTicker(pTicker, "ADVFN")
    sExchange = smfStrExtr(sTicker, "~", ":")
    If sExchange <> "" Then sTicker = smfStrExtr(sTicker, ":", "~")
    Select Case UCase(pPeriod)
       Case "A": sLabel1 = "&mode=annual_reports": sLabel2 = "start_date"
       Case "Q": sLabel1 = "&mode=quarterly_reports": sLabel2 = "istart_date"
       Case Else: vError = "Improper period -- should be A or Q": GoTo ErrorExit
       End Select
    If sAdvFNPrefix = "" Then LoadAdvFNPrefix
    If sExchange = "" Then
       sURL = "https://" & sAdvFNPrefix & ".advfn.com/p.php?pid=financials&mode=quarterly_reports&symbol=" & sTicker
       sExchange = smfStrExtr(RCHGetWebData(sURL, "/stock-market/", 50), "/stock-market/", "/")
       End If
    sURL = "https://" & sAdvFNPrefix & ".advfn.com/stock-market/" & sExchange & "/" & sTicker & "/financials?" _
         & sLabel1 & "&btn=" & sLabel2
    
    '------------------> Determine # of available periods and paging points (5 periods per page)
    nPeriods = smfConvertData(smfStrExtr(Right(smfGetTagContent(sURL, "select", 1, "Select start date", pError:=vError), 35), "'", "'"))
    If pCells = 999 Then
       smfGetADVFNElement = nPeriods + 1
       Exit Function
       End If
    If pCells > nPeriods + 1 Then GoTo ErrorExit
    nRawID = nPeriods - 5 * (Int((pCells - 1) / 5) + 1)
    nColumn = (200 - pCells) Mod 5 + 1
    Select Case nRawID
       Case Is < 0
            If nPeriods < 5 Then
               nPageId = ""
            Else
               nPageId = "&" & sLabel2 & "=0"
               End If
            nColumn = nColumn + nRawID + 1
       Case Is = nPeriods - 5
            nPageId = ""
       Case Else
            nPageId = "&" & sLabel2 & "=" & (nRawID + 1)
       End Select
    
    '------------------> Return data element
    smfGetADVFNElement = RCHGetTableCell(sURL & nPageId, nColumn, pFind1, pFind2, , , , "</TABLE", , vError)
    
    Exit Function

ErrorExit: smfGetADVFNElement = vError
                   
   End Function
Sub smfSetAdvFNPrefix(p1 As String)
    sAdvFNPrefix = p1
    End Sub

Sub LoadAdvFNPrefix()
    sAdvFNPrefix = "www"
    On Error GoTo ErrorExit
    Open ThisWorkbook.Path & "\smf-AdvFN-Prefix.txt" For Input As #1
    Line Input #1, sAdvFNPrefix
    Close #1
ErrorExit:
    End Sub


