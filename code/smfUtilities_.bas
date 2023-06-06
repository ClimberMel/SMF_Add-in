Attribute VB_Name = "smfUtilities_"
Public Const kPages = 1000                      ' Number of data pages to save
Public Const kUnix1970 As Long = 25569          ' CDbl(DateSerial(1970, 1, 1))
Public vError As Variant                        ' Value to return if error
Public aData(1 To kPages, 1 To 2) As String     ' Saved web page data (2) and its ticker-source (1)
Public sData(1 To 3) As String                  ' 1 = Raw data, 2 = Stripped data, 3 = Upper case of stripped data
Public sLog As String
Public sWebCache As String                      ' Set to "N" to force SMF to web pages for selected ranges
Public bASync As Boolean                        ' Set to TRUE for Asyncrhonous XMLHTTP processing


Public Function smfLogInternetCalls(pLog As String)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2014.05.24 -- Created. Function to control whether URL calls are logged into a CSV file
    ' 2016.05.18 -- Add Application.PathSeparator to ease transition between operating systems
    '-----------------------------------------------------------------------------------------------------------*
    On Error Resume Next
    sLog = UCase(pLog)
    Select Case sLog
       Case "Y": smfLogInternetCalls = "Logging on"
       Case "DELETE"
            Kill ThisWorkbook.Path & Application.PathSeparator & "smf-log.csv"
            sLog = "N"
            smfLogInternetCalls = "Logging off, file deleted"
       Case "RESET"
            Kill ThisWorkbook.Path & Application.PathSeparator & "smf-log.csv"
            sLog = "Y"
            smfLogInternetCalls = "Logging on, file reset"
       Case Else: smfLogInternetCalls = "Logging off"
       End Select
    End Function

Public Sub smfOpenLogFile()
    '-----------------------------------------------------------------------------------------------------------*
    ' 2014.05.24 -- Created. Macro to open and format SMF log file
    ' 2014.05.25 -- Added With statement
    ' 2014.05.25 -- Open in ReadOnly mode and use worksheet specific instructions
    ' 2016.05.18 -- Add Application.PathSeparator to ease transition between operating systems
    '-----------------------------------------------------------------------------------------------------------*
    Dim wb As Workbook
    Set wb = Workbooks.Open(ThisWorkbook.Path & Application.PathSeparator & "smf-log.csv", False, True)
    With wb.ActiveSheet
         .Range("A1").EntireRow.Insert
         .Range("A1").Value = "Time Stamp"
         .Range("B1").Value = "Duration"
         .Range("C1").Value = "Called URL"
         .Columns("A:A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
         .Columns("A:A").HorizontalAlignment = xlCenter
         .Columns("B:B").NumberFormat = "0.0000"
         .Columns("B:B").HorizontalAlignment = xlRight
         .Columns("C:C").ColumnWidth = 99.86
         .Range("A2").Select
         ActiveWindow.FreezePanes = True
         End With
    
    End Sub

Public Sub smfForceRecalculation()
Attribute smfForceRecalculation.VB_ProcData.VB_Invoke_Func = "R\n14"
    '-----------------------------------------------------------------------------------------------------------*
    ' 2016.05.29 -- Add reset of iCookieInit global variable, for loading Mac cookies
    ' 2017.07.23 -- Remove iMorningStar variable
    ' 2018.08.27 -- ERASE arrays rather than resetting the individual items
    '-----------------------------------------------------------------------------------------------------------*
    sAdvFNPrefix = ""
    iInit = 0
    iCookieInit = 0         ' Reset Mac Cookie flag
    Erase aData, aGuruFocusItems
    'aGuruFocusItems(1) = "" ' Reset stored GuruFocus array
    'For i1 = 1 To kPages
        'aData(i1, 1) = ""  ' Reset stored ticker array
        'Next i1
    If Val(Application.Version) < 10 Then
       Application.CalculateFull
    Else
       Application.CalculateFullRebuild
       End If
    End Sub

Public Sub smfASyncOn() ' Turn Asynchronous XMLHTTP on
    bASync = True
    End Sub

Public Sub smfASyncOff() ' Turn Asynchronous XMLHTTP off
    bASync = False
    End Sub

Public Function RCHGetURLData1(pURL As String, _
                     Optional ByVal pType As String = "GET") As String
                     
    '-----------------------------------------------------------------------------------------------> Version 2.0i
    ' 2008.07.18 -- Expand oHTTP.Status selections for "OK" to include zero
    ' 2009.01.26 -- Allow "GET" or "POST" requests
    ' 2014.06.13 -- Add bASync parameter
    ' 2017.05.01 -- Add "User-Agent" option
    ' 2023-05-05 -- Mel Pryor
    ' 2023-05-05 -- Note on oHTTP.ReadyState
    '               0   UNSENT  Client has been created. open() not called yet.
    '               1   OPENED  open() has been called.
    '               2   HEADERS_RECEIVED    send() has been called, and headers and status are available.
    '               3   LOADING     Downloading; responseText holds partial data.
    '               4   DONE    The operation is complete.
    '-----------------------------------------------------------------------------------------------------------*
    On Error GoTo ErrorExit
    Dim oHTTP As New xmlhttp
    oHTTP.Open pType, pURL, bASync
    oHTTP.setRequestHeader "User-Agent", "XMLHTTP/1.0"
    oHTTP.send
    Do While bASync
       DoEvents
       If oHTTP.readyState = 4 Then Exit Do
       Loop
    Select Case oHTTP.Status
       Case 0: RCHGetURLData1 = oHTTP.responseText
       Case 200: RCHGetURLData1 = oHTTP.responseText
       Case Else: GoTo ErrorExit
       End Select
    Exit Function
ErrorExit:
    RCHGetURLData1 = vError
    End Function

Public Function RCHGetURLData2(pURL As String) As String
    '-----------------------------------------------------------------------------------------------------------*
    ' 2018.08.14 -- Change Busy to a combination of Busy and ReadyState
    '-----------------------------------------------------------------------------------------------------------*
    On Error GoTo ErrorExit
    Dim oIE As Object
    Set oIE = CreateObject("InternetExplorer.Application")
    oIE.Visible = False
    oIE.Navigate pURL

    Application.Wait Now + (1 / 864000)
    While oIE.Busy Or oIE.readyState <> 4
       Application.Wait Now + (1 / 864000)
       DoEvents
       Wend
    Application.Wait Now + (1 / 864000)
    Do While oIE.Busy Or oIE.readyState <> 4
       Application.Wait Now + (1 / 864000)
       DoEvents
       Loop
    RCHGetURLData2 = oIE.Document.documentElement.outerHTML
    oIE.Quit
       
    
    'With oIE
    '    .Navigate pURL
    '    Do Until Not .Busy
    '        DoEvents
    '        Loop
    '    RCHGetURLData2 = .Document.documentElement.outerHTML
    '    .Quit
    '    End With
    
    Set oIE = Nothing
    Exit Function
ErrorExit:
    RCHGetURLData2 = vError
    End Function

Public Function RCHGetURLData3(pURL As String) As String
    '-----------------------------------------------------------------------------------------------> Version 2.0i
    ' 2009.01.26 -- Drop ".Document" qualifier
    '-----------------------------------------------------------------------------------------------> Version 2.0k
    ' 2009.07.13 -- Add fnWait call
    '-----------------------------------------------------------------------------------------------------------*
    On Error GoTo ErrorExit
    Dim oHTML As New HTMLDocument
    Set oDoc = oHTML.createDocumentFromUrl(pURL, vbNullString)
    Do: DoEvents: Loop Until oDoc.readyState = "complete"
    Call fnWait(2)  ' Wait for JavaScript to run on page?
    RCHGetURLData3 = oDoc.documentElement.outerHTML
    Exit Function
ErrorExit:
    RCHGetURLData3 = vError
    End Function

Public Function RCHGetURLData(ByVal pURL As String, _
                     Optional ByVal pUseIE As Integer = 0) As String
                     
    '-----------------------------------------------------------------------------------------------> Version 2.0i
    ' 2009.01.26 -- Add pUseIE options of 2 and 3
    ' 2009.03.16 -- Add documentation
    ' 2014.05.24 -- Add CSV output for logging of data requests
    ' 2014.05.25 -- Add double quotes around URL
    ' 2016.05.18 -- Add Application.PathSeparator to ease transition between operating systems
    ' 2016.05.18 -- Add call to RCHGetURLData1Mac for Mac usage
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim i1 As Integer
    Dim tStart As Single, tEnd As Single, dDate As Date
    dDate = Now
    tStart = Timer
    
#If Mac Then
    RCHGetURLData = RCHGetURLData1Mac(pURL, pUseIE)
#Else
    Select Case True
       Case pUseIE = 1: RCHGetURLData = RCHGetURLData2(pURL)                  ' IE Object
       Case pUseIE = 2: RCHGetURLData = RCHGetURLData3(pURL)                  ' HTMLDocument
       Case pUseIE = 3: RCHGetURLData = RCHGetURLData1(pURL, "POST")          ' XMLHTTP Post
       Case Else: RCHGetURLData = RCHGetURLData1(pURL, "GET")                 ' XMLHTTP Get
       End Select
#End If
    
    If sLog = "Y" Then
       tEnd = Timer
       i1 = FreeFile()
       Open ThisWorkbook.Path & Application.PathSeparator & "smf-log.csv" For Append As #i1
       Print #i1, dDate & "," & (tEnd - tStart) & ",""" & Left(pURL, 150) & """"
       Close #i1
       End If
    
    End Function

Public Function smfCDec(ByVal pString As String) As Variant
    
    '-----------------------------------------------------------------------------------------------------------*
    ' 2016.05.18 -- Add routine to ease transition between operating systems
    ' 2023-02-21 -- Mel Pryor (ClimberMel@gmail.com)
    '               Takes a string and converts it to a decimal number (or currency number on Mac)
    '-----------------------------------------------------------------------------------------------------------*
    
    smfCDec = pString
    On Error Resume Next
    
    #If Mac Then
        smfCDec = CCur(smfCDec)
    #Else
        smfCDec = CDec(smfCDec)
    #End If

    End Function
                         

Public Function smfGetWebPage(ByVal pURL As String, _
                     Optional ByVal pUseIE As Integer = 0, _
                     Optional ByVal pConvType As Integer = 0) As String
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.02.16 -- Add routine
    ' 2011.04.27 -- Add HTML codes &#48; thru &#57;
    ' 2014.05.31 -- Add sWebCache variable to allow re-retrieval of a web page
    ' 2017.11.08 -- Modify sWebCache processing
    '-----------------------------------------------------------------------------------------------------------*
    For iData = 1 To kPages
        Select Case True
           Case aData(iData, 1) = "" Or (aData(iData, 1) = pUseIE & ":" & pURL And sWebCache = "N")
                'sWebCache = "Y"
                s2 = RCHGetURLData(pURL, pUseIE)
                Select Case pConvType
                     Case 0
                          s2 = Replace(s2, "&amp;", "&")
                          s2 = Replace(s2, "&nbsp;<b>", "<b> ")
                          s2 = Replace(s2, "&nbsp;", " ")
                          s2 = Replace(s2, Chr(9), " ")
                          s2 = Replace(s2, Chr(10), "")
                          s2 = Replace(s2, Chr(13), "")
                          s2 = Replace(s2, "&#48;", "0")
                          s2 = Replace(s2, "&#49;", "1")
                          s2 = Replace(s2, "&#50;", "2")
                          s2 = Replace(s2, "&#51;", "3")
                          s2 = Replace(s2, "&#52;", "4")
                          s2 = Replace(s2, "&#53;", "5")
                          s2 = Replace(s2, "&#54;", "6")
                          s2 = Replace(s2, "&#55;", "7")
                          s2 = Replace(s2, "&#56;", "8")
                          s2 = Replace(s2, "&#57;", "9")
                          s2 = Replace(s2, "&#150;", Chr(150))
                          s2 = Replace(s2, "&#151;", "-")
                          s2 = Replace(s2, "&mdash;", "-")
                          s2 = Replace(s2, "&#160;", " ")
                          s2 = Replace(s2, Chr(160), " ")
                          s2 = Replace(s2, "<TH", "<td")
                          s2 = Replace(s2, "</TH", "</td")
                          s2 = Replace(s2, "<th", "<td")
                          s2 = Replace(s2, "</th", "</td")
                     Case 1
                          s2 = Replace(s2, Chr(10), Chr(13))
                     End Select
                Select Case pURL
                   Case "https://finance.yahoo.com/advances"
                        s2 = Replace(s2, "<sup>1</sup>", "")
                   End Select
                aData(iData, 1) = pUseIE & ":" & pURL
                aData(iData, 2) = s2
                smfGetWebPage = s2
                Exit Function
           Case aData(iData, 1) = pUseIE & ":" & pURL
                smfGetWebPage = aData(iData, 2)
                Exit Function
           Case iData = kPages
                smfGetWebPage = "Error -- Too many web page retrievals"
                Exit Function
           End Select
        Next iData
    smfGetWebPage = "Error"
    End Function

Public Function smfGetAData(p1 As Integer, p2 As Integer)
   smfGetAData = Left(aData(p1, p2), 32767)
   End Function

Public Sub smfFixLinks()
    '-----------------------------------------------------------------------------------------------------------*
    ' 2012.01.02 -- Expand to do all sheets in workbook
    '-----------------------------------------------------------------------------------------------------------*
    Dim Sht As Worksheet
    For Each Sht In Worksheets
        Sht.Cells.Replace _
            What:="'*\RCH_Stock_Market_Functions.xla'!", _
            Replacement:="", _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            MatchCase:=False, _
            SearchFormat:=False, _
            ReplaceFormat:=False
        Next Sht
    End Sub

Public Function fnWait(iSeconds As Integer)
   Dim varStart As Variant
   varStart = Timer
   Do While Timer < varStart + iSeconds
      DoEvents
      Loop
End Function

Function smfStrExtr(pString As String, _
                    pStart As String, _
                    pEnd As String, _
           Optional pConvert As Integer = 0)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2010.01.22 -- Add function
    ' 2010.06.06 -- Add error checking
    ' 2011.07.12 -- Add dummy characters to represent start and end of input string
    ' 2017.07.09 -- Add pConvert parameter
    ' 2023-01-29 -- Mel Pryor (ClimberMel@gmail.com)
    ' 2023-02-06 -- when called from smfGetYahooHistory with formula d1 = Int(smfUnix2Date(smfStrExtr(s1, """date"":", ",")))
    '               it would exit both this function and the calling function
    '-----------------------------------------------------------------------------------------------------------*
    ' If pConvert = 1: Calls smfConvertData
    '-----------------------------------------------------------------------------------------------------------*
    If pStart = "~" Then
       iPos1 = 1
       iPos3 = 2
    Else
       iPos1 = InStr(pString, pStart) + Len(pStart)
       iPos3 = iPos1
       If iPos1 = Len(pStart) Then
          smfStrExtr = ""
          Exit Function
          End If
       End If
    If pEnd = "~" Then iPos2 = Len(pString) + 1 Else iPos2 = InStr(iPos3, pString, pEnd)
    If iPos2 = 0 Then
       smfStrExtr = ""
       Exit Function
       End If
    smfStrExtr = Mid(pString, iPos1, iPos2 - iPos1)
    If pConvert = 1 Then smfStrExtr = smfConvertData(smfStrExtr)
EndStr:
    
End Function

Function smfEval(pData As String)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2012.01.21 -- Add function
    '-----------------------------------------------------------------------------------------------------------*
    smfEval = "Error"
    smfEval = Evaluate(pData)
    End Function

Function smfEvaluateTwice(pData As String)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2012.01.21 -- Add function
    '-----------------------------------------------------------------------------------------------------------*
    smfEvaluateTwice = "Error"
    smfEvaluateTwice = Evaluate(pData)
    If smfEvaluateTwice = pData Then smfEvaluateTwice = Evaluate("=" & pData)
    End Function

Function smfJoin(myRange As Range, myDelimiter As String)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.02.16 -- Add function
    '-----------------------------------------------------------------------------------------------------------*
    smfJoin = ""
    For Each oCell In myRange
        If smfJoin <> "" And oCell.Value <> "" Then smfJoin = smfJoin & myDelimiter
        smfJoin = smfJoin & oCell.Value
        Next oCell
    End Function

Public Function smfWord(ByVal Haystack As String, _
                        ByVal Occurrence As Long, _
               Optional ByVal Delimiter As String = " ", _
               Optional ByVal pConvert As Integer = 0)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.02.16 -- Add function
    ' 2017.09.20 -- Add pConvert parameter
    ' 2023-01-22 -- Mel Pryor (ClimberMel@gmail.com)
    '               Called by --> Verify and Process pNames parameter (Headers)
    '               in smfGetYahooHistory to extract the Header name based on position
    '-----------------------------------------------------------------------------------------------------------*
    On Error GoTo ErrorHandler
    smfWord = Split(Haystack, Delimiter)(Occurrence - 1)
    If pConvert = 1 Then smfWord = smfConvertData(smfWord)
    Exit Function
ErrorHandler:
    smfWord = ""
    End Function

Public Function smfStripHTML(ByVal sHTML As String) As String
    '-----------------------------------------------------------------------------------------------------------*
    ' 2014.04.07 -- Add function
    '-----------------------------------------------------------------------------------------------------------*
    Dim oDoc As HTMLDocument
    Set oDoc = New HTMLDocument
    oDoc.body.innerHTML = sHTML
    smfStripHTML = oDoc.body.innerText
    End Function
  
Public Function smfDate2Unix(ByVal pDate As Date) As Long
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.05.17 -- Add function
    ' 2023-02-21 -- Converts a date object to Unix style of seconds from Epoch
    '-----------------------------------------------------------------------------------------------------------*
    smfDate2Unix = DateDiff("s", kUnix1970, pDate)
    End Function

Public Function smfUnix2Date(pUnixDate As Long) As Date
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.05.17 -- Add function
    ' 2023-02-21 -- Converts a Unix style date back to a regular datetime object
    '-----------------------------------------------------------------------------------------------------------*
    smfUnix2Date = DateAdd("s", pUnixDate, kUnix1970)
    End Function

Public Function smfUnix2DateStr(pUnixDate As Long, _
                       Optional pFormat As String = "yyyy-mm-dd")
    '-----------------------------------------------------------------------------------------------------------*
    ' 2023-06-04 -- Add function
    '               Converts a Unix style date to a regular datetime object in string format
    '-----------------------------------------------------------------------------------------------------------*
    ' DateAdd(interval, number, date)
    '   interval = s is second
    '   pUnixDate is the Unix Date passed to the function
    '   kUnix1970 = 25569          "CDbl(DateSerial(1970, 1, 1))" from Constants at top
    ' So the following will create a serial date for Excel from the Unix date provided
    '-----------------------------------------------------------------------------------------------------------*
    '
    ' Example:
    '   =smfUnix2DateStr(Date)
    '   returns a string of the UNIX date in format "2023-06-04"
    '
    '   =smfUnix2DateStr(Date, "d/m/yy")
    '   returns a string of the UNIX date in format "4/6/23"
    '
    '   =smfUnix2DateStr(smfGetYahooJSONData("~~~~~","price","regularMarketTime",,"num"),"yyyy-mm-dd HH:MM")
    '   returns a string of the UNIX date from the json file as "2023-06-04 15:22"
    '-----------------------------------------------------------------------------------------------------------*
    
    unix2Date = DateAdd("s", pUnixDate, kUnix1970)
    
    ' Then this will return it in a String format based on the pFormat
    
    smfUnix2DateStr = Format(unix2Date, pFormat)

End Function

Public Function smfHTMLDecode(pString As String) As String
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.06.16 -- Add function
    '-----------------------------------------------------------------------------------------------------------*
    smfHTMLDecode = pString
    smfHTMLDecode = Replace(smfHTMLDecode, "&quot;", """")
    smfHTMLDecode = Replace(smfHTMLDecode, "&lt;", "<")
    smfHTMLDecode = Replace(smfHTMLDecode, "&gt;", ">")
    smfHTMLDecode = Replace(smfHTMLDecode, "&nbsp;", " ")
    smfHTMLDecode = Replace(smfHTMLDecode, "&apos;", "'")
    smfHTMLDecode = Replace(smfHTMLDecode, "&#39;", "'")
    smfHTMLDecode = Replace(smfHTMLDecode, "&#039;", "'")
    smfHTMLDecode = Replace(smfHTMLDecode, "&#150;", "")
    smfHTMLDecode = Replace(smfHTMLDecode, "&#151;", "-")
    smfHTMLDecode = Replace(smfHTMLDecode, "&mdash;", "-")
    smfHTMLDecode = Replace(smfHTMLDecode, "&#160;", " ")
    smfHTMLDecode = Replace(smfHTMLDecode, "&amp;", "&")
    End Function
