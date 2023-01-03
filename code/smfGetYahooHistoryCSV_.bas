Attribute VB_Name = "smfGetYahooHistoryCSV_"
Option Explicit

Function smfGetYahooHistoryCSV(ByVal pTicker As String, _
                   Optional ByVal pStartDate As Variant = "", _
                   Optional ByVal pEndDate As Variant = "", _
                   Optional ByVal pPeriod As String = "d", _
                   Optional ByVal pRows As Integer = 10000, _
                   Optional ByVal pCols As Integer = 7)

    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to process CSV file from Yahoo
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.05.31 -- Added
    '-----------------------------------------------------------------------------------------------------------*
                   
    ReDim vData(1 To 1, 1 To 1) As Variant
    vData(1, 1) = "Error"
    
    On Error GoTo ErrorExit
    
    '------------------> Null Return Item
    If pTicker = "None" Or pTicker = "" Then
       vData(1, 1) = "None"
       GoTo ErrorExit
       End If
       
    '------------------> Verify and Process starting and ending dates
    Dim dBegin As Double, dEnd As Double
    vData(1, 1) = "Error on starting date: " & pStartDate
    Select Case True
          Case VarType(pStartDate) = vbDate Or VarType(pStartDate) = vbDouble
               dBegin = smfDate2Unix(pStartDate)
          Case pStartDate = ""
               dBegin = smfDate2Unix(DateValue("1/1/1970"))
          Case Else
               dBegin = smfDate2Unix(DateValue(pStartDate))
          End Select
    vData(1, 1) = "Error on ending date: " & pEndDate
    Select Case True
          Case VarType(pEndDate) = vbDate Or VarType(pEndDate) = vbDouble
               dEnd = smfDate2Unix(Int(pEndDate) + 1)
          Case pEndDate = ""
               dEnd = smfDate2Unix(Int(Now) + 1)
          Case Else
               dEnd = smfDate2Unix(Int(DateValue(pEndDate)) + 1)
          End Select
     If dBegin > dEnd Then
        vData(1, 1) = "Error: Starting date cannot be after ending date: " & pStartDate & "," & pEndDate
        GoTo ErrorExit
        End If
       
    '------------------> Process period
    Dim sPeriod As String, sEvent As String, sInterval As String
    sPeriod = UCase(pPeriod)
    Select Case sPeriod
       Case "D": sEvent = "history": sInterval = "1d"
       Case "W": sEvent = "history": sInterval = "1wk"
       Case "M": sEvent = "history": sInterval = "1mo"
       Case "S": sEvent = "split": sInterval = "1d"
       Case "V": sEvent = "div": sInterval = "1d"
       Case Else
            vData(1, 1) = "Error on period: " & pPeriod
            GoTo ErrorExit
       End Select
    
    '------------------> Determine size of array to return and initialize array
    Dim iRows As Integer, iCols As Integer, i1 As Integer, i2 As Integer
    iRows = pRows  ' Rows
    iCols = pCols  ' Columns
    On Error Resume Next
    iRows = Application.Caller.Rows.Count
    iCols = Application.Caller.Columns.Count
    On Error GoTo ErrorExit
  
    ReDim vData(1 To iRows, 1 To iCols) As Variant
    For i1 = 1 To iRows
        For i2 = 1 To iCols
            vData(i1, i2) = ""
            Next i2
        Next i1
    
    '------------------> Get CSV file
    Dim sURL As String
    sURL = "https://query1.finance.yahoo.com/v7/finance/download/" & UCase(pTicker) & _
           "?period1=" & dBegin & "&period2=" & dEnd & "&interval=" & sInterval & "&events=" & sEvent & "&crumb="
    vData = smfGetCSVFile(sURL, ",", iRows, iCols)

ErrorExit:
    smfGetYahooHistoryCSV = vData
                   
   End Function



Function smfGetYahooHistoryCSVData(Optional ByRef pURL As String = "https://query1.finance.yahoo.com/v7/finance/download/MMM?period1=1493610466&period2=1496202466&interval=1d&events=history&crumb=")

   '-----------------------------------------------------------------------------------------------------------*
   ' User defined function to download historical quotes data file from Yahoo
   '-----------------------------------------------------------------------------------------------------------*
   ' 2017.05.31 -- Adapted from code at http://www.xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
   '-----------------------------------------------------------------------------------------------------------*

   Dim sURL As String
   Dim sCrumb As String, sCookie As String, sResult As String
   Dim i1 As Integer, b1 As Boolean
   Dim oData As Object
   
   sURL = "https://finance.yahoo.com/lookup?s=%7B0%7D"
   Set oData = New WinHttp.WinHttpRequest
   b1 = True
   For i1 = 1 To 5
       With oData
            .Open "GET", sURL, b1
            Select Case i1
               Case 1: .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
               Case Else: .setRequestHeader "Cookie", sCookie
               End Select
            .send
            .waitForResponse
            Select Case i1
               Case 1
                    sCrumb = smfStrExtr(smfStrExtr(.responseText, "CrumbStore", "~"), """:""", """")
                    sCookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
                    sURL = pURL + sCrumb
               Case Else
                    sResult = .responseText
                    If Left(sResult, 4) = "Date" Then Exit For
               End Select
            End With
       b1 = False
       Next i1
   
    smfGetYahooHistoryCSVData = sResult
      
End Function
