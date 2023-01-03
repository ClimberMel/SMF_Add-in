Attribute VB_Name = "smfGetBarChartHistoryCSV_"
Option Explicit

Function smfGetBarChartHistoryCSV(ByVal pTicker As String, _
                   Optional ByVal pPeriod As String = "d", _
                   Optional ByVal pSort As String = "d", _
                   Optional ByVal pRows As Integer = 1000, _
                   Optional ByVal pCols As Integer = 7)

    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to process CSV file from BarChart
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.06.02 -- Added
    '-----------------------------------------------------------------------------------------------------------*
                   
    ReDim vData(1 To 1, 1 To 1) As Variant
    vData(1, 1) = "Error"
    
    On Error GoTo ErrorExit
    
    '------------------> Null Return Item
    If pTicker = "None" Or pTicker = "" Then
       vData(1, 1) = "None"
       GoTo ErrorExit
       End If
       
    '------------------> Process pPeriod
    Dim sPeriod As String
    Select Case UCase(pPeriod)
       Case "": sPeriod = "daily"
       Case "D": sPeriod = "daily"
       Case "W": sPeriod = "weekly"
       Case "M": sPeriod = "monthly"
       Case "Q": sPeriod = "quarterly"
       Case "A": sPeriod = "yearly"
       Case Else
            vData(1, 1) = "Error on period: " & pPeriod
            GoTo ErrorExit
       End Select
       
    '------------------> Process pSort
    Dim sSort As String
    Select Case UCase(pSort)
       Case "": sSort = "desc"
       Case "A": sSort = "asc"
       Case "D": sSort = "desc"
       Case Else
            vData(1, 1) = "Error on sort: " & pSort
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
    sURL = "https://www.barchart.com/proxies/timeseries/queryeod.ashx?symbol=" & UCase(pTicker) & _
           "&data=" & sPeriod & "&maxrecords=" & iRows & "&volume=total&order=" & sSort & "&dividends=true&backadjust=false"
    vData = smfGetCSVFile(sURL, ",", iRows, iCols)

ErrorExit:
    smfGetBarChartHistoryCSV = vData
                   
   End Function

