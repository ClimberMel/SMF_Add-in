Attribute VB_Name = "smfGetGoogleHistoryCSV_"
'@Lang VBA
Option Explicit

Function smfGetGoogleHistoryCSV(ByVal pTicker As String, _
                   Optional ByVal pStartDate As Variant = "", _
                   Optional ByVal pEndDate As Variant = "", _
                   Optional ByVal pPeriod As String = "d", _
                   Optional ByVal pRows As Integer = 10000, _
                   Optional ByVal pCols As Integer = 7)

    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to process CSV file from Google
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.06.12 -- Added
    ' 2017.11.30 -- Change URL ("www" to "finance"
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
    Dim dBegin As Variant, dEnd As Variant
    vData(1, 1) = "Error on starting date: " & pStartDate
    Select Case True
          Case VarType(pStartDate) = vbDate Or VarType(pStartDate) = vbDouble
               dBegin = pStartDate
          Case pStartDate = ""
               dBegin = DateValue("1/1/1970")
          Case Else
               dBegin = DateValue(pStartDate)
          End Select
    vData(1, 1) = "Error on ending date: " & pEndDate
    Select Case True
          Case VarType(pEndDate) = vbDate Or VarType(pEndDate) = vbDouble
               dEnd = Int(pEndDate)
          Case pEndDate = ""
               dEnd = Int(Now)
          Case Else
               dEnd = Int(DateValue(pEndDate))
          End Select
     If dBegin > dEnd Then
        vData(1, 1) = "Error: Starting date cannot be after ending date: " & pStartDate & "," & pEndDate
        GoTo ErrorExit
        End If
       
    '------------------> Process period
    ' No processing at this point
    
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
    sURL = "http://finance.google.com/finance/historical?output=csv&q=" & UCase(pTicker) & _
           "&startdate=" & Format(dBegin, "mmm d, yyyy") & "&enddate=" & Format(dEnd, "mmm d, yyyy")
    vData = smfGetCSVFile(sURL, ",", iRows, iCols)

ErrorExit:
    smfGetGoogleHistoryCSV = vData
                   
   End Function

