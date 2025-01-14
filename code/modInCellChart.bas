Attribute VB_Name = "modInCellChart"
'@Lang VBA
Function smfInCellChart(pVector As Variant, _
               Optional pType As String = "Line", _
               Optional pColor As Long = 203) As String
    
    '-----------------------------------------------------------------------------------------------------------*
    ' Function to create "in cell" charts -- line charts, bar charts, or slope of linear regression
    '-----------------------------------------------------------------------------------------------------------*
    ' 2007.09.12 -- Adapted from http://www.dailydoseofexcel.com/archives/2006/02/05/in-cell-charting/
    ' 2007.09.13 -- Change rCaller .Height and .Width attributes to its MergeArea equivalents
    ' 2007.09.13 -- Add ability to pass a column of data instead of just a row
    '-----------------------------------------------------------------------------------------------> Version 2.0g
    ' 2007.09.18 -- Fix range/array processing for Trend/Min/Max functions
    ' 2007.09.24 -- Move code to delete previous shapes closer to top of module
    '-----------------------------------------------------------------------------------------------> Version 2.0h
    ' Examples of usage:
    '
    '        =smfInCellCharts(A14:I14)
    '        =smfInCellCharts(A14:I14, "Line",  203)
    '        =smfInCellCharts(A14:I14, "Bar",   203)
    '        =smfInCellCharts(A14:I14, "Slope", 203)
    '-----------------------------------------------------------------------------------------------------------*

    Const cMargin = 2       ' A margin to buffer the usable cell area
    Const cGap = 1          ' Size of gap to use between bar charts
    Dim rCaller As Range    ' The calling range for the function
    Dim oRange As Range, oShape As Shape
    Dim dMin As Double, dMax As Double
    Dim dBegX As Double, dBigY As Double
    Dim dEndX As Double, dEndY As Double
    Dim iSize As Integer
    Dim dHeight As Double, dWidth As Double, dTop As Double, dLeft As Double
 
    smfInCellChart = ""
    
    '----------------------------------> Identify the calling range
    Set rCaller = Application.Caller
    dHeight = rCaller.MergeArea.Height
    dWidth = rCaller.MergeArea.Width
    dLeft = rCaller.MergeArea.Left
    dTop = rCaller.MergeArea.Top
 
    '----------------------------------> Delete existing shapes in the calling range
    For Each oShape In rCaller.Worksheet.Shapes
        Set oRange = Intersect(Range(oShape.TopLeftCell, oShape.BottomRightCell), rCaller.MergeArea)
        If Not oRange Is Nothing Then
           If oRange.Address = Range(oShape.TopLeftCell, oShape.BottomRightCell).Address Then oShape.Delete
           End If
        Next oShape
    
    '----------------------------------> Copy input range/array to standard array area
    On Error Resume Next
    iSize = UBound(pVector)
    iSize = pVector.Count
    On Error GoTo 0

    ReDim vData(1 To iSize) As Double
    i = 0
    For Each oItem In pVector
        i = i + 1
        vData(i) = oItem
        Next oItem
    
    '------------------> Determine type of chart to create
    
    Select Case UCase(pType)
       Case "BAR": GoTo Bar_Chart
       Case "LINE": GoTo Line_Chart
       Case "SLOPE": GoTo Slope_Chart
       Case Else
            smfInCellChart = "Incorrect type of chart: " & pType
            GoTo ExitFunction
       End Select

'------------------> Create a bar chart
Bar_Chart:
    Dim sngLeft As Single, sngTop As Single, sngWidth As Single, sngHeight As Single
    Dim sngMin As Single, sngMax As Single, shp As Shape

    '------------------> Determine minimum and maximum chartable values
    dMin = Application.WorksheetFunction.Min(vData)
    dMax = Application.WorksheetFunction.Max(vData)
    If dMin > 0 Then dMin = 0
    If dMin = dMax Then
       dMin = dMin - 1
       dMax = dMax + 1
       End If

    '------------------> Draw the bar for each data point
    With rCaller.Worksheet.Shapes
         For i = 0 To iSize - 1
             sngIntv = (dHeight - (cMargin * 2)) / (dMax - dMin)
             sngLeft = cMargin + cGap + dLeft + (i * (dWidth - (cMargin * 2)) / iSize)
             sngTop = cMargin + dTop + (dMax - IIf(vData(i + 1) < 0, 0, vData(i + 1))) * sngIntv
             sngWidth = (dWidth - (cMargin * 2)) / iSize - (cGap * 2)
             sngHeight = Abs(vData(i + 1)) * sngIntv
             With .AddShape(msoShapeRectangle, sngLeft, sngTop, sngWidth, sngHeight)
                  If pColor > 0 Then .Fill.ForeColor.RGB = pColor Else .Fill.ForeColor.SchemeColor = -pColor
                  End With
             Next i
         End With

    GoTo ExitFunction

'------------------> Create a line chart
Line_Chart:

    '------------------> Determine minimum and maximum chartable values
    dMin = Application.WorksheetFunction.Min(vData)
    dMax = Application.WorksheetFunction.Max(vData)
    If dMin = dMax Then
       dMin = dMin - 1
       dMax = dMax + 1
       End If
    
    '------------------> Draw the lines for each pair of data points
    With rCaller.Worksheet.Shapes
         For i = 0 To iSize - 2
             dBegX = cMargin + dLeft + (i * (dWidth - (cMargin * 2)) / (iSize - 1))
             dBegY = cMargin + dTop + (dMax - vData(i + 1)) * (dHeight - (cMargin * 2)) / (dMax - dMin)
             dEndX = cMargin + dLeft + ((i + 1) * (dWidth - (cMargin * 2)) / (iSize - 1))
             dEndY = cMargin + dTop + (dMax - vData(i + 2)) * (dHeight - (cMargin * 2)) / (dMax - dMin)
             With .AddLine(dBegX, dBegY, dEndX, dEndY)
                  If pColor > 0 Then .Line.ForeColor.RGB = pColor Else .Line.ForeColor.SchemeColor = -pColor
                  End With
             Next i
         End With

    GoTo ExitFunction
    
'------------------> Create a chart of a linear regression slope line
Slope_Chart:

    '------------------> Create linear regression trend line
    vTrend = Application.WorksheetFunction.Trend(vData())

    '------------------> Determine minimum and maximum chartable values
    dMin = Application.WorksheetFunction.Min(vData, vTrend)
    dMax = Application.WorksheetFunction.Max(vData, vTrend)
    If dMin = dMax Then
       dMin = dMin - 1
       dMax = dMax + 1
       End If
    
    '------------------> Draw the regression line
    With rCaller.Worksheet.Shapes
         dBegX = cMargin + dLeft
         dBegY = cMargin + dTop + (dMax - vTrend(1)) * (dHeight - (cMargin * 2)) / (dMax - dMin)
         dEndX = dLeft + dWidth - cMargin
         dEndY = cMargin + dTop + (dMax - vTrend(iSize)) * (dHeight - (cMargin * 2)) / (dMax - dMin)
         With .AddLine(dBegX, dBegY, dEndX, dEndY)
              If pColor > 0 Then .Line.ForeColor.RGB = pColor Else .Line.ForeColor.SchemeColor = -pColor
              .Line.BeginArrowheadStyle = msoArrowheadOval
              .Line.BeginArrowheadLength = msoArrowheadShort
              .Line.BeginArrowheadWidth = msoArrowheadNarrow
              .Line.EndArrowheadStyle = msoArrowheadStealth
              End With
         End With
    
    GoTo ExitFunction
    
ExitFunction:
    End Function

