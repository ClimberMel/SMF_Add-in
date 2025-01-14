Attribute VB_Name = "modPictureTable"
'@Lang VBA
Function RCHImageTable(pTickers As Variant, _
        Optional ByVal pBreaks As Integer = -1, _
        Optional ByVal pChart As String = "6")
                           
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to create a table of images (e.g. Stock Charts)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2006.04.27 -- Created by Randy Harmelink (rharmelinkg@gmail.com)
    '-----------------------------------------------------------------------------------------------> Version 1.2
    ' > Example of an invocation to create a table of two normal StockCharts 6-month Candleglance charts:
    '
    '   =RCHImageTable("IBM,GE")
    '-----------------------------------------------------------------------------------------------------------*
    ' Notes:
    '
    ' Possible fundamental charts to add (Revenue Growth and EPS Growth -- need 26 breaks in between):
    ' > http://tools.morningstar.com/charts/MStarCharts.aspx?Security=MMM&bSize=460&Fundamental=RG&Options=F&Stock=&DateFrom=4/30/2005&DateTo=4/29/2006&FPrime=MMM
    ' > http://tools.morningstar.com/charts/MStarCharts.aspx?Security=MMM&bSize=460&Fundamental=EPSG&Options=F&Stock=&DateFrom=4/30/2005&DateTo=4/29/2006&FPrime=MMM
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    If pBreaks >= 0 Then
       sPrefix1 = ""
       sSuffix1 = ""
       sPrefix2 = ""
       sSuffix2 = Replace(String(pBreaks, "!"), "!", "<br>")
    Else
       sPrefix1 = "<table>"
       sSuffix1 = "</table>"
       sPrefix2 = "<tr><td>"
       sSuffix2 = "</td></tr>"
       End If
    Select Case UCase(pChart)
        Case "6": sDisplay = "<img src=""http://stockcharts.com/c-sc/sc?chart=~~~~~,uu[305,a]dacayaci[pb20!b50][dc]"">"
        Case "12": sDisplay = "<img src=""http://stockcharts.com/c-sc/sc?chart=~~~~~,uu[305,a]dacayaci[pb20!b50][dd]"">"
        Case "P&F": sDisplay = "<img src=""http://stockcharts.com/def/servlet/SharpChartv05.ServletDriver?chart=~~~~~,pltad[pa][da][f!3!!]&pnf=y"">"
        Case "SC1": sDisplay = "<img src=""http://stockcharts.com/c-sc/sc?s=~~~~~&p=D&b=3&g=0&id=t08330678207&r=4815"">"
        Case Else: sDisplay = pChart
        End Select
    sDisplay = sPrefix2 & sDisplay & sSuffix2
    
    RCHImageTable = sPrefix1
    Select Case VarType(pTickers)
        Case vbString
             sTickers = Split(pTickers, ",")
             For i1 = 0 To UBound(sTickers, 1)
                 If sTickers(i1) <> "" Then RCHImageTable = RCHImageTable & Replace(sDisplay, "~~~~~", sTickers(i1))
                 Next i1
        Case Is >= 8192
             For Each oCell In pTickers
                 If oCell.Value <> "" Then RCHImageTable = RCHImageTable & Replace(sDisplay, "~~~~~", oCell.Value)
                 Next oCell
        Case Else
            GoTo ErrorExit
        End Select
    RCHImageTable = RCHImageTable & sSuffix1
    
    Exit Function

ErrorExit:
    RCHImageTable = "Error"
    End Function

