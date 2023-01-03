Attribute VB_Name = "smfCreateComment_"
Public Function RCHCreateComment(pTicker As String, _
                  Optional ByVal pChoice As Integer = 1, _
                  Optional ByVal pWidth As Integer = 0, _
                  Optional ByVal pHeight As Integer = 0, _
                  Optional ByVal pVisible As Integer = 0, _
                  Optional ByVal pTop As Integer = 1, _
                  Optional ByVal pLeft As Integer = 1, _
                  Optional ByVal pScale As Single = 1#, _
                  Optional ByVal pText As String = "", _
                  Optional ByVal pReturn As String = "Chart")
Attribute RCHCreateComment.VB_Description = "Creates a comment box for the cell that can contain text and/or an image (e.g. a stock chart). "

    '-----------------------------------------------------------------------------------------------------------*
    ' Function to create a comment object and insert image/text into it
    '-----------------------------------------------------------------------------------------------------------*
    ' 2007.01.17 -- Change CCur() usage to CDec() because of precision issues
    ' 2007.09.24 -- Set the comment line color to white in order to "hide" it
    ' 2009.06.26 -- Add ActiveWorkbook.Name error check
    ' 2016.05.18 -- Change CDec() to smfCDec() to ease transition between operating systems
    ' 2018.01.24 -- Change AdvFN URL from "http://" to "https://"
    ' 2018.12.27 -- Change "http://stockcharts.com" domains to "https://c.stockcharts.com"
    '-----------------------------------------------------------------------------------------------------------*
    ' Examples of usage:
    '
    '    =RCHCreateComment("MMM")
    '    =RCHCreateComment("MMM",1,350,390,1,1,1)
    '
    '-----------------------------------------------------------------------------------------------------------*
    
    Set oCell = Cells(Application.Caller.Cells.Row, Application.Caller.Cells.Column)
    If ActiveWorkbook.Name <> Application.Caller.Parent.Parent.Name Then Exit Function
    If ActiveSheet.Name <> Application.Caller.Worksheet.Name Then Exit Function
    On Error Resume Next
    oCell.Comment.Delete
    On Error GoTo 0
    Select Case True
       Case UCase(pTicker) = "NONE": RCHCreateComment = "None": Exit Function
       Case pChoice = 0
            sURL = ""
            If pWidth = 0 Then pWidth = 300
            If pHeight = 0 Then pHeight = 200
       Case pTicker = "": GoTo ErrorExit
       Case pScale <= 0: GoTo ErrorExit
       Case pWidth < 0: GoTo ErrorExit
       Case pHeight < 0: GoTo ErrorExit
       Case pChoice = 1      ' Daily Chart of Gallery View from StockCharts
            sURL = "https://c.stockcharts.com/c-sc/sc?chart=" & pTicker & ",uu[h,a]daclyyay[pb50!b200!f][vc60][iue12,26,9!lc20]"
            If pWidth = 0 Then pWidth = 350 * pScale
            If pHeight = 0 Then pHeight = 390 * pScale
       Case pChoice = 2      ' P&F Chart from StockCharts
            sURL = "https://stockcharts.com/def/servlet/SharpChartv05.ServletDriver?chart=" & pTicker & ",pltad[pa][da][f!3!!]&pnf=y"
            If pWidth = 0 Then pWidth = 390 * pScale
            If pHeight = 0 Then pHeight = 314 * pScale
       Case pChoice = 3      ' 6-month Candleglance Chart from StockCharts
            sURL = "https://c.stockcharts.com/c-sc/sc?chart=" & pTicker & ",uu[305,a]dacayaci[pb20!b50][dc]"
            If pWidth = 0 Then pWidth = 229 * pScale
            If pHeight = 0 Then pHeight = 132 * pScale
       Case pChoice = 4      ' 6-month chart from Business Week Online
            sURL = "https://c.stockcharts.com/c-sc/sc?s=" & pTicker & "&p=D&yr=0&mn=6&dy=0&i=t94339682869&r=4806"
            If pWidth = 0 Then pWidth = 638 * pScale
            If pHeight = 0 Then pHeight = 501 * pScale
       Case pChoice = 5      ' 6-month chart Rule #1 Technicals from StockCharts
            sURL = "https://c.stockcharts.com/c-sc/sc?s=" & pTicker & "&p=D&yr=0&mn=6&dy=0&i=t39628903145&r=9933"
            If pWidth = 0 Then pWidth = 350 * pScale
            If pHeight = 0 Then pHeight = 360 * pScale
       Case pChoice = 97
            sURL = "https://www.advfn.com/p.php?pid=financialgraphs" '&a0=13&a1=13&a2=10&a3=8&a4=8&a5=10"
            aSplit = Split(pTicker, ",")
            iNbr = UBound(aSplit, 1)
            If iNbr <> 4 Then GoTo ErrorExit
            For i1 = 0 To iNbr
                sURL = sURL & "&a" & i1 & "=" & aSplit(i1)
                Next i1
            If pWidth = 0 Then pWidth = 263 * pScale
            If pHeight = 0 Then pHeight = 169 * pScale
       Case pChoice = 98
            sURL = "http://ogres-crypt.com/php/chart.php?d="
            aSplit = Split(pTicker, ",")
            iNbr = UBound(aSplit, 1)
            iMax = 0.01
            iMin = 999999999
            For i1 = 0 To iNbr
                iTemp = 0
                On Error Resume Next
                iTemp = smfCDec(aSplit(i1))
                On Error GoTo 0
                If (iTemp > iMax) Then iMax = iTemp
                If (iTemp < iMin And iTemp > 0) Then iMin = iTemp
                Next i1
            For i1 = 0 To iNbr
                iTemp = 0
                On Error Resume Next
                iTemp = smfCDec(aSplit(i1))
                On Error GoTo 0
                iTemp = IIf(iTemp > 0, 1 + 97 * (iTemp - iMin) / (iMax - iMin), 0)
                sURL = sURL & CInt(iTemp) & IIf(i1 = iNbr, "", ",")
                Next i1
            If pWidth = 0 Then pWidth = 36 * iNbr * pScale
            If pHeight = 0 Then pHeight = 90 * pScale
       Case pChoice = 99
            sURL = pTicker
            If pWidth = 0 Then pWidth = 400
            If pHeight = 0 Then pHeight = 300
       Case Else: GoTo ErrorExit
       End Select
    oCell.AddComment ("")
    If sURL <> "" Then oCell.Comment.Shape.Fill.UserPicture sURL
    oCell.Comment.Text Text:=IIf(pText = "", Chr(32), pText)
    oCell.Comment.Shape.Width = pWidth
    oCell.Comment.Shape.Height = pHeight
    oCell.Comment.Shape.Top = pTop + oCell.Top
    oCell.Comment.Shape.Left = pLeft + oCell.Left
    oCell.Comment.Shape.Line.Visible = False             ' Doesn't work
    oCell.Comment.Shape.Line.ForeColor.SchemeColor = 9   ' Set line color to background color
    oCell.Comment.Shape.Shadow.Visible = False
    oCell.Comment.Visible = IIf(pVisible = 1, True, False)
    RCHCreateComment = pReturn
    Exit Function
ErrorExit:
    RCHCreateComment = "Error"
    End Function


