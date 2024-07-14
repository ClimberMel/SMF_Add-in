Attribute VB_Name = "modGetHTMLTable"
'@Lang VBA
Const kDim3 = 20
Public Function RCHGetHTMLTable(ByVal pURL As String, _
                                ByVal pFind1 As String, _
                       Optional ByVal pDir1 As Integer = -1, _
                       Optional ByVal pFind2 As String = "", _
                       Optional ByVal pDir2 As Integer = 1, _
                       Optional ByVal pRowOnly As Boolean = False, _
                       Optional ByVal pDim1 As Integer = 0, _
                       Optional ByVal pDim2 As Integer = 0, _
                       Optional ByVal pType As Integer = 0, _
                       Optional ByVal pCalc As Integer = 1)
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to extract an HTML table from a web page
    '-----------------------------------------------------------------------------------------------------------*
    ' 2005.02.16 -- Fixed pDim1/pDim2 processing
    ' 2006.03.12 -- Fixed COLSPAN='n' interpretation error
    ' 2006.03.12 -- Add process to convert numeric table cells
    ' 2006.07.09 -- Add interpretation of <THEAD> and </THEAD> as if they were <TR> and </TR>
    ' 2006.07.09 -- Remove possibility of "empty" table cells (i.e. <TD></TD> or <TH></TH>)
    ' 2006.09.04 -- Add ability to just return a row of the table
    ' 2007.01.17 -- Change CCur() usage to CDec() because of precision issues
    ' 2007.09.18 -- Modify pDim1/pDim2 processing
    ' 2007.10.13 -- Add LEFT() function to table cell cannot exceed 255 bytes, which causes #VALUE in EXCEL
    ' 2008.03.16 -- Add pType parameter
    ' 2010.08.12 -- Modify pDim1/pDim2 processing so return size can be overridden from worksheet
    ' 2010.10.10 -- Added code to change HTML code &#151; to a normal hyphen
    ' 2010.10.22 -- Added code to change HTML code &mdash; to a normal hyphen
    ' 2011.02.16 -- Convert to use smfGetWebPage() function
    ' 2011.04.28 -- Change cDec() to smfConvertData()
    ' 2011.06.09 -- Added pCalc parameter
    ' 2014.03.27 -- Made pDir1, pFind2, and pDir2 parameters optional by giving them default values
    '-----------------------------------------------------------------------------------------------------------*
    ' > Sample invocation to grab "Price Target Summary" from Yahoo for ticker IBM:
    '
    '   =RCHGetHTMLTable("https://finance.yahoo.com/q/ao?s=IBM", "Mean Target", -3, "Mean Target", 1)
    '-----------------------------------------------------------------------------------------------------------*
    
    '------------------> Leave range alone?
    'If pCalc = 0 Then
    '   On Error Resume Next
    '   RCHGetHTMLTable = Range("A1").Offset(Application.Caller.Row - 1, Application.Caller.Column - 1).Resize(Application.Caller.Rows.Count, Application.Caller.Columns.Count)
    '   On Error GoTo ErrorExit
    '   Exit Function
    '   End If
    
    '------------------> Determine size of array to return
    kDim1 = pDim1  ' Rows
    kDim2 = pDim2  ' Columns
    If pDim1 = 0 Or pDim2 = 0 Then
       If pDim1 = 0 Then kDim1 = 10   ' Old default
       If pDim2 = 0 Then kDim2 = 10   ' Old default
       On Error Resume Next
       kDim1 = Application.Caller.Rows.Count
       kDim2 = Application.Caller.Columns.Count
       On Error GoTo ErrorExit
       End If
    
    '------------------> Initialize returning array
    ReDim vData(1 To kDim1, 1 To kDim2) As Variant
    Dim iTBMinRow(1 To kDim3)
    Dim iTBMaxCol(1 To kDim3)
    Dim iTRMinCol(1 To kDim3)
    Dim iTRMaxRow(1 To kDim3)
    
    For i1 = 1 To kDim1
        For i2 = 1 To kDim2
            vData(i1, i2) = ""
            Next i2
        Next i1
    For i1 = 1 To kDim3
        iTBMinRow(i1) = 0
        iTBMaxCol(i1) = 0
        iTRMinCol(i1) = 0
        iTRMaxRow(i1) = 0
        Next i1
    
    '------------------> Download web page
    sData1 = smfGetWebPage(pURL, pType, 0)
    sData2 = UCase(sData1)
    
    '------------------> Look for the start and the end of the desired data table(s) on the page
    iSel1 = InStr(sData2, UCase(pFind1))
    For i1 = 1 To Abs(pDir1)
        If pDir1 < 0 Then
           iSel1 = InStrRev(sData2, IIf(pRowOnly, "<TR", "<TABLE"), iSel1 - 1)
        Else
           iSel1 = InStr(iSel1 + 1, sData2, IIf(pRowOnly, "<TR", "<TABLE"))
           End If
        Next i1
    
    If pFind2 = "" Then pFind2 = pFind1
    iSel2 = InStr(sData2, UCase(pFind2))
    For i1 = 1 To Abs(pDir2)
        If pDir2 < 0 Then
           iSel2 = InStrRev(sData2, IIf(pRowOnly, "</TR", "</TABLE"), iSel2 - 1)
        Else
           iSel2 = InStr(iSel2 + 1, sData2, IIf(pRowOnly, "</TR", "</TABLE"))
           End If
        Next i1
    
    '------------------> Parse the table into rows and columns
   iTB = 1
   iTR = 1
   iRow = 0
   iCol = 0
   iPos1 = iSel1
   iTD = 0
   Do While True
      iPos1 = InStr(iPos1, sData2, "<")
      If iPos1 = 0 Or iPos1 > iSel2 Then Exit Do
      iPos2 = InStr(iPos1, sData2, ">")
      If iPos2 = 0 Or iPos2 < iPos1 Then Exit Do
      If Mid(sData2, iPos1, 6) = "<TABLE" Then
         iTD = 0                                ' Previous table cell start is not a data cell
         iTB = iTB + 1                          ' Start of new table
         iTBMinRow(iTB) = iRow                  ' Save row that table began at
         If iRow > 0 And iTB > 2 Then iRow = iRow - 1   ' Need next row to start on current row
      ElseIf Mid(sData2, iPos1, 7) = "</TABLE" Then
         If iTB > 0 Then
            If iTB = 2 Then
               iCol = 0
            Else
               iRow = iTBMinRow(iTB)            ' Restore row that table begain at
               iCol = iTBMaxCol(iTB)            ' Set column to max column used by table
               iTBMinRow(iTB) = 0
               iTBMaxCol(iTB) = 0
               End If
            iTB = iTB - 1                       ' End of current table
            End If
      ElseIf Mid(sData2, iPos1, 3) = "<TR" Or Mid(sData2, iPos1, 6) = "<THEAD" Then
         iTR = iTR + 1                          ' Start of new row
         iRow = iRow + 1                        ' Point to next row of array
         iTRMinCol(iTR) = iCol                  ' Save column that row began at
      ElseIf Mid(sData2, iPos1, 4) = "</TR" Or Mid(sData2, iPos1, 7) = "</THEAD" Then
         iTBMaxCol(iTB) = Application.WorksheetFunction.Max(iTBMaxCol(iTB), iCol)
         iCol = iTRMinCol(iTR)                  ' Restore column that the row started at, for next row
         iTR = iTR - 1                          ' End of current row
         If iTR = 0 Then Exit Do
         iTRMaxRow(iTR) = Application.WorksheetFunction.Max(iTRMaxRow(iTR + 1), iRow)
         iRow = iTRMaxRow(iTR)                  ' Set row to max row used during this row
      ElseIf Mid(sData2, iPos1, 3) = "<TD" Or Mid(sData2, iPos1, 3) = "<TH" Then
         iTD = iPos2 + 1                        ' Save possible start of cell data
         sTemp = Mid(sData2, iPos1, iPos2 - iPos1 + 1)
         iPos3 = InStr(sTemp, "COLSPAN=")
         If iPos3 > 0 Then
            iPos4 = InStr(iPos3, sTemp, " ")
            If iPos4 = 0 Then iPos4 = Len(sTemp)
            iColSpan = CInt(Replace(Replace(Mid(sTemp, iPos3 + 8, iPos4 - iPos3 - 8), """", ""), "'", ""))
         Else
            iColSpan = 1
            End If
      ElseIf Mid(sData2, iPos1, 4) = "</TD" Or Mid(sData2, iPos1, 4) = "</TH" Then
         If iTD > 0 Then
            iCol = iCol + 1
            sTemp = Mid(sData1, iTD, iPos1 - iTD)
            Do While True
               iPos3 = InStr(sTemp, "<")
               If iPos3 = 0 Then Exit Do
               iPos4 = InStr(sTemp, ">")
               If iPos4 = 0 Then Exit Do
               sTemp = Mid(sTemp, 1, iPos3 - 1) & Mid(sTemp, iPos4 + 1)
              Loop
            If iRow <= kDim1 And iCol <= kDim2 Then
               vData(iRow, iCol) = smfConvertData(Trim(Left(sTemp, 255)))
               End If
            iCol = iCol + iColSpan - 1
            iTD = 0
            End If
         End If
      iPos1 = iPos2 + 1
      Loop
    
ErrorExit:

    RCHGetHTMLTable = vData
    
    End Function


