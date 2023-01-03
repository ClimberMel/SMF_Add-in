Attribute VB_Name = "smfDownloadTable_"
Public Sub smfUpdateDownloadTable()
    '-----------------------------------------------------------------------------------------------------------*
    ' Macro to download data to fill in a 2-dimensional table
    '-----------------------------------------------------------------------------------------------------------*
    ' 2006.12.15 -- Created as a test process
    ' 2007.07.13 -- Add check for "X" to skip the replacement process
    ' 2007.07.13 -- Add progress information to status bar
    ' 2010.12.02 -- Add ability to refer to data in a prior column of the same row
    ' 2012.07.14 -- Fix element number processing for formula-based element definitions
    ' 2014.05.01 -- Rewrite to allow the update of only selected areas of data
    ' 2014.08.27 -- Modify error messages
    ' 2017.12.08 -- Allow for 50 (instead of 30) backward column reference
    ' 2018.06.23 -- Do native calculation of formula if it has a prefix of "=="
    ' 2020.03.09 -- Replace EVALUATE() function with smfEvaluateTwice(), a fix for Microsoft changes
    '-----------------------------------------------------------------------------------------------------------*
    ' Download table requires several setup items:
    ' 1. The upper left hand corner cell of the table needs to be named name "Ticker"
    ' 2. The cells below the "Ticker" cell should be filled in with ticker symbols, one per cell
    ' 3. The cells to the right of the "Ticker" cell should be filled with column titles
    ' 4. The cells above the column titles need to be filled in with SMF add-in formulas or element numbers.  Use
    '    five tildas as a substitute for a ticker symbol.  For example, any of the following text strings could be
    '    used to get "Market Capitalization" from Yahoo:
    '
    '        941
    '        RCHGetElementNumber("~~~~~", 941)
    '        RCHGetTableCell("https://finance.yahoo.com/q/ks?s=~~~~~",1,">Market Cap")
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim sFormula As String
    On Error GoTo ErrorExit
          
    Set rTicker = Range("Ticker")
    If Selection.Address = Selection.EntireColumn.Address Or (Selection.Rows.Count = 1 And Selection.Columns.Count = 1) Then
       nTickers = rTicker.End(xlDown).Row - rTicker.Row
       nRowOffset = 0
    Else
       nTickers = Selection.Rows.Count
       nRowOffset = Selection.Row - rTicker.Row - 1
       If nRowOffset < 0 Then
          MsgBox "The first highlighted row must be below the ""Ticker"" range, within the data table. Update aborted."
          Exit Sub
          End If
       End If
    If Selection.Address = Selection.EntireRow.Address Or (Selection.Rows.Count = 1 And Selection.Columns.Count = 1) Then
       nFormulas = rTicker.Offset(-1, 1).End(xlToRight).Column - rTicker.Column
       nColOffset = 0
    Else
       nFormulas = Selection.Columns.Count
       nColOffset = Selection.Column - rTicker.Column - 1
       If nColOffset < 0 Then
          MsgBox "The first highlighted column must be to the right of the ""Ticker"" range, within the data table. Update aborted."
          Exit Sub
          End If
       End If
    
    oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    
    For iRow = 1 To nTickers
        sTicker = rTicker.Offset(iRow + nRowOffset, 0)
        If sTicker = "" Then Exit For
        Application.StatusBar = Round(100 * ((iRow - 1) / nTickers), 0) & "% Completed " & _
                                " -- now processing " & sTicker & " -- #" & iRow & " of " & nTickers
        For iCol = 1 To nFormulas
            sFormula = rTicker.Offset(-1, iCol + nColOffset)
            If sFormula = "" Then Exit For
            If UCase(sFormula) <> "X" Then
               If IsNumeric(sFormula) Then
                  If smfGetAParms(1) = "" Then s1 = RCHGetElementNumber("Source", 1)
                  s1 = smfWord(smfGetAParms(0 + sFormula), 3, ";")
                  If Left(s1, 1) = "=" Then
                     sFormula = s1
                  Else
                     sFormula = "RCHGetElementNumber(""~~~~~"", " & sFormula & ")"
                     End If
                  End If
               sFormula = Replace(sFormula, "~~~~~", sTicker)
               For i1 = 1 To 50
                   If InStr(sFormula, "~~~") = 0 Then Exit For
                   If InStr(sFormula, "~~~" & i1 & "~~~") > 0 Then
                      sFormula = Replace(sFormula, "~~~" & i1 & "~~~", rTicker.Offset(iRow + nRowOffset, iCol + nColOffset).Offset(0, -i1).Value2)
                      End If
                   Next i1
               If Left(sFormula, 2) = "==" Then
                  rTicker.Offset(iRow + nRowOffset, iCol + nColOffset) = Mid(sFormula, 2, 999)
                  rTicker.Offset(iRow + nRowOffset, iCol + nColOffset) = rTicker.Offset(iRow + nRowOffset, iCol + nColOffset).Value
               Else
                  rTicker.Offset(iRow + nRowOffset, iCol + nColOffset) = smfEvaluateTwice(sFormula)
                  End If
               End If
            Next iCol
        Next iRow
ErrorExit:
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
   End Sub
