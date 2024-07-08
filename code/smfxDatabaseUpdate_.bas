Attribute VB_Name = "smfxDatabaseUpdate_"
'@Lang VBA
Sub UpdateStockDatabase()
    '-----------------------------------------------------------------------------------------------------------*
    ' Subroutine to update a number of stock databases, one sheet per data source
    '-----------------------------------------------------------------------------------------------------------*
    ' 2006.03.15 -- Created
    '-----------------------------------------------------------------------------------------------> Version 1.2
    ' 2007.10.03 -- Removed Business Week and Telescan sheetname options
    '-----------------------------------------------------------------------------------------------> Version 2.0i
    For Each oSheet In ActiveWorkbook.Sheets
        
        Dim iElement As Integer
        Dim sSymbol As String
        
        sVersion = RCHGetElementNumber("Version")    ' Initialize the list of available elements
        
        For iElement = 1 To kElements
            Select Case True
               Case oSheet.Name = RCHGetElementNumber("Source", iElement): Exit For
               Case iElement = kElements: GoTo Next_WorkSheet
               End Select
            Next iElement
        
        iTicker = 2                                    ' Set initial ticker pointer
        Do While True
           
           iTicker = iTicker + 1                       ' Go to next ticker symbol in list
           sSymbol = oSheet.Cells(iTicker, 1)          ' Get ticker symbol of company
           If sSymbol = "" Then GoTo Next_WorkSheet    ' No more ticker symbols
           
           nDate = oSheet.Cells(iTicker, 2)            ' Get date of last update for company
           If nDate <> 0 Then GoTo Next_Company        ' Valid date, no need to update
           oSheet.Cells(iTicker, 2) = Date             ' Update the last update date
           
           iElement = 0                                ' Set initial element pointer
           iColumn = 2                                 ' Set initial column pointer
           iSheet = 1                                  ' Set sheet pointer for 256+ element sources
           Set oUpdate = oSheet
           
           Do While True
              iElement = iElement + 1                             ' Go to next available element
              sSource = RCHGetElementNumber("Source", iElement)   ' Get data source of element
              If sSource = "EOL" Then GoTo Next_Company
              If sSource <> oSheet.Name Then GoTo Next_Element    ' Not an applicable element for worksheet
              iColumn = iColumn + 1                               ' Go to next output column
              If oUpdate.Cells(2, iColumn) = "" Then
                 oUpdate.Cells(1, iColumn) = iElement
                 oUpdate.Cells(2, iColumn) = RCHGetElementNumber("Element", iElement)
                 End If
              Application.StatusBar = "Now updating ticker " & sSymbol & " on worksheet " & oUpdate.Name
              oUpdate.Cells(iTicker, iColumn) = RCHGetElementNumber(sSymbol, iElement)
              If iColumn = 256 Then
                 iSheet = iSheet + 1
                 Set oUpdate = ActiveWorkbook.Sheets(sSource & "_" & iSheet)
                 oUpdate.Cells(iTicker, 1) = oSheet.Cells(iTicker, 1)
                 oUpdate.Cells(iTicker, 2) = oSheet.Cells(iTicker, 2)
                 iColumn = 2
                 End If
              'Call TickerReset
Next_Element: Loop

Next_Company: Loop

Next_WorkSheet: Next oSheet
    
    Application.StatusBar = False
    
    End Sub
    
    
