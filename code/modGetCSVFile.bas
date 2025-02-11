Attribute VB_Name = "modGetCSVFile"
'@Lang VBA
Public Function smfGetCSVFile(pURL As String, _
               Optional ByVal pDelimiter As String = ",", _
               Optional ByVal pDim1 As Integer = 0, _
               Optional ByVal pDim2 As Integer = 0)
                    
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to return an array for a CSV file
    '-----------------------------------------------------------------------------------------------------------*
    ' 2008.07.19 -- Created function
    '-----------------------------------------------------------------------------------------------> Version 2.0h
    ' 2009.09.28 -- Added pDelimiter parameter
    ' 2010.04.21 -- Added pDim1 and pDim2 parameters
    '-----------------------------------------------------------------------------------------------> Version 2.0k
    ' > Example of an invocation (needs to be array-entered):
    '
    '   =smfGetCSVFile("http://finviz.com/grp_export.ashx?g=industry&v=152&o=-perf52w")
    '-----------------------------------------------------------------------------------------------> Version 2.2
    ' 2023-01-21 -- requires module modGetYahooQuotes for RCHGetYahooQuotes
    '-----------------------------------------------------------------------------------------------------------*
    
    smfGetCSVFile = RCHGetYahooQuotes(pURL, "", pDelimiter:=pDelimiter, pDim1:=pDim1, pDim2:=pDim2)
    
    End Function

