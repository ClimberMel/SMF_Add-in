Attribute VB_Name = "modGetOptionQuote"
Option Explicit
Public Function smfGetOptionQuote(ByVal pTicker As Variant, _
                                        ByVal pPutCall As Variant, _
                                        ByVal pExpiry As Variant, _
                                        ByVal pStrike As Variant, _
                                        ByVal pItem As Variant, _
                               Optional ByVal pSource As String = "Z")
                  
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to get a single option quote item from a specified data source
    '-----------------------------------------------------------------------------------------------------------*
    ' 2011.04.08 -- Created function, meant to replace all the individual data source functions
    '-----------------------------------------------------------------------------------------------------------*
    ' > Examples of invocations to get option bid price for MMM 4/16/2011 $90 call:
    '
    '   =smfGetOptionQuote("MMM","C","4/16/2011",90,"b","Z")
    '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    Dim s1 As String
    Dim sURL As String, sTicker As String
    Dim sItem As String, iCells As Integer
        
    '------------------> Verify the pPutCall parameter
    Dim sPutCall As String
    sPutCall = Left(Trim(UCase(pPutCall)), 1)
    Select Case sPutCall
       Case "P"
       Case "C"
       Case Else
            smfGetOptionQuote = "Invalid Put/Call indicator (must be a P or C): " & pPutCall
            Exit Function
       End Select
    
    '------------------> Verify pExpiry and pStrike
    pStrike = Trim(UCase(pStrike))
    pTicker = Trim(UCase(pTicker))
    Select Case True
       Case Not (VarType(pExpiry) = vbDouble Or IsDate(pExpiry))
            smfGetOptionQuote = "Bad expiration date: " & pExpiry
            Exit Function
       Case IsNumeric(pStrike)
       Case Else
            smfGetOptionQuote = "Bad strike price: " & pStrike
            Exit Function
       End Select
    
    '------------------> Verify pSource
    Select Case UCase(pSource)
       Case "Z", "ZACKS": GoTo Source_Zacks
       Case Else
            smfGetOptionQuote = "Invalid data source: " & pSource
            Exit Function
       End Select

Source_Zacks:

    sURL = "http://www.zacks.com/research/report.php?type=grk&t=" & pTicker
    sTicker = UCase(pTicker) & Format(pExpiry, " mmmyy ") & Format(pStrike, "0.00 ") & sPutCall
       
    '------------------> Verify the pItem parameter and set the # of cells to skip
    sItem = Trim(UCase(pItem))
    Select Case sItem
       Case "U": iCells = 0     ' Last price of underlying equity
       Case "X": iCells = 0     ' Expiration date
       Case "S": iCells = -17   ' Strike price
       Case "Z": iCells = 0     ' Zacks ticker symbol
       Case "B": iCells = 1     ' Bid price
       Case "A": iCells = 2     ' Ask price
       Case "L": iCells = 3     ' Last price
       Case "C": iCells = 4     ' $ Change
       Case "H": iCells = 5     ' Daily high
       Case "G": iCells = 6     ' Daily low
       Case "V": iCells = 7     ' Volume
       Case "I": iCells = 8     ' Open Interest
       Case "6": iCells = 9     ' Implied Volatility
       Case "Y": iCells = 10    ' Theoretical Value
       Case "5": iCells = 11    ' Delta
       Case "4": iCells = 12    ' Gamma
       Case "2": iCells = 13    ' Theta
       Case "1": iCells = 14    ' Vega
       Case "3": iCells = 15    ' Rho
       Case Else
            smfGetOptionQuote = "Unrecognized item ID: " & pItem
            Exit Function
       End Select
    
    '------------------> Retrieve data item
    Select Case sItem
       Case "U": smfGetOptionQuote = RCHGetTableCell(sURL, 1, ">Last")
       Case "X": smfGetOptionQuote = pExpiry
       Case Else
            smfGetOptionQuote = RCHGetTableCell(sURL, iCells, sTicker)
       End Select
    Exit Function

ErrorExit:
    smfGetOptionQuote = "Error"

    End Function


