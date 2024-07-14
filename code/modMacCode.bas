Attribute VB_Name = "modMacCode"
'@Lang VBA
Public cookieArr As Variant
Public iCookieInit As Integer

'-----------------------------------------------------------
'Retrieve Data from Internet for Macintosh
'written by Paul Dyson
'-----------------------------------------------------------
'2016.05.17 -- version1.0
'   If a login is required, the function only works for Morningstar but can be amended for other sites
'2016.05.22 - allow login for any site by placing cookie file in addin location using syntax URL_cookie.txt
'2016.05.22 - adjust for zack site http error
'-----------------------------------------------------------

Public Function RCHGetURLData1Mac(pURL As String, _
                   Optional ByVal pUseIE As Integer = 0) As String
                        
    
    Dim ScriptToRun As String
    Dim iloc As Integer
    
    'Get path to addin in bash form
'    fPath = ThisWorkbook.Path & Application.PathSeparator
'    fPath = Replace(fPath, ":", "/")
'    pos1 = InStr(fPath, "/")
'    bPath = Right(fPath, Len(fPath) - pos1 + 1)
    bPath = bashPath
    
    Select Case True
        Case pUseIE = 3: pType = "--data "
        Case Else: pType = "--get "
    
    End Select
    
    'curl throws an ssl error with http://www.zack...
    'replace http with https
    If InStr(pURL, "zack") > 0 Then
        pURL = Replace(pURL, "http", "https")
    End If
    
    If iCookieInit <> 1 Then
        cookieArr = cookieFiles
        iCookieInit = 1
    End If
    
    iloc = 999
    For i = 0 To UBound(cookieArr)
    
        If InStr(pURL, cookieArr(i, 0)) > 1 Then
            iloc = i
            Exit For
        End If
    
    Next i
    
    If iloc = 999 Then
        ScriptToRun = "do shell script " & Chr(34) & "curl -L " & pType & " '" & pURL & "'" & Chr(34)
    Else
        If InStr(pURL, "bonds") = 0 Then
            ScriptToRun = "do shell script " & Chr(34) & "curl -L -b '" & bPath & cookieArr(iloc, 1) & "' " & pType & " '" & pURL & "'" & Chr(34)
        Else
            ScriptToRun = "do shell script " & Chr(34) & "curl -L " & pType & " '" & pURL & "'" & Chr(34)
        End If
    End If
        
            
    RCHGetURLData1Mac = MacScript(ScriptToRun)
    ebitLoc = InStr(RCHGetURLData1Mac, "EBITDA")
    If InStr(RCHGetURLData1Mac, Chr(13)) = 0 Then
        RCHGetURLData1Mac = RCHGetURLData1Mac & vbCrLf
        Else
        RCHGetURLData1Mac = Replace(RCHGetURLData1Mac, Chr(13), vbCrLf)
        ebitLoc = InStr(RCHGetURLData1Mac, "EBITDA")
        'changed  & Chr(13) to vbCrLf
        RCHGetURLData1Mac = RCHGetURLData1Mac & vbCrLf
        ebitLoc = InStr(RCHGetURLData1Mac, "EBITDA")
    End If
    
    End Function

Private Function cookieFiles() As Variant

    Dim arr As Variant
    Dim retArr() As String
    Dim result As String
    Dim fPath As String
    Dim pos1 As Integer
    Dim bPath As String
    
    bPath = bashPath
    
    ScriptToRun = "do shell script " & Chr(34) & "find '" & bPath & "' -name '*_cookie.txt' -type f | awk -F/ '{print $NF}'" & Chr(34)
    result = MacScript(ScriptToRun)
    arr = Split(result, Chr(13))
    ArrSize = UBound(arr)
    ReDim retArr(ArrSize, 1)
    
    For i = 0 To ArrSize
        splitString = Split(arr(i), "_")
        retArr(i, 0) = splitString(0)
        retArr(i, 1) = arr(i)
    Next i
    
    cookieFiles = retArr

End Function

Private Function bashPath()

    Dim fPath As String
    Dim pos1 As Integer

    fPath = ThisWorkbook.Path & Application.PathSeparator
    fPath = Replace(fPath, ":", "/")
    pos1 = InStr(fPath, "/")
    bashPath = Right(fPath, Len(fPath) - pos1 + 1)

End Function


