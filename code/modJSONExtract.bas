Attribute VB_Name = "modJSONExtract"
'@Lang VBA
Option Explicit

'-----------------------------------------------------------------------------------
' 08/08/23 - ScriptEngine/JSON extract code copied from
'            "https://stackoverflow.com/questions/6627652/parsing-json-in-excel-vba"
'-----------------------------------------------------------------------------------
'            -- only runs in Excel x32 --
'-----------------------------------------------------------------------------------

Private ScriptEngine As ScriptControl

Public Sub InitScriptEngine()
    Set ScriptEngine = New ScriptControl
    ScriptEngine.Language = "JScript"
    ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
End Sub

Public Function DecodeJsonString(ByVal JsonString As String)
    Set DecodeJsonString = ScriptEngine.eval("(" + JsonString + ")")
End Function

Public Function GetProperty(ByVal jsonObject As Object, ByVal propertyName As String) As Variant
    GetProperty = ScriptEngine.Run("getProperty", jsonObject, propertyName)
End Function

Public Function GetObjectProperty(ByVal jsonObject As Object, ByVal propertyName As String) As Object
    Set GetObjectProperty = ScriptEngine.Run("getProperty", jsonObject, propertyName)
End Function

Public Function GetKeys(ByVal jsonObject As Object) As String()
    Dim Length As Integer
    Dim KeysArray() As String
    Dim KeysObject As Object
    Dim Index As Integer
    Dim Key As Variant

    Set KeysObject = ScriptEngine.Run("getKeys", jsonObject)
    Length = GetProperty(KeysObject, "length")
    ReDim KeysArray(Length - 1)
    Index = 0
    For Each Key In KeysObject
        KeysArray(Index) = Key
        Index = Index + 1
    Next
    GetKeys = KeysArray
End Function

Function smfJSONExtractField(pJSONData As String, pJSONFields As String)
    '-----------------------------------------------------------------------------------
    ' 08/28/23 - Function runs in Excel x32 only.
    '
    '-----------------------------------------------------------------------------------
    Dim a1 As Variant, i1 As Integer, oJSON As Object

    InitScriptEngine
    
    a1 = Split(pJSONFields, ".")
    
    Set oJSON = DecodeJsonString(CStr(pJSONData))
    For i1 = 0 To UBound(a1) - 1
        Set oJSON = GetObjectProperty(oJSON, a1(i1))
        Next i1
        
    smfJSONExtractField = GetProperty(oJSON, a1(i1))
    If smfJSONExtractField = Empty Then smfJSONExtractField = "Not Found"

End Function
        
Function smfJSONExtractField_x64(pJSONData As String, pFieldNames As String)
    '-----------------------------------------------------------------------------------
    ' 08/28/23 - Created for Excel x64
    '            Uses "JSONConverter" to parse thru JSON string.
    '-----------------------------------------------------------------------------------

    Dim aFieldNames As Variant, i1 As Integer, oJSON As Object
    
    Dim vJSONKey0 As Variant, vJSONKey1 As Variant, vJSONKey2 As Variant, vJSONKey3 As Variant, vJSONKey4 As Variant
    Dim vJSONKey5 As Variant, vJSONKey6 As Variant, vJSONKey7 As Variant, vJSONKey8 As Variant, vJSONKey9 As Variant
    Dim vJSONFieldValue As Variant
    
    aFieldNames = Split(pFieldNames, ".")
    
    Set oJSON = JsonConverter.ParseJson(pJSONData)
       
    For i1 = 0 To UBound(aFieldNames)
        Select Case i1
            Case 0: vJSONKey0 = smfChkJSONFieldType(aFieldNames(i1))
            Case 1: vJSONKey1 = smfChkJSONFieldType(aFieldNames(i1))
            Case 2: vJSONKey2 = smfChkJSONFieldType(aFieldNames(i1))
            Case 3: vJSONKey3 = smfChkJSONFieldType(aFieldNames(i1))
            Case 4: vJSONKey4 = smfChkJSONFieldType(aFieldNames(i1))
            Case 5: vJSONKey5 = smfChkJSONFieldType(aFieldNames(i1))
            Case 6: vJSONKey6 = smfChkJSONFieldType(aFieldNames(i1))
            Case 7: vJSONKey7 = smfChkJSONFieldType(aFieldNames(i1))
            Case 8: vJSONKey8 = smfChkJSONFieldType(aFieldNames(i1))
            Case 9: vJSONKey9 = smfChkJSONFieldType(aFieldNames(i1))
        End Select
    Next
            
    Select Case i1
        Case 10: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)(vJSONKey5)(vJSONKey6)(vJSONKey7)(vJSONKey8)(vJSONKey9)
        Case 9: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)(vJSONKey5)(vJSONKey6)(vJSONKey7)(vJSONKey8)
        Case 8: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)(vJSONKey5)(vJSONKey6)(vJSONKey7)
        Case 7: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)(vJSONKey5)(vJSONKey6)
        Case 6: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)(vJSONKey5)
        Case 5: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)
        Case 4: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)
        Case Else
            vJSONFieldValue = "JSON key field (" & i1 + 1 & ") out of range"
    End Select
    
    If vJSONFieldValue = Empty Then vJSONFieldValue = "Not Found"
    
    smfJSONExtractField_x64 = vJSONFieldValue

End Function
    
Function smfChkJSONFieldType(pField)
    '-----------------------------------------------------------------------------------
    ' 08/28/23 - Created for Excel x64
    '            "JSONConverter" arrays start at 1 not 0.
    '            "Numeric" key fields also need to be numeric type, not 'variant'
    '-----------------------------------------------------------------------------------
    Dim vField As Variant
    
    If IsNumeric(pField) Then
        vField = pField + 1             ' Also converts to numeric type
    Else
        vField = pField
    End If
           
    smfChkJSONFieldType = vField
    
End Function

Function smfJSONExtractKeys(pJSONData As String, pFieldName As String)
    Dim a1 As Variant, i1 As Integer, o1 As Object, oJSON As Object, v1 As Variant, vReturn As Variant

    InitScriptEngine
    
    a1 = Split(pFieldName, ".")
    
    Set oJSON = DecodeJsonString(CStr(pJSONData))
    For i1 = 0 To UBound(a1)
        Set oJSON = GetObjectProperty(oJSON, a1(i1))
        Next i1
        
    v1 = GetKeys(oJSON)
    ReDim vReturn(1 To UBound(v1) + 2, 1 To 1)
    For i1 = 0 To UBound(v1)
        vReturn(i1 + 1, 1) = v1(i1)
        Next i1
    vReturn(i1 + 1, 1) = "#N/A"
    smfJSONExtractKeys = vReturn
    
End Function


