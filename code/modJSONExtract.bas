Attribute VB_Name = "modJSONExtract"
Option Explicit

Private ScriptEngine As ScriptControl

Public Sub InitScriptEngine()
    Set ScriptEngine = New ScriptControl
    ScriptEngine.Language = "JScript"
    ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
End Sub

Public Function DecodeJsonString(ByVal JsonString As String)
    Set DecodeJsonString = ScriptEngine.Eval("(" + JsonString + ")")
End Function

Public Function GetProperty(ByVal JsonObject As Object, ByVal propertyName As String) As Variant
    GetProperty = ScriptEngine.Run("getProperty", JsonObject, propertyName)
End Function

Public Function GetObjectProperty(ByVal JsonObject As Object, ByVal propertyName As String) As Object
    Set GetObjectProperty = ScriptEngine.Run("getProperty", JsonObject, propertyName)
End Function

Public Function GetKeys(ByVal JsonObject As Object) As String()
    Dim Length As Integer
    Dim KeysArray() As String
    Dim KeysObject As Object
    Dim Index As Integer
    Dim Key As Variant

    Set KeysObject = ScriptEngine.Run("getKeys", JsonObject)
    Length = GetProperty(KeysObject, "length")
    ReDim KeysArray(Length - 1)
    Index = 0
    For Each Key In KeysObject
        KeysArray(Index) = Key
        Index = Index + 1
    Next
    GetKeys = KeysArray
End Function

Function smfJSONExtractField(pJSONData As String, pFieldName As String)
    Dim a1 As Variant, i1 As Integer, oJSON As Object

    InitScriptEngine
    
    a1 = Split(pFieldName, ".")
    
    Set oJSON = DecodeJsonString(CStr(pJSONData))
    For i1 = 0 To UBound(a1) - 1
        Set oJSON = GetObjectProperty(oJSON, a1(i1))
        Next i1
        
    smfJSONExtractField = GetProperty(oJSON, a1(i1))
    If smfJSONExtractField = Empty Then smfJSONExtractField = "Not Found"

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


