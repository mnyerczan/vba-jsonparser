Option Explicit

Private counter As Long

Private Function UTest(Procname As String, ByVal arg1, Optional ByVal arg2 = False, Optional ByVal arg3 = False)
    counter = counter + 1
    Dim res As Variant
    On Error GoTo label
        Select Case Procname
            Case "JsonParser.JsonParserEncode"
                jsonParser.JsonEncode arg1, arg2, arg3
            Case "JsonParser.Parse"
                Set res = jsonParser.Parse(arg1, arg2)
        End Select
    On Error GoTo 0
        UTest = True
        Err.Clear
        Exit Function

label:
    On Error GoTo 0
    UTest = False
End Function



Private Sub TestCases()
    Debug.Print String(20000, vbCr)
    Dim tieApi, googleApi As String

    counter = 0
    Debug.Print "-----------------------------------------------------"
    ' --------------------------------------------------
    ' Syntax check
    ' --------------------------------------------------

    ' Correct format
    Debug.Print "Test " & counter & ": " & (UTest("JsonParserParser.JsonParserEncode", "{""a"": 12}", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParserParser.JsonParserEncode", "{""a"": [21]}", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": 21}", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [.23], ""b"": 21}", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [-23], ""b"": 21}", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [+23], ""b"": 21}", True) = True)

    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": .21}", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": .21, ""c"" : true}", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": -21}", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": -21, ""c"" : true}", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": +21}", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": +21, ""c"" : true}", True) = True)

    ' key / value in array
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "[""a"": 12]", True) = False)
    ' Invalid string
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{a"": 12}", True) = False)
    ' Invalid literal
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": a}", True) = False)
    ' Invalid element in array(1)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [,]}", True) = False)
    ' Lone comma in array(1)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [12,]}", True) = False)
    ' Lone comma in object(1) end
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": 21,}", True) = False)
    ' Invalid key/value in object(2)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": {12}}", True) = False)
    ' Missing separator in the object(2)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": {""12""}}", True) = False)
    ' Missing value in the object(2)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": {""12"" :}}", True) = False)
    ' Lone comma in array(1)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [1, 2, 3, ], ""b"": {""12"" : 33}}", True) = False)
    ' Invalid character in array(1)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [1, 2, 3 @], ""b"": {""12"" : 33}}", True) = False)
    ' Ascii konvert
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [1, 2, 3 @], ""b"": {""12"" : 33}}", True, True) = False)
    ' Forbidden character
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": b, ""a"" : [1, 2, 3], ""b"": {""12"" : 33}}", True) = False)

    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": .b}", True) = False)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": -h}", True) = False)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [], ""b"": +r}", True) = False)

    ' Array
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.Parse", "[1, {""a"":1}]", True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.Parse", "[1, {"""":1}]", True) = False)


    ' Lone comma outside JsonParser.
    ' In this case, the algorithm does not throw an error,
    ' it just doesn't make it into the result
    ' invalid characters.
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": []},ads", True) = True)

    ' --------------------------------------------------
    ' None syntax check
    ' --------------------------------------------------
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [1, 2, 3 @], ""b"": {""12"" : 33}}", False) = True)
    ' Ascii konvertálás
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", "{""a"": [1, 2, 3@], ""b"": {""12"" : 33}}", False, True) = True)

    ' --------------------------------------------------
    ' Google api test
    ' --------------------------------------------------
    googleApi = http.HGet("https://adexperiencereport.googleapis.com/$discovery/rest?version=v1")

    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.JsonParserEncode", googleApi, True) = True)
    Debug.Print "Test " & counter & ": " & (UTest("JsonParser.Parse", googleApi, True) = True)

    Debug.Print "-----------------------------------------------------"
End Sub
