'   Json converter/Parser module.
'
' Author
'   Mnyerczán Sándor
'   <mnyerczan@outlook.hu>
'
' Standard
'   The JavaScript Object Notation (JSON) Data Interchange Format
'   https://tools.ietf.org/html/rfc7158
'
'
'
Option Compare Binary
Option Explicit



Private p As Long               ' Counter
Private Token As Variant
Private translator  As Object
Private sstck As IStack         ' Structural level Stack
Private ga As Boolean           ' Grammatical analisis
Private oa As Long              ' Only ASCII code page



' These are the six structural characters:
'
' begin-array     = ws %x5B ws  ; [ left square bracket
' begin-object    = ws %x7B ws  ; { left curly bracket
' end-array       = ws %x5D ws  ; ] right square bracket
' end-object      = ws %x7D ws  ; } right curly bracket
' name-separator  = ws %x3A ws  ; : colon
' value-separator = ws %x2C ws  ; , comma
Private Enum sc
    leftSquareBracket = &H5B
    leftCurlyBracket = &H7B
    rightSquareBracket = &H5D
    rightCurlyBracket = &H7D
    colon = &H3A
    comma = &H2C
End Enum



' Insignificant whitespace is allowed before or after any of the six
' structural characters.
'   ws = *(
'           %x20 /              ; Space
'           %x09 /              ; Horizontal tab
'           %x0A /              ; Line feed or New line
'           %x0D )              ; Carriage return
Private Enum ws
    spac_e = &H20
    horizontalTab = &H9
    lineFeed = &HA
    carriageReturn = &HD
End Enum



' Convert json data to valid vba Dictionary/Array structure.
'
Public Function Parse(ByVal json As String, Optional ByVal typeReset As Boolean = False) As Variant

    json = JsonEncode(json, True)
    Debug.Print json

    p = 2
    Token = tokenize(json)
    'On Error GoTo ErrorHandler

    If Token(1) = "{" Then
        Set Parse = ParseObj
    ElseIf Token(1) = "[" Then
        Parse = ParseArr
    Else
        err.Raise 1011
    End If
    Token = Null

    If typeReset Then
        If VarType(Parse) = vbObject Then
            Set Parse = Reset(Parse)
        Else
            Parse = Reset(Parse)
        End If
    End If

    Exit Function

ErrorHandler:
    err.Raise 1011, "JsonParser.Parse", "Invalid Json format."
End Function



' Json parsing and coding algorithm. If grammatical analysis
' enabled and the systax is wrong, error will raise
' with explainer message.
'
Public Function JsonEncode( _
        ByVal json As String, _
        Optional ByVal grammaticalAnalysis As Boolean = False, _
        Optional ByVal onlyAscii As Boolean = False) As String

    If Len(json) < 2 Then err.Raise 1020, "json.JsonEncode", _
                            "Empty string or opened structure is not parsable!"

    If Mid(json, 1, 1) <> "{" And Mid(json, 1, 1) <> "[" Then
        err.Raise 1020, "json.JsonEncode", _
              "Syntax error. Expected '" & Chr(sc.leftCurlyBracket) & "', or '" & _
              Chr(sc.leftSquareBracket) & "', at: 1"
    End If

    oa = IIf(onlyAscii, &H7F, &H10FFFF)
    ga = grammaticalAnalysis
    p = 0

    On Error GoTo ErrorHandler

    If ga Then
        Set sstck = New IStack
        sstck.SetType ("Long")

        Set translator = CreateObject("Scripting.dictionary")
        translator.Add Key:=sc.leftCurlyBracket, Item:=sc.rightCurlyBracket
        translator.Add Key:=sc.leftSquareBracket, Item:=sc.rightSquareBracket
        translator.Add Key:=sc.rightCurlyBracket, Item:=sc.leftCurlyBracket
        translator.Add Key:=sc.rightSquareBracket, Item:=sc.leftSquareBracket

        JsonEncode = JsonAnalysisAndEncodeEngine(json)
        Set translator = Nothing
        Set sstck = Nothing
    Else
        JsonEncode = StringEncoder("", json)
    End If

    oa = 0
    On Error GoTo -1
    Exit Function

ErrorHandler:
    Dim a, b As Long  ' after, before
    Dim s, c, e As String

    a = IIf(p + 20 > Len(json), Len(json) - p + 1, 20)
    b = IIf(p < 20, p - 1, 20)

    s = Mid(json, p - b, b)
    c = "   >>" & Mid(json, p, 1)
    e = "<<   " & Mid(json, p + 1, a)

    err.Raise err.Number, err.Source, err.Description & Chr(10) & Chr(10) & s & c & e
End Function



' Decode json encoded string
'
Public Function JsonDecode(ByVal json As String) As String
    Dim i As Long, e As Byte
    Dim s As String, cache As String, char As String, js As String

    i = 1
    e = 0
    js = json

    While i <= Len(js)
        If Mid(js, i, 1) = "\" And Mid(js, i + 1, 1) = "u" Then
            If IsNumeric(Mid(js, i + 2, 1)) Or AscW(Mid(js, i + 2, 1)) >= 97 And AscW(Mid(js, i + 2, 1)) <= 102 And _
                IsNumeric(Mid(js, i + 3, 1)) Or AscW(Mid(js, i + 2, 1)) >= 97 And AscW(Mid(js, i + 2, 1)) <= 102 And _
                IsNumeric(Mid(js, i + 4, 1)) Or AscW(Mid(js, i + 2, 1)) >= 97 And AscW(Mid(js, i + 2, 1)) <= 102 And _
                IsNumeric(Mid(js, i + 5, 1)) Or AscW(Mid(js, i + 2, 1)) >= 97 And AscW(Mid(js, i + 2, 1)) <= 102 Then
                For e = 0 To 3
                    cache = cache & Mid(js, i + 2 + e, 1)
                Next

                char = ChrW("&H" & cache)
                s = s & char

                i = i + 5
                cache = vbNullString
            End If
        ElseIf Mid(js, i, 1) = "\" And Mid(js, i + 1, 1) = Chr(&H22) Then
            s = s & "\" & Chr(&H22)
            i = i + 1
        Else
            s = s & Mid(js, i, 1)
        End If

        i = i + 1
    Wend

    JsonDecode = s
End Function



'-------------------------------------------------------------------
' Support functions
'-------------------------------------------------------------------

' word processing algorithm
'
'    string = quotation-mark *char quotation-mark
'
'    char = unescaped /
'        escape (
'            %x22 /          ; "    quotation mark  U+0022
'            %x5C /          ; \    reverse solidus U+005C
'            %x2F /          ; /    solidus         U+002F
'            %x62 /          ; b    backspace       U+0008
'            %x66 /          ; f    form feed       U+000C
'            %x6E /          ; n    line feed       U+000A
'            %x72 /          ; r    carriage return U+000D
'            %x74 /          ; t    tab             U+0009
'            %x75 4HEXDIG )  ; uXXXX                U+XXXX
'
'    escape = %x5C              ; \
'
'    quotation-mark = %x22      ; "
'
'    unescaped = %x20-21 / %x23-5B / %x5D-10FFFF
'
Private Function StringEncoder(ByVal s As String, js As String) As String
    Dim esp As Byte ' Escape sequence part
    Dim cp As Long

    If Len(js) = 0 Then
        err.Raise 1024, "json.JsonEncode", _
        "Syntax error. Empty string, at: " & p
    End If

    s = s & Chr(&H22) 'Kezdõ idézõjel hozzáadása

    Do
        p = p + 1
        cp = CLng(AscW(Mid(js, p, 1)))
        Select Case cp

            Case &H5C:                                          ' KEY: '\'
                esp = CLng(AscW(Mid(js, p + 1, 1)))
                Select Case esp
                    Case &H22, _
                        &H47, _
                        &H62, _
                        &H66, _
                        &H6E, _
                        &H72, _
                        &H74:
                        s = s & Chr(cp) & Chr(esp)

                    Case Else:
                        s = s & Chr(cp) & Chr(&H5C)

                End Select
                p = p + 1

            Case &H22                                           ' KEY: "
                If ga Then
                    s = s & Chr(cp)
                    StringEncoder = s
                    Exit Do
                Else
                    s = s & Chr(cp)
                End If

            Case ws.spac_e, _
                ws.lineFeed, _
                ws.horizontalTab, _
                ws.carriageReturn:
                s = s & Chr(cp)

            Case &H20 To &H21, _
                &H23 To &H5B, _
                &H5D To oa:                                     ' KEY: 0x20-21 / 0x23-5B / 0x5D-10FFFF

                s = s & ChrW(cp)
            Case Else
                s = s & Chr(&H5C) & Chr(&H75)
                s = s & Right(Chr(&H30) & Chr(&H30) & Chr(&H30) & StrConv(Hex(cp), vbLowerCase), 4)
        End Select

        If p = Len(js) Then Exit Do
    Loop
    StringEncoder = s
End Function



' The algorithm checks the json format. With the exception of the allowed
' control characters and white spaces, just a string, a number, and three
' literals are allowed (true, null, false).

'   JSON can represent four primitive types (strings, numbers, booleans,
'   and null) and two structured types (objects and arrays).
'
'   A string is a sequence of zero or more Unicode characters [UNICODE].
'   Note that this citation references the latest version of Unicode
'   rather than a specific release.  It is not expected that future
'   changes in the UNICODE specification will impact the syntax of JSON.
'
'   An object is an unordered collection of zero or more name/value
'   pairs, where a name is a string and a value is a string, number,
'   boolean, null, object, or array.
'
'   An array is an ordered sequence of zero or more values.
'
Private Function JsonAnalysisAndEncodeEngine(ByVal js As String) As String
    Dim cstck As IStack ' Structural, Counter Stack
    Dim cp As Long ' Code point
    Dim s As String

    If Len(js) = p + 1 Then err.Raise 1036, _
        "JsonEncode", "Syntax error. Expected structural character: '" & _
        Chr(translator(sstck.Up)) & "'"

    Set cstck = New IStack
    cstck.SetType ("Long")

    Do:
        p = p + 1

        cp = CLng(AscW(Mid(js, p, 1)))
        Select Case cp

            ' STRUCTURAL CHARACTERS
            Case sc.leftCurlyBracket:                           ' KEY: "{"
                PlaceChk cstck, cp
                sstck.Push (cp)
                cstck.Push (cp)
                s = s & Chr(cp) & JsonAnalysisAndEncodeEngine(js) & Chr(translator(cp))
                sstck.Pop


            Case sc.leftSquareBracket:                          ' KEY: "["
                PlaceChk cstck, cp
                sstck.Push (cp)
                cstck.Push (cp)
                s = s & Chr(cp) & JsonAnalysisAndEncodeEngine(js) & Chr(translator(cp))
                sstck.Pop


            Case sc.rightCurlyBracket:                          ' KEY: "}"
                StackChk (cp)
                ObjectChk cstck
                Exit Do


            Case sc.rightSquareBracket:                         ' KEY: "]"
                StackChk (cp)
                ArrayChk cstck
                Exit Do

            Case sc.comma:                                      ' KEY: ","
                PlaceChk cstck, cp
                s = s + Chr(cp)
                cstck.Push (sc.comma)

            Case sc.colon:                                      ' KEY: ":"
                PlaceChk cstck, cp
                s = s + Chr(cp)
                cstck.Push (sc.colon)


            ' INSIGNIFICANT WHITESPACES
            Case ws.spac_e                                      ' KEY: " "
                s = s & ChrW(cp)

            Case ws.horizontalTab                               ' KEY: "\t"
                s = s & ChrW(cp)

            Case ws.lineFeed                                    ' KEY: "\n"
                s = s & ChrW(cp)

            Case ws.carriageReturn                               ' KEY: "\r"
                s = s & ChrW(cp)


            ' LITERAL NAMES
            Case &H74                                           ' KEY: "true"
                If Asc(Mid(js, p + 1, 1)) <> &H72 Or _
                    Asc(Mid(js, p + 2, 1)) <> &H75 Or _
                    Asc(Mid(js, p + 3, 1)) <> &H65 Then
                    err.Raise 1022, "json.JsonEncode", _
                        "Syntax error. Invalid literal, at: " & p
                End If
                PlaceChk cstck, cp
                s = s & "true"
                p = p + 3
                cstck.Push (cp)

            Case &H66                                           ' KEY: "false"
                If Asc(Mid(js, p + 1, 1)) <> &H61 Or _
                    Asc(Mid(js, p + 2, 1)) <> &H6C Or _
                    Asc(Mid(js, p + 3, 1)) <> &H73 Or _
                    Asc(Mid(js, p + 4, 1)) <> &H65 Then
                    err.Raise 1022, "json.JsonEncode", _
                        "Syntax error. Invalid literal, at: " & p
                End If
                PlaceChk cstck, cp
                s = s & "false"
                p = p + 4
                cstck.Push (cp)

            Case &H6E                                           ' KEY: "null"
                If Asc(Mid(js, p + 1, 1)) <> &H75 Or _
                    Asc(Mid(js, p + 2, 1)) <> &H6C Or _
                    Asc(Mid(js, p + 3, 1)) <> &H6C Then
                    err.Raise 1022, "json.JsonEncode", _
                        "Syntax error. Invalid literal, at: " & p
                End If
                PlaceChk cstck, cp
                s = s & "null"
                p = p + 3
                cstck.Push (cp)


            ' STRING
            Case &H22:                                          ' KEY: '"'
                PlaceChk cstck, cp
                cstck.Push (cp)
                s = StringEncoder(s, js)


            ' NUMBER
            Case &H30, _
                &H31, _
                &H32, _
                &H33, _
                &H34, _
                &H35, _
                &H36, _
                &H37, _
                &H38, _
                &H39, _
                &H2D, _
                &H2B, _
                &H2E, _
                &H45, _
                &H65:                                           ' KEY:  "0", "1", "2", "3", "4",
                                                                '       "5", "6", "7", "8", "9",
                                                                '       "-", "+", ".", "e", "E"

                PlaceChk cstck, cp
                s = s + Chr(cp)
                ' Save, if it does not already exist: "0"
                If cstck.Up <> &H30 Then cstck.Push (CLng(&H30))

            Case Else:                                          ' KEY: Other forbidden
                err.Raise 1023, "json.JsonEncode", _
                    "Syntax error, forbidden character, at: " & _
                    p & Chr(10) & "Code point:  0x" & Right(&H30 & &H30 & &H30 & Hex(cp), 4) & _
                    Chr(10) & "Character: " & Chr(cp)
        End Select

        If sstck.Count <> 0 Then
            If Len(js) = p Then err.Raise 1024, "json.JsonEncode", _
                    "Syntax error. Missing '" & _
                    Chr(translator(sstck.Up)) & "', at: " & p
        Else
            Exit Do
        End If
    Loop

    JsonAnalysisAndEncodeEngine = s
End Function



' post-process control
'
Private Function ObjectChk(cstck As IStack)
    Select Case cstck.Count Mod 4
        Case 0:
            If cstck.Count > 0 Then err.Raise 1026, "json.JsonEncode", _
                "Syntax error. Unexpected separator '" & Chr(sc.comma) & "', at: " & p
        Case 1:
            err.Raise 1025, "json.JsonEncode", _
                "Syntax error. Expected separator '" & Chr(sc.colon) & "', at: " & p
        Case 2:
            err.Raise 1027, "json.JsonEncode", _
                "Syntax error. Key without value, at: " & p
        Case 3:
            ' Everything is alright.
    End Select
End Function


Private Function ArrayChk(cstck As IStack)
    If cstck.Count > 0 And cstck.Count Mod 2 <> 1 Then
        err.Raise 1023, "json.JsonEncode", _
            "Syntax error. To mutch separator in array, at: " & p
    End If
End Function



' in-process control
'
Private Function PlaceChk(ByVal cstck As IStack, cp)
    ' Array
    If sstck.Up = sc.leftSquareBracket Then
        Select Case cstck.Count Mod 2
            Case 0:
                If cp = sc.comma Then
                    err.Raise 1026, "json.JsonEncode", _
                        "Syntax error. Unexpected separator '" & Chr(cp) & "', at: " & p
                ElseIf cp = sc.colon Then
                    err.Raise 1026, "json.JsonEncode", _
                        "Syntax error. Forbidden separator '" & Chr(cp) & "', at: " & p
                End If
            Case 1:
                If cp <> sc.comma And cstck.Up <> &H30 Then
                    err.Raise 1025, "json.JsonEncode", _
                        "Syntax error. Expected separator '" & Chr(sc.comma) & _
                        "', getted: '" & Chr(cp) & "', at: " & p
                End If
        End Select
    ' Object
    ElseIf sstck.Up = sc.leftCurlyBracket Then
        Select Case cstck.Count Mod 4
            Case 0:
                If cp <> &H22 Then err.Raise 1029, "json.JsonEncode", _
                    "Syntax error. Only string can be key of object, at: " & p
            Case 1:
                If cp <> sc.colon Then err.Raise 1025, "json.JsonEncode", _
                    "Syntax error. Expected separator '" & Chr(sc.colon) & _
                    "', getted: '" & Chr(cp) & "', at: " & p
            Case 2:
                If cp = sc.colon Or cp = sc.comma Then err.Raise 1025, "json.JsonEncode", _
                        "Syntax error. Unexpected token '" & Chr(cp) & "', at: " & p
            Case 3:
                If cp <> sc.comma And cstck.Up <> &H30 Then err.Raise 1025, "json.JsonEncode", _
                    "Syntax error. Expected separator '" & Chr(sc.comma) & _
                    "', getted: '" & Chr(cp) & "', at: " & p
        End Select
    End If
End Function



Private Function StackChk(cp)
    If translator(cp) <> sstck.Up Then
        If sstck.Up Then
            err.Raise 1022, "json.JsonEncode", _
                "Syntax error. Expected structural character '" & _
                    Chr(translator(sstck.Up)) & "', at: " & p
        Else
            err.Raise 1022, "json.JsonEncode", _
                "Syntax error. Expected structural characters '{', or '[', at: " & p
        End If
    End If
End Function



' The above algorithms processes the resulting json text and converts it
' into a suitable vba dictionary and / or array, thus obtaining an
' iterable, well-managed data structure during programming and data use.
Private Function ParseObj() As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.dictionary")
    Dim e As Integer
    Do:
        Select Case Token(p)

            Case "}":
                        Exit Do
            Case ",", ":":
                        ' do nothing
            Case Else:
                        If Token(p + 2) = "[" Then      ' Add dictionary
                            e = p
                            p = p + 3
                            dict.Add Key:=Token(e), Item:=ParseArr()

                        ElseIf Token(p + 2) = "{" Then  ' Add array
                            e = p
                            p = p + 3
                            dict.Add Key:=Token(e), Item:=ParseObj()

                        Else
                            dict.Add Key:=Token(p), Item:=Token(p + 2)
                            p = p + 2
                        End If
        End Select
        p = p + 1
    Loop
    Set ParseObj = dict
End Function



Private Function ParseArr() As Variant
    Dim arr() As Variant
    Dim e As Integer
    e = 0
    Do:
        Select Case Token(p)
            Case "}":
                        ' do nothing
            Case "{":
                        ReDim Preserve arr(e)
                        p = p + 1
                        Set arr(e) = ParseObj

            Case "[":
                        ReDim Preserve arr(e)
                        p = p + 1
                        arr(e) = ParseArr()

            Case "]":
                        Exit Do
            Case ",":
                        e = e + 1
            Case Else:
                        ReDim Preserve arr(e)
                        arr(e) = Token(p)
        End Select
        p = p + 1
    Loop

    ParseArr = arr
End Function



Private Function tokenize(s)
    Dim pattern As String
    
    pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[ee][+\-]?\d+)?|\w+|[^\s""']+?"
    tokenize = Rextract(s, pattern, True)
End Function



Private Function Rextract(s, pattern, Optional bGroup1bias As Boolean, Optional bGlobal As Boolean = True)
    Dim c&, m, n, v
    With CreateObject("vbscript.regexp")
        .Global = bGlobal
        .MultiLine = False
        .Ignorecase = True
        .pattern = pattern

        If .test(s) Then
            Set m = .Execute(s)
            ReDim v(1 To m.Count)
            For Each n In m
                c = c + 1
                v(c) = n.Value
                If bGroup1bias Then
                    If Len(n.submatches(0)) Or n.Value = """""" Then
                        v(c) = n.submatches(0)
            Next
        End If
    End With

    Rextract = v
End Function

' The algorithm recursively traverses the structure of
' dictionaries / arrays and modifies string literals to
' the appropriate data structure.
Private Function Reset(jObj As Variant) As Variant

    ' Dictionary
    If VarType(jObj) = vbObject Then
        Dim k As Variant
        For Each k In jObj.Keys()
            vSwitcher jObj, k
        Next k
        Set Reset = jObj
        Exit Function

    ' Variant()
    ElseIf VarType(jObj) = vbArray + vbVariant Then
        Dim i As Long
        If ArrUBound(jObj) > 0 Then
            For i = 0 To UBound(jObj)
                vSwitcher jObj, i
            Next
        End If
    Else
        If IsNumeric(jObj) Then
            ' A kis számábrázolás miatt
            ' minden számot Decimal-ra konvertál
            '
            ' Integer:  2 bytes
            ' Long:     4 bytes
            ' Decimal:  12 bytes  <-
            If 48 < AscW(Left(jObj, 1)) And AscW(Left(jObj, 1)) <= 57 Then
                jObj = CDec(jObj)
            End If
        ElseIf jObj = "true" Then
            jObj = True
        ElseIf jObj = "false" Then
            jObj = False
        ElseIf jObj = "null" Then
            jObj = Null
        End If
    End If

    Reset = jObj
End Function



Public Function ArrUBound(arr As Variant) As Long
    ArrUBound = 0
    On Error Resume Next
        ArrUBound = UBound(arr) + 1
    On Error GoTo -1
End Function



' Because variant type, needed a switcher beetwen
' object and array definition.
'
Private Function vSwitcher(ByRef jObj As Variant, ByVal k As String)
    If VarType(jObj(k)) = vbObject Then
        Set jObj(k) = Reset(jObj(k))
    Else
        jObj(k) = Reset(jObj(k))
    End If
End Function
