' Stack class

Private cnt As Long
Private stck() As Variant
Private styp As String



Private Enum basicTypes
    sinteger = 2
    slong = 3
    ssingle = 4
    sdouble = 5
    sdate = 7
    sstring = 8
    sobject = 9
    sboolean = 11
    svariant = 12
    sdecimal = 14
    sbyte = 17
    [_First] = sinteger
    [_Last] = sbyte
    'sarray = 8192
End Enum

    
    
Public Function Count()
    Count = cnt
End Function



Public Function Dump()
    Erase stck
    cnt = 0
End Function



Public Function Push(ByRef i)
    If styp = "" Then
        Err.Raise 1101, "Stack.Add", "Stack does not initialized about type."
    End If
    Dim n As Integer
    For n = basicTypes.[_First] To basicTypes.[_Last]
        If VarType(i) = n And TypeName(i) = styp Then
            ReDim Preserve stck(cnt)
            stck(cnt) = i
            cnt = cnt + 1
            Exit Function
        End If
    Next
    
    Err.Raise 1102, "Stack.Add", "Bad type for '" & styp & "' stack: " & _
                     i & " <" & TypeName(i) & ">"
End Function



Public Function Pop()
    If cnt > 0 Then
        Pop = stck(cnt - 1)
        If cnt > 1 Then
            ReDim Preserve stck(cnt - 2)
        Else
            Erase stck
        End If
        cnt = cnt - 1
    End If
End Function



Public Function Up()
    If cnt > 0 Then
        Up = stck(cnt - 1)
    End If
End Function



Public Function SetType(ByVal typ As String)
    If typ <> "Integer" And _
        typ <> "Long" And _
        typ <> "Single" And _
        typ <> "Double" And _
        typ <> "Date" And _
        typ <> "String" And _
        typ <> "Object" And _
        typ <> "Boolean" And _
        typ <> "Variant" And _
        typ <> "Decimal" And _
        typ <> "Byte" And _
        typ <> "Stack" And _
        typ <> "Dictionary" Then
        
        Err.Raise 1100, "Stack.Settype", "Non-existent data type: " & typ
    End If
    
    If styp <> Empty Then
        Err.Raise 1100, "Stack.Settype", "It cannot be changed."
    End If
    
    styp = typ
End Function

