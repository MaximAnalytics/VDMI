'1. String conversion
'2. String properties
'3. String templating
'4. Regexp

Sub test()
    Debug.Assert str.subInStr("some string @1, @2", 1, 2) = "some string 1, 2"
    Debug.Assert Not str.IsNull("A") And str.IsNull("")
    
    Debug.Assert IsArray(str.str_to_array("A,B")) And IsArray(str.str_to_array(Array("A", "B")))
    
End Sub

' String conversion
' str_to_array function
Function str_to_array(str0, Optional delimiter As String = ",") As Variant
    If VarType(str0) = vbString Then
        str_to_array = Split(str0, delimiter)
    ElseIf VarType(str0) = (vbVariant Or vbArray) Then
        str_to_array = str0
    Else
        Err.Raise 1001, "str_to_array", "str0 is not string or variant"
    End If
End Function

Function stringToCol(str0, Optional delimiter = ",") As collection
    ' convert delimited string as collection
    Dim col0 As New collection
    arr = Split(CStr(str0), delimiter)
    For i = LBound(arr) To UBound(arr)
       If Len(arr(i)) = 0 Then
          GoTo nx_i
       End If
       col0.Add arr(i)
nx_i:
    Next i
    Set stringToCol = col0
End Function

' String properties
' str_array_len function
Function str_array_len(str0) As Long
    Dim arr0() As String
    arr0 = str_to_array(str0)
    str_array_len = UBound(arr0) - LBound(arr0) + 1
End Function

Function IsNull(s As String) As Boolean
    IsNull = IsMissing(s) Or IsEmpty(s) Or s = ""
End Function

' String templating
Function substitute_into_string(ParamArray params() As Variant) As String
    Dim str0 As String
    Dim i As Integer
    
    str0 = params(LBound(params))
    If LBound(params) = UBound(params) Then
       substitute_into_string = str0
       Exit Function
    End If
    
    ' Loop through each parameter in the ParamArray
    For i = LBound(params) + 1 To UBound(params)
        ' Replace the placeholder @i with the parameter value
        str0 = Replace(str0, "@" & i, params(i))
    Next i
    
    ' Wrap the resulting code within <code></code> tags
    substitute_into_string = str0
End Function

' better name for substitute_into_string
Function subInStr(ParamArray params() As Variant) As String
    Dim str0 As String
    Dim i As Integer
    
    str0 = params(LBound(params))
    If LBound(params) = UBound(params) Then
       subInStr = str0
       Exit Function
    End If
    
    ' Loop through each parameter in the ParamArray
    For i = LBound(params) + 1 To UBound(params)
        ' Replace the placeholder @i with the parameter value
        str0 = Replace(str0, "@" & i, params(i))
    Next i
    
    ' Wrap the resulting code within <code></code> tags
    subInStr = str0
End Function

' Regexp
Function regexp_match(s As String, pattern As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.pattern = pattern
    regexp_match = regex.test(s)
End Function

