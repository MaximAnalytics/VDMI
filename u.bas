' Utilities
'is_empty_missing(x As Variant) => Checks if x is empty, missing, nothing, or an empty string.
'isnull(x As Variant) => Determines if x is null, defined as an empty variant or an empty string.
'nvl(x As Variant, y As Variant) => Returns y if x is null; otherwise, returns x.
'ifFunc(b As Boolean, x As Variant, y As Variant) => Returns x if b is True; otherwise, returns y.
'mask(b As Boolean, x As Variant) => Returns x if b is True; otherwise, returns an empty variant.
'remove_dollar_sign(s) => Returns a string with all dollar signs removed from s.
'printTemplateString(params() As Variant) => Prints a string with placeholders replaced by parameter values.
'printTypename(x) => Prints the data type name of x.

Sub test_utilities()
    ' Test isnull function
    Debug.Assert isnull(Empty) = True
    Debug.Assert isnull("") = True
    Debug.Assert isnull("Hello") = False
    Debug.Assert isnull(0) = False
    
    ' Test nvl function
    Debug.Assert nvl(Empty, "Default") = "Default"
    Debug.Assert nvl("", "Default") = "Default"
    Debug.Assert nvl("Hello", "Default") = "Hello"
    Debug.Assert nvl(0, 1) = 0
    
    ' Test ifFunc function
    Debug.Assert ifFunc(True, "Yes", "No") = "Yes"
    Debug.Assert ifFunc(False, "Yes", "No") = "No"
    Debug.Assert ifFunc(True, 1, 0) = 1
    Debug.Assert ifFunc(False, 1, 0) = 0
    
    ' Test mask function
    Debug.Assert mask(True, "Visible") = "Visible"
    Debug.Assert mask(False, "Visible") = Empty
    Debug.Assert mask(True, 123) = 123
    Debug.Assert mask(False, 123) = Empty
End Sub

' Logical checks: if_empty_missing=> checks if x is empty, missing, nothing or empty string
Function is_empty_missing(x As Variant) As Boolean
    If IsEmpty(x) Or IsMissing(x) Then
        is_empty_missing = True
        Exit Function
    ElseIf TypeName(x) = "String" Then
        If Len(x) = 0 Then
            is_empty_missing = True
            Exit Function
        End If
    End If
    
    On Error Resume Next
    is_empty_missing = x Is Nothing
    Exit Function
    On Error GoTo 0
    
    is_empty_missing = False
End Function

Function isnull(x As Variant) As Boolean
    ' This function checks if the input variant is null.
    ' Null is defined as being an empty variant or an empty string.
    '
    ' Parameters:
    ' x : The variant to check for null.
    '
    ' Returns:
    ' True if x is null, False otherwise.
    
    isnull = IsEmpty(x) Or (TypeName(x) = "String" And x = "")
End Function

Function nvl(x As Variant, y As Variant) As Variant
    ' This function returns the second argument if the first argument is null.
    ' It mimics the behavior of the NVL function in SQLServer.
    '
    ' Parameters:
    ' x : The value to check for null.
    ' y : The value to return if x is null.
    '
    ' Returns:
    ' x if x is not null, y otherwise.
    
    If isnull(x) Then
        nvl = y
    Else
        nvl = x
    End If
End Function

Function ifFunc(b As Boolean, x As Variant, y As Variant) As Variant
    ' This function returns one of two values based on a boolean condition.
    '
    ' Parameters:
    ' boolean : The condition to evaluate.
    ' x       : The value to return if the condition is True.
    ' y       : The value to return if the condition is False.
    '
    ' Returns:
    ' x if boolean is True, y otherwise.
    
    If b Then
        ifFunc = x
    Else
        ifFunc = y
    End If
End Function

Function mask(b As Boolean, x As Variant) As Variant
    ' This function masks the value x based on a boolean condition.
    ' If the condition is True, it returns x; otherwise, it returns an empty variant.
    '
    ' Parameters:
    ' boolean : The condition to evaluate.
    ' x       : The value to potentially mask.
    '
    ' Returns:
    ' x if boolean is True, Empty otherwise.
    
    If b Then
        mask = x
    Else
        mask = Empty
    End If
End Function

'Formulas
Function remove_dollar_sign(s) As String
remove_dollar_sign = Replace(CStr(s), "$", "")
End Function

' Printing
Sub printTemplateString(ParamArray params() As Variant)
    Dim str0 As String
    Dim i As Integer
    
    str0 = params(LBound(params))
    If LBound(params) = UBound(params) Then
       Exit Sub
    End If
    
    ' Loop through each parameter in the ParamArray
    For i = LBound(params) + 1 To UBound(params)
        ' Replace the placeholder @i with the parameter value
        str0 = Replace(str0, "@" & i, params(i))
    Next i
    
    Debug.Print str0
End Sub

Sub printTypename(x)
    Debug.Print TypeName(x)
End Sub
