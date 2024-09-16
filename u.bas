' Utilities
'is_empty_missing(x As Variant) => Checks if x is empty, missing, nothing, or an empty string.
'isnull(x As Variant) => Determines if x is null, defined as an empty variant or an empty string.
'nvl(x As Variant, y As Variant) => Returns y if x is null; otherwise, returns x.
'ifFunc(b As Boolean, x As Variant, y As Variant) => Returns x if b is True; otherwise, returns y.
'mask(b As Boolean, x As Variant) => Returns x if b is True; otherwise, returns an empty variant.
'remove_dollar_sign(s) => Returns a string with all dollar signs removed from s.
'printTemplateString(params() As Variant) => Prints a string with placeholders replaced by parameter values.
'printTypename(x) => Prints the data type name of x.
' 2. Object attributes: getAttr, hasAttr, setAttr
' 3. Objectlist functions

Sub test_utilities()
    ' Test isnull function
    Debug.Assert IsNull(Empty) = True
    Debug.Assert IsNull("") = True
    Debug.Assert IsNull("Hello") = False
    Debug.Assert IsNull(0) = False
    
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
    
    ' Test Object attribute functions
    Dim ws0 As Worksheet, rng0 As Range
    Set ws0 = w.get_or_create_worksheet("test_ranges", ThisWorkbook)
    Set rng0 = ws0.Cells(1, 1)
    Debug.Assert u.hasAttr(rng0, "value")
    Call u.setAttr(rng0, "value", 1)
    Debug.Assert u.getAttr(rng0, "value") = 1
    
    ' Test Objectlist functions
    Dim wsNames As collection
    Set wsNames = u.GetObjectPropertyList(Worksheets, "name")
    Debug.Assert clls.item_exists("test_ranges", wsNames)
    w.delete_worksheet "test_ranges", ThisWorkbook
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

Function IsNull(x As Variant) As Boolean
    ' This function checks if the input variant is null.
    ' Null is defined as being an empty variant or an empty string.
    '
    ' Parameters:
    ' x : The variant to check for null.
    '
    ' Returns:
    ' True if x is null, False otherwise.
    
    IsNull = IsEmpty(x) Or (TypeName(x) = "String" And x = "")
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
    
    If IsNull(x) Then
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

' 2 Object attributes: getAttr, hasAttr, setAttr
Public Function getAttr(obj As Object, attrName As String) As Variant
    On Error GoTo ErrHandler
    getAttr = CallByName(obj, attrName, VbGet)
    Exit Function
ErrHandler:
    getAttr = CVErr(xlErrValue)
End Function

Public Function hasAttr(obj As Object, attrName As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = CallByName(obj, attrName, VbGet)
    hasAttr = (Err.Number = 0)
    Err.clear
End Function

Public Function setAttr(obj As Object, attrName As String, value As Variant) As Boolean
    On Error GoTo ErrHandler
    CallByName obj, attrName, VbLet, value
    setAttr = True
    Exit Function
ErrHandler:
    setAttr = False
End Function

' 3. Objectlist functions
Public Function Apply(objectList, func As String, ParamArray args() As Variant) As collection
    Dim result As New collection
    Dim item As Variant
    Dim output As Variant
    Dim i As Integer
    Dim params() As Variant
    
    ' Loop through each item in the collection
    For Each item In objectList
        ' Build the parameter array dynamically
        If UBound(args) > -1 Then
            ReDim params(0 To UBound(args))
            For i = LBound(args) To UBound(args)
                params(i) = args(i)
            Next i
        Else
            ReDim params(0 To 0)
        End If
        
        ' Apply the function to the item using Application.Run
        output = Application.Run(func, item, "value")
        ' Add the result to the result collection
        result.Add output
    Next item
    
    ' Return the result collection
    Set Apply = result
End Function

Public Function GetObjectPropertyList(objectList, propName As String) As collection
    Dim result As New collection
    Dim item As Variant
    Dim output As Variant
    Dim i As Integer
    Dim func As String

    ' Loop through each item in the collection
    For Each item In objectList
        ' Apply the function to the item using Application.Run
        func = "getAttr"
        output = Application.Run(func, item, propName)
        ' Add the result to the result collection
        result.Add output
    Next item
    
    ' Return the result collection
    Set GetObjectPropertyList = result
End Function

Sub test()
Dim col As New collection
Dim params() As Variant
ReDim params(0 To 0)
params(0) = col
End Sub

