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
    Dim ws As Worksheet, wb As Workbook
    Set wb = ThisWorkbook
    Set ws = w.get_or_create_worksheet("test", ThisWorkbook)
    
    ' Test isnull function
    Debug.Assert IsNull(Empty) = True
    Debug.Assert IsNull("") = True
    Debug.Assert IsNull("Hello") = False
    Debug.Assert IsNull(0) = False
    Debug.Assert Not IsNull(Array("A")) = IsNull(Array()) = True
    
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
    
    ' Test cases for the InList function
    Debug.Assert InList("A", "A;B") = True
    Debug.Assert InList("A", "AS;B") = False
    Debug.Assert InList("A", Array("A", "B")) = True
    Debug.Assert InList("A", Array("AA", "B")) = False
    Debug.Assert InList("A", clls.toCollection("A", "B")) = True
    Debug.Assert InList("A", clls.toCollection("AA", "B")) = False
    
    ' Test for IsArray
    Debug.Assert u.is_array(Array(), False) And u.is_array("", False) = False
    
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
    
    ' create names to test
    r.create_named_range "test_name1", ws.name, "A1"
    r.create_named_range "test_name2", ws.name, "A2"
    r.create_named_range "test_name3", ws.name, "A3"
    
    Debug.Assert r.name_exist("test_name1") = True And r.name_exist("test_name2") = True And r.name_exist("test_name3") = True
    
    ' Test filtering by prop_values
    Set filteredNames = u.filterObjectsOnProperty(wb.Names, prop_name:="Name", prop_values:=Array("test_name1", "test_name2"))
    Debug.Assert filteredNames.count = 2
    Exit Sub
    ' Test filtering by prop_pattern
    Set filteredNames = u.filterObjectsOnProperty(wb.Names, prop_name:="Name", prop_pattern:="_name3|_name4")
    Debug.Assert filteredNames.count = 1
    
    r.deleteNames Array("test_name1", "test_name2", "test_name3")
    
    Debug.Assert Not (r.name_exist("test_name1") = True And r.name_exist("test_name2") = True And r.name_exist("test_name3") = True)
        
    ' clean up
    w.deleteWorksheets "test_ranges", "test"
End Sub

Sub test()


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
    If IsArray(x) Then
       If UBound(x) = -1 Then
          IsNull = True
       Else
          IsNull = False
       End If
    Else
       IsNull = IsEmpty(x) Or (TypeName(x) = "String" And x = "")
    End If
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
    
    If u.IsNull(x) Then
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

' Function: InList
' This function checks if a given value is present in a list, which can be a string, collection, range, or array.
' Parameters:
'   - value: The value to search for in the list.
'   - list: The list to search within, which can be a string, collection, range, or array.
'   - sep: (Optional) The separator to use if the list is a string. Default is ";".
'
' Returns:
'   - True if the value is found in the list, False otherwise.
Function InList(value As Variant, list As Variant, Optional sep As String = ";") As Boolean
    Dim arr As Variant
    Dim col As collection
    Dim item As Variant
    Dim i As Long
    
    ' Determine the type of the list and convert it to a 1D array or collection
    Select Case TypeName(list)
        Case "String"
            ' Split the string into an array using the separator
            arr = Split(list, sep)
        Case "Collection"
            ' Convert the collection to an array
            Set col = list
            ReDim arr(1 To col.count)
            For i = 1 To col.count
                arr(i) = col(i)
            Next i
        Case "Range"
            ' Convert the range to a 1D array
            arr = Application.Transpose(list.value)
        Case "Double"
            arr = Array(list)
        Case "Integer"
            arr = Array(list)
        Case "Float"
            arr = Array(list)
        Case Else
            If IsArray(list) Then
                If a.is_2d_array(list) Then
                    ' Convert 2D array to 1D array
                    arr = a.ConvertTo1DArray(list)
                Else
                    ' Use the 1D array as is
                    arr = list
                End If
            Else
                ' Raise an error if the list is not a valid type
                Err.Raise vbObjectError + 1, "InList", "Invalid list type: " & TypeName(list)
            End If
    End Select
    
    ' Check if the value is in the array
    InList = a.ItemInArray(value, arr)
End Function

Function is_array(arr As Variant, Optional raise_error As Boolean = True, Optional source As String = "") As Boolean
    ' This function checks if the provided variable is an array.
    ' If raise_error is True and the variable is not an array, it raises an error with the specified source.
    '
    ' Parameters:
    ' arr         : The variable to check.
    ' raise_error : Optional. If True, raises an error if arr is not an array. Default is True.
    ' source      : Optional. The source of the error message if raise_error is True.
    '
    ' Returns:
    ' True if arr is an array, False otherwise.
    
    ' definition: Empty can be any type
    If IsEmpty(arr) Then
       is_array = True
       Exit Function
    End If
    
    If VBA.Information.IsArray(arr) Then
        is_array = True
    Else
        is_array = False
        If raise_error Then
            Err.Raise 1001, source, source & ": input is not an array"
        End If
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
Public Function getAttr(obj, attrName As String) As Variant
    If Not (VarType(obj) = vbObject Or u.InList(TypeName(obj), Array("Range", "Name"))) Then
       Err.Raise 1003, "getAttr", "obj is not object but vartype: " & VarType(obj)
    End If
    
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

Function filterObjectsOnProperty(objects As Variant, Optional prop_name As String = "Name", Optional prop_values As Variant, Optional prop_pattern As String = "") As collection
    ' This function filters a list of VBA objects based on a specified property.
    ' Parameters:
    '   - objects: A list of VBA objects to filter.
    '   - prop_name: The name of the property to filter on. Defaults to "Name".
    '   - prop_values: (Optional) A list of property values to filter by.
    '   - prop_pattern: (Optional) A pattern to match the property value against.
    '
    ' Returns: A collection of objects that match the specified property values or pattern.
    
    Dim result As New collection
    Dim obj As Variant
    Dim propValue As Variant
    Dim colPropValues As collection
    Dim regex As Object
    
    ' Convert prop_values to a collection if provided
    If Not IsMissing(prop_values) Then
        Set colPropValues = clls.toCollection(prop_values)
    End If
    
    ' Initialize regex if prop_pattern is provided
    If prop_pattern <> "" Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.pattern = prop_pattern
        regex.IgnoreCase = True
    End If
    
    ' Loop through each object in the list
    For Each obj In objects
        ' Get the property value
        propValue = u.getAttr(obj, prop_name)
        
        ' Check if the property value is in the provided values
        If Not IsMissing(prop_values) Then
            If clls.item_exists(propValue, colPropValues) Then
                result.Add obj
            End If
        ' Check if the property value matches the pattern
        ElseIf prop_pattern <> "" Then
            If regex.test(propValue) Then
                result.Add obj
            End If
        Else
            ' Raise an error if neither prop_values nor prop_pattern is provided
            Err.Raise vbObjectError + 1, "filterObjectsOnProperty", "Either prop_values or prop_pattern must be provided"
        End If
    Next obj
    
    ' Return the filtered collection
    Set filterObjectsOnProperty = result
End Function



