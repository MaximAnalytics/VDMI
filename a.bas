' TODO: check row, column indices using Redim
' array functions
' 1. array properties
' 2. array transformations/generators
' 3. utility
' 4. subsetting
' 5. Array combining, joining, merging
' 6. Interact with worksheet

' 1. array properties
'printArray(arr) => Prints array elements to immediate window
'array_contains(arr, item) => Checks if array contains item
'array_is_missing(arr) => Checks if array is uninitialized
'is_1d_array(arr) => Determines if array is 1-dimensional
'is_2d_array(arr) => Determines if array is 2-dimensional
'get_array_len(arr) => Gets length of 1-dimensional array
'array_length(arr) => Gets length of array (1D or 2D)
'printArrayRowsColumns(arr, array_name) => Prints array dimensions
'printArrayBounds(arr, array_name) => Prints array bounds
'num_array_rows(arr) => Returns number of rows in array
'num_array_columns(arr) => Returns number of columns in array

' 2. array transformation
'ConvertTo1DArray(arr) => Converts 2D array to 1D array
'convertTo2DArray(arr, axis) => Converts 1D array to 2D array
'to_array(x) => Converts collection or value to array
'create_vector(num_rows, default_value, start_index, header_value, as_2darray) => Creates a vector with optional header
'create_integer_vector(start, endVal) => Creates integer vector from start to end
'create_array(num_rows, num_cols, default_value) => Creates 2D array with default values
'FillArray(arr As Variant, value As Variant)

' 3. utility
'ItemInArray(item, arr) => Checks if item is in array
'indexInBounds(arr, i, axis) => Checks if index is within array bounds
'ArraysAreEqual(arr0, arr1) => Checks if two arrays are equal
'castArrayAsString(arr) => Casts all elements of array to string
'as_collection(arr, index) => Converts array to collection
'ArrayIsEmpty(arr) => returns True if array is empty (has no elements)

' 4. subsetting
'FindArrayIndex(arr, value, axis, index0) => Finds index of value in array
'FindArrayRowIndex(arr, column, value) => Finds row index by column value
'FindArrayColumnIndex(arr, column) => Finds column index by column name
'MatchArrayRowIndex(arr, key_columns, search_row, dbg) => Finds row index by multiple criteria
'getIndexOfArray(value, arr) => Finds index of value in array
'get_array_column(arr, i, row_offset) => Extracts a column from array
'get_array_row(arr, j) => Extracts a row from array
'getRowIndexes(arr) => Returns collection of row indexes
'subset_rows(arr, start_row, end_row) => Subsets array rows
'subset_columns(arr, start_column, end_column, start_row, end_row) => Subsets array columns
'resize_array(arr0, r0, r1, c0, c1) => Resizes array by subsetting
'subset_indices(arr, axis, indices) => Subsets array by indices
'select_column_names(arr, column_names) => Selects columns by names
'get_indices(arr, column_names) => Gets indices of column names
'select_array_columns(arr, column_names) => Selects array columns by names
'QueryArray(arr As Variant, ParamArray criteria()) => Filters arr on criteria() => (var1,val1,...,varN,valN)
'RemoveNullsFromArray(arr, filterColumns()) => Filters out arr on filterColumn = null

'4.2 filtering

' 5. combine, join, merge
'setArrayHeader(arr, header) => Sets header row in array
'concatArrays(arr0, arr1) => Concatenates two arrays vertically
'AppendColumn(arr0, values) => Appends column to array

' 6. Interact with worksheet
'paste_array(arr, addr, ws, wb) => Pastes array to worksheet range

Sub test_array_functions()

    Application.ScreenUpdating = False
    
    ' initialize
    Dim ws0 As Worksheet, matrix() As Integer
    Set ws0 = w.get_or_create_worksheet("test", ThisWorkbook, True)
    
    ' 2 transformation
    Debug.Assert a.toString(Array("A", "B")) = "A;B"
    
    '3 utility
    Debug.Assert a.ItemInArray("A", Array("A", "B")) And Not a.ItemInArray("C", Array("A", "B"))

    'test: create vector, array and test 1d/2d, w/o header
    vec0 = a.create_vector(10, 1, header_value:="some_number", as_2darray:=False)
    vec1 = a.create_vector(10, "A", header_value:="some_string", as_2darray:=True)
    Debug.Assert a.is_1d_array(vec0) And a.array_length(vec0) = 10
    Debug.Assert a.is_2d_array(vec1) And a.num_array_rows(vec1) = 10
    
    'test empty array
    Dim empty_array As Variant
    Debug.Assert a.num_array_rows(empty_array) = -1
    Debug.Assert a.ArrayIsEmpty(empty_array) And a.ArrayIsEmpty(matrix) And Not a.ArrayIsEmpty(vec0)
    
    ReDim matrix(2, 2)
    matrix = FillArray(matrix, 1)
    Debug.Assert Not a.ArrayIsEmpty(matrix)

    'test: array utilities (contains), array conversion
    Debug.Assert a.array_contains(vec0, 1) And a.array_contains(vec1, "A")
    
    'test: create matrix (array multiple columns), check number of columns
    arr0 = a.AppendColumn(vec1, vec0)
    Debug.Assert a.num_array_rows(arr0) = 10 And a.num_array_columns(arr0) = 2
    header_row = a.get_array_row(arr0, 1)
    Debug.Assert header_row(1, 1) = "some_string" And header_row(1, 2) = "some_number"
    
    ' Test the FillArray function with assertion tests
    ' Test with a 1D array
    Dim arr1D As Variant
    arr1D = Array(1, 2, 3, 4, 5)
    arr1D = FillArray(arr1D, 0)
    Debug.Assert is_1d_array(arr1D) And array_contains(arr1D, 0) And Not array_contains(arr1D, 1)
    
    ' Test with a 2D array
    Dim arr2D As Variant
    ReDim arr2D(1, 1)
    a.printArray arr2D
    arr2D = FillArray(arr2D, "X")
    
    Debug.Assert is_2d_array(arr2D) And array_contains(arr2D, "X") And Not array_contains(arr2D, 1)
    
    ' Test with an empty array
    Dim emptyArray As Variant
    emptyArray = Array()
    emptyArray = FillArray(emptyArray, "Empty")
    Debug.Assert array_length(emptyArray) = 0
    
    'test: select array columns, names
    column_some_number = a.select_array_columns(arr0, "some_number")
    column_some_string = a.select_array_columns(arr0, "some_string")
    vec0 = a.create_vector(10, 1, header_value:="some_number", as_2darray:=True)
    
    ' these arrays should be equal
    Debug.Assert a.ArraysAreEqual(vec0, column_some_number) And a.ArraysAreEqual(vec1, column_some_string) And a.ArraysAreEqual(empty_array, empty_array)
    column_2 = a.get_array_column(arr0, 2)
    Debug.Assert a.ArraysAreEqual(vec0, column_2)
    
    ' these shouldnt
    Debug.Assert Not (a.ArraysAreEqual(vec0, vec1) Or a.ArraysAreEqual(vec1, matrix))
    
    '4. subset rows, columns and count rows/columns and resize_rows
    arr1 = a.resize_array(arr0, 2) 'without header
    Debug.Assert a.num_array_columns(arr1) = 2 And a.num_array_rows(arr1) = 9
    arr_subset_cols = a.subset_columns(arr0, 2, 2)
    Debug.Assert a.num_array_columns(arr_subset_cols) = 1 And a.num_array_rows(arr_subset_cols) = 10
    arr_subset_rows = a.subset_rows(arr0, 1, 5)
    Debug.Assert a.num_array_columns(arr_subset_rows) = 2 And a.num_array_rows(arr_subset_rows) = 5
    
    Debug.Print getRowIndexes(Array()).count = 0
    
    ' test: FindArrayIndex
    row1 = a.convertTo2DArray(Array("B", 2), axis:=1)
    arr3 = a.concatArrays(arr0, row1)
    'a.PasteArray arr3, "A1", ws0
    Debug.Assert a.num_array_columns(arr3) = 2 And a.num_array_rows(arr3) = 10 + 1
    a.printArray arr3
    Debug.Assert a.FindArrayIndex(arr3, "some_number", axis:=1) = 2 And a.FindArrayRowIndex(arr3, "some_number", 2)
    
    'in array arr3 find row where `some_string`=="B" and some_number == "2"
    matched_row_1 = a.MatchArrayRowIndex(arr3, Array("some_string"), Array("A"))
    matched_row_2 = a.MatchArrayRowIndex(arr3, Split("some_string,some_number", ","), Array("B", 2))
    Debug.Assert matched_row_1 = 2 And matched_row_2 = 11
    
    ' append arr0, arr1
    arr2 = a.concatArrays(arr0, arr1)
    Debug.Print a.num_array_columns(arr0), a.num_array_columns(arr1), a.num_array_columns(arr2)
    Debug.Assert a.num_array_columns(arr2) = 2 And a.num_array_rows(arr2) = 19
    a.pasteArray arr2, "A1", ws0
    
    arrAppendColumn = a.AppendColumn(arr2, 2, header_value:="new_column")
    a.printArrayRowsColumns arrAppendColumn
    Debug.Assert a.num_array_columns(arrAppendColumn) = 3 And arrAppendColumn(1, 3) = "new_column"
    
    'test: getColumnValue
    a.printArray arr2
    
    ' test getColumnValue
    testdataArray = testdata.getTestDataArray()
    Debug.Assert a.getColumnValue(testdataArray, 1, "Omschrijving") = "Omschrijving"
    Debug.Assert a.getColumnValue(testdataArray, 10, "Omschrijving") = "HG Kunststglans 6xIL ES 15F(S)"
    
    ' test QueryArray
    headerArr = a.get_array_row(testdataArray, 1)
    a.printArray headerArr
    a.printArrayRowsColumns headerArr
    arrayFiltered1 = a.QueryArray(testdataArray, "Qty1", 95, "Ordernummer", "228978")
    Debug.Assert a.num_array_rows(arrayFiltered1) = 4
    
    ' 4.2 array filtering
    ' test RemoveNullsFromArray
    testdataArrayNulls = testdataArray
    Debug.Assert a.getArrayColumnIndex(testdataArray, "Artikel") = 1 And a.getArrayColumnIndex(testdataArray, "Aantal") = 5
    
    'Set 2 rows to be filtered out
    testdataArrayNulls(2, 1) = ""
    testdataArrayNulls(10, 5) = Empty
    testdataArrayNotNull = a.RemoveNullsFromArray(testdataArrayNulls, "Aantal", "Artikel", "Qty1")
    Debug.Print a.num_array_rows(testdataArray) - a.num_array_rows(testdataArrayNotNull) = 2
    
    ' test getNamedArrayValue
    arrayFiltered2 = a.QueryArray(arrayFiltered1, "Artikel", "000900853/02")
    a.printArray arrayFiltered2
    Debug.Assert getNamedArrayValue(arrayFiltered2, "Ordernummer") = 228978 And getNamedArrayValue(arrayFiltered2, "Land") = "NL"
    
    columnArray = Worksheets("test").Range("A1:A20").value
    a.printArray RemoveNullsFromArray(columnArray), True
    Debug.Print a.num_array_rows(RemoveNullsFromArray(columnArray))
    Debug.Assert a.num_array_rows(RemoveNullsFromArray(columnArray, 1)) = 19
    
    Debug.Assert a.num_array_rows(a.RemoveNullsFromVector(Array("1", 1, ""))) = 2
    
    mat1 = concatArrays(Array("X", 5550, "5550"), Array("X", "Y", "Z"), axis:=1)
    a.printArrayAsString mat1
    Debug.Assert a.num_array_rows(mat1) = 2 And a.num_array_rows(FilterArrayOnPattern(mat1, "^[\d]", 2)) = 1
    
    

    Application.ScreenUpdating = True
    
End Sub

' 1 Array properties
Sub printArray(arr, Optional as_string As Boolean = False)
    Dim i As Long
    Dim j As Long
    Dim rowItems As collection
    
    ' Check if the array is 1-dimensional or 2-dimensional
    If IsArray(arr) Then
        If as_string Then
           Debug.Print a.toString(arr, ";", "\n")
        Else
            ' 2-dimensional array
            If is_2d_array(arr) Then
                For i = LBound(arr, 1) To UBound(arr, 1)
                        Set rowItems = New collection
                        For j = LBound(arr, 2) To UBound(arr, 2)
                            rowItems.Add arr(i, j)
                        Next j
                        Debug.Print clls.collectionToString(rowItems, ",")
                Next i
            ' 1-dimensional array
            Else
                For i = LBound(arr) To UBound(arr)
                    Debug.Print arr(i)
                Next i
            End If
        End If
    Else
        ' Not an array
        Debug.Print arr
    End If
End Sub

Sub printArrayAsString(arr)
    a.printArray arr, True
End Sub

Function array_contains(arr As Variant, item As Variant) As Boolean
    Dim i As Long
    If is_2d_array(arr) Then
        ' 2-dimensional array
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
               If arr(i, j) = item Then
                  array_contains = True
                  Exit Function
               End If
            Next j
        Next i
    Else
        For i = LBound(arr) To UBound(arr)
            If arr(i) = item Then
                array_contains = True
                Exit Function
            End If
        Next i
    End If
    array_contains = False
End Function

' for checking if array is initialized with values
Function array_is_missing(arr) As Boolean
    array_is_missing = IsMissing(arr)
End Function

Function is_1d_array(arr) As Boolean
    On Error GoTo is_not_one_dimensional
    ub = UBound(arr, 2)
    is_1d_array = False
    Exit Function
is_not_one_dimensional:
    is_1d_array = True
End Function

Function is_2d_array(arr) As Boolean
    Dim isTwoDimensional As Boolean
    
    ' Check if arr is a two-dimensional array
    If IsArray(arr) Then
        On Error GoTo is_one_dimensional
        ' Check the upper bounds of both dimensions
        If UBound(arr, 1) > 0 And UBound(arr, 2) > 0 Then
            is_2d_array = True
            Exit Function
        End If
is_one_dimensional:
        is_2d_array = False
        Exit Function
        On Error GoTo 0
    Else
        Err.Raise 1001, "is_2d_array", "arr is not array but: " + TypeName(arr)
    End If
    is_2d_array = False
End Function

Function get_array_len(arr) As Long
    get_array_len = UBound(arr) - LBound(arr) + 1
End Function

Function array_length(arr) As Long
    If a.is_2d_array(arr) Then
       array_length = UBound(arr, 1) - LBound(arr, 1) + 1
    Else
       array_length = UBound(arr) - LBound(arr) + 1
    End If
End Function

'2. array transformations
Function ConvertTo1DArray(arr As Variant) As Variant
    Dim result() As Variant
    Dim numRows As Long
    Dim numCols As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    numCols = UBound(arr, 2)
    numRows = UBound(arr, 1)
    ReDim result(1 To numRows * numCols)
    
    k = 1
    If IsArray(arr) Then
        If is_2d_array(arr) Then
            For i = LBound(arr, 1) To UBound(arr, 1)
                For j = LBound(arr, 2) To UBound(arr, 2)
                    result(k) = arr(i, j)
                    k = k + 1
                Next j
            Next i
        Else
            result = arr
        End If
    End If
    
    ConvertTo1DArray = result
End Function

Function convertTo2DArray(arr As Variant, Optional axis = 0) As Variant
    If a.is_2d_array(arr) Then
       convertTo2DArray = arr
       Exit Function
    End If
    ' as column array
    If axis = 0 Then
        numRows = UBound(arr) - LBound(arr) + 1
        ReDim result(1 To numRows, 1)
        k = 1
        For i = LBound(arr) To UBound(arr)
            result(k, 1) = arr(i)
            k = k + 1
        Next i
    ' as row array
    Else
        numCols = UBound(arr) - LBound(arr) + 1
        ReDim result(1 To 1, 1 To numCols)
        k = 1
        For i = LBound(arr) To UBound(arr)
            result(1, k) = arr(i)
            k = k + 1
        Next i
    End If
    convertTo2DArray = result
End Function

Sub paste_array(arr As Variant, addr As String, Optional ws, Optional wb)
    ' Set default worksheet if not provided
    Set ws = r.get_default_ws(ws) ' Assuming r is a helper module or class
    Set wb = r.get_default_wb(wb) ' Assuming r is a helper module or class
    wb.Activate
    
    ' if 1d array then convert
    If a.is_1d_array(arr) Then
       arr = a.convertTo2DArray(arr)
    End If
    
    ' Check if arr is a 2D array
    If is_2d_array(arr) Then
        Dim numRows As Long
        Dim numCols As Long
        
        ' Determine the number of rows and columns in the array
        numRows = UBound(arr, 1) - LBound(arr, 1) + 1
        numCols = UBound(arr, 2) - LBound(arr, 2) + 1
        
        ' Define the destination range
        Dim destRange As Range
        Set destRange = ws.Range(addr).Resize(numRows, numCols)
        
        ' Paste the array values to the destination range
        destRange.value = arr
    End If
End Sub

Sub pasteArray(arr As Variant, addr As String, Optional ws, Optional wb)
    ' Set default worksheet if not provided
    Set ws = r.get_default_ws(ws) ' Assuming r is a helper module or class
    Set wb = r.get_default_wb(wb) ' Assuming r is a helper module or class
    wb.Activate
    
    ' if 1d array then convert
    If a.is_1d_array(arr) Then
       arr = a.convertTo2DArray(arr)
    End If
    
    ' Check if arr is a 2D array
    If is_2d_array(arr) Then
        Dim numRows As Long
        Dim numCols As Long
        
        ' Determine the number of rows and columns in the array
        numRows = UBound(arr, 1) - LBound(arr, 1) + 1
        numCols = UBound(arr, 2) - LBound(arr, 2) + 1
        
        ' Define the destination range
        Dim destRange As Range
        Set destRange = ws.Range(addr).Resize(numRows, numCols)
        
        ' Paste the array values to the destination range
        destRange.value = arr
    End If
End Sub

Function to_array(x As Variant) As Variant()
    ' Check if x is already an array
    If IsArray(x) Then
        to_array = x
    ElseIf TypeName(x) = "Collection" Then
        Dim col0 As collection
        Set col0 = x
        Dim arr As Variant
        arr = Array()
        If col0.count > 0 Then
            ReDim arr(1 To col0.count)
            i = 1
            For Each it In col0
                arr(i) = it
                i = i + 1
            Next
        End If
        to_array = arr
    Else
        ' Create a one-dimensional array with value x
        ReDim arr(1 To 1) As Variant
        arr(1) = x
        to_array = arr
    End If
End Function

Function toString(arr As Variant, Optional columnSeparator As String = ";", Optional rowSeparator As String = vbCrLf) As String
    ' This function converts a 1D or 2D array to a string representation.
    ' The elements in each row are separated by columnSeparator, and rows are separated by rowSeparator.
    '
    ' Parameters:
    ' arr             : The array to be converted to a string.
    ' columnSeparator : The separator to use between columns (default is ";").
    ' rowSeparator    : The separator to use between rows (default is line break).
    '
    ' Returns:
    ' A string representation of the array.
    
    Dim result As String
    Dim i As Long, j As Long
    Dim rowString As String
    
    ' Check if the input is an array
    If Not IsArray(arr) Then
        Err.Raise 1001, "toString", "Input must be an array"
    End If
    
    ' Handle 1D array
    If is_1d_array(arr) Then
        For i = LBound(arr) To UBound(arr)
            result = result & arr(i) & columnSeparator
        Next i
        ' Remove the trailing column separator
        If Len(result) > 0 Then
            result = left(result, Len(result) - Len(columnSeparator))
        End If
    ' Handle 2D array
    ElseIf is_2d_array(arr) Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            rowString = ""
            For j = LBound(arr, 2) To UBound(arr, 2)
                rowString = rowString & arr(i, j) & columnSeparator
            Next j
            ' Remove the trailing column separator
            If Len(rowString) > 0 Then
                rowString = left(rowString, Len(rowString) - Len(columnSeparator))
            End If
            result = result & rowString & rowSeparator
        Next i
        ' Remove the trailing row separator
        If Len(result) > 0 Then
            result = left(result, Len(result) - Len(rowSeparator))
        End If
    Else
        Err.Raise 1002, "toString", "Array must be 1D or 2D"
    End If
    
    ' Return the resulting string
    toString = result
End Function

Function create_vector(num_rows As Integer, Optional default_value As Variant, Optional start_index = 1, Optional header_value = "", Optional as_2darray As Boolean = False) As Variant
    Dim arr() As Variant
    
    If Not IsMissing(default_value) Then
        array_value = default_value
    Else
        array_value = ""
    End If
    Dim i As Integer
    
    If (as_2darray) Then
        ReDim arr(1 To num_rows, 1 To 1)
        For i = start_index To num_rows
            If (i = start_index And header_value <> "") Then
                arr(i, 1) = header_value
            Else
                arr(i, 1) = array_value
            End If
        Next i
    Else
        ReDim arr(1 To num_rows)
        For i = start_index To num_rows
            If (i = start_index And header_value <> "") Then
                arr(i) = header_value
            Else
                arr(i) = array_value
            End If
        Next i
    End If
    'return value
    create_vector = arr
End Function

Function create_integer_vector(Optional Start As Integer = 1, Optional endVal As Integer = 100) As Variant
    Dim i As Integer
    Dim arr() As Integer
    
    If Start > endVal Then
        Err.Raise 5, , "Start value must be less than or equal to end value."
        Exit Function
    End If
    
    ReDim arr(Start To endVal)
    
    For i = Start To endVal
        arr(i) = i
    Next i
    
    create_integer_vector = arr
End Function

Function create_array(num_rows As Integer, num_cols As Integer, Optional default_value As Variant) As Variant
    Dim arr() As Variant
    ReDim arr(1 To num_rows, 1 To num_cols)
    If Not IsMissing(default_value) Then
        Dim i As Integer, j As Integer
        For i = 1 To num_rows
            For j = 1 To num_cols
                arr(i, j) = default_value
            Next j
        Next i
    End If
    create_array = arr
End Function

Function FillArray(arr As Variant, value As Variant) As Variant
    ' This function fills a 1D or 2D array with a specified value.
    '
    ' Parameters:
    ' arr   : The array to be filled.
    ' value : The value to fill the array with.
    '
    ' Returns:
    ' A variant containing the filled array.
    
    Dim i As Long, j As Long
    
    ' Check if the array is 2-dimensional
    If is_2d_array(arr) Then
        ' Fill each element of the 2D array with the specified value
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                arr(i, j) = value
            Next j
        Next i
    Else
        ' Fill each element of the 1D array with the specified value
        For i = LBound(arr) To UBound(arr)
            arr(i) = value
        Next i
    End If
    
    ' Return the filled array
    FillArray = arr
End Function

Function num_array_rows(arr As Variant) As Integer
    If IsArray(arr) Then
        If is_2d_array(arr) Then
           num_array_rows = UBound(arr, 1) - LBound(arr, 1) + 1
        Else
           num_array_rows = 1 ' convention
        End If
    Else
        num_array_rows = -1
    End If
End Function

' better name
Function numArrayRows(arr As Variant) As Integer
    If IsArray(arr) Then
        If is_2d_array(arr) Then
           numArrayRows = UBound(arr, 1) - LBound(arr, 1) + 1
        Else
           numArrayRows = 1 ' convention
        End If
    Else
        numArrayRows = -1
    End If
End Function

Function num_array_columns(arr As Variant) As Integer
    If IsArray(arr) Then
    If is_2d_array(arr) Then
       num_array_columns = UBound(arr, 2) - LBound(arr, 2) + 1
    Else
       num_array_columns = UBound(arr) - LBound(arr) + 1 ' convention
    End If
    Else
        num_array_columns = -1
    End If
End Function

Function castArrayAsString(arr As Variant) As Variant
    If a.is_1d_array(arr) Then
       For i = LBound(arr) To UBound(arr)
          arr(i) = CStr(arr(i))
       Next i
    Else
       For rw = LBound(arr, 1) To UBound(arr, 1)
        For c = LBound(arr, 2) To UBound(arr, 2)
           arr(rw, c) = CStr(arr(rw, c))
        Next c
       Next rw
    End If
    castArrayAsString = arr
End Function

' 3 utility
Sub printArrayRowsColumns(arr As Variant, Optional array_name As String = "")
    If array_name <> "" Then
    Debug.Print "array " & array_name & " has num rows:", a.num_array_rows(arr), "num columns:", a.num_array_columns(arr)
    Else
    Debug.Print "num rows:", a.num_array_rows(arr), "num columns:", a.num_array_columns(arr)
    End If
End Sub

Sub printArrayBounds(arr As Variant, Optional array_name As String = "")
    If IsArray(arr) Then
        If a.is_1d_array(arr) Then
            Debug.Print "lbound is:", LBound(arr), "ubound is:", UBound(arr)
        Else
            Debug.Print "lbound dim 1 is:", LBound(arr, 1), "ubound dim 1 is:", UBound(arr, 1), "lbound dim 2 is:", LBound(arr, 2), "ubound dim 2 is:", UBound(arr, 2)
        End If
    End If
End Sub

Function ItemInArray(item, arr As Variant) As Boolean
    ItemInArray = clls.item_exists(item, a.as_collection(arr))
End Function

'index in bounds
Function indexInBounds(arr As Variant, i As Long, Optional axis As Integer = 1) As Boolean
    Dim b As Boolean
    If a.is_1d_array(arr) Then
       b = i >= LBound(arr) And i <= UBound(arr)
    Else
       b = i >= LBound(arr, axis) And i <= UBound(arr, axis)
    End If
    indexInBounds = b
End Function

Function ArrayIsEmpty(arr0 As Variant) As Boolean
    ' This function checks if the provided array is empty.
    '
    ' Parameters:
    ' arr0 : The array to check for emptiness.
    '
    ' Returns:
    ' True if the array is empty, False otherwise.
    ' Raises an error if arr0 is not an array.
    
    ' Check if arr0 is an array
    If IsEmpty(arr0) Then
        ArrayIsEmpty = True
        Exit Function
    End If
    
    If Not IsArray(arr0) Then
        Err.Raise 13, "ArrayIsEmpty", "Input is not an array"
    End If
    
    ' Check if the array has any elements
    Dim testElement As Variant
    If a.is_1d_array(arr0) Then
       On Error Resume Next ' Use error handling to check for empty array
       testElement = arr0(LBound(arr0))
    ElseIf a.is_2d_array(arr0) Then
       On Error Resume Next ' Use error handling to check for empty array
       testElement = arr0(LBound(arr0), LBound(arr0))
    Else
       Err.Raise 14, "ArrayIsEmpty", "Input array is not 1D or 2D"
    End If
    
    If Err.Number <> 0 Then
        ' If an error occurs, the array is empty
        ArrayIsEmpty = True
    Else
        ' No error, the array has elements
        ArrayIsEmpty = False
    End If
    On Error GoTo 0 ' Reset error handling
End Function

Function ArraysAreEqual(arr0 As Variant, arr1 As Variant, Optional print0 As Boolean = False) As Boolean
    ' This function checks if two arrays are equal in terms of dimensions and element values.
    '
    ' Parameters:
    ' arr0 : The first array to compare.
    ' arr1 : The second array to compare.
    '
    ' Returns:
    ' True if the arrays are equal, False otherwise.
    
    Dim i As Long, j As Long, errmsg As String
    
    ' Check if both inputs are empty
    If (IsEmpty(arr0) Or IsEmpty(arr1)) Then
       If IsEmpty(arr0) = IsEmpty(arr1) Then
          ArraysAreEqual = True
          Exit Function
       Else
          ArraysAreEqual = False
       End If
    End If
    
    ' Check if both inputs are arrays
    If Not IsArray(arr0) Or Not IsArray(arr1) Then
        errmsg = str.subInStr("ArraysAreEqual: Both inputs must be arrays. Input arr0 is `@1`. Input arr1 is `@2`", TypeName(arr0), TypeName(arr1))
        Debug.Print errmsg
        If u.IsNull(arr0) And u.IsNull(arr1) Then
           ArraysAreEqual = True
        Else
           ArraysAreEqual = False
        End If
        Exit Function
    End If
    
    ' Check if both arrays are either 1D or 2D
    If (a.is_1d_array(arr0) And a.is_1d_array(arr1)) Or (a.is_2d_array(arr0) And a.is_2d_array(arr1)) Then
        ' Check if arrays have the same number of rows and columns
        If a.num_array_rows(arr0) <> a.num_array_rows(arr1) Or a.num_array_columns(arr0) <> a.num_array_columns(arr1) Then
            ArraysAreEqual = False
            GoTo print_mismatch
        End If
        
        ' Element-by-element comparison
        If a.is_1d_array(arr0) Then
            ' Compare 1D arrays
            For i = LBound(arr0) To UBound(arr0)
                If arr0(i) <> arr1(i) Then
                    ArraysAreEqual = False
                    GoTo print_mismatch
                End If
            Next i
        Else
            ' Compare 2D arrays
            For i = LBound(arr0, 1) To UBound(arr0, 1)
                For j = LBound(arr0, 2) To UBound(arr0, 2)
                    If arr0(i, j) <> arr1(i, j) Then
                        ArraysAreEqual = False
                        GoTo print_mismatch
                    End If
                Next j
            Next i
        End If
        
        
        
        ' If no mismatches found, arrays are equal
        ArraysAreEqual = True
        Exit Function
        
print_mismatch:
        If print0 And Not ArraysAreEqual Then
           a.printArray arr0
           a.printArray arr1
        End If
        Exit Function
    Else
        Err.Raise 1002, "ArraysAreEqual", "Both arrays must be either 1D or 2D."
    End If
End Function



Function FindArrayIndex(arr As Variant, value As Variant, Optional axis = 0, Optional index0 = 1) As Long
    ' find either the first row index of column `index0` or the first column index of row `index0` where arr(i,index0)==value
    If axis = 0 Then
        For rowIndex = LBound(arr, 1) To UBound(arr, 1)
            If arr(rowIndex, index0) = value Then
                FindArrayIndex = rowIndex
                Exit Function
            End If
        Next rowIndex
    Else
        For columnIndex = LBound(arr, 2) To UBound(arr, 2)
            If arr(index0, columnIndex) = value Then
                FindArrayIndex = columnIndex
                Exit Function
            End If
        Next columnIndex
    End If
    ' Return -1 if the value is not found
    FindArrayIndex = -1
End Function

Function FindArrayRowIndex(arr As Variant, column As Variant, value As Variant) As Long
    ' Get columnIndex of column
    columnIndex = FindArrayColumnIndex(arr, column)
    ' Loop through the array to find the matching value
    Dim rowIndex As Long
    rowIndex = a.FindArrayIndex(arr, value, axis:=0, index0:=columnIndex)
    ' Return rowIndex
    FindArrayRowIndex = rowIndex
End Function

Function FindArrayColumnIndex(arr As Variant, column As Variant) As Long
    ' Get the columnIndex of column in arr header
    Dim columnIndex As Long
    columnIndex = a.FindArrayIndex(arr, column, axis:=1)
    If columnIndex = -1 Then
       Err.Raise 1001, "FindArrayRowIndex", "column not found: " & column
    End If
    FindArrayColumnIndex = columnIndex
End Function

Function MatchArrayRowIndex(arr As Variant, key_columns As Variant, search_row As Variant, Optional dbg As Boolean = False) As Long
    ' multi key, value FindArrayRowIndex
    Dim numRows As Long
    Dim numCols As Long
    Dim i As Long, j As Long
    Dim match As Boolean
    
    ' Get the number of rows and columns in the array
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    'check types
    If Not IsArray(key_columns) Then
       Err.Raise 1001, "MatchArrayRowIndex", "key_columns is not array"
    ElseIf Not IsArray(search_row) Then
       Err.Raise 1002, "MatchArrayRowIndex", "search_row is not array"
    ElseIf a.array_length(key_columns) <> a.array_length(search_row) Then
       Err.Raise 1003, "MatchArrayRowIndex", "length mismatch key_columns, search_row arrays"
    End If
    
    ' Loop through each row in the array
    For i = LBound(arr, 1) To UBound(arr, 1)
        match = True
        
        ' Check if the key column values match
        For j = 0 To UBound(key_columns)
            ' Get the column index for the key column
            Dim colIndex As Long
            colIndex = a.FindArrayColumnIndex(arr, key_columns(j))
            
            ' Check if the values match
            If dbg Then
               Debug.Print "array value is:", arr(i, colIndex), "of type", VarType(arr(i, colIndex)), "search value is:", search_row(j), "of type", VarType(search_row(j)), "match is:", arr(i, colIndex) = search_row(j)
            End If
            If arr(i, colIndex) <> search_row(j) Then
                match = False
                Exit For
            End If
        Next j
        
        ' If all key column values match, return the row index
        If match Then
            MatchArrayRowIndex = i
            Exit Function
        End If
    Next i
    
    ' If no match is found, return -1
    MatchArrayRowIndex = -1
End Function

Function getIndexOfArray(value As Variant, arr As Variant) As Long
    ' Find the first occurrence of the value in the range
    Dim match_index As Variant
    match_index = value
    If VarType(match_index) = vbString Then
        match_index = Application.match(value, rng, 0)
        If IsError(match_index) Then
            Err.Raise vbObjectError + 1, "get_index_of_value", "Index " & value & " not found in range"
        End If
    ElseIf VarType(match_index) = vbLong Or VarType(match_index) = vbInteger Then
        If match_index < 1 Or match_index > rng.Cells.count Then
            Err.Raise vbObjectError + 1, "get_index_of_value", "Index " & value & " not found in range"
        End If
    Else
        Err.Raise vbObjectError + 1, "match_index", "Invalid index type"
    End If
    
    getIndexOfArray = match_index
End Function

' 4. array subsetting
Function get_array_column(arr, i, Optional row_offset = 0) As Variant
    Dim numRows As Long
    Dim numCols As Long
    Dim resultArr() As Variant
    Dim row As Long
    
    numRows = a.num_array_rows(arr)
    numCols = a.num_array_columns(arr)
    
    ' Check if i is out of bounds
    If Not a.indexInBounds(arr, CLng(i), axis:=2) Then
        Err.Raise vbObjectError + 1001, , "Column index out of bounds:", i
        Exit Function
    End If
    
    ReDim resultArr(1 To numRows, 1 To 1)
    
    k = 1
    For row = LBound(arr, 1) To UBound(arr, 1)
        If k <= row_offset Then
           GoTo nx_row
        End If
        resultArr(k, 1) = arr(row, i)
nx_row:
        k = k + 1
    Next row
    
    get_array_column = resultArr
End Function

Function get_array_row(arr As Variant, j As Long) As Variant
    Dim numRows As Long
    Dim numCols As Long
    Dim resultArr() As Variant
    Dim cl0 As Long
    
    numRows = a.num_array_rows(arr)
    numCols = a.num_array_columns(arr)
    
    ' Check if j is out of bounds
    If Not a.indexInBounds(arr, j, axis:=1) Then
        Err.Raise vbObjectError + 1001, , "Row index out of bounds: " & j
        Exit Function
    End If
    
    ReDim resultArr(1 To 1, 1 To numCols)
    
    k = 1
    For cl0 = LBound(arr, 2) To UBound(arr, 2)
        resultArr(1, k) = arr(j, cl0)
        k = k + 1
    Next cl0
    
    get_array_row = resultArr
End Function

Function getRowIndexes(arr As Variant) As collection
    Dim row_indexes As New collection
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        row_indexes.Add i
    Next i
    Set getRowIndexes = row_indexes
End Function

Function getColumnValue(arr As Variant, row_index As Integer, column_index_name As Variant) As Variant
    ' This function retrieves a value from a 2-dimensional array based on the provided row index and column index or column name.
    '
    ' Parameters:
    ' arr               : The 2-dimensional array from which to retrieve the value.
    ' row_index         : The row index of the value to retrieve.
    ' column_index_name : The column index (as an Integer) or column name (as a String) of the value to retrieve.
    '
    ' Returns:
    ' The value at the specified row and column in the array.
    
    Dim column_index As Integer
    
    ' Check if arr is a 2-dimensional array
    If Not is_2d_array(arr) Then
        Err.Raise 1001, "getColumnValue", "Input array must be 2-dimensional"
    End If
    
    ' Check if row_index is within the bounds of the array
    If row_index < LBound(arr, 1) Or row_index > UBound(arr, 1) Then
        Err.Raise 1002, "getColumnValue", "Row index out of bounds"
    End If
    
    ' Determine if column_index_name is a column index (Integer) or column name (String)
    If VarType(column_index_name) = vbInteger Then
        ' Use the provided column index
        column_index = column_index_name
    ElseIf VarType(column_index_name) = vbString Then
        ' Find the column index by matching the column name in the first row (header) of the array
        Dim found As Boolean
        found = False
        For i = LBound(arr, 2) To UBound(arr, 2)
            If StrComp(arr(LBound(arr, 1), i), column_index_name, vbTextCompare) = 0 Then
                column_index = i
                found = True
                Exit For
            End If
        Next i
        If Not found Then
            errmsg = str.subInStr("Column name `@1` not found in array", column_index_name)
            Err.Raise 1003, "getColumnValue", errmsg
        End If
    Else
        ' Raise an error if column_index_name is neither an Integer nor a String
        Err.Raise 1004, "getColumnValue", "Invalid column index or name"
    End If
    
    ' Check if column_index is within the bounds of the array
    If column_index < LBound(arr, 2) Or column_index > UBound(arr, 2) Then
        Err.Raise 1005, "getColumnValue", "Column index out of bounds"
    End If
    
    ' Retrieve and return the value from the array
    getColumnValue = arr(row_index, column_index)
End Function

'cast array as collection
Function as_collection(arr As Variant, Optional index = 1) As collection
Dim arr_col As New collection
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If a.is_2d_array(arr) Then
            arr_col.Add arr(i, index)
        Else
            arr_col.Add arr(i)
        End If
    Next i
Set as_collection = arr_col
End Function

' 2. Array subsetting
Function subset_rows(arr As Variant, start_row As Long, Optional end_row As Long) As Variant
    'convert to 2d array
    If Not is_2d_array(arr) Then
        arr = a.convertTo2DArray(arr)
    End If
    
    If end_row = 0 Then end_row = UBound(arr, 1)
    Dim result() As Variant
    ReDim result(1 To end_row - start_row + 1, LBound(arr, 2) To UBound(arr, 2))
    Dim i As Long, j As Long
    For i = start_row To end_row
        For j = LBound(arr, 2) To UBound(arr, 2)
            result(i - start_row + 1, j) = arr(i, j)
        Next j
    Next i
    subset_rows = result
End Function

Function subset_columns(arr As Variant, Optional start_column As Long, Optional end_column As Long, Optional start_row As Long, Optional end_row As Long) As Variant
    'convert to 2d array
    If Not is_2d_array(arr) Then
        arr = a.convertTo2DArray(arr)
    End If
    
    If start_row = 0 Then start_row = LBound(arr, 1)
    If end_row = 0 Then end_row = UBound(arr, 1)
    If start_column = 0 Then start_column = LBound(arr, 1)
    If end_column = 0 Then end_column = UBound(arr, 2)
    Dim result() As Variant
    ReDim result(1 To end_row - start_row + 1, 1 To end_column - start_column + 1)
    Dim i As Long, j As Long
    For i = start_row To end_row
        For j = start_column To end_column
            result(i - start_row + 1, j - start_column + 1) = arr(i, j)
        Next j
    Next i
    subset_columns = result
End Function

Function resize_array(arr0 As Variant, Optional r0 As Long, Optional r1 As Long, Optional c0 As Long, Optional c1 As Long) As Variant
    ' Subset rows from arr0
    Dim arr1 As Variant
    arr1 = subset_rows(arr0, r0, r1)
    
    ' Subset columns from arr1
    Dim arr2 As Variant
    arr2 = subset_columns(arr1, c0, c1)
    
    ' Return arr2 as the resized array
    resize_array = arr2
End Function

' if axis=0 subset on row indices, if axis=1 on column indices
Function subset_indices(arr As Variant, axis As Integer, indices As Variant) As Variant
    'convert to 2d array
    If Not is_2d_array(arr) Then
        arr = a.convertTo2DArray(arr)
    End If
    
    Dim result() As Variant
    Dim i As Long, j As Long, numCols As Long, numRows As Long
    'subset on rows
    If axis = 0 Then
        ReDim result(LBound(indices) To UBound(indices), LBound(arr, 2) To UBound(arr, 2))
        For i = LBound(indices) To UBound(indices)
            For j = LBound(arr, 2) To UBound(arr, 2)
                index0 = indices(i)
                result(i, j) = arr(index0, j)
            Next j
        Next i
    'subset on columns
    ElseIf axis = 1 Then
        ReDim result(LBound(arr, 1) To UBound(arr, 1), LBound(indices) To UBound(indices))
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(indices) To UBound(indices)
                index0 = indices(j)
                result(i, j) = arr(i, index0)
            Next j
        Next i
    End If
    subset_indices = result
End Function

Function select_column_names(arr, column_names) As Variant
    column_indices = a.get_indices(arr, column_names)
    result = a.subset_indices(arr, 1, column_indices)
    select_column_names = result
End Function

Function get_indices(arr As Variant, column_names As Variant) As Variant
    If Not is_2d_array(arr) Then
        Err.Raise 1002, "not 2d array"
        Exit Function
    End If
    Dim result() As Variant
    ReDim result(1 To UBound(column_names) - LBound(column_names) + 1)
    Dim i As Long, j As Long, column_name As String
    Dim header_row_index As Long
    header_row_index = LBound(arr, 1)
    For i = LBound(column_names) To UBound(column_names)
        column_name = column_names(i)
        Dim found As Boolean
        found = False
        For j = LBound(arr, 2) To UBound(arr, 2)
            If StrComp(arr(header_row_index, j), column_name, vbTextCompare) = 0 Then
                result(i - LBound(column_names) + 1) = j
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            Err.Raise 1003, "column_name " + column_name + " not in array"
            Exit Function
        End If
    Next i
    get_indices = result
End Function

Function select_array_columns(arr As Variant, column_names As Variant) As Variant
    Dim header As Variant
    Dim headerIndex As Object
    Dim selectedColumns As collection
    Dim resultArr() As Variant
    Dim i As Long, j As Long
    Dim colName As Variant
    Dim colIndex As Long
    Dim numRows As Long
    Dim numCols As Long
    
    ' convert to column_names to array if string
    If VarType(column_names) = vbString Then
       column_names = str.str_to_array(column_names)
    ElseIf Not IsArray(column_names) Then
       Err.Raise 1001, "select_array_columns", "column_names is not string or array"
    End If
    
    ' Get the header row
    header = get_array_row(arr, LBound(arr, 1))
    
    ' Create a collection to store selected columns
    Set selectedColumns = New collection
    
    ' Create a dictionary to store the column indexes for faster lookup
    Set headerIndex = CreateObject("Scripting.Dictionary")
    
    ' Populate the dictionary with column names and indexes
    For j = 1 To UBound(header, 2)
        headerIndex(header(1, j)) = j
    Next j
    
    ' Loop over column names and select the columns
    For Each colName In column_names
        ' Check if the column name exists in the header
        If headerIndex.exists(colName) Then
            ' Get the index of the column
            colIndex = headerIndex(colName)
            
            ' Add the column to the selected columns collection
            selectedColumns.Add colIndex
        End If
    Next colName
    
    ' Get the number of selected columns
    numCols = selectedColumns.count
    
    If numCols > 0 Then
        ' Get the number of rows
        numRows = UBound(arr, 1)
        
        ' Re-dimension the result array
        ReDim resultArr(1 To numRows, 1 To numCols)
        
        ' Copy the selected columns to the result array
        For i = 1 To numCols
            colIndex = selectedColumns.item(i)
            For j = 1 To numRows
                resultArr(j, i) = arr(j, colIndex)
            Next j
        Next i
    Else
        Err.Raise 1001, "select_array_columns", "no columns selected!"
    End If
    
    ' Return the selected columns array
    select_array_columns = resultArr
End Function

Function setArrayHeader(arr As Variant, header As Variant) As Variant
    ' Check if arr is a 2-dimensional array
    If Not is_2d_array(arr) Then
        Err.Raise 1001, "setArrayHeader", "arr must be a 2-dimensional array"
        Exit Function
    End If
    
    ' Check if header is an array
    If Not IsArray(header) Then
        Err.Raise 1002, "setArrayHeader", "header must be an array"
        Exit Function
    End If
    
    ' Check if the length of header matches the number of columns in arr
    If UBound(header) - LBound(header) + 1 <> UBound(arr, 2) - LBound(arr, 2) + 1 Then
        Err.Raise 1002, "setArrayHeader", "length of header must match number of columns in arr"
        Exit Function
    End If
    
    ' Check if arr has at least one row
    If UBound(arr, 1) - LBound(arr, 1) + 1 < 1 Then
        Err.Raise 1002, "setArrayHeader", "arr must have at least one row"
        Exit Function
    End If
    
    ' Set the values of header to the first row of arr
    Dim i As Long
    For i = LBound(arr, 2) To UBound(arr, 2)
        arr(LBound(arr, 1), i) = header(i - LBound(arr, 2))
    Next i
    
    ' Return the modified array
    setArrayHeader = arr
End Function

Function getArrayColumnIndex(arr As Variant, column_name) As Long
    Dim headerArr As Variant
    ' Check if arr is a 2-dimensional array
    If Not is_2d_array(arr) Then
        Err.Raise 1001, "getArrayColumnIndex", "arr must be a 2-dimensional array"
        Exit Function
    End If
    
    If a.num_array_rows(arr) < 1 Then
       Err.Raise 1002, "getArrayColumnIndex", "arr has no records"
    End If
    
    headerVector = a.ConvertTo1DArray(a.get_array_row(arr, LBound(arr, 1)))
    
    If VarType(column_name) = vbString Then
       For i = LBound(headerVector) To UBound(headerVector)
          If column_name = headerVector(i) Then
             getArrayColumnIndex = i
             Exit Function
          End If
       Next
       Err.Raise 1004, "getArrayColumnIndex", "column_name not found: " & column_name
    ElseIf VarType(column_name) = vbInteger Then
       getArrayColumnIndex = column_name
    Else
       Err.Raise 1003, "getArrayColumnIndex", "column_name is not string or integer"
    End If
End Function

Function QueryArray(arr As Variant, ParamArray criteria()) As Variant
    ' This function filters a 2-dimensional array based on multiple criteria.
    ' Each pair of criteria consists of a column name and the value to filter by.
    '
    ' Parameters:
    ' arr      : The 2-dimensional array to be filtered.
    ' criteria : An array of criteria pairs where each pair is a column name followed by a value.
    '
    ' Returns:
    ' A 2-dimensional array containing only the rows that match all criteria.
    
    Dim i As Long
    Dim j As Long
    Dim numRows As Long
    Dim numCols As Long
    Dim criteriaCount As Long
    Dim colIndex As Long
    Dim resultArr() As Variant
    Dim tempArr As Variant
    Dim headerArr As Variant
    Dim match As Boolean
    Dim filteredRows As New collection
    
    ' Validate that arr is a 2-dimensional array
    If Not is_2d_array(arr) Then
        Err.Raise 1001, "QueryArray", "Input array must be 2-dimensional"
    End If
    
    ' Validate that criteria has at least 2 arguments and an even number of arguments
    criteriaCount = UBound(criteria) - LBound(criteria) + 1
    If criteriaCount < 2 Or criteriaCount Mod 2 <> 0 Then
        Err.Raise 1002, "QueryArray", "Criteria must have at least 2 arguments and an even number of arguments"
    End If
    
    ' Initialize the result array with the input array
    tempArr = arr
    headerArr = a.get_array_row(arr, 1)
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    ' Loop through each criteria pair and filter the array
    For i = LBound(criteria) To UBound(criteria) Step 2
        ' Find the column index for the current column name
        colIndex = FindArrayColumnIndex(tempArr, criteria(i))
        
        ' Filter the array based on the current criteria pair
        Set filteredRows = New collection
        For j = LBound(tempArr, 1) To UBound(tempArr, 1)
            If tempArr(j, colIndex) = criteria(i + 1) Then
                ' Add the row index to the collection if it matches the criteria
                filteredRows.Add j
            End If
        Next j
        
        ' If no rows are found, return the headerArr
        filteredRowsCount = filteredRows.count
        If filteredRowsCount = 0 Then
            QueryArray = headerArr
            Exit Function
        End If
        
        ' Rebuild the array with only the rows that match the criteria
        Dim k As Long
        ReDim resultArr(1 To filteredRowsCount, 1 To numCols)
        For k = 1 To filteredRowsCount
            For j = 1 To numCols
                r0 = filteredRows(k)
                resultArr(k, j) = tempArr(r0, j)
            Next j
        Next k
        
        ' Append headerArr to resultArr
        resultArr = a.concatArrays(headerArr, resultArr)
        
        ' Update tempArr with the filtered result for the next iteration
        tempArr = resultArr
        
    Next i
    
    ' Return the filtered array
    QueryArray = resultArr
End Function

'4.2 array filtering
Function RemoveNullsFromArray(arr As Variant, ParamArray filterColumns() As Variant) As Variant
    ' This function filters out rows from a 2D array where specified columns contain null values.
    ' arr - The 2D array to be filtered.
    ' filterColumns - The indices of the columns to check for null values.
    Dim i As Long, j As Long
    Dim includeRow As Boolean
    Dim currentRow As Variant
    Dim currentCell As Variant
    Dim headerArr As Variant
    Dim filteredArr As Variant
    
    ' Initialize the filteredArr as header
    If a.num_array_rows(arr) < 2 Then
       Debug.Print "RemoveNullsFromArray: arr has no records"
       RemoveNullsFromArray = arr
    End If
    
    filteredArr = a.get_array_row(arr, LBound(arr, 1))
    
    ' Loop through each row of the array
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        ' Assume the row is to be included until a null value is found
        includeRow = True
        
        ' Check each specified filter column for null values
        For j = 0 To UBound(filterColumns)
            
            ' Get the current cell value
            columnIndex = getArrayColumnIndex(arr, filterColumns(j))
            currentCell = arr(i, columnIndex)
            
            ' Check if the current cell is null (empty or zero-length string)
            If IsEmpty(currentCell) Or currentCell = "" Or u.IsNull(currentCell) Then
                ' If a null value is found, exclude the row and exit the loop
                includeRow = False
                Exit For
            End If
        Next j
        
        ' If the row does not contain null values in the filter columns, add it to the result
        If includeRow Then
            ' Get the current row as a 2D array
            currentRow = a.get_array_row(arr, i)
            ' Add the current row to the resultArr
            filteredArr = a.concatArrays(filteredArr, currentRow)
        End If
    Next i
    
    RemoveNullsFromArray = filteredArr
    
End Function

Function RemoveNullsFromVector(ByVal arr As Variant) As Variant
    RemoveNullsFromVector = RemoveNullsFromArray(convertTo2DArray(arr), 1)
End Function

Sub test()
Dim filteredArr As Variant
Debug.Print a.num_array_rows(filteredArr)
End Sub
Function FilterArrayOnPattern(arr As Variant, pattern As String, ParamArray filterColumns() As Variant) As Variant
    ' This function filters out rows from a 2D array where specified columns match a given regular expression pattern.
    ' arr - The 2D array to be filtered.
    ' pattern - The regular expression pattern to match.
    ' filterColumns - The indices or names of the columns to check for pattern matches.
    
    Dim i As Long, j As Long
    Dim includeRow As Boolean
    Dim currentRow As Variant
    Dim currentCell As Variant
    Dim headerArr As Variant
    Dim filteredArr As Variant
    Dim regex As Object
    
    ' Check if array to filter has any rows
    If a.numArrayRows(arr) < 1 Then
       Debug.Print "FilterArrayOnPattern: arr has no records"
       FilterArrayOnPattern = arr
       Exit Function
    End If
    
    ' Create a regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
    
    ' Loop through each row of the array
    For i = LBound(arr, 1) To UBound(arr, 1)
        ' Assume the row is to be excluded until a pattern match is found
        includeRow = False
        
        ' Check each specified filter column for pattern matches
        For j = 0 To UBound(filterColumns)
            
            ' Get the current cell value
            columnIndex = getArrayColumnIndex(arr, filterColumns(j))
            currentCell = CStr(arr(i, columnIndex))
            
            ' Check if the current cell matches the pattern
            If regex.test(currentCell) Then
                ' If a pattern match is found, include the row and exit the loop
                includeRow = True
                Exit For
            End If
        Next j
        
        ' If the row does not contain pattern matches in the filter columns, add it to the result
        If includeRow Then
            ' Get the current row as a 2D array
            currentRow = a.get_array_row(arr, i)
            ' Add the current row to the result filteredArr
            If a.numArrayRows(filteredArr) < 1 Then
               filteredArr = currentRow
            Else
               filteredArr = a.concatArrays(filteredArr, currentRow)
            End If
        End If
    Next i
    
    FilterArrayOnPattern = filteredArr
End Function

Function getNamedArrayValue(arr As Variant, columnname As String) As Variant
    ' This function retrieves the value from a named array based on the column name.
    ' It raises an error if the array does not have exactly one row.
    '
    ' Parameters:
    ' arr        : The named array from which to retrieve the value.
    ' columnname : The name of the column from which to retrieve the value.
    '
    ' Returns:
    ' The value from the specified column of the named array.
    
    Dim numRows As Long
    Dim columnIndex As Long
    Dim value As Variant
    
    ' Check if the array is 2-dimensional
    If Not is_2d_array(arr) Then
        Err.Raise 1001, "getNamedArrayValue", "Input array must be 2-dimensional"
    End If
    
    ' Get the number of rows in the array
    numRows = num_array_rows(arr)
    
    ' Raise an error if the array does not have exactly two rows: header + values
    If numRows < 2 Then
        Err.Raise 1002, "getNamedArrayValue", "Arr only has header (no data row)"
    ElseIf numRows > 2 Then
        Err.Raise 1002, "getNamedArrayValue", "Arr has multiple data rows"
    End If
    
    ' Find the column index by the column name
    columnIndex = FindArrayColumnIndex(arr, columnname)
    
    ' Raise an error if the column name is not found
    If columnIndex = -1 Then
        Err.Raise 1003, "getNamedArrayValue", "Column name not found: " & columnname
    End If
    
    ' Retrieve the value from the array
    value = arr(2, columnIndex)
    
    ' Return the value
    getNamedArrayValue = value
End Function

'5. Array combining, joining, merging
Function concatArrays(arr0 As Variant, arr1 As Variant, Optional axis As Integer = 1) As Variant
    Dim numCols0 As Long
    Dim numCols1 As Long
    Dim numRows0 As Long
    Dim numRows1 As Long
    Dim outputArr As Variant
    Dim i As Long, j As Long
    
    ' convert to 2s arrays
    arr0 = a.convertTo2DArray(arr0, axis:=axis)
    arr1 = a.convertTo2DArray(arr1, axis:=axis)
    
    ' Check if the arrays have compatible dimensions
    numCols0 = UBound(arr0, 2) - LBound(arr0, 2) + 1
    numCols1 = UBound(arr1, 2) - LBound(arr1, 2) + 1
    numRows0 = UBound(arr0, 1) - LBound(arr0, 1) + 1
    numRows1 = UBound(arr1, 1) - LBound(arr1, 1) + 1
    
    If numCols0 <> numCols1 Then
        ' Raise an error if the number of columns doesn't match
        Err.Raise vbObjectError + 1001, , "Incompatible array dimensions"
        Exit Function
    End If
    
    ' Initialize the output array
    Dim r0 As Long, c0 As Long
    r0 = 1
    c0 = 1
    ReDim outputArr(r0 To numRows0 + numRows1, c0 To numCols0)

    ' Copy the values from arr0 to the output array
    For i = LBound(arr0, 1) To UBound(arr0, 1)
        c0 = 1
        For j = LBound(arr0, 2) To UBound(arr0, 2)
            outputArr(r0, c0) = arr0(i, j)
            c0 = c0 + 1
        Next j
        r0 = r0 + 1
    Next i
    
    ' Copy the values from arr1 to the output array
    For i = LBound(arr1, 1) To UBound(arr1, 1)
        c0 = 1
        For j = LBound(arr1, 2) To UBound(arr1, 2)
            outputArr(r0, c0) = arr1(i, j)
            c0 = c0 + 1
        Next j
        r0 = r0 + 1
    Next i
    
    ' Return the appended array
    concatArrays = outputArr
End Function

Function AppendColumn(arr0 As Variant, Optional values As Variant = "", Optional header_value As String = "") As Variant
    Dim numRows As Long
    Dim i As Long
    Dim r_index As Long
    Dim arr As Variant
    
    ' Copy original array to prevent overwrite
    arr = arr0
    
    ' Determine the number of rows in the original array
    numRows = UBound(arr, 1) - LBound(arr, 1) + 1
    
    ' Resize the original array to add a new column
    ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2) + 1)
    
    ' Check if values is 1 or 2 dimensional array
    If IsArray(values) Then
        ' If values is an array, fill the new column of arr with values
        If a.is_2d_array(values) Then
            For i = LBound(arr, 1) To UBound(arr, 1)
               r_index = i - LBound(arr, 1) + LBound(values, 1)
               arr(i, UBound(arr, 2)) = values(r_index, LBound(values, 2))
            Next i
        Else
            For i = LBound(arr, 1) To UBound(arr, 1)
               r_index = i - LBound(arr, 1) + LBound(values)
               arr(i, UBound(arr, 2)) = values(r_index)
            Next i
        End If
    ElseIf Not IsEmpty(values) Then
        ' If values is not an array and not empty, create values_array from values
        Dim values_array() As Variant
        ReDim values_array(LBound(arr, 1) To UBound(arr, 1))
        
        ' Fill values_array with values for the number of rows in arr
        For i = LBound(arr, 1) To UBound(arr, 1)
            values_array(i) = values
        Next i
        
        If header_value <> "" Then
           values_array(LBound(values_array)) = header_value
        End If
        
        ' Fill the new column of arr with values_array
        For i = LBound(arr, 1) To UBound(arr, 1)
            arr(i, UBound(arr, 2)) = values_array(i)
        Next i
    End If
    
    ' Return the modified array
    AppendColumn = arr
End Function



