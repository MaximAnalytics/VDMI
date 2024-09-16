' Collection functions and subroutines

' 0. Utilities
'printItems(col) => Outputs each item in the collection to the debug console.
'printKeys

' 1. Logical
'item_exists(item, collection) => Checks if an item exists within a given collection.

' 2. Transformations
'shuffle(col, items_to_back, items_to_front) => Shuffles items in a collection by moving specified numbers to back and front.
'sort_collection(col, ascending) => Sorts a collection in ascending or descending order.
'alternate_items(col) => Returns a collection with items arranged alternately from front and back.
'collectionToString(col, delimiter) => Converts a collection to a delimited string.

' 3. generators
'toCollection(x, delim) => Converts a string, array, or range to a collection.
'getCollection(args()) => Creates a collection from a parameter array of items.
' TODO: CopyCollection(col) as collection

' 4. combine, merge, join, unique
'concatCollections(col0, col1) => Concatenates two collections into one.
'getMatchingItems(col0, col1) => Returns a collection of items present in both input collections.
'getComplementItems(col0, col1) => Returns a collection of items in col1 that are not in col0.
'distinctItems(col0) => Returns a collection of distinct items from the input collection.

Sub test_collection_functions()
    Dim col As collection, col1 As collection, col0 As collection, complement0 As collection
    
    ' subsetting
    ' Test pop function
    Dim testCol As New collection
    Set testCol = clls.ItemsToCollection("Item1", "Item2", "Item3")
    Debug.Assert pop(testCol) = "Item3"
    Debug.Assert testCol.count = 2
    
    Debug.Assert pop(testCol, index:=1) = "Item1"
    Debug.Assert pop(testCol, index:=1) = "Item2"
    
    ' Test getObjectByKey function
    Dim testObj As New collection
    testObj.Add New collection, "Key1"
    Debug.Assert getObjectByKey(testObj, "Key1") Is Nothing = False
    Debug.Assert getObjectByKey(testObj, "Key2") Is Nothing
    
    ' Test getItem function
    Dim testItemCol As New collection
    Set testItemCol = clls.ItemsToCollection("First", "Second", "Third")
    Debug.Assert getItem(testItemCol, -1) = "Third"
    Debug.Assert getItem(testItemCol, -2) = "Second"
    Debug.Assert getItem(testItemCol, -4) = Empty
    
    ' 1. Logical
    Set col0 = clls.toCollection("A,B")
    Set col1 = clls.toCollection("B,C,D")
    Set complement0 = clls.getComplementItems(col0, col1)
    Debug.Assert clls.collectionToString(complement0) = "C,D"
    
    ' 2. Transformations
    Set col = a.as_collection(Array("C", "B", "A", 1, 2, 3))
    Set col1 = clls.sort_collection(col, True)
    Debug.Assert clls.collectionToString(col1) = "1,2,3,A,B,C"
    
    Set col = a.as_collection(Array("LN1", "LN2", "LN3"))
    ' move last item to front
    Debug.Assert collectionToString(clls.shuffle(col, 0, 1)) = "LN3,LN1,LN2"
    
    ' move first item to back
    Set col = a.as_collection(Array("LN1", "LN2", "LN3"))
    Debug.Assert collectionToString(clls.shuffle(col, 1, 0)) = "LN2,LN3,LN1"

    ' 4. combine, merge, join, unique
    Dim concatAB As collection, concatA0 As collection, concat0B As collection
    Set concatAB = clls.concatCollections(toCollection("A"), toCollection("B"))
    Debug.Assert concatAB.count = 2
    Set col = New collection
    Set concatA0 = clls.concatCollections(toCollection("A"), col)
    Debug.Assert concatA0(1) = "A"
    Set concat0B = clls.concatCollections(col, toCollection("B"))
    Debug.Assert concat0B(1) = "B"
    
End Sub

'0. Collection utilities
Sub printItems(col As collection, Optional as_string As Boolean = False)
    Dim item As Variant
    If as_string Then
        Debug.Print clls.collectionToString(col)
    Else
        For Each item In col
            Debug.Print item
        Next item
    End If
End Sub

Sub printKeys(col As collection)
    ' This subroutine attempts to print the keys of the items in a collection.
    ' It assumes that the collection has been constructed with keys.
    ' If the collection does not have keys, it will print the index of the item instead.
    '
    ' Parameters:
    '   - col: The collection whose keys are to be printed.
    
    Dim i As Integer
    Dim key As Variant
    Dim item As Variant
    
    ' Loop through each item in the collection
    For i = 1 To col.count
        ' Attempt to retrieve the key using error handling
        On Error Resume Next
        key = col(i) ' Attempt to access the item by its index
        If Err.Number = 0 Then
            ' If no error, print the key
            Debug.Print "Key for item " & i & ": " & key
        Else
            ' If error, print the index instead
            Debug.Print "Item " & i & " does not have a key or error retrieving key."
        End If
        On Error GoTo 0 ' Reset error handling
    Next i
End Sub

' Logical
Function item_exists(item As Variant, collection As collection) As Boolean
    Dim i As Integer
    For i = 1 To collection.count
        If collection(i) = item Then
            item_exists = True
            Exit Function
        End If
    Next i
    item_exists = False
End Function

' Checks if a key exists in a given collection
' Parameters:
'   - col: The collection to check for the key
'   - key: The key to search for in the collection
' Returns: True if the key exists, False otherwise
Function KeyExists(col As collection, key As String) As Boolean
    Dim key_exist As Boolean, item As Variant
    key_exist = False
    ' Attempt to retrieve the item with the given key
    On Error GoTo ErrHandler
    key_exist = IsObject(col.item(key))
    key_exist = True
    On Error GoTo 0
ErrHandler:
    KeyExists = key_exist
End Function

' Subsetting
' Pop an item from a collection by index
Function pop(ByRef col As collection, Optional index As Variant) As Variant
    ' This function removes an item from the collection at the specified index and returns it.
    ' If no index is provided, it defaults to the last item in the collection.
    ' Parameters:
    '   - col: The collection to pop an item from.
    '   - index: (Optional) The index of the item to pop. Defaults to the last item.
    '
    ' Returns: The item that was removed from the collection.
    
    Dim result As Variant
    
    If col.count = 0 Then
        Exit Function
    End If
    
    If IsMissing(index) Then
        index = col.count ' Default to the last item
    End If
    
    ' Check if the index is within the bounds of the collection
    If index >= 1 And index <= col.count Then
        result = col(index)
        col.Remove index
    Else
        Err.Raise 9, "pop", "Index out of bounds"
    End If
    
    ' Return the popped item
    pop = result
End Function

' Get an object from a collection by key
Function getObjectByKey(col As collection, key As Variant) As Object
    ' This function retrieves an object from the collection by its key.
    ' If the key is not found, it returns Nothing.
    ' Parameters:
    '   - col: The collection to retrieve the object from.
    '   - key: The key of the object to retrieve.
    '
    ' Returns: The object associated with the key, or Nothing if the key is not found.
    
    Dim item As Variant
    On Error Resume Next ' Ignore errors to handle the case when the key is not found
    Set getObjectByKey = col(key)
    On Error GoTo 0 ' Reset error handling
End Function

' Get an item from a collection with Python-like indexing
Function getItem(col As collection, shift As Integer) As Variant
    ' This function retrieves an item from the collection using Python-like negative indexing.
    ' Parameters:
    '   - col: The collection to retrieve the item from.
    '   - shift: The index of the item to retrieve. Supports negative indexing.
    '
    ' Returns: The item at the specified index, or Nothing if the index is out of bounds.
    
    Dim index As Integer
    If shift < 0 Then
        index = col.count + shift + 1 ' Convert negative index to positive
    Else
        index = shift
    End If
    
    ' Check if the index is within the bounds of the collection
    If index >= 1 And index <= col.count Then
        item = col(index)
        getItem = item
    Else
        getItem = Empty ' Return Empty if index is out of bounds
    End If
End Function

' Collection ordering
Function alternate_items(col As collection) As collection
    Dim i As Long
    Dim newCol As collection
    Dim frontIndex As Long, backIndex As Long
    
    Set newCol = New collection
    
    frontIndex = 1
    backIndex = col.count
    
    ' Continue until front and back indices cross
    While frontIndex <= backIndex
        ' Add the front item
        newCol.Add col(frontIndex)
        frontIndex = frontIndex + 1
        
        ' If indices haven't crossed, add the back item
        If frontIndex <= backIndex Then
            newCol.Add col(backIndex)
            backIndex = backIndex - 1
        End If
    Wend
    
    Set alternate_items = newCol
End Function

Function shuffle(ByVal col As collection, Optional items_to_back As Integer = 0, Optional items_to_front As Integer = 0)
    ' This function shuffles the items in a collection by moving a specified number of items to the back and front of the collection.
    ' Parameters:
    '   - col: The collection to shuffle.
    '   - items_to_back: The number of items to move to the back of the collection.
    '   - items_to_front: The number of items to move to the front of the collection.
    
    Dim tempToFront As collection, tempToBack As collection, part1 As collection, part2 As collection
    Dim i As Integer
    Set col0 = col
    ' Create a temporary collection to hold the shuffled items
    Set tempToFront = New collection
    Set tempToBack = New collection
    
    ' Move the specified number of items from the front to the back of the collection
    If items_to_back > 0 Then
        For i = 1 To items_to_back
            tempToBack.Add col(1)
            col.Remove 1
        Next i
    End If
    
    ' Move the specified number of items to the front of the collection
    If items_to_front > 0 Then
        For i = 1 To items_to_front
            tempToFront.Add col(col.count)
            col.Remove col.count
        Next i
    End If
    
    ' Add the remaining items from the original collection to the shuffled collection
    Set part1 = clls.concatCollections(tempToFront, col)
    Set part2 = clls.concatCollections(part1, tempToBack)
    Set shuffle = part2
    
    ' Clean up the temporary collection
    Set tempToFront = Nothing
    Set tempToBack = Nothing
End Function

Function sort_collection(col As collection, Optional ascending As Boolean = True) As collection
    Dim arr() As Variant
    Dim i As Integer
    Dim temp As Variant
    Dim sortedCol As New collection
    Dim item
    
    ' Convert Collection to Array
    ReDim arr(1 To col.count)
    For i = 1 To col.count
        If IsDate(col(i)) Or IsNumeric(col(i)) Or VarType(col(i)) = vbString Then
            arr(i) = col(i)
        Else
            Err.Raise 5, , "Item " & i & " is not a sortable type (must be String, Date, or Numeric)."
            Exit Function
        End If
    Next i

    ' Sort Array
    Dim j As Integer
    If ascending Then
        For i = LBound(arr) To UBound(arr) - 1
            For j = i + 1 To UBound(arr)
                If arr(i) > arr(j) Then
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                End If
            Next j
        Next i
    Else
        For i = LBound(arr) To UBound(arr) - 1
            For j = i + 1 To UBound(arr)
                If arr(i) < arr(j) Then
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                End If
            Next j
        Next i
    End If

    ' Create Sorted Collection
    For Each item In arr
        sortedCol.Add item
    Next item

    Set sort_collection = sortedCol
End Function

' Collection set operations
' Returns a collection containing distinct items from the input collection
Function distinctItems(col0 As collection) As collection
    Dim dict As Object
    Dim item As Variant
    Dim result As collection
    
    ' Use a dictionary to maintain unique items
    Set dict = CreateObject("Scripting.Dictionary")
    Set result = New collection
    
    ' Loop through each item in the input collection
    For Each item In col0
        ' Add the item to the dictionary if it's not already present
        If Not dict.exists(item) Then
            dict.Add item, item
            result.Add item ' Add the item to the result collection
        End If
    Next item
    
    ' Return the collection of distinct items
    Set distinctItems = result
End Function

' Concatenates two collections into one
Function concatCollections(col0 As collection, col1 As collection) As collection
    Dim result As collection
    Dim item As Variant
    
    ' Create a new collection to store the concatenated results
    Set result = New collection
    
    ' Add all items from the first collection to the result
    For Each item In col0
        result.Add item
    Next item
    
    ' Add all items from the second collection to the result
    For Each item In col1
        result.Add item
    Next item
    
    ' Return the concatenated collection
    Set concatCollections = result
End Function

' Returns a collection with items that are present in both input collections
Function getMatchingItems(col0 As collection, col1 As collection) As collection
    Dim dict As Object
    Dim item As Variant
    Dim result As collection
    
    ' Use a dictionary to maintain unique items
    Set dict = CreateObject("Scripting.Dictionary")
    Set result = New collection
    
    ' Add all items from the first collection to the dictionary
    For Each item In col0
        dict.Add item, item
    Next item
    
    ' Loop through the second collection
    For Each item In col1
        ' If the item is in the dictionary, add it to the result collection
        If dict.exists(item) Then
            result.Add item
        End If
    Next item
    
    ' Return the collection of matching items
    Set getMatchingItems = result
End Function

Function getComplementItems(col0 As collection, col1 As collection) As collection
    ' Create a new collection to store the complementary items => items in col1 but not in col0
    Dim complementaryItems As New collection
    Dim item As Variant
    ' Loop through each item in col1
    For Each item In col1
        If clls.item_exists(item, col0) Then GoTo nx_item
        ' If the item does not exist in col0, add it to the complementaryItems collection
        complementaryItems.Add item
nx_item:
    Next item
    
    ' Return the collection of complementary items
    Set getComplementItems = complementaryItems
End Function

' Collection transformation
Function collectionToString(col As collection, Optional delimiter As String = ",") As String
    Dim result As String
    Dim i As Integer
    
    ' Loop through each item in the collection
    For i = 1 To col.count
        ' Append the item to the result string
        result = result & col(i)
        
        ' Add delimiter if it's not the last item
        If i < col.count Then
            result = result & delimiter
        End If
    Next i
    
    ' Return the result string
    collectionToString = result
End Function

Function CollectionToArray(col As collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim n As Long
    
    ' Get the number of items in the collection
    n = col.count
    
    ' Dimension the array as (1 to n, 1 to 1)
    ReDim arr(1 To n, 1 To 1)
    
    ' Populate the array with the collection items
    For i = 1 To n
        arr(i, 1) = col(i)
    Next i
    
    ' Return the array
    CollectionToArray = arr
End Function

' Collection generators
' Converts a string, array, or range to a collection
Function toCollection(x As Variant, Optional delim As String = ",") As collection
    Dim result As collection
    Dim cell As Range
    Dim item As Variant
    Dim i As Long
    
    ' Create a new collection to store the results
    Set result = New collection
    
    ' Determine the type of the input and convert accordingly
    If TypeName(x) = "String" Then
        ' Split the string by comma and add each item to the collection
        Dim items() As String
        items = Split(x, delim)
        For i = LBound(items) To UBound(items)
            result.Add Trim(items(i))
        Next i
    ElseIf TypeName(x) = "Range" Then
        ' Add each cell's value in the range to the collection
        For Each cell In x.Cells
            result.Add cell.value
        Next cell
    ElseIf IsArray(x) Then
        ' Add each item in the array to the collection
        For i = LBound(x) To UBound(x)
            result.Add x(i)
        Next i
    Else
        ' Raise an error if the input is not a string, array, or range
        Err.Raise vbObjectError + 1, "toCollection", "Invalid input type"
    End If
    
    ' Return the converted collection
    Set toCollection = result
End Function

Function getCollection(ParamArray args() As Variant) As collection
    ' This function creates a new Collection and adds the items passed through the parameter array.
    ' Parameters:
    '   - args(): A parameter array of items to be added to the collection.
    '
    ' Returns: A Collection object containing the items from the parameter array.
    
    Dim result As collection
    Dim item As Variant
    Dim i As Long
    
    ' Create a new collection to store the results
    Set result = New collection
    
    ' Loop through each item in the parameter array
    For i = LBound(args) To UBound(args)
        ' Add each item to the collection
        result.Add args(i)
    Next i
    
    ' Return the collection with all items added
    Set getCollection = result
End Function

' Collection generation
' Converts a parameter array to a collection
' Parameters:
'   - ParamArray items(): An array of items to be added to the collection
' Returns: A collection containing the items from the parameter array
Function ItemsToCollection(ParamArray items()) As collection
    Dim result As New collection
    Dim i As Long
    
    ' Loop through each item in the parameter array and add it to the collection
    For i = LBound(items) To UBound(items)
        result.Add items(i)
    Next i
    
    ' Return the collection with all items added
    Set ItemsToCollection = result
End Function

