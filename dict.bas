' This function takes an input string and delimiters for key-value pairs and items, and returns a dictionary object
' with keys and values populated based on the input string.
'
' @param input_string The string containing key-value pairs separated by delimiters.
' @param keyvalue_delimiter The delimiter used to separate keys from values within a pair. Default is "=".
' @param item_delimiter The delimiter used to separate different key-value pairs. Default is ";".
' @return A dictionary object with keys and values based on the input string.
Function getDictionaryFromString(input_string As String, Optional keyvalue_delimiter As String = "=", Optional item_delimiter As String = ";") As Object
    Dim dictObj As Object
    Dim items As Variant
    Dim keyvaluepair As Variant
    Dim i As Integer
    Dim key As String
    Dim value As String
    
    ' Create a new dictionary object
    Set dictObj = CreateObject("Scripting.Dictionary")
    
    ' Separate the input string by the item delimiter to get items array
    items = Split(input_string, item_delimiter)
    
    ' Loop over elements of items
    For i = LBound(items) To UBound(items)
        ' Separate each item by the key-value delimiter
        keyvaluepair = Split(items(i), keyvalue_delimiter)
        
        ' Get key and value from the key-value pair
        key = keyvaluepair(0)
        value = keyvaluepair(1)
        
        ' Add to dictObj value with key
        dictObj.Add key, value
    Next i
    
    ' Return the dictionary object
    Set getDictionaryFromString = dictObj
End Function

' This subroutine takes a dictionary object and delimiters for key-value pairs and items, and prints the dictionary
' content in a string format with the specified delimiters.
'
' @param dictObj The dictionary object whose key-value pairs are to be printed.
' @param keyvalue_delimiter The delimiter used to separate keys from values within a pair. Default is "=".
' @param item_delimiter The delimiter used to separate different key-value pairs. Default is ";".
Sub printDictionaryKeyValue(dictObj As Object, Optional keyvalue_delimiter As String = "=", Optional item_delimiter As String = ";")
    Dim key As Variant
    Dim result As String
    
    ' Initialize the result string
    result = ""
    
    ' Loop through each key-value pair in the dictionary
    For Each key In dictObj
        ' Append the key-value pair to the result string with delimiters
        result = result & key & keyvalue_delimiter & dictObj(key) & item_delimiter
    Next key
    
    ' Remove the trailing item delimiter
    If Len(result) > 0 Then
        result = left(result, Len(result) - Len(item_delimiter))
    End If
    
    ' Print the result
    Debug.Print result
End Sub

Function dictionaryToString(dictObj As Object, Optional keyvalue_delimiter As String = "=", Optional item_delimiter As String = ";")
    Dim key As Variant
    Dim result As String
    
    ' Initialize the result string
    result = ""
    
    ' Loop through each key-value pair in the dictionary
    For Each key In dictObj
        ' Append the key-value pair to the result string with delimiters
        result = result & key & keyvalue_delimiter & dictObj(key) & item_delimiter
    Next key
    
    ' Remove the trailing item delimiter
    If Len(result) > 0 Then
        result = left(result, Len(result) - Len(item_delimiter))
    End If
    
    ' Return the result
    dictionaryToString = result
End Function

' This function inverts the keys and values of a given dictionary object.
' If a value is repeated, an error is raised because keys in a dictionary must be unique.
'
' @param dictObj The dictionary object to invert.
' @return A new dictionary object with values as keys and keys as values.
Function invertDictionaryObject(dictObj As Scripting.Dictionary) As Scripting.Dictionary
    Dim invertedDict As Scripting.Dictionary
    Dim key As Variant
    Dim value As Variant
    
    ' Initialize the new dictionary object
    Set invertedDict = New Scripting.Dictionary
    
    ' Loop through each key-value pair in the input dictionary
    For Each key In dictObj.Keys
        value = dictObj(key)
        
        ' Check if the value already exists as a key in the inverted dictionary
        If invertedDict.exists(value) Then
            ' Raise an error if a duplicate key is found
            Err.Raise vbObjectError + 513, "invertDictionaryObject", "Duplicate value found: " & value
        Else
            ' Add the value as a key and the key as a value in the inverted dictionary
            invertedDict.Add value, key
        End If
    Next key
    
    ' Return the inverted dictionary
    Set invertDictionaryObject = invertedDict
End Function

Sub test_dict_functions()
Dim dictObject As Scripting.Dictionary, input_string As String
input_string = "key1=value1;key2=value2"
Set dictObject = dict.getDictionaryFromString(input_string)
dict.printDictionaryKeyValue dictObject

Debug.Assert dictObject.count = 2
Debug.Assert dictObject.item("key1") = "value1"
' test `invertDictionaryObject`
dict.printDictionaryKeyValue dict.invertDictionaryObject(dictObject)
Debug.Assert dict.invertDictionaryObject(dictObject).item("value1") = "key1"

End Sub


