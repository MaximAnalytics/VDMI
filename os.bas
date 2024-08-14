'1. utilities

Sub test_os()
    Debug.Assert os.pathJoin("D:", "home") = "D:\home"
End Sub

' operating system functions like Py
Function getcwd() As String
    getcwd = ThisWorkbook.path
End Function

' This function joins multiple path components into a single path.
' It mimics the behavior of Python's os.path.join function.
'
' @param ParamArray paths() - An array of path components to join.
' @return A string representing the joined path.
Function pathJoin(ParamArray paths() As Variant) As String
    Dim i As Integer
    Dim joinedPath As String
    Dim pathComponent As String
    
    ' Initialize the joined path as an empty string
    joinedPath = ""
    
    ' Loop through each path component in the ParamArray
    For i = LBound(paths) To UBound(paths)
        ' Get the current path component
        pathComponent = CStr(paths(i))
        
        ' Remove any leading or trailing path separators from the component
        pathComponent = Trim(pathComponent)
        If Right(pathComponent, 1) = Application.PathSeparator Then
            pathComponent = left(pathComponent, Len(pathComponent) - 1)
        End If
        If left(pathComponent, 1) = Application.PathSeparator Then
            pathComponent = Mid(pathComponent, 2)
        End If
        
        ' Append the path component to the joined path
        If joinedPath = "" Then
            joinedPath = pathComponent
        Else
            joinedPath = joinedPath & Application.PathSeparator & pathComponent
        End If
    Next i
    
    ' Return the joined path
    pathJoin = joinedPath
End Function

