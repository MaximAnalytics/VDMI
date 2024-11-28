'1. utilities

Sub test_os()
    Debug.Assert os.pathJoin("D:", "home") = "D:\home"
    
      ' Test pathSplit function
    Dim parts As Variant
    parts = pathSplit("D:\test")
    Debug.Assert parts(0) = "D:\" And parts(1) = "test"
    
    ' Test isFile function
    Debug.Assert isFile(ThisWorkbook.FullName) = True
    Debug.Assert isFile(ThisWorkbook.path) = False
    
    ' Test isDir function
    Debug.Assert isDir(ThisWorkbook.path) = True
    Debug.Assert isDir(ThisWorkbook.FullName) = False
    
    ' Test getDesktopPath function
    Debug.Print "Desktop Path: " & getDesktopPath()
    
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


' This function splits a path into a tuple (head, tail) where tail is the last part of the path
' and head is everything leading up to that. It mimics the behavior of Python's os.path.split function.
'
' @param path - The path to split.
' @return A variant array where the first element is the head and the second element is the tail.
Function pathSplit(path As String) As Variant
    Dim fso As Object
    Dim folderPath As String
    Dim fileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get the folder path and file name
    folderPath = fso.GetParentFolderName(path)
    fileName = fso.GetFileName(path)
    
    ' Return the result as a variant array
    pathSplit = Array(folderPath, fileName)
End Function

' This function checks if a given path is a file.
' It mimics the behavior of Python's os.path.isfile function.
'
' @param path - The path to check.
' @param raise_error - (Optional) If True, raises an error if the path is not a file. Default is False.
' @return True if the path is a file, False otherwise.
Function isFile(path As String, Optional raise_error As Boolean = False) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    isFile = fso.FileExists(path)
    If raise_error And Not isFile Then
        Err.Raise vbObjectError + 1, "isFile", "Path is not a file: " & path
    End If
End Function

' This function checks if a given path is a directory.
' It mimics the behavior of Python's os.path.isdir function.
'
' @param path - The path to check.
' @param raise_error - (Optional) If True, raises an error if the path is not a directory. Default is False.
' @return True if the path is a directory, False otherwise.
Function isDir(path As String, Optional raise_error As Boolean = False) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    isDir = fso.FolderExists(path)
    If raise_error And Not isDir Then
        Err.Raise vbObjectError + 1, "isDir", "Path is not a directory: " & path
    End If
End Function

' This function returns the path to the desktop on a Windows computer.
'
' @return A string representing the path to the desktop.
Function getDesktopPath() As String
    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")
    
    getDesktopPath = wshShell.SpecialFolders("Desktop")
End Function

' This function returns the path to the Documents folder on a Windows computer.
'
' @return A string representing the path to the Documents folder.
Function getDocumentsPath() As String
    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")
    
    getDocumentsPath = wshShell.SpecialFolders("MyDocuments")
End Function


