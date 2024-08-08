'filesystem
'1 utilities
'2 vb modules

Sub tests_fs()
    '1 utilities
    Debug.Assert pathExist(ThisWorkbook.path) = True
    Debug.Assert pathExist("fakepath") = False

    Dim testpaths As String
    Dim testpath As String
    
    ' Set test paths including the valid path ThisWorkbook.path
    testpath = ThisWorkbook.path
    testpaths = "fakepath1;fakepath2;" & testpath & ";fakepath3"
    
    ' Assertion test
    Debug.Assert getFirstValidPath(testpaths) = testpath
    
    Debug.Print "tests_fs completed!"
End Sub

'1 utilities
Function fileIsExcel(fl As File) As Boolean
    fileIsExcel = False
    If (InStr(fl.Type, "Excel") > 0) = True Then
    fileIsExcel = True
    End If
End Function

' Function to check if a given path exists
Function pathExist(path As String) As Boolean
    ' This function checks if the specified path exists.
    '
    ' Parameters:
    ' path - The path to check for existence.
    '
    ' Returns:
    ' True if the path exists, False otherwise.
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    pathExist = fso.FolderExists(path) Or fso.FileExists(path)
End Function

' Function to get the first valid path from a list of paths
Function getFirstValidPath(paths As String, Optional sep As String = ";") As String
    ' This function returns the first valid path from a list of paths separated by a specified delimiter.
    ' If no valid path is found, it raises an error.
    '
    ' Parameters:
    ' paths - A string containing a list of paths separated by the specified delimiter.
    ' sep - (Optional) The delimiter used to separate the paths in the list. Default is ";".
    '
    ' Returns:
    ' The first valid path from the list.
    
    Dim pathArray As Variant
    Dim i As Integer
    
    ' Split the paths string into an array using the specified delimiter
    pathArray = Split(paths, sep)
    
    ' Loop through each path in the array and check if it exists
    For i = LBound(pathArray) To UBound(pathArray)
        If pathExist(Trim(pathArray(i))) Then
            getFirstValidPath = Trim(pathArray(i))
            Exit Function
        End If
    Next i
    
    ' Raise an error if no valid path is found
    Err.Raise vbObjectError + 1, "getFirstValidPath", "No valid path found in the list."
End Function

'2 vb modules
Sub exportModuleCode(module_name As String, path As String, Optional extension As String = "")
    Dim module_code As String
    Dim file_path As String
    Dim file_number As Integer
    
    ' Get the module code
    module_code = GetModuleCode(module_name)
    
    ' Determine the module type and set the default extension if not provided
    module_type = TypeName(ThisWorkbook.VBProject.VBComponents(module_name))
    If extension = "" Then
        If module_type = "VBComponent" Then
            extension = IIf(ThisWorkbook.VBProject.VBComponents(module_name).Type = vbext_ct_ClassModule, "cls", "bas")
        Else
            extension = "txt"
        End If
    End If
        
    ' Create the file path
    file_path = path & "\" & module_name & "." & extension
    
    ' Open the file for writing
    file_number = FreeFile
    Open file_path For Output As file_number
    
    ' Write the module code to the file
    Print #file_number, module_code
    
    ' Close the file
    Close file_number
    
    ' Message
    u.printTemplateString "Code module `@1` exported as `@2`", module_name, file_path
End Sub

Function GetModuleCode(module_name As String) As String
    Dim module_code As String
    Dim module_object As Object
    
    ' Get the module object
    Set module_object = ThisWorkbook.VBProject.VBComponents(module_name).codeModule
    
    ' Get the module code
    module_code = module_object.Lines(1, module_object.CountOfLines)
    
    ' Return the module code
    GetModuleCode = module_code
End Function


' This subroutine takes a semicolon-separated string of module names and exports the code for each module.
'
' @param module_names A semicolon-separated string containing the names of the modules to export.
' @param path The file path where the exported code should be saved.
Sub exportModuleCodes(module_names As String, path As String, Optional extension As String = "")
    Dim moduleNameArray As Variant
    Dim i As Integer
    
    ' Split the module_names string into an array using ";" as the delimiter
    moduleNameArray = Split(module_names, ";")
    
    ' Loop through each module name in the array and export its code
    For i = LBound(moduleNameArray) To UBound(moduleNameArray)
        exportModuleCode CStr(moduleNameArray(i)), path, extension:=extension
    Next i
End Sub

' This subroutine takes a semicolon-separated string of module file names and imports each module file from the specified directory.
'
' @param module_files A semicolon-separated string containing the names of the module files to import.
' @param path The file path from where the module files should be imported.
Sub importModules(module_files As String, path As String)
    Dim moduleNameArray As Variant
    Dim i As Integer
    Dim file_path As String
    Dim file_content As String
    Dim module_name As String
    Dim file_number As Integer
    
    ' Split the module_files string into an array using ";" as the delimiter
    moduleNameArray = Split(module_files, ";")
    
    ' Loop through each module file name in the array and import its content
    For i = LBound(moduleNameArray) To UBound(moduleNameArray)
        ' Create the full file path
        file_path = path & "\" & moduleNameArray(i)
        
        ' Extract the module name from the file name
        module_name = left(moduleNameArray(i), InStrRev(moduleNameArray(i), ".") - 1)
        
        ' Open the file for reading
        file_number = FreeFile
        Open file_path For Input As file_number
        
        ' Read the entire content of the file
        file_content = Input$(LOF(file_number), file_number)
        
        ' Close the file
        Close file_number
        
        ' Add a new module to the VBProject and set its content
        With ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
            .name = module_name
            .codeModule.AddFromString file_content
        End With
        
        ' Message
        Debug.Print "Module '" & module_name & "' imported from '" & file_path & "'."
    Next i
End Sub


