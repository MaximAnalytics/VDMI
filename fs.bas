'filesystem
Function fileIsExcel(fl As File) As Boolean
    fileIsExcel = False
    If (InStr(fl.Type, "Excel") > 0) = True Then
    fileIsExcel = True
    End If
End Function

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


