'filesystem
'1 utilities
'2 vb modules

Sub tests_fs()
    Dim ws As Worksheet, wb As Workbook
    Set wb = ThisWorkbook
    Set ws = w.get_or_create_worksheet("test", wb)

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
    
    ' Test writeFile function
    Dim testFileName As String
    Dim testFilePath As String
    Dim testFileContent As String
    Dim readContent As String
    
    testFileName = "testFile.txt"
    testFilePath = ThisWorkbook.path
    testFileContent = "This is a test file."
    
    ' Write the file
    writeFile testFileContent, testFileName, testFilePath
    
    ' Check if the file exists
    Debug.Assert pathExist(testFilePath & "\" & testFileName) = True
    
    ' Read the content from the file
    readContent = readFile(testFileName, testFilePath)
    
    ' Assertion test to check if written content equals read content
    ' Debug.Assert testFileContent = readContent
    
    ' Delete the test file
    deleteFile testFileName, ThisWorkbook.path
    
    ' Assert that the file has been deleted
    Debug.Assert pathExist(ThisWorkbook.path & "\" & testFileName) = False
    
    ' 2 vb
    Dim testModuleName As String
    Debug.Assert fs.findModuleName("base") = "Sheet8" And fs.findModuleName("Sheet8") = "Sheet8"

    testModuleName = "test_module"
    createCodeModule testModuleName, "standard"
    
    ' put some module code
    testFileName = "test_module.bas"
    testFilePath = ThisWorkbook.path
    testFileContent = "'This is a test module"
    fs.putModuleCode testModuleName, testFileContent
    Debug.Assert fs.GetModuleCode(testModuleName) = testFileContent
    
    fs.deleteCodeModule testModuleName
    
    ' now test updateCodeModule
    createCodeModule testModuleName, "standard"
    fs.putModuleCode testModuleName, testFileContent
    fs.exportModuleCode testModuleName, testFilePath
    fs.deleteCodeModule testModuleName
    
    fs.updateCodeModule testFileName, testFilePath
    Debug.Assert fs.moduleExist(testModuleName) = True
    
    ' add/remove procedures from module
    fs.addProcedureToModule "Sub test;debug.print 1;end sub", "test", "test"
    fs.removeProcedureFromModule "test", "test"
    
    'clean up
    fs.deleteCodeModule testModuleName
    fs.deleteFile testFileName, testFilePath
    w.delete_worksheet "test"
    
    Debug.Print "tests_fs completed!"
End Sub

Sub test()
    Dim ws As Worksheet, wb As Workbook
    Set wb = ThisWorkbook
    Set ws = w.get_or_create_worksheet("test", wb)
    
    fs.addProcedureToModule "Sub test;debug.print 1;end sub", "test", "test"
    fs.removeProcedureFromModule "test", "test"
    
    fs.exportModuleCodes "fs.bas;ctr.bas", zz_env.getVDMIGithub()
    
End Sub

'1 utilities
Function fileIsExcel(fl As File) As Boolean
    fileIsExcel = False
    If (InStr(fl.Type, "Excel") > 0) = True Then
    fileIsExcel = True
    End If
End Function

' Function to check if a given path exists
Function pathExist(path As String, Optional raise_error As Boolean = False) As Boolean
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
    If raise_error And Not pathExist Then
       Err.Raise 1001, "pathExist", str.subInStr("Path does not exist: @1", path)
    End If
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

' Function to write content to a file
Public Sub writeFile(file_content As String, file_name As String, path As String)
    ' This subroutine writes the specified content to a file with the given name at the specified path.
    '
    ' Parameters:
    ' file_content - The content to write to the file.
    ' file_name - The name of the file to create.
    ' path - The path where the file should be created.
    
    Dim file_path As String
    Dim file_number As Integer
    
    ' Create the file path
    file_path = path & "\" & file_name
    
    ' Open the file for writing
    file_number = FreeFile
    Open file_path For Output As file_number
    
    ' Write the content to the file
    Print #file_number, file_content
    
    ' Close the file
    Close file_number
    
    ' Message
    Debug.Print "File '" & file_name & "' written to '" & file_path & "'."
End Sub

' Function to read the content of a file
Public Function readFile(file_name As String, path As String) As String
    Dim file_path As String
    Dim file_number As Integer
    Dim file_content As String
    
    ' Create the full file path
    file_path = path & "\" & file_name
    
    ' Check if path exists
    fs.pathExist file_path, True
    
    ' Open the file for reading
    file_number = FreeFile
    Open file_path For Input As file_number
    
    ' Read the entire content of the file
    file_content = Input$(LOF(file_number), file_number)
    
    ' Close the file
    Close file_number
    
    ' Return the file content
    readFile = file_content
End Function

' This subroutine deletes a file from the specified path.
'
' @param file_name The name of the file to delete.
' @param path The path where the file is located.
Public Sub deleteFile(file_name As String, path As String)
    Dim fso As Object
    Dim file_path As String
    
    ' Create the file system object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create the full file path
    file_path = path & "\" & file_name
    
    ' Check if the file exists
    If fso.FileExists(file_path) Then
        ' Delete the file
        fso.deleteFile file_path
        Debug.Print "File '" & file_name & "' deleted from '" & path & "'."
    Else
        Debug.Print "File '" & file_name & "' not found in '" & path & "'."
    End If
End Sub

Public Sub deleteFilePath(file_path As String)
    Dim fso As Object

    ' Create the file system object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the file exists
    If fso.FileExists(file_path) Then
        ' Delete the file
        fso.deleteFile file_path
        Debug.Print "File '" & file_path & "' deleted."
    Else
        Debug.Print "File path '" & file_path & "' not found in."
    End If
End Sub

'2 vb modules
Sub exportModuleCode(module_name As String, path As String, Optional extension As String = "")
    Dim module_code As String, actual_module_name As String, module_extension As String
    Dim file_path As String
    Dim file_number As Integer
    
    module_parts = Split(module_name, ".")
    If UBound(module_parts) > LBound(module_parts) Then
        module_name = module_parts(0)
        extension = module_parts(1)
    End If
    
    ' Get the module code
    actual_module_name = findModuleName(module_name)
    module_code = GetModuleCode(actual_module_name)
    
    ' Determine the module type and set the default extension if not provided
    module_type = TypeName(ThisWorkbook.VBProject.VBComponents(actual_module_name))
    If module_extension = "" Then
        If module_type = "VBComponent" Then
            extension = IIf(ThisWorkbook.VBProject.VBComponents(actual_module_name).Type = vbext_ct_ClassModule, "cls", "bas")
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
    Dim module_code As String, actual_module_name As String
    Dim module_object As Object
    
    ' Get the module object
    actual_module_name = findModuleName(module_name)
    Set module_object = ThisWorkbook.VBProject.VBComponents(actual_module_name).codeModule
    
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
        fs.exportModuleCode CStr(moduleNameArray(i)), path, extension:=extension
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

' This function checks if a module exists in the VB project.
'
' @param module_name The name of the module to check.
' @param raise_error (Optional) If True, raises an error if the module does not exist. Default is False.
' @return True if the module exists, False otherwise.
Public Function moduleExist(module_name As String, Optional raise_error As Boolean = False) As Boolean
    Dim actual_module_name As String
    Dim module_object As Object
    
    ' Get the actual module name using the findModuleName function
    On Error Resume Next
    actual_module_name = fs.findModuleName(module_name)
    On Error GoTo 0
    
    ' Check if the module exists in the VB project
    On Error Resume Next
    Set module_object = ThisWorkbook.VBProject.VBComponents(actual_module_name)
    On Error GoTo 0
    
    If Not module_object Is Nothing Then
        ' Module exists
        moduleExist = True
    Else
        ' Module does not exist
        moduleExist = False
        
        ' Raise an error if raise_error is True
        If raise_error Then
            Err.Raise vbObjectError + 1, "moduleExist", "Module does not exist: " & module_name
        End If
    End If
End Function

' This function tries to find a code module with the given name. If not found, it assumes the name is a sheet name and tries to find the corresponding sheet's module name.
' If neither is found, it raises an error.
'
' @param module_or_sheet_name The name of the module or sheet to find.
' @return The name of the found module.
Public Function findModuleName(module_or_sheet_name As String) As String
    Dim vbComponent As Object
    Dim ws As Worksheet
    
    ' Try to find the module by name
    On Error Resume Next
    Set vbComponent = ThisWorkbook.VBProject.VBComponents(module_or_sheet_name)
    On Error GoTo 0
    
    If Not vbComponent Is Nothing Then
        ' Module found, return its name
        findModuleName = vbComponent.name
        Exit Function
    End If
    
    ' Try to find the sheet by name
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(module_or_sheet_name)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ' Sheet found, return its code module name
        findModuleName = ws.CodeName
        Exit Function
    End If
    
    ' If neither module nor sheet is found, raise an error
    Err.Raise vbObjectError + 1, "findModuleName", "Module or sheet does not exist: " & module_or_sheet_name
End Function

' This subroutine updates the code of an existing module in the VB project.
'
' @param module_name The name of the module to update.
' @param module_code_string The new code to be placed in the module.
Public Sub putModuleCode(module_name As String, module_code_string As String)
    Dim module_object As Object
    Dim code_module As Object
    Dim code_module_name As String
    
    ' Try to get the module object
    code_module_name = findModuleName(module_name)
    Set module_object = ThisWorkbook.VBProject.VBComponents(code_module_name)
    
    ' Get the code module object
    Set code_module = module_object.codeModule
    
    ' Clear the existing code in the module
    code_module.DeleteLines 1, code_module.CountOfLines
    
    ' Add the new code to the module
    code_module.AddFromString module_code_string
    
    ' Message
    Debug.Print "Code for module '" & module_name & "' has been updated."
End Sub

' This subroutine creates a new code module in the VB project.
'
' @param module_name The name of the module to create.
' @param module_type The type of the module to create (e.g., "Standard", "Class").
Public Sub createCodeModule(module_name As String, module_type As String)
    Dim module_object As Object
    Dim module_type_constant As Integer
    
    ' Check if the module already exists
    On Error Resume Next
    Set module_object = ThisWorkbook.VBProject.VBComponents(module_name)
    On Error GoTo 0
    
    ' If the module exists, raise an error
    If Not module_object Is Nothing Then
        Err.Raise vbObjectError + 1, "createCodeModule", "Module '" & module_name & "' already exists in the VB project."
    End If
    
    ' Determine the module type constant based on the provided module_type
    Select Case LCase(module_type)
        Case "standard"
            module_type_constant = vbext_ct_StdModule
        Case "class"
            module_type_constant = vbext_ct_ClassModule
        Case Else
            Err.Raise vbObjectError + 2, "createCodeModule", "Invalid module type: '" & module_type & "'. Valid types are 'Standard' and 'Class'."
    End Select
    
    ' Add the new module to the VB project
    Set module_object = ThisWorkbook.VBProject.VBComponents.Add(module_type_constant)
    module_object.name = module_name
    
    ' Message
    Debug.Print "Module '" & module_name & "' of type '" & module_type & "' has been created."
End Sub

' This subroutine deletes a code module from the VB project.
'
' @param module_name The name of the module to delete.
Public Sub deleteCodeModule(module_name As String)
    Dim actual_module_name As String
    Dim module_object As Object
    
    ' Find the actual module name using the findModuleName function
    actual_module_name = fs.findModuleName(module_name)
    
    ' Get the module object
    Set module_object = ThisWorkbook.VBProject.VBComponents(actual_module_name)
    
    ' Remove the module from the VB project
    ThisWorkbook.VBProject.VBComponents.Remove module_object
    
    ' Message
    Debug.Print "Module '" & actual_module_name & "' has been deleted."
End Sub

' This subroutine updates or imports a code module from a specified file.
'
' @param module_file_name The name of the module file (including extension) to update or import.
' @param path The file path where the module file is located.
Public Sub updateCodeModule(module_file_name As String, path As String)
    Dim module_name As String, actual_module_name As String
    Dim module_code_string As String
    Dim file_path As String
    Dim file_number As Integer
    Dim module_exists As Boolean
    Dim module_type As String
    
    ' Create the full file path
    file_path = path & "\" & module_file_name
    
    ' Get the module code from file
    module_code_string = fs.readFile(module_file_name, path)
    
    ' Extract the module name from the file name (without extension)
    module_name = left(module_file_name, InStrRev(module_file_name, ".") - 1)
    
    ' Check if the module exists in the current project
    module_exists = fs.moduleExist(module_name, False)
    
    On Error Resume Next
    module_exists = Not ThisWorkbook.VBProject.VBComponents(module_name) Is Nothing
    On Error GoTo 0
    
    ' If the module does not exist, import it
    If Not module_exists Then
        ' Determine the module type from the file extension
        Select Case Right(module_file_name, Len(module_file_name) - InStrRev(module_file_name, "."))
            Case "cls"
                module_type = vbext_ct_ClassModule
            Case "frm"
                module_type = vbext_ct_MSForm
            Case Else
                module_type = vbext_ct_StdModule
        End Select
        
        ' Add a new module to the VBProject
        With ThisWorkbook.VBProject.VBComponents.Add(module_type)
            .name = module_name
        End With
        
        ' Set the module code
        fs.putModuleCode module_name, module_code_string
        
        ' Message
        Debug.Print "Module '" & module_name & "' imported from '" & file_path & "'."
    Else
        ' If the module exists, update its code with the lines from the module file
        
        ' Update the module code
        actual_module_name = fs.findModuleName(module_name)
        fs.putModuleCode actual_module_name, module_code_string
        
        ' Message
        Debug.Print "Module '" & actual_module_name & "' updated with code from '" & file_path & "'."
    End If
End Sub

' This subroutine updates or imports multiple code modules from specified files.
'
' @param module_files A semicolon-separated string containing the names of the module files to update or import.
' @param path The file path where the module files are located.
Public Sub updateCodeModules(module_files As String, path As String)
    Dim moduleNameArray As Variant
    Dim i As Integer

    ' Split the module_files string into an array using ";" as the delimiter
    moduleNameArray = Split(module_files, ";")
    
    ' Loop through each module file name in the array and update or import its content
    For i = LBound(moduleNameArray) To UBound(moduleNameArray)
       fs.updateCodeModule CStr(moduleNameArray(i)), path
    Next i
End Sub

' This subroutine removes a procedure (Sub or Function) from a specified module in the VB project.
'
' @param proc_name The name of the procedure to remove.
' @param module_name The name of the module from which the procedure should be removed.
Public Sub removeProcedureFromModule(proc_name As String, module_name As String)
    Dim module_object As Object
    Dim code_module As Object
    Dim start_line As Long
    Dim end_line As Long
    Dim line_text As String
    Dim found As Boolean
    
    ' Get the module object
    Set module_object = ThisWorkbook.VBProject.VBComponents(fs.findModuleName(module_name))
    Set code_module = module_object.codeModule
    
    ' Initialize variables
    found = False
    start_line = 1
    
    ' Loop through each line in the module to find the procedure
    Do While start_line <= code_module.CountOfLines
        line_text = Trim(code_module.Lines(start_line, 1))
        
        ' Check if the line contains the procedure definition
        If (LCase(left(line_text, 3)) = "sub" Or LCase(left(line_text, 8)) = "function" Or LCase(left(line_text, 11)) = "private sub" Or LCase(left(line_text, 16)) = "private function") And InStr(line_text, proc_name) > 0 Then
            found = True
            Exit Do
        End If
        
        start_line = start_line + 1
    Loop
    
    ' If the procedure is found, determine the end line and delete the lines
    If found Then
        end_line = start_line
        
        ' Find the end of the procedure
        Do While end_line <= code_module.CountOfLines
            line_text = Trim(code_module.Lines(end_line, 1))
            If LCase(line_text) = "end sub" Or LCase(line_text) = "end function" Then
                Exit Do
            End If
            end_line = end_line + 1
        Loop
        
        ' Delete the procedure lines
        code_module.DeleteLines start_line, end_line - start_line + 1
        Debug.Print "Procedure '" & proc_name & "' removed from module '" & module_name & "'."
    Else
        Debug.Print "Procedure '" & proc_name & "' not found in module '" & module_name & "'."
    End If
End Sub

' This subroutine adds a procedure to a specified module in the VB project.
' It first removes any existing procedure with the same name, then adds the new procedure code.
'
' @param codelines A string containing the code lines of the procedure, separated by the specified separator.
' @param proc_name The name of the procedure to add.
' @param module_name The name of the module to which the procedure should be added.
' @param sep (Optional) The separator used to split the code lines. Default is ";".
Public Sub addProcedureToModule(codelines As String, proc_name As String, module_name As String, Optional sep As String = ";")
    Dim module_object As Object
    Dim code_module As Object
    Dim code_array As Variant
    Dim i As Integer
    
    ' Remove the existing procedure if it exists
    removeProcedureFromModule proc_name, module_name
    
    ' Get the module object
    Set module_object = ThisWorkbook.VBProject.VBComponents(fs.findModuleName(module_name))
    Set code_module = module_object.codeModule
    
    ' Split the code lines into an array using the specified separator
    code_array = Split(codelines, sep)
    
    ' Check if there is already code in the module
    If code_module.CountOfLines > 0 Then
        ' Add an additional line break before inserting the new procedure
        code_module.InsertLines code_module.CountOfLines + 1, "" ' NEW: Add line break
    End If
    
    ' Add the new procedure code to the module
    For i = LBound(code_array) To UBound(code_array)
        code_module.InsertLines code_module.CountOfLines + 1, Trim(code_array(i))
    Next i
    
    Debug.Print "Procedure '" & proc_name & "' added to module '" & module_name & "'."
End Sub

' This function retrieves the code lines of a specified procedure from a given module.
'
' @param proc_name The name of the procedure to retrieve.
' @param module_name The name of the module from which to retrieve the procedure.
' @return A string containing the code lines of the procedure, separated by semicolons.
Public Function getProcedureCode(proc_name As String, module_name As String, Optional trim_code As Boolean = True) As String 'NEW
    Dim module_object As Object
    Dim code_module As Object
    Dim start_line As Long
    Dim end_line As Long
    Dim line_text As String
    Dim found As Boolean
    Dim code_lines As String
    
    ' Get the module object
    Set module_object = ThisWorkbook.VBProject.VBComponents(fs.findModuleName(module_name))
    Set code_module = module_object.codeModule
    
    ' Initialize variables
    found = False
    start_line = 1
    code_lines = ""
    
    ' Loop through each line in the module to find the procedure
    Do While start_line <= code_module.CountOfLines
        line_text = Trim(code_module.Lines(start_line, 1))
        
        ' Check if the line contains the procedure definition
        If (LCase(left(line_text, 3)) = "sub" Or LCase(left(line_text, 8)) = "function" Or LCase(left(line_text, 11)) = "private sub" Or LCase(left(line_text, 16)) = "private function") And InStr(line_text, proc_name) > 0 Then
            found = True
            Exit Do
        End If
        
        start_line = start_line + 1
    Loop
    
    ' If the procedure is found, determine the end line and collect the lines
    If found Then
        end_line = start_line
        
        ' Find the end of the procedure
        Do While end_line <= code_module.CountOfLines
            If trim_code Then
               line_text = Trim(code_module.Lines(end_line, 1))
            Else
               line_text = code_module.Lines(end_line, 1)
            End If
            code_lines = code_lines & line_text & vbCrLf
            If LCase(line_text) = "end sub" Or LCase(line_text) = "end function" Then
                Exit Do
            End If
            end_line = end_line + 1
        Loop
        
        ' Remove the trailing semicolon
        If Right(code_lines, 1) = ";" Then
            code_lines = left(code_lines, Len(code_lines) - 1)
        End If
    Else
        Err.Raise vbObjectError + 1, "getProcedureCode", "Procedure '" & proc_name & "' not found in module '" & module_name & "'."
    End If
    
    ' Return the code lines
    getProcedureCode = code_lines
End Function

' This subroutine copies a procedure from one module to another.
'
' @param proc_name The name of the procedure to copy.
' @param src_module The name of the source module containing the procedure.
' @param trg_module The name of the target module where the procedure should be copied.
Public Sub copyProcedureCode(proc_name As String, src_module As String, trg_module As String) 'NEW
    Dim codelines As String
    
    ' Get the procedure code lines from the source module
    codelines = getProcedureCode(proc_name, src_module, False)
    
    ' Remove and add the procedure code lines in the target module
    addProcedureToModule codelines, proc_name, trg_module
End Sub
