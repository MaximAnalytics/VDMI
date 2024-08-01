Sub test_vb_functions()

    ' copy modules from workbook wb0 to wb1
    Dim wb0 As Workbook, wb1 As Workbook
    Set wb0 = ThisWorkbook
    Set wb1 = Workbooks("template.xltm")
    
    vb.copyModuleCodes wb0, wb1, "a;fs"
End Sub

' This subroutine copies the VBA code from one worksheet module to another within the same workbook.
'
' @param sourceSheetName The name of the source worksheet whose module code is to be copied.
' @param targetSheetName The name of the target worksheet where the code will be pasted.
Sub CopyWorksheetCode(sourceSheetName As String, targetSheetName As String)
    Dim sourceModule As Object
    Dim targetModule As Object
    Dim codeContent As String
    Dim sheetName As String
    Dim ws0 As Worksheet
    
    ' Get the module object for the source worksheet
    Set ws0 = ThisWorkbook.Sheets(sourceSheetName)
    Set sourceModule = ThisWorkbook.VBProject.VBComponents(ws0.CodeName).codeModule
    
    ' Get the module object for the target worksheet
    Set ws0 = ThisWorkbook.Sheets(targetSheetName)
    Set targetModule = ThisWorkbook.VBProject.VBComponents(ws0.CodeName).codeModule
    
    ' Store the entire content of the source module
    codeContent = sourceModule.Lines(1, sourceModule.CountOfLines)
    
    ' Clear any existing code in the target module
    targetModule.DeleteLines 1, targetModule.CountOfLines
    
    ' Paste the code into the target module
    targetModule.AddFromString codeContent
    
    ' Display the user that the operation is complete
    Debug.Print "Code has been copied from " & sourceSheetName & " to " & targetSheetName, vbInformation, "Code Copied"
End Sub

' This method copies a code module from one workbook to another.
' Parameters:
'   wb0 - The source workbook from which the module code will be copied.
'   wb1 - The destination workbook to which the module code will be copied.
'   module_name - The name of the module to be copied.
Sub copyModuleCode(wb0 As Workbook, wb1 As Workbook, module_name As String)
    Dim srcModule As Object
    Dim destModule As Object
    Dim codeString As String
    Dim lineNum As Long
    
    ' Ensure the source workbook has the specified module
    On Error Resume Next
    Set srcModule = wb0.VBProject.VBComponents(module_name)
    On Error GoTo 0
    
    If srcModule Is Nothing Then
        Debug.Print "Module " & module_name & " not found in the source workbook.", vbExclamation
        Exit Sub
    End If
    
    ' Create a new module in the destination workbook
    Set destModule = wb1.VBProject.VBComponents.Add(vbext_ct_StdModule)
    destModule.name = module_name
    
    ' Copy the code from the source module to the destination module
    codeString = ""
    For lineNum = 1 To srcModule.codeModule.CountOfLines
        codeString = codeString & srcModule.codeModule.Lines(lineNum, 1) & vbCrLf
    Next lineNum
    
    destModule.codeModule.AddFromString codeString
    
    Debug.Print "Module " & module_name & " has been copied successfully.", vbInformation
End Sub

' This method copies VBA modules from one workbook to another.
' Parameters:
'   wb0 - The source workbook from which modules will be copied.
'   wb1 - The destination workbook to which modules will be copied.
'   module_names - A semi-colon separated string of module names to be copied.
Public Sub copyModuleCodes(wb0 As Workbook, wb1 As Workbook, module_names As String)
    Dim moduleArray() As String
    Dim moduleName As Variant
    
    ' Split the module_names string into an array of module names
    moduleArray = Split(module_names, ";")
    
    ' Loop through each module name in the array
    For Each moduleName In moduleArray
        ' Call the copyModuleCode method for each module name
        Call vb.copyModuleCode(wb0, wb1, Trim(moduleName))
    Next moduleName
End Sub
