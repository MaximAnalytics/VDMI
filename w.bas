' Workbook, worksheet functions and subs
' 1. worksheet getters, subsetters, generators
'`get_or_create_worksheet(wsname, wb, overwrite, clear)` => Creates or retrieves a worksheet with the given name, optionally overwriting or clearing it.
'`getWorksheet(ws_name, wb)` => Retrieves a worksheet by name or reference from the given workbook, raising an error if not found.
'`delete_worksheet(wsname, wb)` => Deletes the worksheet with the specified name from the given workbook.
'`copy_ws(wsname, new_wsname)` => Creates a copy of the specified worksheet, optionally renaming the new worksheet.
'`move_ws(wsname, before_ws, after_ws, wb)` => Moves a worksheet within the workbook to a new position, specified before or after another sheet.

' 2. worksheet logical, utility functions
'`sheet_exists(ws, wb)` => Checks if a worksheet exists in the specified workbook, returning a boolean result.

' 3. worksheet clearing
'`clearWorksheet(ws_name, wb)` => Clears contents, formats, and conditional formatting from the specified worksheet.

' 4. worksheet rows, columns functions and protect sheet
'`subset_columns(ws, column_indexes)` => Returns a range object representing a subset of columns from the specified worksheet.
'`freeze_top_rows(ws, n)` => Freezes the top 'n' rows of the specified worksheet for easier viewing.
'`protect_sheet(wsname)` => Protects the worksheet with the given name, preventing changes.
'`test_protect_sheets()` => Unprotects all sheets in the active workbook.

' 5. worksheet sorting
'`order_sheets(sheetNames())` => Reorders worksheets in the active workbook according to the provided array of sheet names.
'`orderSheets(sheetNamesCol)` => Reorders worksheets in the active workbook based on the order of names in the given collection.

' 6. workbook generators, getters
'`create_empty_workbook(new_workbook_name, path)` => Creates a new empty workbook with a single sheet, saving it to the specified path.
'`test_create_empty_workbook()` => Tests the creation of an empty workbook with a default name and saves it to the current directory.

Sub test_w_functions()
    ' 1. workbook
    Dim wbname As String
    wbname = "template"
    w.createMacroEnabledTemplate wbname, zz_env.VDMI_DATAPATH, False
    
    ' ws_copy
    w.copy_ws "base", "base2", True
    Debug.Assert w.sheet_exists("base2")
    w.delete_worksheet "base2"
    Debug.Assert Not w.sheet_exists("base2")
End Sub

' 1. workbook
Function get_or_create_worksheet(wsName As String, wb As Workbook, Optional overwrite As Boolean = False, Optional clear As Boolean) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Worksheets(wsName)
    If overwrite Then
      Application.DisplayAlerts = False
      ws.Delete
      Set ws = Nothing
      Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.name = wsName
    End If
    
    If clear Then
      r.clear_formatting ws.Cells
    End If
    
    Set get_or_create_worksheet = ws
End Function

Sub delete_worksheet(wsName, Optional wb As Workbook)
Dim ws As Worksheet
Set wb = r.get_default_wb(wb)
current_wsname = ActiveSheet.name
On Error Resume Next
    Set ws = wb.Worksheets(wsName)
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
On Error GoTo 0
' return to the active sheet if this is not the deleted sheet
If current_wsname <> wsName Then
    wb.Sheets(current_wsname).Activate
End If
End Sub

Function copy_ws(wsName As String, Optional new_wsname As String, Optional overwrite As Boolean = False) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(wsName)
    If new_wsname = "" Then
       'new_wsname = wsName & " (copy)"
    End If
    ws.Copy Before:=ThisWorkbook.Worksheets(1)
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Visible = xlSheetVisible
    If overwrite And w.sheet_exists(new_wsname, ThisWorkbook) Then
       w.delete_worksheet new_wsname, ThisWorkbook
    End If
    If new_wsname <> "" Then
       ws.name = new_wsname
    End If
    Set copy_ws = ws
End Function

Sub move_ws(wsName, Optional before_ws = "", Optional after_ws = "", Optional wb)
    Dim wb0 As Workbook
    Dim ws As Worksheet
    Dim before_ws_index As Integer
    Dim after_ws_index As Integer
    
    ' Set default workbook if not provided
    Set wb0 = r.get_default_wb(wb:=wb)
    
    ' Check if worksheet with given name exists
    On Error Resume Next
    Set ws = wb0.Sheets(wsName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Err.Raise 1002, , "Worksheet with name '" & wsName & "' does not exist."
        Exit Sub
    End If
    
    ' Check if before_ws is provided and exists
    If Not before_ws = "" Then
        On Error Resume Next
        Set ws = wb0.Sheets(before_ws)
        On Error GoTo 0
        
        If ws Is Nothing Then
            Err.Raise 1001, , "Worksheet with name '" & before_ws & "' does not exist."
            Exit Sub
        Else
            before_ws_index = ws.index
        End If
    End If
    
    ' Check if after_ws is provided and exists
    If Not after_ws = "" Then
        On Error Resume Next
        Set ws = wb0.Sheets(after_ws)
        On Error GoTo 0
        
        If ws Is Nothing Then
            Err.Raise 1002, , "Worksheet with name '" & after_ws & "' does not exist."
            Exit Sub
        Else
            after_ws_index = ws.index
        End If
    End If
    
    ' Move the worksheet
    If before_ws_index <> 0 Then
        wb0.Sheets(wsName).Move Before:=wb0.Sheets(before_ws_index)
    ElseIf after_ws_index <> 0 Then
        wb0.Sheets(wsName).Move After:=wb0.Sheets(after_ws_index)
    End If
End Sub

Sub order_sheets(ParamArray sheetNames() As Variant)
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim ws As Worksheet
    Dim i As Integer
    Dim wsName As String, prev_ws_index As Long: prev_ws_index = 1
    
    ' Loop through each sheet name in the paramarray
    For i = LBound(sheetNames) To UBound(sheetNames)
        wsName = sheetNames(i)
        Debug.Print "sheet name is:", wsName
        
        ' Check if the sheet exists in the workbook
        On Error Resume Next
            Set ws = wb.Sheets(sheetNames(i))
        On Error GoTo 0
        
        ' Take the first sheet as starting point
        If i = LBound(sheetNames) Then
          GoTo next_i
        End If
        
        ' If the sheet exists, move it to the desired position
        If Not ws Is Nothing Then
            ws.Move After:=wb.Sheets(prev_ws_index)
            prev_ws_index = ws.index
        End If
next_i:
    
    Next i
End Sub

Sub orderSheets(sheetNamesCol As collection)
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim ws As Worksheet
    Dim i As Integer
    Dim wsName As String, prev_ws_index As Long
    Dim sheetNames As Variant
    
    ' Generate sheetNames array from collection
    sheetNames = a.to_array(sheetNamesCol)
    
    ' Loop through each sheet name in the paramarray
    c = 0
    For i = LBound(sheetNames) To UBound(sheetNames)
        
        ' Get the worksheet name
        wsName = sheetNames(i)
        
        ' Check if the sheet exists in the workbook
        On Error Resume Next
            Set ws = wb.Sheets(sheetNames(i))
        On Error GoTo 0
        c = c + 1
        Debug.Print "sheet name is:", wsName

        ' Take the first existing sheet as starting point
        If c = 1 Then
          prev_ws_index = ws.index
          GoTo next_i
        End If
        
        ' If the sheet exists, move it to the desired position
        If Not ws Is Nothing Then
            ws.Move After:=wb.Sheets(prev_ws_index)
            prev_ws_index = ws.index
        End If
next_i:
    Next i
End Sub

' This VBA code defines a method to create a macro-enabled Excel template.
' The method takes three parameters: template_name, path, and an optional closeNewWorkbook flag.

Sub createMacroEnabledTemplate(template_name As String, path As String, Optional closeNewWorkbook As Boolean = False)
    ' Declare variables
    Dim newWorkbook As Workbook
    Dim templateFullPath As String
    
    ' Construct the full path for the new template
    templateFullPath = path & "\" & template_name & ".xltm"
    
    ' Add a new workbook
    Set newWorkbook = Workbooks.Add
    
    ' Save the new workbook as a macro-enabled template
    newWorkbook.SaveAs Filename:=templateFullPath, FileFormat:=xlOpenXMLTemplateMacroEnabled
    
    ' Check if the new workbook should be closed
    If closeNewWorkbook Then
        newWorkbook.Close SaveChanges:=False
    End If
    
    ' Clean up
    Set newWorkbook = Nothing
End Sub


' 2.
Function sheet_exists(ws As Variant, Optional wb As Workbook) As Boolean
    Dim wsName As String
    Dim wb1 As Workbook
    Dim wsTemp As Worksheet
    Dim sheetExist As Boolean

    ' Determine if ws is a string or a worksheet object
    If TypeName(ws) = "String" Then
        wsName = ws
    ElseIf TypeName(ws) = "Worksheet" Then
        wsName = ws.name
    Else
        Err.Raise 1001, , "ws must be a worksheet name or worksheet instance"
    End If

    ' Get the workbook
    If wb Is Nothing Then
        Set wb1 = r.get_default_wb(wb) ' Assuming 'r' is a predefined object with method get_default_wb
    Else
        Set wb1 = wb
    End If

    ' Check if sheet exists
    sheetExist = False
    For Each wsTemp In wb1.Worksheets
        If wsTemp.name = wsName Then
            sheetExist = True
            Exit For
        End If
    Next wsTemp

    sheet_exists = sheetExist
End Function


' 3.
Function getWorksheet(ws_name As Variant, Optional wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Set wb = r.get_default_wb(wb)
    
    If TypeName(ws_name) = "Worksheet" Then
        Set ws = ws_name
    ElseIf VarType(ws_name) = vbString Then
        If sheet_exists(ws_name, wb:=wb) Then
        Set ws = wb.Worksheets(ws_name)
        Else
        Err.Raise 1001, "getWorksheet", "worksheet doesnt exist"
        End If
    Else
        Err.Raise 1002, "getWorksheet", "ws_name is not Worksheet or Worksheet name"
    End If
    
    Set getWorksheet = ws
End Function

Sub clearWorksheet(ws_name As Variant, Optional wb As Workbook)
    Dim ws As Worksheet
    Set ws = getWorksheet(ws_name, wb)
    
    If Not ws Is Nothing Then
        With ws
            .Cells.ClearContents
            .Cells.ClearFormats
            .Cells.FormatConditions.Delete
            .columns.Hidden = False
            .Rows.Hidden = False
        End With
    End If
End Sub

Sub hideWorksheets(ParamArray wsNames() As Variant)
    ' This subroutine hides one or multiple worksheets in ThisWorkbook based on the provided worksheet names.
    '
    ' Parameters:
    ' wsNames - A ParamArray of worksheet names (as strings) to be hidden.
    
    Dim wsName As Variant
    Dim ws As Worksheet
    
    ' Loop through each name provided in the ParamArray
    For Each wsName In wsNames
        ' Check if the worksheet name is a string and not empty
        If VarType(wsName) = vbString And wsName <> "" Then
            ' Check if the worksheet exists in ThisWorkbook
            On Error Resume Next ' Ignore error if worksheet does not exist
            Set ws = ThisWorkbook.Sheets(wsName)
            If Not ws Is Nothing Then
                ' Hide the worksheet if it exists
                ws.Visible = xlSheetHidden
            End If
            Set ws = Nothing ' Reset ws for the next iteration
            On Error GoTo 0 ' Resume normal error handling
        End If
    Next wsName
End Sub

' 4.
Function subset_columns(ws As Worksheet, column_indexes As Variant) As Range
    Dim result_range As Range
    Dim i As Integer
    
    Set result_range = ws.columns(column_indexes(1))
    
    For i = 2 To UBound(column_indexes)
        Set result_range = Union(result_range, ws.columns(column_indexes(i)))
    Next i
    
    Set subset_columns = result_range
End Function

Sub freeze_top_rows(ws As Worksheet, n As Integer)
    ws.Activate
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitRow = n
        .FreezePanes = True
    End With
End Sub

' protect sheet
Sub protect_sheet(wsName As String)
    
    ' Protect the worksheet
    Dim ws As Worksheet
    Set ws = w.get_or_create_worksheet(wsName, ThisWorkbook)
    ws.Protect
        
End Sub

Sub test_protect_sheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
       ws.Unprotect
    Next
End Sub

' 5. workbook getters, subsetters, generators
' create empty workbook
Sub create_empty_workbook(new_workbook_name As String, path As String)
    ' Save current state
    Application.DisplayAlerts = False
    ThisWorkbook.Save 'As ThisWorkbook.path & "\" & ThisWorkbook
        
    ' Add new sheet ws1 and remove all other sheets
    Dim ws As Worksheet, ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets.Add
    For Each ws In ThisWorkbook.Sheets
        If ws.name <> ws1.name Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    
    ' Rename ws1 to Sheet1 and reset worksheets index
    ws1.name = "Sheet1"
    
    ' Remove all named ranges
    Dim n As name
    For Each n In ThisWorkbook.Names
        r.delete_named_range n.name, ws, ThisWorkbook, False
    Next
    
    ' Break all links to external sources
    Dim Links As Variant
    Links = ThisWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    If Not u.is_empty_missing(Links) Then
        For i = LBound(Links) To UBound(Links)
        ThisWorkbook.BreakLink _
            name:=Links(i), _
            Type:=xlLinkTypeExcelLinks
        Next i
    End If
    
    ' Remove all code in ThisWorkbook module
    Dim codeModule As vbide.codeModule
    Set codeModule = ThisWorkbook.VBProject.VBComponents("ThisWorkbook").codeModule
    codeModule.DeleteLines 1, codeModule.CountOfLines
    
    ' Remove all VBA modules starting with "main" or "Module"
    Dim VBComponent As VBComponent
    For Each VBComponent In ThisWorkbook.VBProject.VBComponents
        If left(VBComponent.name, 4) = "main" Or left(VBComponent.name, 6) = "Module" Then
            ThisWorkbook.VBProject.VBComponents.Remove VBComponent
        End If
    Next VBComponent

    ' Save current workbook as path\new_workbook_name
    ThisWorkbook.SaveAs path & "\" & new_workbook_name
    
    Application.DisplayAlerts = True
End Sub

Sub test_create_empty_workbook()
    w.create_empty_workbook "template.xlsm", os.getcwd()
End Sub

