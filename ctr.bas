' 1. controls: buttons, etc
' 2. message boxes and windows
Option Explicit

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    
' 0 tests
Sub test_controls_functions()
    ctr.test_formButtonFunctions
    
    ctr.test_cmdButtonFunctions
End Sub
    
' 1. controls: buttons, etc
Sub add_button(btn_name As String, Optional addr As String = "$A$1", Optional h As Long = 50, Optional w As Long = 100, Optional ws As Worksheet, Optional wb As Workbook, _
                Optional overwrite As Boolean = True, Optional label As String = "Button", Optional placement As XlPlacement = xlFreeFloating, Optional event_handler_function As String = "", _
                Optional left_offset = 0)
    Dim ws0 As Worksheet, wb0 As Workbook, btn As Button
    'Get the default worksheet, workbook
    Set ws0 = r.get_default_ws(ws)
    Set wb0 = r.get_default_wb(wb)

    'Check if the button already exists and remove it if overwrite is True
    If overwrite = False Then
        For Each btn In ws0.Buttons
            If btn.name = btn_name Then
                Debug.Print "button exists, dont overwrite"
                Exit Sub
            End If
        Next btn
    Else
        On Error Resume Next
        ws0.Buttons(btn_name).Delete
        On Error GoTo 0
    End If
    
    'Create the button and set its properties
    Dim left As Long
    left = ws0.Range(addr).left + left_offset
    Set btn = ws0.Buttons.Add(left:=left, top:=ws0.Range(addr).top, width:=w, Height:=h)
    btn.name = btn_name
    btn.Text = label
    btn.placement = placement
    
    'Add the event handler
    If event_handler_function <> "" Then
        Dim code As String
        code = "Private Sub " & btn_name & "_Click()" & vbCrLf
        code = code & "    " & event_handler_function & vbCrLf
        code = code & "End Sub"
        With wb0.VBProject.VBComponents(ws0.CodeName).codeModule
            .InsertLines .CountOfLines + 1, code
        End With
    End If
End Sub

Sub positionButton(btn_name As String, Optional addr As String = "$A$1", Optional ws As Worksheet, Optional left_offset = 0)
   Dim left0 As Long, top0 As Long, btn As Shape
   Set ws = r.get_default_ws(ws)
   Set btn = getButton(btn_name, ws:=ws)
   left0 = ws.Range(addr).left + left_offset
   top0 = ws.Range(addr).top
   With btn
    .left = left0
    .top = top0
   End With
End Sub


Function remove_button(btn_name As String, Optional ws As Worksheet, Optional wb As Workbook)
    Dim ws0 As Worksheet, wb0 As Workbook
    'Get the default worksheet, workbook
    Set ws0 = r.get_default_ws(ws)
    Set wb0 = r.get_default_wb(wb)
    'Check if the button exists and remove it
    On Error Resume Next
    ws0.Buttons(btn_name).Delete
    On Error GoTo 0
End Function

Sub move_button(btn_name As String, addr As String, ws0 As Worksheet, Optional left_offset As Long)
    Dim btn As Button
    
    'default left offset
    If IsMissing(left_offset) Then
       'left_offset = 0.25 * main.BTN_WIDTH
    End If
    Set btn = ws0.Buttons(btn_name)
    
    With btn
        .top = Range(addr).top
        .left = Range(addr).left + left_offset
    End With
End Sub


' 2. message box and windows
Public Function HasMsgBox() As Boolean
    Dim hwnd As LongPtr
    
    ' Find the window with the specified title (the message text)
    hwnd = FindWindow("#32770", vbNullString)  ' #32770 is the class name for a MessageBox
    
    HasMsgBox = (hwnd <> 0)
End Function

Public Function CheckMessageBox(msgText As String) As Boolean
    Dim hwnd As LongPtr
    Dim buffer As String
    Dim textLength As Long
    
    ' Find the window with the specified title (the message text)
    hwnd = FindWindow("#32770", vbNullString)  ' #32770 is the class name for a MessageBox
    
    ' If the window is found
    If hwnd <> 0 Then
        ' Prepare buffer to hold the window's title
        buffer = String(256, vbNullChar)
        
        ' Get the window's title text
        textLength = GetWindowText(hwnd, buffer, Len(buffer))
        
        ' Trim the buffer to the length of the text
        buffer = left(buffer, textLength)
        
        ' Check if the message box text matches the expected message
        If InStr(buffer, msgText) > 0 Then
            CheckMessageBox = True
            Exit Function
        End If
    End If
    
    ' If no match is found
    CheckMessageBox = False
End Function

' 3 Form control buttons

' Function to check if a button exists
Function buttonExists(btn_name As String, Optional raise_error As Boolean = False, Optional ws As Worksheet, Optional wb As Workbook) As Boolean
    Dim ws0 As Worksheet
    Set ws0 = r.get_default_ws(ws)
    
    On Error Resume Next
    buttonExists = Not ws0.Buttons(btn_name) Is Nothing
    On Error GoTo 0
    
    If raise_error And Not buttonExists Then
        Err.Raise vbObjectError + 1, "buttonExists", "Button '" & btn_name & "' does not exist."
    End If
End Function

' Function to get a button
Function getButton(btn_name As String, Optional raise_error As Boolean = True, Optional ws As Worksheet, Optional wb As Workbook) As Button
    Dim ws0 As Worksheet
    Set ws0 = r.get_default_ws(ws)
    
    On Error Resume Next
    Set getButton = ws0.Buttons(btn_name)
    On Error GoTo 0
    
    If raise_error And getButton Is Nothing Then
        Err.Raise vbObjectError + 1, "getButton", "Button '" & btn_name & "' does not exist."
    End If
End Function

' Sub to delete a button
Sub deleteButton(btn_name As String, Optional raise_error As Boolean = False, Optional ws As Worksheet, Optional wb As Workbook)
    Dim ws0 As Worksheet
    Set ws0 = r.get_default_ws(ws)
    
    If buttonExists(btn_name, ws:=ws0) Then
        ws0.Buttons(btn_name).Delete
    ElseIf raise_error Then
        Err.Raise vbObjectError + 1, "deleteButton", "Button '" & btn_name & "' does not exist."
    End If
End Sub

' Function to list all buttons on a worksheet
Function listButtons(Optional ws As Worksheet, Optional wb As Workbook) As collection
    Dim ws0 As Worksheet
    Dim btn As Button
    Dim btnCollection As New collection
    
    Set ws0 = r.get_default_ws(ws)
    
    For Each btn In ws0.Buttons
        btnCollection.Add btn
    Next btn
    
    Set listButtons = btnCollection
End Function

' Function to list all button names on a worksheet
Function listButtonNames(Optional ws As Worksheet, Optional wb As Workbook) As collection
    Dim btnCollection As collection
    Dim btnNames As New collection
    Dim btn As Button
    
    Set btnCollection = listButtons(ws, wb)
    
    For Each btn In btnCollection
        btnNames.Add btn.name
    Next btn
    
    Set listButtonNames = btnNames
End Function

' Sub to create a button
Sub createButton(btn_name As String, caption As String, Optional length As Long = 100, Optional width As Long = 50, Optional assign_macro As String = "", _
                 Optional raise_error As Boolean = True, Optional ws As Worksheet, Optional wb As Workbook, Optional overwrite As Boolean = False)
    Dim ws0 As Worksheet
    Dim btn As Button
    
    Set ws0 = r.get_default_ws(ws)
    
    If buttonExists(btn_name, ws:=ws0) Then
        If overwrite Then
            ctr.deleteButton btn_name, ws:=ws, wb:=wb
        ElseIf raise_error Then
            Err.Raise vbObjectError + 1, "createButton", "Button '" & btn_name & "' already exists."
            Exit Sub
        End If
        
    End If
    
    Set btn = ws0.Buttons.Add(left:=ws0.Range("A1").left, top:=ws0.Range("A1").top, width:=width, Height:=length)
    btn.name = btn_name
    btn.caption = caption
    
    If assign_macro <> "" Then
        btn.OnAction = assign_macro
    End If
End Sub

' Sub to assign a macro to a button
Sub assignMacroToButton(btn_name As String, Optional assign_macro As String = "", Optional overwrite As Boolean = False, Optional raise_error As Boolean = True, _
                        Optional ws As Worksheet, Optional wb As Workbook)
    Dim btn As Button
    Set btn = getButton(btn_name, raise_error, ws, wb)
    
    If Not btn Is Nothing Then
        If btn.OnAction <> "" And Not overwrite Then
            If raise_error Then
                Err.Raise vbObjectError + 1, "assignMacroToButton", "Button '" & btn_name & "' already has a macro assigned."
            End If
            Exit Sub
        End If
        btn.OnAction = assign_macro
    End If
End Sub

' Sub to size a button
Sub sizeButton(btn_name As String, Optional length As Long = 100, Optional width As Long = 50, Optional ws As Worksheet, Optional wb As Workbook)
    Dim btn As Button
    Set btn = getButton(btn_name, True, ws, wb)
    
    If Not btn Is Nothing Then
        btn.width = width
        btn.Height = length
    End If
End Sub

' Sub to position a button
Sub positionFormButton(btn_name As String, Optional top As Long, Optional left As Long, Optional address As String = "", Optional left_offset As Long = 0, _
                   Optional top_offset As Long = 0, Optional ws As Worksheet, Optional wb As Workbook)
    Dim btn As Button
    Dim ws0 As Worksheet
    Set ws0 = r.get_default_ws(ws)
    Set btn = getButton(btn_name, True, ws, wb)
    
    If Not btn Is Nothing Then
        If address <> "" Then
            btn.top = ws0.Range(address).top - top_offset
            btn.left = ws0.Range(address).left - left_offset
        Else
            btn.top = top
            btn.left = left
        End If
    End If
End Sub

'4. Command (activex) buttons
' Function to check if a command button exists
Function cmdButtonExists(btn_name As String, Optional raise_error As Boolean = False, Optional ws As Worksheet, Optional wb As Workbook) As Boolean
    Dim ws0 As Worksheet
    Set ws0 = r.get_default_ws(ws)
    
    On Error Resume Next
    cmdButtonExists = Not ws0.OLEObjects(btn_name) Is Nothing
    On Error GoTo 0
    
    If Not cmdButtonExists And raise_error Then
        Err.Raise vbObjectError + 1, "cmdButtonExists", "Command button '" & btn_name & "' does not exist."
    End If
End Function

' Function to get a command button
Function getCmdButton(btn_name As String, Optional raise_error As Boolean = True, Optional ws As Worksheet, Optional wb As Workbook) As OLEObject
    If cmdButtonExists(btn_name, raise_error, ws, wb) Then
        Set getCmdButton = r.get_default_ws(ws).OLEObjects(btn_name)
    End If
End Function

' Sub to delete a command button
Sub deleteCmdButton(btn_name As String, Optional raise_error As Boolean = False, Optional ws As Worksheet, Optional wb As Workbook)
    If cmdButtonExists(btn_name, raise_error, ws, wb) Then
        r.get_default_ws(ws).OLEObjects(btn_name).Delete
    End If
End Sub

' Function to list all command buttons
Function listCmdButtons(Optional ws As Worksheet, Optional wb As Workbook) As collection
    Dim ws0 As Worksheet, btn As OLEObject, btns As New collection
    Set ws0 = r.get_default_ws(ws)
    
    For Each btn In ws0.OLEObjects
        If TypeName(btn.Object) = "CommandButton" Then
            btns.Add btn
        End If
    Next btn
    
    Set listCmdButtons = btns
End Function

' Function to list all command button names
Function listCmdButtonNames(Optional ws As Worksheet, Optional wb As Workbook) As collection
    Dim btns As collection, btn As OLEObject, btnNames As New collection
    Set btns = listCmdButtons(ws, wb)
    
    For Each btn In btns
        btnNames.Add btn.name
    Next btn
    
    Set listCmdButtonNames = btnNames
End Function

' Sub to create a command button
Sub createCmdButton(btn_name As String, caption As String, Optional length As Long = 100, Optional width As Long = 50, _
                    Optional assign_macro As String = "", Optional raise_error As Boolean = True, Optional ws As Worksheet, Optional wb As Workbook, Optional overwrite As Boolean = False)
    Dim ws0 As Worksheet, btn As OLEObject
    Set ws0 = r.get_default_ws(ws)
    
    If cmdButtonExists(btn_name, False, ws, wb) Then
        If overwrite Then
           ctr.deleteCmdButton btn_name, ws:=ws, wb:=wb
        ElseIf raise_error Then
           Err.Raise vbObjectError + 1, "createCmdButton", "Command button '" & btn_name & "' already exists."
           Exit Sub
        End If
    End If
    
    Set btn = ws0.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, _
                                 left:=ws0.Range("A1").left, top:=ws0.Range("A1").top, width:=width, Height:=length)
    With btn
        .name = btn_name
        .Object.caption = caption
    End With
    
    If assign_macro <> "" Then
        ctr.assignMacroToCmdButton btn_name, assign_macro:=assign_macro, ws:=ws, wb:=wb
    End If
End Sub

' Sub to assign a macro to a command button
Sub assignMacroToCmdButton(btn_name As String, Optional assign_macro As String = "", Optional overwrite As Boolean = False, _
                           Optional raise_error As Boolean = True, Optional ws As Worksheet, Optional wb As Workbook)
    Dim btn As OLEObject
    Dim sheetModuleName As String
    Dim clickEventCode As String, wsName As String, clickHandlerName As String
    Set btn = getCmdButton(btn_name, raise_error:=True, ws:=ws, wb:=wb)
    
    ' Find the VB module of the sheet
    wsName = btn.Parent.name
    sheetModuleName = fs.findModuleName(wsName)
    
    'Get the click handler name: BtnName_Click()
    clickHandlerName = CStr(btn.name) & "_Click()"
    
    ' Write the code lines for the button click event
    clickEventCode = "Private Sub " & clickHandlerName & vbCrLf
    clickEventCode = clickEventCode & "    Call " & assign_macro & "()" & vbCrLf
    clickEventCode = clickEventCode & "End Sub"
    
    ' Add the click event procedure to the sheet module
    fs.addProcedureToModule clickEventCode, clickHandlerName, sheetModuleName
    
End Sub

Sub test()
    Dim btn As OLEObject
    Set btn = getCmdButton("btn_update_isah_data_LN 1", raise_error:=True)
    Debug.Print btn.ZOrder
    
End Sub

' Sub to size a command button
Sub sizeCmdButton(btn_name As String, Optional length As Long = 100, Optional width As Long = 50, Optional ws As Worksheet, Optional wb As Workbook)
    Dim btn As OLEObject
    Set btn = getCmdButton(btn_name, True, ws, wb)
    
    If Not btn Is Nothing Then
        With btn
            .width = width
            .Height = length
        End With
    End If
End Sub

' Sub to position a command button
Sub positionCmdButton(btn_name As String, Optional top As Long, Optional left As Long, Optional address As String = "", _
                      Optional left_offset As Long = 0, Optional top_offset As Long = 0, Optional ws As Worksheet, Optional wb As Workbook)
    Dim btn As OLEObject, ws0 As Worksheet
    Set ws0 = r.get_default_ws(ws)
    Set btn = getCmdButton(btn_name, True, ws, wb)
    
    If Not btn Is Nothing Then
        If address <> "" Then
            With ws0.Range(address)
                btn.top = .top + top_offset
                btn.left = .left + left_offset
            End With
        Else
            btn.top = top
            btn.left = left
        End If
    End If
End Sub

' Add the following test code to `ctr.bas`

Sub test_formButtonFunctions()
    Dim ws As Worksheet
    Dim btn_name As String
    Dim caption As String
    Dim btn_macro As String
    Dim btn_address As String
    
    ' Set test parameters
    Set ws = w.get_or_create_worksheet("TestSheet", ThisWorkbook)
    btn_name = "new_button"
    caption = "MyButton"
    btn_macro = "ctr.test_button"
    btn_address = "$B$2"
    
    ' Create button
    createButton btn_name, caption, ws:=ws, overwrite:=True
    Debug.Assert buttonExists(btn_name, ws:=ws) = True
    Debug.Assert getButton(btn_name, ws:=ws).name = btn_name
    Debug.Assert u.InList(btn_name, listButtonNames(ws:=ws)) = True
    
    
    ' Assign macro to button
    assignMacroToButton btn_name, assign_macro:=btn_macro, ws:=ws
    Debug.Print getButton(btn_name, ws:=ws).OnAction
    Debug.Assert Split(getButton(btn_name, ws:=ws).OnAction, "!")(1) = btn_macro

    ' Position button
    positionFormButton btn_name, address:="$B$2", left_offset:=10, top_offset:=10, ws:=ws
    
    ' Delete button
    deleteButton btn_name, ws:=ws
    Debug.Assert buttonExists(btn_name, ws:=ws) = False
End Sub

' Add the following test code to `ctr.bas`

Sub test_cmdButtonFunctions()
    Dim ws As Worksheet
    Dim btn_name As String, caption As String, btn_macro As String, btn_address As String
    
    ' Set test parameters
    Set ws = w.get_or_create_worksheet("TestSheet", ThisWorkbook, False)
    btn_name = "new_cmd_button"
    caption = "MyCmdButton"
    btn_macro = "ctr.test_button"
    btn_address = "$D$2"
    
    ' Create a command button
    createCmdButton btn_name, caption, ws:=ws, raise_error:=False
    Debug.Assert cmdButtonExists(btn_name, ws:=ws) = True
    Debug.Assert getCmdButton(btn_name, ws:=ws).name = btn_name
    Debug.Assert u.InList(btn_name, listCmdButtonNames(ws)) = True
    
    ' Assign a macro to the command button
    assignMacroToCmdButton btn_name, assign_macro:=btn_macro, ws:=ws
    
    ' Position the command button
    positionCmdButton btn_name, address:=btn_address, left_offset:=10, top_offset:=10, ws:=ws

    ' Delete the command button
    deleteCmdButton btn_name, ws:=ws
    Debug.Assert cmdButtonExists(btn_name, ws:=ws) = False
End Sub

Sub test_button()
    MsgBox "Button clicked!"
End Sub
