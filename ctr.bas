' 1. controls: buttons, etc
' 2. message boxes and windows
Option Explicit

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long

' 1. controls: buttons, etc
Sub add_button(btn_name As String, Optional addr As String = "$A$1", Optional h As Long = 50, Optional w As Long = 100, Optional ws As Worksheet, Optional wb As Workbook, _
                Optional overwrite As Boolean = True, Optional label As String = "Button", Optional placement As xlplacement = xlFreeFloating, Optional event_handler_function As String = "", _
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
    Debug.Print "button left " & left
    Set btn = ws0.Buttons.Add(left:=left, Top:=ws0.Range(addr).Top, Width:=w, Height:=h)
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
        .Top = Range(addr).Top
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


()
