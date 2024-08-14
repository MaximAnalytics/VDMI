Private Sub Worksheet_Change(ByVal Target As Range)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim connectionStrings As Range
    Set connectionStrings = ThisWorkbook.Names(main.CONNECTION_STRINGS_NAMED_RANGE).RefersToRange
    If Not Intersect(Target, Me.Range(main.DATABASE_DROPDOWN_ADDR)) Is Nothing Then
        Dim selectedName As String
        Dim i As Long
        selectedName = Target.value
        For i = 1 To connectionStrings.Rows.count
            If connectionStrings.Cells(i, 1).value = selectedName Then
                ' Do something with the connection string
                ' For example, store it in another cell
                Me.Range(main.SELECTED_CONNECTION_STRING_ADDR).value = connectionStrings.Cells(i, 2).value
                Me.Range(main.SELECTED_DATABASE_NAME_ADDR).value = connectionStrings.Cells(i, 3).value
                Exit For
            End If
        Next i
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
