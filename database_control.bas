Sub CreateDropdownWithNamedRange()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim namedRange As Range
    Dim namedRangeValues As Range
    Dim namesRange As Range
    Dim databaseNameRange As Range
    Dim connectionStringsRange As Range
    Dim dropdownList As String
    Dim i As Long
    
    ' Set the worksheet where the dropdown will be created
    Set ws = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME) ' Change "Sheet1" to your actual sheet name
    
    ' Set the target cell for the dropdown
    Set targetCell = ws.Range(main.DATABASE_DROPDOWN_ADDR) ' Change "TargetCellAddress" to the cell where you want the dropdown
    ws.Range(main.DATABASE_DROPDOWN_LABEL_ADDR).value = "Selecteer ISAH database naam"
        
    ' Assuming the first column of the named range contains the 'name' values
    Set namedRange = ws.Range(main.CONNECTION_STRINGS_NAMED_RANGE)
    Set databaseNameRange = r.get_column(namedRange, main.DATABASE_NAME_COLUMN_NAME, ws, ThisWorkbook, 1)
    
    ' Format the connection string range
    format_connection_strings_range main.CONNECTION_STRINGS_NAMED_RANGE
    
    ' Clear any existing data validation from the target cell
    targetCell.Validation.Delete
    
    ' Add data validation to the target cell with the list of 'name' values
    With targetCell.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='" & ws.name & "'!" & databaseNameRange.address
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Sub format_connection_strings_range(range_name As String)
    Dim rng0 As Range
    Set rng0 = r.get_range(range_name)
    r.add_outside_border rng0
    r.add_outside_border rng0.Rows(1)
    r.add_outside_border rng0.columns(1)
    
    Set header0 = r.get_row(rng0, 1, offset_column:=1)
    header0.Interior.Color = WT_HEADER_COLOR
    r.get_column(rng0, 1, offset_row:=1).Interior.Color = WT_IDS_COLOR
    
    Dim rng_values As Range
    Set rng_values = r.get_range_values(rng0)
    rng_values.Interior.Color = WT_VALUES_COLOR
    rng_values.NumberFormat = "General"
    
End Sub


