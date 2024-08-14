' Functions and subs

' Named ranges
' format_name(range_name As String) => Formats range name for Excel
' create_named_range(name, sheetName, address, default_value, formula, header_row, id_row, overwrite, expand_range, clear) => Creates/overwrites a named range
' delete_named_range(name, ws, wb, clear) => Deletes a named range
' update_named_range(name, new_range, wb) => Updates reference of a named range
' resize_named_range(name, add_rows, add_cols, wb) => Resizes a named range
' subsetNamedRange(name, startrow, endrow, startcol, endcol, wb) => Creates subset of a named range
' fit_named_range_to_values(name, wb) => Fits named range to non-empty cells
' add_formula_column_to_named_range(name, newColumn, formulaDefinition) => Adds formula column to named range
' name_exist(name, ws, wb) => Checks if named range exists

' Formulas
' fill_formula_range(rng, formulaDefinition, only_values) => Fills range with formula
' test_add_formula() => Test function for adding formula to named range

' Subsetting
' safe_offset(rng0, offset_row, offset_column) => Safely offsets a range
' get_range(rng, ws, wb, offset_row, offset_column) => Gets a range object from rng input: range, string (named range), worksheet
' extend_range(rng, ws, wb, add_rows, add_cols) => Extends a range by adding `add_rows` rows or `add_cols` columns
' expand_range(rng, ws, wb, max_num_cols, c1) => Expands range to last data cell
' getResizedRange(rng, add_rows, add_cols, num_rows, num_cols) => Resizes a range
' subset_range(rng0, startrow, endrow, startcol, endcol) => Subsets a range
' get_last_row(rng, ws, wb, range2) => Gets last row with data
' get_last_col(rng, ws, wb) => Gets last column with data
' get_header(rng, ws, wb) => Gets header row of a range
' get_range_values(rng, ws, wb, offset_row, offset_column) => Gets values of a range
' get_column(rng, index, ws, wb, offset_row) => Gets a column from a range
' get_column_values(rng, index, ws, wb) => Gets values of a column
' select_columns(rng, column_names) => Selects columns from range

' Logical
' range_contains(rng0, x) => Checks if range contains value(s)

' Columns and rows
' get_column_index(rng, column_name) => Gets index of a column by name
' get_column_indexes(rng, column_names) => Gets indexes of multiple columns
' column_exist(rng0, column_name) => Checks if column exists in range
' get_row(rng, index, ws, wb, offset_column) => Gets a row from a range
' get_row_index(rng, row_index) => Gets index of a row
' get_value(rng, r_index, c_index) => Gets value from a cell in range
' set_value(rng, r_index, c_index, value) => Sets value to a cell in range
' get_index_of_value(value, rng) => Gets index of value in range

' Addresses
' get_range_address(ws0, r0, r1, c0, c1) => Gets address of a range
' get_column_address(rng0) => Gets address of a column

' Formatting
' clear_formatting(rng0, ws, wb) => Clears formatting of a range
' clear_range(rng0, ws, wb, clear_formatting) => Clears contents and optionally formatting
' clear_range_values(rng, clear_formatting) => Clears values of a range
' setConditionalFormatting(rng, greenValue, redValue) => Sets conditional formatting
' format_row_column_size(ws, r_height, c_width) => Sets row height and column width
' autofit_columns_rows(rng) => Auto-fits columns and rows
' autofit_columns(rng) => Auto-fits columns
' get_column_formats(rng) => Gets formats of columns
' get_color_index(rng) => Gets color index of range
' add_outside_border(rng) => Adds outside border to range
' add_all_borders(rng) => Adds all borders to range
' get_color_i() => Gets color index of cell
' format_columns =>

' Worksheet, Workbook
' get_default_ws(ws) => Gets default worksheet
' get_default_wb(wb) => Gets default workbook
' protect_sheet(ws, unlock_ranges) => Protects sheet and optionally unlocks ranges

' Transformations
' rng_to_1d_array(rng0) => Converts range to 1D array
' rng_to_2d_array(rng0) => Converts range to 2D array
' get_unique_vals(rng) => Gets unique values from range
' to_column_names(column_names) => Converts to column names array

' Filtering, sorting
' FilterRange(rng, column_name, filter_value) => Filters range by column value
' filter_range(rng, column_name, filter_value, column_name_2, filter_value_2, remove_filter, xl_operator) => Filters range and returns array
' sort_range_by_columns(rng, column_names, sort_order) => Sorts range by columns
' sort_range_by_columns_2(rng, column_names, sort_order) => Sorts range by columns using SortFields

' Copy/Paste
' copy_range(rng0, addr, ws) => Copies range to another location
' paste_array(arrayToCopy, addr, ws) => Pastes array to range

' Utilities
' str_to_array(str0) => Converts string to array
' get_array_len(arr) => Gets length of array

' constants
Global Const MAX_XL_ROWS As Long = 1048576
Global Const FILTER_COLUMN_NAME = "Cap.Grp"

'tests
Sub test_range_functions()
   Dim n As name, rng0 As Range, ws0 As Worksheet, wb0 As Workbook, wstest As Worksheet, rng1 As Range, named_range_name As String, formulaTemplate As String
   Dim numRows As Long, numCols As Long
   Set wb0 = ThisWorkbook
   Set wstest = w.get_or_create_worksheet("test", wb0, True)
   
   ' test `expand_range`
   Set rng0 = r.expand_range("A1", Worksheets("TEST_DATA"))
   Debug.Assert rng0.Rows.count = 311
    
    'filter on single value
    filter_values = Array("LN 1")
    filtered_array = r.filter_range(rng0, r.FILTER_COLUMN_NAME, filter_values, remove_filter:=True, xl_operator:=xlFilterValues)
    Debug.Assert a.num_array_rows(filtered_array) > 1
    
    'filter on multiple values
    filter_values = Array("LN 1", "HCM", "U800")
    filtered_array = r.filter_range(rng0, r.FILTER_COLUMN_NAME, filter_values, remove_filter:=False, xl_operator:=xlFilterValues)
    Debug.Assert a.num_array_rows(filtered_array) > 1
    
    'named ranges
    Dim Column1 As Range, Column2 As Range
    nrows = 100
    named_range_name = "test"
    r.create_named_range named_range_name, wstest.name, "A1:D" & nrows, header_row:="id,val1,val2,val3", overwrite:=True
    address = r.get_range(named_range_name).address
    Debug.Assert address = "$A$1:$D$100"
    
    r.set_column_values "test", "val1", a.to_array(1)
    r.set_column_values "test", "val2", a.create_integer_vector(1, 100 - 1)
    Set Column1 = r.get_column(named_range_name, "val1")
    Set Column2 = r.get_column(named_range_name, "val2")
    Debug.Assert Column1.Cells(2) = 1 And Column2.Cells(3) = 2
    
    'subset named range
    r.subsetNamedRange named_range_name, 1, 1, 1, 1
    Debug.Assert r.get_range(named_range_name).address = "$A$1"
    
    named_range_name = "test2"
    r.create_named_range named_range_name, wstest.name, "F1:G5", header_row:="val4,val5", overwrite:=True, default_value:="A"
    Set rng0 = r.get_range(named_range_name)
    rng0.Select
    
    new_values = a.create_array(5, 5, "B")
    
    Exit Sub
    
    ' Resize the named range to match the dimensions of the values array
    numRows = UBound(new_values, 1)
    numCols = UBound(new_values, 2)
    Set rng1 = r.getResizedRange(rng0, num_rows:=numRows, num_cols:=numCols)
    Debug.Assert rng1.Rows.count = numRows And rng1.columns.count = numCols
    
    ' test updateNamedRangeWithValues
    r.updateNamedRangeWithValues named_range_name, new_values
    Set rng1 = r.get_range(named_range_name)
    Debug.Assert rng1.Rows.count = numRows And rng1.columns.count = numCols And rng1.Cells(1, 1) = new_values(1, 1)
    
    ' insert columns at start, second position and end
    r.add_named_range_column "test", "first_column", pos:=1, values:=a.to_array("A")
    r.add_named_range_column "test", "second_column", pos:=2, values:=a.to_array("B")
    r.add_named_range_column "test", "last_column", pos:=0, values:=a.to_array("Z")
    
    ' remove column id
    r.remove_named_range_column "test", "id"
    
    ' insert formula to add columns val1+val2
    r.add_named_range_column "test", "val1+val2", pos:=0
    formulaTemplate = "=@1+@2"
    ref1 = r.get_column(named_range, "val1").Cells(2, 1).address
    ref2 = r.get_column(named_range, "val2").Cells(2, 1).address
    formulaDef = Replace(str.subInStr(formulaTemplate, ref1, ref2), "$", "")
    
    Set rng0 = r.get_range(named_range)
    Set formulaRange = r.subset_range(r.get_column(named_range, "val1+val2"), 2)
    formulaRange.Select
    r.fill_formula_range formulaRange, formulaDef
    
    ' sorting
    r.sort_range_by_columns_2 rng0, Array("Cap.Grp", "Aantal")
    
End Sub



Sub test_fit_named_range()
r.fit_named_range_to_values "LN_5_orders", searchUpRange:=Range("A999")
Debug.Print r.get_range("LN_5_orders").address
End Sub


Sub test_subsetting()
    Dim rng0 As Range, new_column_name As String
    Set rng0 = r.expand_range("A1", Sheets("TEST_DATA"), ThisWorkbook)
    ' test getColumnIndex
    column_index = r.getColumnIndex(rng0, "Omschrijving")
    Debug.Assert column_index = 3
    
    ' insert before "Omschrijving"
    new_column_name = "NEW"
    r.InsertColumnIntoRange rng0, "Bulkcode", new_column_name, 1
    Debug.Assert rng0.Cells(2, r.getColumnIndex(rng0, new_column_name)) = 1
    
    r.DeleteColumnFromRange rng0, new_column_name
    Debug.Assert r.column_exist(rng0, new_column_name) = False

End Sub

Sub SetColumnValues(rng As Range, Optional values)
    ' This subroutine sets the values of a range excluding the first row (header).
    ' If values is an array, it checks if the number of rows in the range matches the length of the array.
    ' If they match, it loops over the items of the array and sets them to the cells in the range.
    ' If values is not an array, it sets each cell in the range to the provided value.
    
    Dim valuesRng As Range, errstr As String
    Dim i As Long
    
    ' Set valuesRng as rng without the first row (header)
    If rng.Rows.count <= 1 Then
       Exit Sub
    End If
    Set valuesRng = rng.Offset(1, 0).Resize(rng.Rows.count - 1, rng.columns.count)
    
    ' Check if values is an array
    If IsArray(values) Then
        ' Check if the number of rows in valuesRng matches the length of the array
        values = a.ConvertTo1DArray(values)
        num_values = a.array_length(values)
        If valuesRng.Rows.count = num_values Then
            ' Loop over items of values and set to cells in valuesRng
            For i = LBound(values) To UBound(values)
                valuesRng.Cells(i - LBound(values) + 1, 1).value = values(i)
            Next i
        Else
            ' Raise an error if the number of values does not match the number of rows
            errstr = "Number of values. Number of rows: @1, number of values elements: @2"
            errstr = str.subInStr(errstr, valuesRng.Rows.count, num_values)
            Err.Raise vbObjectError + 1, "SetColumnValues", errstr
        End If
    Else
        ' If values is not an array, set each cell in valuesRng to values
        valuesRng.value = values
    End If
End Sub

Sub set_column_values(named_range As String, column_name As String, values As Variant)
    Dim rng As Range
    Dim columnname As Range
    Dim numRows As Long
    Dim i As Long
    
    ' Get the named range
    Set rng = r.get_range(named_range)
    
    ' Check if the column exists in the named range
    If Not r.column_exist(rng, column_name) Then
        Err.Raise 1000, "", "Column not found in named range: " & column_name
        Exit Sub
    End If
    
    ' Get the number of rows in the named range
    numRows = rng.Rows.count - 1
    
    ' If range has no values rows then exit sub
    If numRows < 1 Then
       Exit Sub
    End If
    
    ' Check if values is not array, then set to array
    If Not IsArray(values) Then
       Err.Raise 1001, "", "values is not array"
    ElseIf a.is_2d_array(values) Then
       values = a.ConvertTo1DArray(values)
       'Err.Raise 1002, "", "values is 2d array"
    End If
    
    ' Check the length of the values array
    Set columnRange = r.get_column(rng, column_name)
    If UBound(values) - LBound(values) + 1 = 1 Then
        ' Set the same value in all cells of the column
        columnRange.Offset(1).Resize(numRows).value = values(LBound(values))
    ElseIf UBound(values) - LBound(values) + 1 = numRows Then
        ' Set all values in the column
        Set valuesRange = columnRange.Offset(1).Resize(numRows)
        For i = LBound(values) To UBound(values)
            valuesRange.Cells(i) = values(i)
        Next i
    Else
        ' Raise an error if the number of values mismatched with the number of rows
        Err.Raise vbObjectError + 1, , "Number of values mismatched with number of rows."
    End If
End Sub

Sub add_named_range_column(named_range As String, column_name As String, Optional pos As Long = 0, Optional overwrite As Boolean = True, _
    Optional values As Variant, Optional wb As Variant)
    Dim rng As Range
    Dim headerRow As Range
    Dim newColumn As Range
    Dim newRange As Range
    Dim columnIndex As Long
    Dim orig_address As String
    
    ' Get the named range
    Set rng = r.get_range(named_range, wb:=wb)
    orig_address = rng.address
    
    ' Check if the column already exists in the named range
    If r.column_exist(rng, column_name) Then
        If Not overwrite Then
        Err.Raise "", "", "Column already in named range."
        End If
        Exit Sub
    End If
    
    ' Get the header row
    Set headerRow = rng.Rows(1)
    
    ' Get the column index to insert the new column
    If pos > 0 And pos <= rng.columns.count Then
        ' Insert the new column
        columnIndex = pos
        Set newColumn = headerRow.Cells(1, columnIndex).EntireColumn
        newColumn.Insert shift:=xlToRight
        
        ' if inserting the starting column then resize range to include an additional column
        If columnIndex = 1 Then
           r.update_named_range named_range, r.getResizedRange(r.get_range(orig_address), add_cols:=1)
        End If
        
    ElseIf pos > rng.columns.count Or pos = 0 Then
        ' Append to end
        ' resize named range
        r.update_named_range named_range, r.getResizedRange(rng, add_cols:=1)
        Set headerRow = r.get_header(named_range)
        orig_address = r.get_range(named_range).address
        
        ' Insert the new column
        columnIndex = headerRow.columns.count
        Set newColumn = headerRow.Cells(1, columnIndex).EntireColumn
        newColumn.Insert shift:=xlToRight
        r.update_named_range named_range, orig_address
    Else
        Exit Sub
    End If

    ' Set the header of the new column
    Set headerRow = r.get_range(named_range) '.Rows(1)
    headerRow.Cells(columnIndex).value = column_name
    
    ' Optional insert values
    If Not u.is_empty_missing(values) Then
      r.set_column_values named_range, column_name, values
    End If
    
    
End Sub

Sub remove_named_range_column(named_range As String, column_name As String)
    Dim rng As Range
    Dim headerRow As Range
    Dim columnRange As Range
    Dim newRange As Range
    
    ' Get the named range
    Set rng = r.get_range(named_range)
    
    ' Check if the column exists in the named range
    If Not r.column_exist(rng, column_name) Then
        Err.Raise 1000, "", "Column not found in named range."
        Exit Sub
    End If
    
    ' Get the header row
    Set headerRow = rng.Rows(1)
    
    ' Get the column range
    Set columnRange = headerRow.Find(column_name)
    
    ' Delete the column, named range gets updated automatically
    columnRange.EntireColumn.Delete
    
    ' Update the named range
    'Set newRange = rng.Resize(rng.Rows.count, rng.columns.count - 1)
    ' r.update_named_range named_range, newRange
End Sub

' range functions
Function format_name(range_name As String) As String
    Dim new_name As String
    new_name = range_name
    'Replace spaces with underscores
    new_name = Replace(new_name, " ", "_")
    'Remove invalid characters
    new_name = Application.WorksheetFunction.Substitute(new_name, ":", "")
    new_name = Application.WorksheetFunction.Substitute(new_name, "/", "")
    new_name = Application.WorksheetFunction.Substitute(new_name, "\", "")
    new_name = Application.WorksheetFunction.Substitute(new_name, "?", "")
    new_name = Application.WorksheetFunction.Substitute(new_name, "*", "")
    format_name = new_name
End Function

' named ranges
Function create_named_range(ByVal name As String, ByVal sheetName As String, ByVal address As String, _
Optional ByVal default_value As Variant, Optional ByVal formula, Optional header_row As String = "", Optional id_row As String = "", _
Optional overwrite As Boolean = True, Optional expand_range As Boolean = False, Optional clear As Boolean = True)
    
    Dim rng As Range
    Dim exists As Boolean, num_rows As Long, num_cols As Long, wb As Workbook
    
    ' format as proper name for named_range
    name = r.format_name(name)
    
    ' workbook
    Set wb = r.get_default_wb()
    
    'Check if named range already exists
    On Error Resume Next
    Set rng = ThisWorkbook.Names(name).RefersToRange
    exists = Not rng Is Nothing
    On Error GoTo 0
    
    'If named range exists, delete it
    If overwrite Then
        delete_named_range name, ws, wb, clear:=clear
    End If
    
    'Create new named range
    Set rng = wb.Worksheets(sheetName).Range(address)
    
    ' optional: expand range
    If expand_range Then
      Set rng = r.expand_range(rng, ws, wb)
    End If
    
    ' add named range to Names
    wb.Names.Add name:=name, RefersTo:=rng
    
    'Set default value if provided
    If Not IsMissing(default_value) Then
       rng.value = default_value
    ElseIf Not IsMissing(formula) Then
       ' Set formula if provided
       rng.formula = formula
    End If
    num_cols = rng.columns.count
    num_rows = rng.Rows.count

    ' optionally fill header/first row
    Dim header_array() As String
    Dim id_array() As String
    
    ' If header_row is specified, set the values to the first row of the named range
    If Len(header_row) > 0 Then
        header_array = str_to_array(header_row)
        num_cols = get_array_len(header_array)
    End If
    
    ' If id_row is specified, set the values to the first column of the named range
    If Len(id_row) > 0 Then
        id_array = str_to_array(id_row)
        num_rows = get_array_len(id_array)
    End If
    
    If rng.columns.count <> num_cols Or rng.Rows.count <> num_rows Then
       Debug.Print rng.address
       Set rng = r.getResizedRange(rng.Cells(1, 1), num_rows - 1, num_cols - 1)
       Debug.Print num_rows, num_cols, rng.address
    End If
    
    If Len(header_row) > 0 Then
       rng.Rows(1).value = str_to_array(header_row)
    End If
    
    If Len(id_row) > 0 Then
      rng.columns(1).value = WorksheetFunction.Transpose(str_to_array(id_row))
    End If
    
    ' update named range
    wb.Names(name).RefersTo = rng
    
End Function

Sub delete_named_range(name As String, Optional ws As Variant, Optional wb As Variant, Optional clear As Boolean = True)
    ' Set default worksheet and workbook
    Set ws0 = get_default_ws(ws:=ws)
    Set wb0 = get_default_wb(wb:=wb)
    name = r.format_name(name)
    On Error Resume Next
    ' clear fill
    If clear Then
       clear_formatting wb0.Names(name).RefersToRange
    End If
    wb0.Names(name).Delete
    On Error GoTo 0

End Sub

Sub update_named_range(name As String, new_range, Optional wb)
    Dim rng0 As Range
    Set wb = r.get_default_wb(wb)
    Set rng0 = r.get_range(new_range)
    name = r.format_name(name)
    wb.Names(name).RefersTo = rng0
End Sub

Sub expandNamedRange(name As String, Optional wb, Optional dbg As Boolean = False)
    Dim rng0 As Range, rng1 As Range, wb0 As Workbook
    Set wb0 = r.get_default_wb(wb)
    Set rng0 = r.get_range(name)
    Set rng1 = r.expand_range(rng0, rng0.Worksheet, wb0, dbg:=dbg)
    r.update_named_range name, rng1, wb:=wb0
End Sub

Sub resize_named_range(name As String, Optional add_rows As Long = 0, Optional add_cols As Long = 0, Optional wb)
    Dim rng1 As Range, rng0 As Range
    Set rng0 = r.get_range(name, wb:=wb)
    Set rng1 = r.getResizedRange(rng0, add_rows:=add_rows, add_cols:=add_cols)
    r.update_named_range name:=name, new_range:=rng1, wb:=wb
End Sub

Sub subsetNamedRange(name As String, Optional startrow As Variant, Optional endrow As Variant, Optional startcol As Variant, Optional endcol As Variant, _
    Optional wb)
    Dim rng1 As Range, rng0 As Range
    Set rng0 = r.get_range(name, wb:=wb)
    Set rng1 = r.subset_range(rng0, startrow:=startrow, endrow:=endrow, startcol:=startcol, endcol:=endcol)
    r.update_named_range name:=name, new_range:=rng1, wb:=wb
End Sub

' fit the named range to cells with values
Sub fit_named_range_to_values(name As String, Optional wb, Optional searchUpRange As Range)
    Dim rng0 As Range, rng1 As Range, wb0 As Workbook
    Set wb0 = r.get_default_wb(wb)
    Set rng0 = r.get_range(name)
    
    ' the fitted range
    Set rng1 = r.expand_range(rng0, wb:=wb0, searchUpRange:=searchUpRange)
    
    r.update_named_range name, new_range:=rng1, wb:=wb0
End Sub

' Update a named range with values as formulas
Sub updateNamedRangeWithValues(named_range_name As String, values As Variant, Optional clear_formatting As Boolean = False)
    ' This subroutine updates a named range with the provided 2D array of values.
    ' The values are pasted as formulas.
    ' Parameters:
    '   - named_range: The name of the range to update.
    '   - values: A 2D array of values to paste into the named range.
    
    Dim rng As Range
    Dim numRows As Long, numCols As Long
    
    ' Check if values is a 2D array
    If Not a.is_2d_array(values) Then
        Err.Raise 1001, "updateNamedRangeWithValues", "values must be a 2D array"
    ElseIf a.array_length(values) < 0 Then
        Exit Sub 'array is empty
    End If
    
    ' Get the named range
    Set rng = r.get_range(named_range_name)
    Debug.Print "updateNamedRangeWithValues: initial range is " & rng.address
    
    ' Clear contents and formatting
    r.clear_range rng, rng.Worksheet, clear_formatting:=clear_formatting
    
    ' Resize the named range to match the dimensions of the values array
    numCols = a.num_array_columns(values)
    numRows = a.num_array_rows(values)
    
    Set rng = r.getResizedRange(rng, num_rows:=numRows, num_cols:=numCols)
    Debug.Print "updateNamedRangeWithValues: resized range is " & rng.address
     
    ' Paste the values
    r.paste_array values, rng.address, rng.Worksheet
    
    ' Update the named range
    r.update_named_range named_range_name, rng
End Sub

' formulas
Sub add_formula_column_to_named_range(name As String, Optional newColumn = "formula", Optional formulaDefinition = "=A1+B1")
    Dim rng As Range
    Dim newcolumn_range As Range, formula_range As Range
    'Dim formulaDefinition As String
    Dim numRows As Long
    
    ' Set the named range
    Set rng = get_range(name)
    
    ' Get the number of rows in the named range
    numRows = rng.Rows.count
    
    ' Define the formula for the new column
    ' formulaDefinition = "=YourFormula"
    
    ' Add a new column to the right of the named range
    Set newcolumn_range = rng.Offset(0, rng.columns.count).Resize(numRows, 1)
    
    ' Set the header of the new column
    newcolumn_range.Cells(1).value = newColumn
    
    ' Set the formula for the new column
    newcolumn_range.Cells(2).formula = formulaDefinition
    
    ' Fill the formula downwards if the named range has more than 2 rows
    ' Set the formula for the remaining cells in the new column
    If rng.Rows.count > 1 Then
        Set formula_range = rng.Offset(1, rng.columns.count).Resize(numRows - 1, 1)
        formula_range.Cells(1).AutoFill Destination:=formula_range, Type:=xlFillDefault
        formula_range.Activate
    End If
    
    ' Update the named range to include the new column
    update_named_range name, rng.Resize(numRows, rng.columns.count + 1)
End Sub

Sub fill_formula_range(rng, formulaDefinition, Optional only_values As Boolean = False)
    Dim rng0 As Range, formula_range As Range
    ' Set the range
    Set formula_range = get_range(rng)
    
    ' optional: exclude header (first row) if only_values
    If formula_range.Rows.count > 1 And only_values Then
       Set formula_range = r.subset_range(formula_range, startrow:=2)
    End If
    
    Debug.Print formula_range.address
    Debug.Print formulaDefinition
    
    ' Set the formula for the new column
    formula_range.Cells(1).formula = formulaDefinition
    If formula_range.Rows.count > 1 Then
        formula_range.Cells(1).AutoFill Destination:=formula_range, Type:=xlFillDefault
    End If
End Sub

Sub test_add_formula()
r.create_named_range "test", "test", "$G$1:$H$13", clear:=False
r.add_formula_column_to_named_range "test", formulaDefinition:="=G2+H2"
End Sub

Function safe_offset(rng0 As Range, Optional offset_row = 0, Optional offset_column = 0) As Range
    If rng0.Rows.count < r.MAX_XL_ROWS Then
       Set rng0 = rng0.Offset(offset_row, offset_column)
    End If
    Set safe_offset = rng0
End Function

Function get_range(rng As Variant, Optional ws As Variant, Optional wb As Variant, Optional offset_row = 0, Optional offset_column = 0) As Range
    Dim rng0 As Range, ws0 As Worksheet, wb0 As Workbook
    
    ' Set default worksheet and workbook
    If IsMissing(ws) Then
        Set ws0 = ActiveSheet
    ElseIf IsEmpty(ws) Then
        Set ws0 = ActiveSheet
    ElseIf ws Is Nothing Then
        Set ws0 = ActiveSheet
    Else
        Set ws0 = ws
    End If
    
    If IsMissing(wb) Then
        Set wb0 = ActiveWorkbook
    ElseIf IsEmpty(ws) Then
        Set wb0 = ActiveWorkbook
    ElseIf wb Is Nothing Then
       Set wb0 = ActiveWorkbook
    Else
       Set wb0 = wb
    End If
    
    ' Convert rng to range object
    If TypeName(rng) = "Range" Then
        Set rng0 = rng
        rng0.Worksheet.Activate
    ElseIf TypeName(rng) = "String" Then
        ' Check if rng is a named range
        Dim range_name As String: range_name = r.format_name(CStr(rng))
        On Error Resume Next
        Set rng0 = wb0.Names(range_name).RefersToRange
        rng0.Worksheet.Activate
        On Error GoTo 0
        
        If rng0 Is Nothing Then
            Debug.Print "Input is not named range: " & range_name
            ' Try to use rng as range address
            On Error Resume Next
            Set rng0 = ws0.Range(rng)
            rng0.Worksheet.Activate
            On Error GoTo 0
            
            If rng0 Is Nothing Then
                ' Failed to convert rng to range object
                Err.Raise 1002, , "Invalid range: " & range_name
            End If
        End If
    ElseIf TypeName(rng) = "Worksheet" Then
        Set ws0 = rng
        Set rng0 = r.expand_range("A1", ws0, wb)
    ElseIf TypeName(rng) = "Nothing" Then
        Set rng0 = Nothing
    Else
        ' Invalid argument type
        Err.Raise 1003, , "Invalid argument type: " & TypeName(rng)
    End If
    
    If offset_row <> 0 Or offset_column <> 0 Then
       Set rng0 = r.safe_offset(rng0, offset_row, offset_column)
    End If
    
    Set get_range = rng0
End Function

Function extend_range(rng As Variant, Optional ws As Variant, Optional wb As Variant, Optional add_rows As Integer = 0, Optional add_cols As Integer = 0) As Range
    ' Get the initial range object
    Dim rng0 As Range
    Set rng0 = r.get_range(rng, ws, wb)
    
    ' Extend the range if necessary
    If add_rows > 0 Or add_cols > 0 Then
        Dim lastRow As Long, lastCol As Long
        lastRow = rng0.Rows(rng0.Rows.count).row + add_rows
        lastCol = rng0.columns(rng0.columns.count).column + add_cols
        Set rng0 = ws.Range(rng0.Cells(1, 1), ws.Cells(lastRow, lastCol))
    End If
    
    ' Return the extended range
    Set extend_range = rng0
End Function

Function expand_range(rng As Variant, Optional ws As Variant, Optional wb As Variant, Optional max_num_cols = 0, Optional c1 As Long = 0, _
    Optional dbg As Boolean = False, Optional searchUpRange As Range) As Range
    Dim rng0 As Range, rng1 As Range, ws0 As Worksheet
    Set rng0 = get_range(rng, ws, wb)
    Set ws0 = rng0.Worksheet
    Dim lRow As Long
    Dim lColumn As Long
    
    ' find last row of range
    lRow = r.get_last_row(rng0, ws0, wb, range2:=searchUpRange)  ' absolute row number
    
    ' get last column with data OR passed c1
    If c1 = 0 Then
        
       lColumn = r.get_last_col(rng0, ws0, wb)
    Else
       lColumn = c1
    End If
    
    ' set maximum last column (to prevent large ranges)
    If max_num_cols > 0 Then
      lColumn = WorksheetFunction.Min(lColumn, max_num_cols)
    End If
    
    If lRow >= MAX_XL_ROWS Then
      Set expand_range = rng0
      Exit Function
    End If
    
    
    r0 = rng0.Cells(1, 1).row
    c0 = rng0.Cells(1, 1).column
    With ws0
    Set rng1 = .Range(.Cells(r0, c0), .Cells(lRow, lColumn))
    End With
    
    If dbg = True Then
       Debug.Print "expanding range ", rng0.address, " to row ", lRow, " column ", lColumn
    End If
    Set expand_range = rng1
    
    Exit Function
    
End Function

Function getResizedRange(rng As Range, Optional add_rows As Long = 0, Optional add_cols As Long = 0, Optional num_rows As Long = 0, Optional num_cols As Long = 0) As Range
    Dim new_range As Range
    Dim start_cell As Range
    
    ' Get the top-left cell of the original range
    Set start_cell = rng.Cells(1, 1)
    
    ' Resize the range: either add or set the new number of rows, columns
    If add_rows <> 0 Or add_cols <> 0 Then
       Set new_range = start_cell.Resize(rng.Rows.count + add_rows, rng.columns.count + add_cols)
    ElseIf num_rows > 0 Or num_cols > 0 Then
       If num_rows < 1 Then
          Err.Raise 1, "getResizedRange", str.subInStr("new number of rows: @1", num_rows)
       ElseIf num_cols < 1 Then
          Err.Raise 1, "getResizedRange", str.subInStr("new number of columns: @1", num_cols)
       End If
       Set new_range = start_cell.Resize(num_rows, num_cols)
    Else
       Set new_range = rng
    End If
    
    ' Return the resized range
    Set getResizedRange = new_range
End Function

'subsetting on row, column numbers
Function subset_range(rng0 As Range, _
    Optional startrow As Variant, Optional endrow As Variant, Optional startcol As Variant, Optional endcol As Variant) As Range
    Dim rng As Range
    Dim minRow As Long, maxRow As Long, minCol As Long, maxCol As Long
    
    If IsMissing(startrow) Or IsEmpty(startrow) Then
        minRow = rng0.row
    Else
        minRow = startrow + rng0.row - 1
    End If
    
    If IsMissing(endrow) Or IsEmpty(endrow) Then
        maxRow = rng0.Rows.count + rng0.row - 1
    Else
        maxRow = endrow + rng0.row - 1
    End If
    
    If IsMissing(startcol) Or IsEmpty(startcol) Then
        minCol = rng0.column
    Else
        minCol = startcol + rng0.column - 1
    End If
    
    If IsMissing(endcol) Or IsEmpty(endcol) Then
        maxCol = rng0.columns.count + rng0.column - 1
    Else
        maxCol = endcol + rng0.column - 1
    End If
    
    Set rng = Range(Cells(minRow, minCol), Cells(maxRow, maxCol))
    Set subset_range = Intersect(rng, rng0)
End Function

Function get_last_row(Optional rng As Variant, Optional ws As Variant, Optional wb As Variant, Optional range2 As Range = Nothing, _
    Optional start_column_index As Long = 1) As Long
    ' Set default values for worksheet and workbook if none are provided
    Dim ws0 As Worksheet
    Set ws0 = get_default_ws(ws)
    Dim wb0 As Workbook
    Set wb0 = get_default_wb(wb)
    
    ' Get the range to search for the last row with a value
    Dim rng0 As Range
    Set rng0 = get_range(rng, ws:=ws0, wb:=wb0)
    
    ' Find the last row with a value starting from the first cell in the range
    Dim lastRow As Long
    Set ws0 = rng0.Worksheet
    lastRow = ws0.Cells(rng0.Cells(1).row, rng0.Cells(start_column_index).column).End(xlDown).row

    ' Optional: Find the last row with a value starting from the second range
    If Not range2 Is Nothing Then
       lastRow2 = ws0.Cells(range2.Cells(1).row, range2.Cells(1).column).End(xlUp).row
       lastRow = WorksheetFunction.MAX(lastRow, lastRow2)
    End If
    
    If lastRow = r.MAX_XL_ROWS Then
      ' assume no row with value found, then set to first row of rng0
      lastRow = rng0.Cells(1).row
    End If
    
    ' Return the last row
    get_last_row = lastRow
End Function

Function get_last_col(Optional rng As Variant, Optional ws As Variant, Optional wb As Variant) As Long
    ' Set default values for worksheet and workbook if none are provided
    Dim ws0 As Worksheet
    Set ws0 = get_default_ws(ws)
    Dim wb0 As Workbook
    Set wb0 = get_default_wb(wb)
    
    ' Get the range to search for the last column with a value
    Dim rng0 As Range
    Set rng0 = get_range(rng, ws:=ws0, wb:=wb0)
    
    ' Find the last column with a value starting from the first cell in the range
    Dim lastCol As Long
    Set ws0 = rng0.Worksheet
    lastCol = ws0.Cells(rng0.Cells(1).row, rng0.Cells(1).column).End(xlToRight).column
    ' Return the last column
    get_last_col = lastCol
End Function

' returns the header (first row) of range
Function get_header(rng, Optional ws As Worksheet, Optional wb As Workbook) As Range
    Dim rng0 As Range
    Set rng0 = get_range(rng, ws, wb)
    Set get_header = rng0.Rows(1)
End Function

Function get_range_values(rng, Optional ws As Worksheet, Optional wb As Workbook, Optional offset_row = 1, Optional offset_column = 1) As Range
    Dim rng0 As Range, ws0 As Worksheet
    Dim rng1 As Range
    
    Set rng0 = get_range(rng, ws, wb)
    Set ws0 = get_default_ws(ws)
    
    'Set rng1 to the specified range, excluding the first row
    If rng0.Rows.count <= 1 Then
       Err.Raise 1, "r.get_range_values", "range has no interior"
       Exit Function
    End If
    
    Set rng1 = ws0.Range(rng0.address)
    If rng1.Rows.count < r.MAX_XL_ROWS Then
       Set rng1 = rng1.Offset(offset_row, offset_column).Resize(rng0.Rows.count - offset_row, rng0.columns.count - offset_column)
    End If
    
    'Return rng1
    Set get_range_values = rng1
End Function

Function get_column(rng, index As Variant, Optional ws As Worksheet, Optional wb As Workbook, Optional offset_row = 0) As Range
    Dim rng0 As Range
    Dim col_index As Long
    Set rng0 = get_range(rng, ws, wb)
    
    col_index = r.get_column_index(rng0, index)
    If offset_row < 1 Then
       Set get_column = rng0.columns(col_index)
    Else
       Set rng1 = rng0.columns(col_index)
       Set get_column = rng1.Offset(offset_row).Resize(rng1.Rows.count - offset_row)
    End If
End Function

Function get_column_values(rng As Range, index As Variant, Optional ws As Worksheet, Optional wb As Workbook) As Range
    Dim rng0 As Range
    Set rng0 = get_column(rng, index, ws, wb)
    If rng0.Rows.count > 1 Then
       Set rng0 = safe_offset(rng0, offset_row:=1)
       Set get_column_values = rng0.Resize(rng0.Rows.count - 1)
    Else
       Set get_column_values = Nothing
    End If
End Function

Function get_column_index(rng As Range, column_name) As Long
    Dim header0 As Range
    Set header0 = r.get_header(rng)
    get_column_index = r.get_index_of_value(column_name, header0)
End Function

Function getColumnIndex(rng As Range, column_name_index) As Long
    ' This function returns the column index based on the column name or index provided.
    ' If the column_name is an integer, it is treated as the column index.
    ' If the column_name is a string, it is treated as the column header name.
    
    Dim header0 As Range
    Dim match_index As Variant
    
    If IsNumeric(column_name_index) Then
        ' If column_name is an integer, check if it is within the range's column count
        column_index = column_name_index
        If column_index > 0 And column_index <= rng.columns.count Then
            ' Return the column index as is
            getColumnIndex = column_index
        Else
            ' Raise an error if the column index is out of bounds
            Err.Raise vbObjectError + 1, "get_column_index", "Column integer greater than range's column count."
        End If
    Else
        ' If column_name is a string, find the column index by matching the header name
        column_name = column_name_index
        Set header0 = r.get_header(rng)
        match_index = Application.match(column_name, header0, 0)
        If IsError(match_index) Then
            ' Raise an error if the column name is not found
            Err.Raise vbObjectError + 1, "get_column_index", "Column name not found in range header."
        Else
            ' Return the found column index
            getColumnIndex = match_index
        End If
    End If
End Function

Function get_column_indexes(rng As Range, column_names) As Variant
    Dim header0 As Range, column_names_col As collection
    Set header0 = r.get_header(rng)
    column_names_array = r.to_column_names(column_names)
    Set column_names_col = a.as_collection(column_names_array)
    
    column_indexes = a.create_vector(column_names_col.count)
    c = 1
    For Each column_name In column_names_col
    column_indexes(c) = r.get_index_of_value(CStr(column_name), header0)
    c = c + 1
    Next
    get_column_indexes = column_indexes
End Function

Function column_exist(rng0 As Range, column_name) As Boolean
    On Error GoTo not_exist
    index = r.get_index_of_value(column_name, r.get_header(rng0))
    On Error GoTo 0
    column_exist = True
    Exit Function
not_exist:
   column_exist = False
End Function

Function get_row(rng, index As Variant, Optional ws As Worksheet, Optional wb As Workbook, Optional offset_column = 0) As Range
    Dim rng0 As Range, rng1 As Range
    Dim row_index As Long
    Set rng0 = get_range(rng, ws, wb)
    
    row_index = r.get_index_of_value(index, rng0.Rows(1))
    
    If offset_column < 1 Then
       Set get_row = rng0.Rows(row_index)
    Else
       Set rng1 = rng0.Rows(row_index)
       Set get_row = rng1.Offset(ColumnOffset:=offset_column).Resize(ColumnSize:=rng1.columns.count - offset_column)
    End If
End Function

Function get_row_index(rng, row_index) As Long
    Dim row0 As Range, rng0 As Range
    Set rng0 = r.get_range(rng)
    Set row0 = rng0.columns(1)
    get_row_index = r.get_index_of_value(row_index, row0)
End Function

Function get_value(rng, r_index, c_index)
    Dim rng0 As Range, row_index As Long, col_index As Long
    Set rng0 = get_range(rng, ws, wb)
    row_index = r.get_row_index(rng0, r_index)
    col_index = r.get_column_index(rng0, c_index)
    get_value = rng0.Cells(row_index, col_index).value
End Function

Sub set_value(rng, r_index, c_index, value)
    Dim rng0 As Range, row_index As Long, col_index As Long
    Set rng0 = get_range(rng, ws, wb)
    row_index = r.get_row_index(rng0, r_index)
    col_index = r.get_column_index(rng0, c_index)
    rng0(row_index, col_index).value = value
End Sub

Function range_contains(rng0 As Range, x As Variant) As Boolean
    Dim arr0 As Variant
    Dim cell As Range
    Dim i As Long
    
    If VarType(x) = vbString Then
        arr0 = Split(x, ",")
    ElseIf VarType(x) = vbArray Then
        arr0 = x
    Else
        Err.Raise vbObjectError + 1, "range_contains", "Invalid input type"
    End If
    
    For i = LBound(arr0) To UBound(arr0)
        For Each cell In rng0.Cells
            'On Error GoTo not_contains
            If cell.value = arr0(i) Then
                range_contains = True
                Exit Function
            End If
            On Error GoTo 0
        Next cell
    Next i
    
    range_contains = False
not_contains:
    range_contains = False
End Function

Public Function get_default_ws(Optional ws As Variant) As Worksheet
    If IsMissing(ws) Then
        Set ws0 = ActiveSheet
    ElseIf IsEmpty(ws) Then
        Set ws0 = ActiveSheet
    ElseIf ws Is Nothing Then
        Set ws0 = ActiveSheet
    ElseIf TypeName(ws) = "String" Then
        Set ws0 = ThisWorkbook.Sheets(ws)
    Else
        Set ws0 = ws
    End If
    Set get_default_ws = ws0
End Function

Public Function get_default_wb(Optional wb As Variant) As Workbook
    If IsMissing(wb) Then
        Set wb0 = ActiveWorkbook
    ElseIf IsEmpty(wb) Then
        Set wb0 = ActiveWorkbook
    ElseIf wb Is Nothing Then
        Set wb0 = ActiveWorkbook
    Else
       Set wb0 = wb
    End If
    Set get_default_wb = wb0
End Function


Function name_exist(name As String, Optional ws As Worksheet, Optional wb As Workbook) As Boolean
    Dim ws0 As Worksheet, wb0 As Workbook
    
    ' Get default worksheet if not specified
    'Set ws0 = r.get_default_ws(ws:=ws)
    ' Get default workbook if not specified
    Set wb0 = r.get_default_wb(wb:=wb)
    
    ' Check if named range exists in worksheet
    On Error Resume Next
    Dim rng As Range
    name = r.format_name(name)
    Set rng = wb0.Names(name).RefersToRange
    If rng Is Nothing Then
        name_exist = False
    Else
        name_exist = True
    End If
    On Error GoTo 0
End Function

' Addresses
Function get_range_address(ws0 As Worksheet, Optional r0 = 1, Optional r1 = 1, Optional c0 = 1, Optional c1 = 1) As String
    Dim rangeAddress As String
    rangeAddress = ws0.Cells(r0, c0).address & ":" & ws0.Cells(r1, c1).address
    get_range_address = rangeAddress
End Function

Function removeContextWithinBrackets(file_name As String) As String
    ' This function removes the context within square brackets [] if it exists in the given string.
    ' If no square brackets are found, it returns the original string.
    '
    ' Parameters:
    ' file_name - The input string that may contain context within square brackets.
    '
    ' Returns:
    ' The string with the context within square brackets removed, or the original string if no brackets are found.
    
    Dim startPos As Long
    Dim endPos As Long
    
    ' Find the position of the opening bracket [
    startPos = InStr(file_name, "[")
    
    ' Find the position of the closing bracket ]
    endPos = InStr(file_name, "]")
    
    ' Check if both brackets are found
    If startPos > 0 And endPos > 0 And endPos > startPos Then
        ' Remove the context within the brackets
        removeContextWithinBrackets = left(file_name, startPos - 1) & Mid(file_name, endPos + 1)
    Else
        ' Return the original string if no brackets are found
        removeContextWithinBrackets = file_name
    End If
End Function

Function getRangeFullAddress(rng As Range, Optional removeFileName As Boolean = True, Optional removeDollarSigns As Boolean = False)
    Dim address As String, full_address As String
    address = rng.address(External:=True)
    If removeFileName Then
       full_address = removeContextWithinBrackets(address)
    Else
       full_address = address
    End If
    
    If removeDollarSigns Then
       full_address = Replace(full_address, "$", "")
    End If
    getRangeFullAddress = full_address
End Function

Function get_column_address(rng0 As Range) As String
    startcol = rng0.columns(1).column
    endcol = rng0.columns(rng0.columns.count).column
    
    startColAddress = Split(Cells(1, startcol).address, "$")(1)
    endColAddress = Split(Cells(1, endcol).address, "$")(1)
    
    get_column_address = u.remove_dollar_sign(startColAddress & ":" & endColAddress)
End Function

' cell formatting
Sub clear_formatting(rng0, Optional ws, Optional wb)
    Dim rng As Range
    Set rng = r.get_range(rng0, ws, wb)
    rng.Interior.ColorIndex = xlColorIndexNone ' White color
    rng.Cells.ClearContents
    rng.ClearFormats 'Clear all formatting
    rng.NumberFormat = "General" 'Set number format to general
End Sub

Sub clear_range(rng0, Optional ws, Optional wb, Optional clear_formatting As Boolean = True)
    Dim rng As Range
    Set rng = r.get_range(rng0, ws, wb)
    rng.Cells.ClearContents
    If clear_formatting Then
        rng.Interior.ColorIndex = xlColorIndexNone ' White color
        rng.ClearFormats 'Clear all formatting
        ' Clear conditional formatting
        rng.FormatConditions.Delete
        rng.NumberFormat = "General" 'Set number format to general
    End If
End Sub

Sub clear_range_values(rng, Optional clear_formatting As Boolean = True)
    Dim rng0 As Range
    Set rng0 = r.get_range(rng)
    If rng0.Rows.count > 1 Then
       Set rng0 = safe_offset(rng0, offset_row:=1) 'rng0.Offset(1)
       r.clear_range rng0, clear_formatting:=clear_formatting
    End If
End Sub

Sub setConditionalFormatting(rng As Range, greenValue As Variant, redValue As Variant)
    ' Clear any existing conditional formatting
    rng.FormatConditions.Delete
    
    ' Set up the green formatting rule
    Dim greenRule As FormatCondition
    Set greenRule = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=greenValue)
    With greenRule.Interior
        .pattern = xlSolid
        .Color = RGB(0, 255, 0) ' Green color
    End With
    
    ' Set up the red formatting rule
    Dim redRule As FormatCondition
    Set redRule = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=redValue)
    With redRule.Interior
        .pattern = xlSolid
        .Color = RGB(255, 0, 0) ' Red color
    End With
End Sub

Sub protect_sheet(ws As Worksheet, Optional unlock_ranges = "")
    
    ' Protect the worksheet
    ws.Protect
    
    ' Convert comma-separated string to array
    Dim unlock_ranges_arr() As String
    unlock_ranges_arr = str_to_array(unlock_ranges)
    
    ' Loop through each range in the array and unlock it
    Dim rng As Variant
    Dim rng0 As Range
    For Each rng In unlock_ranges_arr
        Set rng0 = get_range(rng, ws)
        rng0.Locked = False
    Next rng
        
        
End Sub

Sub format_row_column_size(ws As Worksheet, r_height As Double, c_width As Double)
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long
    
    ' Get the last row and column of the worksheet
    lastRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookAt:=xlPart, _
                LookIn:=xlFormulas, SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, MatchCase:=False).row
    lastCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookAt:=xlPart, _
                LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, MatchCase:=False).column
    
    ' Check if any row has a different height
    For i = 1 To lastRow
        If ws.Rows(i).RowHeight <> r_height Then
            ws.Rows(i).RowHeight = r_height
        End If
    Next i
    
    ' Check if any column has a different width
    For j = 1 To lastCol
        If ws.columns(j).ColumnWidth <> c_width Then
            ws.columns(j).ColumnWidth = c_width
        End If
    Next j
    
End Sub

Sub autofit_columns_rows(rng As Range)
    Dim i As Long, j As Long
    For i = 1 To rng.columns.count
        rng.columns(i).EntireColumn.AutoFit
    Next i
    For j = 1 To rng.Rows.count
        rng.Rows(j).EntireRow.AutoFit
    Next j
End Sub

Sub autofit_columns(rng As Range)
    Dim i As Long
    For i = 1 To rng.columns.count
        'rng.Columns(i).EntireColumn.AutoFit
        rng.Worksheet.columns(rng.Cells(1, i).column).EntireColumn.AutoFit
    Next i
End Sub

Function get_column_formats(rng As Range) As Variant
    Dim columnFormatArr() As Variant
    Dim i As Long
    
    ' Re-dimension the column format array
    ReDim columnFormatArr(1 To rng.columns.count)
    
    ' Loop through each cell in the specified column and store its format
    For i = 1 To rng.columns.count
        columnFormatArr(i) = rng.columns(i).NumberFormat
    Next i
    
    ' Return the column format array
    get_column_formats = columnFormatArr
End Function

Function rng_to_1d_array(rng0) As Variant
    Dim arr() As Variant, rng As Range
    Dim i As Long
    
    Set rng = get_range(rng0)
    
    ' Resize the array to match the size of the input range
    ReDim arr(1 To rng.Cells.count)
    
    ' Loop through the range and populate the array
    For i = 1 To rng.Cells.count
        arr(i) = rng.Cells(i).value
    Next i
    
    ' Return the array
    rng_to_1d_array = arr
End Function

Function rng_to_2d_array(rng0) As Variant
    Dim arr() As Variant, rng As Range
    Dim i As Long
    Dim j As Long
    
    Set rng = get_range(rng0)
    
    ' Resize the array to match the size of the input range
    ReDim arr(1 To rng.Rows.count, 1 To rng.columns.count)
    
    ' Loop through the range and populate the array
    For i = 1 To rng.Rows.count
        For j = 1 To rng.columns.count
            arr(i, j) = rng.Cells(i, j).value
        Next j
    Next i
    
    ' Return the array
    rng_to_2darray = arr
End Function

Function get_unique_vals(rng) As Variant
    Dim rng0 As Range
    Dim dict As Object
    Dim i As Long
    
    ' Create a dictionary object to store unique values
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop through the array and add unique values to the dictionary
    If rng Is Nothing Then
        Err.Raise 1001, , "range is nothing"
    End If
    
    arr = rng_to_1d_array(rng)
    For i = LBound(arr) To UBound(arr)
        If Not dict.exists(arr(i)) Then
            dict.Add arr(i), 0
        End If
    Next i
    
    ' Convert the dictionary keys to an array and return it
    get_unique_vals = dict.Keys()
End Function

Function FilterRange(rng As Range, column_name As String, filter_value As Variant) As Range
    Dim headers As Range
    Dim columnIndex As Long
    Dim filteredRange As Range
    
    ' Enable AutoFilter for the range
    rng.AutoFilter
    
    ' Find the index of the column matching the column_name
    columnIndex = r.get_column_index(rng, column_name)
    
    ' Apply the filter to the range
    rng.AutoFilter field:=columnIndex, Criteria1:=filter_value
    
    ' Set the filtered range as the visible cells (excluding headers)
    On Error Resume Next
    Set filteredRange = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    ' Remove the filter
    rng.AutoFilter
    
    ' Return the filtered range
    Set FilterRange = filteredRange
End Function

' apply AutoFilter to range and return array
Function filter_range(rng As Range, column_name As String, filter_value As Variant, Optional column_name_2 As String, _
                      Optional filter_value_2 As Variant, Optional remove_filter As Boolean = True, Optional xl_operator As XlAutoFilterOperator = xlAnd) As Variant
    Dim headers As Range
    Dim columnIndex As Long
    Dim filteredRange As Range
    
    ' Check inputs
    If xl_operator = xlFilterValues And Not IsArray(filter_value) Then
       Err.Raise 1001, "", "filter_value is not an array"
    End If
    
    ' remove AutoFilter from range if set on worksheet
    If rng.Worksheet.AutoFilterMode Then
        ' Remove autofilter
        rng.AutoFilter
    End If
    
    ' sort the range on column_name
    r.sort_range_by_columns rng, column_name, xlAscending
    
    ' Enable AutoFilter for the range
    rng.AutoFilter
    
    ' Find the index of the column matching the column_name
    columnIndex = r.get_column_index(rng, column_name)
    
    ' Apply the filter to the range
    rng.AutoFilter field:=columnIndex, Criteria1:=filter_value, Operator:=xl_operator
    
    ' Optionally apply the second filter if set
    If Not IsEmpty(column_name_2) And column_name_2 <> "" Then
        columnIndex = r.get_column_index(rng, column_name_2)
        rng.AutoFilter field:=columnIndex, Criteria1:=filter_value_2
    End If
    
    ' Set the filtered range as the visible cells (excluding headers)
    Set filteredRange = rng.SpecialCells(xlCellTypeVisible)
     
    ' return based on number of areas (1,2)
    If filteredRange.Areas.count <= 1 Then
       Debug.Print "filteredRange has single area:", filteredRange.Areas(1).address
       filter_range = filteredRange.value
    Else
       Debug.Print "filteredRange has multiple areas:", filteredRange.Areas.count
       filter_range = filteredRange.Areas(1).value
       Debug.Print filteredRange.Areas(1).address
       For i = 2 To filteredRange.Areas.count
           Debug.Print filteredRange.Areas(i).address
           filter_range = a.concatArrays(filter_range, filteredRange.Areas(i).value)
       Next i
    End If
    
    ' Remove the filter
    If remove_filter Then
       rng.AutoFilter
    End If
          
End Function

Function get_index_of_value(value As Variant, rng As Range) As Variant
    ' Find the first occurrence of the value in the range
    Dim match_index As Variant
    match_index = value
    If VarType(match_index) = vbString Then
        match_index = Application.match(value, rng, 0)
        If IsError(match_index) Then
            Err.Raise vbObjectError + 1, "get_index_of_value", "Index " & value & " not found in range"
        End If
    ElseIf VarType(match_index) = vbLong Or VarType(match_index) = vbInteger Then
        If match_index < 1 Or match_index > rng.Cells.count Then
            Err.Raise vbObjectError + 1, "get_index_of_value", "Index " & value & " not found in range"
        End If
    Else
        Err.Raise vbObjectError + 1, "match_index", "Invalid index type"
    End If
    
    get_index_of_value = match_index
    
End Function

Function to_column_names(column_names As Variant) As Variant
    Dim result As Variant
    
    ' Check if column_names is a string
    If TypeName(column_names) = "String" Then
        ' Convert string to 1-dimensional array
        result = Split(column_names, ",")
    ElseIf IsArray(column_names) Then
        ' Check if column_names is a 2-dimensional array
        If a.is_2d_array(column_names) = True Then
            ' Convert 2-dimensional array to 1-dimensional array
            result = a.ConvertTo1DArray(column_names)
        Else
            ' column_names is already a 1-dimensional array, do nothing
            result = column_names
        End If
    Else
        ' Raise an error for invalid input
        Err.Raise vbObjectError + 9999, , "Invalid input. Expected string or array."
    End If
    
    to_column_names = result
End Function

Sub sort_range_by_columns(rng As Range, column_names As Variant, Optional sort_order As XlSortOrder = xlAscending)
    ' sort_column_indices can be a single column index or an array of column indices to sort by
    
    Dim sort_range As Range
    Dim sort_column_indices As Variant
    Dim header0 As Range
    
    ' get column_names to sort on as array
    column_names_array = r.to_column_names(column_names)
    
    ReDim sort_column_indices(LBound(column_names_array) To UBound(column_names_array))
    Set header0 = r.get_header(rng)
    For i = LBound(sort_column_indices) To UBound(sort_column_indices)
       sort_column_indices(i) = get_index_of_value(column_names_array(i), header0)
    Next i
    
    ' Apply the sort to the sort range
    Set sort_range = rng
    With sort_range
        '.Sort Key1:=.Columns(sort_column_indices(0)), Order1:=sort_order, _
              Orientation:=xlSortColumns, Header:=xlYes
        For i = LBound(sort_column_indices) To UBound(sort_column_indices)
            .Sort Key1:=.columns(sort_column_indices(i)), Order1:=sort_order, _
                  Orientation:=xlSortColumns, header:=xlYes
        Next i
    End With
End Sub

' using worksheet.sortfields
Sub sort_range_by_columns_2(rng As Range, column_names As Variant, Optional sort_order As XlSortOrder = xlAscending)
    Dim rng0 As Range, ws0 As Worksheet
    Dim ws As Worksheet
    Dim sortKey0 As Range
    Dim default_sort_columns As Variant
    
    ' Set the worksheet and range variables
    Set rng0 = r.get_range(rng)
    Set ws0 = rng0.Worksheet
    
    ' Set the sort range including the header row
    Set SortRange = rng0
    
    ' Set the sort keys
    column_names_array = r.to_column_names(column_names)
    column_indexes = r.get_column_indexes(rng0, column_names_array)
    
    ' Sort the range based on the sort keys
    With ws0.Sort
        .SortFields.clear
        For i = LBound(column_names_array) To UBound(column_names_array)
            Set sortKey0 = rng0.Rows(1).Find(column_names_array(i))
            .SortFields.Add key:=sortKey0, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Next i
        .SetRange SortRange
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' insert, delete columns, rows
Sub InsertColumnIntoRange(rng As Range, column_name_index As Variant, Optional new_column_name As String, _
    Optional values As Variant)
    ' This function inserts a new column into the specified range before the column identified by column_name_index.
    ' If values are provided, it sets these values in the new column.
    '
    ' Parameters:
    ' rng - The range where the new column will be inserted.
    ' column_name_index - The name or index of the column before which the new column will be inserted.
    ' values - (Optional) The values to be set in the new column.
    '
    ' Returns: The range including the new column.
    
    Dim columnIndex As Long
    Dim newColumn As Range
    Dim ws As Worksheet
    Dim column_name As String
    
    ' Set the worksheet based on the range
    Set ws = rng.Worksheet
    
    ' Get the column index using the getColumnIndex function
    columnIndex = getColumnIndex(rng, column_name_index)
    
    ' If column_name is provided, check if it exists using column_exist
    If new_column_name <> "" Then
        If column_exist(rng, new_column_name) Then
            Err.Raise Number:=vbObjectError + 1001, Description:="Column '" & new_column_name & "' already exists in the range."
        End If
    End If
    
    ' Insert the new column at the specified index
    Set newColumn = rng.columns(columnIndex)
    newColumn.Insert shift:=xlToRight
    
    ' Get the new column by index
    Set newColumn = r.get_column(rng, columnIndex)
    
    If new_column_name <> "" Then
       newColumn.Cells(1, 1).value = new_column_name
    End If
    
    If IsMissing(values) = False Then
        r.SetColumnValues newColumn, values
    End If
End Sub

Sub DeleteColumnFromRange(rng As Range, column_name_index As Variant)
    ' This subroutine deletes a column from the specified range based on the column name or index.
    ' If the column_name_index is an integer, it is treated as the column index.
    ' If the column_name_index is a string, it is treated as the column header name.
    '
    ' Parameters:
    ' rng - The range from which the column will be deleted.
    ' column_name_index - The name or index of the column to be deleted.
    
    Dim columnIndex As Long
    Dim columnToDelete As Range
    Dim ws As Worksheet
    
    ' Set the worksheet based on the range
    Set ws = rng.Worksheet
    
    ' Get the column index using the getColumnIndex function
    columnIndex = getColumnIndex(rng, column_name_index)
    
    ' Delete the column at the specified index
    Set columnToDelete = rng.columns(columnIndex).EntireColumn
    columnToDelete.Delete shift:=xlToLeft
    
End Sub

Sub AppendColumnToRange(rng As Range, Optional new_column_name As String, _
    Optional values As Variant, Optional overwrite As Boolean = True)
    ' This subroutine appends a new column to the end of the specified range.
    ' If values are provided, it sets these values in the new column.
    '
    ' Parameters:
    ' rng - The range to which the new column will be appended.
    ' new_column_name - (Optional) The name for the header of the new column.
    ' values - (Optional) The values to be set in the new column.
    
    Dim ws As Worksheet
    Dim lastColumn As Long
    Dim newColumn As Range
    
    ' Set the worksheet based on the range
    Set ws = rng.Worksheet
    
    ' Check if column exists
    If new_column_name <> "" And r.column_exist(rng, new_column_name) Then
       If Not overwrite Then
          Exit Sub
       Else
          r.DeleteColumnFromRange rng, new_column_name
       End If
    End If
    
    ' Find the last column index of the range
    firstRow = rng.Rows(1).row
    lastRow = rng.Rows(rng.Rows.count).row
    lastColumn = rng.columns(rng.columns.count).column
    
    ' Append the new column at the end of the range
    Set newColumn = ws.Range(Cells(firstRow, lastColumn + 1), Cells(lastRow, lastColumn + 1))
    
    ' If values are provided, set them in the new column
    If Not IsMissing(values) Then
        r.SetColumnValues newColumn, values
    End If
    
    ' If new_column_name is provided, set it as the header of the new column
    If new_column_name <> "" Then
       newColumn.Cells(1, 1).value = new_column_name
    End If
End Sub

' copy paste
Sub copy_range(rng0, addr As String, Optional ws As Worksheet = Nothing)
    Dim rng As Range, ws0 As Worksheet
    
    ' Set source worksheet
    Set rng = r.get_range(rng0, ws:=ws)
    Set ws0 = rng0.Worksheet
    
    ' Set destination worksheet if not provided
    If ws Is Nothing Then Set ws = rng.Worksheet

    ' Copy the range
    rng.Copy
    
    ' Paste values and formats to the destination address on the worksheet
    ws.Range(addr).PasteSpecial xlPasteValues
    ws.Range(addr).PasteSpecial xlPasteFormats
    
    ' Clear the clipboard
    Application.CutCopyMode = False
        
    ' Return to source worksheet
    ws0.Activate
End Sub

Sub copyRangeFormulas(rng0, addr As String, Optional ws As Worksheet = Nothing)
    Dim rng As Range, ws0 As Worksheet
    
    ' Set source worksheet
    Set rng = r.get_range(rng0, ws:=ws)
    Set ws0 = rng0.Worksheet
    
    ' Set destination worksheet if not provided
    If ws Is Nothing Then Set ws = rng.Worksheet

    ' Copy the range
    rng.Copy
    
    ' Paste formulas and formats to the destination address on the worksheet
    ws.Range(addr).PasteSpecial xlPasteFormulas
    ws.Range(addr).PasteSpecial xlPasteFormats
    
    ' Clear the clipboard
    Application.CutCopyMode = False
        
    ' Return to source worksheet
    ws0.Activate
End Sub

Sub paste_array(arrayToCopy, addr As String, Optional ws As Worksheet = Nothing)
    Dim desRange As Range
    Set ws = r.get_default_ws(ws)
    Set desRange = ws.Range(addr)
    desRange.Resize(UBound(arrayToCopy, 1), UBound(arrayToCopy, 2)).value = arrayToCopy
End Sub

Function select_columns(rng As Range, column_names) As Range
    Dim headers As Range
    Dim column As Range
    Dim selectedRange As Range
    Dim selectedHeaders As Range
    
    Set header0 = r.get_header(rng)
    
    For Each column In header0.Cells
        If IsError(Application.match(column.value, column_names, 0)) Then
            ' Column name not found in the selected names, skip it
        Else
            If selectedRange Is Nothing Then
                ' Set the initial selected range
                Set selectedRange = rng.columns(column.column)
                Set selectedHeaders = column
            Else
                ' Union the subsequent selected ranges
                Set selectedRange = Union(selectedRange, rng.columns(column.column))
                Set selectedHeaders = Union(selectedHeaders, column)
            End If
        End If
    Next column
    
    ' Include the header row in the selected range
    Set selectedRange = Union(selectedHeaders, selectedRange)

    'Return
    Set select_columns = selectedRange
    
End Function

' formatting
Function get_color_index(rng)
    Dim color_index As Long
    color_index = r.get_range(rng).Interior.Color 'Replace "A1" with the cell or range containing the color you want to get
    get_color_index = color_index
End Function

Sub add_outside_border(rng)
    Dim rng0 As Range
    Set rng0 = r.get_range(rng)
    With rng0
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).ColorIndex = xlAutomatic
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).ColorIndex = xlAutomatic
    End With
End Sub

Sub add_all_borders(rng)
    Dim rng0 As Range
    Set rng0 = r.get_range(rng)
    ' Add outside border and border for each row
    rng0.BorderAround Weight:=xlThin, ColorIndex:=xlColorIndexAutomatic
    rng0.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rng0.Borders(xlInsideHorizontal).Weight = xlThin
    rng0.Borders(xlInsideVertical).LineStyle = xlContinuous
    rng0.Borders(xlInsideVertical).Weight = xlThin
End Sub

Sub ClearAllBorders(rng As Range)
    ' This subroutine removes all borders from the specified range.
    '
    ' Parameters:
    ' rng - The range from which all borders will be removed.
    
    Dim rng0 As Range
    Set rng0 = r.get_range(rng)
    With rng0.Borders
        .LineStyle = xlNone
    End With
End Sub

Sub get_color_i()
    Dim color_index As Long
    Dim rgb_code As String
    color_index = Range("A1").Interior.Color 'Replace "A1" with the cell or range containing the color you want to get
End Sub

Sub format_columns(ws As Worksheet, column_format_mapping As Variant, Optional wb As Workbook, Optional ky_delim As String = "=", _
                   Optional it_delim As String = ";")
    ' Define variables
    Dim mapping As Scripting.Dictionary
    Dim rng As Range
    Dim colName As Variant
    Dim colFormat As Variant
    Dim colIndex As Long
    
    ' Get the workbook
    Set wb = r.get_default_wb(wb)

    ' Get the range from the worksheet
    Set rng = r.get_range(ws, wb:=wb)
    
    ' Get the mapping from column_format_mapping
    If TypeName(column_format_mapping) = "String" Then
        ' Convert string to dictionary
        Set mapping = dict.getDictionaryFromString(CStr(column_format_mapping), ky_delim, it_delim)
    ElseIf TypeName(column_format_mapping) = "Dictionary" Then
        ' Use the provided dictionary
        Set mapping = column_format_mapping
    Else
        ' Raise an error if column_format_mapping is not a string or dictionary
        Err.Raise 1001, "format_columns", "column_format_mapping must be a string or dictionary"
    End If
    
    ' Loop over keys in mapping
    For Each colName In mapping.Keys
        colFormat = mapping(colName)
        
        ' Find the associated column index
        colIndex = r.get_column_index(rng, colName)
        
        ' Check if the column index is valid
        If colIndex > 0 Then
            ' Set the column format
            rng.columns(colIndex).NumberFormat = colFormat
        Else
            ' Raise an error if the column name is not found
            Err.Raise 1002, "format_columns", "Column name not found: " & colName
        End If
    Next colName
End Sub

' helpers

' returns string as one dimensional array
Function str_to_array(str0) As Variant
    Dim delimiter As String
    delimiter = ","
    If VarType(str0) = vbString Then
        str_to_array = Split(str0, delimiter)
    ElseIf VarType(str0) = (vbVariant Or vbArray) Then
        str_to_array = str0
    Else
        Dim arr0() As String
        ReDim arr0(UBound(args))
        For i = 0 To UBound(args)
            arr0(i) = CStr(args(i))
        Next i
        str_to_array = arr0
    End If
End Function

Function get_array_len(arr) As Long
    get_array_len = UBound(arr) - LBound(arr) + 1
End Function

' TESTS
Sub test_set_bold_row()
    Dim orders_rng As Range, orders_values_rng As Range
    Set orders_rng = get_orders_range("LN 1")
    enddates = r.get_column_values(orders_rng, main.ENDDATE_COLUMN)
    overflow_row_index = main.find_week_overflow_row(enddates)
    Debug.Print overflow_row_index, enddates(overflow_row_index, 1)
    Set orders_values_rng = r.get_range_values(orders_rng)
    orders_values_rng.Rows(overflow_row_index).Font.Bold = True
End Sub



