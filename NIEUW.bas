' This subroutine checks if the change occurred in the column with the index CapGrpColumnIndex.
' If so, it calls the subroutine main.handle_input_capgrp with the range of the changed column.
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim CapGrpColumnIndex As Long
    Dim capgrp_column_range As Range

    ' Set the worksheet to "NIEUW"
    Set ws = Target.Worksheet

    ' Define the column index for CapGrp
    CapGrpColumnIndex = main.INPUT_CAPGRP_COLUMN_INDEX ' Change this to the actual column index for CapGrp

    ' Check if the change occurred in the CapGrp column
    If Not Intersect(Target, ws.columns(CapGrpColumnIndex)) Is Nothing Then
        ' Set the range for the CapGrp column excluding the header
        Set capgrp_column_range = ws.Range(ws.Cells(2, CapGrpColumnIndex), ws.Cells(ws.Rows.count, CapGrpColumnIndex).End(xlUp))

        ' Call the main.handle_input_capgrp subroutine
        Application.EnableEvents = False
        main.handle_input_capgrp capgrp_column_range
        Application.EnableEvents = True
    End If
End Sub


