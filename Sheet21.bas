Private Sub Worksheet_Open()
    main.control_sheet_update_database_settings Range("A1")
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    main.control_sheet_update_database_settings Target
End Sub



