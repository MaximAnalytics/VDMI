' Define init_workbook function to create necessary sheets and headers
Private Sub Workbook_Open()
    Application.EnableEvents = False
    Application.DisplayAlerts = True
    
    ' Initialize the global named ranges
    On Error Resume Next
       main.init_named_ranges
    On Error GoTo 0
    
    ' Order sheets: start with `instructie`, `overzicht`, `Template`, `BULK`
    On Error Resume Next
       main.init_worksheet_sorting
    On Error GoTo 0

    ' Active the start sheet if exists
    On Error Resume Next
       ThisWorkbook.Sheets(main.START_SHEET_NAME).Activate
    On Error GoTo 0
    
    ' Initialize articles per pallet sheet deactivate bcs 32-bit problem
    On Error Resume Next
       ' main.init_articles_per_pallet
    On Error GoTo 0
    
    ' Enable events
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub




