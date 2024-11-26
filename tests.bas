'tests for planning_automatisering_macro.xlsm
'0. initialize: set data base clear all sheets
'1. import ISAH data capgrp sheets

Global Const testdatabase = "JKR2"
Global Const release_database = "PROD"
Global Const current_prodwk = 29
Global Const previous_prodwk = 28
Global Const CAPGRP_START_ROW = 14
Global Const NUMBER_ARTICLES_LN1 = 83

Dim wb As Workbook, capgrp_sheet As String, ws As Worksheet, ordersRng As Range, worktimesRng As Range
Dim IsahSheet As Worksheet, BulkSheet As Worksheet

Const P_RELEASE As Boolean = True

Sub test_all()

GoTo start_test

    ' 0. Initialize:
    ThisWorkbook.Activate
    tests.set_database tests.testdatabase
    tests.set_init_capgrp_sheets
    tests.test_btn_clear_sheet

    ' 1. IMPORT DATE (overhalen orders): INPUT PRODWK29 LN1
    tests.set_input_isah_to_wk29_ln1
    tests.test_btn_import_art_ln1
    
    ' 2. CONTROLS: PRODWK29
    tests.set_input_isah_to_wk29
    tests.test_add_capgrp ' remove and re-add LN18
    
    tests.test_btn_import_art_all
    tests.test_btn_import_bulk
    tests.test_bulk_sheet_values

    ' 3. ADD/REMOVE/MUTATE/PRINT capgrp LN1 orders
    tests.test_prodwk29_ln1
    
    ' 4. ISAH import/export
    tests.test_workbook_connection 'workbook connection is needed for database operations
    tests.insert_update_isah_testdata_tables
    tests.test_isah_staging_ln1
    tests.test_isah_imports
    tests.test_isah_export
    
    ' 5. State control
    tests.test_state_control
    
    ' 6. test adding multiple new capgrps
    tests.test_set_new_capgrps
    
    ' 7. test add new orders
    tests.test_add_new_orders

    ' 8. test week 29: ISAH import export
    tests.test_isah_wk29_import_export
    
start_test:
    If P_RELEASE Then
       tests.set_for_release
    End If
    
Exit Sub

End Sub

Sub test_isah_wk29_import_export()
'GoTo clean_up
    main.btn_clear_sheet_Click
    tests.set_input_isah_to_wk29
    main.btn_import_art_Click
    tests.insert_isah_testdata_wk29
    main.btn_isah_database_export_Click
    
clean_up:
    
End Sub

Sub test_set_new_capgrps()
   'new worksheets
   w.deleteWorksheets "LN 6", "NW", "PROM", "INPK", "LN20", "LN22", "LN24", "LN26"
   
   tests.set_input_isah_to_wk46_new_capgrps
   main.btn_add_capgrp_sheets_Click

   Debug.Assert w.sheet_exists("LN22") And w.sheet_exists("LN24") And w.sheet_exists("LN26")
   Debug.Assert w.sheet_exists("NW") And w.sheet_exists("PROM") And w.sheet_exists("INPK")
   
   ' clear ws
   w.deleteWorksheets "LN 6", "NW", "PROM", "INPK", "LN20", "LN22", "LN24", "LN26"
End Sub

Sub test_add_new_orders()
    Application.ScreenUpdating = False
    tests.set_input_isah_to_wk48
    tests.set_input_new_orders_to_wk48
    
    ' check for new capgrps and add then
    main.btn_add_capgrp_sheets_Click
    
    ' set LN1 to week 48
    main.set_capgrp_weeknumber "LN 1", 48
    
    ' import all orders to capgrp tabs
    main.btn_import_art_Click
    
    ' add new orders to capgrp tab
    main.btn_add_new_orders_Click
    
    ' now check if each of the records in NIEUW is in the correct sheet
    Dim new_orders_rng As Range, record As Dictionary, Records As collection, capgrp As String, orders_rng As Range
    Set new_orders_rng = main.get_new_orders_range()
    Set Records = r.getRowsAsRecords(new_orders_rng)
    
    capgrp_ = ""
    For Each record In Records
        capgrp = record.item("Cap.Grp")
        If capgrp_ <> capgrp Then
           Debug.Print "checking new orders for: " & capgrp
           Set orders_rng = main.get_orders_range(capgrp)
        End If
        
        If orders_rng.Rows.count <= 1 Then
           GoTo nx
        End If

        ' check productieorder in orders_rng.productieorder
        productie_orders = r.get_column_values(orders_rng, "Productieorder")
        chk_productie_order = record.item("Productieorder")
        a.printArray productie_orders
        Debug.Assert u.InList(CLng(chk_productie_order), productie_orders)
        
        capgrp_ = capgrp
nx:
    Next
    Application.ScreenUpdating = True
End Sub

Sub test()
    insert_isah_testdata "[T_ProdBillOfMat$]", "[Testmultifill].[dbo].[T_ProdBillOfMat_TEST]"
End Sub

Sub test_btn_import_articles_per_pallet()
    main.init_articles_per_pallet
End Sub

Sub set_database(database_name As String)
    ' set database connection string to refer to the local test database
    ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME).Range(main.DATABASE_DROPDOWN_ADDR).value = database_name
End Sub

Sub set_for_release()
    ' remove invalid names
    r.cleanNamesWithReferenceError ThisWorkbook
    
    ' set database
    ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME).Range(main.DATABASE_DROPDOWN_ADDR).value = tests.release_database
    
    ' disable events
    Application.EnableEvents = False
    
    ' hide testing sheets
    w.hideWorksheets "tests", "TEST_DATA", "base", "planning", "test", "CAPGRP", "TESTMULTIFILL_ISAH_WK46_LNXX", "TESTMULTIFILL_ISAH_WK48", "TESTMULTIFILL_NEW_ORDERS_WK48", "TESTMULTIFILL_ISAH_PRODWK29"
    
    ' clear input and orders
    tests.test_btn_clear_sheet
    
    Worksheets(main.CONTROL_SHEET_NAME).Activate
    
    ' renable events
    Application.EnableEvents = True
    ThisWorkbook.Save
End Sub

Sub set_init_capgrp_sheets()
    Dim init_capgrp_sheet_names As collection: Set init_capgrp_sheet_names = clls.toCollection("LN 1;LN 2;LN 3;LN 5;LN 9;LN10;LN14;LN15;LN18;LN19", ";")
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' remove existing capgrp sheets
    For Each wsName In init_capgrp_sheet_names
        w.delete_worksheet wsName, ThisWorkbook
    Next
    
    ' cleanNamesWithReferenceError(optional workbook) => in workbook delete Names with #REF error
    r.cleanNamesWithReferenceError ThisWorkbook
    
    ' create new ones and try to activate
    main.init_capgrp_sheets capgrp_sheet_names:=init_capgrp_sheet_names
    Application.EnableEvents = True
    For Each wsName In init_capgrp_sheet_names
        ' try to active capgrp sheet
        ThisWorkbook.Sheets(wsName).Activate
    Next
    
    Worksheets(main.CONTROL_SHEET_NAME).Activate
    
exit_sub:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Sub test_btn_clear_sheet()
    main.btn_clear_sheet_Click
    
    ' test if `INPUT_ISAH_SHEET` sheet has been cleared
    Set wb = ThisWorkbook
    Set IsahSheet = wb.Sheets(main.INPUT_ISAH_SHEET)
    Debug.Assert r.get_last_row(IsahSheet) = 1
    
    ' test if capgrp sheets have been cleared
    For Each capgrp In main.get_capgrp_sheet_names()
       Set ws = wb.Sheets(capgrp)
       lastRow = r.get_last_row(ws)
       Debug.Print "lastRow is:", lastRow
       Debug.Assert lastRow = 1
       
       Set ordersRng = main.get_orders_range(CStr(capgrp))
       Debug.Assert ordersRng.Rows.count = 1
    Next
End Sub

Sub test_btn_import_art_ln1()
    Set wb = ThisWorkbook
    Dim capgrp_sheets As collection: Set capgrp_sheets = main.get_capgrp_sheet_names()
    Dim ordersRngRowsCount As Long
    
    ' set LN 1 year and weeknumber to 29
    main.set_capgrp_weeknumber "LN 1", 29
    main.set_capgrp_year "LN 1", 2023
    
    ' test if for all capgrp sheets, year is set to 29
    For Each capgrp In capgrp_sheets
       Debug.Assert main.get_capgrp_weeknumber(CStr(capgrp)) = 29
    Next
   
    ' import articles for all capgrps
    main.btn_import_art_Click
    
    ' test if capgrp sheet LN 1 has been filled
    capgrp = "LN 1"
    Set ws = wb.Sheets(capgrp)
    Set ordersRng = main.get_orders_range(CStr(capgrp))
    Debug.Assert r.get_last_row("A14", ws) > tests.CAPGRP_START_ROW
    Debug.Assert ordersRng.Rows.count > 1
    
    ' test if capgrp sheet <> LN 1 has no rows
    For Each capgrp In capgrp_sheets
       If capgrp = "LN 1" Then
          GoTo next_capgrp
       End If
       
       Set ws = wb.Sheets(capgrp)
       Debug.Assert r.get_last_row("A14", ws) = tests.CAPGRP_START_ROW
       Set ordersRng = main.get_orders_range(CStr(capgrp))
       ordersRngRowsCount = ordersRng.Rows.count
       Debug.Assert ordersRngRowsCount = 1
next_capgrp:
    Next
    
    Exit Sub
    
End Sub

Sub test_btn_import_art_all()
    Set wb = ThisWorkbook
    Dim capgrp_sheets As collection: Set capgrp_sheets = main.get_capgrp_sheet_names()
    Dim ordersRngRowsCount As Long
    
    ' set LN 1 year and weeknumber to 29
    main.set_capgrp_weeknumber "LN 1", 29
    main.set_capgrp_year "LN 1", 2023
    
    ' test if for all capgrp sheets, year is set to 29
    For Each capgrp In capgrp_sheets
       Debug.Assert main.get_capgrp_weeknumber(CStr(capgrp)) = 29
    Next
   
    ' import articles for all capgrps
    main.btn_import_art_Click

    ' test if capgrp sheet of capgrp with data has rows
    For Each capgrp In capgrp_sheets
       If capgrp > "LN 1" Then
          GoTo next_capgrp
       End If
       
       Set ws = wb.Sheets(capgrp)
       'TODO figure out why this one fails
       Debug.Assert r.get_last_row("A14", ws) > tests.CAPGRP_START_ROW
       Set ordersRng = main.get_orders_range(CStr(capgrp))
       ordersRngRowsCount = ordersRng.Rows.count
       Debug.Assert ordersRngRowsCount > 1
next_capgrp:
    Next
    
    Exit Sub
    
End Sub

Sub test_ln1_updated()
    Set wb = ThisWorkbook
    capgrp = "LN 1"
    Set ws = wb.Sheets(capgrp)
    Debug.Print r.get_last_row("A14", ws)
End Sub

Sub test_btn_import_bulk()
    main.btn_import_bulk_Click
End Sub

Sub test_bulk_sheet_values()
    Dim ws0 As Worksheet, bulkRange As Range, calcRange As Range
    Set ws0 = ThisWorkbook.Sheets(main.BULK_SHEET_NAME)
    Set bulkRange = r.get_range(main.BULK_ORDERS_RANGE_NAME, ws0)
    arr_bulk = bulkRange
    Set calcRange = r.get_range(main.BULK_ORDERS_CALC_RANGE_NAME, ws0)
    arr_bulk_calc = calcRange
    
    'validate dimensions
    Debug.Assert bulkRange.columns.count >= 15 And bulkRange.Rows.count = 64 And calcRange.Rows.count = 64
    
    ' get calculated values
    a.printArray a.QueryArray(arr_bulk, "Ordernr", 506676)
    firstDateLn1 = a.getNamedArrayValue(a.QueryArray(arr_bulk, "Ordernr", 506676), "Lijn 1")
    lastDateLn1 = a.getNamedArrayValue(a.QueryArray(arr_bulk, "Ordernr", 506611), "Lijn 1")
    Debug.Assert firstDateLn1 = "Ma 17 06:00" And lastDateLn1 = "Do 27 07:30"
    
    ' change workdaytimes: set first block = 0
    main.get_worktimes_values_range("LN 1").Cells(1, 1) = 0
     
    ' return to BULK sheet
    ws0.Activate
    arr_bulk = bulkRange
    firstDateLn1 = a.getNamedArrayValue(a.QueryArray(arr_bulk, "Ordernr", 506676), "Lijn 1")
    lastDateLn1 = a.getNamedArrayValue(a.QueryArray(arr_bulk, "Ordernr", 506611), "Lijn 1")
    Debug.Print firstDateLn1, lastDateLn1
    Debug.Assert firstDateLn1 = "Ma 17 08:15" And lastDateLn1 = "Do 27 09:45"
    
    ' restore worktimes LN 1
    main.get_worktimes_values_range("LN 1").Cells(1, 1) = 1
    
    ws0.Activate
    arr_bulk = bulkRange
    firstDateLn1 = a.getNamedArrayValue(a.QueryArray(arr_bulk, "Ordernr", 506676), "Lijn 1")
    lastDateLn1 = a.getNamedArrayValue(a.QueryArray(arr_bulk, "Ordernr", 506611), "Lijn 1")
    Debug.Assert firstDateLn1 = "Ma 17 06:00" And lastDateLn1 = "Do 27 07:30"
    
End Sub

Sub test_prodwk29_ln1()
    Dim current_volgnummer As Double, change_row_number As Double, volgnummer_index As Integer, articles_column_index As Integer
    
    volgnummer_index = 1
    articles_column_index = 2
    
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("LN 1")
    ws.Activate
    
    ' count orders rows (should be 83 for week 29)
    Set ordersRng = main.get_orders_range("LN 1")
    Debug.Assert ordersRng.Rows.count = 83
    'validate volgnummer
    current_volgnummer = ordersRng.Cells(2, volgnummer_index).value
    Debug.Assert current_volgnummer = 1
    
    'add/remove order records
    change_row_number = 3
    ordersRng.Cells(change_row_number, 1).Select
    current_volgnummer = ordersRng.Cells(change_row_number, volgnummer_index).value
    main.btn_add_record_Click
    Debug.Assert ordersRng.Rows.count = 84
    
    'validate volgnummer is recalculate for added record
    Debug.Assert ordersRng.Cells(change_row_number, volgnummer_index).value = current_volgnummer
    
    main.btn_delete_record_Click
    Debug.Assert ordersRng.Rows.count = 83
    'validate volgnummer is recalculate after deleted record
    Debug.Assert ordersRng.Cells(change_row_number, volgnummer_index).value = current_volgnummer
    
    ' get the startdate column number
    Dim startDateColNum As Long
    startDateColNum = r.get_column_index(ordersRng, "Starttijd")
    
    ' change workingdates
    Dim startdate0 As Date
    chk_startdate_str = "2023-07-17 06:00"
    startdate0 = ordersRng.Cells(2, startDateColNum)
    Debug.Assert dt.format_datetime(startdate0) = chk_startdate_str
    
    ' deactivate first starttime block
    main.get_worktimes_range("LN 1").Cells(2, 2) = 0
    chk_startdate_str = "2023-07-17 08:15"
    startdate0 = ordersRng.Cells(2, startDateColNum)
    Debug.Assert dt.format_datetime(startdate0) = chk_startdate_str
        
    ' reactivate first starttime block
    main.get_worktimes_range("LN 1").Cells(2, 2) = 1
    chk_startdate_str = "2023-07-17 06:00"
    startdate0 = ordersRng.Cells(2, startDateColNum)
    Debug.Assert dt.format_datetime(startdate0) = chk_startdate_str
    
    ' insert "ombouw" on row 10
    ordersRng.Rows(10).Select
    main.btn_add_record_Click
    'ordersRng = main.get_orders_range("LN 1")
    ordersRng.Rows(10).Select
    ordersRng.Cells(10, 4) = "ombouw"
    Debug.Assert main.get_orders_range("LN 1").Rows.count = 84

    ' count the Artikelen on LN1
    Debug.Assert WorksheetFunction.CountA(main.get_orders_range("LN 1").columns(articles_column_index)) = NUMBER_ARTICLES_LN1
    
    'export LN1
    Dim pdf_path As String
    main.btn_export_pdf_Click
    pdf_path = main.get_capgrp_print_location("LN 1")
    Debug.Assert fs.pathExist(pdf_path)
    fs.deleteFilePath pdf_path
    
    'print LN1
    main.btn_print_dates_Click
    
End Sub

Sub test_pdf_export()
    Dim pdf_path As String
    ThisWorkbook.Worksheets("LN 1").Activate
    main.btn_export_pdf_Click
    pdf_path = main.get_capgrp_print_location("LN 1")
    Debug.Assert fs.pathExist(pdf_path)
    fs.deleteFilePath pdf_path
    
End Sub

Sub test_add_capgrp()
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    ' add zero capgrps
    num_sheets = wb.Sheets.count
    main.btn_add_capgrp_sheets_Click
    Debug.Assert num_sheets = wb.Sheets.count
    
    ' remove capgrp sheet LN18 and re-add
    w.delete_worksheet "LN18", wb
    main.btn_add_capgrp_sheets_Click
    Debug.Assert num_sheets = wb.Sheets.count
    Debug.Assert ActiveSheet.name = main.CONTROL_SHEET_NAME
End Sub

Sub test_state_control()
    Dim wb0 As Workbook, ordersRng As Range, worktimesRng As Range, capgrp_sheet As String
    Set wb0 = ThisWorkbook
    
    wb0.Sheets("overzicht").Activate
    For Each key In main.get_capgrp_sheet_names()
        capgrp_sheet = key
        ' TODO FIX exception LN19
        If key = "LN 1" Or key = "LN19" Then
           GoTo nx_capgrp
        End If
        Debug.Print "Testing state control on sheet: " & capgrp_sheet
        
        ' start with a fresh sheet
        main.clear_orders_range capgrp_sheet
        main.set_worktimes_range_values capgrp_sheet, 1
        
        ' import orders and restore previous state
        orders_arr_prev = main.get_orders_range(capgrp_sheet).formula
        worktimes_arr_prev = main.get_worktimes_range(capgrp_sheet).Value2
        main.btn_update_isah_data_Click
        
        ' check if orders_range is filled
        address = main.get_orders_range(capgrp_sheet).address
        Debug.Print address
        If address = "$A$14" Then
           GoTo nx_capgrp
        End If
        
        capgrpStateArray = getCapgrpStateArray(-2, True)  'last state is the state before the previous state
        orders_arr_prev = capgrpStateArray(0)
        
        main.btn_restore_prev_state_Click
        orders_arr_cur = main.get_orders_range(capgrp_sheet).Value2
        worktimes_arr_cur = main.get_worktimes_range(capgrp_sheet).Value2
        
        ' check if imported orders have been cleared after restore previous state
        Debug.Print TypeName(orders_arr_prev), TypeName(orders_arr_cur)
        'Debug.Assert orders_arr_cur = orders_arr_prev And TypeName(orders_arr_prev) = "Empty"
        Debug.Assert a.ArraysAreEqual(worktimes_arr_cur, worktimes_arr_prev)
        
        ' again import orders
        main.btn_update_isah_data_Click
        tests.check_state_after_restore "tests.state_change_aantal", capgrp_sheet
        tests.check_state_after_restore "tests.state_change_set_worktime", capgrp_sheet
        
        ' restore after worksheet buttons/actions
        ActiveSheet.Range("A15").Select
        tests.check_state_after_restore "main.btn_add_record_Click", capgrp_sheet
        tests.check_state_after_restore "main.btn_delete_record_Click", capgrp_sheet
        
        
        'Exit Sub
        'Debug.Assert orders_range_prev = orders_range_cur
nx_capgrp:
    Next
End Sub


' This subroutine stores the current state of the active worksheet, executes a given subroutine,
' restores the previous state, and compares the current state with the stored state.
'
' Parameters:
' sub_name : The name of the subroutine to execute.
'
' Example usage:
' check_state_after_restore "MySubroutine"
Sub check_state_after_restore(sub_name As String, capgrp_sheet As String)
    
    ' Store the current state of the active worksheet (e.g., value in cell A1)
    orders_arr_prev = main.get_orders_range(capgrp_sheet).formula
    worktimes_arr_prev = main.get_worktimes_range(capgrp_sheet).Value2
    
    ' Execute the subroutine by name
    Application.Run sub_name
    
    ' Call the restore subroutine to restore the previous state
    main.btn_restore_prev_state_Click
    
    ' Get the current state after restoration
    orders_arr_cur = main.get_orders_range(capgrp_sheet).formula
    worktimes_arr_cur = main.get_worktimes_range(capgrp_sheet).Value2
    
    ' Compare the current state with the stored previous state
    Debug.Print "Comparing restored state after subroutine " & sub_name
    Debug.Assert a.ArraysAreEqual(orders_arr_prev, orders_arr_cur, True)
    Debug.Assert a.ArraysAreEqual(worktimes_arr_prev, worktimes_arr_cur, True)
    
End Sub

Sub state_change_aantal()
    ActiveSheet.Range("H15").value = 999
End Sub

Sub state_change_set_worktime()
    Dim worktimesValuesRng As Range
    Set worktimesValuesRng = main.get_worktimes_values_range(ActiveSheet.name)
    worktimesValuesRng.Cells(1, 1).value = 0
End Sub

' ISAH tests
Sub test_isah_imports()
   Dim checkMsgText As String
   ' import the number of articles per pallet to sheet `aantal_per_pallet` and check the message box
   main.btn_isah_update_aantal_per_pallet_Click
   
   'checkMsgText = "Sheet geupdatet: " & main.NUMBER_PER_PALLET_SHEET_NAME
   'Debug.Assert ctr.CheckMessageBox(checkMsgText) = True
End Sub

Sub test_msgbox()
    MsgBox "Sheet geupdatet: " & main.NUMBER_PER_PALLET_SHEET_NAME
    'Debug.Print ctr.HasMsgBox()
End Sub

Sub test_isah_export()
    main.btn_isah_database_export_Click
    Debug.Assert ActiveSheet.name = "EXPORT_ISAH"
End Sub

Sub test_buttons_capgrp()
    capgrp_sheet = "LN 1"
    Set wb = ThisWorkbook
    Set ws = wb.Sheets(capgrp_sheet)
    ws.Activate
    
    ' import previous prodweek ISAH data to capgrp worksheet
    Set rng = main.get_weeknumber_range(ws.name)
    rng.Cells(2, 2).value = tests.previous_prodwk
    
    'check if other capgrp_sheets are set to same prodwk
    weeknumber_2 = main.get_weeknumber_range("LN 2").Cells(2, 2).value
    Debug.Assert tests.previous_prodwk = weeknumber_2
    
    ws.Activate
    
    ' import ISAH data to capgrp worksheet
    main.btn_update_isah_data_Click
    Set ordersRng = main.get_orders_range(capgrp_sheet)
    Debug.Assert ordersRng.Rows.count <= 1
    
    ' set prodweek ISAH to current_prodwk and import ISAH
    Set rng = main.get_weeknumber_range(ws.name)
    rng.Cells(2, 2).value = tests.current_prodwk
    main.btn_update_isah_data_Click
    Set ordersRng = main.get_orders_range(capgrp_sheet)
    Debug.Assert ordersRng.Rows.count > 1
    
    ' add record and count number of records in ordersRng
    r0 = ordersRng.Rows.count
    Range("A15").Select
    main.btn_add_record_Click
    r1 = main.get_orders_range(capgrp_sheet).Rows.count
    Debug.Assert r0 + 1 = r1
    
    main.btn_delete_record_Click
    r2 = main.get_orders_range(capgrp_sheet).Rows.count
    Debug.Assert r2 = r0
    
    ' copy and insert first row and recalculate dates
    Set ordersRng = main.get_orders_range(capgrp_sheet)
    ordersRng.Rows(2).Select
    Selection.Copy
    Selection.Insert shift:=xlDown
    
    main.btn_calculate_dates_Click
    
    Selection.Delete shift:=xlUp
    
    main.btn_calculate_dates_Click
    
    'change working times
    Set worktimesRng = main.get_worktimes_range(capgrp_sheet)


End Sub

' DATABASE TESTS
Sub test_multifill_rds()
    Dim connstr As String: connstr = "Driver={ODBC Driver 11 for SQL Server};Server=MF-ERP\MSSQLSERVER_ISAH;User Id=PlanningExcel;Password=a2AZCY8mkr&Qt5#LbB"
    Dim conn As ADODB.Connection, sql0 As String
    Set conn = db.openDBconn(connstr:=connstr)
    sql0 = "SELECT 1"
    db.printRecordset db.queryDB(conn, sql0)
    On Error GoTo close_connection
    
    GoTo close_connection
close_connection:
    conn.Close
End Sub

' TESTDATA
Sub insert_isah_testdata(source_table As String, target_table As String, Optional print_statement As Boolean = True)
    
    ' get testdata from excel as recordset
    Dim fl0 As File
    Dim fs0 As New filesystemobject: Set fs0 = New filesystemobject
    filePath = os.pathJoin(zz_env.getVDMITestPath(), "ISAH_mock_tables.xlsx")
    Set fl0 = fs0.GetFile(filePath)
    
    Dim conn As ADODB.Connection, rs0 As New ADODB.Recordset, sqlconn As ADODB.Connection
    Set conn = db.openExcelConn(fl0)
    
    Dim sql0 As String
    sql0 = "SELECT * FROM " & source_table & ";"
    Set rs0 = db.queryDB(conn, sql0)
    
    'create the insert statement string
    Dim insert_statement As String
    insert_statement = db.sqlInsertStatement(rs0, target_table, mssql)
    Debug.Print insert_statement
    
    'clean up
    Set fs0 = Nothing
    conn.Close
        
    'insert data into sql table
    ' execute `insert_statement` against MSSQL_HOME
    Set sqlconn = db.openDBconn(main.getISAHconnstr())
    
    'clear previous test data
    db.truncateTable sqlconn, target_table
    
    Dim statements As New collection
    Set statements = str.stringToCol(insert_statement, db.MSSQL_LINE_BREAK)
    
    c = 0
    For Each stat In statements
        If print_statement Then
           Debug.Print stat
        End If
        sqlconn.Execute CStr(stat)
        c = c + 1
    Next
    
    'clean up
    sqlconn.Close

End Sub

Sub update_isah_testdata(source_table As String, target_table As String, set_columns As String, where_columns As String)
    Dim conn As ADODB.Connection, sqlconn As ADODB.Connection
    filePath = os.pathJoin(zz_env.getVDMITestPath(), "ISAH_mock_tables.xlsx")
    Set conn = db.openExcelConn(filePath)
    
    Dim rs0 As ADODB.Recordset
    Set rs0 = db.queryDB(conn, "SELECT * FROM " & source_table & ";")
    
    Dim update_statement As String
    update_statement = db.sqlUpdateStatement(rs0, target_table, set_columns, where_columns, mssql)
    Debug.Print update_statement
    
    conn.Close
    
    Set sqlconn = db.openDBconn(main.getISAHconnstr())
    executeSqlStatements sqlconn, update_statement, db.MSSQL_LINE_BREAK
    sqlconn.Close
End Sub

Sub insert_update_isah_testdata_tables()
    If Not u.InList(main.getISAHProfileName(), "JKR;JKR2") Then
       Err.Raise 1001, "check profile name"
    End If
    insert_isah_testdata "[T_ProductionHeader$]", "[Testmultifill].[dbo].[T_ProductionHeader_TEST]"
    insert_isah_testdata "[T_ProdBillOfOper$]", "[Testmultifill].[dbo].[T_ProdBillOfOper_TEST]"
    insert_isah_testdata "[T_ProdBillOfMat$]", "[Testmultifill].[dbo].[T_ProdBillOfMat_TEST]"
    insert_isah_testdata "[T_Part_BasicMat$]", "[Testmultifill].[dbo].[T_Part_BasicMat]", False
    update_isah_testdata "[Update_T_ProductionHeader$]", "[Testmultifill].[dbo].[T_ProductionHeader_TEST]", "StartDate;EndDate", "ProdHeaderDossierCode"
End Sub

Sub insert_isah_testdata_wk29()
    If Not u.InList(main.getISAHProfileName(), "JKR;JKR2") Then
       Err.Raise 1001, "check profile name"
    End If
    insert_isah_testdata "[T_ProductionHeader_wk29$]", "[Testmultifill].[dbo].[T_ProductionHeader_TEST]"
    insert_isah_testdata "[T_ProdBillOfOper_wk29$]", "[Testmultifill].[dbo].[T_ProdBillOfOper_TEST]"
    insert_isah_testdata "[T_ProdBillOfMat_wk29$]", "[Testmultifill].[dbo].[T_ProdBillOfMat_TEST]"
End Sub

Sub test_isah_staging_ln1()
    Dim ordersStagingRng As Range
    main.isah_export_stage_orders
    'check the number of product orders for LN1
    ordersStagingArr = r.get_range("isah_staging_orders_range")
    ordersStagingArrLN1 = a.QueryArray(ordersStagingArr, "CAPGRP", "LN 1")
    num_rows = a.numArrayRows(ordersStagingArrLN1)
    Debug.Assert num_rows = tests.NUMBER_ARTICLES_LN1
    
End Sub

Sub set_input_isah(wsName)
    Dim ws0 As Worksheet, ws1 As Worksheet
    Set wb = ThisWorkbook
    Set ws1 = wb.Sheets(main.INPUT_ISAH_SHEET)
    w.clearWorksheet ws1
    Set ws0 = wb.Sheets(wsName)
    Set rng0 = r.expand_range("A1", ws0)
    r.copy_range rng0, "A1", ws1
    
    'format the columns
    main.format_isah_input_range
End Sub

Sub set_new_orders_data(wsName)
    Dim ws0 As Worksheet, ws1 As Worksheet
    Set wb = ThisWorkbook
    Set ws1 = wb.Sheets(main.ISAH_NEW_ORDERS_SHEET_NAME)
    w.clearWorksheet ws1
    Set ws0 = wb.Sheets(wsName)
    Set rng0 = r.expand_range("A1", ws0)
    r.copy_range rng0, "A1", ws1
End Sub

Sub set_input_isah_to_wk29()
    Call tests.set_input_isah(wsName:="TESTMULTIFILL_ISAH_PRODWK29")
End Sub

Sub set_input_isah_to_wk29_ln1()
    Call tests.set_input_isah(wsName:="TESTMULTIFILL_ISAH_WK29_LN1")
End Sub

Sub set_input_isah_to_wk29_bulk()
    Call tests.set_input_isah(wsName:="TESTMULTIFILL_ISAH_WK29_U000")
End Sub

Sub set_input_isah_to_wk46_new_capgrps()
    Call tests.set_input_isah(wsName:="TESTMULTIFILL_ISAH_WK46_LNXX")
End Sub

Sub set_input_isah_to_wk48()
    Call tests.set_input_isah(wsName:="TESTMULTIFILL_ISAH_WK48")
End Sub

Sub set_input_new_orders_to_wk48()
    Call tests.set_new_orders_data(wsName:="TESTMULTIFILL_NEW_ORDERS_WK48")
End Sub


Sub test_queries()
Dim conn0 As ADODB.Connection, sql0 As String
Set conn0 = db.openExcelConn(ThisWorkbook)
sql0 = "SELECT a.ProdHeaderOrdNr, a.ProdHeaderDossierCode, a.next_StartDate_header, b.min_bom_required_date, b.max_bom_required_date , SWITCH(a.[StartDate_header] = b.[max_bom_required_date],1,1=1,0) as check_bom_required_date FROM [EXPORT_ISAH$] a LEFT JOIN [CHECK_PROD_BILL_OF_MAT$] b ON a.ProdHeaderDossierCode=b.ProdHeaderDossierCode"
sql0 = "SELECT a.ProdHeaderOrdNr, a.ProdHeaderDossierCode, a.next_StartDate_header, b.min_bom_required_date, b.max_bom_required_date , SWITCH(a.[StartDate_header] = b.[max_bom_required_date],1,1=1,0) as check_bom_required_date FROM [EXPORT_ISAH$] a LEFT JOIN [CHECK_PROD_BILL_OF_MAT$] b ON a.ProdHeaderDossierCode=b.ProdHeaderDossierCode"

Set rs0 = db.queryFromWorkbook(sql0, conn0)
conn0.Close
End Sub

Sub test_last_art_capgrp()
    Dim col0 As collection
    Set col0 = a.as_collection(main.get_art_capgrps())
    Debug.Print col0.item(col0.count)
End Sub

Sub test_set_database()
tests.set_database "JKR"
Debug.Print main.getISAHconnstr, main.getISAHdbname
End Sub

' Workbook connection tests
Sub test_workbook_connection()
    Debug.Assert WorkBookConnection Is Nothing
    Debug.Print main.getWorkbookConnection().Attributes
    Set wbconn = main.getWorkbookConnection()
    Debug.Assert wbconn Is WorkBookConnection
    wbconn.Close
    Set wbconn = Nothing
    Set WorkBookConnection = Nothing
    Debug.Assert WorkBookConnection Is Nothing
End Sub

