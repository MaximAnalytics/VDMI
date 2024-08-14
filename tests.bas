'tests for planning_automatisering_macro.xlsm
'0. initialize: set data base clear all sheets
'1. import ISAH data capgrp sheets

Global Const testdatabase = "JKR"
Global Const release_database = "TEST"
Global Const current_prodwk = 29
Global Const previous_prodwk = 28
Global Const CAPGRP_START_ROW = 14
Global Const NUMBER_ARTICLES_LN1 = 83

Dim wb As Workbook, capgrp_sheet As String, ws As Worksheet, ordersRng As Range, worktimesRng As Range
Dim IsahSheet As Worksheet, BulkSheet As Worksheet

Const P_RELEASE As Boolean = True

Sub test_all()
    ' 0. Initialize:
    tests.set_database tests.testdatabase
    tests.test_btn_clear_sheet
    
    ' 1. FLOW: INPUT PRODWK29 LN1
    tests.set_input_isah_to_wk29_ln1
    tests.test_btn_import_art_ln1

    ' FLOW: PRODWK29
    tests.set_input_isah_to_wk29
    tests.test_add_capgrp ' remove and re-add LN18
    tests.test_btn_import_art_all
    tests.test_btn_import_bulk
    tests.test_bulk_sheet_values
    
    ' ZOOM-IN: LN1
    tests.test_prodwk29_ln1
    
    ' ISAH import/export
    If testdatabase = "JKR" Then
       tests.insert_update_isah_testdata_tables
    End If
    tests.test_isah_staging_ln1
    tests.test_isah_imports
    tests.test_isah_export
    
    ' State control
    tests.test_state_control
    
    If P_RELEASE Then
       tests.set_for_release
    End If
    
Exit Sub

End Sub

Sub test_btn_import_articles_per_pallet()
    'main.BTN_ADD_CAPGRP_ADDR)
End Sub

Sub set_database(database_name As String)
    ' set database connection string to refer to the local test database
    ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME).Range(main.DATABASE_DROPDOWN_ADDR).value = database_name
End Sub

Sub set_for_release()
    ' disable events
    Application.EnableEvents = False
    ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME).Range(main.DATABASE_DROPDOWN_ADDR).value = tests.release_database
    
    ' hide testing sheets
    w.hideWorksheets "tests", "TEST_DATA", "base", "planning", "test"
    
    ' clear input and orders
    tests.test_btn_clear_sheet
    
    Worksheets(main.CONTROL_SHEET_NAME).Activate
    
    ' renable events
    Application.EnableEvents = True
    ThisWorkbook.Save
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

Sub test_isah_staging_ln1()
    Dim ordersStagingRng As Range
    main.isah_export_stage_orders
    'check the number of product orders for LN1
    ordersStagingArr = r.get_range("isah_staging_orders_range")
    ordersStagingArrLN1 = a.QueryArray(ordersStagingArr, "CAPGRP", "LN 1")
    num_rows = a.num_array_rows(ordersStagingArrLN1)
    Debug.Assert num_rows = tests.NUMBER_ARTICLES_LN1
    
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
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("LN 1")
    ws.Activate
    
    ' count orders rows (should be 83 for week 29)
    Set ordersRng = main.get_orders_range("LN 1")
    Debug.Assert ordersRng.Rows.count = 83
    
    'add/remove order records
    ordersRng.Cells(3, 1).Select
    main.btn_add_record_Click
    Debug.Assert ordersRng.Rows.count = 84
    
    main.btn_delete_record_Click
    Debug.Assert ordersRng.Rows.count = 83
    
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
    Debug.Assert WorksheetFunction.CountA(main.get_orders_range("LN 1").columns(1)) = NUMBER_ARTICLES_LN1
    
    'print LN1
    main.btn_print_dates_Click
    
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
        If key <> "LN 1" Then
           'GoTo nx_capgrp
        End If
        Debug.Print "Testing state control on sheet: " & capgrp_sheet
        
        ' start with a fresh sheet
        main.clear_orders_range capgrp_sheet
        main.set_worktimes_range_values capgrp_sheet, 1
        
        ' import orders and restore previous state
        orders_arr_prev = main.get_orders_range(capgrp_sheet).formula
        worktimes_arr_prev = main.get_worktimes_range(capgrp_sheet).Value2
        main.btn_update_isah_data_Click
        
        capgrpStateArray = getCapgrpStateArray(-2, True)  'last state is the state before the previous state
        orders_arr_prev = capgrpStateArray(0)
        Debug.Print TypeName(orders_arr_prev)
        'Exit Sub
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

Sub test_()
Dim wb0 As Workbook, rng As Range, rng1 As Range, capgrp_sheet As String, numCols As Long, numRows As Long
capgrp_sheet = "INPAK"
Set ws = Worksheets(capgrp_sheet)
ws.Activate
ws.Range("A15").Activate

values = main.get_orders_range(capgrp_sheet).Value2
Set rng = main.get_orders_range(capgrp_sheet)

main.btn_add_record_Click
numCols = a.num_array_columns(values)
numRows = a.num_array_rows(values)
Set rng1 = r.getResizedRange(rng, num_rows:=numRows, num_cols:=numCols)
Debug.Print rng.address, rng1.address


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
    'TODO: use function a.InList(dbprofile, JKR;JKR2)
    If main.getISAHProfileName() <> "JKR" And main.getISAHProfileName() <> "JKR2" Then
       Err.Raise 1001, "check profile name"
    End If
    insert_isah_testdata "[T_ProductionHeader$]", "[Testmultifill].[dbo].[T_ProductionHeader_TEST]"
    insert_isah_testdata "[T_ProdBillOfOper$]", "[Testmultifill].[dbo].[T_ProdBillOfOper_TEST]"
    insert_isah_testdata "[T_ProdBillOfMat$]", "[Testmultifill].[dbo].[T_ProdBillOfMat_TEST]"
    insert_isah_testdata "[T_Part_BasicMat$]", "[Testmultifill].[dbo].[T_Part_BasicMat]", False
    update_isah_testdata "[Update_T_ProductionHeader$]", "[Testmultifill].[dbo].[T_ProductionHeader_TEST]", "StartDate;EndDate", "ProdHeaderDossierCode"
End Sub

Sub test()
    insert_isah_testdata "[T_ProdBillOfMat$]", "[Testmultifill].[dbo].[T_ProdBillOfMat_TEST]"
End Sub

Sub set_input_isah(wsName)
    Dim ws0 As Worksheet, ws1 As Worksheet
    Set wb = ThisWorkbook
    Set ws1 = wb.Sheets(main.INPUT_ISAH_SHEET)
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

Sub test_queries()
Dim conn0 As ADODB.Connection, sql0 As String
Set conn0 = db.openExcelConn(ThisWorkbook)
sql0 = "SELECT a.ProdHeaderOrdNr, a.ProdHeaderDossierCode, a.next_StartDate_header, b.min_bom_required_date, b.max_bom_required_date , SWITCH(a.[StartDate_header] = b.[max_bom_required_date],1,1=1,0) as check_bom_required_date FROM [EXPORT_ISAH$] a LEFT JOIN [CHECK_PROD_BILL_OF_MAT$] b ON a.ProdHeaderDossierCode=b.ProdHeaderDossierCode"
sql0 = "SELECT a.ProdHeaderOrdNr, a.ProdHeaderDossierCode, a.next_StartDate_header, b.min_bom_required_date, b.max_bom_required_date , SWITCH(a.[StartDate_header] = b.[max_bom_required_date],1,1=1,0) as check_bom_required_date FROM [EXPORT_ISAH$] a LEFT JOIN [CHECK_PROD_BILL_OF_MAT$] b ON a.ProdHeaderDossierCode=b.ProdHeaderDossierCode"

Set rs0 = db.queryFromWorkbook(sql0, conn0)
conn0.Close
End Sub

'Sub query_from_isah()
'    Dim conn0 As ADODB.Connection, sql0 As String
'    Set conn0 = main.getISAHprodbom()
'End Sub

' test state_control
' for all capgrp sheets: clear state -> reverse the following steps: import isah data

' extra tests
'Sub test_2()
'    ' update_ranges
'    Dim rng1 As Range, unique_bulkcodes As collection, rng0 As Range
'    Set wb0 = ThisWorkbook
'    capgrp = "INPK"
'    range_name = "INPK_orders"
'    Set rng0 = wb0.Names(range_name).RefersToRange
'    Debug.Print rng0.Rows(1).row
'    'Set ws0 = Worksheets("INPK")
'
'    capgrp = "LN 1"
'    init_buttons capgrp
'    Exit Sub
'    Set rng0 = get_range(main.ORDERS_RANGE_ADDR, ws:=ThisWorkbook.Worksheets(capgrp))
'    Debug.Print r.get_last_row(rng0), r.get_last_col(rng0)
'    Exit Sub
'
'    Set rng0 = r.expand_range(main.ORDERS_RANGE_ADDR, ws:=ThisWorkbook.Worksheets(capgrp), c1:=10)
'    Exit Sub
'    Debug.Print main.ORDERS_RANGE_ADDR, rng0.address
'    r1 = r.get_last_row(rng0, ws:=rng0.Worksheet)
'    Debug.Print r1
'    'ctr.move_button "btn_update_isah_data_" & capgrp, main.BTN_UPDATE_ISAH_ADDRESS, ws0
'End Sub

Sub test_last_art_capgrp()
    Dim col0 As collection
    Set col0 = a.as_collection(main.get_art_capgrp())
    Debug.Print col0.item(col0.count)
End Sub
'
'Sub test_capgrpState()
'    Dim capgrp_sheet As String, CapgrpState As collection, CapgrpStates As collection, col1 As collection, wb As Workbook
'    Set wb = ThisWorkbook
'    capgrp_sheet = "LN 1"
'    wb.Sheets(capgrp_sheet).Activate
'
'    Dim Counter As collection
'    Set Counter = clls.toCollection(a.create_integer_vector(1, 10))
'    For Each c In Counter
'        storeCapgrpState
'        Debug.Print c, tests.getCapgrpStates(capgrp_sheet).count
'    Next
'
'    Exit Sub
'
'    tests.storeCapgrpState
'    Debug.Print WorksheetStateCollection.count
'    Set CapgrpStates = WorksheetStateCollection(capgrp_sheet)
'    Debug.Print CapgrpStates.count
'
'    capgrpStateArray = clls.getItem(CapgrpStates, -6)
'    'capgrpStateArray = getCapgrpStateArray(-4)
'    If a.num_array_rows(capgrpStateArray) > -1 Then
'        orders_arr_prev = capgrpStateArray(0)
'        worktimes_arr_prev = capgrpStateArray(1)
'        a.printArray orders_arr_prev
'        'a.printArray worktimes_arr_prev
'    Else
'     Debug.Print "capgrpStateArray is -1"
'    End If
'
'
'
'End Sub

