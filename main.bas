' Define enumerations

' Define global constants
' modes
Global Const END_USER_MODE As Boolean = False
Global Const PRINT_AS_PDF As Boolean = True

' Error handling
Global Const GENERIC_ERR_MSG As String = "Fout opgetreden"
Global Const ERR_MSG_ISAH_INPUT_EMPTY As String = "Geen ISAH orders aanwezig op Template"
Global Const ERR_MSG_NO_ARTICLES As String = "Geen artikelen gevonden in T_Part"
Global Const ERR_MSG_NO_BULK_CODES As String = "Geen ISAH bulken aanwezig op Template"

' Settings (P_)
Global Const P_DEBUG = True
Global Const test_mode = "TEST" 'values are: HOME, TEST, PROD
Global Const WORKSHEET_IGNORE_MULTIROW_EVENTS = False

' Persistence
Public WorksheetStateCollection As New collection
Global Const P_STORE_STATE = True
Global Const P_NUM_STORED_STATES = 5

' ISAH INPUT COLUMNS
Global Const INPUT_DATA_HEADER As String = "Artikel,Ordernummer,Productieorder,Omschrijving,Bulkcode,Aantal,Qty1,Duur,Resources,Flesformaat,Sluiting,Pallettype,ProdWk,Land,Oproep"
Global Const INPUT_DATA_FORMATS As String = "General,General,General,General,General,0,0,0.00,General,General,General,General,General,General,General,General,General,General"
Global Const ISAH_TEMPLATE_COLUMN_FORMATS As String = "Artikel=General;Ordernummer=General;Productieorder=General;Omschrijving=General;Bulkcode=General;Aantal=0;Qty1=0;Duur=0.00;Resources=General;Flesformaat=General;Sluiting=General;Pallettype=General;ProdWk=General;Land=General;Oproep=General"
Global Const OUTPUT_DATA_FORMATS As String = "General,General,General,General,General,0,0,0.00,General,General,General,General,General,General,General,ddd hh:mm,ddd hh:mm"
Global Const ORDERS_RANGE_COLUMN_FORMATS As String = "Aantal=0;Qty1=0;Duur=0.00;Starttijd=ddd hh:mm;Eindtijd=ddd hh:mm"
Global Const INPUT_ISAH_SHEET = "Template"
Global Const PRODWK_COLUMN As String = "ProdWk"
Global Const CAPGRP_COLUMN_NAME As String = "Cap.Grp"
Global Const BULKCODE_COLUMN As String = "Bulkcode"
Global Const STARTDATE_COLUMN As String = "Starttijd"
Global Const DESCRIPTION_COLUMN As String = "Omschrijving"
Global Const ENDDATE_COLUMN As String = "Eindtijd"
Global Const DURATION_COLUMN As String = "Duur"
Global Const ART_COLUMN As String = "Artikel"
Global Const INPUT_ISAH_SHEET_SORT_KEY As String = "Cap.Grp,Bulkcode,Aantal,Flesformaat"
Global Const COLUMNS_HIDE_FOR_PRINT As String = "Flesformaat,Sluiting"
Global Const QTY_COLUMN As String = "Qty1"
Global Const INPUT_CAPGRP_COLUMN_INDEX = 20

' LAYOUT
Global Const ORDERS_RANGE_ADDR As String = "A14"
Global Const WEEKNUMBER_RANGE_ADDRESS As String = "C2"
Global Const WEEKNUMBER_RANGE_IDS As String = ",Weeknummer"
Global Const YEAR_RANGE_ADDRESS As String = "C4"
Global Const YEAR_RANGE_IDS As String = ",Jaar"
Global Const WORKDAYS_RANGE_NAME As String = "workdays_input_rng"
Global Const WORKDAYS_RANGE_HEADER As String = ",06:00-08:00,08:15-10:15,10:30-12:30,13:00-15:00,15:15-17:15,17:45-20:15,20:30-23:00"
Global Const MAP_WORKDAYS_TIMES_TO_LABEL As String = "06:00-08:00=B1;08:15-10:15=B2;10:30-12:30=B3;13:00-15:00=B4;15:15-17:15=B5;17:45-20:15=B6;20:30-23:00=B7"
Global Const numberOfWorkTimeBlocks As Integer = 7
Global Const XL_MAX_NUMBER_COLUMNS As Long = 26

Global Const WORKDAYS_RANGE_IDS As String = ",ma,di,woe,do,vrij,ma2,di2,woe2,do2,vrij2"
Global Const WORKDAYS_RANGE_ADDRESS As String = "E2"
Global Const ORDERS_RANGE_MAX_COLUMN_NUMBER As Integer = 15
Global Const ORDERS_RANGE_VOLGNUMMER_COLUMN = "volgnummer"

Global Const BTN_UPDATE_ISAH_ADDRESS As String = "A2"
Global Const BTN_ADD_RECORD_ADDR As String = "A5"
Global Const BTN_DELETE_RECORD_ADDR As String = "A8"
Global Const BTN_RESTORE_PREV_STATE_ADDR As String = "A11"
Global Const BTN_EXPORT_PDF_ADDR As String = "C8"
Global Const BTN_PRINT_DATES_ADDR As String = "C11"
Global Const BTN_CALCULATE_DATES_ADDR As String = "D11"
Global Const BTN_WIDTH = 90
Global Const BTN_HEIGHT = 30
Global Const BTN_LEFT_OFFSET = 20
Global Const BTN_SECOND_COLUMN_LEFT_OFFSET = 9

' BUTTONS
Global Const BTN_RESTORE_PREV_STATE_LABEL = "Ga terug"

' ORDERS sheet layout
Global Const N_TOP_ROWS_FREEZE As Integer = 14
Global Const INPUT_FIELD_COLOR = 65535
Global Const WT_HEADER_COLOR = 15123099
Global Const WT_IDS_COLOR = 11389944
Global Const WT_VALUES_COLOR = 13431551
Global Const TANK_LO_COLUMN_WIDTH = 8.5
Global Const COLUMN_AB_WIDTH = 70
Global Const COLUMN_CD_WIDTH = 110

'PRINT LAYOUT
Global Const PRINT_SHEET_NAME = "PRINT"
Global Const PRINT_ORDERS_ADDRESS = "A11"
Global Const PRINT_WORKDAYS_ADDRESS = "J2"
Global Const PRINT_EXTRA_INFO_ADDRESS = "S2:U9"
Global Const PRINT_EXTRA_INFO_FIELDS = "Field1;Field2;Field3"
Global Const PRINT_FILE_RANGE_ADDRESS As String = "C6"
Global Const PRINT_FILE_RANGE_IDS As String = ",Locatie PDF bestand"
Global Const PRINT_FILE_RANGE_HEADER As String = ","
Global Const PRINT_RENAME_COLUMNS = "volgnummer=#;Ordernummer=OrderNr;Productieorder=ProdOrd;Resources=Res;# pallets=#Plts;Flesformaat=Fles;Sluiting=Slt;ProdWk=Wk;Land=Ld;Pallettype=PltType"
Global Const PRINT_TITLE_START = "A2"
Global Const PRINT_TITLE_RANGE = "A2:H6"
Global Const PRINT_SHOW_PDF_EXPORTED As Boolean = True

' COMPONENTS
Global Const BASE_SHEET_NAME As String = "base"
Global Const START_SHEET_NAME As String = "instructie"

' DEFAULTS
Global Const DEFAULT_PRINT_FILE_PATH = ""

' BULK ORDERS SHEET
Global Const BULK_ORDERS_HEADER = "Productieorder,Artikel,Qty1"
Global Const BULK_SHEET_NAME = "BULK"
Global Const BULK_SHEET_START_ADDR = "A10"
Global Const BULK_DATETIME_COLUMN_FORMAT = "yyyy-mm-dd hh:mm"
Global Const BULK_ORDER_SHEET_HEADER = "Ordernr;Bulkcode;Omschrijving;Hoeveelheid;Opmerking"
Global Const BULK_ORDER_SHEET_ORDERNR = "Ordernr"
Global Const BULK_ORDER_SHEET_BULKCODE = "Bulkcode"
Global Const BULK_ORDER_SHEET_QTY = "Hoeveelheid"
Global Const BULK_ORDER_SHEET_DSC = "Omschrijving"
Global Const BULK_ORDERS_RANGE_NAME As String = "bulk_orders_range"
Global Const BULK_ORDERS_SORTING_COLUMN As String = "SORTERING"
Global Const BULK_ORDERS_CALC_RANGE_NAME As String = "bulk_orders_calc_range"

Global Const CAPGRP_BULKCODE_STARTDATE_RANGE = "$H:$T" 'range on capgrp containing columns Bulkcode ... Starttijd
Global Const STARTDATE_COLUMN_INDEX = 13

' CONTROL SHEET
Global Const CONTROL_SHEET_NAME = "overzicht"
Global Const BTN_IMPORT_ART_ADDR = "B2"
Global Const BTN_IMPORT_BULK_ADDR = "B8"
Global Const BTN_ADD_CAPGRP_ADDR = "B14"
Global Const BTN_ISAH_EXPORT_ADDR = "B20"
'Global Const BTN_CLEAR_SHEET = "B26"
Global Const BTN_ADD_NEW_ORDERS_ADDR = "D12"
Global Const CONNECTION_STRINGS_NAMED_RANGE = "settings_connection_strings"
Global Const DATABASE_DROPDOWN_LABEL_ADDR = "H2"
Global Const DATABASE_DROPDOWN_ADDR = "H3"
Global Const DATABASE_NAME_COLUMN_NAME = "naam"
Global Const SELECTED_CONNECTION_STRING_ADDR = "H4"
Global Const SELECTED_DATABASE_NAME_ADDR = "H5"

' METADATA
Global Const NUMBER_PER_PALLET_SHEET_NAME = "aantal_per_pallet"
Global Const NUMBER_PER_PALLET_NAMED_RANGE = "metadata_number_per_pallet"
Global Const NUMBER_OF_PALLETS_NAME = "# pallets"
Global Const INPAK_CAPGRP_LIST = "XXX" '"NIVP;VDAM;INPK;SPEC;HCM "
Global Const CAPGRP_SHEET_PATTERN = "^LN(.*?)"
Global Const MAP_NUMBER_PER_PALLET_COL_TO_FMT = "Artikel=@;Omschrijving=General;aantal_per_pallet=0"
Global Const GAPGRP_FILTER_SHEET = "CAPGRP"

' ISAH database
Global Const ISAH_STAGING_SHEET_NAME = "EXPORT_ISAH"
Global Const ISAH_STAGING_COLUMNS = "Productieorder;Starttijd;Eindtijd;Resources;Aantal;Duur"
Global Const ISAH_STAGING_COLUMNS_DB_NAMES = "ProdHeaderOrdNr;StartDate;EndDate;MachGrpCode;Qty;Duur"
Global Const ISAH_STAGING_CAGGRP_COLUMN = "CAPGRP"
Global Const ISAH_STAGING_UPDATE_COLUMNS = "match_prod_header;StartDate_header;EndDate_header;ProdHeaderDossierCode;match_prod_boo;StartDate_boo;EndDate_boo;next_StartDate_header;next_Enddate_header;check_dates_header;next_StartDate_boo;next_Enddate_boo;check_dates_boo;next_StartTime_boo;next_StandCapacity_boo;next_MachPlanTime_boo;check_ProdBOOStatusCode"
Global Const ISAH_STAGING_UPDATE_COLUMNS_FORMATS = "match_prod_header=0;StartDate_header=yyyy-mm-dd hh:mm;EndDate_header=yyyy-mm-dd hh:mm;ProdHeaderDossierCode=0;match_prod_boo=0;StartDate_boo=yyyy-mm-dd hh:mm;EndDate_boo=yyyy-mm-dd hh:mm;next_StartDate_header=yyyy-mm-dd hh:mm;next_Enddate_header=yyyy-mm-dd hh:mm;check_dates_header=0;next_StartDate_boo=yyyy-mm-dd hh:mm;next_Enddate_boo=yyyy-mm-dd hh:mm;check_dates_boo=0;next_StartTime_boo=General;next_StandCapacity_boo=0;next_MachPlanTime_boo=0;check_ProdBOOStatusCode=General"

Global Const ISAH_STAGING_RANGE_NAME = "isah_staging_orders_range"
Global Const ISAH_STAGING_ORDERNR_INDEX = 1
Global Const ISAH_STAGING_ORDERNR_COLUMN = "ProdHeaderOrdNr"
Global Const ISAH_DATABASE_ORDERNR_COLUMN = "ProdHeaderOrdNr"
Global Const ISAH_DATABASE_CAGGRP_COLUMN = "MachGrpCode"
Global Const ISAH_DATABASE_DOSSIERCODE_COLUMN = "ProdHeaderDossierCode"
Global Const ISAH_DATABASE_DATE_COLUMNS = "convert(VARCHAR(20), StartDate) as StartDate, convert(VARCHAR(20), EndDate) as EndDate"

Global Const ISAH_CHECK_BOM_REQUIRED_DATE_SHEET = "CHECK_PROD_BILL_OF_MAT"
Global Const ISAH_MATCH_BOM_REQUIRED_DATE_SHEET = "JOIN_ISAH_EXPORT_PROD_BOM"
Global Const CHECK_BOM_REQUIRED_DATE_COLUMNS_FORMATS = "ProdHeaderDossierCode=0;min_bom_required_date=yyyy-mm-dd hh:mm;max_bom_required_date=yyyy-mm-dd hh:mm"

'ISAH database constants
Global Const ISAH_MANUAL_UPDATE_PRODBOOSTATUSCODE = "20" ' in ProdBOO set field ProdBOOStatusCode to this value

' Worksheet sorting
Global Const SHEETS_START_ORDER = "instructie;overzicht;Template;BULK"
Global Const LAST_CAPGRP_SHEET_NAME = "LN18"

' NEW ORDERS
Global Const ISAH_NEW_ORDERS_SHEET_NAME = "NIEUW"


' Public (global) variables
Public WorkBookConnection As Object

' variables
Dim capgrp As String, capgrp_sheet As String, range_name As String, workdaytimes_range As Range
Dim r0 As Long, r1 As Long, c0 As Long, c1 As Long
Dim rng0 As Range

' NEW DESIGN
' 1. Initialization and Setup
' 2. Data Retrieval and Processing
' 3. Data Manipulation and UpdatesE
' 4. Database and ISAH Operations
' 5. Printing and Layout
' 6. UI (BUTTON, DROPDOWN) HANDLERS

' 1. Initialization and Setup:
'    - init_capgrp_sheets
'    - init_capgrp_sheet
'    - init_capgrp_worksheet_code
'    - init_capgrp_sheets_ALL
'    - init_orders_range
'    - init_workdaytimes_range
'    - init_weeknumber_range
'    - init_print_file_range
'    - init_buttons
'    - init_worksheet_sorting
'    - init_control_sheet
'    - layout_control_sheet_buttons
'    - init_named_ranges
'    - handle_input_capgrp

' 2. Data Retrieval and Processing:
'    - get_isah_input_range
'    - get_isah_capgrp
'    - get_template_capgrp_names
'    - get_isah_capgrps
'    - get_capgrp_sheet_names
'    - get_art_capgrps
'    - get_last_art_capgrp
'    - get_bulk_capgrp
'    orders: get_orders_range, set_orders_range_values
'    - get_worktimes_range_name
'    - get_worktimes_range
'    - get_worktimes_values_range
'    - set_worktimes_range_values
'    - get_isah_export_range
'    - get_weeknumber_range
'    - get_capgrp_startdate
'    - get_capgrp_year
'    - get_capgrp_weeknumber
'    - get_capgrp_print_location
'    - get_row_in_named_range

' 3. Data Manipulation and Updates:
'    orders: copy_selected_orders, clear_orders_range, update_orders_range, fit_order_range_to_values
'    - update_bulk_capgrp_orders
'    - update_bulk_sorting
'    - restore_isah_default_sorting
'    - add_date_columns
'    - add_tank_lo_columns
'    - format_workdaytimes_range
'    - assign_macro_to_btn
'    - create_buttons
'    - get_workdaytimes_array
'    - set_capgrp_weeknumber
'    - set_capgrp_year
'    - set_default_weeknumber_year
'    - update_start_end_times
'    - find_week_overflow_row
'    - calculate_start_end_times
'    - find_block_starttime
'    - find_block_endtime
'    - get_last_block
'    - clear_all_capgrp_sheets
'    - insert_number_of_pallets_formula
'    - insert_record
'    - delete_record
'    - update_orders_color_format
'    - update_orders_columns_width
'    - fill_bulkcode_color
'    - fill_bulkcode_color_rbg

' 4. Database and ISAH Operations:
'    - getISAHconnection
'    - isah_export_test_connection
'    - checkIsahTestQuery
'    - getISAHdbname
'    - getISAHProfileName
'    - ISAHProfileIsHome
'    - getISAHprodheader
'    - getISAHprodboo
'    - getISAHprodbom
'    - getISAHpart
'    - getISAHconnstr
'    - isah_export_stage_orders
'    - isah_export_match_prodheader
'    - isah_export_update_prodboo_grp
'    - isah_export_match_prodboo
'    - isah_export_update_prodheader
'    - isah_export_update_prodboo
'    - isah_export_update_prodbom
'    - isah_export_check_bom_dates
'    - isah_export_match_bom_dates
'    - isah_export_run_all
'    - isah_import_articles

' 5. Printing and Layout:
'    - print_planning
'    - hide_columns_for_print
'    - delete_columns_for_print

' 6. Button Handlers:
'    - btn_update_isah_data_Click
'    - btn_add_record_Click
'    - btn_delete_record_Click
'    - btn_restore_prev_state_Click
'    - btn_print_dates_Click
'    - btn_calculate_dates_Click
'    - btn_import_art_Click
'    - btn_import_bulk_Click
'    - btn_add_capgrp_sheets_Click
'    - btn_isah_database_export_Click
'    - btn_clear_sheet_Click
'    - btn_isah_update_aantal_per_pallet_Click
'    - btn_import_testdata_Click

' 7. State Management:
'    - SafeStoreCurrentState

' 8. Checks and Exceptions:
'    - check_isah_input_empty
'    - check_isah_input_columns
'    - check_isah_input

' 9. Open Workbook Methods:
'    - init_articles_per_pallet

' 10. Clean up functions
'    - remove_capgrp_sheet

' 11. Utility and Helper Functions:
'    - create_yellow_gradient
'    - create_green_red_gradient
'    - get_color_palette
'    - get_random_color
'    - get_random_color_palette
'    - get_random_color_indices
'    - get_color_index_light
'    - get_color_indices_light
'    - returnStringWithinBrackets
'    - test_copy_selected_orders
'    - test_calculate_start_end_times
'    - test_insert_number_of_pallets_formula
'    - test_get_color_index_light
'    - test_

' INITIALIZERS: create capgrp sheets, UI panels, buttons
Sub init_capgrp_sheets(Optional capgrp_sheet_filter As String, Optional capgrp_sheet_names As collection = Nothing)

  ' get the unique article capgrps, bulk capgrps (starting with U) are handled in different sheet
  Dim capgrp_sheets As collection, rng0 As Range, rng1 As Range, rng2 As Range, ws0 As Worksheet, ws1 As Worksheet
  Dim capgrp As String, range_name As String, orders_rng As Range, prev_wsname As String, curEnableEvents As Boolean, curScreenUpdating As Boolean
  
  ' parameters
  Dim b_init_buttons As Boolean: b_init_buttons = True
  prev_wsname = "BULK" 'previous worksheet name, used to maintain sorting of worksheets
  
  ' store current EnableEvents, ScreenUpdating
  curEnableEvents = Application.EnableEvents
  curScreenUpdating = Application.ScreenUpdating
  Application.EnableEvents = False ' to prevent events from worksheet pasting
  Application.ScreenUpdating = False
  
  On Error GoTo handle_error
  
  ' capgrp_to_create: either the capgrps in the Template sheet or passed capgrp_sheet_names
  Dim capgrp_to_create As collection
  If capgrp_sheet_names Is Nothing Then
     Set capgrp_sheets = main.get_template_capgrp_names()
  Else
     Set capgrp_sheets = capgrp_sheet_names
  End If
  
  ' Get all capgrp sheets to initialize
  For Each c In capgrp_sheets
     ' 0 select capgrp orders from input ISAH and copy to capgrp_sheet
     capgrp = c
     If Not str.IsNull(capgrp_sheet_filter) And capgrp <> capgrp_sheet_filter Then
        GoTo nextiteration
     ElseIf capgrp = capgrp_sheet_filter Then
        Debug.Print "using filter, initialize capgrp sheet: " + capgrp
     Else
        Debug.Print "initialize capgrp sheet: " + capgrp
     End If
     
     ' 1. if sheet exists, delete and add, copy base sheet CODE and rename
     delete_worksheet capgrp, ThisWorkbook
     Set ws0 = w.get_or_create_worksheet(capgrp, ThisWorkbook)
     vb.CopyWorksheetCode "base", capgrp
     w.move_ws capgrp, after_ws:=prev_wsname, wb:=ThisWorkbook
     prev_wsname = capgrp
     
     ' 2 initialize layout: weeknumber_range, workdaytimes_range
     range_name = main.get_worktimes_range_name(capgrp)
     init_workdaytimes_range capgrp, range_name
     init_weeknumber_range capgrp
     init_print_file_range capgrp
     
     ' 3 copy capgrp orders, initialize as named range and sort on bulkcode
     copy_selected_orders capgrp, clear_ws:=False
     range_name = main.get_orders_range_name(capgrp)
     init_orders_range capgrp, range_name, overwrite:=True
     
     ' 4 update start and end times
     update_start_end_times capgrp

     ' 5 initialize control buttons
     main.init_buttons capgrp, True
     
     ' 5.1 position buttons
     main.position_buttons capgrp
     
     ' 6 initialize extra info range
     init_extra_info capgrp
       
     ' freeze panes
     w.freeze_top_rows ThisWorkbook.Sheets(capgrp), main.N_TOP_ROWS_FREEZE
      
     'Exit For
nextiteration:
  Next c
  
GoTo clean_up
handle_error:
    On Error GoTo 0
    Application.EnableEvents = curEnableEvents
    Application.ScreenUpdating = curScreenUpdating
    Err.Raise Err.Number 'rethrow error for user
    
clean_up:
    Application.EnableEvents = curEnableEvents
    Application.ScreenUpdating = curScreenUpdating
End Sub

Sub init_capgrp_worksheet_code()
  Set capgrp_sheet_names = main.get_capgrp_sheet_names()
  For Each capgrp_ In capgrp_sheet_names
     fs.copyProcedureCode "Worksheet_Activate", main.BASE_SHEET_NAME, CStr(capgrp_)
     fs.copyProcedureCode "Worksheet_Change", main.BASE_SHEET_NAME, CStr(capgrp_)
  Next
End Sub

Sub init_capgrp_sheet()
    capgrp_sheet = "LN 1"
    init_capgrp_sheets capgrp_sheet
End Sub

Sub init_capgrp_sheets_ALL()
    For Each c In main.get_template_capgrp_names()
        init_capgrp_sheets CStr(c)
    Next
    main.init_worksheet_sorting
    main.init_capgrp_worksheet_code
End Sub

Sub insert_volgnummer_into_orders(capgrp As String, Optional volgnummer_index As Integer = 1)
    Dim orders_rng As Range
    Set orders_rng = main.get_orders_range(capgrp)
    r.InsertColumnIntoRange orders_rng, volgnummer_index, main.ORDERS_RANGE_VOLGNUMMER_COLUMN
    
    'orders range is mutated so update in named range
    r.update_named_range main.get_orders_range_name(capgrp), orders_rng
    
    'main.get_orders_range(capgrp).Select
    'Debug.Print main.get_orders_range(capgrp).address
    main.calculate_volgnummer capgrp, volgnummer_index
    
End Sub

Sub calculate_volgnummer(capgrp As String, Optional volgnummer_index As Integer = 1)
    Dim orders_rng As Range, i As Long
    
    Set orders_rng = main.get_orders_range(capgrp)
    
    'TODO figure out why volgnummer is not in range
    'If Not r.column_exist(orders_rng, "volgnummer") Then
    '   Exit Sub
    'End If
    
    If orders_rng.count > 1 Then
        For i = 2 To orders_rng.Rows.count
            orders_rng.Cells(i, volgnummer_index).value = i - 1
        Next i
    End If
    
    orders_rng.columns(volgnummer_index).EntireColumn.AutoFit
End Sub

Sub test_volgnummer()
    Application.EnableEvents = False
    insert_volgnummer_into_orders "LN 1"
    main.get_orders_range("LN 1").Select
    Application.EnableEvents = True
End Sub

Sub restore_ln1()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    main.init_capgrp_sheet
    main.update_orders_range "LN 1"
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub test_init_orders_range()
    Application.EnableEvents = False
    Dim range_name As String, capgrp As String
    capgrp = "LN 6"
    range_name = capgrp & "_orders"
    main.init_orders_range capgrp, range_name, True
    Application.EnableEvents = True
End Sub

Sub init_orders_range(capgrp As String, range_name As String, overwrite As Boolean)
    ' Initializes the orders range for a given capacity group (capgrp).
    ' This subroutine creates or retrieves a named range for orders, formats the range,
    ' and inserts necessary columns such as start and end dates.
    '
    ' Parameters:
    ' capgrp - The name of the capacity group for which the orders range is being initialized.
    ' range_name - The name of the range to be created or retrieved.
    ' overwrite - A boolean indicating whether to overwrite the existing range if it exists.
    '
    ' The subroutine performs the following actions:
    ' - Creates or retrieves a named range for the orders.
    ' - Formats the columns in the range according to predefined formats.
    ' - Inserts a formula for calculating the number of pallets.
    ' - Adds additional columns such as "Tank" and "L/O".
    ' - Sorts the range by bulk code and applies color formatting.
    ' - Adjusts column widths for better readability.
    Dim rng0 As Range, ws0 As Worksheet, rng1 As Range, c1 As Long

    ' create or get named range
    If overwrite Or Not r.name_exist(range_name) Then
      ' get the length of the input data header (r1), expand downwards to last row
      input_data_columns = str.str_to_array(main.INPUT_DATA_HEADER)
      c1 = a.numArrayColumns(input_data_columns)
      Set rng0 = r.expand_range(main.ORDERS_RANGE_ADDR, ws:=ThisWorkbook.Worksheets(capgrp), c1:=c1, dbg:=main.P_DEBUG)
      r.delete_named_range range_name, clear:=True
      r.create_named_range range_name, capgrp, rng0.address, overwrite:=overwrite, expand_range:=False
      'add startdate, enddate columns without formatting
      Set rng0 = add_date_columns(range_name)
    Else
      ' get the named range with orders
      Set rng0 = r.get_range(range_name)
    End If
    
    ' insert `volgnummer` at the first column of the orders_range
    insert_volgnummer_into_orders capgrp
    
    ' formatting
    Dim i As Long
    Set ws0 = rng0.Worksheet
    ws0.Activate
    formats_array = str.str_to_array(main.OUTPUT_DATA_FORMATS)
    For i = LBound(formats_array) To UBound(formats_array)
        rng0.columns(i + 1).Select
        With Selection
        .NumberFormat = formats_array(i)
        .HorizontalAlignment = xlCenter
        End With
    Next i
    
    ' 20230720 insert the #pallets formula
    main.insert_number_of_pallets_formula capgrp
    
    ' 20240209 add Tank, L/O
    main.add_tank_lo_columns capgrp
        
    ' sort on bulkcode and color formatting bulkcode column
    r.sort_range_by_columns rng0, main.BULKCODE_COLUMN
    
    ' add color, column/row height/width formatting
    main.update_orders_color_format capgrp
    main.update_orders_columns_width capgrp
    
End Sub

Sub init_workdaytimes_range(capgrp As String, range_name As String)
    Dim rng0 As Range, workdays As Range, worktimes As Range, workdaytimes_range As Range
    range_name = capgrp & "_workdays"
    r.create_named_range range_name, capgrp, main.WORKDAYS_RANGE_ADDRESS, header_row:=main.WORKDAYS_RANGE_HEADER, _
    id_row:=main.WORKDAYS_RANGE_IDS, overwrite:=True
    
    ' fill workdaytimes range with default value 1
    Set workdaytimes_range = r.get_range(range_name)
    Set workdays = r.get_column(range_name, 1, offset_row:=1)
    Set worktimes = r.get_row(range_name, 1, offset_column:=1)
    For Each wd In workdays.Cells
        For Each wt In worktimes.Cells
        r.set_value range_name, wd, wt, 1
        Next
    Next
    
    ' autofit columns
    r.autofit_columns r.get_range(range_name)
    
    ' format columns
    format_workdaytimes_range range_name
End Sub

' restore days column, times row in worktimes range
Sub init_workdaytimes_days_times(capgrp As String)
    Dim rng As Range
    Set rng = get_worktimes_range(capgrp)
    rng.Rows(1).value = r.str_to_array(main.WORKDAYS_RANGE_HEADER)
    rng.columns(1).value = WorksheetFunction.Transpose(r.str_to_array(main.WORKDAYS_RANGE_IDS))
End Sub

Sub init_weeknumber_range(capgrp As String)
    ' initialize weeknumber input
    range_name = capgrp & "_input_weeknumber"
    r.create_named_range range_name, capgrp, main.WEEKNUMBER_RANGE_ADDRESS, header_row:=",", id_row:=main.WEEKNUMBER_RANGE_IDS, overwrite:=True ' replace later
    main.set_capgrp_weeknumber capgrp, 0
    r.get_range(range_name).Cells(2, 2).Interior.Color = main.INPUT_FIELD_COLOR
    
    ' initialize year input
    range_name = capgrp & "_input_year"
    r.create_named_range range_name, capgrp, main.YEAR_RANGE_ADDRESS, header_row:=",", id_row:=main.YEAR_RANGE_IDS, overwrite:=True, default_value:=2023 ' replace later
    main.set_capgrp_year capgrp, 0
    r.get_range(range_name).Cells(2, 2).Interior.Color = main.INPUT_FIELD_COLOR
    
    ' refit
    r.autofit_columns r.get_range(range_name)
End Sub

Function get_weeknumber_range(capgrp As String) As Range
    Set get_weeknumber_range = r.get_range(capgrp & "_input_weeknumber")
End Function

Function get_weeknumber(capgrp As String) As Integer
    get_weeknumber = CInt(main.get_weeknumber_range(capgrp).Cells(2, 2).value)
End Function

Function get_yearnumber_range(capgrp As String) As Range
    Set get_yearnumber_range = r.get_range(capgrp & "_input_year")
End Function

Function get_yearnumber_range_name(capgrp As String) As String
    get_yearnumber_range_name = capgrp & "_input_year"
End Function

Sub init_print_file_range(capgrp As String)
    range_name = capgrp & "_input_print_location"
    r.create_named_range range_name, capgrp, main.PRINT_FILE_RANGE_ADDRESS, header_row:=",", id_row:=main.PRINT_FILE_RANGE_IDS, overwrite:=True
    main.set_capgrp_print_location capgrp
    r.get_range(range_name).Cells(2, 2).Interior.Color = main.INPUT_FIELD_COLOR
End Sub

Sub init_extra_info(capgrp As String)
    ' Initializes the extra info range for a given capacity group (capgrp).
    ' This subroutine creates a named range for extra info and sets its location.
    '
    ' Parameters:
    ' capgrp - The name of the capacity group for which the extra info range is being initialized.
    '
    ' The subroutine performs the following actions:
    ' - Creates a named range for the extra info.
    ' - Sets the location of the extra info range.
    Dim range_name As String
    Dim extraInfoRange As Range, ws0 As Worksheet
    
    range_name = capgrp & "_extra_info"
    r.create_named_range range_name, capgrp, main.PRINT_EXTRA_INFO_ADDRESS, overwrite:=True
    
    Set ws0 = ThisWorkbook.Sheets(capgrp)
    Set extraInfoRange = ws0.Range(main.PRINT_EXTRA_INFO_ADDRESS)
    
    ' Simple extra info range
    extraInfoRange.Merge
    
    ' Add borders to ExtraInfo + WrapText
    With extraInfoRange
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    
End Sub

Sub init_buttons(capgrp As String, Optional overwrite As Boolean = True)
    Dim wb As Workbook, capgrp_btn As String
    Set wb = ThisWorkbook
    Dim left_offset As Long
    left_offset = main.BTN_LEFT_OFFSET
    
    capgrp_btn = Replace(capgrp, " ", "_")
    
    ' first column of buttons: Overhalen orders, Regel toevoegen, Regel verwijderen, Ga Terug
    ctr.createCmdButton "btn_update_isah_data_" & capgrp_btn, "Overhalen orders", ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, length:=main.BTN_HEIGHT, width:=main.BTN_WIDTH
    
    ctr.createCmdButton "btn_add_record_" & capgrp_btn, "Regel toevoegen", ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, length:=main.BTN_HEIGHT, width:=main.BTN_WIDTH
    
    ctr.createCmdButton "btn_delete_record_" & capgrp_btn, "Regel verwijderen", ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, length:=main.BTN_HEIGHT, width:=main.BTN_WIDTH
    
    ctr.createCmdButton "btn_restore_prev_state_" & capgrp_btn, main.BTN_RESTORE_PREV_STATE_LABEL, ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, length:=main.BTN_HEIGHT, width:=main.BTN_WIDTH
    
    ' second column of buttons: Exporteren PDF, Printen planning,
    ctr.createCmdButton "btn_export_pdf_" & capgrp_btn, "Exporteren PDF", ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, length:=main.BTN_HEIGHT, width:=main.BTN_WIDTH
    
    ctr.createCmdButton "btn_print_dates_" & capgrp_btn, "Printen planning", ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, length:=main.BTN_HEIGHT, width:=main.BTN_WIDTH
    
    ' third column of buttons: Actualiseer tijden
    ctr.createCmdButton "btn_calculate_dates_" & capgrp_btn, "Actualiseer tijden", ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, assign_macro:="btn_calculate_dates_Click", length:=main.BTN_HEIGHT, width:=main.BTN_WIDTH
    
    ctr.assignMacroToCmdButton "btn_update_isah_data_" & capgrp_btn, assign_macro:="btn_update_isah_data_Click"
    ctr.assignMacroToCmdButton "btn_add_record_" & capgrp_btn, assign_macro:="btn_add_record_Click"
    ctr.assignMacroToCmdButton "btn_delete_record_" & capgrp_btn, assign_macro:="btn_delete_record_Click"
    ctr.assignMacroToCmdButton "btn_restore_prev_state_" & capgrp_btn, assign_macro:="btn_restore_prev_state_Click"
    ctr.assignMacroToCmdButton "btn_export_pdf_" & capgrp_btn, assign_macro:="btn_export_pdf_Click"
    ctr.assignMacroToCmdButton "btn_print_dates_" & capgrp_btn, assign_macro:="btn_print_dates_Click"
    ctr.assignMacroToCmdButton "btn_calculate_dates_" & capgrp_btn, assign_macro:="btn_calculate_dates_Click"
    
End Sub

Sub position_buttons(capgrp As String)
    'Declare variables
    Dim wb As Workbook, capgrp_btn As String
    Dim left_offset As Long
    
    'Initialize variables
    Set wb = ThisWorkbook
    left_offset = main.BTN_LEFT_OFFSET
    capgrp_btn = Replace(capgrp, " ", "_")
    
    ' Position the "Overhalen orders" button
    ctr.positionCmdButton "btn_update_isah_data_" & capgrp_btn, address:=main.BTN_UPDATE_ISAH_ADDRESS, ws:=wb.Sheets(capgrp), left_offset:=left_offset
    
    ' Position the "Regel toevoegen" button
    ctr.positionCmdButton "btn_add_record_" & capgrp_btn, address:=main.BTN_ADD_RECORD_ADDR, ws:=wb.Sheets(capgrp), left_offset:=left_offset
    
    ' Position the "Regel verwijderen" button
    ctr.positionCmdButton "btn_delete_record_" & capgrp_btn, address:=main.BTN_DELETE_RECORD_ADDR, ws:=wb.Sheets(capgrp), left_offset:=left_offset
    
    ' Position the "Ga Terug" button
    ctr.positionCmdButton "btn_restore_prev_state_" & capgrp_btn, address:=main.BTN_RESTORE_PREV_STATE_ADDR, ws:=wb.Sheets(capgrp), left_offset:=left_offset
    
    ' Position the "Exporteren PDF" button
    ctr.positionCmdButton "btn_export_pdf_" & capgrp_btn, address:=main.BTN_EXPORT_PDF_ADDR, ws:=wb.Sheets(capgrp), left_offset:=main.BTN_SECOND_COLUMN_LEFT_OFFSET
    
    ' Position the "Printen planning" button
    ctr.positionCmdButton "btn_print_dates_" & capgrp_btn, address:=main.BTN_PRINT_DATES_ADDR, ws:=wb.Sheets(capgrp), left_offset:=main.BTN_SECOND_COLUMN_LEFT_OFFSET
    
    ' Position the "Actualiseer tijden" button
    ctr.positionCmdButton "btn_calculate_dates_" & capgrp_btn, address:=main.BTN_CALCULATE_DATES_ADDR, ws:=wb.Sheets(capgrp), left_offset:=main.BTN_SECOND_COLUMN_LEFT_OFFSET
End Sub

Sub test_init_position_buttons()
    init_buttons "LN 1"
    position_buttons "LN 1"
End Sub

Sub init_worksheet_sorting()
    Dim sheetOrder As New collection, newSheets As collection, lnSheetNames As collection
    ' start with SHEETS_START_ORDER
    Set newSheets = clls.toCollection(main.SHEETS_START_ORDER, ";")
    
    ' add the Template capgrp sheets
    Set lnSheetNames = main.get_capgrp_sheet_names()
    Set sheetOrder = clls.concatCollections(newSheets, lnSheetNames)
    
    ' finally: NIEUW, ...
    Set newSheets = clls.toCollection("NIEUW;EXPORT_ISAH;" & main.NUMBER_PER_PALLET_SHEET_NAME, ";")
    Set sheetOrder = clls.concatCollections(sheetOrder, newSheets)
    
    ' order the sheets
    Application.EnableEvents = False
    On Error Resume Next
    w.orderSheets sheetOrder
    On Error GoTo 0
    Application.EnableEvents = True
End Sub

' control sheet
Sub init_control_sheet()
     Dim ws As Worksheet

     ' if sheet CONTROL_SHEET_NAME exists, delete. copy base sheet and rename
     w.delete_worksheet main.CONTROL_SHEET_NAME
     w.copy_ws main.BASE_SHEET_NAME, main.CONTROL_SHEET_NAME
     Set ws = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)

     ' initialize buttons
     ctr.add_button "btn_import_art", main.BTN_IMPORT_ART_ADDR, ws:=ws, _
     overwrite:=True, label:="Importeren alle artikelen", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_import_art", "btn_import_art_Click"
     
     ctr.add_button "btn_import_bulk", main.BTN_IMPORT_BULK_ADDR, ws:=ws, _
     overwrite:=True, label:="Importeren bulken", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_import_bulk", "btn_import_bulk_Click"

     ctr.add_button "btn_add_capgrp_sheets", main.BTN_ADD_CAPGRP_ADDR, ws:=ws, _
     overwrite:=True, label:="Toevoegen nieuwe capaciteitsgroepen", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_add_capgrp_sheets", "btn_add_capgrp_sheets_Click"
     
     ctr.add_button "btn_isah_database_export", main.BTN_ISAH_EXPORT_ADDR, ws:=ws, overwrite:=True, label:="Exporteren naar ISAH database", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_isah_database_export", "btn_isah_database_export_Click"
    
     'ctr.add_button "btn_clear_sheet", main.BTN_CLEAR_SHEET, ws:=ws, overwrite:=True, label:="Alle productielijnen wissen", w:=main.BTN_WIDTH
     'main.assign_macro_to_btn "btn_clear_sheet", "btn_clear_sheet_Click"
     
     ctr.add_button "btn_add_new_orders", main.BTN_ADD_NEW_ORDERS_ADDR, ws:=ws, overwrite:=True, label:="Orders van tabblad NIEUW toevoegen", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_add_new_orders", "btn_add_new_orders_Click"
End Sub

Sub layout_control_sheet_buttons()
     Dim ws As Worksheet
     Set ws = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)
     
     ' initialize buttons
     ctr.add_button "btn_import_art", main.BTN_IMPORT_ART_ADDR, ws:=ws, _
     overwrite:=False, label:="Importeren alle artikelen", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_import_art", "btn_import_art_Click"
     
     ctr.add_button "btn_import_bulk", main.BTN_IMPORT_BULK_ADDR, ws:=ws, _
     overwrite:=False, label:="Importeren bulken", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_import_bulk", "btn_import_bulk_Click"

     ctr.add_button "btn_add_capgrp_sheets", main.BTN_ADD_CAPGRP_ADDR, ws:=ws, _
     overwrite:=False, label:="Toevoegen nieuwe capaciteitsgroepen", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_add_capgrp_sheets", "btn_add_capgrp_sheets_Click"
     
     ctr.add_button "btn_isah_database_export", main.BTN_ISAH_EXPORT_ADDR, ws:=ws, overwrite:=False, label:="Exporteren naar ISAH database", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_isah_database_export", "btn_isah_database_export_Click"
    
     'ctr.add_button "btn_clear_sheet", main.BTN_CLEAR_SHEET, ws:=ws, overwrite:=false, label:="Alle productielijnen wissen", w:=main.BTN_WIDTH
     'main.assign_macro_to_btn "btn_clear_sheet", "btn_clear_sheet_Click"

     ctr.add_button "btn_add_new_orders", main.BTN_ADD_NEW_ORDERS_ADDR, ws:=ws, overwrite:=False, label:="Orders van tabblad NIEUW toevoegen", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_add_new_orders", "btn_add_new_orders_Click"
     
End Sub

Sub layout_position_control_buttons()
    ' Documentation: This subroutine arranges buttons on an Excel worksheet based on specified parameters.
    ' It sets the size and position of each button in the btn_names array.
    ' The first 4 buttons are aligned relative to cell A1, and the next 4 buttons are aligned relative to cell D1.

    Dim btn_names As Variant
    Dim w As Double, h As Double
    Dim right_distance As Double, top_distance As Double, top_interval As Double
    Dim i As Integer
    Dim btn As Object
    Dim ws As Worksheet
    
    ' Initialize the worksheet
    Set ws = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)
    
    ' Initialize button names
    btn_names = Array("btn_import_art", "btn_import_bulk", "Button 14", "Button 15", "btn_isah_database_export", "btn_add_capgrp_sheets", "", "Button 13")
    
    ' Set button size
    w = 88 ' Width of the button
    h = 52  ' Height of the button
    
    ' Set alignment parameters
    right_distance = 32 ' Distance from the right of the reference cell
    top_distance = 18    ' Initial top distance from the reference cell
    top_interval = 75   ' Interval increase for each subsequent button
    
    ' Loop through each button name in the array
    For i = LBound(btn_names) To UBound(btn_names)
        If (btn_names(i) = "") Then
        GoTo next_i
        End If
        
        ' Create or reference the button
        Set btn = ws.Buttons(btn_names(i))
        
        ' Set button size
        btn.width = w
        btn.Height = h
        
        ' Determine button position
        If i < 4 Then
            ' First 4 buttons aligned to cell A1
            btn.left = ws.Range("A1").left + right_distance
            btn.top = ws.Range("A1").top + top_distance + (i * top_interval)
        Else
            ' Next 4 buttons aligned to cell D1
            btn.left = ws.Range("D1").left + right_distance
            btn.top = ws.Range("D1").top + top_distance + ((i - 4) * top_interval)
        End If
next_i:
    Next i
End Sub

Sub test_control_button()
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)
Set btn = ws.Buttons("btn_import_art")
Debug.Print btn.width, btn.Height
Debug.Print Range("A1").left, Range("D1").left, btn.left
Debug.Print Range("A2").top, Range("A7").top
End Sub

' 2.Data Retrieval and Processing
' Public variables
Function getWorkbookConnection() As ADODB.Connection
   If WorkBookConnection Is Nothing Then
      Set WorkBookConnection = db.openExcelConn(ThisWorkbook)
      Set getWorkbookConnection = WorkBookConnection
      Exit Function
   Else
      Set getWorkbookConnection = WorkBookConnection
   End If
End Function

' GETTERS
Function get_isah_input_range() As Range
    Dim rng0 As Range, ws0 As Worksheet
    Set ws0 = ThisWorkbook.Sheets(main.INPUT_ISAH_SHEET)
    Set rng0 = r.expand_range(ws0.Cells(1), ws0, dbg:=main.P_DEBUG)
    Set get_isah_input_range = rng0
End Function

Sub format_isah_input_range()
    Dim ws0 As Worksheet, rng0 As Range
    Set ws0 = ThisWorkbook.Sheets(main.INPUT_ISAH_SHEET)
    Set rng0 = r.expand_range("A1", ws0)
    r.formatRangeColumns rng0, ISAH_TEMPLATE_COLUMN_FORMATS, ThisWorkbook
End Sub

Function get_isah_capgrp() As Variant
    Dim rng0 As Range, rng1 As Range
    Set rng0 = r.get_range(ThisWorkbook.Sheets(main.INPUT_ISAH_SHEET))
    Set rng1 = r.get_column_values(rng0, main.CAPGRP_COLUMN_NAME)
    get_isah_capgrp = r.get_unique_vals(rng1)
End Function

'get the capgrp names (LN 1, LN 2, ...) from the Template sheet, to use for initializing capgrp sheets
Function get_template_capgrp_names() As collection
    Dim col0 As New collection, map_capgrp_inpak_col As collection, capgrp_sheet As String
    Set art_capgrp_col = a.as_collection(main.get_art_capgrps())
    Set map_capgrp_inpak_col = a.as_collection(Split(CStr(main.INPAK_CAPGRP_LIST), ";"))
    
    ' merge the INPAK capgrps to INPAK
    For Each c In art_capgrp_col
        capgrp = c
        'capgrps to map to INPAK, add INPAK if not added yet
        If clls.item_exists(capgrp, map_capgrp_inpak_col) Then
            If Not clls.item_exists("INPAK", col0) Then
               capgrp_sheet = "INPAK"
            Else
               GoTo next_c
            End If
        Else
            capgrp_sheet = capgrp
        End If
        col0.Add capgrp_sheet
next_c:
    Next c
    
    ' move INPAK to back
    If clls.item_exists("INPAK", col0) Then
       Dim col1 As New collection
       For Each it In col0
          If it <> "INPAK" Then
            col1.Add it
          End If
       Next it
       col1.Add "INPAK"
       Set get_template_capgrp_names = col1
    Else
       Set get_template_capgrp_names = col0
    End If
End Function

' map capgrp_sheet to isah_capgrp(s), returns array
Function get_isah_capgrps(capgrp_sheet) As Variant
    Dim inpak_capgrp_array As Variant
    inpak_capgrp_array = Split(CStr(main.INPAK_CAPGRP_LIST), ";")
    If capgrp_sheet = "INPAK" Then
       isah_capgrps = inpak_capgrp_array
    Else
       isah_capgrps = Array(capgrp_sheet)
    End If
    get_isah_capgrps = isah_capgrps
End Function

' get capgrp names from sheets
Function get_capgrp_sheet_names() As collection
    Dim art_capgrps As Variant, col0 As New collection
    art_capgrps = main.get_art_capgrps()
    For Each ws0 In ThisWorkbook.Sheets
        If a.ItemInArray(ws0.name, art_capgrps) Or str.regexp_match(ws0.name, main.CAPGRP_SHEET_PATTERN) Then
           col0.Add ws0.name
        Else
           GoTo nx_ws
        End If
nx_ws:
    Next
    Set get_capgrp_sheet_names = clls.sort_collection(col0, True)
End Function

Sub test_get_art_capgrps()

' Get filter patterns from the GAPGRP_FILTER_SHEET
Dim rng0 As Range
Set rng0 = r.expand_range("A1", main.GAPGRP_FILTER_SHEET)
filterPatterns = r.get_column_values(rng0, "Capgrp")
'a.printArray filterPatterns

a.printArray main.get_art_capgrps()

a.printArray a.FilterVectorWithPattern(main.get_isah_capgrp(), "^U", True)

a.printArray main.get_bulk_capgrp()
a.printArray a.FilterVectorWithPattern(main.get_isah_capgrp(), "^U", False)

End Sub

Function get_art_capgrps() As Variant
    ' return ISAH capgrp NOT starting with U
    Dim isah_capgrp_array As Variant
    isah_capgrp_array = main.get_isah_capgrp()
    get_art_capgrps = a.FilterVectorWithPattern(isah_capgrp_array, "^U", True)
End Function

Function get_last_art_capgrp() As String
    Dim col0 As collection
    Set col0 = a.as_collection(main.get_art_capgrps())
    If col0.count > 0 Then
    get_last_art_capgrp = col0.item(col0.count)
    Else
    get_last_art_capgrp = ThisWorkbook.Sheets(1).name
    End If
End Function

Function get_bulk_capgrp() As Variant
    ' return ISAH capgrp starting with U
    Dim isah_capgrp_array As Variant
    isah_capgrp_array = main.get_isah_capgrp()
    get_bulk_capgrp = a.FilterVectorWithPattern(isah_capgrp_array, "^U", False)
End Function

Function get_orders_range(capgrp As String) As Range
    ' Retrieves the orders range for a specified capacity group (capgrp).
    ' This function returns the range object associated with the named range for orders.
    '
    ' Parameters:
    ' capgrp - The name of the capacity group for which the orders range is being retrieved.
    '
    ' Returns:
    ' A Range object representing the orders range for the specified capacity group.
    Dim ord_range As Range, range_name As String
    range_name = main.get_orders_range_name(capgrp)
    Set get_orders_range = r.get_range(range_name)
End Function

Sub set_orders_range_values(capgrp As String, values)
    ' Sets the values for the orders range of a specified capacity group (capgrp).
    ' This subroutine updates the named range for orders with the provided values.
    '
    ' Parameters:
    ' capgrp - The name of the capacity group for which the orders range is being updated.
    ' values - The values to be set in the orders range.
    '
    ' Note: The subroutine assumes that the named range for orders has been previously created
    ' and is available in the workbook.
    Dim range_name As String
    range_name = main.get_orders_range_name(capgrp)
    r.updateNamedRangeWithValues range_name, values
End Sub

Function get_worktimes_range_name(capgrp) As String
    get_worktimes_range_name = capgrp & "_workdays"
End Function

Function get_worktimes_range(capgrp) As Range
    Dim range_name As String
    range_name = capgrp & "_workdays"
    Set get_worktimes_range = r.get_range(range_name)
End Function

Function get_worktimes_values_range(capgrp) As Range
    Dim range_name As String, rng0 As Range
    range_name = main.get_worktimes_range_name(capgrp)
    Set rng0 = r.get_range(range_name)
    If rng0.Rows.count > 1 Then
       Set rng0 = r.subset_range(rng0, startrow:=2, startcol:=2)
    Else
       Set rng0 = Nothing
    End If
    Set get_worktimes_values_range = rng0
End Function

Sub set_worktimes_range_values(capgrp, values)
    Dim worktimesRng As Range
    Set worktimesRng = main.get_worktimes_values_range(capgrp)
    If Not IsArray(values) Then
       values = a.create_array(worktimesRng.Rows.count, worktimesRng.columns.count, values)
    End If
    r.paste_array values, worktimesRng.address, worktimesRng.Worksheet
End Sub

Function get_isah_export_range() As Range
    Dim rng0 As Range, ws0 As Worksheet
    Set ws0 = ThisWorkbook.Sheets(main.ISAH_STAGING_SHEET_NAME)
    Set rng0 = r.expand_range(ws0.Cells(1), ws0, dbg:=main.P_DEBUG)
    Set get_isah_export_range = rng0
End Function

' LOGICAL
Function IsCapgrpSheet(capgrp_sheet As String) As Boolean
    IsCapgrpSheet = clls.item_exists(capgrp_sheet, main.get_capgrp_sheet_names())
End Function

' LAYOUT/CONFIG

' ISAH CAPGRP PLANNING FLOW: per capgrp (sheet), import orders from `INPUT_ISAH_SHEET` and calculate startendtimes
Sub copy_selected_orders(capgrp As String, Optional clear_ws As Boolean = False, Optional remove_filter As Boolean = True)
    ' Copies selected orders for a specified capacity group (capgrp) from the ISAH input sheet.
    ' This subroutine filters and copies orders based on the capacity group and production week,
    ' and pastes them into the corresponding worksheet for the capacity group.
    '
    ' Parameters:
    ' capgrp - The name of the capacity group for which orders are being copied.
    ' clear_ws - An optional boolean indicating whether to clear the target worksheet before copying.
    ' remove_filter - An optional boolean indicating whether to remove the filter after copying.
    '
    ' The subroutine performs the following actions:
    ' - Filters orders from the ISAH input sheet based on the capacity group and production week.
    ' - Copies the filtered orders to the corresponding worksheet for the capacity group.
    ' - Restores the default sorting of the ISAH input sheet.
   Dim ws0 As Worksheet, rng0 As Range, ws1 As Worksheet, prodWk As String
   Set ws0 = ThisWorkbook.Sheets(main.INPUT_ISAH_SHEET)
   Set rng0 = r.get_range(ws0)
   
   ' 20230602 added  prodWk input to orders filter
   prodWk = CStr(main.get_capgrp_weeknumber(capgrp))
   
   ' map the capgrp (sheet) to the isah capgrps (array). the capgrp sheet INPAK is mapped the ISAH capgrps HCM, NIVP, etc
   isah_capgrp_array = main.get_isah_capgrps(capgrp)
   If main.P_DEBUG Then
      Debug.Print "filtering ISAH orders on capgrps:"
      a.printArray isah_capgrp_array
   End If
   filtered_array = r.filter_range(rng0, main.CAPGRP_COLUMN_NAME, isah_capgrp_array, main.PRODWK_COLUMN, prodWk, remove_filter:=remove_filter, xl_operator:=xlFilterValues)
   input_data_columns = str.str_to_array(main.INPUT_DATA_HEADER)
   selected_array = a.select_array_columns(filtered_array, input_data_columns)
   
   ' restore default sorting of isah input sheet
   main.restore_isah_default_sorting
   
   ' 2 paste selected capgrp orders on capgrp sheet
   Set ws1 = w.get_or_create_worksheet(capgrp, ThisWorkbook, overwrite:=False, clear:=clear_ws)
   r.paste_array selected_array, main.ORDERS_RANGE_ADDR, ws1
   
   ws1.Activate
   
End Sub

Sub test_copy_selected_orders()
    main.copy_selected_orders "LN 5", False, False
End Sub

' select bulk orders from Template sheet and update BULK sheet
Sub update_bulk_capgrp_orders()
   Dim ws0 As Worksheet, rng0 As Range, ws1 As Worksheet, prodWk As String, capgrp_sheets As collection, formula_code As String
   Dim start_address As String, ordernr_index As Long, bulkcode_index As Long, qty_index As Long, dsc_index As Long
   Dim BulkSheet As Worksheet, IsahSheet As Worksheet
   
   ' settings
   Dim wkDayNameLong As Boolean: wkDayNameLong = False
   Dim capgrpColumnWidth As Integer
   If wkDayNameLong Then
      capgrpColumnWidth = 16
   Else
      capgrpColumnWidth = 12
   End If

   Application.EnableEvents = False

   ' 1 prepare the BULK sheet
   Dim range_values As Range, range_name As String, CAPGRP_COLUMN_NAME As String, num_rows As Long
   Set BulkSheet = ThisWorkbook.Sheets(main.BULK_SHEET_NAME)
   w.clearWorksheet BulkSheet.name, ThisWorkbook
   BulkSheet.Activate
   range_name = main.BULK_ORDERS_RANGE_NAME
   If r.name_exist(range_name) Then
       r.delete_named_range range_name, clear:=True
   End If
   
   ' Add named range with header
   r.create_named_range main.BULK_ORDERS_RANGE_NAME, main.BULK_SHEET_NAME, main.BULK_SHEET_START_ADDR, expand_range:=False, clear:=False, header_row:=Replace(main.BULK_ORDER_SHEET_HEADER, ";", ",")
   
   ' 2.1 get BULK orders from Template
   Set IsahSheet = ThisWorkbook.Sheets(main.INPUT_ISAH_SHEET)
   Set rng0 = r.get_range(IsahSheet)
   ' column indices
   ordernr_index = r.get_column_index(rng0, "Productieorder")
   bulkcode_index = r.get_column_index(rng0, "Artikel")
   qty_index = r.get_column_index(rng0, "Qty1")
   dsc_index = r.get_column_index(rng0, "Omschrijving")
   
   ' 2.2 filter orders from the bulk capgrp only
   bulk_capgrp_array = main.get_bulk_capgrp()
   If Not IsArray(bulk_capgrp_array) Or a.ArrayIsEmpty(bulk_capgrp_array) Then
      ' no BULK capgrps found in ISAH Template, skip procedure
      Debug.Print "no BULK capgrps found in ISAH Template"
      MsgBox main.ERR_MSG_NO_BULK_CODES
      GoTo finally
   End If
   
   ' 3 get arrays insert later into BULK sheet from ISAH Template sheet => TODO: get_column(array,index,offset_row)
   filtered_array = r.filter_range(rng0, main.CAPGRP_COLUMN_NAME, bulk_capgrp_array, remove_filter:=True, xl_operator:=xlFilterValues)
   input_data_columns = str.str_to_array(main.BULK_ORDERS_HEADER)
   selected_array = filtered_array
   selected_values_array = a.subset_rows(selected_array, LBound(selected_array) + 1)
   ordernr_arr = a.subset_columns(selected_values_array, ordernr_index, ordernr_index)
   bulkcode_arr = a.subset_columns(selected_values_array, bulkcode_index, bulkcode_index)
   qty_arr = a.subset_columns(selected_values_array, qty_index, qty_index)
   dsc_arr = a.subset_columns(selected_values_array, dsc_index, dsc_index)
   
   ' restore default sorting of isah input sheet
   main.restore_isah_default_sorting
   If a.numArrayRows(selected_array) <= 1 Then
      GoTo finally
   End If

   ' resize to fit ordernr_arr,bulkcode,qty
   num_rows = a.numArrayRows(ordernr_arr)
   r.resize_named_range range_name, add_rows:=num_rows
   
   ' set ordernr, bulkcode, qty values
   r.set_column_values range_name, main.BULK_ORDER_SHEET_ORDERNR, values:=ordernr_arr
   r.set_column_values range_name, main.BULK_ORDER_SHEET_BULKCODE, values:=bulkcode_arr
   r.set_column_values range_name, main.BULK_ORDER_SHEET_QTY, values:=qty_arr
   r.set_column_values range_name, main.BULK_ORDER_SHEET_DSC, values:=dsc_arr
   
   ' get the initial column count c0
   Dim c0 As Long
   Set rng0 = r.get_range(range_name)
   c0 = rng0.columns.count
   
   ' add Lijn (capgrp sheet) columns to bulk_orders_range
   Dim formulaTemplate As String, formula_range As Range, capgrp_sheet As String

   ' formulaTemplate: VLOOKUP for getting the minimal startdate from the capgrp_sheet
   ' 20240211: wrap formula in formatting function `formatDateVDMI`
   If wkDayNameLong Then
      formulaTemplate = "=formatDateVDMILong(IF(ISERROR(VLOOKUP(@1,'@2'!@3,@4,FALSE)),"" "",VLOOKUP(@1,'@2'!@3,@4,FALSE)))"
   Else
      formulaTemplate = "=formatDateVDMIShort(IF(ISERROR(VLOOKUP(@1,'@2'!@3,@4,FALSE)),"" "",VLOOKUP(@1,'@2'!@3,@4,FALSE)))"
   End If
   start_address = "$" & u.remove_dollar_sign(r.get_column(range_name, main.BULK_ORDER_SHEET_BULKCODE).Cells(2).address)

   ' for each capgrp_sheet, create a column with VLOOKUP on BULK sheet
   Dim lastFormulaRangeAddr As String
   Set capgrp_sheets = main.get_capgrp_sheet_names()
   For Each c In capgrp_sheets
       capgrp_sheet = c
       If Not w.sheet_exists(c, ThisWorkbook) Then
           GoTo next_c
       End If
       CAPGRP_COLUMN_NAME = Replace(capgrp_sheet, "LN", "Lijn")
       r.add_named_range_column range_name, CAPGRP_COLUMN_NAME
        
       ' fill capgrp_column with VLOOKUP
       formulaDef = str.subInStr(formulaTemplate, start_address, capgrp_sheet, main.CAPGRP_BULKCODE_STARTDATE_RANGE, main.STARTDATE_COLUMN_INDEX)
       Set formula_range = r.get_column(range_name, CAPGRP_COLUMN_NAME)
       Set formula_range = r.subset_range(formula_range, startrow:=1)
       r.fill_formula_range formula_range, formulaDef, True
        
       ' format formula range as datetime
       formula_range.Cells.NumberFormat = main.BULK_DATETIME_COLUMN_FORMAT
       
       ' get last column, first row
       lastFormulaRangeAddr = formula_range.Cells(1, 1).address
       
next_c:
    Next c
    
    ' create calculation range for SORTERING column, strip out the formating function formatDateVDMIShort/Long
    Dim capgrp_columns_range As Range
    Set rng0 = r.get_range(range_name)
    Set capgrp_columns_range = r.subset_range(rng0, startcol:=c0 + 1)

    Dim calcRange As Range, START_CALCULATION_ADDR As String, calcFormulasRange As Range, cl As Range, sortColumnInputAddress As String
    START_CALCULATION_ADDR = r.safe_offset(Range(lastFormulaRangeAddr), offset_column:=1).address
    r.copyRangeFormulas capgrp_columns_range, START_CALCULATION_ADDR, Worksheets(main.BULK_SHEET_NAME)
    Set calcRange = r.expand_range(START_CALCULATION_ADDR, rng0.Worksheet)
    Set calcFormulasRange = r.subset_range(calcRange, 2)
    For Each cl In calcFormulasRange.Cells
        cl.formula = "=" & returnStringWithinBrackets(cl.formula)
    Next

    ' calculation range formatting
    calcRange.columns.AutoFit
    calcRange.Cells.NumberFormat = main.BULK_DATETIME_COLUMN_FORMAT
    
    ' set as named range `main.BULK_ORDERS_CALC_RANGE_NAME`
    r.create_named_range main.BULK_ORDERS_CALC_RANGE_NAME, BulkSheet.name, calcRange.address, overwrite:=True, clear:=False
    
    ' add SORTERING column as value minimum of dates in calculation range `calcRange`
    r.add_named_range_column range_name, main.BULK_ORDERS_SORTING_COLUMN
    sortColumnInputAddress = calcRange.Rows(2).address
    columns_address = u.remove_dollar_sign(sortColumnInputAddress)
    formulaDef = str.subInStr("=MIN(@1)", columns_address)
    r.fill_formula_range r.get_column(range_name, main.BULK_ORDERS_SORTING_COLUMN), formulaDef, True
        
    ' set capgrp columns width to standard width
    Set rng0 = r.get_range(range_name)
    Set capgrp_columns_range = r.subset_range(rng0, startcol:=c0 + 1)
    capgrp_columns_range.ColumnWidth = capgrpColumnWidth
    
    ' set SORTERING column width to 16
    rng0.columns(rng0.columns.count).ColumnWidth = 16
    
    ' format header to bold and add borders
    rng0.Rows(1).Font.Bold = True
    r.add_all_borders rng0
    r.get_range(range_name).Select
    
    ' update bulk sorting
    main.update_bulk_sorting
    
    ' re focus on the original named range
    ' Set rng0 = r.get_range(range_name)
    ' rng0.Select
    ' Exit Sub
    
finally:
    BulkSheet.Activate
    Application.EnableEvents = True
    Exit Sub
End Sub

Sub update_bulk_sorting()
    If w.sheet_exists(main.BULK_SHEET_NAME) Then
       Dim rng0 As Range
       Set rng0 = r.get_range(main.BULK_ORDERS_RANGE_NAME)
       If r.column_exist(rng0, main.BULK_ORDERS_SORTING_COLUMN) Then
          r.sort_range_by_columns_2 rng0, main.BULK_ORDERS_SORTING_COLUMN
       End If
       ' return to ws0
    End If
End Sub

Function returnStringWithinBrackets(inputString As String) As String
    ' This function returns the substring within the first opening bracket "("
    ' and the last closing bracket ")" of the input string.
    '
    ' Parameters:
    ' inputString - The string from which to extract the substring.
    '
    ' Returns:
    ' A string that is between the first opening bracket and the last closing bracket.
    
    Dim startPos As Integer
    Dim endPos As Integer
    Dim extractedString As String
    
    ' Find the position of the first opening bracket "("
    startPos = InStr(1, inputString, "(")
    
    ' Find the position of the last closing bracket ")"
    endPos = InStrRev(inputString, ")")
    
    ' Check if both brackets are found and the positions are valid
    If startPos > 0 And endPos > startPos Then
        ' Extract the string between the brackets
        extractedString = Mid(inputString, startPos + 1, endPos - startPos - 1)
    Else
        ' If brackets are not found or positions are invalid, return an empty string
        extractedString = ""
    End If
    
    ' Return the extracted string
    returnStringWithinBrackets = extractedString
End Function

' get the capgrp columns on the BULK sheet
Function get_capgrp_columns_bulk_sheet() As Range
   Dim capgrp_sheets As collection, rng0 As Range, capgrp_sheet As String
   Set capgrp_sheets = main.get_capgrp_sheet_names()
   Set rng0 = r.get_range(main.BULK_ORDERS_RANGE_NAME)
   
   startcolindex = 9999
   endcolindex = 0
   For Each c In capgrp_sheets
       capgrp_sheet = c
       CAPGRP_COLUMN_NAME = Replace(CStr(capgrp_sheet), "LN", "Lijn")
       If r.column_exist(rng0, CAPGRP_COLUMN_NAME) Then
          col_index = r.get_column_index(rng0, CAPGRP_COLUMN_NAME)
          startcolindex = WorksheetFunction.Min(startcolindex, col_index)
          endcolindex = WorksheetFunction.MAX(endcolindex, col_index)
       End If
   Next c
   
   Debug.Print startcolindex, endcolindex
   
   ' return range
   Set get_capgrp_columns_bulk_sheet = r.subset_range(rng0, startcol:=startcolindex, endcol:=endcolindex)
   get_capgrp_columns_bulk_sheet.Select

End Function

Sub restore_isah_default_sorting()
    Dim rng0 As Range, ws0 As Worksheet
    Dim ws As Worksheet
    Dim sortKey1 As Range
    Dim sortKey2 As Range
    Dim sortKey3 As Range
    Dim sortKey4 As Range
    Dim default_sort_columns As Variant
    
    ' Set the worksheet and range variables
    Set ws0 = ThisWorkbook.Sheets(main.INPUT_ISAH_SHEET)
    Set rng0 = r.get_range(ws0)
    ' Set the sort range excluding the header row
    Set SortRange = rng0 'rng0.Offset(1).Resize(rng0.Rows.count - 1)
    
    ' Set the sort keys
    default_sort_columns = r.str_to_array(main.INPUT_ISAH_SHEET_SORT_KEY)
    Set sortKey1 = rng0.Rows(1).Find(default_sort_columns(0))
    Set sortKey2 = rng0.Rows(1).Find(default_sort_columns(1))
    Set sortKey3 = rng0.Rows(1).Find(default_sort_columns(2))
    Set sortKey4 = rng0.Rows(1).Find(default_sort_columns(3))
    
    ' Sort the range based on the sort keys
    With ws0.Sort
        .SortFields.clear
        .SortFields.Add key:=sortKey1, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=sortKey2, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=sortKey3, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=sortKey4, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange SortRange
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Function add_date_columns(range_name As String) As Range
    Dim rng0 As Range
    Set rng0 = r.get_range(range_name)
    Set ws = rng0.Worksheet
    
    ' check if "STARTDATE" exist
    If Not column_exist(rng0, main.STARTDATE_COLUMN) Then
       Debug.Print "STARTDATE does not exist"
       ws.Cells(rng0.Rows(1).row, rng0.columns.count + 1) = main.STARTDATE_COLUMN
       ' resize rng0
       Set rng0 = r.getResizedRange(rng0, 0, 1)
    End If
  
    If Not column_exist(rng0, main.ENDDATE_COLUMN) Then
       Debug.Print "ENDDATE does not exist"
       ws.Cells(rng0.Rows(1).row, rng0.columns.count + 1) = main.ENDDATE_COLUMN
       Set rng0 = r.getResizedRange(rng0, 0, 1)
    End If
    
    ' update the named range with date columns
    r.update_named_range range_name, rng0
    
    Set add_date_columns = rng0

End Function

'added on 20240209
Sub add_tank_lo_columns(capgrp As String)
    Dim rng0 As Range, column0 As Range
    Set rng0 = main.get_orders_range(capgrp)
    
    ' insert columns "Tank", "L/O" before "Bulkcode"
    If Not r.column_exist(rng0, "Tank") And Not r.column_exist(rng0, "L/O") Then
       r.InsertColumnIntoRange rng0, "Bulkcode", "Tank", ""
       r.InsertColumnIntoRange rng0, "Bulkcode", "L/O", ""
    End If
    Exit Sub
End Sub

Sub format_workdaytimes_range(range_name As String)
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

Sub assign_macro_to_btn(btn_name, macro_name)
    ActiveSheet.Shapes.Range(Array(btn_name)).Select
    Selection.OnAction = macro_name
End Sub


Sub create_buttons()
capgrp = "INPK"
    Dim wb As Workbook
    Set wb = ThisWorkbook
    init_buttons capgrp, True
End Sub

' input workdaytimes from sheet, returns array of (n_worktimes * n_workday, 3) with columns (starttime, endtime, ind_active)
Function get_workdaytimes_array(capgrp As String, startDate As Date)
    Dim date0 As Date
    Dim range_name As String
    range_name = capgrp & "_workdays"
    
    Set workdaytimes_range = r.get_range(range_name)
    Set workdays = r.get_column(range_name, 1, offset_row:=1)
    Set worktimes = r.get_row(range_name, 1, offset_column:=1)
    
    ' container
    ReDim work_startend_times(1 To workdays.Cells.count * worktimes.Cells.count, 1 To 3)
    
    c = 1
    For Each wd In workdays.Cells
        ' get the workday
        date0 = dt.vdmi_get_day_of_week(startDate, CStr(wd))
        For Each wt In worktimes.Cells
        wt0 = Split(wt, "-")(0)
        wt1 = Split(wt, "-")(1)
        dt0 = dt.get_datetime_value(date0, wt0)
        dt1 = dt.get_datetime_value(date0, wt1)
        'Debug.Print dt0, dt1
        work_startend_times(c, 1) = dt0
        work_startend_times(c, 2) = dt1
        work_startend_times(c, 3) = r.get_value(workdaytimes_range, wd, wt)
        c = c + 1
        Next wt
    Next wd
    
    get_workdaytimes_array = work_startend_times
End Function

' returns the first day of inputted year-week
Function get_capgrp_startdate(capgrp As String) As Date
    Dim isoyear As Integer, weeknumber As Integer
    ThisWorkbook.Sheets(capgrp).Activate
    isoyear = r.get_range(capgrp & "_input_year").Cells(2, 2).value
    weeknumber = r.get_range(capgrp & "_input_weeknumber").Cells(2, 2).value
    get_capgrp_startdate = dt.first_day_isoweek(weeknumber, isoyear)
End Function

Function get_capgrp_year(capgrp As String) As Integer
    get_capgrp_year = r.get_range(capgrp & "_input_year").Cells(2, 2).value
End Function

Function get_capgrp_weeknumber(capgrp As String) As Integer
    get_capgrp_weeknumber = r.get_range(capgrp & "_input_weeknumber").Cells(2, 2).value
End Function

Sub set_capgrp_weeknumber(capgrp_sheet As String, weeknumber As Integer)
    r.get_range(capgrp_sheet & "_input_weeknumber").Cells(2, 2).value = weeknumber
End Sub

Sub set_capgrp_year(capgrp_sheet As String, year As Integer)
    r.get_range(capgrp_sheet & "_input_year").Cells(2, 2).value = year
End Sub

Function get_capgrp_print_location(capgrp As String) As String
    ' Retrieves the print location for a specified capacity group (capgrp).
    ' If the print location is an absolute path, it validates the path.
    ' If the print location is not an absolute path, it joins with the documents path to construct an absolute path and validates it.
    '
    ' Parameters:
    ' capgrp - The name of the capacity group for which the print location is being retrieved.
    '
    ' Returns:
    ' A string representing the validated absolute path for the print location.
    
    Dim printLocation As String
    printLocation = r.get_range(capgrp & "_input_print_location").Cells(2, 2).value
    
    ' Check if the print location is an absolute path
    If left(printLocation, 1) = "\" Or InStr(printLocation, ":") > 0 Then
        ' Validate the absolute path
        If validate_path(printLocation) Then
            get_capgrp_print_location = printLocation
        Else
            MsgBox "Directory does not exist: " & printLocation
            get_capgrp_print_location = ""
        End If
    Else
        ' Join with the documents path to construct an absolute path
        Dim documentsPath As String
        documentsPath = os.getDocumentsPath() ' Assuming getDocumentsPath is defined in os.bas
        Dim fullPath As String
        fullPath = os.pathJoin(documentsPath, printLocation)
        
        ' Validate the constructed path
        If validate_path(fullPath) Then
            get_capgrp_print_location = fullPath
        Else
            MsgBox "Directory does not exist: " & fullPath
            get_capgrp_print_location = ""
        End If
    End If
End Function

Function validate_path(path As String) As Boolean
    ' Validates a given path by checking if the directory part exists.
    '
    ' Parameters:
    ' path - The path to validate.
    '
    ' Returns:
    ' True if the directory part of the path exists, False otherwise.
    Dim parts As Variant
    parts = os.pathSplit(path)
    Dim directoryPart As String
    directoryPart = parts(0)
    
    If os.isDir(directoryPart) Then
        validate_path = True
    Else
        validate_path = False
    End If
End Function

Sub set_capgrp_print_location(capgrp As String, Optional overwrite As Boolean = False)
    ' Sets the print location for a specified capacity group (capgrp).
    ' If the print location is empty, it sets a default value.
    ' If overwrite is True, it overwrites the current print location.
    '
    ' Parameters:
    ' capgrp - The name of the capacity group for which the print location is being set.
    ' overwrite - An optional boolean indicating whether to overwrite the current print location.
    Dim range_name As String
    range_name = capgrp & "_input_print_location"
    
    If overwrite Or r.get_range(range_name).Cells(2, 2).value = "" Then
        r.get_range(range_name).Cells(2, 2).value = capgrp & "_planning.pdf"
    End If
End Sub



' This subroutine sets default values for capgrp sheet inputs weeknumber and year
Sub set_default_weeknumber_year(capgrp_sheet As String)

 Dim current_prodwk As Integer, current_year As Integer, new_prodwk As Integer, prodwk_year As Integer
 ' Get first new_prodwk from ISAH INPUT
 new_prodwk = WorksheetFunction.Min(r.get_column(main.get_isah_input_range, main.PRODWK_COLUMN, offset_row:=1))

 ' set ProdWk and Year if not set
 current_prodwk = main.get_capgrp_weeknumber(capgrp_sheet)
 current_year = main.get_capgrp_year(capgrp_sheet)
 If current_prodwk = 0 Then
    Call main.set_capgrp_weeknumber(capgrp_sheet, new_prodwk)
 End If
 
 If current_year = 0 Then
    prodwk_year = dt.determine_year_based_on_weeknum(current_prodwk)
    Call main.set_capgrp_year(capgrp_sheet, prodwk_year)
 End If
 
End Sub

' 20230714: make the week overflow row bold
Sub update_start_end_times(capgrp As String, Optional dbg As Boolean = False)
    'get the workdaytimes array, starting from startDate
    Dim startDate As Date
    startDate = main.get_capgrp_startdate(capgrp)
    
    ' compute the start and end times of capgrp starting from startDate
    ' check if orders range has interior cells
    Dim ord_range As Range, range_name As String
    range_name = capgrp & "_orders"
    Set ord_range = r.get_range(range_name)
    If ord_range.Rows.count <= 1 Then
       Exit Sub
    End If
    
    start_end_times = calculate_start_end_times(capgrp, startDate, dbg:=dbg)
    
    ' paste start_end_times columns in `capgrp'_orders
    Dim ws As Worksheet, startdate_column_address As String
    Set ws = ThisWorkbook.Sheets(capgrp)
    range_name = capgrp & "_orders"
    startdate_column_address = r.get_column(range_name, main.STARTDATE_COLUMN, ws:=ws, offset_row:=1).Cells(1, 1).address
    a.pasteArray start_end_times, startdate_column_address, ws
    
    ' fit columns
    r.autofit_columns_rows r.get_column(range_name, main.STARTDATE_COLUMN, ws:=ws)
    r.autofit_columns_rows r.get_column(range_name, main.ENDDATE_COLUMN, ws:=ws)
    
    ' determine the overflow row index in orders_rng and set to BOLD font
    Dim orders_rng As Range, orders_values_rng
    Set orders_rng = get_orders_range(capgrp)
    ' TODO rmltr
    'enddates = r.get_column_values(orders_rng, main.ENDDATE_COLUMN)
    enddates = a.getArrayColumnValues(orders_rng.value, main.ENDDATE_COLUMN)
    
    ' if enddates not array, then exit
    If IsArray(EndDate) Then
        overflow_row_index = main.find_week_overflow_row(enddates)
        Debug.Print "overflow_row_index is ", overflow_row_index
        If overflow_row_index > 0 Then
           orders_rng.Rows(overflow_row_index + 1).Font.Bold = True
        End If
    End If
End Sub

'find the row index where orders first overflow into the next week (second monday)
Function find_week_overflow_row(enddates)
    Dim i As Long, mondayCount As Integer, r As Long
    Dim uniqueMondays As New collection, mondayFound As Boolean
    Dim enddatetime As Date
    
    ' Assume that enddates has been assigned values
    
    ' Initialize variables
    mondayCount = 0
    r = 0
    mondayFound = False
    
    ' Loop over the dates in the enddates array
    On Error Resume Next
    For i = LBound(enddates, 1) To UBound(enddates, 1)
        On Error GoTo 0
        Debug.Print IsArray(enddates)
        enddatetime = enddates(i, 1) 'cast to proper Date
        On Error Resume Next
        enddate0 = dt.set_date_timepart(enddatetime, "00:00") 'set time to 00:00
        ' Check if the date is a Monday
        If Weekday(enddate0, vbMonday) = 1 Then
            ' Check if the Monday is unique
            uniqueMondays.Add CDate(enddate0), CStr(enddate0)
            If Err.Number = 0 Then
                ' Increment the count of unique Mondays
                mondayCount = mondayCount + 1
                If mondayCount = 2 Then
                    ' If this is the second unique Monday, set r to the row number and break the loop
                    r = i
                    mondayFound = True
                    Exit For
                End If
            Else
                ' Reset error if duplicate item
                Err.clear
            End If
        End If
    Next i
    On Error GoTo 0
    
    ' Check if the second Monday was not found
    If Not mondayFound Then
        r = 0
    End If
    
    find_week_overflow_row = r

End Function

Sub test_calculate_start_end_times()
capgrp = "INPK"
    Dim startDate As Date
    startDate = get_capgrp_startdate(capgrp)
start_end_times = calculate_start_end_times(capgrp, startDate, dbg:=True)

Debug.Print a.numArrayRows(start_end_times)

End Sub

' returns (n_orders,2) array with columns block_starttime, block_endtime
Function calculate_start_end_times(capgrp As String, startDate As Date, Optional dbg As Boolean = True)

    'get the workdaytimes array, starting from startDate
    wdt = get_workdaytimes_array(capgrp, startDate)

    ' now get from ORDERS from worksheet
    Dim ord As Variant, ord_range As Range, article_code As String, duration As Double
    range_name = capgrp & "_orders"
    
    Set ord_range = r.get_range(range_name)
    ord = r.get_range_values(ord_range, offset_row:=1, offset_column:=0).value
    dur_index = r.get_column_index(ord_range, main.DURATION_COLUMN)
    art_index = r.get_column_index(ord_range, main.ART_COLUMN)
    
    ' get the duration, articles array and create empty array for the calculated startdates `startdates_out`
    duration_arr = a.getArrayColumn(ord, dur_index)
    articles_arr = a.getArrayColumn(ord, art_index)
    startdates_out = a.create_array(a.numArrayRows(ord), 1)
    
    'get the first starttime from the workdaytimes array (wdt)
    Dim startTime As Double, endtime As Double
    jT = a.numArrayRows(wdt)
    startTime = wdt(1, 1)
    
    ' create empty start_end_times array
    Dim start_end_planned() As Variant, j0 As Long, dur As Double
    ReDim start_end_times(LBound(articles_arr) To UBound(articles_arr), 1 To 2)
    
    ' loop over each article i
    j0 = 1
    For Each i In a.getRowIndexes(articles_arr)
        duration = duration_arr(i, 1)
        If Len(duration) <= 0 Then
           GoTo next_art
        End If
        
        block_starttime = find_block_starttime(wdt, j0, startTime, dbg:=dbg)
        j0 = block_starttime(0)
        startTime = block_starttime(1)
        start_end_times(i, 1) = startTime
        dur = CDbl(duration_arr(i, 1))
        block_endtime = find_block_endtime(wdt, j0, startTime, dur, dbg:=dbg)
        start_end_times(i, 2) = block_endtime(1)
        
        ' set endtime as next starttime
        startTime = block_endtime(1)
        j0 = block_endtime(0)
        
        ' checks:
        starttime_i = start_end_times(i, 1)
        endtime_i = start_end_times(i, 2)
        If starttime_i > endtime_i Then
           Err.Raise 1001, Description:="Startime is greater than endtime at i " & i
        End If
next_art:
    Next i
        
    calculate_start_end_times = start_end_times
End Function

Function find_block_starttime(wdt As Variant, j0 As Long, ByVal startTime, Optional prep As Long = 0, _
Optional dbg As Boolean = True) As Variant
    Dim jT As Long
    jT = UBound(wdt, 1) - LBound(wdt, 1) + 1
    
    Dim starttime2 As Double
    starttime2 = m.round_up_to_nearest_quarter(CDbl(startTime))
    If j0 > jT Then
        last_block_time = get_last_block(wdt)
        If dbg Then
           Debug.Print "last block reached, return starttime", last_block_time(1)
        End If
        find_block_starttime = last_block_time
        Exit Function
    End If
    
    Dim j As Long, ind_active As Long, startworktime As Double, endworktime As Double
    For j = j0 To jT
        startworktime = wdt(j, 1)
        endworktime = wdt(j, 2)
        ind_active = wdt(j, 3)
        If (m.gte_dbl(starttime2, startworktime) Or m.lte_dbl(starttime2, endworktime)) And ind_active = 1 Then ' use gte_dbl because of precision issue
            find_block_starttime = Array(j, starttime2)
            If dbg Then
               Debug.Print "found block", j, "with starttime", starttime2
            End If
            Exit Function
        ' added 20230601 to prevent first starttime to be "stuck" on 6:00
        ElseIf j < jT Then
            starttime2 = wdt(j + 1, 1)
        End If
    Next j
    
    If j = jT + 1 Then
       If dbg Then
       Debug.Print "last block reached", j
       End If
    End If
End Function

Function find_block_endtime(wdt As Variant, j0 As Long, startTime As Double, dur As Double, Optional dbg As Boolean = True _
) As Variant
    Dim dur1 As Double, startworktime As Double, endworktime As Double, j As Long
    dur1 = dur ' is in hours
    
    ' check if end of workdaytime reached
    Dim jT As Long
    jT = UBound(wdt, 1)
    If j0 > jT Then
        find_block_endtime = get_last_block(wdt)
        If dbg Then
        Debug.Print "find_block_endtime: last block reached", j
        End If
        'Err.Raise 1001, Description:="last block reached"
        find_block_endtime = get_last_block(wdt)
        Exit Function
    End If
    
    Dim endtime As Double
    endtime = dt.add_hours(startTime, dur1)

    'cases: 1. starttime, endtime fits in active block j0
    startworktime = wdt(j0, 1)
    endworktime = wdt(j0, 2)
    ind_active = wdt(j0, 3)
    If (m.gte_dbl(startTime, startworktime) And m.lte_dbl(endtime, endworktime)) And ind_active = 1 Then  ' use gte_dbl because of precision issue
        endtime = m.round_up_to_nearest_quarter(endtime)
        find_block_endtime = Array(j0, endtime)
        If dbg Then
        Debug.Print "endtime ", endtime, "fits in block", j
        End If
        Exit Function
    End If
    
    ' 2. endtime does not fit in block j0, fill remainer of current block and continue to next block with residual duration dur1
    If j0 = jT Then
       find_block_endtime = get_last_block(wdt)
        If dbg Then
        Debug.Print "endtime ", endtime, "doesnt fit in block", j0, "but is last block"
        End If
    End If
    dur1 = dur1 - 24 * (endworktime - startTime)
    If dbg Then
       Debug.Print "moving to next block", j0 + 1, "with residual duration", dur1
    End If
    
    For j = j0 + 1 To jT
        startworktime = wdt(j, 1)
        endworktime = wdt(j, 2)
        ind_active = wdt(j, 3)
        endtime = dt.add_hours(startworktime, dur1)
        If ind_active = 1 Then
            If (m.lte_dbl(endtime, endworktime)) Then ' fits in block j
                ' if fits in block then round up to nearest quarter
                endtime = m.round_up_to_nearest_quarter(endtime)
                find_block_endtime = Array(j, endtime)
                If dbg Then
                Debug.Print "endtime ", endtime, "fits in block", j
                End If
                Exit Function
            Else 'calculate residual duration and move to next block j+1
                dur1 = dur1 - 24 * (endworktime - startworktime)
                If dbg Then
                Debug.Print "moving to next block", j + 1, "with residual duration", dur1
                End If
            End If
        End If
    Next j
    
    find_block_endtime = get_last_block(wdt)
    If dbg Then
    Debug.Print "find_block_endtime: last block reached", j
    End If
    'Err.Raise 1001, Description:="last block reached"
       
End Function

Function get_last_block(wdt As Variant) As Variant
    Dim lastRow As Long
    lastRow = UBound(wdt, 1)
    get_last_block = Array(lastRow + 1, wdt(lastRow, 2))
End Function

' update ISAH data from template
Sub clear_orders_range(Optional capgrp As String = "")
    ' Clears the orders range for a specified capacity group (capgrp).
    ' This subroutine clears the contents and formatting of the orders range,
    ' and reduces the named range to a single cell.
    '
    ' Parameters:
    ' capgrp - An optional string specifying the name of the capacity group for which the orders range is being cleared.
    '          If not provided, the active sheet's name is used.
    Dim range_name As String, ordersRng As Range, rangeToClear As Range
    If capgrp = "" Then
       capgrp = ActiveSheet.name
    End If
    range_name = main.get_orders_range_name(capgrp)
    If r.name_exist(range_name) Then
       Set ordersRng = main.get_orders_range(capgrp)
       Set rangeToClear = r.getResizedRange(ordersRng.Cells(1, 1), add_rows:=998, add_cols:=main.XL_MAX_NUMBER_COLUMNS - 1)
       r.clear_range rangeToClear, clear_formatting:=True
       ' reduce orders_range to single row
       r.subsetNamedRange range_name, 1, 1, 1, 1 ' ordersRng.columns.count, ThisWorkbook
    End If
End Sub

Public Sub clear_all_capgrp_sheets()
    Dim worktimes_range As Range
    Dim capgrp_sheet As String
    
    ' clear capgrp sheets
    For Each capgrp_sheet0 In main.get_capgrp_sheet_names()
        capgrp_sheet = CStr(capgrp_sheet0)
        If main.P_DEBUG Then
           Debug.Print "clearing capgrp sheet:", capgrp_sheet
        End If
        main.clear_orders_range CStr(capgrp_sheet)
        Set worktimes_range = get_worktimes_values_range(capgrp_sheet)
        worktimes_range.Cells.value = 1
        
        ' clear capgrp sheet inputs
        main.set_capgrp_weeknumber capgrp_sheet, 0
        main.set_capgrp_year capgrp_sheet, 0
nx_capgrp:
    Next
        
    ' clear isah input
    Dim rng0 As Range
    Set rng0 = main.get_isah_input_range()
    If rng0.Rows.count > 1 Then
        r.subset_range(rng0, 2).Cells.ClearContents
    End If
    
    ' clear BULK
    r.clear_range_values main.BULK_ORDERS_RANGE_NAME, clear_formatting:=True
    
    ' clear EXPORT_ISAH
    r.clear_range_values main.ISAH_STAGING_RANGE_NAME, clear_formatting:=True
    
    ' clear NIEUW
    w.clearWorksheet main.ISAH_NEW_ORDERS_SHEET_NAME
End Sub

' TODO: replace capgrp with capgrp_sheet
Sub update_orders_range(capgrp As String)
    ' Updates the orders range for a specified capacity group (capgrp) with new data.
    ' This subroutine clears the existing orders range, sets default inputs, copies selected orders,
    ' and performs formatting and sorting operations on the updated range.
    '
    ' Parameters:
    ' capgrp - The name of the capacity group for which the orders range is being updated.
    '
    ' The subroutine performs the following actions:
    ' - Clears the existing orders range and sets default inputs for the capacity group.
    ' - Copies selected orders from the ISAH input sheet to the orders range.
    ' - Adds necessary columns such as start and end dates.
    ' - Sorts the range by bulk code and applies color formatting.
    ' - Updates start and end times for the orders.
    ' - Adjusts column widths and moves buttons to their designated positions.
    ' - Updates sorting on the BULK sheet if times change on the capgrp sheet.
    ' - Stores the current state in the WorksheetStateCollection if enabled.
    Dim ws1 As Worksheet, ws0 As Worksheet, wb0 As Workbook
    Dim r1 As Long, rng1 As Range, range_name As String
    Set wb0 = ThisWorkbook

    ' 1 clear target orders range
    range_name = capgrp & "_orders"
    If Not r.name_exist(range_name) Then
       Debug.Print "named range does not exist, create: " & range_name
       main.init_orders_range capgrp, range_name, False
    End If
    
    ' 2 clear orders range, set the default capgrp inputs and copy selected orders and refit orders range to copied values
    main.clear_orders_range capgrp
    main.set_default_weeknumber_year capgrp
    main.copy_selected_orders capgrp, clear_ws:=False, remove_filter:=True
    main.fit_order_range_to_values capgrp
    
    ' add STARTDATE, ENDDATE columns to orders_range
    add_date_columns range_name
    
    ' fit named range to include new columns
    main.fit_order_range_to_values capgrp
    Set rng1 = r.get_range(range_name, wb:=wb0)
    Set ws0 = rng1.Worksheet
    
    ' insert `volgnummer`
    insert_volgnummer_into_orders capgrp
    
    If rng1.Rows.count <= 1 Then
       Debug.Print "update_orders_range: no orders found for capgrp: `" & capgrp & "`"
       Exit Sub
    End If
    
    ' 3 sort on bulkcode and color formatting bulkcode column
    r.sort_range_by_columns rng1, main.BULKCODE_COLUMN
    
    ' 4.1 update start end times if orders are not empty
    If rng1.Rows.count > 1 Then
       main.update_start_end_times capgrp
       ' 202307 4.2 20 insert the #pallets formula
       main.insert_number_of_pallets_formula capgrp
       
       ' 20240209 fill firt row of columns Tank, L/O
       main.add_tank_lo_columns capgrp

       ' column/row formatting
       main.update_orders_columns_width capgrp
    End If
    
    ' 5 position buttons
    main.position_buttons capgrp
     
    ' 5.1 format Starttijd, Eindtijd + general formatting
    main.format_orders_range capgrp, rng1
    
    ' 6 update sorting on BULK if times change on capgrp sheet
    main.update_bulk_sorting
    ws0.Activate
    
    ' 7 store state in WorksheetStateCollection
    If main.P_STORE_STATE Then
       WorksheetStateCollection.Add "1"
    End If
    
    Exit Sub
End Sub

Sub format_orders_range(capgrp As String, Optional orders_rng As Range = Nothing)
    If orders_rng Is Nothing Then
       Set orders_rng = main.get_orders_range(capgrp)
    Else
       Set orders_rng = orders_rng 'performance
    End If
    
    r.formatRangeColumns orders_rng, main.ORDERS_RANGE_COLUMN_FORMATS
    Dim i As Long
    If orders_rng.Rows.count > 1 Then
        For i = 1 To orders_rng.columns.count
            With orders_rng
            .HorizontalAlignment = xlCenter
            End With
        Next
        
        ' color formatting
        main.update_orders_color_format capgrp
    End If
    
    Exit Sub
    
End Sub

Sub update_orders_range_formulas(capgrp As String, Optional orders_rng As Range = Nothing)
    Dim rng1 As Range
    If orders_rng Is Nothing Then
       Set rng1 = main.get_orders_range(capgrp)
    Else
       Set rng1 = orders_rng 'performance
    End If

    If rng1.Rows.count > 1 Then
       main.update_start_end_times capgrp
       ' 202307 4.2 20 insert the #pallets formula
       main.insert_number_of_pallets_formula capgrp
       
       ' 20240209 fill firt row of columns Tank, L/O
       main.add_tank_lo_columns capgrp

       ' column/row formatting
       main.update_orders_columns_width capgrp
    End If
End Sub

Function get_orders_range_name(capgrp_name As String) As String
    get_orders_range_name = capgrp_name & "_orders"
End Function

Sub fit_order_range_to_values(capgrp_name As String)
    ' Adjusts the size of the orders range to fit the actual values for a specified capacity group (capgrp_name).
    ' This subroutine resizes the named range for orders to match the number of rows with values,
    ' ensuring that the range accurately reflects the data it contains.
    '
    ' Parameters:
    ' capgrp_name - The name of the capacity group for which the orders range is being adjusted.
    Dim range_name As String, ordersRng As Range, search_column_index As Long
    range_name = main.get_orders_range_name(capgrp_name)
    Set ordersRng = main.get_orders_range(capgrp_name)
    
    ' find last row of values in `orders_range`, search down from the first row of the range, search up from Range("A999")
    r.fit_named_range_to_values range_name, wb:=ThisWorkbook, searchUpRange:=Range("A999")
End Sub

Sub test_insert_number_of_pallets_formula()
    insert_number_of_pallets_formula "LN 1"
End Sub

Sub insert_number_of_pallets_formula(capgrp As String)
    Dim rng As Range, name As String, formulaTemplateString As String

    ' capgrp = "LN_1"
    
    ' Get the named range
    Set rng = main.get_orders_range(capgrp)
    name = capgrp & "_orders"

    ' Check if column NUMBER_OF_PALLETS_NAME has been added to named range
    If Not r.column_exist(rng, main.NUMBER_OF_PALLETS_NAME) Then
    
        ' Get the column index of "Resources"
        Dim resourcesColumnIndex As Long
        resourcesColumnIndex = get_column_index(rng, "Resources")
        
        ' Get the column index of "Flesformaat"
        Dim flesformaatColumnIndex As Long
        flesformaatColumnIndex = get_column_index(rng, "Flesformaat")
        
        ' Insert a new column between "Resources" and "Flesformaat"
        rng.columns(flesformaatColumnIndex).Insert shift:=xlToRight
        rng.columns(flesformaatColumnIndex).Cells(1) = main.NUMBER_OF_PALLETS_NAME
        
        ' Get the column index of the newly inserted column
        Dim newColumnIndex As Long
        newColumnIndex = flesformaatColumnIndex
        
        ' Update the named range to include the newly inserted column
        update_named_range name, rng, ActiveWorkbook
    End If
    
    If rng.Rows.count > 1 Then
    
        ' Get the column index of NUMBER_OF_PALLETS
        Dim numberOfPalletsIndex As Long
        numberOfPalletsIndex = r.get_column_index(rng, main.NUMBER_OF_PALLETS_NAME)
        
        ' Get the address of the second row of the column named "Artikel"
        Dim artikelColumnIndex As Long, quantityColumnIndex As Long
        artikelColumnIndex = get_column_index(rng, main.ART_COLUMN)
        quantityColumnIndex = get_column_index(rng, main.QTY_COLUMN)
        
        ' Remove the dollar sign from the address
        Dim address As String, qty_address
        address = Replace(rng.Cells(2, artikelColumnIndex).address, "$", "")
        qty_address = Replace(rng.Cells(2, quantityColumnIndex).address, "$", "")
        
        ' Fill the formula formulaTemplateString in the newly inserted column
        Dim formulaDefinition As String, formulaRange As Range
        formulaTemplateString = "=IF(@1="""","""",CEILING(@2/VLOOKUP(@1,@3,3,FALSE),1))"
        
        formulaDefinition = str.subInStr(formulaTemplateString, address, qty_address, main.NUMBER_PER_PALLET_NAMED_RANGE)
        Set formulaRange = rng.columns(numberOfPalletsIndex)
        If formulaRange.Rows.count > 1 Then
            Set formulaRange = formulaRange.Offset(1, 0).Resize(formulaRange.Rows.count - 1, 1)
            'formulaRange.Activate
            r.fill_formula_range formulaRange, formulaDefinition
        End If
    
    End If
End Sub

Sub insert_record(capgrp As String)
    Dim orders_rng As Range, ws As Worksheet, range_name As String, abs_row As Long
    Dim active_row As Range, offset_row As Range, new_orders_range As Range
    range_name = capgrp & "_orders"
    abs_row = CLng(ActiveCell.row)
    
    Set orders_rng = r.get_range(range_name)
    Set offset_row = r.safe_offset(orders_rng.Rows(orders_rng.Rows.count), 1)
    Set ws = orders_rng.Worksheet
    
    If Not Intersect(orders_rng, ActiveCell) Is Nothing And abs_row > 1 Then
        ' don't insert when on the header row
        If orders_rng.Cells(1, 1).row = ActiveCell.row Then
           Exit Sub
        End If
        Set active_row = main.get_row_in_named_range(range_name, abs_row)
        
        'Insert the record into the named range
        active_row.Insert shift:=xlDown
        
        ' Update capgrp formulas
        main.update_orders_range_formulas capgrp, orders_rng
        
        ' general formatting
        main.format_orders_range capgrp
        
        ' Recalculate volgnummers
        main.calculate_volgnummer capgrp
        
    ElseIf Not Intersect(offset_row, ActiveCell) Is Nothing Then
        ' combine current range with offset row
        Set new_orders_range = Range(orders_rng.address, offset_row.address)
        r.update_named_range range_name, new_orders_range
        
        ' Update capgrp formulas
        main.update_orders_range_formulas capgrp, new_orders_range
        
        ' general formatting
        main.format_orders_range capgrp
        
        ' Recalculate volgnummers
        main.calculate_volgnummer capgrp
    End If
End Sub

Sub delete_record(capgrp As String)
    Dim rng0 As Range, ws As Worksheet, range_name As String, abs_row As Long
    Dim active_row As Range
    range_name = capgrp & "_orders"
    abs_row = CLng(ActiveCell.row)
    
    Set rng0 = r.get_range(range_name)
    Set ws = rng0.Worksheet
    
    If Not Intersect(rng0, ActiveCell) Is Nothing And abs_row > 1 Then
        Set active_row = main.get_row_in_named_range(range_name, abs_row)
        
        'Insert the record into the named range
        active_row.Delete shift:=xlUp
        
        'Recalculate start enddates
        main.update_start_end_times capgrp
        
        ' Recalculate volgnummers
        main.calculate_volgnummer capgrp
    End If
End Sub

Function get_row_in_named_range(range_name As String, row_num As Long) As Range
    'Get the row in the named range using the absolute row number
    Dim first_row As Long, rng0 As Range, ws As Worksheet
    Set rng0 = r.get_range(range_name)
    Set ws = rng0.Worksheet
    first_row = ws.Range(rng0.address).row
    Set get_row_in_named_range = ws.Range(rng0.address).Rows(row_num - first_row + 1)
End Function

' orders color formatting
Sub update_orders_color_format(capgrp As String)
    Dim rng0 As Range, rng1 As Range, unique_bulkcodes As collection
    Set rng0 = r.get_range(capgrp & "_orders")
    Set rng1 = r.get_column_values(rng0, main.BULKCODE_COLUMN)
    If Not rng1 Is Nothing Then
        Set unique_bulkcodes = a.as_collection(r.get_unique_vals(rng1))
        If unique_bulkcodes.count > 0 Then
           fill_bulkcode_color rng1, unique_bulkcodes
        End If
    End If
End Sub

' orders column height, width
Sub update_orders_columns_width(capgrp As String)
   Dim rng0 As Range
   Set rng0 = r.get_range(capgrp & "_orders")
   r.autofit_columns_rows rng0

   'set widths of columns Tank, L/O to `TANK_LO_COLUMN_WIDTH`
   r.get_column(rng0, "Tank").ColumnWidth = main.TANK_LO_COLUMN_WIDTH
   rng0.columns(1).ColumnWidth = Round(main.COLUMN_AB_WIDTH / 5.6, 0)
   rng0.columns(2).ColumnWidth = Round(main.COLUMN_AB_WIDTH / 5.6, 0)
   rng0.columns(3).ColumnWidth = Round(main.COLUMN_CD_WIDTH / 5.6, 0)
   rng0.columns(4).ColumnWidth = Round(main.COLUMN_CD_WIDTH / 5.6, 0)
   'r.get_column(rng0, "L/O").ColumnWidth = main.TANK_LO_COLUMN_WIDTH => autofit is better bsc of first cell of column F
    
End Sub

' COLORS: move later to separate module
Function create_yellow_gradient(i As Long, n_shades As Long) As Variant
    Dim r As Integer, g As Integer, b As Integer
    r = 255
    g = 255 - (i - 1) * WorksheetFunction.Floor(255 / n_shades, 1)
    b = 0
    create_yellow_gradient = Array(r, g, b)
End Function

Function create_green_red_gradient(i As Long, n_shades As Long) As Variant
    Dim r As Integer, g As Integer, b As Integer
    r = (i) * WorksheetFunction.Floor(255 / n_shades, 1)
    g = 255 - (i) * WorksheetFunction.Floor(255 / n_shades, 1)
    b = 0
    create_green_red_gradient = Array(r, g, b)
End Function

Function get_color_palette(n_colors As Integer) As collection
    Dim i As Long
    Dim color_palette As collection
    Dim step_value As Long
    Dim r As Long, g As Long, b As Long
    Dim lower_bound As Long, upper_bound As Long

    ' Initialize the color palette
    Set color_palette = New collection

    ' Define the lower and upper bounds for the color values
    lower_bound = 100
    upper_bound = 255

    ' Calculate the step value based on the number of colors and the new range
    step_value = (upper_bound - lower_bound) \ n_colors
    
    ' Generate the colors
    For i = 0 To n_colors - 1
        r = lower_bound + (i * step_value) Mod (upper_bound - lower_bound + 1)
        g = lower_bound + (2 * i * step_value) Mod (upper_bound - lower_bound + 1)
        b = lower_bound + (3 * i * step_value) Mod (upper_bound - lower_bound + 1)

        ' Add the color to the palette
        color_palette.Add Array(r, g, b)
    Next i

    ' Return the color palette
    Set get_color_palette = color_palette
End Function

Function get_random_color() As Variant
    Dim r As Long, g As Long, b As Long
    
    ' Generate random RGB values, excluding 0 and 255
    r = Int((254 - 1 + 1) * Rnd + 1)
    g = Int((254 - 1 + 1) * Rnd + 1)
    b = Int((254 - 1 + 1) * Rnd + 1)
    
    ' Return the color as an RGB array
    get_random_color = Array(r, g, b)
End Function

Function get_random_color_palette(n_colors As Integer) As collection
    Dim i As Long
    Dim colors As Object
    Dim newColor As Variant
    Dim color_palette As collection

    ' Initialize the colors dictionary and color_palette collection
    Set colors = CreateObject("Scripting.Dictionary")
    Set color_palette = New collection

    ' While we have less colors than needed
    While colors.count < n_colors
        ' Generate a random color
        newColor = get_random_color()
        
        ' Convert color array to string to use as key in the dictionary
        colorKey = Join(newColor, ",")
        
        ' If this color is not already in the dictionary, add it
        If Not colors.exists(colorKey) Then
            colors.Add colorKey, Nothing
            color_palette.Add newColor
        End If
    Wend

    ' Return the color palette
    Set get_random_color_palette = color_palette
End Function

Function get_random_color_indices(N As Integer) As collection
    Dim colorIndices As New collection
    Dim randNum As Integer
    Dim i As Integer
    Dim j As Integer
    Dim isExists As Boolean

    For i = 1 To N
        Do
            ' Generate a random number from 2 to 56
            randNum = Int((56 - 2 + 1) * Rnd + 2)

            ' Check if the number exists in the colorIndices
            isExists = False
            For j = 1 To colorIndices.count
                If colorIndices.item(j) = randNum Then
                    isExists = True
                    Exit For
                End If
            Next j
            
            ' If the number exists, then we add 1 to the number and continue the loop
            ' If it exceeds 56, we reset it to 2
            If isExists Then
                randNum = randNum + 1
                If randNum > 56 Then
                    randNum = 2
                End If
            End If
            
        Loop While isExists

        ' Add the number to colorIndices
        colorIndices.Add randNum

    Next i
    
    Set get_random_color_indices = colorIndices
End Function

Sub fill_bulkcode_color_rbg(rng As Range, codes As collection)
    Dim count As Integer, cell As Range, i As Long, rbg_colors As collection
    count = 1
    Set rbg_colors = clls.alternate_items(get_color_palette(codes.count))
    For Each code In codes
        rgb_array = rbg_colors.item(count)
        For Each cell In rng.Cells
            If CStr(cell.value) = CStr(code) Then
                cell.Interior.Color = RGB(rgb_array(0), rgb_array(1), rgb_array(2))
            End If
        Next cell
        count = count + 1
    Next code
End Sub

Function get_color_index_light(index As Integer)
    Dim arr As Variant
    arr = Array(11, 25, 49, 51, 52, 55, 56)
    ' first modulo of 56 (last colorindex)
    cIndex = index Mod 56
    ' then exclude black and white
    If cIndex = 0 Or cIndex = 1 Then
    cIndex = 3
    ElseIf cIndex = 2 Then
    cIndex = 4
    ElseIf a.array_contains(arr, cIndex) Then
    cIndex = cIndex + 1 Mod 56
    Else
    cIndex = cIndex
    End If
    get_color_index_light = cIndex
End Function

Sub test_get_color_index_light()
    Dim color_indices As collection
    integers = a.create_integer_vector(1, 100)
    Set color_indices = main.get_color_indices_light(N:=100)
    'For i = 1 To 100
       'Debug.Print i, color_indices.item(i)
    'Next i
End Sub

Function get_color_indices_light(N As Integer, Optional skip_numbers As Variant) As collection
    Dim i As Integer, j As Integer
    Dim output As New collection
    
    If IsMissing(skip_numbers) Then
       skip_numbers = Array(1, 2, 9, 11, 25, 30, 49, 51, 52, 55, 56) ' exclude white, black and dark colors
    End If
    
    'handle case if all possible j are in skip_numbers
    'If n = a.get_array_len(skip_numbers) Then
    '   Debug.Print "check array skip_numbers"
    '   Set get_color_indices_light = Nothing
    'End If
    
    'Loop from 1 to n
    j = 1
    For i = 1 To N
        j = j Mod 56
        'Check if j is in skip_numbers and adjust accordingly
        Do While a.array_contains(skip_numbers, j)
            j = j + 1
        Loop
        If j > 56 Then
           j = 3
        End If
        output.Add j
        j = j + 1
    Next i
    
    'Check if the length of the output collection equals n
    Debug.Print output.count
    If output.count <> N Then
        Err.Raise 9999, , "The output collection length does not equal to n."
    End If
    
    Set get_color_indices_light = output
    
    Exit Function

End Function

Sub fill_bulkcode_color(rng As Range, codes As collection)
    Dim count As Integer, cell As Range, i As Long, color_indices As collection, sorted_codes As collection
    count = 1
    Set sorted_codes = clls.sort_collection(codes) ' sort the passed bulkcodes such that each bulkcode maps to the same colorindex
    Set color_indices = main.get_color_indices_light(sorted_codes.count)
    For Each code In sorted_codes
        color_index = color_indices.item(count)
        For Each cell In rng.Cells
            If CStr(cell.value) = CStr(code) Then
                cell.Interior.ColorIndex = color_index
            End If
        Next cell
        count = count + 1
    Next code
End Sub

' ISAH NIEUW ORDERS
Function get_new_orders_range() As Range
    Set get_new_orders_range = r.get_range(ThisWorkbook.Sheets(main.ISAH_NEW_ORDERS_SHEET_NAME))
End Function

' ISAH DATABASE EXPORT

' create connection to isah under current credentials/connection string
Function getISAHconnection() As ADODB.Connection
    Set getISAHconnection = db.openDBconn(main.getISAHconnstr())
End Function

' test isah connection
Sub isah_export_test_connection()
    If checkIsahTestQuery() Then
        MsgBox "verbinding met ISAH geslaagd!"
    End If
End Sub

Function checkIsahTestQuery() As Boolean
    'input for T_ProductionHeader: connect to db and execute query
    Dim sql1 As String, sqlconn As ADODB.Connection, rs0 As ADODB.Recordset
    Debug.Print "Try to connect with string: " + main.getISAHconnstr()
    
    On Error GoTo connection_error:
        Set sqlconn = main.getISAHconnection()
        ' connection management: make sure to close connections on error
    On Error GoTo close_connection
        sql1 = "SELECT 1"
        Set rs0 = db.queryDB(sqlconn, sql1)
        If True Then
           db.printRecordset rs0, False, False
        End If
        sqlconn.Close
    On Error GoTo 0
        GoTo no_error
  
connection_error:
    MsgBox Err.Description
    checkIsahTestQuery = False
    Exit Function
close_connection:
    sqlconn.Close
    MsgBox Err.Description
    checkIsahTestQuery = False
    Exit Function
no_error:
    checkIsahTestQuery = True
End Function

' Function to get ISAH database name based database/connection string dropdown
Function getISAHdbname() As String
    getISAHdbname = ThisWorkbook.Worksheets(main.CONTROL_SHEET_NAME).Range(main.SELECTED_DATABASE_NAME_ADDR).value
End Function

Function getISAHProfileName() As String
    getISAHProfileName = ThisWorkbook.Worksheets(main.CONTROL_SHEET_NAME).Range(main.DATABASE_DROPDOWN_ADDR).value
End Function

' function to check if the profile is a home profile
Function ISAHProfileIsHome(profile As String) As Boolean
    ISAHProfileIsHome = a.ItemInArray(profile, Array("JKR", "JKR2"))
End Function

' Function to get ISAH production header based on ISAHProfilename
Function getISAHprodheader() As String
    Dim dbname As String, table_name As String, full_table_name As String, dbprofile As String
    dbname = main.getISAHdbname()
    full_table_name = "[@1].[dbo].[@2]"
    dbprofile = main.getISAHProfileName()
    If ISAHProfileIsHome(dbprofile) Then
      table_name = "T_ProductionHeader_TEST"
    ElseIf dbname = "NewMultifill" Or dbname = "Testmultifill" Then
      table_name = "T_ProductionHeader"
    Else
      Err.Raise 1001, "getISAHprodheader", "Expecting dbname either TestMultifill or NewMultifill, not " & dbname
    End If
    
    getISAHprodheader = str.subInStr(full_table_name, dbname, table_name)
End Function

' Function to get ISAH production bill of operations based on ISAHProfilename
Function getISAHprodboo() As String
    Dim dbname As String, table_name As String, full_table_name As String, dbprofile As String
    dbname = main.getISAHdbname()
    full_table_name = "[@1].[dbo].[@2]"
    dbprofile = main.getISAHProfileName()
    If ISAHProfileIsHome(dbprofile) Then
      table_name = "T_ProdBillOfOper_TEST"
    ElseIf dbname = "NewMultifill" Or dbname = "Testmultifill" Then
      table_name = "T_ProdBillOfOper"
    Else
      Err.Raise 1001, "getISAHprodboo", "Expecting dbname either Testmultifill or NewMultifill, not " & dbname
    End If
    
    getISAHprodboo = str.subInStr(full_table_name, dbname, table_name)
End Function

' Function to get ISAH production bill of materials based on ISAHProfilename
Function getISAHprodbom() As String
    Dim dbname As String, table_name As String, full_table_name As String, dbprofile As String
    dbname = main.getISAHdbname()
    full_table_name = "[@1].[dbo].[@2]"
    dbprofile = main.getISAHProfileName()
    If ISAHProfileIsHome(dbprofile) Then
      table_name = "T_ProdBillOfMat_TEST"
    ElseIf dbname = "NewMultifill" Or dbname = "Testmultifill" Then
      table_name = "T_ProdBillOfMat"
    Else
      Err.Raise 1001, "getISAHprodbom", "Expecting dbname either Testmultifill or NewMultifill, not " & dbname
    End If
    
    getISAHprodbom = str.subInStr(full_table_name, dbname, table_name)
End Function

' Function to get ISAH T_Part based on ISAHProfilename
Function getISAHpart() As String
    Dim dbname As String, table_name As String, full_table_name As String, dbprofile As String
    dbname = main.getISAHdbname()
    full_table_name = "[@1].[dbo].[@2]"
    dbprofile = main.getISAHProfileName()
    If ISAHProfileIsHome(dbprofile) Then
      table_name = "T_Part_BasicMat"
    ElseIf dbname = "NewMultifill" Or dbname = "Testmultifill" Then
      table_name = "T_Part"
    Else
      Err.Raise 1001, "getISAHprodboo", "Expecting dbname either Testmultifill or NewMultifill, not " & dbname
    End If
    
    getISAHpart = str.subInStr(full_table_name, dbname, table_name)
End Function

' Function to get ISAH connection string based on VDMI test mode
Function getISAHconnstr() As String
    getISAHconnstr = ThisWorkbook.Worksheets(main.CONTROL_SHEET_NAME).Range(main.SELECTED_CONNECTION_STRING_ADDR).value
End Function

Sub init_named_ranges()
   Dim ws0 As Worksheet
   ' initialize global named ranges
   ' isah staging sheet
   Set ws0 = w.get_or_create_worksheet(main.ISAH_STAGING_SHEET_NAME, ThisWorkbook)
   r.create_named_range main.ISAH_STAGING_RANGE_NAME, ws0.name, "A1", clear:=True

   ' number of pallets per article sheet
   Set ws0 = w.get_or_create_worksheet(main.NUMBER_PER_PALLET_SHEET_NAME, ThisWorkbook)
   r.create_named_range main.NUMBER_PER_PALLET_NAMED_RANGE, ws0.name, "A1", clear:=False
   r.expandNamedRange main.NUMBER_PER_PALLET_NAMED_RANGE, ThisWorkbook
End Sub

' This subroutine processes each cell in the provided range capgrp_column_range.
' It trims the value of each cell and sets the trimmed value back to the cell.
Sub handle_input_capgrp(capgrp_column_range As Range)
    Dim cell As Range
    Dim trimmedValue As String

    ' Loop through each cell in the range, starting from the second row
    For Each cell In capgrp_column_range
        ' Check if the cell value is a string
        If VarType(cell.value) = vbString Then
            ' Trim the value and set it back to the cell
            trimmedValue = Trim(cell.value)
            cell.value = trimmedValue
        End If
    Next cell
End Sub

Sub isah_export_stage_orders()
    Dim rng0 As Range
    Dim wb0 As Workbook
    Set wb0 = ThisWorkbook

    ' append all capgrp order ranges and paste columns `ISAH_STAGING_COLUMNS` to sheet `ISAH_STAGING_SHEET_NAME`
    ' first get named ranges of orders
    Set capgrp_sheets = main.get_capgrp_sheet_names()
    Dim r0 As Integer, N As Long, orders_arr_all As Variant, ws0 As Worksheet
    Dim columns_to_select As Variant, orderNrCapgrp As String, weekNr As Integer, capgrp As String
      
    'parameters
    columns_to_select = Split(main.ISAH_STAGING_COLUMNS, ";")
    isah_database_columns = Split(main.ISAH_STAGING_COLUMNS_DB_NAMES, ";")
      
    ' sheets: ISAH staging and Template
    Set ws0 = wb0.Worksheets(main.ISAH_STAGING_SHEET_NAME)
    w.clearWorksheet ws0, wb0
    arrForProductieOrder = main.get_isah_input_range()
    
    ' append all capgrp order ranges to array `orders_arr_all`
    N = 1
    For Each c In capgrp_sheets
        capgrp = c
        orders_arr = main.get_orders_range(capgrp)
        r0 = a.numArrayRows(orders_arr) 'number of rows of current capgrp
        If (r0 < 2) Then
            GoTo next_capgrp_sheet
        End If
        
        ' subset columns and set `isah_database_columns` as header
        orders_arr = a.select_array_columns(orders_arr, columns_to_select) '=> TODO FIX ?
        orders_arr = a.setArrayHeader(orders_arr, isah_database_columns)
        
        ' filter out where column ProdHeaderOrdNr is not series of digits
        
        ' if CAPGRP = "INPAK" then get the right Cap.Grp
        If capgrp = "INPAK" Then
           If a.numArrayRows(orders_arr) > 0 Then
           
              orders_arr = a.AppendColumn(orders_arr, "", main.ISAH_STAGING_CAGGRP_COLUMN)
              cl_index = a.FindArrayColumnIndex(orders_arr, "ProdHeaderOrdNr")
              
              For Each rw_index In a.getRowIndexes(orders_arr)
                  If rw_index <= 1 Then
                     GoTo nx_i
                  End If
                  orderNr = Trim(orders_arr(rw_index, cl_index))
                  
                  'For INPAK recover the original Cap.Grp => v20240301
                  If main.P_DEBUG Then
                    Debug.Print "finding capgrp of Productiorder " & CStr(orderNr)
                  End If
                  
                  ' if orderNr is not defined then skip (for example Ombouwregel)
                  If CStr(orderNr) = "" Then
                    GoTo nx_i
                  End If
                  
                  weekNr = main.get_weeknumber(capgrp)
                  orders_arr_filtered = a.QueryArray(arrForProductieOrder, "Productieorder", CStr(orderNr), "ProdWk", weekNr)
                  orderNrCapgrp = Trim(a.getNamedArrayValue(orders_arr_filtered, "Cap.Grp"))
                  orders_arr(rw_index, UBound(orders_arr, 2)) = orderNrCapgrp
nx_i:
              Next rw_index
           End If
        Else
           ' add column CAPGRP = `capgrp` to orders_arr
           capgrp_column_arr = a.create_vector(r0, capgrp, header_value:=main.ISAH_STAGING_CAGGRP_COLUMN, as_2darray:=True)
           orders_arr = a.AppendColumn(orders_arr, capgrp_column_arr)
        End If
         
        ' append to `orders_arr_all`
        If N = 1 Then
           orders_arr_all = orders_arr
        Else
           orders_arr_values = a.resize_array(orders_arr, r0:=2)
           orders_arr_all = a.concatArrays(orders_arr_all, orders_arr_values)
        End If
        N = N + 1
next_capgrp_sheet:
    Next
    
    If a.numArrayRows(orders_arr_all) <= 1 Then
       MsgBox "No orders on sheet"
       Exit Sub
    End If
    
    ' in result array `orders_arr_all`, filter out all rows where ProdHeaderOrdNr does not match pattern ^\\d (starts with digit)
    OrdersHeaderArray = a.getArrayRow(orders_arr_all, 1)
    column_index = CInt(a.getArrayColumnIndex(orders_arr_all, "ProdHeaderOrdNr"))
    OrdersNoHeader = a.subset_rows(orders_arr_all, 2)
    OrdersWithOrdNrArray = a.FilterArrayOnPattern(OrdersNoHeader, "^\d", column_index)
    OrdersWithOrdNrArray = a.concatArrays(OrdersHeaderArray, OrdersWithOrdNrArray)
    
    ' paste array `orders_arr_all` and create named range `main.ISAH_STAGING_RANGE_NAME`
    a.pasteArray OrdersWithOrdNrArray, "A1", ws0
    r1 = a.numArrayRows(OrdersWithOrdNrArray)
    c1 = a.numArrayColumns(OrdersWithOrdNrArray)
    Set rng0 = ws0.Range(r.get_range_address(ws0, 1, r1, 1, c1))
    r.create_named_range main.ISAH_STAGING_RANGE_NAME, ws0.name, "A1", clear:=False
    r.expandNamedRange main.ISAH_STAGING_RANGE_NAME, ThisWorkbook, dbg:=True
    Set rng0 = r.get_range(main.ISAH_STAGING_RANGE_NAME, ws0, wb:=wb0)

    ' add extra columns to `ordersRange`:
    Dim ordersRangeName As String: ordersRangeName = main.ISAH_STAGING_RANGE_NAME
    Dim IsahColumns As collection
    Set IsahColumns = str.stringToCol(main.ISAH_STAGING_UPDATE_COLUMNS, ";")
    For Each colName In IsahColumns:
        r.add_named_range_column ordersRangeName, CStr(colName)
    Next
    
    'set worksheet column formats
    w.formatWorksheetColumns ws0, main.ISAH_STAGING_UPDATE_COLUMNS_FORMATS
    
    'autofit columns
    ws0.columns.AutoFit
    
    Exit Sub
    
End Sub

Sub isah_export_match_prodheader()
    Dim wb0 As Workbook: Set wb0 = ThisWorkbook
    
    'parameters
    Dim ordernr_column As String, capgrp_column As String
    ordernr_db_column = main.ISAH_DATABASE_ORDERNR_COLUMN
    capgrp_column = main.ISAH_STAGING_CAGGRP_COLUMN
    capgrp_db_column = main.ISAH_DATABASE_CAGGRP_COLUMN
    dossiercode_column = main.ISAH_DATABASE_DOSSIERCODE_COLUMN
    
    Dim ordersRange As Range: Set ordersRange = r.get_range(main.ISAH_STAGING_RANGE_NAME, wb:=wb0)
    'orders_to_query = r.get_column_values(ordersRange, main.ISAH_STAGING_ORDERNR_INDEX, wb:=wb0)
    orders_to_query = a.getArrayColumnValues(ordersRange.value, main.ISAH_STAGING_ORDERNR_INDEX)
    
    ' T_ProductionHeader query
    Dim sql0 As String, table_name As String, sql_template As String
    table_name = main.getISAHprodheader()
    where_statement = db.sqlWhereInCondition(orders_to_query, main.ISAH_DATABASE_ORDERNR_COLUMN, mssql)
    sql_template = "SELECT LTRIM(RTRIM(@1)) as @2, @3, @4 FROM @5 @6;"
    sql0 = str.subInStr(sql_template, ordernr_db_column, ordernr_db_column, main.ISAH_DATABASE_DATE_COLUMNS, dossiercode_column, table_name, where_statement)
    
    'input for T_ProductionHeader: connect to db and execute query
    Dim sqlconn As ADODB.Connection, rs0 As ADODB.Recordset
    Set sqlconn = main.getISAHconnection()
  
On Error GoTo close_connection
    Set rs0 = db.queryDB(sqlconn, sql0)
    'db.printRecordset rs0, False, False
    result_array = db.RecordSetToArray(rs0)
    sqlconn.Close
On Error GoTo 0
    GoTo no_error
    
close_connection:
    If main.P_DEBUG Then
       Debug.Print "Error querying: ", sql0
    End If
    sqlconn.Close
    Err.Raise Err
    
no_error:
    ' fill range using prodheader_array
    ' from ordersRange.orderNrColumn, find associated row in prodheader_array and get values: `match_prod_header`=0/1, StartDate_header, EndDate_header
    Dim orderNrColumn As Range, matchColumn As Range, StartDateColumn As Range, EndDateColumn As Range, CapGrpColumn As Range, dossierCodeColumn As Range
    Set orderNrColumn = r.get_column_values(ordersRange, main.ISAH_STAGING_ORDERNR_INDEX, wb:=wb0)

    'loop over values in `orderNrColumn` and find values in `prodheader_array`
    Set matchColumn = r.get_column_values(ordersRange, "match_prod_header", wb:=wb0)
    Set StartDateColumn = r.get_column_values(ordersRange, "StartDate_header", wb:=wb0)
    Set EndDateColumn = r.get_column_values(ordersRange, "EndDate_header", wb:=wb0)
    Set dossierCodeColumn = r.get_column_values(ordersRange, dossiercode_column, wb:=wb0)
    
    Dim cl As Range, rw As Long
    Dim key_array As Variant, search_array As Variant
    
    rw = 1
    For Each cl In orderNrColumn.Cells
        order_value = CStr(cl.value)
        key_array = Array("ProdHeaderOrdNr")
        search_array = Array(CStr(order_value))
        matchRowIndex = a.MatchArrayRowIndex(result_array, key_array, search_array, dbg:=False)
        If matchRowIndex > 0 Then
            matchColumn.Cells(rw, 1) = 1
            StartDateColumn.Cells(rw, 1) = result_array(matchRowIndex, 2)
            EndDateColumn.Cells(rw, 1) = result_array(matchRowIndex, 3)
            dossierCodeColumn.Cells(rw, 1) = result_array(matchRowIndex, 4)
        Else
            matchColumn.Cells(rw, 1) = 0
        End If
        rw = rw + 1
    Next
    
    ' set matchColumn 0=Red, 1=Green
    r.setConditionalFormatting matchColumn, 1, 0
    
End Sub

Sub isah_export_update_prodboo_grp()
    ' This procedure updates the MachGrpCode in the BillOfOperations table for each order.
    ' If the line or resource has been changed, then this will be updated in ISAH.
    ' It sets the ProdBOOStatusCode to '20' to indicate the dossier requires a manual update.
    ' The procedure iterates over each order in a predefined range, constructs SQL update statements,
    ' and executes them to apply the changes to the database.
    '
    ' The procedure performs the following actions:
    ' - Retrieves table names for headers and BillOfOperations.
    ' - Constructs a query template to find ProdHeaderDossierCode.
    ' - Loops over each order in the staging area and builds update statements.
    ' - Updates the newStandCapacity and newMachPlanTime in the Excel staging area.
    ' - Converts values to a format compatible with the database.
    ' - Adds the constructed SQL update statements to a collection.
    ' - Opens a database connection and executes each SQL statement to update the records.
    ' - Closes the database connection upon completion or error.
    '
    ' Note: The procedure contains error handling to ensure the database connection is closed in case of an error.
    
    Dim sqlTemplate As String, sqlUpdate As String, newMachGrpCodeRes As String, newMachGrpCode As String, orderNr As String
    Dim newQty As Double, newStandCapacity As Double 'constrain to ISAH precision, scale
    Dim updateStandCapacity As String
    Dim tableBOO As String, tableHeader As String
    Dim wb0 As Workbook: Set wb0 = ThisWorkbook
    
    ' Get table names
    tableHeader = main.getISAHprodheader()
    tableBOO = main.getISAHprodboo()
    
    ' Get the staging orders range and array
    Dim ordersRange As Range: Set ordersRange = r.get_range(main.ISAH_STAGING_RANGE_NAME, wb:=wb0, offset_row:=0)
    ordersRangeArray = ordersRange
    
    ' Query template
    sqlSubQueryTemplate = "(SELECT DISTINCT ProdHeaderDossierCode FROM @1 where ProdHeaderOrdNr ='@2')" 'to find ProdHeaderDossierCode
    
    'Loop over each matched staging order and build update statments, collect update statements in collection `sqlStatements`
    Dim sqlStatements As New collection, ProdBOOLineNrs As collection, LineNr As Integer, matchInd As Integer, i As Integer
    ' Generate collection of LineNr 1,2 to loop over
    Set ProdBOOLineNrs = clls.getCollection(1, 2)
    For i = LBound(ordersRangeArray, 1) + 1 To UBound(ordersRangeArray, 1) 'start on the second row of ordersRangeArray
        ' get values from ordersRangeArray
        newMachGrpCode = a.getColumnValue(ordersRangeArray, i, main.ISAH_STAGING_CAGGRP_COLUMN)
        newMachGrpCodeRes = a.getColumnValue(ordersRangeArray, i, "MachGrpCode")
        dossierCode = a.getColumnValue(ordersRangeArray, i, main.ISAH_DATABASE_DOSSIERCODE_COLUMN)
        orderNr = a.getColumnValue(ordersRangeArray, i, main.ISAH_STAGING_ORDERNR_COLUMN)
        newQty = a.getColumnValue(ordersRangeArray, i, "Qty")
        calcDur = a.getColumnValue(ordersRangeArray, i, "Duur")
        
        If newQty <= 0 Then
           GoTo nx_i
        End If
        
        ' Calculate newStandCapacity from newQty / calcDur and new MachPlanTime
        newStandCapacity = Round(m.getRatio(newQty, calcDur, 0), 3)
        updateQty = Round(newQty, 3) 'make sure that Qty can be updated
        newMachPlanTime = Round(calcDur * 3600, 3) 'make sure that MachPlanTime can be updated
        
        ' update excel newStandCapacity column
        col_index = r.getColumnIndex(ordersRange, "next_StandCapacity_boo")
        ordersRange.Cells(i, col_index).value = newStandCapacity
        col_index = r.getColumnIndex(ordersRange, "next_MachPlanTime_boo")
        ordersRange.Cells(i, col_index).value = newMachPlanTime
        
        updateStandCapacity = Replace(CStr(newStandCapacity), ",", ".") ' FIX ME: handle dutch xlsx
        updateMachPlanTime = Replace(CStr(newMachPlanTime), ",", ".") ' FIX ME: handle dutch xlsx
        
        ' Debug.Print newQty, calcDur, newStandCapacity, updateStandCapacity, updateMachPlanTime
        
        ' Parameterize sub query
        sqlSubQuery = str.subInStr(sqlSubQueryTemplate, tableHeader, orderNr)

        For Each ProdBOOLineNr In ProdBOOLineNrs
            LineNr = ProdBOOLineNr
            ' update BOO template
            sqlTemplate = "UPDATE @1 SET MachGrpCode = '@2', " + _
            "Qty = @3, StandCapacity = @4, ProdBOOStatusCode = '@7', MachPlanTime =  @8" + _
            "WHERE ProdHeaderDossierCode=@5 AND ProdBOOLineNr=@6"
            ' update Line 1
            If LineNr = 1 Then
                sqlUpdate = str.subInStr(sqlTemplate, tableBOO, newMachGrpCode, updateQty, updateStandCapacity, sqlSubQuery, ProdBOOLineNr, _
                main.ISAH_MANUAL_UPDATE_PRODBOOSTATUSCODE, updateMachPlanTime)
                sqlStatements.Add sqlUpdate
            ' update Line 2 where the Resource code <> ""
            ElseIf LineNr = 2 And newMachGrpCodeRes <> "" Then
                sqlUpdate = str.subInStr(sqlTemplate, tableBOO, newMachGrpCodeRes, updateQty, updateStandCapacity, sqlSubQuery, ProdBOOLineNr, _
                main.ISAH_MANUAL_UPDATE_PRODBOOSTATUSCODE, updateMachPlanTime)
                sqlStatements.Add sqlUpdate
            End If
        Next
        GoTo nx_i
nx_i:
    Next i

    ' connect to database and execute sqlStatements
    Dim conn As ADODB.Connection
    Set conn = main.getISAHconnection()
    
    c = 0
    ' connection management: make sure to close connections on error
    On Error GoTo close_connection
    ' transaction management
    conn.BeginTrans
    
    On Error GoTo TransactionError ' Set up error handling
    
    For Each stat In sqlStatements
       conn.Execute CStr(stat)
       c = c + 1
    Next
    conn.CommitTrans
    On Error GoTo 0
    GoTo no_error

'errorhandler
TransactionError:
    conn.RollbackTrans
    If main.P_DEBUG Then
       Debug.Print "Transaction failed, updates have been rolled back"
    End If
     
close_connection:
    If main.P_DEBUG Then
       Debug.Print "Database error executing: " & stat
    End If
    conn.Close
    Err.Raise Err
    
no_error:
    conn.Close
    
End Sub

Sub isah_export_match_prodboo()
    ' 20240120: return StandCapacity from ISAH
    Dim wb0 As Workbook: Set wb0 = ThisWorkbook
    
    'parameters
    Dim ordernr_column As String, capgrp_column As String
    ordernr_db_column = main.ISAH_DATABASE_ORDERNR_COLUMN
    capgrp_column = main.ISAH_STAGING_CAGGRP_COLUMN
    capgrp_db_column = main.ISAH_DATABASE_CAGGRP_COLUMN
    dossiercode_column = main.ISAH_DATABASE_DOSSIERCODE_COLUMN
    
    ' get the ProdHeaderDossierCode
    Dim ordersRange As Range: Set ordersRange = r.get_range(main.ISAH_STAGING_RANGE_NAME, wb:=wb0)
    'dossiercode_to_query = r.get_column_values(ordersRange, dossiercode_column, wb:=wb0)
    'capgrp_to_query = r.get_column_values(ordersRange, main.ISAH_STAGING_CAGGRP_COLUMN, wb:=wb0)
    
    dossiercode_to_query = a.getArrayColumnValues(ordersRange.value, dossiercode_column)
    capgrp_to_query = a.getArrayColumnValues(ordersRange.value, main.ISAH_STAGING_CAGGRP_COLUMN)
    
    ' build match query for T_ProdBillOfOper  using CAPGRP and DossierCode as key,
    ' return ISAH columns ProdHeaderDossierCode, MachGrpCode, StartDate, EndDate
    Dim sql1 As String, or_statement As String, OR_TEMPLATE As String, dossiercode_value As String, capgrp_value As String
    OR_TEMPLATE = "(@1='@2' AND @3='@4')"
    table_name = main.getISAHprodboo()
    
    Dim where_parts As New collection
    For i = LBound(dossiercode_to_query, 1) To UBound(dossiercode_to_query, 1)
        dossiercode_value = CStr(dossiercode_to_query(i, 1))
        If dossiercode_value <> "" Then
            capgrp_value = capgrp_to_query(i, 1)
            or_statement = str.subInStr(OR_TEMPLATE, dossiercode_column, dossiercode_value, capgrp_db_column, capgrp_value)
            where_parts.Add or_statement
        Else
            where_parts.Add "1=0"
        End If
    Next
    where_statement = clls.collectionToString(where_parts, " OR ")
    sql_template = "SELECT LTRIM(RTRIM(@1)) as @1, LTRIM(RTRIM(@2)) as @2, @3, StandCapacity FROM @4 WHERE @5;"
    sql1 = str.subInStr(sql_template, dossiercode_column, capgrp_db_column, main.ISAH_DATABASE_DATE_COLUMNS, table_name, where_statement)
    
    If P_DEBUG Then
       Debug.Print sql1
    End If
    
    ' connect to db and query matching records in T_ProdBillOfOper
    Dim sqlconn As ADODB.Connection, rs0 As ADODB.Recordset
    Set sqlconn = main.getISAHconnection()

    ' connection management: make sure to close connections on error
On Error GoTo close_connection
    Set rs0 = db.queryDB(sqlconn, sql1)
    If main.P_DEBUG Then
       db.printRecordset rs0, False, False
    End If
    prodboo_array = db.RecordSetToArray(rs0)
    If main.P_DEBUG Then
       a.printArray prodboo_array
    End If
On Error GoTo 0
    GoTo no_error
    
close_connection:
    sqlconn.Close
    Err.Raise Err
    
no_error:
    ' fill range using prodboo_array
    Dim dossierCodeColumn As Range, matchColumn As Range, StartDateColumn As Range, EndDateColumn As Range, CapGrpColumn As Range
    Set dossierCodeColumn = r.get_column_values(ordersRange, dossiercode_column, wb:=wb0)
    Set CapGrpColumn = r.get_column_values(ordersRange, main.ISAH_STAGING_CAGGRP_COLUMN, wb:=wb0)
    
    'loop over values in `orderNrColumn` and find values in `prodheader_array`
    Set matchColumn = r.get_column_values(ordersRange, "match_prod_boo", wb:=wb0)
    Set StartDateColumn = r.get_column_values(ordersRange, "StartDate_boo", wb:=wb0)
    Set EndDateColumn = r.get_column_values(ordersRange, "EndDate_boo", wb:=wb0)
    
    rw = 1
    For Each cl In dossierCodeColumn.Cells
        dossiercode_value = dossierCodeColumn.Cells(rw).value
        If dossiercode_value = "" Then
        GoTo nx_row
        End If
        capgrp_value = CapGrpColumn.Cells(rw).value
        key_array = Array(dossiercode_column, capgrp_db_column)
        search_array = Array(CStr(dossiercode_value), CStr(capgrp_value))
        matchRowIndex = a.MatchArrayRowIndex(prodboo_array, key_array, search_array, dbg:=False)
        If matchRowIndex > 0 Then
            matchColumn.Cells(rw, 1) = 1
            StartDateColumn.Cells(rw, 1) = prodboo_array(matchRowIndex, 3)
            EndDateColumn.Cells(rw, 1) = prodboo_array(matchRowIndex, 4)
        Else
            matchColumn.Cells(rw, 1) = 0
        End If
nx_row:
        rw = rw + 1
    Next
    
    ' set matchColumn 0=Red, 1=Green
    r.setConditionalFormatting matchColumn, 1, 0
End Sub

Sub isah_export_update_prodheader()
    Dim wbconn As ADODB.Connection, rs0 As ADODB.Recordset, sqlconn As ADODB.Connection
    Dim wb0 As Workbook
    Dim sql_template As String, sql0 As String, sql1 As String, table_name As String, update_statements As String, table_name_isah As String
    Dim set_columns As String, key_columns As String
    Dim wherecondition As String, orderNumbers As Variant
    Dim t_prodheader_array As Variant, ordersRange As Range
    Dim table_suffix As String
    
    'Parameters
    table_name = main.ISAH_STAGING_SHEET_NAME & "$"
    select_column_list = "ProdHeaderOrdNr, StartDate, EndDate "
    where_condition = "match_prod_header=1"
    set_columns = "StartDate;EndDate"
    key_columns = "ProdHeaderOrdNr"
    table_name_isah = main.getISAHprodheader()
    table_suffix = "_header"
    select_check_columns_list = "LTRIM(RTRIM(ProdHeaderOrdNr)) as ProdHeaderOrdNr, StartDate AS next_StartDate_header, EndDate AS next_EndDate_header"
    
    ' Get the global connection or connect to this workbook
    Set wbconn = main.getWorkbookConnection()
    
    ' Query from table "EXPORT_ISAH" where `match_prod_header`=1
    sql_template = "SELECT @1 FROM [@2] WHERE @3;"
    sql0 = str.subInStr(sql_template, select_column_list, table_name, where_condition)
    Set rs0 = db.queryDB(wbconn, sql0)
    match_orders_array = db.RecordSetToArray(rs0)
    
    ' Check orders to match, if array is empty then exit sub
    If a.numArrayRows(match_orders_array) <= 1 Then
       Debug.Print "No EXPORT_ISAH records with match_prod_header=1 found"
       ' Close connection to this workbook
       'wbconn.Close
       Exit Sub
    End If
    
    ' Create SQL update statement using `ProdHeaderOrdNr` as key column and `StartDate`, `EndDate` as update columns
    'update_statements = db.sqlUpdateStatement(rs0, table_name_isah, set_columns, key_columns, mssql)
    update_statements = db.sqlUpdateCaseStatement(rs0, table_name_isah, set_columns, key_columns, mssql)
    If main.P_DEBUG Then
       Debug.Print update_statements
    End If
    
    ' Create query statement `sql1` to select columns from `rs0` where `match_prod_header`=1
    orderNumbers = a.resize_array(a.getArrayColumn(match_orders_array, 1), r0:=2)
    wherecondition = db.sqlWhereInCondition(orderNumbers, main.ISAH_DATABASE_ORDERNR_COLUMN, mssql)
    sql1 = str.subInStr("SELECT @1 FROM @2 @3;", select_check_columns_list, table_name_isah, wherecondition)
    If main.P_DEBUG Then
       Debug.Print sql1
    End If
    
'On Error GoTo close_connection
    ' Open connection to ISAH database
    Set sqlconn = main.getISAHconnection()
    
    ' Begin transaction
    sqlconn.BeginTrans
    
    ' Execute SQL update statements
    db.executeSqlStatements sqlconn, update_statements, db.MSSQL_LINE_BREAK, False, False
    
    ' Commit transaction
    sqlconn.CommitTrans
    
    ' Query from `sqlconn` and store result as array `result_array`
    result_array = db.RecordSetToArray(db.queryDB(sqlconn, sql1))
    
    ' Close connection to ISAH database
    sqlconn.Close
    
GoTo no_error
    
close_connection:
    ' Rollback transaction in case of error
    sqlconn.RollbackTrans
    
    ' Close connection to ISAH database
    sqlconn.Close
    Err.Raise Err
    
no_error:
    ' fill range using prodheader_array
    ' from ordersRange.orderNrColumn, find associated row in prodheader_array and get values: `match_prod_header`=0/1, StartDate_header, EndDate_header
    Set ordersRange = r.get_range(main.ISAH_STAGING_RANGE_NAME, wb:=wb0)
    Dim orderNrColumn As Range, StartDateColumn As Range, EndDateColumn As Range, NextStartDateColumn As Range, NextEndDateColumn As Range  ', capGrpColumn As Range, DossierCodeColumn As Range
    Dim CheckDatesColumn As Range
    Set orderNrColumn = r.get_column_values(ordersRange, main.ISAH_STAGING_ORDERNR_INDEX, wb:=wb0)

    'loop over values in `orderNrColumn` and find values in `prodheader_array`
    Set StartDateColumn = r.get_column_values(ordersRange, "StartDate", wb:=wb0)
    Set EndDateColumn = r.get_column_values(ordersRange, "EndDate", wb:=wb0)
    Set NextStartDateColumn = r.get_column_values(ordersRange, "next_StartDate" & table_suffix, wb:=wb0)
    Set NextEndDateColumn = r.get_column_values(ordersRange, "next_EndDate" & table_suffix, wb:=wb0)
    Set CheckDatesColumn = r.get_column_values(ordersRange, "check_dates" & table_suffix, wb:=wb0)
    
    Dim cl As Range, rw As Long
    Dim key_array As Variant, search_array As Variant
    
    rw = 1
    For Each cl In orderNrColumn.Cells
        order_value = CStr(cl.value)
        key_array = Array("ProdHeaderOrdNr")
        search_array = Array(CStr(order_value))
        matchRowIndex = a.MatchArrayRowIndex(result_array, key_array, search_array, dbg:=False)
        If matchRowIndex > 0 Then
            NextStartDateColumn.Cells(rw, 1) = result_array(matchRowIndex, 2)
            NextEndDateColumn.Cells(rw, 1) = result_array(matchRowIndex, 3)
            If StartDateColumn.Cells(rw, 1) <> NextStartDateColumn.Cells(rw, 1) And StartDateColumn.Cells(rw, 1) <> NextEndDateColumn.Cells(rw, 1) Then
               check_value = 0
            Else
               check_value = 1
            End If
            CheckDatesColumn.Cells(rw, 1) = check_value
        End If
nx_row:
    rw = rw + 1
    Next

    ' set matchColumn 0=Red, 1=Green
    r.setConditionalFormatting CheckDatesColumn, 1, 0
End Sub

Sub isah_export_update_prodboo()
    ' Update the ISAH BillOfProd (BOO) table:
    ' 1. Establishing a connection to the Excel workbook and querying data from `main.ISAH_STAGING_SHEET_NAME`.
    ' 2. Checking if the queried records are have condition match_prod_boo=1 and exiting if not.
    ' 3. Constructing SQL update statements based on the queried data.
    ' 4. Opening a connection to the ISAH database and executing the update statements.
    ' 5. Querying the ISAH database to check the updates and storing the results in an array.
    ' 6. Updating the Excel workbook with the results from the ISAH database.
    ' 7. Applying conditional formatting to the workbook based on the update results.
    '
    ' The subroutine handles errors by closing the database connection and re-raising the error.
    ' It also includes clean-up code to ensure that all connections are closed properly.
    
    Dim wbconn As ADODB.Connection, rs0 As ADODB.Recordset, sqlconn As ADODB.Connection
    Dim wb0 As Workbook
    Dim sql_template As String, sql0 As String, sql1 As String, table_name As String, update_statements As String, table_name_isah As String
    Dim set_columns As String, key_columns As String
    Dim wherecondition As String, dossierCodes As Variant, machgrpCodes As Variant
    Dim t_prodheader_array As Variant, ordersRange As Range
    Dim table_suffix As String
    
    ' Parameters from staging sheet
    table_name = main.ISAH_STAGING_SHEET_NAME & "$"
    select_column_list = "ProdHeaderDossierCode, StartDate, EndDate, ROUND(cdbl(TimeValue(StartDate))*24*3600,2) AS StartTime"
    where_condition = "match_prod_boo=1"
    set_columns = "StartDate;EndDate;StartTime" '20231218: update ISAH StartDate, EndDate and StartTime in seconds
    key_columns = "ProdHeaderDossierCode" '20231218: match update on ProdHeaderDossierCode
    table_name_isah = main.getISAHprodboo()
    table_suffix = "_boo"
    
    ' Get the global connection or connect to this workbook
    Set wbconn = main.getWorkbookConnection()
    
    ' Query from sheet "EXPORT_ISAH" where `match_prod_boo`=1
    sql_template = "SELECT @1 FROM [@2] WHERE @3;"
    sql0 = str.subInStr(sql_template, select_column_list, table_name, where_condition)
    Set rs0 = db.queryDB(wbconn, sql0)
    
    match_dossiercode_capgrp_array = db.RecordSetToArray(rs0)
    
    ' Check dossiercodes, capgrp to match, if array is empty then exit sub
    If a.numArrayRows(match_dossiercode_capgrp_array) <= 1 Then
        Debug.Print "No EXPORT_ISAH records with match_prod_boo=1 found"
        Exit Sub
    End If
    
    ' Create SQL update statement using `ProdHeaderOrdNr` as key column and `StartDate`, `EndDate` as update columns
    'update_statements = db.sqlUpdateStatement(rs0, table_name_isah, set_columns, key_columns, mssql, force_string:=True)
    update_statements = db.sqlUpdateCaseStatement(rs0, table_name_isah, set_columns, key_columns, mssql, force_string:=True)
    If main.P_DEBUG Then
       Debug.Print update_statements
    End If
    
    ' Create query statement `sql_check_columns` to select columns from `rs0` where `match_prod_header`=1
    Dim sql_check_columns As String
    select_check_columns_list = "LTRIM(RTRIM(ProdHeaderDossierCode)) as ProdHeaderDossierCode," + _
    "MIN(StartDate) OVER (PARTITION BY ProdHeaderDossierCode) AS next_StartDate_boo, " + _
    "MIN(EndDate) OVER (PARTITION BY ProdHeaderDossierCode) AS next_EndDate_boo, " + _
    "MIN(CAST(FORMAT(CAST((StartTime+1)/86400.000 AS datetime), 'HH:mm') AS varchar)) OVER (PARTITION BY ProdHeaderDossierCode) AS next_StartTime_boo, " + _
    "MIN(ProdBOOStatusCode) OVER (PARTITION BY ProdHeaderDossierCode) AS ProdBOOStatusCode"
    
    ' Key matching values for update statement
    dossierCodes = a.resize_array(a.getArrayColumn(match_dossiercode_capgrp_array, 1), r0:=2)
    machgrpCodes = a.resize_array(a.getArrayColumn(match_dossiercode_capgrp_array, 2), r0:=2)
    
    Dim where_values_col As New collection, wherepart As String
    Dim dossiercode_value As String, machgrpcode_value As String
    sql_template = "(@1=@2)"
    For i = LBound(dossierCodes) To UBound(dossierCodes)
        dossiercode_column = main.ISAH_DATABASE_DOSSIERCODE_COLUMN
        machgrpcode_column = main.ISAH_DATABASE_CAGGRP_COLUMN
        dossiercode_value = db.xlToDBvalue(CStr(dossierCodes(i, 1)), mssql, force_string:=True)
        wherepart = str.subInStr(sql_template, dossiercode_column, dossiercode_value)
        where_values_col.Add wherepart
    Next i
    
    wherecondition = clls.collectionToString(where_values_col, " OR ")
    sql_check_columns = str.subInStr("SELECT DISTINCT @1 FROM @2 WHERE @3;", select_check_columns_list, table_name_isah, wherecondition)
    
On Error GoTo close_connection
    ' Open connection to ISAH database
    Set sqlconn = main.getISAHconnection()
    
    ' Begin transaction
    sqlconn.BeginTrans
    
    ' Execute SQL update statements
    db.executeSqlStatements sqlconn, update_statements, db.MSSQL_LINE_BREAK, False, False
    
    ' Commit transaction
    sqlconn.CommitTrans
    
    ' Query `sql_check_columns` from `sqlconn` and store result as array `ProdBOOArray`
    ProdBOOArray = db.RecordSetToArray(db.queryDB(sqlconn, sql_check_columns))
    If main.P_DEBUG Then
       a.printArray ProdBOOArray
    End If
    
    ' Close connection to ISAH database
    sqlconn.Close
    
    On Error GoTo 0
    GoTo no_error
    
close_connection:
    ' Rollback transaction in case of error
    sqlconn.RollbackTrans
    
    ' Close connection to ISAH database
    sqlconn.Close
    Err.Raise Err

no_error:
    ' fill range using ProdBOOArray
    ' from ordersRange.orderNrColumn, find associated row in prodheader_array and get values: `match_prod_header`=0/1, StartDate_header, EndDate_header
    Dim dossierCodeColumn As Range, StartDateColumn As Range, EndDateColumn As Range, NextStartDateColumn As Range, NextEndDateColumn As Range, _
    CheckDatesColumn As Range, CapGrpColumn As Range, NextStartTimeColumn As Range, CheckProdBOOStatusCodeColumn As Range
    
    Set ordersRange = r.get_range(main.ISAH_STAGING_RANGE_NAME, wb:=wb0)
    Set dossierCodeColumn = r.get_column_values(ordersRange, main.ISAH_DATABASE_DOSSIERCODE_COLUMN, wb:=wb0)
    Set CapGrpColumn = r.get_column_values(ordersRange, main.ISAH_STAGING_CAGGRP_COLUMN, wb:=wb0)
    
    'loop over values in `orderNrColumn` and find values in `prodheader_array`
    Set StartDateColumn = r.get_column_values(ordersRange, "StartDate", wb:=wb0)
    Set EndDateColumn = r.get_column_values(ordersRange, "EndDate", wb:=wb0)
    Set NextStartDateColumn = r.get_column_values(ordersRange, "next_StartDate" & table_suffix, wb:=wb0)
    Set NextEndDateColumn = r.get_column_values(ordersRange, "next_EndDate" & table_suffix, wb:=wb0)
    Set NextStartTimeColumn = r.get_column_values(ordersRange, "next_StartTime" & table_suffix, wb:=wb0)
    Set CheckDatesColumn = r.get_column_values(ordersRange, "check_dates" & table_suffix, wb:=wb0)
    Set CheckProdBOOStatusCodeColumn = r.get_column_values(ordersRange, "check_ProdBOOStatusCode", wb:=wb0)
    
    Dim cl As Range, rw As Long
    Dim key_array As Variant, search_array As Variant
    
    rw = 1
    For Each cl In dossierCodeColumn.Cells
        If cl.value = "" Then
            GoTo nx_row
        End If
        dossiercode_value = CStr(dossierCodeColumn.Cells(rw, 1))
        key_array = Array(main.ISAH_DATABASE_DOSSIERCODE_COLUMN) ', main.ISAH_DATABASE_CAGGRP_COLUMN)
        search_array = Array(dossiercode_value) ', machgrpcode_value)
        matchRowIndex = a.MatchArrayRowIndex(ProdBOOArray, key_array, search_array, dbg:=False)
        If matchRowIndex > 0 Then
            NextStartDateColumn.Cells(rw, 1) = ProdBOOArray(matchRowIndex, 2)
            NextEndDateColumn.Cells(rw, 1) = ProdBOOArray(matchRowIndex, 3)
            NextStartTimeColumn.Cells(rw, 1) = ProdBOOArray(matchRowIndex, 4)
            ' Check ProdBOO NextStartDateColumn matches input StartDateColumn
            If StartDateColumn.Cells(rw, 1) <> NextStartDateColumn.Cells(rw, 1) And EndDateColumn.Cells(rw, 1) <> NextEndDateColumn.Cells(rw, 1) Then
               check_value = 0
            Else
               check_value = 1
            End If
            CheckDatesColumn.Cells(rw, 1) = check_value
            
            ' Check ProdBOO ProdBOOStatusCode equals "20"
            CheckProdBOOStatusCodeColumn.Cells(rw, 1) = ProdBOOArray(matchRowIndex, 5)
        End If
nx_row:
    rw = rw + 1
    Next

    ' set matchColumn 0=Red, 1=Green
    r.setConditionalFormatting CheckDatesColumn, 1, 0
    ' conditional formatting for check ProdBOOStatusCode
    r.setConditionalFormatting CheckProdBOOStatusCodeColumn, "20", "10"
    
End Sub

Sub isah_export_update_prodbom()
    ' This procedure updates the RequiredDate in the ISAHprodbom table based on the EXPORT_ISAH sheet
    Dim source_table As String, target_table As String
    Dim rs0 As ADODB.Recordset
    Dim bom_updates_array As Variant
    Dim set_columns As String, key_columns As String
    Dim update_statements As String
    Dim updateStatements As New collection
    Dim sqlUpdate As Variant
    Dim wb0 As Workbook: Set wb0 = ThisWorkbook
    Dim sql0 As String
    
    ' Define the source and target tables
    source_table = "EXPORT_ISAH"
    target_table = main.getISAHprodbom()
    
    ' Query the distinct values of 'ProdHeaderDossierCode' and 'Startdate_header' from the source table
    Dim wbconn0 As New ADODB.Connection
    Set wbconn0 = main.getWorkbookConnection()
    sql0 = "SELECT DISTINCT Cstr(ProdHeaderDossierCode) AS ProdHeaderDossierCode, next_StartDate_header AS RequiredDate FROM [" & source_table & "$] WHERE ProdHeaderDossierCode IS NOT NULL AND Startdate_header IS NOT NULL"
    Set rs0 = db.queryDB(wbconn0, sql0)

    If db.RecordSetNumberRecords(rs0) = 0 Then
       Debug.Print "isah_export_update_prodbom: ISAH_EXPORT has no valid ProdHeaderDossierCode RequiredDate to update in ISAH"
       Exit Sub
    End If
    
    ' Define the columns for the SET and WHERE clauses of the update statement
    set_columns = "RequiredDate"
    key_columns = "ProdHeaderDossierCode"
    
    ' Construct the SQL update statements and add them to the collection
    ' update_statements = db.sqlUpdateStatement(rs0, target_table, set_columns, key_columns, mssql)
    update_statements = db.sqlUpdateCaseStatement(rs0, target_table, set_columns, key_columns, mssql)
    
    ' Initialize the collection to store SQL update statements
    Set updateStatements = str.stringToCol(update_statements, ";")
    
    ' Print the update statements
    For Each sqlUpdate In updateStatements
        If main.P_DEBUG Then
           Debug.Print sqlUpdate
        End If
    Next sqlUpdate
    
    ' clean up
    rs0.Close
    Set rs0 = Nothing
    
    ' Update ISAH table BOM
    Dim conn As ADODB.Connection
    
On Error GoTo close_connection
    Set conn = main.getISAHconnection()
    
    ' Begin transaction
    conn.BeginTrans
    
    For Each sqlUpdate In updateStatements
        db.executeSql conn, CStr(sqlUpdate)
    Next sqlUpdate
    
    ' Commit transaction
    conn.CommitTrans
    
On Error GoTo 0
GoTo no_error
  
close_connection:
    ' Rollback transaction in case of error
    conn.RollbackTrans
    ' Close connection to ISAH database
    If main.P_DEBUG Then
       Debug.Print "Database error: " & sqlUpdate
    End If
    conn.Close
    Set conn = Nothing
    Err.Raise Err
  
no_error:
    ' Clean up
    conn.Close
    Set conn = Nothing
    
End Sub

Sub isah_export_check_bom_dates()

    ' 1. get ProdBillOfMat checks from ISAH, write to sheet main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET
    Dim wbconn0 As ADODB.Connection, conn As ADODB.Connection, sql0 As String, prodHeaderDossierCodeValuesList As String
    
    Set wbconn0 = main.getWorkbookConnection()
    sql0 = str.subInStr("SELECT DISTINCT ProdHeaderDossierCode FROM [@1$]", main.ISAH_STAGING_SHEET_NAME)
    prodHeaderDossierCodeValues = a.toVector(db.RecordSetToArray(db.queryDB(wbconn0, sql0)))
    prodHeaderDossierCodeValuesList = db.toSqlList(prodHeaderDossierCodeValues, force_string:=True)
    
    sql0 = main_isah_queries.check_ProdBillOfMat(bom_table_name:=main.getISAHprodbom(), _
                                                 prodheader_dossier_code_list:=prodHeaderDossierCodeValuesList _
                                                )
    
    
    Set conn = main.getISAHconnection()
    db.writeQueryToSheet conn, sql0, main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET, main.CHECK_BOM_REQUIRED_DATE_COLUMNS_FORMATS, write_empty_records:=True
    conn.Close
    
    w.formatWorksheetColumns ThisWorkbook.Sheets(main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET), main.CHECK_BOM_REQUIRED_DATE_COLUMNS_FORMATS
    
End Sub


Sub isah_export_match_bom_dates()
    ' 2. join with ISAH_EXPORT with ISAH_CHECK_BOM_REQUIRED_DATE_SHEET, add check column and write to ISAH_MATCH_BOM_REQUIRED_DATE_SHEET
    Dim rng0 As Range
    
    ' check if table main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET has actually been filled
    Set rng0 = ThisWorkbook.Sheets(main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET).Cells(1, 1)
    Set rng0 = r.expand_range(rng0)

    If rng0.Rows.count <= 1 Then
       Debug.Print str.subInStr("Sheet not filled: @1", main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET)
       'Exit Sub
    End If
        rng0.Activate

GoTo check_range
    Dim wbconn0 As ADODB.Connection, sql0 As String, rs0 As ADODB.Recordset
    Set wbconn0 = main.getWorkbookConnection()
    sql0 = join_ISAH_EXPORT_CHECK_PROD_BOM()
    db.writeQueryToSheet wbconn0, sql0, main.ISAH_MATCH_BOM_REQUIRED_DATE_SHEET
    
check_range:
    ' append column `check_bom_required_date` to EXPORT_ISAH
    Dim checkRange As Range, IsahExportRange As Range, checkColumn As Range
    Set checkRange = r.expand_range("A1", ThisWorkbook.Sheets(main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET))
    
    If checkRange.Rows.count <= 1 Then
    Set checkColumn = r.get_column(checkRange, "max_bom_required_date")
    Else
    Set checkColumn = r.get_column(checkRange, "max_bom_required_date", offset_row:=1)
    End If
    checkColumnIndex = checkColumn.column
    
    ' Add check column to ISAH_EXPORT
    Set IsahExportRange = main.get_isah_export_range()
    r.AppendColumnToRange IsahExportRange, "check_bom_required_date" ', values
    Set IsahExportRange = main.get_isah_export_range()
    Set checkColumn = r.get_column(IsahExportRange, "check_bom_required_date")
    
    ' Match CHECK_PROD_BILL_OF_MAT.max_bom_required_date against ISAH_EXPORT.StartDate_header
    ProdHeaderDossierCodeColumnIndex = r.get_column(IsahExportRange, "ProdHeaderDossierCode").column
    StartDateColumnIndex = r.get_column(IsahExportRange, "next_StartDate_header").column

    ' Fill the formula formulaTemplateString in the newly added column
    Dim formulaDefinition As String, formulaRange As Range, lookupColumnAddress As String, lookupRangeAddress As String, StartDateColumnAddress As String
    formulaTemplateString = "=IF(VLOOKUP(@1,@2,@3,FALSE)=@4,1,0)"
    lookupColumnAddress = Replace(IsahExportRange.Cells(2, ProdHeaderDossierCodeColumnIndex).address, "$", "")
    lookupRangeAddress = r.getRangeFullAddress(checkRange, removeFileName:=True, removeDollarSigns:=False)
    StartDateColumnAddress = Replace(IsahExportRange.Cells(2, StartDateColumnIndex).address, "$", "")
    formulaDefinition = str.subInStr(formulaTemplateString, lookupColumnAddress, lookupRangeAddress, checkColumnIndex, StartDateColumnAddress)
    
    Set formulaRange = checkColumn
    If formulaRange.Rows.count > 1 Then
       Set formulaRange = formulaRange.Offset(1, 0).Resize(formulaRange.Rows.count - 1, 1)
       r.fill_formula_range formulaRange, formulaDefinition
    End If
    
    ' set checkColumn 0=Red, 1=Green
    r.setConditionalFormatting checkColumn, 1, 0

    ' hide worksheets
    w.hideWorksheets main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET, main.ISAH_MATCH_BOM_REQUIRED_DATE_SHEET
    
End Sub

Public Sub isah_export_run_all()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim wbconn As ADODB.Connection
    Dim t As Timer
    Set t = New Timer
    
    ' initialize the global connection to thisworkbook, needed below
    Set wbconn = main.getWorkbookConnection()
    
    t.Start ' Start the timer
    
    ' prepare the staging sheet `EXPORT_ISAH` using capgrp sheets
    main.isah_export_stage_orders
    
    ' connect to ISAH database and update the MachGrpCode, Qty and StandCapacity in `ProdBillOfOperation` => 0.1 sec
    main.isah_export_update_prodboo_grp

    ' connect to ISAH database and match orders to dossier in `ProdHeader` table, 0 sec (?)
    main.isah_export_match_prodheader

    ' connect to ISAH database and match dossiers in `ProdBillOfOperation` table,  0.1 sec
    main.isah_export_match_prodboo

    ' update StartTime, EndTime in `ProdHeader` table, 1.2 sec
    main.isah_export_update_prodheader

    ' update StartTime, EndTime in `ProdBillOperation` table, 1 sec
    main.isah_export_update_prodboo

    ' update RequiredDate in `ProdBillOfMat` table, 1 sec
    main.isah_export_update_prodbom

    ' check RequiredDate in `ProdBillOfMat` table, 1 sec
    main.isah_export_check_bom_dates
    main.isah_export_match_bom_dates

    Debug.Print t.StopTimer 'Stop the timer
    GoTo clean_up
    
clean_up:
    wbconn.Close
    Set wbconn = Nothing
    Set WorkBookConnection = Nothing
    
End Sub

' ISAH ARTICLE IMPORT
Sub isah_import_articles()
    ' Objective: Run query against database and paste the return records on sheet.
    ' Definitions: database connection is `conn`, target worksheet is `ws`, query string is `sql`, recordset is `rs0`
    
    ' Declare variables
    Dim conn As ADODB.Connection
    Dim ws As Worksheet
    Dim sql As String
    Dim rs0 As ADODB.Recordset
    Dim tabname As String
    Dim recordsArray As Variant
    
    ' Get database connection using connection string
    Set conn = main.getISAHconnection()
    
    ' Get table name
    tabname = getISAHpart()
    
    ' Define SQL query
    sql = "SELECT LTRIM(RTRIM(PartCode)) as Artikel, Description as Omschrijving, BasicMat as Aantal_per_pallet FROM " & tabname & " WHERE BasicMat <> ''"
    
    ' Get recordset from query
    Set rs0 = db.queryDB(conn, sql)
    
    ' Check if recordset has more than 1 record
    If rs0.RecordCount > 1 Then
        ' Get target worksheet
        Set ws = w.get_or_create_worksheet(main.NUMBER_PER_PALLET_SHEET_NAME, ThisWorkbook)
        
        ' Clear worksheet contents and formatting
        w.clearWorksheet ws
        
        ' Convert recordset to array
        recordsArray = db.RecordSetToArray(rs0)
        
        ' Paste recordset as array on A1
        ws.Range("A1").Resize(UBound(recordsArray, 1), UBound(recordsArray, 2)).value = recordsArray
        
        ' Set filled range as named range
        r.update_named_range main.NUMBER_PER_PALLET_NAMED_RANGE, ws.Range("A1").Resize(UBound(recordsArray, 1), UBound(recordsArray, 2)), ThisWorkbook
    
        ' Set column formats
        r.formatRangeColumns ThisWorkbook.Sheets(main.NUMBER_PER_PALLET_SHEET_NAME), main.MAP_NUMBER_PER_PALLET_COL_TO_FMT
        
        ' calculate worksheet such that references are updated
        Application.Calculate
    Else
        ' Show message box if no articles found
        MsgBox main.ERR_MSG_NO_ARTICLES
    End If
    
    ' Clean up
    rs0.Close
    Set rs0 = Nothing
    conn.Close
    Set conn = Nothing
End Sub

' 6. UI (BUTTON, DROPDOWN) HANDLERS
Public Sub btn_update_isah_data_Click()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim capgrpSheet As Worksheet
    Set capgrpSheet = ThisWorkbook.Worksheets(ActiveSheet.name)
    
    ' EXCEPTIONS: ISAH input sheet is empty
    If Not main.check_isah_input() Then
       capgrpSheet.Activate
       Application.ScreenUpdating = True
       Application.EnableEvents = True
       Exit Sub
    End If
    
    main.update_orders_range capgrpSheet.name
    capgrpSheet.Range("A1").Activate
    
    ' after updating isah data store the resulting state
    main.SafeStoreCurrentState ActiveSheet
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub btn_add_record_Click()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    main.insert_record ActiveSheet.name
    ' after insertion store the resulting state
    main.SafeStoreCurrentState ActiveSheet
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub btn_delete_record_Click()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    main.delete_record ActiveSheet.name
    ' after deletion store the resulting state
    main.SafeStoreCurrentState ActiveSheet
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub btn_restore_prev_state_Click()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    state_control.restoreLastCapgrpState
    state_control.removeLastState
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub btn_print_dates_Click()
    Application.ScreenUpdating = False 'Otherwise main.print_planning doesnt insert the headers correctly, somehow
    Application.EnableEvents = False
    main.print_planning ActiveSheet.name, print_pdf:=False, delete_print_sheet:=True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub btn_export_pdf_Click()
    Application.ScreenUpdating = False 'Otherwise main.print_planning doesnt insert the headers correctly, somehow
    Application.EnableEvents = False
    main.print_planning ActiveSheet.name, print_pdf:=True, delete_print_sheet:=True, show_pdf_exported_msg:=main.PRINT_SHOW_PDF_EXPORTED
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub btn_calculate_dates_Click()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    capgrp = ActiveSheet.name
    If main.get_orders_range(capgrp).Rows.count > 1 Then
       main.update_start_end_times capgrp, main.P_DEBUG
    End If
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub control_sheet_update_database_settings(Optional ByVal Target As Range)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim connectionStrings As Range
    Dim worksheetChanged As Boolean
    Dim ws0 As Worksheet: Set ws0 = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)
    
    If Not Target Is Nothing Then
       worksheetChanged = Not Intersect(Target, ws0.Range(main.DATABASE_DROPDOWN_ADDR)) Is Nothing
    Else
       worksheetChanged = False
    End If
    
    Set connectionStrings = ThisWorkbook.Names(main.CONNECTION_STRINGS_NAMED_RANGE).RefersToRange
    If worksheetChanged Or True Then
        Dim selectedName As String
        Dim i As Long
        selectedName = ws0.Range(main.DATABASE_DROPDOWN_ADDR).value
        For i = 1 To connectionStrings.Rows.count
            If connectionStrings.Cells(i, 1).value = selectedName Then
                ' Do something with the connection string
                ' For example, store it in another cell
                ws0.Range(main.SELECTED_CONNECTION_STRING_ADDR).value = connectionStrings.Cells(i, 2).value
                ws0.Range(main.SELECTED_DATABASE_NAME_ADDR).value = connectionStrings.Cells(i, 3).value
                Exit For
            End If
        Next i
    Else
    
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' 5. print planning
Public Sub print_planning(capgrp As String, Optional print_pdf As Boolean = True, Optional delete_print_sheet As Boolean = True, Optional show_pdf_exported_msg As Boolean = False)
    ' Prints the planning for a specified capacity group (capgrp).
    ' If print_pdf is True, it exports the planning to a PDF file.
    ' If delete_print_sheet is True, it deletes the print sheet after printing.
    ' If show_pdf_exported_msg is True, it shows a message box with the absolute path of the exported PDF.
    '
    ' Parameters:
    ' capgrp - The name of the capacity group for which the planning is being printed.
    ' print_pdf - An optional boolean indicating whether to export the planning to a PDF file.
    ' delete_print_sheet - An optional boolean indicating whether to delete the print sheet after printing.
    ' show_pdf_exported_msg - An optional boolean indicating whether to show a message box with the PDF export path.
    
    Dim ws As Worksheet, rng0 As Range, range_name As String, headerRange As Range
    Dim ws0 As Worksheet, ws1 As Worksheet, rng_orders As Range, rng_workdays As Range, wkNumber As String
    Dim pdf_path As String
    
    ' if printing PDF, get and validate the location beforehand
    If print_pdf = True Then
       pdf_path = main.get_capgrp_print_location(capgrp)
       If pdf_path = "" Then
          Exit Sub
       End If
    End If
    
    ' Get the current weeknumber
    wkNumber = main.get_capgrp_weeknumber(capgrp)
    
    'create new sheet main.PRINT_SHEET_NAME, paste rng0 on new sheet
    Set ws1 = w.get_or_create_worksheet(main.PRINT_SHEET_NAME, ThisWorkbook, True)
    
    ' 20240416: fit the capgrp orders value to range in case of detachment of sheet values to range
    main.fit_order_range_to_values capgrp
    
    ' copy orders range
    Set rng_orders = main.get_orders_range(capgrp)
    Set ws0 = rng_orders.Worksheet
    
    ' Copy values and formats from source range to target range
    r.copy_range rng_orders, main.PRINT_ORDERS_ADDRESS, ws1

    ' Set the print orders range on sheet `PRINT_SHEET_NAME`
    ws1.Activate
    Dim add_rows As Integer, add_cols As Integer
    add_rows = rng_orders.Rows.count - 1
    add_cols = rng_orders.columns.count - 1
    Set rng0 = r.extend_range(main.PRINT_ORDERS_ADDRESS, ws1, add_rows:=add_rows, add_cols:=add_cols)
    
    ' Rename `rng_orders` columns using dictionary
    Dim renameColumnsDict As Scripting.Dictionary
    Set renameColumnsDict = dict.getDictionaryFromString(main.PRINT_RENAME_COLUMNS)
    Set headerRange = r.get_header(rng0, ws1, ThisWorkbook)
    For Each cl In headerRange.Cells
        If renameColumnsDict.exists(cl.value) Then
           cl.value = renameColumnsDict.item(cl.value)
        End If
    Next cl

    ' Autofit columns in the target range
    rng0.EntireColumn.AutoFit

    ' Add outside border and border for each row
    rng0.BorderAround Weight:=xlMedium, ColorIndex:=xlColorIndexAutomatic
    rng0.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rng0.Borders(xlInsideHorizontal).Weight = xlThin
    
    ' Get the header range of the range
    Set headerRange = rng0.Rows(1)
    ' Set header row to bold font
    headerRange.Rows(1).Font.Bold = True
    headerRangeAddress = headerRange.address  'for later, see below
    
    ' Remove interior colors
    rng0.Interior.ColorIndex = xlColorIndexNone
    
    ' Delete columns `COLUMNS_HIDE_FOR_PRINT`
    main.delete_columns_for_print rng0
    
    ' Copy `capgrp_workdays` values and formats to print sheet
    Set rng_workdays = r.get_range(capgrp & "_workdays")
    ' truncate workdays to first 8 rows
    Set rng_workdays = r.subset_range(rng_workdays, 1, 8)
    r.copy_range rng_workdays, main.PRINT_WORKDAYS_ADDRESS, ws1
    
    ' Rename `worktimes` columns using dict
    Dim renameWorkTimesColumnsDict As Scripting.Dictionary, workdaysRng As Range
    Set renameWorkTimesColumnsDict = dict.getDictionaryFromString(main.MAP_WORKDAYS_TIMES_TO_LABEL)
    Set workdaysRng = r.extend_range(main.PRINT_WORKDAYS_ADDRESS, ws1, add_cols:=main.numberOfWorkTimeBlocks)
    Set headerRange = r.get_header(workdaysRng, ws1, ThisWorkbook)
    For Each cl In headerRange.Cells
        If renameWorkTimesColumnsDict.exists(cl.value) Then
           cl.value = renameWorkTimesColumnsDict.item(cl.value)
        End If
    Next cl
    
    ' Copy `capgrp_extra_info` values and formats to print sheet
    Dim rng_extra_info As Range
    Set rng_extra_info = r.get_range(capgrp & "_extra_info")
    r.copy_range rng_extra_info, main.PRINT_EXTRA_INFO_ADDRESS, ws1
    
    ' Style `capgrp_workdays` range
    ws1.Activate
    Set rng0 = r.extend_range(main.PRINT_WORKDAYS_ADDRESS, ws1, add_rows:=rng_workdays.Rows.count - 1, add_cols:=rng_workdays.columns.count - 1)
    With rng0
        ' Autofit columns in the target range
        .EntireColumn.AutoFit
        
        ' Add outside border and border for each row
        .BorderAround Weight:=xlMedium, ColorIndex:=xlColorIndexAutomatic
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
    
        ' Get the header range of the range
        .HorizontalAlignment = xlCenter
        .columns(1).HorizontalAlignment = xlLeft
        Set headerRange = .Rows(1)
        With headerRange
            ' Set header row to bold font
            .Font.Bold = True
        End With

        ' Remove interior colors
        .Interior.ColorIndex = xlColorIndexNone
    End With
        
    ' setup PrintArea, which includes both copied `workdays` and `orders` ranges
    Application.ScreenUpdating = True
    Dim print_area_address As String, r1 As Long, c1 As Long, print_rng As Range
    Set ws = ws1
    r1 = r.get_last_row(main.PRINT_ORDERS_ADDRESS, ws, range2:=Range("A999")) '20231225: do double-sized search (down and up) to "correct" for gaps in the search range
    c1 = r.get_last_col(main.PRINT_ORDERS_ADDRESS, ws)
    print_area_address = r.get_range_address(ws, 1, r1, 1, c1)
    Set print_rng = ws.Range(print_area_address)
    
    ws.PageSetup.PrintArea = print_area_address
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False

    ' Set the margins to as small as possible
    With ws.PageSetup
      .LeftMargin = Application.InchesToPoints(0.25)
      .RightMargin = Application.InchesToPoints(0.25)
      .TopMargin = Application.InchesToPoints(0.5)
      .HeaderMargin = Application.InchesToPoints(0)
      '.BottomMargin = Application.InchesToPoints(0.25)
      '.FooterMargin = Application.InchesToPoints(0.1)
    End With
    
    'Set the header text and format
    header_text = "Planning wk " & wkNumber & " " & capgrp
    ws.Range(main.PRINT_TITLE_START).value = header_text
    ws.Range(main.PRINT_TITLE_RANGE).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    With Selection.Font
        .name = "Calibri"
        .Size = 48
    End With
    
    With ws.PageSetup
     .TopMargin = Application.InchesToPoints(0.75)
     .Orientation = xlLandscape
    End With

    ' Set the footer text, add to footer 2 lines:
    Dim WorktimesMapString As String
    WorktimesMapString = dict.dictionaryToString(dict.invertDictionaryObject(dict.getDictionaryFromString(main.MAP_WORKDAYS_TIMES_TO_LABEL)))
    textFooterFirstLine = Replace(WorktimesMapString, ";", ", ")
    textFooterSecondLineLeft = "FP-2/1 03/24"
    textFooterSecondLineRight = dt.format_datetime(Now())
    r.SetCustomFooter ws.PageSetup, textFooterFirstLine, textFooterSecondLineLeft, textFooterSecondLineRight
    
    ' Insert `print_rng` header at linebreaks
    ' First get the page breaks rows
    Dim pageBreakType As XlPageBreak, pageBreakRows As New collection
    For i = 1 To print_rng.Rows.count
     pageBreakType = print_rng.Rows(i).PageBreak
     If pageBreakType = xlPageBreakAutomatic Then
        pageBreakRows.Add i
     End If
    Next
    clls.printItems pageBreakRows

    ' Then loop over the pagebreak rows, increment after each header row insert (as pagebreak row will have moved down). If the target pageBreak row
    ' is at the bottom of the print range, do not insert header
    Dim targetRange As Range
    Set headerRange = print_rng.Worksheet.Range(headerRangeAddress)
    c = 0
    For Each i In pageBreakRows
        targetRow = i '+ c
        If targetRow >= print_rng.Rows.count Then
           GoTo nx_break
        End If
        Set targetRange = print_rng.Rows(targetRow)
        headerRange.Select
        Selection.Copy
        targetRange.Insert shift:=xlDown
        c = c + 1
nx_break:
    Next
    
    Application.ScreenUpdating = False
    
    ' Start printing dialog
    If print_pdf = False Then
      'Open the Print dialogue box
      Application.Dialogs(xlDialogPrint).Show 'arg11:=True
    Else
      'Suppress the save dialog box
      Application.DisplayAlerts = False
        
      'Print the named range to a PDF file
       print_rng.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdf_path, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
       If main.P_DEBUG Then
          Debug.Print str.subInStr("PDF `@1` exported to `@2`", capgrp, pdf_path)
       End If
              
      'Restore the display alerts setting
       Application.DisplayAlerts = True
    End If
        
    ' Delete print sheet
    If (delete_print_sheet) Then
       w.delete_worksheet main.PRINT_SHEET_NAME
       ' Activate source sheet
       ws0.Activate
    End If
    
    If print_pdf And show_pdf_exported_msg Then
       MsgBox "PDF exported to: " & pdf_path
    End If
        
End Sub

Sub print_add_planning_extra_info()
    ' Add ExtraInfo range
    Dim extraInfoFields As Variant, wsPrint As Worksheet
    extraInfoFields = Split(main.PRINT_EXTRA_INFO_FIELDS, ";")
    Set wsPrint = ThisWorkbook.Worksheets(main.PRINT_SHEET_NAME)
    
    Dim extraInfoRange As Range
    Set extraInfoRange = wsPrint.Range(main.PRINT_EXTRA_INFO_ADDRESS)
    
    ' Simple extra info range
    extraInfoRange.Merge
    
    ' Add borders to ExtraInfo
    extraInfoRange.Borders.LineStyle = xlContinuous
    
    Exit Sub
    
    ' Merge and fill the first row
    With extraInfoRange.Rows(1)
        .Merge
        .value = "Extra info"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' Merge and fill the 2nd and 3rd rows
    With extraInfoRange.Range(Cells(2, 1), Cells(3, 1))
        .Merge
        .value = extraInfoFields(0)
    End With
    
    With extraInfoRange.Range(Cells(2, 2), Cells(3, 3))
        .Merge
    End With
    
    ' Merge and fill the 4th and 5th rows
    With extraInfoRange.Range(Cells(4, 1), Cells(5, 1))
        .Merge
        .value = extraInfoFields(1)
    End With
    With extraInfoRange.Range(Cells(4, 2), Cells(5, 3))
        .Merge
    End With
    
    ' Merge and fill the 6th and 7th rows
    With extraInfoRange.Range(Cells(6, 1), Cells(7, 1))
        .Merge
        .value = extraInfoFields(2)
    End With
    With extraInfoRange.Range(Cells(6, 2), Cells(7, 3))
        .Merge
    End With
    
End Sub

Sub test_print_pdf()
main.print_planning "LN 1", True, False
'main.print_add_planning_extra_info

End Sub


Sub hide_columns_for_print(rng0 As Range)

    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = rng0.Worksheet
    
    ' Assume worksheet is fully unhidden
    ws.Cells.EntireColumn.Hidden = False
    
    ' Loop through array and hide the columns
    column_names = Split(main.COLUMNS_HIDE_FOR_PRINT, ",")
    For i = LBound(column_names) To UBound(column_names)
        column_name = column_names(i)
        If r.column_exist(rng0, column_name) Then
           col_index = r.get_column(rng0, column_name).column
           ws.columns(col_index).EntireColumn.Hidden = True
        End If
    Next i

End Sub

Sub delete_columns_for_print(rng0 As Range)
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = rng0.Worksheet
    
    ' Assume worksheet is fully unhidden
    ws.Cells.EntireColumn.Hidden = False
    
    ' Loop through array and delete the columns
    column_names = Split(main.COLUMNS_HIDE_FOR_PRINT, ",")
    For i = LBound(column_names) To UBound(column_names)
        column_name = column_names(i)
        If r.column_exist(rng0, column_name) Then
           col_index = r.get_column(rng0, column_name).column
           ws.columns(col_index).Delete
        End If
    Next i
End Sub

'BUTTONS CONTROL SHEET
'1. import data for each art capgrp sheet => "Importeren alle artikelen"
Sub btn_import_art_Click()
  Dim ctrl_sheet As Worksheet: Set ctrl_sheet = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)
  
  ' EXCEPTIONS: ISAH input sheet is empty
  If Not main.check_isah_input() Then
     GoTo clean_up
  End If
  
  Dim ws As Worksheet
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  
  ' 20241218: Import articles sheet as its deactivated in Workbook_Open
  main.isah_import_articles
  
  ' Format isah input range columns to expected data types
  main.format_isah_input_range
  
  ' Add new capgrp tabs before importing orders
  main.add_capgrp_sheets
  
On Error GoTo control_sheet
  ' loop over capgrps as update orders
  For Each c In main.get_capgrp_sheet_names()
     capgrp_sheet = CStr(c)
     Debug.Print "updating sheet " & capgrp_sheet
     main.update_orders_range capgrp_sheet
  Next
GoTo clean_up
 
control_sheet:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Err.Raise Err.Number 'rethrow error for user

clean_up:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Sub

'2. import bulk orders from Template to BULK
Sub btn_import_bulk_Click()
  Dim ctrl_sheet As Worksheet: Set ctrl_sheet = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)

  ' EXCEPTIONS: ISAH input sheet is empty
  If Not main.check_isah_input() Then
     GoTo clean_up
  End If
  
  Application.ScreenUpdating = False
  Application.EnableEvents = False

On Error GoTo control_sheet
  main.update_bulk_capgrp_orders
GoTo clean_up
 
control_sheet:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Err.Raise Err.Number 'rethrow error for user

clean_up:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Sub

' find new capgrps in Template and add as sheet
Public Sub btn_add_capgrp_sheets_Click()
  Dim ctrl_sheet As Worksheet: Set ctrl_sheet = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)

  ' EXCEPTIONS: ISAH input sheet is empty
  If Not main.check_isah_input() Then
     GoTo clean_up
  End If
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  
  On Error GoTo control_sheet

  ' add capgrp sheets
  main.add_capgrp_sheets

  ' sort sheets to make sure new capgrp is in right location
  main.init_worksheet_sorting
  GoTo clean_up
  
control_sheet:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Err.Raise Err.Number 'rethrow error for user

clean_up:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Sub

Sub add_capgrp_sheets()
  Dim capgrp_sheets As collection, capgrp_sheet As String
  Set capgrp_sheets = main.get_template_capgrp_names() ' get Template capgrp codes
  For Each c In capgrp_sheets
     capgrp_sheet = c
     If Not w.sheet_exists(capgrp_sheet) Then
        'initialize NEW capgrp sheet
        main.init_capgrp_sheets capgrp_sheet
        
     End If
  Next
End Sub

Public Sub btn_isah_database_export_Click()
  Dim ctrl_sheet As Worksheet: Set ctrl_sheet = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  
On Error GoTo control_sheet
  main.isah_export_run_all
  GoTo clean_up
  
control_sheet:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Err.Raise Err.Number 'rethrow error for user
  
clean_up:
  Set ws = ThisWorkbook.Sheets(main.ISAH_STAGING_SHEET_NAME)
  ws.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Sub

' clear all capgrp sheet orders, workdaytimes and control inputs
Public Sub btn_clear_sheet_Click()
  Dim ctrl_sheet As Worksheet: Set ctrl_sheet = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)
  Application.ScreenUpdating = False
  Application.EnableEvents = False

On Error GoTo control_sheet
  main.clear_all_capgrp_sheets
  
GoTo clean_up

control_sheet:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Err.Raise Err.Number 'rethrow error for user

clean_up:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  
End Sub

Public Sub btn_isah_update_aantal_per_pallet_Click(Optional showMsgBox As Boolean = True)
  Dim ctrl_sheet As Worksheet: Set ctrl_sheet = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)
  Application.ScreenUpdating = False
  Application.EnableEvents = False

  On Error GoTo control_sheet
  main.isah_import_articles
  If showMsgBox Then
     MsgBox "Sheet geupdatet: " & main.NUMBER_PER_PALLET_SHEET_NAME
  End If
  GoTo clean_up
  
control_sheet:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Err.Raise Err.Number 'rethrow error for user

clean_up:
  ctrl_sheet.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True

End Sub

Public Sub btn_import_testdata_Click()
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)
  
  tests.set_input_isah_to_wk29
  
  ws.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Sub

Public Sub btn_add_new_orders_Click()
    ' This subroutine processes new orders from the "NIEUW" sheet and adds them to the appropriate capgrp sheets.
    ' It checks for the existence of the capgrp sheet, verifies if the order already exists, and ensures the week number matches.
    ' If any condition is not met, it displays an error message and exits the subroutine.
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim ws As Worksheet, new_orders_rng As Range, capgrp As String, article As String, record As Dictionary
    Dim errmsg As String, prodOrder As String, capgrp_weeknumber As Integer
    
    ' Get the worksheet and range for new orders
    Set ws = ThisWorkbook.Sheets(main.ISAH_NEW_ORDERS_SHEET_NAME)
    Set new_orders_rng = r.get_range(ws)
    
    ' Check if there are any records in the new orders range
    If new_orders_rng.Rows.count <= 1 Then
        MsgBox "No records on " & ws.name
        GoTo clean_up
    End If
    
    ' Loop over all values in the new_orders_rng column Cap.Grp
    For row_index = 2 To new_orders_rng.Rows.count
        Set record = r.getRowAsRecord(new_orders_rng, row_index)
        capgrp = record.item("Cap.Grp")
        
        ' Check if the capgrp sheet exists
        If Not w.sheet_exists(capgrp) Then
            errmsg = "Capgrp `" & capgrp & "` sheet does not exist"
            MsgBox errmsg
            GoTo clean_up
        End If
        
        ' Check if the capgrp orders range has more than 1 row
        If main.get_orders_range(capgrp).Rows.count <= 1 Then
            errmsg = "Capgrp `" & capgrp & "` sheet does not have orders"
            MsgBox errmsg
            GoTo clean_up
        End If
nx:
    Next row_index
    
    ' Loop over all rows of new_orders_rng
    For row_index = 2 To new_orders_rng.Rows.count
        Set record = r.getRowAsRecord(new_orders_rng, row_index)
        capgrp = record.item("Cap.Grp")
        article = record.item("Artikel")
        prodOrder = record.item("Productieorder")
        
        ' Activate the capgrp sheet and get the orders range
        ThisWorkbook.Sheets(capgrp).Activate
        Dim orders_rng As Range
        Set orders_rng = main.get_orders_range(capgrp)
        
        ' Check if the article already exists in the orders range
        ' TODO rmltr
        ' productie_orders = r.get_column_values(orders_rng, "Productieorder")
        productie_orders = a.getArrayColumnValues(orders_rng.value, "Productieorder")
        
        If u.InList(CLng(prodOrder), productie_orders) Then
            errmsg = "Sheet `" & capgrp & "` productieorder `" & prodOrder & "` already inserted."
            MsgBox errmsg
            GoTo clean_up
        End If
        
        ' Get the week number from the capgrp sheet and compare with the record
        capgrp_weeknumber = main.get_capgrp_weeknumber(capgrp)
        If capgrp_weeknumber <> record.item("ProdWk") Then
            errmsg = "Sheet `" & capgrp & "` weeknumber is `" & capgrp_weeknumber & "`, new order weeknumber `" & record.item("ProdWk") & "`"
            MsgBox errmsg
            GoTo clean_up
        End If
        
        ' Activate the orders_rng first column last cell with row offset 1
        r.safe_offset(orders_rng.Cells(orders_rng.Rows.count, 1), offset_row:=1).Activate
        
        ' Call btn_add_record_Click to add a new record
        Call btn_add_record_Click
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        ' Insert the record into the orders range
        Set orders_rng = main.get_orders_range(capgrp) 'refresh the orders range
        r.insertRecordIntoRange orders_rng, record, orders_rng.Rows.count, validate_columns:=False

nx2:
    Next row_index

clean_up:
  Set ws = ThisWorkbook.Sheets(main.CONTROL_SHEET_NAME)
  ws.Activate
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Sub

' STATE MANAGEMENT
Sub SafeStoreCurrentState(ws As Worksheet)
    ' store current state
    ' TODO update from tests
    If main.P_DEBUG Then
       On Error GoTo 0
       state_control.storeCapgrpState ws
       On Error GoTo handle_error
    Else
       On Error Resume Next
       state_control.storeCapgrpState ws
       Debug.Print "Error in: SafeStoreCurrentState"
       On Error GoTo handle_error
    End If
handle_error:
    Exit Sub
End Sub

' CHECKS AND EXCEPTIONS
Function check_isah_input_empty() As Boolean
    Dim rng0 As Range
    Set rng0 = main.get_isah_input_range()
    If rng0.Rows.count <= 1 Then
       MsgBox main.ERR_MSG_ISAH_INPUT_EMPTY
       check_isah_input_empty = False
    Else
       check_isah_input_empty = True
    End If
End Function

Function check_isah_input_columns() As Boolean
    Dim rng0 As Range, header As Range
    Set rng0 = main.get_isah_input_range()
    Set header = r.get_header(rng0, rng0.Worksheet, ThisWorkbook)
    
    Dim header_columns As collection, required_columns As collection, missing_columns As collection
    Set header_columns = clls.toCollection(header)
    Set required_columns = clls.toCollection(main.INPUT_DATA_HEADER)
    Set missing_columns = clls.getComplementItems(header_columns, required_columns)
    If missing_columns.count > 0 Then
        missing_columns_list = clls.collectionToString(missing_columns, ", ")
        errmsg = "Following columns are missing: " & missing_columns_list
        MsgBox errmsg
        check_isah_input_columns = False
    Else
       Debug.Print "All required columns in header", main.INPUT_ISAH_SHEET
        check_isah_input_columns = True
    End If
End Function

Function check_isah_input() As Boolean
    If Not check_isah_input_empty() Then
       check_isah_input = False
    ElseIf Not check_isah_input_columns() Then
       check_isah_input = False
    Else
       check_isah_input = True
    End If
End Function

' 9. OPEN WORKBOOK METHODS
Sub init_articles_per_pallet()
   If main.checkIsahTestQuery() Then
      main.btn_isah_update_aantal_per_pallet_Click False
   End If
End Sub

' 10. Clean up functions
Sub remove_capgrp_sheet(capgrp As String)
    ' This subroutine removes a capgrp sheet and its associated named ranges.
    ' It first finds and deletes all named ranges that start with the capgrp name,
    ' and then deletes the capgrp sheet itself.
    '
    ' Parameters:
    ' capgrp - The name of the capacity group (capgrp) sheet to be removed.
    
    Dim wb As Workbook
    Dim namesToDelete As collection
    Dim nameObj As name
    Dim ws As Worksheet
    
    ' Set the workbook
    Set wb = ThisWorkbook
    
    ' Check if the sheet exists before attempting to delete
    If w.sheet_exists(capgrp, wb) Then
        ' Delete the capgrp sheet
        Application.DisplayAlerts = False
        wb.Sheets(capgrp).Delete
        Application.DisplayAlerts = True
    End If
    
    ' Find all named ranges that start with the capgrp name
    Set namesToDelete = u.filterObjectsOnProperty(wb.Names, "Name", prop_pattern:=capgrp & ".*")
    
    ' Delete each named range
    r.deleteNames namesToDelete, wb:=wb
    
End Sub

Sub remove_all_capgrps()
    For Each capgrp_ In main.get_capgrp_sheet_names()
    main.remove_capgrp_sheet CStr(capgrp_)
    Next
End Sub






