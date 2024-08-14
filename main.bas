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
Global Const OUTPUT_DATA_FORMATS As String = "General,General,General,General,General,0,0,0.00,General,General,General,General,General,General,General,ddd hh:mm,ddd hh:mm"
Global Const INPUT_ISAH_SHEET = "Template"
Global Const PRODWK_COLUMN As String = "ProdWk"
Global Const capgrp_column As String = "Cap.Grp"
Global Const BULKCODE_COLUMN As String = "Bulkcode"
Global Const STARTDATE_COLUMN As String = "Starttijd"
Global Const DESCRIPTION_COLUMN As String = "Omschrijving"
Global Const ENDDATE_COLUMN As String = "Eindtijd"
Global Const DURATION_COLUMN As String = "Duur"
Global Const ART_COLUMN As String = "Artikel"
Global Const INPUT_ISAH_SHEET_SORT_KEY As String = "Cap.Grp,Bulkcode,Aantal,Flesformaat"
Global Const COLUMNS_HIDE_FOR_PRINT As String = "Flesformaat,Sluiting"
Global Const QTY_COLUMN As String = "Qty1"

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
Global Const XL_MAX_NUMBER_COLUMNS As Long = 256

Global Const WORKDAYS_RANGE_IDS As String = ",ma,di,woe,do,vrij,ma2,di2,woe2,do2,vrij2"
Global Const WORKDAYS_RANGE_ADDRESS As String = "E2"
Global Const ORDERS_RANGE_MAX_COLUMN_NUMBER As Integer = 15

Global Const BTN_UPDATE_ISAH_ADDRESS As String = "A2"
Global Const BTN_ADD_RECORD_ADDR As String = "A5"
Global Const BTN_DELETE_RECORD_ADDR As String = "A8"
Global Const BTN_RESTORE_PREV_STATE_ADDR As String = "C8"
Global Const BTN_PRINT_DATES_ADDR As String = "A11"
Global Const BTN_CALCULATE_DATES_ADDR As String = "C11"
Global Const BTN_WIDTH = 90
Global Const BTN_HEIGHT = 30
Global Const BTN_LEFT_OFFSET = 10

' button labels
Global Const BTN_RESTORE_PREV_STATE_LABEL = "Ga terug"

' order sheet layout
Global Const N_TOP_ROWS_FREEZE As Integer = 14
Global Const INPUT_FIELD_COLOR = 65535
Global Const WT_HEADER_COLOR = 15123099
Global Const WT_IDS_COLOR = 11389944
Global Const WT_VALUES_COLOR = 13431551
Global Const TANK_LO_COLUMN_WIDTH = 8.5

'PRINT LAYOUT
Global Const PRINT_SHEET_NAME = "PRINT"
Global Const PRINT_ORDERS_ADDRESS = "A11"
Global Const PRINT_WORKDAYS_ADDRESS = "K2"
Global Const PRINT_FILE_RANGE_ADDRESS As String = "C6"
Global Const PRINT_FILE_RANGE_IDS As String = ",Locatie print bestand"
Global Const PRINT_FILE_RANGE_HEADER As String = ","
Global Const PRINT_RENAME_COLUMNS = "Ordernummer=OrderNr;Productieorder=ProdOrd;Resources=Res;# pallets=#Plts;Flesformaat=Fles;Sluiting=Slt;ProdWk=Wk;Land=Ld;Pallettype=PltType"
Global Const PRINT_TITLE_START = "A2"
Global Const PRINT_TITLE_RANGE = "A2:H6"

' COMPONENTS
Global Const BASE_SHEET_NAME As String = "base"
Global Const START_SHEET_NAME As String = "instructie"

' DEFAULTS
Global Const DEFAULT_YEAR = 2023
Global Const DEFAULT_WEEKNUMBER = 20
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

Global Const CAPGRP_BULKCODE_STARTDATE_RANGE = "$G:$S" 'range on capgrp containing columns Bulkcode ... Starttijd
Global Const STARTDATE_COLUMN_INDEX = 13


' CONTROL SHEET
Global Const CONTROL_SHEET_NAME = "overzicht"
Global Const BTN_IMPORT_ART_ADDR = "B2"
Global Const BTN_IMPORT_BULK_ADDR = "B8"
Global Const BTN_ADD_CAPGRP_ADDR = "B14"
Global Const BTN_ISAH_EXPORT_ADDR = "B20"
Global Const BTN_CLEAR_SHEET = "B26"
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
Global Const INPAK_CAPGRP_LIST = "NIVP;VDAM;INPK;SPEC;HCM "
Global Const CAPGRP_SHEET_PATTERN = "^LN(.*?)|INPAK"
Global Const MAP_NUMBER_PER_PALLET_COL_TO_FMT = "Artikel=@;Omschrijving=General;aantal_per_pallet=0"

' ISAH database
Global Const ISAH_STAGING_SHEET_NAME = "EXPORT_ISAH"
Global Const ISAH_STAGING_COLUMNS = "Productieorder;Starttijd;Eindtijd;Resources;Aantal;Duur"
Global Const ISAH_STAGING_COLUMNS_DB_NAMES = "ProdHeaderOrdNr;StartDate;EndDate;MachGrpCode;Qty;Duur"
Global Const ISAH_STAGING_CAGGRP_COLUMN = "CAPGRP"
Global Const ISAH_STAGING_UPDATE_COLUMNS = "match_prod_header;StartDate_header;EndDate_header;ProdHeaderDossierCode;match_prod_boo;StartDate_boo;EndDate_boo;next_StartDate_header;next_Enddate_header;check_dates_header;next_StartDate_boo;next_Enddate_boo;check_dates_boo;next_StartTime_boo;next_StandCapacity_boo;next_MachPlanTime_boo;check_ProdBOOStatusCode"
Global Const ISAH_STAGING_UPDATE_COLUMNS_FORMATS = "0;yyyy-mm-dd hh:mm;yyyy-mm-dd hh:mm;0;0;yyyy-mm-dd hh:mm;yyyy-mm-dd hh:mm;yyyy-mm-dd hh:mm;yyyy-mm-dd hh:mm;0;yyyy-mm-dd hh:mm;yyyy-mm-dd hh:mm;0;General;0;0;General"

Global Const ISAH_STAGING_RANGE_NAME = "isah_staging_orders_range"
Global Const ISAH_STAGING_ORDERNR_INDEX = 1
Global Const ISAH_STAGING_ORDERNR_COLUMN = "ProdHeaderOrdNr"
Global Const ISAH_DATABASE_ORDERNR_COLUMN = "ProdHeaderOrdNr"
Global Const ISAH_DATABASE_CAGGRP_COLUMN = "MachGrpCode"
Global Const ISAH_DATABASE_DOSSIERCODE_COLUMN = "ProdHeaderDossierCode"
Global Const ISAH_DATABASE_DATE_COLUMNS = "convert(VARCHAR(20), StartDate) as StartDate, convert(VARCHAR(20), EndDate) as EndDate"

Global Const ISAH_CHECK_BOM_REQUIRED_DATE_SHEET = "CHECK_PROD_BILL_OF_MAT"
Global Const ISAH_MATCH_BOM_REQUIRED_DATE_SHEET = "JOIN_ISAH_EXPORT_PROD_BOM"
Global Const CHECK_BOM_REQUIRED_DATE_SHEET_COLUMNS_MAP = "ProdHeaderDossierCode=0;min_bom_required_date=yyyy-mm-dd hh:mm;max_bom_required_date=yyyy-mm-dd hh:mm"

'ISAH database constants
Global Const ISAH_MANUAL_UPDATE_PRODBOOSTATUSCODE = "20" ' in ProdBOO set field ProdBOOStatusCode to this value

' Worksheet sorting
Global Const SHEETS_START_ORDER = "instructie;overzicht;Template;BULK;X"
Global Const LAST_CAPGRP_SHEET_NAME = "LN18"

' variables
Dim capgrp As String, capgrp_sheet As String, range_name As String, workdaytimes_range As Range
Dim r0 As Long, r1 As Long, c0 As Long, c1 As Long
Dim rng0 As Range

' NEW DESIGN
' 1.1 initialize from input sheet: new tabs per capgrp=>
' 1.2 append STARTDATE, ENDDATE columns to each sheet
' 2.1 dynamically update STARTDATE, ENDDATE
' 3 export sheets
' 4 ISAH export

' INITIALIZERS: create capgrp sheets, UI panels, buttons
Sub init_capgrp_sheets(Optional capgrp_sheet_filter As String)

  ' get the unique article capgrps, bulk capgrps (starting with U) are handled in different sheet
  Dim capgrp_sheets As collection, rng0 As Range, rng1 As Range, rng2 As Range, ws0 As Worksheet, ws1 As Worksheet
  Dim capgrp As String, range_name As String, orders_rng As Range, prev_wsname As String
  
  ' parameters
  Dim b_init_buttons As Boolean: b_init_buttons = True
  prev_wsname = "BULK" 'previous worksheet name, used to maintain sorting of worksheets
  Application.EnableEvents = False ' to prevent events from worksheet pasting
  Application.ScreenUpdating = False
  
  On Error GoTo handle_error
  
  ' Get all capgrp sheets to initialize
  Set capgrp_sheets = main.get_template_capgrp_names()
  For Each c In capgrp_sheets
     ' 0 select capgrp orders from input ISAH and copy to capgrp_sheet
     capgrp = c
     If Not str.is_empty(capgrp_sheet_filter) And capgrp <> capgrp_sheet_filter Then
        GoTo nextiteration
     ElseIf capgrp = capgrp_sheet_filter Then
        Debug.Print "using filter, initialize capgrp sheet: " + capgrp
     Else
        Debug.Print "initialize capgrp sheet: " + capgrp
     End If
     
     ' if sheet exists, delete and add, copy base sheet CODE and rename
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

     ' 5 initialize controls
     main.init_buttons capgrp, True
     
     ' freeze panes
     w.freeze_top_rows ThisWorkbook.Sheets(capgrp), main.N_TOP_ROWS_FREEZE
      
     'Exit For
nextiteration:
  Next c
  
GoTo clean_up
handle_error:
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number 'rethrow error for user
    
clean_up:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub init_capgrp_sheet()
    capgrp_sheet = "LN 1"
    init_capgrp_sheets capgrp_sheet
End Sub

Sub init_capgrp_worksheet_code()
  Dim capgrp As String
  Set capgrp_sheet_names = main.get_capgrp_sheet_names()
  For Each c In capgrp_sheet_names
     capgrp = c
     vb.CopyWorksheetCode main.BASE_SHEET_NAME, capgrp
  Next
End Sub

Sub init_capgrp_sheets_ALL()
    For Each c In main.get_template_capgrp_names()
        init_capgrp_sheets CStr(c)
    Next
    main.init_worksheet_sorting
    main.init_capgrp_worksheet_code
End Sub

Sub init_orders_range(capgrp As String, range_name As String, overwrite As Boolean)
    Dim rng0 As Range, ws0 As Worksheet, rng1 As Range, c1 As Long

    ' create or get named range
    If overwrite Or Not r.name_exist(range_name) Then
      ' get the length of the input data header (r1), expand downwards to last row
      input_data_columns = str.str_to_array(main.INPUT_DATA_HEADER)
      c1 = a.num_array_columns(input_data_columns)
      Set rng0 = r.expand_range(main.ORDERS_RANGE_ADDR, ws:=ThisWorkbook.Worksheets(capgrp), c1:=c1, dbg:=main.P_DEBUG)
      r.delete_named_range range_name, clear:=False
      r.create_named_range range_name, capgrp, rng0.address, overwrite:=overwrite, expand_range:=False
      'add startdate, enddate columns without formatting
      Set rng0 = add_date_columns(range_name)
    Else
      ' get the named range with orders
      Set rng0 = r.get_range(range_name)
    End If
    
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

Sub init_print_file_range(capgrp As String)
    range_name = capgrp & "_input_print_location"
    r.create_named_range range_name, capgrp, main.PRINT_FILE_RANGE_ADDRESS, header_row:=",", id_row:=main.PRINT_FILE_RANGE_IDS, overwrite:=True
    r.get_range(range_name).Cells(2, 2).value = main.DEFAULT_PRINT_FILE_PATH & "planning_" & capgrp & ".pdf"
    r.get_range(range_name).Cells(2, 2).Interior.Color = main.INPUT_FIELD_COLOR
End Sub

Sub init_buttons(capgrp As String, Optional overwrite As Boolean = True)
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim left_offset As Long
    left_offset = main.BTN_LEFT_OFFSET
    
    ctr.add_button "btn_update_isah_data_" & capgrp, main.BTN_UPDATE_ISAH_ADDRESS, ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, label:="Overhalen orders", h:=main.BTN_HEIGHT, w:=main.BTN_WIDTH, left_offset:=left_offset
    assign_macro_to_btn "btn_update_isah_data_" & capgrp, "btn_update_isah_data_Click"
    
    ctr.add_button "btn_add_record_" & capgrp, main.BTN_ADD_RECORD_ADDR, ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, label:="Regel toevoegen", h:=main.BTN_HEIGHT, w:=main.BTN_WIDTH, left_offset:=left_offset
    assign_macro_to_btn "btn_add_record_" & capgrp, "btn_add_record_Click"
    
    ctr.add_button "btn_delete_record_" & capgrp, main.BTN_DELETE_RECORD_ADDR, ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, label:="Regel verwijderen", h:=main.BTN_HEIGHT, w:=main.BTN_WIDTH, left_offset:=left_offset
    assign_macro_to_btn "btn_delete_record_" & capgrp, "btn_delete_record_Click"
    
    ctr.add_button "btn_restore_prev_state_" & capgrp, main.BTN_RESTORE_PREV_STATE_ADDR, ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, label:=main.BTN_RESTORE_PREV_STATE_LABEL, h:=main.BTN_HEIGHT, w:=main.BTN_WIDTH, left_offset:=left_offset
    assign_macro_to_btn "btn_restore_prev_state_" & capgrp, "btn_restore_prev_state_Click"
    
    ctr.add_button "btn_print_dates_" & capgrp, main.BTN_PRINT_DATES_ADDR, ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, label:="Printen planning", h:=main.BTN_HEIGHT, w:=main.BTN_WIDTH, left_offset:=left_offset
    assign_macro_to_btn "btn_print_dates_" & capgrp, "btn_print_dates_Click"
    
    ctr.add_button "btn_calculate_dates_" & capgrp, main.BTN_CALCULATE_DATES_ADDR, ws:=wb.Sheets(capgrp), _
    overwrite:=overwrite, label:="Actualiseer tijden", h:=main.BTN_HEIGHT, w:=main.BTN_WIDTH, left_offset:=left_offset
    assign_macro_to_btn "btn_calculate_dates_" & capgrp, "btn_calculate_dates_Click"
End Sub

Sub init_worksheet_sorting()
    Dim sheetOrder As New collection, newSheets As collection, lnSheetNames As collection
    ' start with SHEETS_START_ORDER
    Set newSheets = clls.toCollection(main.SHEETS_START_ORDER, ";")
    Set sheetOrder = clls.concatCollections(sheetOrder, newSheets)
    
    ' add the LN sheets
    Set lnSheetNames = main.get_capgrp_sheet_names()
    Set sheetOrder = clls.concatCollections(sheetOrder, lnSheetNames)
    
    ' finally: INPAK, ...
    Set newSheets = clls.toCollection("INPAK;" & main.NUMBER_PER_PALLET_SHEET_NAME, ";")
    Set sheetOrder = clls.concatCollections(sheetOrder, newSheets)
    
    ' order the sheets
    w.orderSheets sheetOrder
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
    
     ctr.add_button "btn_clear_sheet", main.BTN_CLEAR_SHEET, ws:=ws, overwrite:=True, label:="Alle productielijnen wissen", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_clear_sheet", "btn_clear_sheet_Click"
End Sub

Sub layout_control_sheet_buttons()
     Dim ws As Worksheet
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
    
     ctr.add_button "btn_clear_sheet", main.BTN_CLEAR_SHEET, ws:=ws, overwrite:=True, label:="Alle productielijnen wissen", w:=main.BTN_WIDTH
     main.assign_macro_to_btn "btn_clear_sheet", "btn_clear_sheet_Click"
End Sub

' GETTERS
Function get_isah_input_range() As Range
    Dim rng0 As Range, ws0 As Worksheet
    Set ws0 = ThisWorkbook.Sheets(main.INPUT_ISAH_SHEET)
    Set rng0 = r.expand_range(ws0.Cells(1), ws0, dbg:=main.P_DEBUG)
    Set get_isah_input_range = rng0
End Function

Function get_isah_capgrp() As Variant
    Dim rng0 As Range, rng1 As Range
    Set rng0 = r.get_range(ThisWorkbook.Sheets(main.INPUT_ISAH_SHEET))
    Set rng1 = r.get_column_values(rng0, main.capgrp_column)
    get_isah_capgrp = r.get_unique_vals(rng1)
End Function

'get the capgrp names (LN 1, LN 2, ...) from the Template sheet, to use for initializing capgrp sheets
Function get_template_capgrp_names() As collection
    Dim col0 As New collection, map_capgrp_inpak_col As collection, capgrp_sheet As String
    Set art_capgrp_col = a.as_collection(main.get_art_capgrp())
    Set map_capgrp_inpak_col = a.as_collection(Split(CStr(main.INPAK_CAPGRP_LIST), ";"))
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
    Dim ws0 As Worksheet
    Dim col0 As New collection
    For Each ws0 In ThisWorkbook.Sheets
        If str.regexp_match(ws0.name, main.CAPGRP_SHEET_PATTERN) Then
            col0.Add ws0.name
        Else
            GoTo nx_ws
        End If
nx_ws:
    Next
    Set get_capgrp_sheet_names = clls.sort_collection(col0, True)
End Function

Function get_art_capgrp() As Variant
    arr = main.get_isah_capgrp()
    ' skip bulk capgrps
    Dim result() As Variant
    ReDim result(LBound(arr) To UBound(arr))
    Dim i As Long, j As Long
    j = 0
    For i = LBound(arr) To UBound(arr)
        If left(arr(i), 1) <> "U" Then
            result(j) = arr(i)
            j = j + 1
        End If
    Next i
    ReDim Preserve result(0 To j - 1)
    get_art_capgrp = result
End Function

Function get_last_art_capgrp() As String
    Dim col0 As collection
    Set col0 = a.as_collection(main.get_art_capgrp())
    If col0.count > 0 Then
    get_last_art_capgrp = col0.item(col0.count)
    Else
    get_last_art_capgrp = ThisWorkbook.Sheets(1).name
    End If
End Function

Function get_bulk_capgrp() As Variant
    arr = main.get_isah_capgrp()
    ' keep bulk capgrps, starting with "U"
    Dim col0 As New collection
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If left(arr(i), 1) = "U" Then
            col0.Add arr(i)
        End If
    Next i
    If col0.count > 0 Then
       get_bulk_capgrp = a.to_array(col0)
    Else
       Dim arr0 As Variant ' empty "Variant"
       get_bulk_capgrp = arr0
    End If
End Function

Function get_orders_range(capgrp As String) As Range
    Dim ord_range As Range, range_name As String
    range_name = main.get_orders_range_name(capgrp)
    Set get_orders_range = r.get_range(range_name)
End Function

Sub set_orders_range_values(capgrp As String, values)
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
   filtered_array = r.filter_range(rng0, main.capgrp_column, isah_capgrp_array, main.PRODWK_COLUMN, prodWk, remove_filter:=remove_filter, xl_operator:=xlFilterValues)
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
   Dim range_values As Range, range_name As String, capgrp_column_name As String, num_rows As Long
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
   If Not IsArray(bulk_capgrp_array) Then
      ' no BULK capgrps found in ISAH Template, skip procedure
      Debug.Print "no BULK capgrps found in ISAH Template"
      MsgBox main.ERR_MSG_NO_BULK_CODES
      GoTo finally
   End If
   
   ' 3 get arrays insert later into BULK sheet from ISAH Template sheet => TODO: get_column(array,index,offset_row)
   filtered_array = r.filter_range(rng0, main.capgrp_column, bulk_capgrp_array, remove_filter:=True, xl_operator:=xlFilterValues)
   input_data_columns = str.str_to_array(main.BULK_ORDERS_HEADER)
   selected_array = filtered_array
   selected_values_array = a.subset_rows(selected_array, LBound(selected_array) + 1)

   ordernr_arr = a.subset_columns(selected_values_array, ordernr_index, ordernr_index)
   bulkcode_arr = a.subset_columns(selected_values_array, bulkcode_index, bulkcode_index)
   qty_arr = a.subset_columns(selected_values_array, qty_index, qty_index)
   dsc_arr = a.subset_columns(selected_values_array, dsc_index, dsc_index)
   
   ' restore default sorting of isah input sheet
   main.restore_isah_default_sorting
   If a.num_array_rows(selected_array) <= 1 Then
      GoTo finally
   End If

   ' resize to fit ordernr_arr,bulkcode,qty
   num_rows = a.num_array_rows(ordernr_arr)
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
   Set capgrp_sheets = main.get_capgrp_sheet_names()
   For Each c In capgrp_sheets
       capgrp_sheet = c
       If Not w.sheet_exists(c, ThisWorkbook) Then
           GoTo next_c
       End If
       capgrp_column_name = Replace(capgrp_sheet, "LN", "Lijn")
       r.add_named_range_column range_name, capgrp_column_name
        
       ' fill capgrp_column with VLOOKUP
       formulaDef = str.subInStr(formulaTemplate, start_address, capgrp_sheet, main.CAPGRP_BULKCODE_STARTDATE_RANGE, main.STARTDATE_COLUMN_INDEX)
       Set formula_range = r.get_column(range_name, capgrp_column_name)
       Set formula_range = r.subset_range(formula_range, startrow:=1)
       r.fill_formula_range formula_range, formulaDef, True
        
       ' format formula range as datetime
       formula_range.Cells.NumberFormat = main.BULK_DATETIME_COLUMN_FORMAT
next_c:
    Next c
    
    ' create calculation range for SORTERING column, strip out the formating function formatDateVDMIShort/Long
    Dim capgrp_columns_range As Range
    Set rng0 = r.get_range(range_name)
    Set capgrp_columns_range = r.subset_range(rng0, startcol:=c0 + 1)

    Dim calcRange As Range, START_CALCULATION_RANGE As String, calcFormulasRange As Range, cl As Range, sortColumnInputAddress As String
    START_CALCULATION_RANGE = "Q10"
    r.copyRangeFormulas capgrp_columns_range, START_CALCULATION_RANGE, Worksheets(main.BULK_SHEET_NAME)
    Set calcRange = r.expand_range(START_CALCULATION_RANGE, rng0.Worksheet)
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
       capgrp_column_name = Replace(CStr(capgrp_sheet), "LN", "Lijn")
       If r.column_exist(rng0, capgrp_column_name) Then
          col_index = r.get_column_index(rng0, capgrp_column_name)
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
    get_capgrp_print_location = r.get_range(capgrp & "_input_print_location").Cells(2, 2).value
End Function

' This subroutine sets default values for capgrp sheet inputs weeknumber and year
Sub set_capgrp_default_inputs(capgrp_sheet As String)

 Dim current_prodwk As Integer, current_year As Integer, new_prodwk As Integer
 ' Get first new_prodwk from ISAH INPUT
 new_prodwk = WorksheetFunction.Min(r.get_column(main.get_isah_input_range, main.PRODWK_COLUMN, offset_row:=1))

 ' set ProdWk and Year if not set
 current_prodwk = main.get_capgrp_weeknumber(capgrp_sheet)
 current_year = main.get_capgrp_year(capgrp_sheet)
 If current_year = 0 Then
    Call main.set_capgrp_year(capgrp_sheet, CInt(year(Now())))
 End If
 If current_prodwk = 0 Then
    Call main.set_capgrp_weeknumber(capgrp_sheet, new_prodwk)
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
    a.paste_array start_end_times, startdate_column_address, ws
    
    ' fit columns
    r.autofit_columns_rows r.get_column(range_name, main.STARTDATE_COLUMN, ws:=ws)
    r.autofit_columns_rows r.get_column(range_name, main.ENDDATE_COLUMN, ws:=ws)
    
    ' determine the overflow row index in orders_rng and set to BOLD font
    Dim orders_rng As Range, orders_values_rng
    Set orders_rng = get_orders_range(capgrp)
    enddates = r.get_column_values(orders_rng, main.ENDDATE_COLUMN)
    
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

Debug.Print a.num_array_rows(start_end_times)

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
    duration_arr = a.get_array_column(ord, dur_index)
    articles_arr = a.get_array_column(ord, art_index)
    startdates_out = a.create_array(a.num_array_rows(ord), 1)
    
    'get the first starttime from the workdaytimes array (wdt)
    Dim starttime As Double, endtime As Double
    jT = a.num_array_rows(wdt)
    starttime = wdt(1, 1)
    
    ' create empty start_end_times array
    Dim start_end_planned() As Variant, j0 As Long, dur As Double
    ReDim start_end_times(LBound(articles_arr) To UBound(articles_arr), 1 To 2)
    
    ' loop over each article i
    j0 = 1
    For Each i In a.get_row_indexes(articles_arr)
        duration = duration_arr(i, 1)
        If Len(duration) <= 0 Then
           GoTo next_art
        End If
        
        block_starttime = find_block_starttime(wdt, j0, starttime, dbg:=dbg)
        j0 = block_starttime(0)
        starttime = block_starttime(1)
        start_end_times(i, 1) = starttime
        dur = CDbl(duration_arr(i, 1))
        block_endtime = find_block_endtime(wdt, j0, starttime, dur, dbg:=dbg)
        start_end_times(i, 2) = block_endtime(1)
        
        ' set endtime as next starttime
        starttime = block_endtime(1)
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

Function find_block_starttime(wdt As Variant, j0 As Long, ByVal starttime, Optional prep As Long = 0, _
Optional dbg As Boolean = True) As Variant
    Dim jT As Long
    jT = UBound(wdt, 1) - LBound(wdt, 1) + 1
    
    Dim starttime2 As Double
    starttime2 = m.round_up_to_nearest_quarter(CDbl(starttime))
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

Function find_block_endtime(wdt As Variant, j0 As Long, starttime As Double, dur As Double, Optional dbg As Boolean = True _
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
    endtime = dt.add_hours(starttime, dur1)

    'cases: 1. starttime, endtime fits in active block j0
    startworktime = wdt(j0, 1)
    endworktime = wdt(j0, 2)
    ind_active = wdt(j0, 3)
    If (m.gte_dbl(starttime, startworktime) And m.lte_dbl(endtime, endworktime)) And ind_active = 1 Then  ' use gte_dbl because of precision issue
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
    dur1 = dur1 - 24 * (endworktime - starttime)
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
    Dim range_name As String, ordersRng As Range, rangeToClear As Range
    If capgrp = "" Then
       capgrp = ActiveSheet.name
    End If
    range_name = main.get_orders_range_name(capgrp)
    If r.name_exist(range_name) Then
       Set ordersRng = main.get_orders_range(capgrp)
       Set rangeToClear = r.getResizedRange(ordersRng.Cells(1, 1), add_rows:=998, add_cols:=main.XL_MAX_NUMBER_COLUMNS - 1)
       r.clear_range rangeToClear, clear_formatting:=True
       ' reduce orders_range to single cell
       r.subsetNamedRange range_name, 1, 1, 1, 1
    End If
End Sub

Public Sub clear_all_capgrp_sheets()
    Dim worktimes_range As Range
    Dim capgrp_sheet As String
    
    ' clear capgrp sheets
    For Each capgrp_sheet0 In main.get_capgrp_sheet_names()
        capgrp_sheet = CStr(capgrp_sheet0)
        Debug.Print "clearing capgrp sheet:", capgrp_sheet
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
End Sub

' TODO: replace capgrp with capgrp_sheet
Sub update_orders_range(capgrp As String)
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
    main.set_capgrp_default_inputs capgrp
    main.copy_selected_orders capgrp, clear_ws:=False, remove_filter:=True
    main.fit_order_range_to_values capgrp
    
    ' add STARTDATE, ENDDATE columns to orders_range
    add_date_columns range_name
    
    ' fit named range to include new columns
    main.fit_order_range_to_values capgrp
    Set rng1 = r.get_range(range_name, wb:=wb0)
    
    If rng1.Rows.count <= 1 Then
       Debug.Print "update_orders_range: no orders found for capgrp: `" & capgrp & "`"
       Exit Sub
    End If
    
    ' 3 sort on bulkcode and color formatting bulkcode column
    r.sort_range_by_columns rng1, main.BULKCODE_COLUMN
    
    ' general formatting
    Set ws0 = rng1.Worksheet
    If rng1.Rows.count > 1 Then
        Dim i As Long
        ws0.Activate
        formats_array = str.str_to_array(main.OUTPUT_DATA_FORMATS)
        For i = LBound(formats_array) To UBound(formats_array)
            rng1.columns(i + 1).Select
            With Selection
            .NumberFormat = formats_array(i)
            .HorizontalAlignment = xlCenter
            End With
        Next i
        
        ' color formatting
        main.update_orders_color_format capgrp
    End If
    
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
    
    ' 5 move buttons to address
    ctr.move_button "btn_update_isah_data_" & capgrp, main.BTN_UPDATE_ISAH_ADDRESS, ws0, left_offset:=0.25 * main.BTN_WIDTH
    ctr.move_button "btn_add_record_" & capgrp, main.BTN_ADD_RECORD_ADDR, ws0, left_offset:=0.25 * main.BTN_WIDTH
    ctr.move_button "btn_delete_record_" & capgrp, main.BTN_DELETE_RECORD_ADDR, ws0, left_offset:=0.25 * main.BTN_WIDTH
    ctr.move_button "btn_print_dates_" & capgrp, main.BTN_PRINT_DATES_ADDR, ws0, left_offset:=0.25 * main.BTN_WIDTH
    ctr.move_button "btn_calculate_dates_" & capgrp, main.BTN_CALCULATE_DATES_ADDR, ws0, left_offset:=0
    ctr.move_button "btn_restore_prev_state_" & capgrp, main.BTN_RESTORE_PREV_STATE_ADDR, ws0, left_offset:=0
    
    ' 6 update sorting on BULK if times change on capgrp sheet
    main.update_bulk_sorting
    ws0.Activate
    
    ' 7 store state in WorksheetStateCollection
    If main.P_STORE_STATE Then
       WorksheetStateCollection.Add "1"
    End If
    
    Exit Sub
End Sub

Function get_orders_range_name(capgrp_name As String) As String
    get_orders_range_name = capgrp_name & "_orders"
End Function

Sub fit_order_range_to_values(capgrp_name As String)
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
        formulaTemplateString = "=CEILING(@2/VLOOKUP(@1,@3,3,FALSE),1)"
        
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
    Dim rng0 As Range, ws As Worksheet, range_name As String, abs_row As Long
    Dim active_row As Range
    range_name = capgrp & "_orders"
    abs_row = CLng(ActiveCell.row)
    
    Set rng0 = r.get_range(range_name)
    Set ws = rng0.Worksheet
    
    If Not Intersect(rng0, ActiveCell) Is Nothing And abs_row > 1 Then
        ' don't insert when on the header row
        If rng0.Cells(1, 1).row = ActiveCell.row Then
           Exit Sub
        End If
        Set active_row = main.get_row_in_named_range(range_name, abs_row)
        
        'Insert the record into the named range
        active_row.Insert shift:=xlDown
        
        'Update the named range to include the new row
        'r.update_named_range range_name, r.getResizedRange(rng0, add_rows:=1)
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
        fill_bulkcode_color rng1, unique_bulkcodes
    End If
End Sub

' orders column height, width
Sub update_orders_columns_width(capgrp As String)
   Dim rng0 As Range
   Set rng0 = r.get_range(capgrp & "_orders")
   r.autofit_columns_rows rng0

   'set widths of columns Tank, L/O to `TANK_LO_COLUMN_WIDTH`
   r.get_column(rng0, "Tank").ColumnWidth = main.TANK_LO_COLUMN_WIDTH
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

Function get_random_color_indices(n As Integer) As collection
    Dim colorIndices As New collection
    Dim randNum As Integer
    Dim i As Integer
    Dim j As Integer
    Dim isExists As Boolean

    For i = 1 To n
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
    Set color_indices = main.get_color_indices_light(n:=100)
    For i = 1 To 100
    Debug.Print i, color_indices.item(i)
    Next i
End Sub

Function get_color_indices_light(n As Integer, Optional skip_numbers As Variant) As collection
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
    For i = 1 To n
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
    If output.count <> n Then
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

' Function to get ISAH production header based on ISAHProfilename
Function getISAHprodheader() As String
    Dim dbname As String, table_name As String, full_table_name As String
    dbname = main.getISAHdbname()
    full_table_name = "[@1].[dbo].[@2]"
    dbprofile = main.getISAHProfileName()
    'TODO: use function a.InList(dbprofile, JKR;JKR2)
    If dbprofile = "JKR" Or dbprofile = "JKR2" Then
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
    Dim dbname As String, table_name As String, full_table_name As String
    dbname = main.getISAHdbname()
    full_table_name = "[@1].[dbo].[@2]"
    dbprofile = main.getISAHProfileName()
    'TODO: use function a.InList(dbprofile, JKR;JKR2)
    If dbprofile = "JKR" Or dbprofile = "JKR2" Then
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
    Dim dbname As String, table_name As String, full_table_name As String
    dbname = main.getISAHdbname()
    full_table_name = "[@1].[dbo].[@2]"
    dbprofile = main.getISAHProfileName()
    'TODO: use function a.InList(dbprofile, JKR;JKR2)
    If dbprofile = "JKR" Or dbprofile = "JKR2" Then
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
    Dim dbname As String, table_name As String, full_table_name As String
    dbname = main.getISAHdbname()
    full_table_name = "[@1].[dbo].[@2]"
    dbprofile = main.getISAHProfileName()
    'TODO: use function a.InList(dbprofile, JKR;JKR2)
    If dbprofile = "JKR" Or dbprofile = "JKR2" Then
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

Sub isah_export_stage_orders()
    Dim rng0 As Range
    Dim wb0 As Workbook
    Set wb0 = ThisWorkbook

    ' append all capgrp order ranges and paste columns `ISAH_STAGING_COLUMNS` to sheet `ISAH_STAGING_SHEET_NAME`
    ' first get named ranges of orders
    Set capgrp_sheets = main.get_capgrp_sheet_names()
    Dim r0 As Integer, n As Long, orders_arr_all As Variant, ws0 As Worksheet
    Dim columns_to_select As Variant, orderNrCapgrp As String
      
    'parameters
    columns_to_select = Split(main.ISAH_STAGING_COLUMNS, ";")
    isah_database_columns = Split(main.ISAH_STAGING_COLUMNS_DB_NAMES, ";")
      
    ' sheets: ISAH staging and Template
    Set ws0 = wb0.Worksheets(main.ISAH_STAGING_SHEET_NAME)
    w.clearWorksheet ws0, wb0
    arrForProductieOrder = main.get_isah_input_range()

    ' append all capgrp order ranges to array `orders_arr_all`
    n = 1
    For Each c In capgrp_sheets
        capgrp = c
        orders_arr = main.get_orders_range(capgrp)
        r0 = a.num_array_rows(orders_arr) 'number of rows of current capgrp
        If (r0 < 2) Then
            GoTo next_capgrp_sheet
        End If
        
        ' subset columns and set `isah_database_columns` as header
        orders_arr = a.select_array_columns(orders_arr, columns_to_select) '=> TODO FIX ?
        orders_arr = a.setArrayHeader(orders_arr, isah_database_columns)
        
        'if CAPGRP = "INPAK" then get the right Cap.Grp
        If capgrp = "INPAK" Then
           If a.num_array_rows(orders_arr) > 0 Then
              orders_arr = a.AppendColumn(orders_arr, "", main.ISAH_STAGING_CAGGRP_COLUMN)
              cl_index = a.FindArrayColumnIndex(orders_arr, "ProdHeaderOrdNr")
              'a.printArray orders_arr
              For Each rw_index In a.get_row_indexes(orders_arr)
                  If rw_index <= 1 Then
                     GoTo nx_i
                  End If
                  orderNr = Trim(orders_arr(rw_index, cl_index))
                  'For INPAK recover the original Cap.Grp => v20240301
                  orders_arr_filtered = a.QueryArray(arrForProductieOrder, "Productieorder", CStr(orderNr))
                  'a.printArray orders_arr_filtered
                  orderNrCapgrp = Trim(a.getNamedArrayValue(orders_arr_filtered, "Cap.Grp"))
                  orders_arr(rw_index, UBound(orders_arr, 2)) = orderNrCapgrp
nx_i:
              Next rw_index
              'a.printArray orders_arr
              'Exit Sub
           End If
        Else
           ' add column CAPGRP = `capgrp` to orders_arr
           capgrp_column_arr = a.create_vector(r0, capgrp, header_value:=main.ISAH_STAGING_CAGGRP_COLUMN, as_2darray:=True)
           orders_arr = a.AppendColumn(orders_arr, capgrp_column_arr)
        End If
         
        ' append to `orders_arr_all`
        If n = 1 Then
           orders_arr_all = orders_arr
        Else
           orders_arr_values = a.resize_array(orders_arr, r0:=2)
           orders_arr_all = a.concatArrays(orders_arr_all, orders_arr_values)
        End If
        n = n + 1
next_capgrp_sheet:
    Next
    
    ' in result array `orders_arr_all`, filter out all rows where ProdHeaderOrdNr is NULL (Empty or '')
    OrdersWithOrdNrArray = a.RemoveNullsFromArray(orders_arr_all, "ProdHeaderOrdNr")
    
    ' paste array `orders_arr_all` and create named range `main.ISAH_STAGING_RANGE_NAME`
    a.paste_array OrdersWithOrdNrArray, "A1", ws0
    r1 = a.num_array_rows(OrdersWithOrdNrArray)
    c1 = a.num_array_columns(OrdersWithOrdNrArray)
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
    
    'set column formats
    Dim ordersRange As Range: Set ordersRange = r.get_range(main.ISAH_STAGING_RANGE_NAME, wb:=wb0)
    Dim columnRange As Range
    columnsArray = Split(main.ISAH_STAGING_UPDATE_COLUMNS, ";")
    columnFormatsArray = Split(main.ISAH_STAGING_UPDATE_COLUMNS_FORMATS, ";")
    For i = LBound(columnsArray) To UBound(columnsArray)
       colName = columnsArray(i)
       colFormat = columnFormatsArray(i)
       Set columnRange = r.get_column(ordersRange, colName, wb:=wb0)
       columnRange.NumberFormat = colFormat
    Next i
    
    'autofit columns
    ws0.columns.AutoFit
    
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
    orders_to_query = r.get_column_values(ordersRange, main.ISAH_STAGING_ORDERNR_INDEX, wb:=wb0)
    
    ' T_ProductionHeader query
    Dim sql0 As String, table_name As String, sql_template As String
    table_name = main.getISAHprodheader()
    where_statement = db.sqlWhereInCondition(orders_to_query, main.ISAH_DATABASE_ORDERNR_COLUMN, mssql)
    sql_template = "SELECT LTRIM(RTRIM(@1)) as @2, @3, @4 FROM @5 @6;"
    sql0 = str.subInStr(sql_template, ordernr_db_column, ordernr_db_column, main.ISAH_DATABASE_DATE_COLUMNS, dossiercode_column, table_name, where_statement)
    Debug.Print sql0
    
    'input for T_ProductionHeader: connect to db and execute query
    Dim sqlconn As ADODB.Connection, rs0 As ADODB.Recordset
    Set sqlconn = main.getISAHconnection()
  
On Error GoTo close_connection
    Set rs0 = db.queryDB(sqlconn, sql0)
    db.printRecordset rs0, False, False
    result_array = db.RecordSetToArray(rs0)
    sqlconn.Close
On Error GoTo 0
    GoTo no_error
    
close_connection:
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
        
        'TODO FIX: EXCEPTION for CAPGRP="INPAK"
        'If newMachGrpCode = "INPAK" Then
        '   GoTo nx_i
        'End If
                 
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
    Dim conn0 As ADODB.Connection
    Set conn0 = main.getISAHconnection()
    
    c = 0
    ' connection management: make sure to close connections on error
    On Error GoTo close_connection
    For Each stat In sqlStatements
       Debug.Print stat
       conn0.Execute CStr(stat)
       c = c + 1
    Next
    On Error GoTo 0
GoTo no_error

'errorhandler
close_connection:
    conn0.Close
    Err.Raise Err
no_error:
    conn0.Close
    
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
    dossiercode_to_query = r.get_column_values(ordersRange, dossiercode_column, wb:=wb0)
    capgrp_to_query = r.get_column_values(ordersRange, main.ISAH_STAGING_CAGGRP_COLUMN, wb:=wb0)
    
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
    
    ' thisworkbook as wb0
    Set wb0 = ThisWorkbook
    
    ' Connect to this workbook
    Set wbconn = db.openExcelConn(ThisWorkbook)
    
    ' Query from table "EXPORT_ISAH" where `match_prod_header`=1
    sql_template = "SELECT @1 FROM [@2] WHERE @3;"
    sql0 = str.subInStr(sql_template, select_column_list, table_name, where_condition)
    Set rs0 = db.queryDB(wbconn, sql0)
    match_orders_array = db.RecordSetToArray(rs0)
    
    ' Check orders to match, if array is empty then exit sub
    If a.num_array_rows(match_orders_array) <= 1 Then
       Debug.Print "No EXPORT_ISAH records with match_prod_header=1 found"
       ' Close connection to this workbook
       wbconn.Close
       Exit Sub
    End If
    
    ' Create SQL update statement using `ProdHeaderOrdNr` as key column and `StartDate`, `EndDate` as update columns
    update_statements = db.sqlUpdateStatement(rs0, table_name_isah, set_columns, key_columns, mssql)
    Debug.Print update_statements
    
    ' Create query statement `sql1` to select columns from `rs0` where `match_prod_header`=1
    orderNumbers = a.resize_array(a.get_array_column(match_orders_array, 1), r0:=2)
    wherecondition = db.sqlWhereInCondition(orderNumbers, main.ISAH_DATABASE_ORDERNR_COLUMN, mssql)
    sql1 = str.subInStr("SELECT @1 FROM @2 @3;", select_check_columns_list, table_name_isah, wherecondition)
    Debug.Print sql1
    
    ' Close connection to this workbook
    wbconn.Close
    
On Error GoTo close_connection
    ' Open connection to ISAH database
    Set sqlconn = main.getISAHconnection()
    
    ' Execute SQL update statements
    db.executeSqlStatements sqlconn, update_statements, db.MSSQL_LINE_BREAK
    
    ' Query from `sqlconn` and store result as array `result_array`
    result_array = db.RecordSetToArray(db.queryDB(sqlconn, sql1))
    
    ' Close connection to ISAH database
    sqlconn.Close
On Error GoTo 0
    GoTo no_error
    
close_connection:
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
    
    ' thisworkbook as wb0
    Set wb0 = ThisWorkbook
    
    ' Connect to this workbook
    Set wbconn = db.openExcelConn(ThisWorkbook)
    
    ' Query from sheet "EXPORT_ISAH" where `match_prod_boo`=1
    sql_template = "SELECT @1 FROM [@2] WHERE @3;"
    sql0 = str.subInStr(sql_template, select_column_list, table_name, where_condition)
    Set rs0 = db.queryDB(wbconn, sql0)
    
    match_dossiercode_capgrp_array = db.RecordSetToArray(rs0)
    
    ' Check dossiercodes, capgrp to match, if array is empty then exit sub
    If a.num_array_rows(match_dossiercode_capgrp_array) <= 1 Then
        Debug.Print "No EXPORT_ISAH records with match_prod_boo=1 found"
        ' Close connection to this workbook
        wbconn.Close
        Exit Sub
    End If
    
    ' Create SQL update statement using `ProdHeaderOrdNr` as key column and `StartDate`, `EndDate` as update columns
    update_statements = db.sqlUpdateStatement(rs0, table_name_isah, set_columns, key_columns, mssql, force_string:=True)
    Debug.Print update_statements
    
    ' Create query statement `sql_check_columns` to select columns from `rs0` where `match_prod_header`=1
    Dim sql_check_columns As String
    select_check_columns_list = "LTRIM(RTRIM(ProdHeaderDossierCode)) as ProdHeaderDossierCode," + _
    "MIN(StartDate) OVER (PARTITION BY ProdHeaderDossierCode) AS next_StartDate_boo, " + _
    "MIN(EndDate) OVER (PARTITION BY ProdHeaderDossierCode) AS next_EndDate_boo, " + _
    "MIN(CAST(FORMAT(CAST((StartTime+1)/86400.000 AS datetime), 'HH:mm') AS varchar)) OVER (PARTITION BY ProdHeaderDossierCode) AS next_StartTime_boo, " + _
    "MIN(ProdBOOStatusCode) OVER (PARTITION BY ProdHeaderDossierCode) AS ProdBOOStatusCode"
    
    ' Key matching values for update statement
    dossierCodes = a.resize_array(a.get_array_column(match_dossiercode_capgrp_array, 1), r0:=2)
    machgrpCodes = a.resize_array(a.get_array_column(match_dossiercode_capgrp_array, 2), r0:=2)
    
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
    
    ' Close connection to this workbook
    wbconn.Close
     
On Error GoTo close_connection
    ' Open connection to ISAH database
    Set sqlconn = main.getISAHconnection()
    
    ' Execute SQL update statements
    db.executeSqlStatements sqlconn, update_statements, db.MSSQL_LINE_BREAK
    
    ' Query `sql_check_columns` from `sqlconn` and store result as array `ProdBOOArray`
    ProdBOOArray = db.RecordSetToArray(db.queryDB(sqlconn, sql_check_columns))
    a.printArray ProdBOOArray
    
    ' Close connection to ISAH database
    sqlconn.Close
    
    On Error GoTo 0
    GoTo no_error
    
close_connection:
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
    Dim xlsconn As New ADODB.Connection
    sql0 = "SELECT DISTINCT Cstr(ProdHeaderDossierCode) AS ProdHeaderDossierCode, next_StartDate_header AS RequiredDate FROM [" & source_table & "$] WHERE ProdHeaderDossierCode IS NOT NULL AND Startdate_header IS NOT NULL"
    Set rs0 = db.queryFromWorkbook(sql0, xlsconn)

    If db.RecordSetNumberRecords(rs0) = 0 Then
       Debug.Print "isah_export_update_prodbom: ISAH_EXPORT has no valid ProdHeaderDossierCode RequiredDate to update in ISAH"
       Exit Sub
    End If
    
    ' Define the columns for the SET and WHERE clauses of the update statement
    set_columns = "RequiredDate"
    key_columns = "ProdHeaderDossierCode"
    
    ' Construct the SQL update statements and add them to the collection
    update_statements = db.sqlUpdateStatement(rs0, target_table, set_columns, key_columns, mssql)
    
    ' Initialize the collection to store SQL update statements
    Set updateStatements = str.stringToCol(update_statements, ";")
    
    ' Print the update statements
    For Each sqlUpdate In updateStatements
        Debug.Print sqlUpdate
    Next sqlUpdate
    
    ' clean up
    rs0.Close
    Set rs0 = Nothing
    xlsconn.Close
    Set xlsconn = Nothing
    
    ' Update ISAH table BOM
    Dim conn As ADODB.Connection
    
On Error GoTo close_connection
    Set conn = main.getISAHconnection()
    For Each sqlUpdate In updateStatements
        db.executeSql conn, CStr(sqlUpdate)
        'Debug.Print sqlUpdate
    Next sqlUpdate
    
On Error GoTo 0
GoTo no_error
  
close_connection:
    ' Close connection to ISAH database
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
    Dim conn0 As ADODB.Connection, sql0 As String
    sql0 = main_isah_queries.check_ProdBillOfMat(main.getISAHprodbom())
    Set conn0 = main.getISAHconnection()
    db.writeQueryToSheet conn0, sql0, main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET, main.CHECK_BOM_REQUIRED_DATE_SHEET_COLUMNS_MAP
    conn0.Close
    
End Sub

Sub isah_export_match_bom_dates()

    ' 2. join with ISAH_EXPORT with ISAH_CHECK_BOM_REQUIRED_DATE_SHEET, add check column and write to ISAH_MATCH_BOM_REQUIRED_DATE_SHEET
    Dim conn0 As ADODB.Connection, sql0 As String, rs0 As ADODB.Recordset, rng0 As Range
    sql0 = join_ISAH_EXPORT_CHECK_PROD_BOM()
    
    ' check if table main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET has actually been filled
    Set rng0 = ThisWorkbook.Sheets(main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET).Cells(1, 1)
    If rng0.value = "" Then
       Debug.Print str.subInStr("Sheet not filled: @1", main.ISAH_CHECK_BOM_REQUIRED_DATE_SHEET)
       Exit Sub
    End If
    
    Set conn0 = db.openExcelConn(ThisWorkbook)
    db.writeQueryToSheet conn0, sql0, main.ISAH_MATCH_BOM_REQUIRED_DATE_SHEET
    
    ' append column `check_bom_required_date` to EXPORT_ISAH
    Dim checkRange As Range, IsahExportRange As Range, checkColumn As Range
    Set checkRange = r.expand_range("A1", ThisWorkbook.Sheets(main.ISAH_MATCH_BOM_REQUIRED_DATE_SHEET))
    Set checkColumn = r.get_column(checkRange, "check_bom_required_date", offset_row:=1)
    checkColumnIndex = checkColumn.column
    
    ' Append checkColumn values to IsahExportRange
    values = checkColumn
    Set IsahExportRange = main.get_isah_export_range()
    r.AppendColumnToRange IsahExportRange, "check_bom_required_date", values
    
    Set IsahExportRange = main.get_isah_export_range()
    Set checkColumn = r.get_column(IsahExportRange, "check_bom_required_date")
    
    ' Fill the formula formulaTemplateString in the newly inserted column
    Dim formulaDefinition As String, formulaRange As Range, lookupColumnAddress As String, lookupRangeAddress As String
    formulaTemplateString = "=VLOOKUP(@1,@2,@3,FALSE)"
    lookupColumnAddress = Replace(IsahExportRange.Cells(2, 1).address, "$", "")
    lookupRangeAddress = r.getRangeFullAddress(checkRange, removeFileName:=True, removeDollarSigns:=False)
    formulaDefinition = str.subInStr(formulaTemplateString, lookupColumnAddress, lookupRangeAddress, checkColumnIndex)
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
    ' prepare the staging sheet `EXPORT_ISAH` using capgrp sheets
    main.isah_export_stage_orders
    
    ' connect to ISAH database and update the MachGrpCode, Qty and StandCapacity in `ProdBillOfOperation`
    main.isah_export_update_prodboo_grp

    ' connect to ISAH database and match orders to dossier in `ProdHeader` table
    main.isah_export_match_prodheader
    
    ' connect to ISAH database and match dossiers in `ProdBillOfOperation` table
    main.isah_export_match_prodboo
    
    ' update StartTime, EndTime in `ProdHeader` table
    main.isah_export_update_prodheader
    
    ' update StartTime, EndTime in `ProdBillOperation` table
    main.isah_export_update_prodboo
    
    ' update RequiredDate in `ProdBillOfMat` table
    main.isah_export_update_prodbom
    
    ' check RequiredDate in `ProdBillOfMat` table
    main.isah_export_check_bom_dates
    main.isah_export_match_bom_dates
        
End Sub

Sub test_()
    ' check RequiredDate in `ProdBillOfMat` table
    main.isah_export_check_bom_dates
    'main.isah_export_match_bom_dates
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
        r.format_columns ThisWorkbook.Sheets(main.NUMBER_PER_PALLET_SHEET_NAME), main.MAP_NUMBER_PER_PALLET_COL_TO_FMT
        
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

' BUTTON HANDLERS
Public Sub btn_update_isah_data_Click()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim capgrpSheet As Worksheet
    Set capgrpSheet = ThisWorkbook.Worksheets(ActiveSheet.name)
    
    ' EXCEPTIONS: ISAH input sheet is empty
    If Not main.check_isah_input() Then
       capgrpSheet.Activate
       Exit Sub
    End If
    
    main.update_orders_range capgrpSheet.name
    capgrpSheet.Range("A1").Activate
    
    ' after updating isah data store the resulting state
    main.SafeStoreCurrentState
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub btn_add_record_Click()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    main.insert_record ActiveSheet.name
    ' after insertion store the resulting state
    main.SafeStoreCurrentState
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub btn_delete_record_Click()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    main.delete_record ActiveSheet.name
    ' after deletion store the resulting state
    main.SafeStoreCurrentState
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


' print planning
Public Sub print_planning(capgrp As String, Optional print_pdf As Boolean = True, Optional delete_print_sheet As Boolean = True)
    Dim ws As Worksheet, rng0 As Range, range_name As String, headerRange As Range
    Dim ws0 As Worksheet, ws1 As Worksheet, rng_orders As Range, rng_workdays As Range, wkNumber As String
    
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
    header_text = "Planning week " & wkNumber & " " & capgrp
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
    main.SetCustomFooter ws.PageSetup, textFooterFirstLine, textFooterSecondLineLeft, textFooterSecondLineRight
    
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
      'Print the named range
      Application.Dialogs(xlDialogPrint).Show
    Else
      'Suppress the save dialog box
      Application.DisplayAlerts = False
        
      'Print the named range to a PDF file
       print_rng.ExportAsFixedFormat Type:=xlTypePDF, Filename:=main.get_capgrp_print_location(capgrp), Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
      'Restore the display alerts setting
       Application.DisplayAlerts = True
    End If
        
    ' Delete print sheet
    If (delete_print_sheet) Then
       w.delete_worksheet main.PRINT_SHEET_NAME
       ' Activate source sheet
       ws0.Activate
    End If
    
End Sub

Sub SetCustomFooter(pgSetup As PageSetup, ctext, ltext, rtext)
    ' This subroutine sets a custom two-line footer for the active worksheet.
    ' The first line is centered and contains "text1".
    ' The second line has "left text" aligned to the left and "right text" aligned to the right.
    
    ' Clear any existing footers
    With pgSetup
    .LeftFooter = ""
    .CenterFooter = ""
    .RightFooter = ""
    ' Set the first line of the footer, centered
    .CenterFooter = ctext
    ' Set the second line of the footer with left and right aligned texts
    ' The character Chr(10) is used to insert a line break
    .LeftFooter = ltext & Chr(10)
    .RightFooter = rtext & Chr(10)
    
    End With
    
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
  
  Dim capgrp_sheets As collection, capgrp_sheet As String
  
On Error GoTo control_sheet
  Set capgrp_sheets = main.get_template_capgrp_names() ' get Template capgrp codes
  For Each c In capgrp_sheets
     capgrp_sheet = c
     If Not w.sheet_exists(capgrp_sheet) Then
        'initialize NEW capgrp sheet
        main.init_capgrp_sheets capgrp_sheet
     End If
  Next
  
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

' STATE MANAGEMENT
Sub SafeStoreCurrentState()
    ' store current state
    ' TODO update from tests
    If main.P_DEBUG Then
       On Error GoTo 0
       state_control.storeCapgrpState
       On Error GoTo handle_error
    Else
       On Error Resume Next
       state_control.storeCapgrpState
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


' OPEN WORKBOOK METHODS
Sub init_articles_per_pallet()
   If main.checkIsahTestQuery() Then
      main.btn_isah_update_aantal_per_pallet_Click False
   End If
End Sub


