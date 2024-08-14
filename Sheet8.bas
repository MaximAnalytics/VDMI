Private Sub Worksheet_Activate()
    ' safely store the worksheet initial state
    main.SafeStoreCurrentState
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wb0 As Workbook, ws0 As Worksheet, listenerRng As Range, rng0 As Range, num_cols As Long, capgrp As String, _
    orders_range_name As String, worktimes_range_name As String, worktimes0 As Range, ordersRange As Range, _
    durationRange As Range, qtyRange As Range, wkNumberRange As Range, targetRange As Range, _
    ordersRangeFooter As Range
    
    Set wb0 = ThisWorkbook
    Set ws0 = Target.Worksheet
    capgrp = ws0.name
    orders_range_name = main.get_orders_range_name(capgrp)
    worktimes_range_name = main.get_worktimes_range_name(capgrp)

    If Target Is Nothing Then
       Exit Sub
    End If
    target_address = Target.address
    
On Error GoTo handle_error
    Application.ScreenUpdating = False
    Application.EnableEvents = False
        
    ' events updating wkNumberRange inputs on other sheets
    If capgrp = "LN 1" Then
        Dim next_capgrp As String, wkNumberRangeNext As Range
        Set wkNumberRange = main.get_weeknumber_range(capgrp)
        If Not Intersect(Target, wkNumberRange) Is Nothing Then
           If main.P_DEBUG Then
              Debug.Print "weeknumber changed on " & ws0.name & ", set on other sheets.."
           End If
                
           For Each c In main.get_capgrp_sheet_names()
              next_capgrp = c
              If capgrp <> next_capgrp Then
                 Set wkNumberRangeNext = main.get_weeknumber_range(next_capgrp)
                 wkNumberRangeNext.Cells(2, 2).value = wkNumberRange.Cells(2, 2).value
              End If
           Next c
        End If
        ' return to current ws
        ws0.Activate
    End If

    ' events updating ordersRange
    If r.name_exist(orders_range_name, ws0, wb0) And r.name_exist(worktimes_range_name, ws0, wb0) Then
        ' 20240102: check if ordersRange is filled
        Set ordersRange = wb0.Names(orders_range_name).RefersToRange
        If ordersRange.Cells(1, 1).value = "" Then
            Debug.Print "Worksheet_Change: orders range not filled"
            main.SafeStoreCurrentState
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Exit Sub
        End If
        Set durationRange = r.get_column(ordersRange, main.DURATION_COLUMN)
        Set qtyRange = r.get_column(ordersRange, main.QTY_COLUMN)
        Set worktimes0 = wb0.Names(worktimes_range_name).RefersToRange
        If main.P_DEBUG Then
           warningString = str.subInStr("Worksheet_Change on @1, target address is: @2", ws0.name, target_address)
           Debug.Print warningString
        End If
        
        ' Event variables (boolean indicators)
        Dim eventOrdersRange As Boolean
        Dim eventWorkDayTimesRange As Boolean
        Dim eventOrdersFooter As Boolean
        Dim eventDurationRange As Boolean
        
        eventOrdersRange = Not Intersect(Target, ordersRange) Is Nothing
        eventWorkDayTimesRange = Not Intersect(Target, worktimes0) Is Nothing
        eventDurationRange = Not Intersect(Target, durationRange) Is Nothing
        
        ' Handle events on ordersRange but not triggered from changes in worktimes (row inserts, deletes, updates)
        ' If Target is header OR multiple rows, then skip.
        ' 20240107: use `main.WORKSHEET_IGNORE_MULTIROW_EVENTS`=false to test multirow inserts
        If eventOrdersRange Then
            ' if target is header, then skip
            If Target.Cells(1, 1).row = ordersRange.Rows(1).row Then
                warningString = str.subInStr("Worksheet_Change on @1, target address is header, ignore", ws0.name)
                Debug.Print warningString
                main.SafeStoreCurrentState
                Application.ScreenUpdating = True
                Application.EnableEvents = True
                Exit Sub
            End If
            
            ' if target does NOT come from workdaytimes AND is multirow
            If Not eventWorkDayTimesRange Then
                ' if target is multirows
                If Target.Cells.count > 1 And main.WORKSHEET_IGNORE_MULTIROW_EVENTS Then
                     If main.P_DEBUG Then
                        Debug.Print "target is multiple cells, dont do anything: " & target_address
                     End If
                   main.SafeStoreCurrentState
                   Application.ScreenUpdating = True
                   Application.EnableEvents = True
                   Exit Sub
                End If
            End If
        End If
        
        ' if target is from footer: resize the ordersRange => does not work yet, some problem with `r.update_named_range`
        Set ordersRangeFooter = ordersRange.Rows(ordersRange.Rows.count + 1)
        eventOrdersFooter = Not Intersect(Target, ordersRangeFooter) Is Nothing
        If eventOrdersFooter Then
           Set ordersRange = r.expand_range(ordersRange, ws0, wb0)
           r.update_named_range orders_range_name, ordersRange, wb0
           warningString = str.subInStr("Worksheet_Change on @1, target address is orderRange footer, new orderRange is @2", ws0.name, _
           ordersRange.address)
           Debug.Print warningString
        End If
        
        ' update orders: start_end_times, color formats, bulk sorting
        If eventDurationRange Or eventWorkDayTimesRange Or eventOrdersFooter Then
            If main.P_DEBUG Then
               Debug.Print "data changed on sheet " & ws0.name & ", cell:" & target_address
            End If
            main.update_start_end_times capgrp
            main.update_orders_color_format capgrp
            ' worktimesrange formatting if target
            If eventWorkDayTimesRange Then
               r.ClearAllBorders main.get_worktimes_values_range(capgrp)
               r.add_outside_border main.get_worktimes_values_range(capgrp)
            End If
        End If
                
        ' return to current ws
        ws0.Activate
    Else
        If main.P_DEBUG Then
           Debug.Print "named range doesnt exist: " & orders_range_name
        End If
    End If
    
clean_up:
    ' safely store the resulting end state
    main.SafeStoreCurrentState
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
handle_error:
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ws0.Activate
    MsgBox "Onbekende fout opgetreden!", vbCritical     'Show error to user but dont break to VB editor
End Sub



