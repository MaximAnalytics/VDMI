' functions, procedures to store, restore and print the active capgrp sheet state
Sub storeCapgrpState()
    Dim capgrp_sheet As String, States As collection
    
    ' get the states collection of the current capgrp sheet
    capgrp_sheet = ActiveSheet.name
    If Not main.IsCapgrpSheet(capgrp_sheet) Then
        Exit Sub
    End If
    
    Set States = getCapgrpStates(capgrp_sheet)
    If States Is Nothing Then
      Debug.Print "Activesheet is not a capgrp sheet"
      Exit Sub
    End If
    
    ' store orders, worktimes as array
    orders_arr_0 = main.get_orders_range(capgrp_sheet).formula
    orders_arr = u.ifFunc(IsArray(orders_arr_0), orders_arr_0, Empty)
    worktimes_arr = main.get_worktimes_values_range(capgrp_sheet).Value2
    States.Add Array(orders_arr, worktimes_arr)
    
    ' check the number of stored states, make sure that at most `P_NUM_STORED_STATES` are stored
    If States.count > main.P_NUM_STORED_STATES Then
       clls.pop States, index:=1
    End If
    
    WorksheetStateCollection.Remove capgrp_sheet
    WorksheetStateCollection.Add States, key:=capgrp_sheet
    
    If main.P_DEBUG Then
       Debug.Print "Store current state on capgrp sheet: " + capgrp_sheet
    End If
End Sub

Function getCapgrpStates(capgrp_sheet As String) As collection
    Dim States As New collection
    If main.IsCapgrpSheet(capgrp_sheet) Then
        Set getCapgrpStates = Nothing
    End If
    
    If Not clls.KeyExists(WorksheetStateCollection, capgrp_sheet) Then
        Debug.Print "add key: " + capgrp_sheet
        WorksheetStateCollection.Add States, key:=capgrp_sheet
    Else
        Set States = WorksheetStateCollection(capgrp_sheet)
    End If
    Set getCapgrpStates = States
End Function

Sub removeLastState()
    Dim States As collection, capgrp_sheet As String
    capgrp_sheet = ActiveSheet.name
    Set States = getCapgrpStates(capgrp_sheet)
    If Not States Is Nothing Then
       clls.pop States, States.count
    End If
    WorksheetStateCollection.Remove capgrp_sheet
    WorksheetStateCollection.Add States, key:=capgrp_sheet
End Sub

Function getCapgrpStateArray(Optional shift = -1, Optional dbg As Boolean = True) As Variant
    Dim capgrp_sheet As String, CapgrpStates As collection
    capgrp_sheet = ActiveSheet.name
    If main.IsCapgrpSheet(capgrp_sheet) And clls.KeyExists(WorksheetStateCollection, capgrp_sheet) Then
       Set CapgrpStates = WorksheetStateCollection(capgrp_sheet)
       getCapgrpStateArray = clls.getItem(CapgrpStates, CInt(shift))
    Else
       If dbg Then
          Debug.Print "no CapgrpStates found"
       End If
       getCapgrpStateArray = "" 'has length -1
    End If
End Function

Sub restoreLastCapgrpState()
    Dim capgrp_sheet As String, CapgrpStates As collection
    capgrp_sheet = ActiveSheet.name
    capgrpStateArray = getCapgrpStateArray(-2, True)  'last state is the state before the previous state
    If a.num_array_rows(capgrpStateArray) > -1 Then
        orders_arr_prev = capgrpStateArray(0)
        worktimes_arr_prev = capgrpStateArray(1)
        'a.printArray orders_arr_prev
        'a.printArray worktimes_arr_prev
        
        main.set_worktimes_range_values capgrp_sheet, worktimes_arr_prev
        If Not IsEmpty(orders_arr_prev) And TypeName(orders_arr_prev) <> "String" Then
           main.set_orders_range_values capgrp_sheet, orders_arr_prev
        Else
           Debug.Print "previous orders are empty, clear orders_range"
           main.clear_orders_range capgrp_sheet
        End If
    Else
        Debug.Print "capgrpStateArray is -1"
    End If
End Sub

Sub printLastCapgrpState(Optional index = -1)
    Dim capgrp_sheet As String, CapgrpStates As collection
    capgrp_sheet = ActiveSheet.name
    capgrpStateArray = getCapgrpStateArray(index, True)  'last state is the previous state, see worksheet "base" code
    If a.num_array_rows(capgrpStateArray) > -1 Then
        orders_arr_prev = capgrpStateArray(0)
        worktimes_arr_prev = capgrpStateArray(1)
        a.printArray orders_arr_prev
        a.printArray worktimes_arr_prev
    Else
        Debug.Print "capgrpStateArray is -1"
    End If
End Sub

Sub printNumberCapgrpStates()
    If main.IsCapgrpSheet(ActiveSheet.name) Then
       Debug.Print getCapgrpStates(ActiveSheet.name).count
    End If
End Sub
