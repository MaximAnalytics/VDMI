Global Const MSSQL_LINE_BREAK = vbCrLf
Dim errmsg0 As String

' Database functions
' requirements: reference to Microsoft ActiveX Data Objects 6.1 Library, Microsoft Scripting Runtime
' 0. enumerations

' 1. connections
'openMSSQLconn(driver, server, dbname, un, pw) => Establishes a connection to an MSSQL database using provided credentials and returns the connection object.
'openDBconn(connstr) => Opens a database connection with a given connection string and returns the connection object.
'openExcelConn(xls0, connstr, columnsAsText0, printConnStr) => Opens a connection to an Excel workbook and returns it as an ADODB.Connection object.

' 2. data manipulation statements: CRUD
'executeSql(conn0, sql0, close_connection) => Executes a SQL statement on the provided ADODB.Connection and optionally closes the connection.
'executeSqlStatements(conn0, statements, linebr) => Executes a batch of SQL statements separated by a specified line break character.
'queryDB(conn0, sql0, close_connection, dbg) => Executes a SQL query using the provided connection and returns a Recordset.
'queryFromWorkbook(sql0, xlsconn) => Executes a SQL query on an Excel workbook connection and returns a Recordset.
'truncateTable(conn0, tabname) => Truncates the specified table in the database.
'writeQueryToSheet(conn0 As Object, sql0 As String, wsName As String, Optional map_format_column_type As String = "")

' 3. Recordset:
'printRecordset(rs0, print_datatypes, print_field_name) => Prints the contents of a Recordset to the debug output.
'RecordSetNumberRecords(rs0) => Returns the number of records in the provided Recordset.
'RecordSetToArray(rs0) => Converts a Recordset to a two-dimensional array.
'commaSeparateFields(rs0) => Returns a comma-separated string of field names from a Recordset.
'commaSeparateValues(rs0, xlsvalues, dbtype0) => Returns a comma-separated string of field values from a Recordset.

' 4. sql writers
'sqlInsertStatement(rs0, table_name, dbtype0) => Generates an SQL INSERT statement for each record in the provided Recordset.
'sqlSetCondition(record, set_columns, dbtype0) => Constructs an SQL SET clause based on the provided Recordset and column names.
'sqlWhereCondition(record, where_columns, dbtype0, force_string) => Constructs an SQL WHERE clause based on the provided Recordset and column names.
'sqlWhereInCondition(where_values, where_column, dbtype0, force_string) => Constructs an SQL WHERE IN clause for a specified column and list of values.
'sqlUpdateStatement(rs0, table_name, set_columns, where_columns, dbtype0, force_string) => Generates an SQL UPDATE statement for each record in the provided Recordset.

' 9. utilities
'xlToDBvalue(xlvalue, dbType, defvalue, force_string) => Converts an Excel value to a database-compatible string or numeric representation.
'xl_to_xlsdb_value(xlvalue) => Converts an Excel value to a string representation compatible with Excel database queries.

' 0. enumerations
Public Enum dbType
[_first] = 1
oracledb = 1
mysql = 2
mssql = 3
[_last] = 3
End Enum

Public Sub test_db()
Dim filePath As String, conn0 As ADODB.Connection, DataSourceString As String, ConnectionStringMap As Scripting.Dictionary, sql0 As String, rs0 As ADODB.Recordset

    ' 1 connections
    filePath = zz_env.getExcelTestDataFile()
    Set conn0 = db.openExcelConn(filePath)
    Set ConnectionStringMap = dict.getDictionaryFromString(conn0.ConnectionString)
    DataSourceString = ConnectionStringMap.item("Data Source")
    Debug.Assert DataSourceString = filePath
    
    ' 2 data manipulation: DML
    sql0 = "SELECT DISTINCT BasicMat FROM [T_Part_BasicMat$]"
    Set rs0 = db.queryDB(conn0, sql0)
    Debug.Assert RecordSetNumberRecords(rs0) = 317
    recordsArr = db.RecordSetToArray(rs0)
    Debug.Assert a.num_array_columns(recordsArr) = 1 And a.num_array_rows(recordsArr) = 318 ' 317 plus header row
    
    ' query empty recordset
    sql0 = "SELECT DISTINCT BasicMat FROM [T_Part_BasicMat$] WHERE BasicMat<'0'"
    Set rs0 = db.queryDB(conn0, sql0)
    Debug.Assert RecordSetNumberRecords(rs0) = 0 And rs0.Fields.Count = 1 And db.RecordsSetHasFields(rs0)
    Debug.Assert a.ItemInArray("BasicMat", db.GetColumnNames(rs0))
    a.printArray db.RecordSetToArray(rs0)
    writeQueryToSheet conn0, sql0, "query"
     
    ' clean up
    conn0.Close
    w.delete_worksheet "query"
    
    '4 sqlwriters
    Debug.Assert sqlWhereInCondition(Array("A", "B", "C"), "column", mssql) = "WHERE column IN ('A','B','C')"
    
        
End Sub

' 1. connections: connect to MSSQL database, connect to excel file as database
Function openMSSQLconn(driver As String, server As String, dbname As String, un As String, pw As String) As Object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    With conn
        .ConnectionString = "DRIVER={" & driver & "};" & _
                            "SERVER=" & server & ";" & _
                            "DATABASE=" & dbname & ";" & _
                            "UID=" & un & ";" & _
                            "PWD=" & pw & ";"
        .Open
    End With
    Set openMSSQLconn = conn
End Function

Function openDBconn(connstr As String)
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    With conn
        .ConnectionString = connstr
        .Open
    End With
    Set openDBconn = conn
End Function

' connect to excel workbook file using ADODB, returns closed connection
Public Function openExcelConn(xls0, Optional connstr As String, Optional columnsAsText0 As Boolean = False, _
Optional printConnStr As Boolean = False) As ADODB.Connection
'On Error GoTo errhandler
errmsg1 = "error in db.openExcelConn "
Dim conn As New ADODB.Connection, fs0 As New filesystemobject: Set fs0 = New filesystemobject
Dim fl0 As File, filePath As String

'get open workbook or xls file as ADO connection
If (TypeName(xls0) = "Workbook") Then
    Dim wb0 As Workbook: Set wb0 = xls0
    wb0.Save
    Set fl0 = fs0.GetFile(wb0.path & "\" & wb0.name)
    GoTo chk_file
        With conn
        If (columnsAsText0 = True) = True Then
            .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Extended Properties").value = "Excel 8.0;HDR=YES;IMEX=1"
        .Open wb0.path & "\" & wb0.name
        Else
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Extended Properties").value = "Excel 8.0"
        .Open wb0.path & "\" & wb0.name
        End If
        End With
ElseIf TypeName(xls0) = "String" Then
    ' check if file name refers to filePath
    filePath = xls0
    fs.pathExist filePath, True
    Set fl0 = fs0.GetFile(filePath)
    GoTo chk_file
ElseIf TypeName(xls0) = "File" Then
    Set fl0 = xls0
    If fs.fileIsExcel(fl0) Then
chk_file:
    'determine excel version
        u.printTemplateString "Excel file type is: `@1`", fl0.Type
        Select Case fl0.Type
        Case "Microsoft Office Excel 97-2003 Worksheet"
                With conn
                If (columnsAsText0 = True) = True Then
                    .Provider = "Microsoft.Jet.OLEDB.4.0"
                .Properties("Extended Properties").value = "Excel 8.0;HDR=YES;IMEX=1"
                .Open fl0.path
                Else
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .Properties("Extended Properties").value = "Excel 8.0"
                .Open fl0.path
                End If
                End With
        Case "Microsoft Office Excel Worksheet"
                With conn
                If (columnsAsText0 = True) = True Then
                    .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Xml;HDR=YES;IMEX=1"
                .Open fl0.path
                Else
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Xml"
                .Open fl0.path
                End If
                End With
        Case "Microsoft Excel Worksheet"
                With conn
                If (columnsAsText0 = True) = True Then
                    .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Xml;HDR=YES;IMEX=1"
                .Open fl0.path
                Else
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Xml"
                .Open fl0.path
                End If
                End With
        Case "Microsoft Office Excel Macro-Enabled Worksheet"
                With conn
                If (columnsAsText0 = True) = True Then
                    .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Macro;HDR=YES;IMEX=1"
                .Open fl0.path
                Else
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Macro"
                .Open fl0.path
                End If
                End With
        Case "Microsoft Excel-werkblad met macro's"
                With conn
                If (columnsAsText0 = True) = True Then
                    .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Macro;HDR=YES;IMEX=1"
                .Open fl0.path
                Else
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Macro"
                .Open fl0.path
                End If
                End With
        Case "excel.exe"
                With conn
                If (columnsAsText0 = True) = True Then
                    .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Macro;HDR=YES;IMEX=1"
                .Open fl0.path
                Else
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Macro"
                .Open fl0.path
                End If
                End With
        Case "Microsoft Excel Macro-Enabled Worksheet", "Microsoft Excel Macro-enabled Worksheet"
                With conn
                If (columnsAsText0 = True) = True Then
                    .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Macro;HDR=YES;IMEX=1"
                .Open fl0.path
                Else
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Properties("Extended Properties").value = "Excel 12.0 Macro"
                .Open fl0.path
                End If
                End With
        Case Else
            GoTo check_type
        End Select
    Else
        GoTo check_type
    End If
Else
check_type:
    errmsg1 = errmsg1 & " check type xls0 is workbook or file. type xls0 is : " & TypeName(xls0)
    GoTo ErrHandler
End If

' get connection string, return to caller
'print connection string
If printConnStr = True Then
   Debug.Print "connected to xls with: "; conn.ConnectionString
End If

' clean-up: set filesystemobject to Nothing
Set fs0 = Nothing
Set openExcelConn = conn

Exit Function
ErrHandler: Err.Raise 1001, "openExcelConn", errmsg0 & " " & errmsg1 & " " & Error(Err)
End Function

Public Sub executeSql(conn0 As ADODB.Connection, sql0 As String, Optional close_connection As Boolean = False)
    'parameter declaration
    Dim rs0 As ADODB.Recordset: Set rs0 = New ADODB.Recordset
    'On Error GoTo errhandler
    errmsg0 = "error in function db.executeSql"
    
    'body
    ' if conn0 is closed then open
    If conn0.State = 0 Then
       conn0.Open
    End If
    
    'open send sql0 to connection conn0
    With rs0
        Debug.Print "queryDB: querying " & sql0
        .Open sql0, conn0, adOpenStatic
    End With
    
    If close_connection Then
      conn0.Close
    End If
    
    'exit procedure
    Exit Sub
    
'end procedure
End Sub

Public Sub executeSqlStatements(conn0 As ADODB.Connection, statements As String, linebr As String)
    ' execute string of statements one-by-one
    Dim sql_statements As New collection
    Set sql_statements = str.stringToCol(statements, linebr)
    
    c = 0
    ' connection management: make sure to close connections on error
    On Error GoTo close_connection
    For Each stat In sql_statements
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
   Exit Sub
End Sub

' 2. CRUD
' CREATE: send insert statement
Public Function sqlInsertStatement(rs0 As Recordset, table_name As String, dbtype0 As dbType) As String
    ' construct insert statement from each record in RecordSet `rs0`
    Dim sql0 As String, sql1 As String

    'body
    'initialize
    c = 0
    With rs0
        Do While .EOF = False
            infieldlist0 = db.commaSeparateFields(rs0)
            invaluelist0 = db.commaSeparateValues(rs0, dbtype0:=dbtype0)
            sql0 = "INSERT INTO " & table_name & " (" & infieldlist0 & ") VALUES (" & invaluelist0 & ");"
            If (c = 0) Then
              sql1 = sql0
            Else
              sql1 = sql1 & db.MSSQL_LINE_BREAK & sql0
            End If
            c = c + 1
            .MoveNext
        Loop
    .MoveFirst ' restore Recordset order
    End With

return_value:
    sqlInsertStatement = sql1
    'exit procedure
    Exit Function
    
End Function

Public Function sqlSetCondition(record As Recordset, set_columns As String, dbtype0 As dbType) As String
'construct set statement

'parameter declaration
Dim fieldNames0 As collection
Set fieldNames0 = str.stringToCol(set_columns, ";")

'body
str0 = "SET " & fieldNames0.item(1) & " = " & xlToDBvalue(record.Fields(fieldNames0.item(1)).value, dbtype0)
If (fieldNames0.Count > 1) = True Then
    For i = 2 To fieldNames0.Count
        str0 = str0 & " , " & fieldNames0.item(i) & " = " & xlToDBvalue(record.Fields(fieldNames0.item(i)).value, dbtype0)
    Next i
End If

return_value:
sqlSetCondition = str0

'exit procedure
Exit Function

'end procedure
End Function


Public Function sqlWhereCondition(record As Recordset, where_columns As String, dbtype0 As dbType, Optional force_string As Boolean = False) As String
'construct set statement

'parameter declaration
Dim fieldNames0 As collection
Set fieldNames0 = str.stringToCol(where_columns, ";")

'body
str0 = "WHERE " & fieldNames0.item(1) & " = " & xlToDBvalue(record.Fields(fieldNames0.item(1)).value, dbtype0, force_string:=force_string)
If (fieldNames0.Count > 1) = True Then
    For i = 2 To fieldNames0.Count
    str0 = str0 & " AND " & fieldNames0.item(i) & " = " & xlToDBvalue(record.Fields(fieldNames0.item(i)).value, dbtype0, force_string:=force_string)
    Next i
End If

return_value:
sqlWhereCondition = str0

'exit procedure
Exit Function

'end procedure
End Function

Public Function sqlWhereInCondition(where_values As Variant, where_column As String, dbtype0 As dbType, Optional force_string As Boolean = False) As String
    Dim sql0 As String, where_values_col As collection
    sql0 = "("
    c = 1
    Set where_values_col = a.as_collection(where_values)
    For Each value0 In where_values_col
        If (c < where_values_col.Count) Then
        sql0 = sql0 & db.xlToDBvalue(value0, dbtype0, force_string:=force_string) & ","
        Else
        sql0 = sql0 & db.xlToDBvalue(value0, dbtype0, force_string:=force_string)
        End If
        c = c + 1
    Next
    sql0 = sql0 & ")"
    sql0 = "WHERE " & where_column & " IN " & sql0
    sqlWhereInCondition = sql0
End Function

Public Function sqlUpdateStatement(rs0 As Recordset, table_name As String, set_columns As String, where_columns As String, dbtype0 As dbType, Optional force_string As Boolean = False) As String
    ' construct insert statement from each record in RecordSet `rs0`
    Dim sql0 As String, sql1 As String, set_condition As String, where_condition As String
    
    'body
    'initialize
    c = 0
    With rs0
        Do While .EOF = False
            set_condition = db.sqlSetCondition(rs0, set_columns, dbtype0)
            where_condition = db.sqlWhereCondition(rs0, where_columns, dbtype0:=dbtype0, force_string:=force_string)
            sql0 = "UPDATE " & table_name & " " & set_condition & " " & where_condition & ";"
            If (c = 0) Then
              sql1 = sql0
            Else
              sql1 = sql1 & db.MSSQL_LINE_BREAK & sql0
            End If
            c = c + 1
            .MoveNext
        Loop
    .MoveFirst ' restore Recordset order
    End With

return_value:
    sqlUpdateStatement = sql1
    'exit procedure
    Exit Function
    
End Function

' READ: send select statement and return recordset
Public Function queryDB(conn0 As ADODB.Connection, sql0 As String, Optional close_connection As Boolean = False, Optional dbg As Boolean = False) As ADODB.Recordset
    'parameter declaration
    Dim rs0 As ADODB.Recordset: Set rs0 = New ADODB.Recordset
    
    'body
    ' if conn0 is closed then open
    If conn0.State = 0 Then
       conn0.Open
    End If
    
    'open send sql0 to connection conn0
    With rs0
        Debug.Print "queryDB: querying " & sql0
        .Open sql0, conn0, adOpenStatic
    End With
    
    'error checking: check if query returns any results
    If (rs0.BOF = False And rs0.EOF = False) = True Then
      'query contains rows
    Else
       'query contains no rows
       errmsg0 = " in queryDB: query " & sql0 & " returns no rows"
       Debug.Print errmsg0
    End If
    
return_value:
    Set queryDB = rs0
    
    If close_connection Then
      conn0.Close
    End If
    
    'exit procedure
    Exit Function
    
End Function

Sub writeQueryToSheet(conn0 As Object, sql0 As String, wsName As String, Optional map_format_column_type As String = "")
    ' This subroutine queries a database using an SQL statement and outputs the results to a specified worksheet.
    ' It creates the target worksheet if it does not exist, and if it does exist, it clears it of all content and formatting.
    '
    ' Parameters:
    ' conn0 - The database connection object
    ' sql0 - The SQL query string
    ' wsname - The name of the worksheet where the results will be output
    
    Dim ws As Worksheet
    Dim rs0 As ADODB.Recordset
    Dim fieldCount As Integer
    Dim i As Integer
    
    ' Check if the worksheet exists, and if so, clear it; otherwise, create it
    Set ws = w.get_or_create_worksheet(wsName, ThisWorkbook, clear:=True)
    
    ' Execute the SQL query
    Set rs0 = db.queryDB(conn0, sql0, False)
    
    ' Check if the recordset is empty and RecordsSetHasFields(rs)
    If db.RecordSetNumberRecords(rs0) = 0 And Not db.RecordsSetHasFields(rs0) Then
       Debug.Print "writeQueryToSheet: Recordset contains no fields"
       Exit Sub
    End If
    arr0 = db.RecordSetToArray(rs0)

    ' paste array to sheet
    a.pasteArray arr0, "A1", ws, ThisWorkbook
    
    ' Autofit the columns for better readability
    ws.Columns.AutoFit
    
    If Len(map_format_column_type) > 0 Then
        r.format_columns ws, map_format_column_type, ThisWorkbook
    End If
    
    ' Clean up
    rs0.Close
    Set rs0 = Nothing
End Sub

Function queryFromWorkbook(sql0 As String, xlsconn As ADODB.Connection) As ADODB.Recordset
    Dim wb0 As Workbook, rs0 As New ADODB.Recordset
    Set wb0 = ThisWorkbook
    Set xlsconn = db.openExcelConn(wb0)
    Set rs0 = db.queryDB(xlsconn, sql0)
    Set queryFromWorkbook = rs0
End Function

Public Sub truncateTable(conn0 As ADODB.Connection, tabname)
    Dim sql0 As String
    sql0 = "TRUNCATE TABLE " & tabname & ";"
    db.executeSql conn0, sql0
End Sub

' 3. Recordset
Public Function GetColumnNames(rs As ADODB.Recordset) As Variant
    Dim colNames() As String
    Dim i As Integer
    
    ' Check if the recordset is open and contains fields
    If Not (rs Is Nothing) And rs.State = adStateOpen And rs.Fields.Count > 0 Then
        ' Resize the array to hold all column names
        ReDim colNames(1 To rs.Fields.Count)
        
        ' Loop through the fields and retrieve the column names
        For i = 1 To rs.Fields.Count
            colNames(i) = rs.Fields(i - 1).name
        Next i
        
        ' Return the column names as a variant array
        GetColumnNames = colNames
    Else
        ' If the recordset is not valid, return an empty array
        GetColumnNames = Array()
    End If
End Function

Public Sub printRecordset(rs0 As Recordset, Optional print_datatypes As Boolean = False, Optional print_field_name As Boolean = True)
'parameter declaration
Dim r0 As Long

' checks
If rs0.State = 0 Then
    Debug.Print errmsg0 & " recordset is closed"
    Exit Sub
End If

' Check if the recordset is empty, if so goto return_array
If rs0.EOF Or rs0.BOF Then
    Debug.Print "RecordSet is empty"
    Exit Sub
End If

'body
With rs0
r0 = 1
    Do While .EOF = False
    str0 = ""
        On Error Resume Next
        If print_field_name = False And r0 = 1 Then
        hdr = util.anything_to_list(.Fields, " ")
        Debug.Print "record num " & hdr
        End If
        
        For c = 1 To .Fields.Count
            val0 = .Fields(c - 1).value
            name0 = ""
            type0 = ""
                If print_datatypes = True Then
                type0 = .Fields(c - 1).Type
                End If
                If print_field_name = True Then
                name0 = .Fields(c - 1).name
                End If
            
            str0 = str0 & " " & name0 & " " & type0 & " " & val0
        Next c
        On Error GoTo 0
    Debug.Print "record "; r0; " "; str0
    r0 = r0 + 1
    .MoveNext
    Loop
.MoveFirst 'restore Recordset
End With

'exit procedure
Exit Sub

'end procedure
End Sub

Function RecordSetNumberRecords(rs0 As ADODB.Recordset) As Long
    ' Check if the recordset is empty
    If rs0.EOF Or rs0.BOF Then
        RecordSetNumberRecords = 0
        Exit Function
    End If
    
    ' Get the number of records in the recordset
    RecordSetNumberRecords = rs0.RecordCount
End Function

Public Function RecordsSetHasFields(rs0 As ADODB.Recordset) As Boolean
     RecordsSetHasFields = (Not a.ArrayIsEmpty(GetColumnNames(rs0)))
End Function

Function RecordSetToArray(rs0 As ADODB.Recordset) As Variant
    ' Get the number of fields in the recordset
    Dim numFields As Integer
    numFields = rs0.Fields.Count
    
    ' Get the number of records in the recordset
    Dim numRecords As Long
    numRecords = RecordSetNumberRecords(rs0)
    
    ' Create the output array
    Dim outputArr() As Variant
    ReDim outputArr(1 To numRecords + 1, 1 To numFields)
    
    ' Add the field names as the first row of the array
    Dim field As ADODB.field
    Dim fieldIndex As Integer
    fieldIndex = 1
    For Each field In rs0.Fields
        outputArr(1, fieldIndex) = field.name
        fieldIndex = fieldIndex + 1
    Next field

    ' Check if the recordset is empty, if so goto return_array
    If rs0.EOF Or rs0.BOF Then
        GoTo return_array
    End If
    
    ' Add the record values to the array
    Dim rowIndex As Long
    rowIndex = 2
    rs0.MoveFirst
    Do Until rs0.EOF
        For fieldIndex = 1 To numFields
            outputArr(rowIndex, fieldIndex) = rs0.Fields(fieldIndex - 1).value
        Next fieldIndex
        rs0.MoveNext
        rowIndex = rowIndex + 1
    Loop
    
    ' Start at beginning of RecordSet
    rs0.MoveFirst
    
    ' Return the output array
return_array:
    RecordSetToArray = outputArr
End Function

' 4. sql writers


' 9. utilities
Public Function commaSeparateFields(rs0 As Recordset) As String
    'parameter declaration
    infields = ""
    'body
    For Each fll In rs0.Fields
        infields = infields & fll.name & "|"
    Next
    infields = left(infields, Len(infields) - 1)
    infields = Replace(Trim(infields), "|", " , ")
    str0 = infields

return_value:
    commaSeparateFields = str0
    
    'exit procedure
    Exit Function
    
'end procedure
End Function

Public Function commaSeparateValues(rs0 As Recordset, Optional xlsvalues As Boolean = True, Optional dbtype0 As dbType = oracledb) As String
    ' convert RecordSet of Name,Value pairs to comma seperated string
    ' parameter declaration
    infields = ""
    'body
    For Each fll In rs0.Fields
    If (xlsvalues = True) = True Then
       fllvalue = xlToDBvalue(fll.value, dbtype0, 0)
    Else
       fllvalue = fll.value
    End If
       infields = infields & fllvalue & "|"
    Next
    infields = left(infields, Len(infields) - 1)
    infields = Replace(Trim(infields), "|", " , ")
    str0 = infields
    
return_value:
    commaSeparateValues = str0
    'exit procedure
    Exit Function

End Function

Public Function xlToDBvalue(xlvalue, dbType As dbType, Optional defvalue As String = "*", Optional force_string As Boolean = False)
' convert excel value with type `VarType` to numeric or string representation that database `dbType` can interprete

'body
'if empty: send NULL to database
If (IsEmpty(xlvalue) = True Or VarType(xlvalue) = vbNull Or VarType(xlvalue) = 1) = True Then
      If (dbType = mysql) Then
      xlvalue = "NULL"
      ElseIf (dbType = oracledb) Then
      xlvalue = "NULL"
      ElseIf (dbType = mssql) Then
      xlvalue = "NULL"
      Else
      MsgBox errmsg0 & " : dbType not recognized"
      End
      End If
'not empty: determine data type
Else
    'excel string type
    If (VarType(xlvalue) = 8) = True Or force_string Then
        Select Case UCase(Trim(xlvalue))
        Case "NULL"
        xlvalue = "NULL"
        Case Else
        xlvalue = "'" & Trim(xlvalue) & "'"
        End Select
    'excel date type or 'excel vartype=7" vbdate
    ElseIf (VarType(xlvalue) = vbDate Or VarType(xlvalue) = 7) = True Then
        If (dbType = mysql) Then
          'the mysql db date format , for oracle use the to_date function: to_date('xlvalue','mm/dd/yyyy')
          xlvalue = "'" & WorksheetFunction.Text(xlvalue, "yyyy-mm-dd") & "'"
        ElseIf (dbType = oracledb) Then
          xlvalue = "to_date(" & "'" & WorksheetFunction.Text(xlvalue, "mm/dd/yyyy") & "','mm/dd/yyyy')"
        ElseIf (dbType = mssql) Then
          xlvalue = "CAST(" & "'" & WorksheetFunction.Text(xlvalue, "yyyy-mm-dd hh:mm:ss") & "' AS DATETIME)"
        Else
          MsgBox errmsg0 & " : dbType not recognized"
          End
        End If
    'excel numeric type
    ElseIf (IsNumeric(xlvalue) = True) = True Then
        xlvalue = xlvalue
    Else
        MsgBox errmsg0 & " : cell data type not recognized " & VarType(xlvalue)
        End
    End If
End If

return_value:
xlToDBvalue = xlvalue

'exit procedure
Exit Function

'end procedure
End Function

Public Function xl_to_xlsdb_value(xlvalue)

'body
'if empty: send NULL to database
If (IsEmpty(xlvalue) = True Or VarType(xlvalue) = vbNull) = True Then
      xlvalue0 = "NULL"
'not empty: determine data type
Else
    'excel string type
    If (VarType(xlvalue) = 8) = True Then
        Select Case UCase(Trim(CStr(xlvalue)))
        Case "NULL"
        xlvalue0 = "NULL"
        Case Else
        xlvalue0 = "'" & Trim(CStr(xlvalue)) & "'"
        End Select
    'excel date type
    ElseIf (IsDate(xlvalue) = True) = True Then
        xlvalue0 = CDbl(DateValue(xlvalue))
    'excel numeric type
    ElseIf (IsNumeric(xlvalue) = True) = True Then
        xlvalue0 = xlvalue
    'excel other type
    Else
        xlvalue0 = "'" & Trim(CStr(xlvalue)) & "'"
    End If
End If

return_value:
xl_to_xlsdb_value = xlvalue0

'exit procedure
Exit Function

End Function


