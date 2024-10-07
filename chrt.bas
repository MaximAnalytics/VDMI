Sub save_chart_as_pdf(chrt_name As String, pdf_file_name As String, ws As Worksheet)
    Dim chart As chart
    
    'Get the chart object by its name or index
    Set chart = ws.ChartObjects(chrt_name).chart
    'Or: Set chart = ActiveSheet.ChartObjects(1).Chart
    
    'Export the chart as a PDF file
    chart.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdf_file_name
    
    'Display a message box indicating that the PDF file has been saved
    MsgBox "The chart has been saved as a PDF file: " & pdf_file_name, vbInformation, "Export Chart"


End Sub

Sub CopyChartToNewWorkbook(ws0 As Worksheet, wb0 As Workbook, chart_name As String, chart_pdf_file As String)
    Dim myChart As chart
    Dim myWorkbook As Workbook
    Dim myWorksheet As Worksheet
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    
    Set wb0 = r.get_default_wb(wb0)
    
    ' Set the chart object
    Set myChart = ws0.ChartObjects(chart_name).chart
    
    ' Copy the chart object to a new worksheet in a new workbook
    Set newWorkbook = Workbooks.Add
    Set newWorksheet = newWorkbook.Sheets("Sheet1")
    myChart.ChartArea.Copy
    newWorksheet.Paste Destination:=newWorksheet.Range("A1")
    
    ' Print the chart as PDF
'    With newWorksheet.PageSetup
'        .Orientation = xlLandscape
'        .Zoom = False
'        .FitToPagesWide = 1
'        .FitToPagesTall = 1
'    End With
'    newWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=main.CHART_PDF_FILE, Quality:=xlQualityStandard
'
    ' Close the new workbook without saving
    newWorkbook.Close SaveChanges:=False
    
    ' Activate the original workbook and worksheet
    wb0.Activate
    ws0.Activate
End Sub


Sub PrintChartAsPDF(ws0 As Worksheet, wb0 As Workbook, chart_name As String, chart_pdf_file As String)

    Call chrt.CopyChartToNewWorkbook(ws0, wb0, chart_name, chart_pdf_file)
    
    ' Show button with message
    MsgBox "figuur gekopieerd naar " & chart_pdf_file, vbInformation
    
    Exit Sub
    
End Sub


Sub init_chart()
    Const CHART_ROW_START As Long = 8
    Dim CHART_COL_START As Long
    
    Dim ws As Worksheet
    Dim cht As ChartObject
    
    ' Set the worksheet object
    Set ws = ThisWorkbook.Worksheets(INPUT_DATA_SHEET)
    
    ' Set the chart start column
    CHART_COL_START = UBound(Split(INPUT_DATA_HEADER, ",")) + 3
    
    ' Delete the existing chart if it exists
    For Each cht In ws.ChartObjects
        If cht.name = chart_name Then
            cht.Delete
            Exit For
        End If
    Next cht
    
    ' Create the chart
    Set cht = ws.ChartObjects.Add(left:=ws.Cells(CHART_ROW_START, CHART_COL_START).left, _
                                   Top:=ws.Cells(CHART_ROW_START, CHART_COL_START).Top, _
                                   Width:=CHART_WIDTH, Height:=CHART_HEIGHT)
    cht.name = chart_name
    cht.chart.ChartType = xlBarStacked
    cht.chart.HasLegend = False

End Sub
'
'Sub update_chart()
'
'    Dim wb0 As Workbook
'    Dim gantt_chart As ChartObject
'    Set wb0 = ThisWorkbook '""
'    Set gantt_chart = wb0.Worksheets(main.INPUT_DATA_SHEET).ChartObjects(chart_name)
'
'    ' Clear the existing series from the chart
'    Do While gantt_chart.chart.SeriesCollection.count > 0
'        gantt_chart.chart.SeriesCollection(1).Delete
'    Loop
'
'    ' Define the range for the start dates and durations
'    Dim startdate_range As Range
'    Dim duration_range As Range
'    Dim label_range As Range
'    Dim rng0 As Range
'
'    ' range is static but should be dynamic, get the new range if current named range INPUT_CALC_RNG_NAME is changed
'    Set rng0 = wb0.Names(INPUT_CALC_RNG_NAME).RefersToRange
'    Set startdate_range = rng0.columns(STARTDATE_COLUMN_INDEX).Offset(1, 0).Resize(rng0.Rows.count - 1, 1)
'    Set duration_range = rng0.columns(DURATION_COLUMN_INDEX).Offset(1, 0).Resize(rng0.Rows.count - 1, 1)
'    Set id_range = rng0.columns(1).Offset(1, 0).Resize(rng0.Rows.count - 1, 1)
'
'    ' Add the start date and duration series to the chart
'    gantt_chart.chart.SeriesCollection.NewSeries
'    gantt_chart.chart.FullSeriesCollection(1).name = "=calculated_data!" & rng0.columns(STARTDATE_COLUMN_INDEX).Cells(1, 1).address
'    gantt_chart.chart.FullSeriesCollection(1).values = "=calculated_data!" & startdate_range.address '"=calculated_data!$B$2:$B$3"
'
'    gantt_chart.chart.SeriesCollection.NewSeries
'    gantt_chart.chart.FullSeriesCollection(2).name = "=calculated_data!" & rng0.columns(DURATION_COLUMN_INDEX).Cells(1, 1).address
'    gantt_chart.chart.FullSeriesCollection(2).values = "=calculated_data!" & duration_range.address
'    gantt_chart.chart.FullSeriesCollection(2).XValues = "=calculated_data!" & id_range.address
'
'    ' Update the chart formatting in case of changed parameters
'    Call update_chart_formatting
'
'End Sub
'
'
'Sub update_chart_formatting()
'    ' Set the chart x-axis minimum value
'    Call chart_set_axes
'
'    ' Set the chart title
'    Call chart_set_title
'
'    ' Set the chart series fill
'    Call chart_set_fill
'
'    'Activate the chart with the given name
'    ThisWorkbook.Worksheets(INPUT_DATA_SHEET).ChartObjects(chart_name).Activate
'    ' chart set x axis formatting
'    ActiveChart.Axes(xlvalue).Select
'    Selection.MinorTickMark = xlInside
'    Selection.TickLabels.NumberFormat = "dd/mm/yyyy"
'    ActiveChart.Axes(xlvalue).MajorUnit = 1
'    ActiveChart.Axes(xlvalue).MinorUnit = 0.125
'
'End Sub
'
'
'Sub chart_set_data_labels()
'    'Get the input data column range
'    Dim calcDataLabelsColumn As Range
'    Set calcDataLabelsColumn = get_calculated_data_column(LABEL_COLUMN_INDEX)
'
'    'Activate the chart with the given name
'    ThisWorkbook.Worksheets(INPUT_DATA_SHEET).ChartObjects(chart_name).Activate
'
'    'Set the data labels of the active chart
'    ActiveChart.FullSeriesCollection(2).Select
'    ActiveChart.FullSeriesCollection(2).ApplyDataLabels
'    ActiveChart.FullSeriesCollection(2).DataLabels.Select
'    ActiveChart.SeriesCollection(2).DataLabels.Format.TextFrame2.TextRange. _
'        InsertChartField msoChartFieldRange, "=input_data!" & calcDataLabelsColumn.address, 0
'    Selection.ShowRange = True
'    Selection.ShowValue = False
'End Sub
'
'Sub chart_set_axes()
'    'Activate the chart with the given name
'    ThisWorkbook.Worksheets(INPUT_DATA_SHEET).ChartObjects(chart_name).Activate
'
'    'Set the minimum value for the x-axis to the value of the named range
'    Dim dateValue As Double
'    dateValue = Range(STARTDATE_NAME_RANGE).value
'    ActiveChart.Axes(xlvalue).Select
'    ActiveChart.Axes(xlvalue).MinimumScale = dateValue
'
'    ' categories in reverse order
'    ActiveChart.Axes(xlCategory).Select
'    ActiveChart.Axes(xlCategory).ReversePlotOrder = True
'End Sub
'
'Public Sub chart_set_title()
'    'Activate the chart with the given name
'    Dim cht As ChartObject
'    Set cht = ThisWorkbook.Worksheets(INPUT_DATA_SHEET).ChartObjects(chart_name)
'
'    Dim chartName As String
'    chartName = Range(CHART_NAME_RANGE).value
'    cht.chart.SetElement (msoElementChartTitleAboveChart)
'    cht.chart.ChartTitle.Text = chartName
'    'Selection.Format.TextFrame2.TextRange.Characters.Text = chartName
'End Sub
'
'Sub chart_set_fill()
'    'Activate the chart with the given name
'    ThisWorkbook.Worksheets(INPUT_DATA_SHEET).ChartObjects(chart_name).Activate
'    'Set the fill of the first series to none
'    ActiveChart.SeriesCollection(1).Format.Fill.Visible = msoFalse
'End Sub




