'test() => Executes the test subroutine demonstrating date and time functions
'first_day_isoweek(weeknum As Integer, year As Integer) => Returns the first day of the specified ISO week number and year
'vdmi_get_day_of_week(startDate As Date, day As String) => Gets day `day` of week(startDate) as Date.
'vdmi_last_datetime_of_day(date0) => Returns the last datetime (23:00) of the given date
'vdmi_first_datetime_of_day(date0) => Returns the first datetime (09:00) of the given date
'get_datetime_value(dateString, Optional timeString) => Converts a date string and optional time string to a datetime value
'add_hours(datetime_value As Double, h) => Adds a specified number of hours to a datetime value
'format_datetime(date_value, Optional str_format) => Formats a date value into a string using the specified or default format
'set_date_timepart(date0 As Date, Optional timepart) => Sets the time part of a date to the specified or default time

Sub test_dt_functions()
    Dim date0 As Date
    datestr = "28/07/2023 23:00:00"
    date0 = datestr
    thisyear = year(Now())
    nextyear = year(Now()) + 1
    Debug.Print date0
    
    Debug.Assert dt.vdmi_get_day_of_week("2024-02-05", "ma") = "2024-02-05" 'date of "ma" in the week of 2024-02-05
    Debug.Assert dt.vdmi_get_day_of_week("2024-02-05", "di") = "2024-02-06" 'date of "di" in the week of 2024-02-05
    
    ' Assertion test for the formatDateVDMI function
    Debug.Print formatDateVDMI(#2/11/2024 8:29:00 PM#)
    Debug.Assert formatDateVDMI(#2/11/2024 8:29:00 PM#) = "Zon 11 20:29"
    Debug.Assert formatDateVDMI(#2/12/2024 8:29:00 PM#) = "Ma 12 20:29"
    Debug.Assert formatDateVDMI(#3/1/2024 8:29:00 PM#) = "Vrij 01 20:29"
    
    Debug.Assert determine_year_based_on_weeknum(1) = nextyear
    Debug.Assert determine_year_based_on_weeknum(54) = thisyear
    
End Sub

' date and time functions
Function first_day_isoweek(weeknum As Integer, year As Integer) As Date
    Dim jan1 As Date
    Dim jan1Weekday As Integer
    Dim daysToMonday As Integer
    Dim firstDay As Date
    
    ' Get January 1st of the specified year
    jan1 = DateSerial(year, 1, 1)
    
    ' Get the weekday of January 1st (1 = Sunday, 2 = Monday, ..., 7 = Saturday)
    jan1Weekday = Weekday(jan1, vbMonday)
    
    ' Calculate the number of days to previous Monday on Jan 1st: this is the start of the ISO week
    daysToMonday = jan1Weekday
    
    ' Calculate the first day of the ISO week
    firstDay = DateAdd("d", (weeknum - 1) * 7 - daysToMonday + 1, jan1)
    
    ' Return the first day of the ISO week
    first_day_isoweek = firstDay
    
End Function

Function vdmi_get_day_of_week(startDate As Date, day As String) As Date
    Dim resultDate As Date
    
    ' Convert the day abbreviation to lowercase for case-insensitive comparison
    day = LCase(day)
    
    ' Determine the number of days to add based on the day abbreviation
    Select Case day
        Case "ma" ' Monday
            resultDate = startDate + 0
        Case "di" ' Tuesday
            resultDate = startDate + 1
        Case "woe" ' Wednesday
            resultDate = startDate + 2
        Case "do" ' Thursday
            resultDate = startDate + 3
        Case "vrij" ' Friday
            resultDate = startDate + 4
        Case "ma2" ' Monday
            resultDate = startDate + 7
        Case "di2" ' Tuesday
            resultDate = startDate + 8
        Case "woe2" ' Wednesday
            resultDate = startDate + 9
        Case "do2" ' Thursday
            resultDate = startDate + 10
        Case "vrij2" ' Friday
            resultDate = startDate + 11
        Case Else
            ' Invalid day abbreviation, return the start date
            Err.Raise 2003, Description:="Invalid day abbreviation: " & day
    End Select
    
    ' Return the resulting date
    vdmi_get_day_of_week = resultDate
End Function

Function vdmi_last_datetime_of_day(date0) As Double
vdmi_last_datetime_of_day = get_datetime_value(date0, "23:00")
End Function

Function vdmi_first_datetime_of_day(date0) As Double
vdmi_first_datetime_of_day = get_datetime_value(date0, "09:00")
End Function

Function formatDateVDMI(date0 As Variant, Optional wkDayNameIsLong As Boolean = False, Optional dbg As Boolean = False) As String
    ' This function formats the date as per VDMI requirements.
    ' If date0 is of vbDate type, it returns a formatted date string in the format:
    ' "day of week" "day of month" hh:mm
    ' Otherwise, it returns date0.
    ' "day of week" is the capitalized day name in Dutch (Monday => Maandag, etc.).
    ' "day of month" is the zero-padded day index of the month of date0.
    
    Dim dayNames As Variant
    If wkDayNameIsLong Then
       dayNames = Array("Zondag", "Maandag", "Dinsdag", "Woensdag", "Donderdag", "Vrijdag", "Zaterdag")
    Else
       dayNames = Array("Zon", "Ma", "Di", "Wo", "Do", "Vrij", "Zat")
    End If
    
    If VarType(date0) = vbDate Or VarType(date0) = vbDouble Then
        Dim date1 As Date
        If VarType(date0) = vbDouble Then
           ' type to cast to proper date
           date1 = date0
        Else
           date1 = date0
        End If
    
        Dim dayOfWeek As String
        Dim dayOfMonth As String
        Dim formattedTime As String
        
        ' Get the Dutch day name
        dayOfWeek = dayNames(Weekday(date1, vbSunday) - 1)
        
        ' Get the zero-padded day of the month
        dayOfMonth = Format(day(date1), "00")
        
        ' Get the formatted time
        formattedTime = Format(date1, "hh:mm")
        
        ' Combine the parts to create the formatted date string
        formatDateVDMI = dayOfWeek & " " & dayOfMonth & " " & formattedTime
    Else
        ' If date0 is not a date, return it as is
        If dbg Then
           Debug.Print "input " & date0 & " is not vbDate but " & VarType(date0)
        End If
        formatDateVDMI = date0
    End If
End Function

Function formatDateVDMILong(date0 As Variant) As String
    formatDateVDMILong = dt.formatDateVDMI(date0, True, dbg:=False)
End Function

Function formatDateVDMIShort(date0 As Variant) As String
    formatDateVDMIShort = dt.formatDateVDMI(date0, False, dbg:=False)
End Function

Function get_datetime_value(dateString, Optional timeString = "00:00") As Double
Dim date_value As Double
If VarType(dateString) = vbDate Then
   dateString = Format(dateString, "yyyy-mm-dd")
   date_value = CDbl(dateValue(dateString))
ElseIf VarType(dateString) = vbDouble Then
   date_value = dateString
ElseIf VarType(dateString) = vbInteger Then
   date_value = dateString
End If
get_datetime_value = date_value + CDbl(TimeValue(timeString))
End Function

Function add_hours(datetime_value As Double, h)
add_hours = datetime_value + h / 24
End Function

Function format_datetime(date_value, Optional str_format = "yyyy-mm-dd hh:mm") As String
format_datetime = Format(date_value, str_format)
End Function

'set the timepart of a date
Function set_date_timepart(date0 As Date, Optional timepart As String = "00:00") As Date
    ' Use the DateValue function to extract the date part of the date argument
    Dim datePart As Date
    datePart = dateValue(date0)
    
    ' Use the TimeValue function to convert the timepart argument to a time
    timepart = TimeValue(timepart)
    
    ' Combine the date and time to create a new Date value
    set_date_timepart = datePart + timepart
End Function

' WEEKS
Function determine_year_based_on_weeknum(weeknum As Integer) As Integer
    ' Determines the year based on the given week number.
    ' If the current week number is less than or equal to the given week number,
    ' it sets the year as the current year (value of the year today).
    ' If the current week number is greater than the given week number,
    ' it sets the year as the next year (value of the year today + 1).
    '
    ' Parameters:
    ' weeknum - The week number to compare with the current week number.
    '
    ' Returns:
    ' An integer representing the determined year.

    Dim currentWeeknum As Integer
    Dim currentYear As Integer

    ' Get the current week number and year
    currentWeeknum = datePart("ww", Date, vbMonday, vbFirstJan1)
    currentYear = year(Date)

    ' Determine the year based on the comparison of week numbers
    If currentWeeknum <= weeknum Then
        determine_year_based_on_weeknum = currentYear
    Else
        determine_year_based_on_weeknum = currentYear + 1
    End If
End Function
