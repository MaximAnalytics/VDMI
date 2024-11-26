' maths
Function Round_to_Nearest_Fraction(N As Double, f As Double) As Double
    Dim roundedValue As Double
    roundedValue = WorksheetFunction.Round(N / f, 0) * f
    Round_to_Nearest_Fraction = roundedValue
End Function

Function Round_Up_to_Nearest_Fraction(N As Double, f As Double) As Double
    Dim roundedValue As Double
    roundedValue = WorksheetFunction.Ceiling(N / f, 1) * f
    Round_Up_to_Nearest_Fraction = roundedValue
End Function

Function round_up_to_nearest_quarter(N As Double) As Double
    Dim quarter_fraction As Double
    quarter_fraction = 1 / (24 * 4)
    round_up_to_nearest_quarter = Round_Up_to_Nearest_Fraction(N, quarter_fraction)
End Function

Function gte_dbl(x As Double, y As Double) As Boolean
  gte_dbl = Round(x, 10) >= Round(y, 10)
End Function

Function lte_dbl(x As Double, y As Double) As Boolean
  lte_dbl = Round(x, 10) <= Round(y, 10)
End Function

Function IsEven(N As Integer) As Boolean
    If N Mod 2 = 0 Then
        IsEven = True
    Else
        IsEven = False
    End If
End Function

Function getRatio(numerator As Variant, denominator As Variant, Optional defvalue As Variant = 0) As Variant
    ' This function returns the ratio of numerator to denominator.
    ' If an error occurs during division (e.g., division by zero), it returns the default value.
    '
    ' Parameters:
    ' numerator: The numerator of the ratio.
    ' denominator: The denominator of the ratio.
    ' defvalue: The default value to return in case of an error (default is 0).
    '
    ' Returns:
    ' The result of the division or the default value if an error occurs.
    
    On Error GoTo ErrorHandler ' Enable error handling
    
    ' Perform division
    getRatio = numerator / denominator
    Exit Function ' Exit before reaching the error handler
    
ErrorHandler:
    ' If an error occurs, return the default value
    getRatio = defvalue
End Function

