' Timer Class
Private startTime As Double

' Start method: records the current time
Public Sub Start()
    startTime = Timer
End Sub

' Stop method: returns the difference in seconds
Public Function StopTimer() As Double
    StopTimer = Timer - startTime
End Function

