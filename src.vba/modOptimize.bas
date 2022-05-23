Option Explicit
Dim prevCalc, prevEvents, prevScreen, prevPageBreaks
Dim execTimer As HighResPerformanceTimer, currTime As Double
#If VBA7 Then
    Private Declare PtrSafe Function CreateThread Lib "kernel32" (ByVal LpThreadAttributes As Long, ByVal DwStackSize As Long, ByVal LpStartAddress As Long, ByVal LpParameter As Long, ByVal dwCreationFlags As Long, ByRef LpThreadld As Long) As Long
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal HANDLE As Long) As Long
#Else
    Private Declare Function CreateThread Lib "kernel32" (ByVal LpThreadAttributes As Long, ByVal DwStackSize As Long, ByVal LpStartAddress As Long, ByVal LpParameter As Long, ByVal dwCreationFlags As Long, ByRef LpThreadld As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal HANDLE As Long) As Long
#End If

Sub OptimizeVBA(isOn As Boolean)
    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (isOn)
    Application.ScreenUpdating = Not (isOn)
    ActiveSheet.DisplayPageBreaks = Not (isOn)
End Sub

Sub OptimizeOn()
    'Optimize VBA exectution
    prevCalc = Application.Calculation: Application.Calculation = xlCalculationManual
    prevEvents = Application.EnableEvents: Application.EnableEvents = False
    prevScreen = Application.ScreenUpdating: Application.ScreenUpdating = False
    prevPageBreaks = ActiveSheet.DisplayPageBreaks: ActiveSheet.DisplayPageBreaks = False
End Sub

Sub OptimizeOff()
    'Turn off VBA optimization
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    ActiveSheet.DisplayPageBreaks = prevPageBreaks
End Sub

Sub StartHighResPerformanceTimer()
    Set execTimer = New HighResPerformanceTimer
    execTimer.StartCounter
End Sub
Function StopHighResPerformanceTimer() As Double
    StopHighResPerformanceTimer = execTimer.TimeElapsed
    Set execTimer = Nothing
End Function
'*****Low Res (seconds)*****
Sub StartLowResPerformanceTimer()
    currTime = Timer
End Sub
Function StopLowResPerformanceTimer() As Double
    StopLowResPerformanceTimer = Timer - currTime
End Function



