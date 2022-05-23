Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr) 'MS Office 64 Bit
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long) 'MS Office 32 Bit
#End If
Sub WaitUntilTime(waitTime As Date)
    Application.Wait waitTime
End Sub
Sub WaitForSeconds(seconds As Long)
    Application.Wait DateAdd("s", seconds, Now)
End Sub
Sub WaitForMinutes(minutes As Long)
    Application.Wait DateAdd("n", minutes, Now)
End Sub
Sub WaitForHours(hours As Long)
    Application.Wait DateAdd("h", hours, Now)
End Sub
Sub RefreshScreen()
'Refresh screen when using Application.ScreenUpdating = False
    DoEvents
End Sub
Sub RunMacroAtTime(timeToRun As Date, nameOfMacroToRun As String)
    Application.OnTime timeToRun, nameOfMacroToRun
End Sub
Sub RunMacroInSeconds(seconds As Long, nameOfMacroToRun As String)
    Application.OnTime DateAdd("s", seconds, Now), nameOfMacroToRun
End Sub
Sub RunMacroInMinutes(minutes As Long, nameOfMacroToRun As String)
    Application.OnTime DateAdd("n", minutes, Now), nameOfMacroToRun
End Sub
Sub RunMacroInHours(hours As Long, nameOfMacroToRun As String)
    Application.OnTime DateAdd("h", hours, Now), nameOfMacroToRun
End Sub


