
Option Explicit

' Written by Philip Treacy, June 2016
' http://www.myonlinetraininghub.com/excel-progress-bar-for-vba


Private Sub CancelButton_Click()

    ShutDownProgressbar

End Sub

Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
       
    If CloseMode = vbFormControlMenu Then
        
        ShutDownProgressbar
    
    End If
    
End Sub

Private Sub ShutDownProgressbar()

    Unload ProgressBar
    End

End Sub