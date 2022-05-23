Option Explicit

Sub InitProgressBar(MaxValue As Long)

    With ProgressBarSimple
    
        .BarColor.Tag = .BarBorder.Width / MaxValue
        .BarColor.Width = 0
        .ProgressText = ""
        .Show vbModeless
        
    End With
    
End Sub

Sub CloseProgressBar()

    Unload ProgressBarSimple
    
End Sub
Sub ShowProgress(progress As Long)

    With ProgressBarSimple

        'Round Up
        .BarColor.Width = Round(.BarColor.Tag * progress, 0)
        .ProgressText.Caption = Round((.BarColor.Width / .BarBorder.Width * 100), 0) & "% complete"
    
    End With

End Sub

