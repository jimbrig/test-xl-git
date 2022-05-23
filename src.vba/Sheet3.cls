Option Explicit




Private Sub BtnCopySQL_Click()
  Range("Z3").Copy
  MsgBox "Copied to clipboard!"
End Sub

Private Sub ComboBoxTableSelection_Change()
  
  TxtBoxDisplaySQL.Text = Range("Z3").Value

End Sub

