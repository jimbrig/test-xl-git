Option Explicit

Sub CreateNamedRanges()

  Dim wksht As Worksheet
  Dim rngName As String
  
  For Each wksht In ThisWorkbook.Worksheets
    
    If wksht.Name = "ToC" Then GoTo NextWorksheet
    If wksht.Name = "SQL" Then GoTo NextWorksheet
    If wksht.Name = "Input" Then GoTo NextWorksheet
    If wksht.Name = "Data Dictionary" Then GoTo NextWorksheet
    
    With wksht
      .Activate
      .Range("A5").Select
    End With
    
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    rngName = CStr(wksht.Name)
    ThisWorkbook.Names.Add Name:=rngName, RefersTo:=Selection
    
NextWorksheet: Next wksht

End Sub

Sub DeleteNames()
'Update 20140314
Dim rng As Name
For Each rng In Application.ActiveWorkbook.Names
  If InStr(rng.Name, "_FilterDatabase") > 0 Then GoTo NextRng
    If InStr(rng.Name, "_xlfn.CONCAT") > 0 Then GoTo NextRng
    If rng.Name = "_xlfn.SINGLE" Then GoTo NextRng
    If rng.Name = "val_date" Then GoTo NextRng
    If rng.Name = "ToC" Then GoTo NextRng
    rng.Delete
NextRng: Next
End Sub
