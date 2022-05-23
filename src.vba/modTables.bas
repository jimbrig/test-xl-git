Option Explicit

Sub ListTables()
'Updated by Extendoffice 20180503
    Dim xTable As ListObject
    Dim xSheet As Worksheet
    Dim i As Long
    i = -1
    Sheets.Add.Name = "Table Name"
    For Each xSheet In Worksheets
        For Each xTable In xSheet.ListObjects
            i = i + 1
            Sheets("Table Name").Range("A1").Offset(i).Value = xTable.Name
        Next xTable
    Next
End Sub
Sub RangeToTable()

  Dim WB As Workbook
  Set WB = ThisWorkbook
  Dim rng As Name
  Dim wksht As Variant
  Dim wksht_name As String
  Dim tblName As String
  
  For Each rng In WB.Names
    
    If InStr(rng.Name, "_FilterDatabase") > 0 Then GoTo NextRng
    If InStr(rng.Name, "_xlfn.CONCAT") > 0 Then GoTo NextRng
    If rng.Name = "_xlfn.SINGLE" Then GoTo NextRng
    If rng.Name = "val_date" Then GoTo NextRng
    If rng.Name = "ToC" Then GoTo NextRng
  
    wksht_name = rng.RefersToRange.Worksheet.Name
    Set wksht = ThisWorkbook.Worksheets(wksht_name)
    
    If wksht.ListObjects.Count > 0 Then GoTo NextRng
      
    Debug.Print "Name: " & rng.Name & "; Value: " & rng.Value
    wksht.ListObjects.Add(xlSrcRange, Range(rng.Name)).Name = "tbl_" & wksht.Name
    
NextRng: Next rng

End Sub

