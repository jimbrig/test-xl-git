Option Explicit

Sub ExportCSV_InputsOnly()

  Application.ScreenUpdating = False
  Application.AlertBeforeOverwriting = False
  Application.DisplayAlerts = False

  Dim WB As Workbook
  Dim NewWB As Workbook
  Dim CSVFile As String
    
  Set WB = ThisWorkbook
  Dim rng As Name
  Dim wksht As Worksheet
  Dim tblName As String
  
  For Each rng In WB.Names
    
    If InStr(rng.Name, "_FilterDatabase") > 0 Then GoTo NextRng
    If rng.Name = "_xlfn.SINGLE" Then GoTo NextRng
    If rng.Name = "val_date" Then GoTo NextRng
    If rng.Name = "ToC" Then GoTo NextRng
    If InStr(rng.Name, "Input_") > 0 Then
      Debug.Print "Name: " & rng.Name & "; Value: " & rng.Value
      rng.RefersToRange.Copy
      CSVFile = ThisWorkbook.Path & Application.PathSeparator & "Inputs" & Application.PathSeparator & Replace(rng.Name, "Input_", "") & ".csv"
      Debug.Print "Exporting Range " & rng.Name & " to CSV Path: " & CSVFile & "."
      
      Set NewWB = Workbooks.Add
      With NewWB
        ActiveSheet.PasteSpecial
        .SaveAs fileName:=CSVFile, FileFormat:=xlCSV, CreateBackup:=True
        .Close
      End With
      
    End If
    
NextRng: Next rng

  Application.ScreenUpdating = True
  Application.AlertBeforeOverwriting = True
  Application.DisplayAlerts = True
  
  MsgBox "Done!", vbInformation
    
End Sub

Sub ExportCSV()

  Dim prog As ProgressBar
  Set prog = New ProgressBar

  Dim WB As Workbook
  Dim NewWB As Workbook
  Dim CSVFile As String
  Dim CSVFilePath As String
    
  Set WB = ThisWorkbook
  Dim rng As Name
  Dim wksht As Worksheet
  Dim tblName As String
  
  Dim FSO As New FileSystemObject
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  OptimizeVBA True
  Application.DisplayAlerts = False
  
  Call prog.Initialize("Exporting CSV Files:", WB.Names.Count - 6)
  
  On Error Resume Next
  
  For Each rng In WB.Names
    
    If InStr(rng.Name, "_FilterDatabase") > 0 Then GoTo NextRng
    If rng.Name = "_xlfn.SINGLE" Then GoTo NextRng
    If rng.Name = "_xlfn.VALUETOTEXT" Then GoTo NextRng
    If rng.Name = "val_date" Then GoTo NextRng
    If rng.Name = "ToC" Then GoTo NextRng
    If rng.Name = "SQL" Then GoTo NextRng
    
    prog.AddProgress
    Application.Wait Now + TimeValue("00:00:01")
    DoEvents
    
    Debug.Print "Name: " & rng.Name & "; Value: " & rng.Value
    
    CSVFilePath = ThisWorkbook.Path & Application.PathSeparator & "CSV" & Application.PathSeparator
    
    If FSO.FolderExists(CSVFilePath) = False Then
      FSO.CreateFolder CSVFilePath
    End If

    rng.RefersToRange.Copy
    
    CSVFile = CSVFilePath & rng.Name & ".csv"
    
    Set NewWB = Workbooks.Add
    
    With NewWB
      ActiveSheet.PasteSpecial
      .SaveAs fileName:=CSVFile, FileFormat:=xlCSV, CreateBackup:=True, ConflictResolution:=2
      .Close
    End With
    
NextRng: Next rng

  On Error GoTo 0

  Set FSO = Nothing
  Application.DisplayAlerts = True
  OptimizeVBA False
  MsgBox ("Done!. Successfully exported CSV files to the CSV directory.")
  
End Sub

Sub ExportSQL()

  Dim SQLRange As Name
  Dim i As Long
  Dim sql, fileName As String
  Dim NewFile As TextStream
  Dim FSO As New FileSystemObject
  Dim NewFilePath As String
  
  Dim prog As ProgressBar
  Set prog = New ProgressBar
  
  OptimizeVBA True
    
  Set FSO = CreateObject("Scripting.FileSystemObject")
  Set SQLRange = ThisWorkbook.Names("SQL")
  
  Call prog.Initialize("Exporting SQL Files:", SQLRange.RefersToRange.Rows.Count)
  
  For i = 1 To SQLRange.RefersToRange.Rows.Count
  
    prog.AddProgress
    Application.Wait Now + TimeValue("00:00:01")
    DoEvents
    
    NewFilePath = ThisWorkbook.Path & Application.PathSeparator & "SQL" & Application.PathSeparator
    
    If FSO.FolderExists(NewFilePath) = False Then
      FSO.CreateFolder NewFilePath
    End If
    
    fileName = NewFilePath & ThisWorkbook.Sheets("SQL Generation").Range("A4").Offset(i, 0).Value & ".sql"
    sql = ThisWorkbook.Sheets("SQL Generation").Range("A4").Offset(i, 1).Value
    Set NewFile = FSO.CreateTextFile(fileName, True)
    NewFile.Write (sql)
    NewFile.Close
  Next
  
  Set FSO = Nothing
  
  OptimizeVBA False
  
  MsgBox ("Done!. Successfully exported SQL scripts to the SQL directory.")
    
End Sub
