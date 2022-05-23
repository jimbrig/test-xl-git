Option Explicit

Const StartRow As Integer = 5

Sub GenerateTableCreate()

  Dim wksht As Worksheet
  Dim xltbl As ListObject
  Dim x, w, t As Long
  Dim colCreate As String
  Dim colname As String
  Dim trailingComma As String
  Dim strOutput As String
  Dim createStmnt, endStmnt, commentStmnt As String
  Dim outCell As Range
  
  Dim prog As ProgressBar
  Set prog = New ProgressBar
  
  OptimizeVBA True
  
  Set outCell = ThisWorkbook.Worksheets("SQL Generation").Range("B4")
  
  w = 0
  t = 0
  
  Call prog.Initialize("SQL Generation Progress:", ThisWorkbook.Sheets.Count - 4)
  
  ' InitProgressBar (ThisWorkbook.Sheets.Count - 4)
    
  For Each wksht In ThisWorkbook.Sheets
  
    If wksht.ListObjects.Count = 0 Then GoTo NextWksht
    
    prog.AddProgress
    Application.Wait Now + TimeValue("00:00:01")
    DoEvents
     
    w = w + 1
    wksht.Activate
    
    For Each xltbl In wksht.ListObjects
      If InStr(xltbl.Name, "tbl_") = 0 Then GoTo NextTbl
      
      t = t + 1
      
      createStmnt = "CREATE TABLE stg." & wksht.Name & " ("
      endStmnt = ");" & vbCrLf
      commentStmnt = "COMMENT ON TABLE stg." & wksht.Name & " IS 'Staging Table for initial loading of source: " & wksht.Name & "';"
    
      strOutput = createStmnt & vbCrLf
    
      For x = 1 To xltbl.ListColumns.Count
        colname = CStr(xltbl.ListColumns(x).Name)
        trailingComma = ","
        If xltbl.ListColumns(x).Name = "valuation_date" Then
          trailingComma = ""
        End If
        colCreate = "  " & """" & colname & """" & " TEXT" & trailingComma & vbCrLf
        'Debug.Print colCreate
        strOutput = strOutput + colCreate
        'Debug.Print strOutput
        
      Next x
      
      strOutput = strOutput + endStmnt + commentStmnt
      Debug.Print strOutput

NextTbl: Next xltbl

outCell.Offset(w, 0).Value = strOutput
outCell.Offset(w, -1).Value = wksht.Name

strOutput = strOutput + vbCrLf + vbCrLf

' ShowProgress (w)

NextWksht: Next wksht

ThisWorkbook.Worksheets("SQL Generation").Activate
Rows("5").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.RowHeight = 15

' CloseProgressBar

OptimizeVBA False
MsgBox ("Done!. Successfully generated SQL DDL for each table")

End Sub

Sub GenerateTableCreate_Life()

  Dim wksht As Worksheet
  Dim xltbl As ListObject
  Dim x, w, t As Long
  Dim colCreate As String
  Dim colname As String
  Dim trailingComma As String
  Dim strOutput As String
  Dim createStmnt, endStmnt, commentStmnt As String
  Dim outCell As Range
  
  Dim prog As ProgressBar
  Set prog = New ProgressBar
  
  OptimizeVBA True
  
  Set outCell = ThisWorkbook.Worksheets("SQL Generation").Range("C4")
  
  w = 0
  t = 0
  
  Call prog.Initialize("SQL Generation Progress:", ThisWorkbook.Sheets.Count - 4)
  
  ' InitProgressBar (ThisWorkbook.Sheets.Count - 4)
    
  For Each wksht In ThisWorkbook.Sheets
  
    If wksht.ListObjects.Count = 0 Then GoTo NextWksht
    
    prog.AddProgress
    Application.Wait Now + TimeValue("00:00:01")
    DoEvents
     
    w = w + 1
    wksht.Activate
    
    For Each xltbl In wksht.ListObjects
      If InStr(xltbl.Name, "tbl_") = 0 Then GoTo NextTbl
      
      t = t + 1
      
      createStmnt = "CREATE TABLE life." & LCase$(wksht.Name) & " ("
      endStmnt = ");" & vbCrLf
      commentStmnt = "COMMENT ON TABLE life." & LCase$(wksht.Name) & " IS 'Life Table for table: " & wksht.Name & "';"
    
      strOutput = createStmnt & vbCrLf
    
      For x = 1 To xltbl.ListColumns.Count
        colname = GetLifeSchemaColName(CStr(xltbl.ListColumns(x).Name))
        trailingComma = ","
        If xltbl.ListColumns(x).Name = "valuation_date" Then
          trailingComma = ""
        End If
        colCreate = "  " & """" & colname & """" & " TEXT" & trailingComma & vbCrLf
        'Debug.Print colCreate
        strOutput = strOutput + colCreate
        'Debug.Print strOutput
        
      Next x
      
      strOutput = strOutput + endStmnt + commentStmnt
      Debug.Print strOutput

NextTbl: Next xltbl

outCell.Offset(w, 0).Value = strOutput
outCell.Offset(w, -1).Value = wksht.Name

strOutput = strOutput + vbCrLf + vbCrLf

' ShowProgress (w)

NextWksht: Next wksht

ThisWorkbook.Worksheets("SQL Generation").Activate
Rows("5").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.RowHeight = 15

' CloseProgressBar

OptimizeVBA False
MsgBox ("Done!. Successfully generated SQL DDL for each table")

End Sub



Public Function GetLifeSchemaColName(colname As String)

  Dim out As String
  
  out = LCase$(colname)
  
  out = Replace( _
  Replace( _
    Replace( _
      Replace( _
        Replace( _
          Replace(out, " ", "_"), _
            "_-_", "_"), _
          "/", ""), _
        "&", ""), _
      "pre-mat", "pre_mat"), _
    "post-mat", "post_mat")
    
  GetLifeSchemaColName = out
     

End Function



