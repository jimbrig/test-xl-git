Option Explicit

Public Function GetSheetName() As String

  GetSheetName = ActiveSheet.Name
  
End Function

Sub RefreshSheetNames()

  Dim w As Worksheet

  For Each w In ThisWorkbook.Worksheets
  
    w.Activate
    w.Range("A1").Formula = "=GetSheetName()"
    w.Range("A1").Calculate
       
  Next

End Sub