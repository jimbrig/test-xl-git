Option Explicit

' Written by Philip Treacy, June 2016
' http://www.myonlinetraininghub.com/excel-progress-bar-for-vba

Sub Using_Collection_Loops_ProgressBar()
  
  Dim lngNumValues() As Long
  Dim strBorder As String
  Dim lngMin As Long, lngMax As Long
  Dim rng As Range, rngInput As Range
  Dim col As Collection
  Dim i As Long, j As Long, k As Long
  Dim varDummy As Variant
  
  
  '---------------------------------
  ' CHANGE THIS PART TO DEFINE RUNS
  '---------------------------------
  
  ReDim lngNumValues(1 To 5) As Long
  lngNumValues(1) = 10
  lngNumValues(2) = 100
  lngNumValues(3) = 1000
  lngNumValues(4) = 10000
  lngNumValues(5) = 50000
  
  '---------------------------------
  ' THIS IS WHERE THE MAGIC HAPPENS
  '---------------------------------
  
  ' Set up output table headers
  strBorder = String(92, "-")
  Debug.Print strBorder
  
    k = 1
          
      'Set up random numbers
      lngMin = 1
      lngMax = lngNumValues(5)
      
    'Initilaize the progress bar width
    InitProgressBar (lngMax)

              
      Cells.Clear
      Set rngInput = Range("A1").Resize(lngMax, 1)
      rngInput.Formula = "=RANDBETWEEN(" & lngMin & "," & lngMax & ")"
           
      Set col = New Collection
      
      ' First add all items to collection (ignore errors when duplicates are added)
      For Each rng In rngInput
        On Error Resume Next
          col.Add rng.Value2, CStr(rng.Value2)
        On Error GoTo 0
      Next rng
      
      
      
      ' Loop through every possible nubmer in range (error returned when number not found)
      On Error Resume Next
        For j = lngMin To lngMax
          varDummy = col(CStr(j))
          If Not Err.Number = 0 Then
            Debug.Print j
            Err.Clear
          End If
            'Must DoEvents to allow code to update bar and show it
            DoEvents
            
            ShowProgress (j)
            
        Next j
      On Error Resume Next
      

        CloseProgressBar

    
End Sub


Sub TestProgressBar()

    Dim j As Long
    Dim max As Long

    max = 50000

    'Initilaize the progress bar width
    InitProgressBar (max)

    For j = 1 To max
    
        DoEvents
        ShowProgress (j)
        
    Next j
    
    CloseProgressBar

End Sub

