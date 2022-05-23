Option Explicit

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

    #If VBA7 Then
      Declare PtrSafe Function apiGetVersionEx Lib "kernel32" _
      Alias "GetVersionExA" _
      (lpVersionInformation As Any) _
      As Long
    #Else
      Declare Function apiGetVersionEx Lib "kernel32" _
      Alias "GetVersionExA" _
      (lpVersionInformation As Any) _
      As Long
    #End If

 
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Public WINDOWS_VER


Sub ListSheetsInfo()
'contextures.com
Dim ws As Worksheet
Dim lCount As Long
Dim wsTemp As Worksheet
Dim rngF As Range
Dim sh As Shape
Dim lFields As Long
Dim lTab As Long
Dim lCh As Long
Dim lSh As Long
Dim rngDV As Range
Dim rngCF As Range
Dim vFml As Variant
Dim vDV As Variant
Dim vCF As Variant
Dim vTbl As Variant
Dim vSh As Variant
Dim vCh As Variant
Dim vPT As Variant
Dim colLC As Long
Dim strNA As String
Dim strLC As String
Application.EnableEvents = False
Application.ScreenUpdating = False
On Error Resume Next
  
Set wsTemp = Worksheets.Add(Before:=Sheets(1))
lCount = 2
lFields = 14 'don't count tab color
strNA = " --"
colLC = 7 'column with last cell link

With wsTemp
  .Range(.Cells(1, 1), _
    .Cells(1, lFields + 1)).Value _
        = Array( _
            "Order", _
            "Sheet Name", _
            "Code Name", _
            "Protected", _
            "Used Range", _
            "Range Cells", _
            "Last Cell", _
            "DV Cells", _
            "CF Cells", _
            "Tables", _
            "Formulas", _
            "Pivot Tables", _
            "Shapes", _
            "Charts", _
            "Tab Color")
End With

For Each ws In ActiveWorkbook.Worksheets
  'Do not list Very Hidden sheets (
  If ws.Name <> wsTemp.Name _
    And ws.Visible <> 2 Then
    strLC = ws.Cells _
      .SpecialCells(xlCellTypeLastCell) _
        .Address
    If ws.ProtectContents = True Then
      vTbl = strNA
      vSh = strNA
      vCh = strNA
      vPT = strNA
      vFml = strNA
      vDV = strNA
      vCF = strNA

    Else
      vTbl = ws.ListObjects.Count
      vPT = ws.PivotTables.Count
      vSh = ws.Shapes.Count
      vCh = ws.ChartObjects.Count
      vSh = vSh - vCh
    
      Set rngF = Nothing
      vFml = 0
      Set rngF = ws.Cells _
        .SpecialCells(xlCellTypeFormulas, 23)
      If Not rngF Is Nothing Then
        vFml = rngF.Cells.Count
      End If
    
      Set rngDV = Nothing
      vDV = 0
      Set rngDV = ws.Cells _
        .SpecialCells(xlCellTypeAllValidation)
      If Not rngDV Is Nothing Then
        vDV = rngDV.Cells.Count
      End If
    
      Set rngCF = Nothing
      vCF = 0
  'xlCellTypeAllFormatConditions
      Set rngCF = ws.Cells.SpecialCells(-4172)
      If Not rngCF Is Nothing Then
        vCF = rngCF.Cells.Count
      End If
    End If
      
    With wsTemp
      .Range(.Cells(lCount, 1), _
        .Cells(lCount, lFields)).Value _
        = Array( _
            ws.Index, _
            ws.Name, _
            ws.CodeName, _
            ws.ProtectContents, _
            ws.UsedRange.Address, _
            ws.UsedRange.Cells.Count, _
            strLC, _
            vDV, _
            vCF, _
            vTbl, _
            vFml, _
            vPT, _
            vSh, _
            vCh, _
            vPT)
      lTab = 0
      lTab = ws.Tab.Color
      If lTab > 0 Then
        On Error Resume Next
        With .Cells(lCount, lFields + 1)
'            .Value = lTab
          .Interior.Color = lTab
        End With
      End If
      'add hyperlink to sheet name in col B
      .Hyperlinks.Add _
          Anchor:=.Cells(lCount, 2), _
          Address:="", _
          subAddress:="'" & ws.Name & "'!A1", _
          ScreenTip:=ws.Name, _
          TextToDisplay:=ws.Name
      'add hyperlink to last cell
      .Hyperlinks.Add _
          Anchor:=.Cells(lCount, colLC), _
          Address:="", _
          subAddress:="'" & ws.Name _
              & "'!" & strLC, _
          ScreenTip:=strLC, _
          TextToDisplay:=strLC
      lCount = lCount + 1
    End With
  End If
Next ws
 
With wsTemp
  .ListObjects.Add(xlSrcRange, _
    .Range("A1").CurrentRegion, , xlYes) _
      .Name = ""
  .ListObjects(1) _
    .ShowTableStyleRowStripes = False
End With

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub


Sub WinVer()
Dim osvi As OSVERSIONINFO
Dim strOut As String

    osvi.dwOSVersionInfoSize = Len(osvi)
    If CBool(apiGetVersionEx(osvi)) Then
        With osvi
            ' Win 2000
            
           If .dwMajorVersion > 5 Then
           WINDOWS_VER = 6
           End If
           End With
     End If
 
End Sub


Sub Sheet_Index()
'I have the below macro assigned to a shortcut key - Ctrl+q
'It shows a very interesting thing….
'The behaviour of SendKeys is different not only
'between 2003 and 2007 but also between XP and VISTA….

If Application.CommandBars("workbook tabs").Controls(16).Caption Like "More Sheets*" Then
  Application.ScreenUpdating = False
  If WINDOWS_VER > 5 Then
    If Application.Version = "12.0" Then
      Application.SendKeys "{end}~"""
      Application.CommandBars("workbook tabs").ShowPopup
    Else
      Application.SendKeys "{end}~"
      Application.CommandBars("workbook tabs").Controls(16).Execute
    End If
  Else
    Application.SendKeys "{end}~"
    Application.CommandBars("workbook tabs").ShowPopup
  End If
  Application.ScreenUpdating = True
Else
  Application.CommandBars("workbook tabs").ShowPopup
End If

Application.ScreenUpdating = True

End Sub

Sub CreateTableOfContents()
'Below is a macro that creates a Table of Contents sheet and
'puts a hyperlink to every sheet that isn’t hidden and is not the current sheet.
'The links do not work for sheets that are graphs and I do not know
'how to either make them work or test that they are graphs and not include them.

Dim shtName As String
Dim shtLink As String
Dim rowNum As Integer
Dim Sht As Worksheet
Dim i As Long

Set Sht = ThisWorkbook.Sheets("Contents")
Sht.Select

Sht.Range("A1").Value = "Table of Contents"
rowNum = 6

For i = 1 To Sheets.Count
  'Does not create a link if the Sheet isn’t visible or the sheet is the current sheet
  If Sheets(i).Visible = True _
        And Sheets(i).Name <> ActiveSheet.Name _
        And IsSheet(Sheets(i).Name) Then
    shtName = Sheets(i).Name
    If shtName = "-->" Then GoTo NextSheet
    shtLink = "'" & shtName & "'!A1"
    Sht.Cells(rowNum, 1).Select
    'inserts the hyperlink to the sheet and cell A1
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", _
      subAddress:=shtLink, TextToDisplay:=shtName
    rowNum = rowNum + 1
  End If
NextSheet: Next i

End Sub

Public Function IsSheet(cName As String) As Boolean

Dim tmpChart As Chart

On Error Resume Next
Set tmpChart = Charts(cName)
On Error GoTo 0
IsSheet = IIf(tmpChart Is Nothing, True, False)

End Function


Sub GenerateSheets()

  Dim SheetRng As Range
  Dim Sht As Worksheet
  Dim i As Long
  Dim rowNum As Long
  
  rowNum = 7
  
  Set SheetRng = Sheets("Contents").Range("L7:L38")
  
  For i = 1 To SheetRng.Cells.Count
  
    Set Sht = Sheets.Add
    Sht.Name = Sheets("Contents").Cells(rowNum, 12).Value
    rowNum = rowNum + 1
    
  Next
  

End Sub




