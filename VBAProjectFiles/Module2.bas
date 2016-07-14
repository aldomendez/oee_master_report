Attribute VB_Name = "OEE_General_Report"

Dim dataSheet As String
Dim pivoteSheet As String

Sub Generate_pivote_and_graph(depto, process)
'
' Create Macro
'

'
    
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Week"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Month"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=WEEKNUM(RC[-14])"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "=DATE(TEXT(RC[-15],""yyyy""),TEXT(RC[-15],""mm""),1)"

    Range("R2:S2").Select
    Selection.AutoFill Destination:=Range("R2:S" & Range("Q" & Rows.Count).End(xlUp).Row)

    Range("S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "mmm"
    Range("K11").Select
    pvt_range = pivote_range()
    Sheets.Add
    ActiveSheet.Name = pivoteSheet
    Sheets(dataSheet).Select
    pivotedest = pivoteSheet & "!R3C1"
    
    Debug.Print pvt_range
    Debug.Print pivoteSheet
    Debug.Print pivotedest

''PMQPSK-DC_TEST-data'!$A$1:$S$88

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        pvt_range, Version:=xlPivotTableVersion12). _
        CreatePivotTable TableDestination:=pivotedest, TableName:= _
        "PivotTable28", DefaultVersion:=xlPivotTableVersion12
        
    Sheets(pivoteSheet).Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Week")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Month")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields("AVAIL"), "Sum of AVAIL", xlSum
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields("PERF"), "Sum of PERF", xlSum
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields("YIELD"), "Sum of YIELD", xlSum
    ActiveSheet.PivotTables("PivotTable28").AddDataField ActiveSheet.PivotTables( _
        "PivotTable28").PivotFields("OEE"), "Sum of OEE", xlSum
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Sum of AVAIL")
        .Caption = "Average of AVAIL"
        .Function = xlAverage
    End With
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Sum of PERF")
        .Caption = "Average of PERF"
        .Function = xlAverage
    End With
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Sum of YIELD")
        .Caption = "Average of YIELD"
        .Function = xlAverage
    End With
    With ActiveSheet.PivotTables("PivotTable28").PivotFields("Sum of OEE")
        .Caption = "Average of OEE"
        .Function = xlAverage
    End With
    Range("B5:E39").Select
    Selection.Style = "Percent"
    Range("E12").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
    Range("D16").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=Range("'" & pivoteSheet & "'!$A$3:$E$39")
    ActiveWorkbook.ShowPivotChartActiveFields = True
    ActiveChart.ChartType = xlLine
    ActiveWorkbook.ShowPivotChartActiveFields = False
    
    Dim RngToCover As Range
    Dim ChtOb As ChartObject
    Set RngToCover = ActiveSheet.Range("G3:W39")
    ActiveSheet.ChartObjects(1).Select
    Set ChtOb = ActiveChart.Parent
    'Set chtopb = ActiveSheet.ChartObjects(1).Parent
    ChtOb.Height = RngToCover.Height ' resize
    ChtOb.Width = RngToCover.Width   ' resize
    ChtOb.Top = RngToCover.Top       ' reposition
    ChtOb.Left = RngToCover.Left     ' reposition
    
    'ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\almendez\Desktop\OEE\Engines.OSA_ASSY.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
End Sub

Function pivote_range()
'Dynamically Retrieve Range Address of Data
  Set StartPoint = ActiveSheet.Range("A1")
  Set DataRange = ActiveSheet.Range(StartPoint, StartPoint.SpecialCells(xlLastCell))
  
  newRange = "'" & ActiveSheet.Name & "'!" & _
    DataRange.Address(ReferenceStyle:=xlR1C1)
  Debug.Print newRange
  pivote_range = newRange
End Function
Sub create_link_to_menu()
'
' create_link_to_menu Macro
'

'
    Sheets(dataSheet).Select
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "Menu!A1", TextToDisplay:="Menu!A1"
    ActiveCell.FormulaR1C1 = "Menu"
    Range("A2").Select
End Sub

Sub create_hiperlink()
    Sheets(pivoteSheet).Select
    Range("A1").Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "'Menu'!A1", TextToDisplay:="Menu"
End Sub

Sub select_range()
    Dim lastRow As Long
    lastRow = Range("Q" & Rows.Count).End(xlUp).Row
    Debug.Print lastRow
    Range("S2").AutoFill Destination:=Range("S2:S" & Range("Q" & Rows.Count).End(xlUp).Row)
End Sub

Sub GenerateReport()
    Call clean_workbooks
    report_index = 0
    Do While Sheets("Menu").Range("A2").Offset(report_index, 0).Value <> ""
        depto = Sheets("Menu").Range("A2").Offset(report_index, 0).Value
        process = Sheets("Menu").Range("A2").Offset(report_index, 1).Value
        
        Call setPivoteSheetName(depto, process)
        Call setDataSheetName(depto, process)
        
        ActiveSheet.Hyperlinks.Add Anchor:=Sheets("Menu").Range("A2").Offset(report_index, 4), Address:="", SubAddress:= _
            "'" & dataSheet & "'!A1", TextToDisplay:="Data"
        ActiveSheet.Hyperlinks.Add Anchor:=Sheets("Menu").Range("A2").Offset(report_index, 5), Address:="", SubAddress:= _
            "'" & pivoteSheet & "'!A1", TextToDisplay:="Pivote"
        
        Call Import_CSV_File_From_URL(depto, process)
        Call Generate_pivote_and_graph(depto, process)
        Call create_link_to_menu
        Call create_hiperlink
        report_index = report_index + 1
        Sheets("Menu").Select
    Loop
End Sub

Function setPivoteSheetName(depto, process)
    pivoteSheet = depto & "_" & process & ""
End Function

Function setDataSheetName(depto, process)
    dataSheet = depto & "_" & process & "-data"
End Function

Sub Import_CSV_File_From_URL(depto, process)
    
    Dim URL As String
    Dim destCell As Range
    
    URL = "http://wmatvmlr401/lr4/dataFeeder/index.php/oee/:depto/:process"
    
    URL = Replace(URL, ":depto", depto)
    URL = Replace(URL, ":process", process)
    
    Debug.Print URL
    Debug.Print dataSheet
    
    Sheets.Add
    
    ActiveSheet.Name = dataSheet
    
    Set destCell = Worksheets(dataSheet).Range("A1")
    
    With destCell.Parent.QueryTables.Add(Connection:="TEXT;" & URL, Destination:=destCell)
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
    
    destCell.Parent.QueryTables(1).Delete

End Sub

Sub clean_workbooks()
Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Debug.Print ws.Name
        If ws.Name = "Menu" Then
        Else
            ws.Delete
        End If
        Count = Count + 1
    Next ws

Application.DisplayAlerts = True
End Sub

Public Sub ExportAll(targetPath As String)
'Two possible issues must be solved to correctly run this code. The first is
'to allow the code to access the VBA Object model programmatically. This
'can be done by correctly setting this from the Trust Center (from Excel Options).

'The second is to reference the Microsoft Visual Basic for Applications
'Extensibility 5.3 object library (from Tools?References) to be able to
'use the VBIDE object model.


 Dim xlApp As Excel.Application
 Dim xlWb As Excel.Workbook
 Dim VBComp As VBIDE.VBComponent
 
 ' Load workbook
 Set xlApp = Application
 'xlApp.Visible = False
 Set xlWb = ActiveWorkbook 'xlApp.Workbooks.Open(sWorkbook)
 
 ' Loop through all files (components) in the workbook
 For Each VBComp In xlWb.VBProject.VBComponents
     ' Export the file
     If VBComp.Type = vbext_ct_StdModule Then _
        VBComp.export targetPath & VBComp.Name & ".bas"
 Next VBComp
End Sub

Sub export()
    Call ExportAll("C:\apps\oee_master_report")
End Sub































