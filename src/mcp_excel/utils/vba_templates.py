"""VBA code templates for common Excel automation tasks."""

VBA_TEMPLATES = {
    "hello_world": {
        "description": "Simple Hello World macro",
        'code': 'Sub HelloWorld()\n    MsgBox "Hello, World!"\nEnd Sub',
    },
    "loop_range": {
        "description": "Loop through a range and process cells",
        "code": """Sub LoopThroughRange()
    Dim cell As Range
    For Each cell In Range("A1:A10")
        If cell.Value <> "" Then
            cell.Interior.Color = vbYellow
        End If
    Next cell
End Sub""",
    },
    "filter_data": {
        "description": "AutoFilter example",
        "code": """Sub FilterData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1").CurrentRegion.AutoFilter Field:=1, Criteria1:=">100"
End Sub""",
    },
    "create_chart": {
        "description": "Create a chart from data",
        "code": """Sub CreateChart()
    Dim chart As Chart
    Set chart = Charts.Add
    chart.SetSourceData Source:=ActiveSheet.UsedRange
    chart.ChartType = xlColumnClustered
End Sub""",
    },
    "copy_to_sheet": {
        "description": "Copy data to another sheet",
        "code": """Sub CopyToSheet()
    Dim src As Worksheet, dst As Worksheet
    Set src = Sheets("Source")
    Set dst = Sheets("Destination")
    src.UsedRange.Copy dst.Range("A1")
End Sub""",
    },
    "send_email": {
        "description": "Send email via Outlook",
        "code": """Sub SendEmail()
    Dim olApp As Object
    Dim olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    With olMail
        .To = "recipient@example.com"
        .Subject = "Report"
        .Body = "Please see attached report."
        .Send
    End With
End Sub""",
    },
    "format_cells": {
        "description": "Format cells with colors and styles",
        "code": """Sub FormatCells()
    Dim rng As Range
    Set rng = Range("A1:D10")
    ' Header formatting
    With rng.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = vbWhite
    End With
    ' Data formatting
    With rng.Rows("2:" & rng.Rows.Count)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
End Sub""",
    },
    "pivot_table": {
        "description": "Create a PivotTable",
        "code": """Sub CreatePivotTable()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Set wsData = Sheets("Data")
    Set wsPivot = Sheets.Add
    wsPivot.Name = "PivotTable"
    Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.UsedRange)
    Set pvt = pvtCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="PivotTable1")
End Sub""",
    },
    "validate_input": {
        "description": "Data validation with dropdown list",
        "code": """Sub AddDataValidation()
    Dim rng As Range
    Set rng = Range("A1:A100")
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="Option1,Option2,Option3"
        .ErrorTitle = "Invalid Input"
        .ErrorMessage = "Please select from the list."
    End With
End Sub""",
    },
    "workbook_events": {
        "description": "ThisWorkbook event handlers",
        "code": """' Place in ThisWorkbook module
Private Sub Workbook_Open()
    MsgBox "Welcome to " & ThisWorkbook.Name
End Sub

Private Sub Workbook_BeforeSave(Cancel As Boolean)
    If MsgBox("Save changes?", vbYesNo) = vbNo Then
        Cancel = True
    End If
End Sub""",
    },
    "error_handling": {
        "description": "Error handling template",
        "code": """Sub SafeMacro()
    On Error GoTo ErrorHandler
    ' Your code here
    Range("A1").Value = "Success"
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub""",
    },
    "file_operations": {
        "description": "Read/write text files",
        "code": """Sub ExportToText()
    Dim fso As Object
    Dim ts As Object
    Dim rng As Range
    Dim cell As Range
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile("C:\\export.txt", True)
    Set rng = Range("A1:A" & Cells(Rows.Count, 1).End(xlUp).Row)
    For Each cell In rng
        ts.WriteLine cell.Value
    Next cell
    ts.Close
End Sub""",
    },
}
