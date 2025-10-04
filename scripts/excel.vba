Sub Init()
'
' Init Macro
'
Range("A1:G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut Destination:=Range("T2")
    
    Range("B2:F5").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Expense Tracker Month Year"
    
    Range("B7:F7").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Transaction Date"
    
    Range("B8:C8").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Start Date"
    
    Range("B9:C9").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "End Date"
       
    Range("D8:F8").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "DD/MM/YYYY"
    
    Range("D9:F9").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "DD/MM/YYYY"

    Range("B11:C12").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Income"
    
    Range("B13:C14").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Expense"
    
    Range("B15:C16").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Investment"
    
    Range("D11:E12").Select
    Selection.Merge
    
    Range("D13:E14").Select
    Selection.Merge
    
    Range("D15:E16").Select
    Selection.Merge
    
    Range("B18:C19").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Income"
    
    Range("E18:F19").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Investment"
    
    Range("H18:I19").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Total Expense"
    
    Range("K18:L19").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Need"
    
    Range("N18:O19").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Want"
    

End Sub

Sub Pivote()
'
' Pivote Macro
'
'
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sourceRange As Range
    Dim i As Integer
    Dim destinations As Variant
    Dim pivotNames As Variant
    Dim pivotTableName As String
    Dim pivotCache As pivotCache
    Dim pt As PivotTable
    
    Set ws = ActiveSheet
    
    ' Define pivot table destinations and names
    destinations = Array("B20", "E20", "H20", "K20", "N20")
    pivotNames = Array(ws.Name & "_pivote1", ws.Name & "_pivote2", ws.Name & "_pivote3", ws.Name & "_pivote4", ws.Name & "_pivote5")
        
    ' Validate and clean pivot table names (remove invalid characters)
    For i = 0 To 4
        pivotNames(i) = Replace(Replace(Replace(pivotNames(i), " ", "_"), ":", ""), "/", "")
        If Len(pivotNames(i)) > 255 Then pivotNames(i) = Left(pivotNames(i), 255)
    Next i
    
    ' Clear existing pivot tables to avoid conflicts
    For Each pt In ws.PivotTables
        pt.TableRange2.Clear
    Next pt
    
    ' Find last row in column G with valid data
    lastRow = ws.Cells(ws.Rows.Count, "Z").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No valid data found in column G starting from T2.", vbCritical
        Exit Sub
    End If
    
    ' Set source range from A31 to G:lastRow
    Set sourceRange = ws.Range("T2:Z" & lastRow)
    
    ' Create a single PivotCache for all pivot tables
    Set pivotCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceRange, Version:=8)
    
    ' Loop to create pivot tables at each destination
    For i = 0 To 4
        pivotTableName = pivotNames(i)
        
        ' Create pivot table
        pivotCache.CreatePivotTable TableDestination:=ws.Range(destinations(i)), TableName:=pivotTableName, DefaultVersion:=8
        
        ' Configure pivot table
        With ws.PivotTables(pivotTableName)
            .ColumnGrand = True
            .HasAutoFormat = True
            .DisplayErrorString = False
            .DisplayNullString = True
            .EnableDrilldown = True
            .ErrorString = ""
            .MergeLabels = False
            .NullString = ""
            .PageFieldOrder = 2
            .PageFieldWrapCount = 0
            .PreserveFormatting = True
            .RowGrand = True
            .SaveData = True
            .PrintTitles = False
            .RepeatItemsOnEachPrintedPage = True
            .TotalsAnnotation = False
            .CompactRowIndent = 1
            .InGridDropZones = False
            .DisplayFieldCaptions = True
            .DisplayMemberPropertyTooltips = False
            .DisplayContextTooltips = True
            .ShowDrillIndicators = True
            .PrintDrillIndicators = False
            .AllowMultipleFilters = False
            .SortUsingCustomLists = True
            .FieldListSortAscending = False
            .ShowValuesRow = False
            .CalculatedMembersInFilters = False
            .RowAxisLayout xlCompactRow
            .RepeatAllLabels xlRepeatLabels
        End With
        
        With ws.PivotTables(pivotTableName).pivotCache
            .RefreshOnFileOpen = False
            .MissingItemsLimit = xlMissingItemsDefault
        End With
        
        ' Show field list only for the first pivot table
        ActiveWorkbook.ShowPivotTableFieldList = (i = 0)
        
        ' Add fields
        With ws.PivotTables(pivotTableName).PivotFields("category")
            .Orientation = xlRowField
            .Position = 1
        End With
        ws.PivotTables(pivotTableName).AddDataField ws.PivotTables(pivotTableName).PivotFields("amount"), "Sum of amount", xlSum
    Next i
    
    ' Add filters
    '
    '
    On Error GoTo PivoteErrorHandler
    With ActiveSheet.PivotTables(ws.Name & "_pivote1").PivotFields("category")
    Dim item As PivotItem
    For Each item In .PivotItems
        Select Case item.Name
            Case "gift-credit", "interest", "dividend", "salary"
                item.Visible = True
            Case Else
                item.Visible = False
        End Select
    Next item
    End With
    With ActiveSheet.PivotTables(ws.Name & "_pivote2").PivotFields("category")
    On Error GoTo PivoteErrorHandler
        For Each item In .PivotItems
            Select Case item.Name
                Case "investment", "investment-fee", "investment-mf", "investment-gold", "investment-stock", _
                "investment-fd", "investment-redeem"
                    item.Visible = True
                Case Else
                    item.Visible = False
            End Select
        Next item
    End With
    With ActiveSheet.PivotTables(ws.Name & "_pivote3").PivotFields("category")
    On Error GoTo PivoteErrorHandler
        For Each item In .PivotItems
            Select Case item.Name
                Case "creditcard", "food", "grocery", "insurance", "others - Account Payment", "others - Merchant Payment", _
                "refund", "rent", "maintanance", "shopping", "travel", "recharge", "insurance-fee", "insurance-redeem", _
                "medical", "entertainment", "cash", "petrol", "grocery_meat", "electricity", "water", "gas", "trip", _
                "gift-debit", "maid"
                    item.Visible = True
                Case Else
                    item.Visible = False
            End Select
        Next item
    End With
    With ActiveSheet.PivotTables(ws.Name & "_pivote4").PivotFields("category")
    On Error GoTo PivoteErrorHandler
        For Each item In .PivotItems
            Select Case item.Name
                Case "creditcard", "food", "grocery", "insurance", "others - Account Payment", "others - Merchant Payment", _
                "refund", "rent", "maintanance", "travel", "recharge", _
                "medical", "cash", "petrol", "grocery_meat", "electricity", "water", "gas", "maid"
                    item.Visible = True
                Case Else
                    item.Visible = False
            End Select
        Next item
    End With
    With ActiveSheet.PivotTables(ws.Name & "_pivote5").PivotFields("category")
        For Each item In .PivotItems
            Select Case item.Name
                Case "shopping", "entertainment", "gift-debit", "trip"
                    item.Visible = True
                Case Else
                    item.Visible = False
            End Select
        Next item
PivoteErrorHandler:
    Application.StatusBar = "Error Adding filters to pivote. check once after execution"
  Resume Next
    End With
    ' set grand total
    '
    '


    Dim grandTotal As Double
    Dim found As Boolean
    Dim targetCells As Variant

    
    ' Define pivot table names and target cells
    pivotNames = Array(ws.Name & "_pivote1", ws.Name & "_pivote2", ws.Name & "_pivote3")
    targetCells = Array("D11", "D15", "D13")
    
    ' Loop through the three pivot tables
    For i = 0 To 2
        pivotTableName = pivotNames(i)
        
        ' Find the pivot table
        found = False
        For Each pt In ws.PivotTables
            If pt.Name = pivotTableName Then
                found = True
                Exit For
            End If
        Next pt
        
        If Not found Then
            MsgBox "Pivot table " & pivotTableName & " not found.", vbCritical
        End If
        
        ' Get the grand total (Sum of amount) from the pivot table
        With pt.DataBodyRange
            grandTotal = .Cells(.Rows.Count, 1).Value
        End With
        
        ' Place the grand total in the target cell
        ws.Range(targetCells(i)).Value = grandTotal
        

    Next i
End Sub

Sub Diff()

    Dim sh As Worksheet
    Dim shprv As Worksheet
    Dim prevSheetName As String

    ' Initialize the current sheet
    Set sh = ActiveSheet

    ' Find the previous sheet
    Dim i As Integer
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = sh.Name Then
            If i > 1 Then
                Set shprv = Worksheets(i - 1)
            Else
                MsgBox "No previous sheet"
                Exit Sub
            End If
            Exit For
        End If
    Next i

    ' Validate the previous sheet
    If shprv Is Nothing Then
        MsgBox "No previous sheet"
        Exit Sub
    End If

    ' Get the name of the previous sheet
    prevSheetName = shprv.Name

    ' Apply formulas
    With sh
         Range("F11").FormulaR1C1 = "=ABS(RC[-2])-ABS('" & prevSheetName & "'!RC[-2])"
        .Range("F13").FormulaR1C1 = "=ABS(RC[-2])-ABS('" & prevSheetName & "'!RC[-2])"
        .Range("F15").FormulaR1C1 = "=ABS(RC[-2])-ABS('" & prevSheetName & "'!RC[-2])"
        .Range("F12").FormulaR1C1 = "=ABS(R[-1]C[-2])/ABS('" & prevSheetName & "'!R[-1]C[-2])"
        .Range("F14").FormulaR1C1 = "=ABS(R[-1]C[-2])/ABS('" & prevSheetName & "'!R[-1]C[-2])"
        .Range("F16").FormulaR1C1 = "=ABS(R[-1]C[-2])/ABS('" & prevSheetName & "'!R[-1]C[-2])"
    End With
End Sub

Sub Validation()
'
' validation Macro
'

'
    Range("V3").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=List!$A$2:$A$38"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("Z3").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=List!$C$2:$C$4"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Sub Chart()

'
' Chart Macro
'
    Dim ch As Shape
    Dim chartName As String
    Dim sheetName As String
    Dim ws As Worksheet
    Dim pt As PivotTable

    Set ws = ActiveSheet
    Set pt = ws.PivotTables(ws.Name & "_pivote3")
    sheetName = ws.Name
    
    
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    Set ch = ActiveSheet.Shapes(1)
    chartName = sheetName & "_Chart"
    ch.Name = chartName
    ActiveChart.SetSourceData Source:=pt.TableRange1
    ActiveSheet.Shapes(chartName).Left = Range("K2").Left
    ActiveSheet.Shapes(chartName).Top = Range("K2").Top
    ActiveSheet.Shapes(chartName).Width = 580
    ActiveSheet.Shapes(chartName).Height = 225
    ActiveChart.ChartColor = 12
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 202
    ActiveChart.SetElement (msoElementLegendNone)
    ActiveChart.SetElement (msoElementChartTitleNone)
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)
    ActiveChart.SetElement (msoElementDataLabelNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisShow)
    
    
    Range("B11:C12").Select
    ActiveSheet.Shapes.AddChart2(251, xlPie).Select
    Set ch = ActiveSheet.Shapes(2)
    chartName = sheetName & "_Chart_Pie"
    ch.Name = chartName
    ActiveChart.SetSourceData Source:=Range("B11:E16")
    ActiveSheet.Shapes(chartName).Select
    ActiveSheet.Shapes(chartName).Left = Range("H2").Left
    ActiveSheet.Shapes(chartName).Top = Range("H2").Top
    ActiveSheet.Shapes(chartName).Height = 225
    ActiveSheet.Shapes(chartName).Width = 225
    ActiveChart.FullSeriesCollection(1).Delete
    ActiveChart.FullSeriesCollection(1).Delete
    ActiveChart.FullSeriesCollection(2).Delete
    ActiveChart.ChartGroups(1).FullCategoryCollection(2).IsFiltered = True
    ActiveChart.ChartGroups(1).FullCategoryCollection(4).IsFiltered = True
    ActiveChart.ChartGroups(1).FullCategoryCollection(6).IsFiltered = True
    ActiveChart.FullSeriesCollection(1).XValues = "='" & sheetName & "'!$B$11:$C$16"
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.ChartColor = 13
    ActiveChart.PlotArea.Select
    ActiveChart.SetElement (msoElementChartTitleNone)
    ActiveChart.SetElement (msoElementDataLabelBestFit)
    ActiveChart.ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.ShowPercentage = True
    Selection.NumberFormat = "0.00%"
    Selection.ShowValue = False
    
    ' pie chart colors
    ActiveSheet.ChartObjects(chartName).Activate
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Points(1).Interior.Color = RGB(169, 209, 142)
    ActiveChart.FullSeriesCollection(1).Points(2).Interior.Color = RGB(255, 198, 117)
    ActiveChart.FullSeriesCollection(1).Points(3).Interior.Color = RGB(148, 180, 203)
End Sub

Sub Styling()

' income, expenses, investments colors

    Dim ws As Worksheet
    Set ws = ActiveSheet
      
    Range("B11:E12").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B13:E14").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("B15:E16").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With

' Pivote table total

    Set pt = ActiveSheet.PivotTables(ws.Name & "_pivote1")
    With pt.TableRange1
        Set lastCell = .Cells(.Rows.Count, .Columns.Count)
    End With
    lastCell.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Set pt = ActiveSheet.PivotTables(ws.Name & "_pivote2")
    With pt.TableRange1
        Set lastCell = .Cells(.Rows.Count, .Columns.Count)
    End With
    lastCell.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Set pt = ActiveSheet.PivotTables(ws.Name & "_pivote3")
    With pt.TableRange1
        Set lastCell = .Cells(.Rows.Count, .Columns.Count)
    End With
    lastCell.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Set pt = ActiveSheet.PivotTables(ws.Name & "_pivote4")
    With pt.TableRange1
        Set lastCell = .Cells(.Rows.Count, .Columns.Count)
    End With
    lastCell.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Set pt = ActiveSheet.PivotTables(ws.Name & "_pivote5")
    With pt.TableRange1
        Set lastCell = .Cells(.Rows.Count, .Columns.Count)
    End With
    lastCell.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    '
    ' diff
    '
    '
    Range("F11:F16").Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("F11:F16").Select
    Selection.Font.Size = 16
    
    '
    ' coma
    '

    '
    Range("A1:Z1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Comma"
     Range("T3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("F12").Select
    Selection.Style = "Percent"
    Range("F14").Select
    Selection.Style = "Percent"
    Range("F16").Select
    Selection.Style = "Percent"
        Range("T3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "m/d/yyyy"
    
    '
    ' conditional formating for categories
    '

    '
    Range("U3:V3").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=OR($V3=""gift"", $V3=""interest"", $V3=""salary"", $V3=""refund"", $V3=""dividend"", $V3=""insurance-redeem"")"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=OR($V3=""investment"", $V3=""investment-mf"",  $V3=""investment-gold"", $V3=""investment-stock"", $V3=""investment-fd"",  $V3=""investment-redeem"",)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=OR($V3=""transfer"")"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399975585192419
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    '
    ' border
    '
    '
    '
    ' title border
    Range("B2:F5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    
    ' title text
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 22
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
     ' transaction date table borders
    
    Range("B7:F9").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
     End With
    
    ' transaction date table text alignment
      Range("B7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
    
    ' main table text alignment
    
    Range("B11:F16").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
    ' main table border
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
   With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
     End With
    
    ' pivote table borders

'

'
    Range("B20:C20").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
     End With
    
    Range("E20:F21").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
     End With
    
    Range("H20:I21").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
     End With
     
    Range("K20:L21").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
     End With
     

    Range("N20:O21").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
     End With
    
    ' pivote table text alignment

    Range("B18:C19,E18:F19,H18:I19,K18:L19,N18:O19").Select
    Range("N18").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' main transactin table border
    
    Range("T2:Z2").Select
    
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
     End With
    
    ' main transaction table header color
    Range("T2:Z2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    ' colors for pivote header
    
    Range("B18:C19").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("E18:F19").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("H18:I19").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("K18:L19").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("N18:O19").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    ' font size
    '
    Range("B11:E16").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("B18:C19,E18:F19,H18:I19,K18:L19,N18:O19").Select
    Range("N18").Activate
    With Selection.Font
        .Name = "Calibri"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Bold = True
    Range("B11:C16").Select
    Selection.Font.Bold = True
    
    
    Range("T2:Z2").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Calibri"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    


End Sub
Sub Main()
'
' main Macro
'
        Call Init
        Call Pivote
        Call Diff
        Call Styling
        Call Validation
        Call Chart

End Sub


