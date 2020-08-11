VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Consolidation1 
   Caption         =   "Consolidation Macro"
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4860
   OleObjectBlob   =   "Bin Consolidation V2.1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Consolidation1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload Me

Application.ScreenUpdating = False

Worksheets(1).Select
Worksheets.Add(after:=Sheets(1)).Name = "B1"
Worksheets.Add(after:=Sheets(2)).Name = "B2"
Worksheets.Add(after:=Sheets(3)).Name = "B3"
Worksheets.Add(after:=Sheets(4)).Name = "M1"
Worksheets.Add(after:=Sheets(5)).Name = "M2"
Worksheets.Add(after:=Sheets(6)).Name = "M3"

Worksheets(1).Select


Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=5, Criteria1:="BIN"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=8, Criteria1:="0"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=4, Criteria1:="=B1*" _
        , Operator:=xlAnd, Criteria2:="<>*PDI*"
    Range("A1:P1000000").Select
    Selection.Copy
    Worksheets(2).Select
    Range("A1").Select
    ActiveSheet.Paste
    Worksheets(1).Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Range("A1").Select

    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=5, Criteria1:="BIN"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=8, Criteria1:="0"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=4, Criteria1:="=B2*" _
        , Operator:=xlAnd, Criteria2:="<>*PDI*"
    Range("A1:P1000000").Select
    Selection.Copy
    Worksheets(3).Select
    Range("A1").Select
    ActiveSheet.Paste
    Worksheets(1).Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Range("A1").Select


    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=5, Criteria1:="BIN"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=8, Criteria1:="0"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=4, Criteria1:="=B3*" _
        , Operator:=xlAnd, Criteria2:="<>*PDI*"
    Range("A1:P1000000").Select
    Selection.Copy
    Worksheets(4).Select
    Range("A1").Select
    ActiveSheet.Paste
    Worksheets(1).Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Range("A1").Select

    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=5, Criteria1:="BIN"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=8, Criteria1:="0"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=4, Criteria1:="=M1*" _
        , Operator:=xlAnd, Criteria2:="<>*PDI*"
    Range("A1:P1000000").Select
    Selection.Copy
    Worksheets(5).Select
    Range("A1").Select
    ActiveSheet.Paste
    Worksheets(1).Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Range("A1").Select

    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=5, Criteria1:="BIN"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=8, Criteria1:="0"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=4, Criteria1:="=M2*" _
        , Operator:=xlAnd, Criteria2:="<>*PDI*"
    Range("A1:P1000000").Select
    Selection.Copy
    Worksheets(6).Select
    Range("A1").Select
    ActiveSheet.Paste
    Worksheets(1).Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Range("A1").Select


    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=5, Criteria1:="BIN"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=8, Criteria1:="0"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=4, Criteria1:="=M3*" _
        , Operator:=xlAnd, Criteria2:="<>*PDI*"
    Range("A1:P1000000").Select
    Selection.Copy
    Worksheets(7).Select
    Range("A1").Select
    ActiveSheet.Paste
    Worksheets(1).Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Range("A1").Select
    


'_______


    Worksheets(2).Select
    Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "B1!R1C1:R500000C16", Version:=6).CreatePivotTable TableDestination:= _
        "B1!R1C29", TableName:="PivotTable1", DefaultVersion:=6
    Worksheets(2).Select
    Cells(1, 29).Select
    With ActiveSheet.PivotTables("PivotTable1")
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
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Location"), "Count of Location", xlCount
    Columns("AC:AD").Select
    Selection.Copy
    Range("AA1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AC:AD").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Count of Locations"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(R2C27:R500000C28,MATCH(RC[-16],R2C27:R500000C27,0),2)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q" & Range("A" & Rows.Count).End(xlUp).Row)
    Range("Q2:Q500000").Select
    Columns("Q:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("S1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "B1!R1C1:R1048576C17", Version:=6).CreatePivotTable TableDestination:= _
        "B1!R1C19", TableName:="PivotTable1", DefaultVersion:=6
    Worksheets(2).Select
    Cells(1, 19).Select
    With ActiveSheet.PivotTables("PivotTable1")
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
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of Locations")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Qty"), "Sum of Qty", xlSum
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of Locations"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of Locations")
        .PivotItems("1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of Locations"). _
        EnableMultiplePageItems = True
        
Range("W1").Select
    ActiveCell.FormulaR1C1 = "Skus more than 1 location"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[5],"">1"")"
    Range("W2").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Range("W4").Select
    ActiveCell.FormulaR1C1 = "Pages"
    Range("W5").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-4])/47"
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
Columns("S:U").Select
    Selection.ColumnWidth = 20
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
        ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "Please turn in when Completed"
        .CenterHeader = "Printed &D"
        .RightHeader = "Page &P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Completed By ______________________"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
        ActiveWindow.View = xlPageBreakPreview


'_______

    Worksheets(3).Select
    Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "B2!R1C1:R500000C16", Version:=6).CreatePivotTable TableDestination:= _
        "B2!R1C29", TableName:="PivotTable2", DefaultVersion:=6
    Worksheets(3).Select
    Cells(1, 29).Select
    With ActiveSheet.PivotTables("PivotTable2")
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
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Location"), "Count of Location", xlCount
    Columns("AC:AD").Select
    Selection.Copy
    Range("AA1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AC:AD").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Count of Locations"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(R2C27:R500000C28,MATCH(RC[-16],R2C27:R500000C27,0),2)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q" & Range("A" & Rows.Count).End(xlUp).Row)
    Range("Q2:Q500000").Select
    Columns("Q:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("S1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "B2!R1C1:R1048576C17", Version:=6).CreatePivotTable TableDestination:= _
        "B2!R1C19", TableName:="PivotTable2", DefaultVersion:=6
    Worksheets(3).Select
    Cells(1, 19).Select
    With ActiveSheet.PivotTables("PivotTable2")
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
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Count of Locations")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Qty"), "Sum of Qty", xlSum
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Count of Locations"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Count of Locations")
        .PivotItems("1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Count of Locations"). _
        EnableMultiplePageItems = True

Range("W1").Select
    ActiveCell.FormulaR1C1 = "Skus more than 1 location"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[5],"">1"")"
    Range("W2").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Range("W4").Select
    ActiveCell.FormulaR1C1 = "Pages"
    Range("W5").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-4])/47"
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
       End With
    
    Columns("S:U").Select
    Selection.ColumnWidth = 20
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
        ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "Please turn in when Completed"
        .CenterHeader = "Printed &D"
        .RightHeader = "Page &P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Completed By ______________________"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
        ActiveWindow.View = xlPageBreakPreview
    
'_______

    Worksheets(4).Select
    Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "B3!R1C1:R500000C16", Version:=6).CreatePivotTable TableDestination:= _
        "B3!R1C29", TableName:="PivotTable3", DefaultVersion:=6
    Worksheets(4).Select
    Cells(1, 29).Select
    With ActiveSheet.PivotTables("PivotTable3")
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
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable3").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Location"), "Count of Location", xlCount
    Columns("AC:AD").Select
    Selection.Copy
    Range("AA1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AC:AD").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Count of Locations"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(R2C27:R500000C28,MATCH(RC[-16],R2C27:R500000C27,0),2)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q" & Range("A" & Rows.Count).End(xlUp).Row)
    Range("Q2:Q500000").Select
    Columns("Q:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("S1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "B3!R1C1:R1048576C17", Version:=6).CreatePivotTable TableDestination:= _
        "B3!R1C19", TableName:="PivotTable3", DefaultVersion:=6
    Worksheets(4).Select
    Cells(1, 19).Select
    With ActiveSheet.PivotTables("PivotTable3")
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
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable3").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Count of Locations")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Qty"), "Sum of Qty", xlSum
    ActiveSheet.PivotTables("PivotTable3").PivotFields("Count of Locations"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Count of Locations")
        .PivotItems("1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable3").PivotFields("Count of Locations"). _
        EnableMultiplePageItems = True

Range("W1").Select
    ActiveCell.FormulaR1C1 = "Skus more than 1 location"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[5],"">1"")"
    Range("W2").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Range("W4").Select
    ActiveCell.FormulaR1C1 = "Pages"
    Range("W5").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-4])/47"
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
   End With

    Columns("S:U").Select
    Selection.ColumnWidth = 20
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "Please turn in when Completed"
        .CenterHeader = "Printed &D"
        .RightHeader = "Page &P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Completed By ______________________"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
        ActiveWindow.View = xlPageBreakPreview


'_______


    Worksheets(5).Select
    Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "M1!R1C1:R500000C16", Version:=6).CreatePivotTable TableDestination:= _
        "M1!R1C29", TableName:="PivotTable4", DefaultVersion:=6
    Worksheets(5).Select
    Cells(1, 29).Select
    With ActiveSheet.PivotTables("PivotTable4")
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
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable4").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("Location"), "Count of Location", xlCount
    Columns("AC:AD").Select
    Selection.Copy
    Range("AA1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AC:AD").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Count of Locations"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(R2C27:R500000C28,MATCH(RC[-16],R2C27:R500000C27,0),2)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q" & Range("A" & Rows.Count).End(xlUp).Row)
    Range("Q2:Q500000").Select
    Columns("Q:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("S1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "M1!R1C1:R1048576C17", Version:=6).CreatePivotTable TableDestination:= _
        "M1!R1C19", TableName:="PivotTable4", DefaultVersion:=6
    Worksheets(5).Select
    Cells(1, 19).Select
    With ActiveSheet.PivotTables("PivotTable4")
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
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable4").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Count of Locations")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("Qty"), "Sum of Qty", xlSum
    ActiveSheet.PivotTables("PivotTable4").PivotFields("Count of Locations"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Count of Locations")
        .PivotItems("1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable4").PivotFields("Count of Locations"). _
        EnableMultiplePageItems = True
        
Range("W1").Select
    ActiveCell.FormulaR1C1 = "Skus more than 1 location"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[5],"">1"")"
    Range("W2").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Range("W4").Select
    ActiveCell.FormulaR1C1 = "Pages"
    Range("W5").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-4])/47"
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
       End With
    
    Columns("S:U").Select
    Selection.ColumnWidth = 20
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "Please turn in when Completed"
        .CenterHeader = "Printed &D"
        .RightHeader = "Page &P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Completed By ______________________"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
        ActiveWindow.View = xlPageBreakPreview
    
'_______
        
    Worksheets(6).Select
    Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "M2!R1C1:R500000C16", Version:=6).CreatePivotTable TableDestination:= _
        "M2!R1C29", TableName:="PivotTable5", DefaultVersion:=6
    Worksheets(6).Select
    Cells(1, 29).Select
    With ActiveSheet.PivotTables("PivotTable5")
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
    End With
    With ActiveSheet.PivotTables("PivotTable5").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable5").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Location"), "Count of Location", xlCount
    Columns("AC:AD").Select
    Selection.Copy
    Range("AA1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AC:AD").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Count of Locations"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(R2C27:R500000C28,MATCH(RC[-16],R2C27:R500000C27,0),2)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q" & Range("A" & Rows.Count).End(xlUp).Row)
    Range("Q2:Q500000").Select
    Columns("Q:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("S1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "M2!R1C1:R1048576C17", Version:=6).CreatePivotTable TableDestination:= _
        "M2!R1C19", TableName:="PivotTable5", DefaultVersion:=6
    Worksheets(6).Select
    Cells(1, 19).Select
    With ActiveSheet.PivotTables("PivotTable5")
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
    End With
    With ActiveSheet.PivotTables("PivotTable5").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable5").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Count of Locations")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Qty"), "Sum of Qty", xlSum
    ActiveSheet.PivotTables("PivotTable5").PivotFields("Count of Locations"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Count of Locations")
        .PivotItems("1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable5").PivotFields("Count of Locations"). _
        EnableMultiplePageItems = True
        
Range("W1").Select
    ActiveCell.FormulaR1C1 = "Skus more than 1 location"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[5],"">1"")"
    Range("W2").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Range("W4").Select
    ActiveCell.FormulaR1C1 = "Pages"
    Range("W5").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-4])/47"
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
       End With
       
    Columns("S:U").Select
    Selection.ColumnWidth = 20
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "Please turn in when Completed"
        .CenterHeader = "Printed &D"
        .RightHeader = "Page &P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Completed By ______________________"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
        ActiveWindow.View = xlPageBreakPreview
    
    
'_______
        
    Worksheets(7).Select
    Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "M3!R1C1:R500000C16", Version:=6).CreatePivotTable TableDestination:= _
        "M3!R1C29", TableName:="PivotTable6", DefaultVersion:=6
    Worksheets(7).Select
    Cells(1, 29).Select
    With ActiveSheet.PivotTables("PivotTable6")
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
    End With
    With ActiveSheet.PivotTables("PivotTable6").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable6").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
        "PivotTable6").PivotFields("Location"), "Count of Location", xlCount
    Columns("AC:AD").Select
    Selection.Copy
    Range("AA1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AC:AD").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Count of Locations"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(R2C27:R500000C28,MATCH(RC[-16],R2C27:R500000C27,0),2)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q" & Range("A" & Rows.Count).End(xlUp).Row)
    Range("Q2:Q500000").Select
    Columns("Q:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("S1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "M3!R1C1:R1048576C17", Version:=6).CreatePivotTable TableDestination:= _
        "M3!R1C19", TableName:="PivotTable6", DefaultVersion:=6
    Worksheets(7).Select
    Cells(1, 19).Select
    With ActiveSheet.PivotTables("PivotTable6")
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
    End With
    With ActiveSheet.PivotTables("PivotTable6").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable6").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Count of Locations")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
        "PivotTable6").PivotFields("Qty"), "Sum of Qty", xlSum
    ActiveSheet.PivotTables("PivotTable6").PivotFields("Count of Locations"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Count of Locations")
        .PivotItems("1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable6").PivotFields("Count of Locations"). _
        EnableMultiplePageItems = True
        
Range("W1").Select
    ActiveCell.FormulaR1C1 = "Skus more than 1 location"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[5],"">1"")"
    Range("W2").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Range("W4").Select
    ActiveCell.FormulaR1C1 = "Pages"
    Range("W5").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-4])/47"
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
       End With
       
    Columns("S:U").Select
    Selection.ColumnWidth = 20
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "Please turn in when Completed"
        .CenterHeader = "Printed &D"
        .RightHeader = "Page &P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Completed By ______________________"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
        ActiveWindow.View = xlPageBreakPreview
        
Worksheets(2).Select
Application.ScreenUpdating = True
        
End Sub

Private Sub CommandButton2_Click()
Unload Me
Application.ScreenUpdating = False
Worksheets(1).Select
Worksheets.Add(after:=Sheets(1)).Name = "LNC"

Worksheets(1).Select
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=5, Criteria1:="INV"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=8, Criteria1:="0"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=4, Criteria1:="=L*"
    Range("A1:P1000000").Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Worksheets(2).Select
    Range("A1").Select
    ActiveSheet.Paste
    Worksheets(1).Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Range("A1").Select

Worksheets(2).Select
    Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "LNC!R1C1:R500000C16", Version:=6).CreatePivotTable TableDestination:= _
        "LNC!R1C29", TableName:="PivotTable7", DefaultVersion:=6
    Worksheets(2).Select
    Cells(1, 29).Select
    With ActiveSheet.PivotTables("PivotTable7")
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
    End With
    With ActiveSheet.PivotTables("PivotTable7").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable7").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable7").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable7").AddDataField ActiveSheet.PivotTables( _
        "PivotTable7").PivotFields("Location"), "Count of Location", xlCount
    Columns("AC:AD").Select
    Selection.Copy
    Range("AA1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AC:AD").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("ab1").FormulaR1C1 = "Count of LPNs"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Count of LPNs"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(R2C27:R500000C28,MATCH(RC[-16],R2C27:R500000C27,0),2)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q" & Range("A" & Rows.Count).End(xlUp).Row)
    Range("Q2:Q500000").Select
    Columns("Q:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("S1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "LNC!R1C1:R1048576C17", Version:=6).CreatePivotTable TableDestination:= _
        "LNC!R1C19", TableName:="PivotTable7", DefaultVersion:=6
    Worksheets(2).Select
    Cells(1, 19).Select
    With ActiveSheet.PivotTables("PivotTable7")
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
    End With
    With ActiveSheet.PivotTables("PivotTable7").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable7").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable7").PivotFields("Count of LPNs")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable7").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable7").PivotFields("Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable7").AddDataField ActiveSheet.PivotTables( _
        "PivotTable7").PivotFields("Qty"), "Sum of Qty", xlSum
    ActiveSheet.PivotTables("PivotTable7").PivotFields("Count of LPNs"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable7").PivotFields("Count of LPNs")
        .PivotItems("1").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable7").PivotFields("Count of LPNs"). _
        EnableMultiplePageItems = True
        
Range("W1").Select
    ActiveCell.FormulaR1C1 = "Skus more than 1 LPNs"
    Range("W2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[5],"">1"")"
    Range("W2").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Range("W4").Select
    ActiveCell.FormulaR1C1 = "Pages"
    Range("W5").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-4])/47"
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
Columns("S:U").Select
    Selection.ColumnWidth = 20
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDashDotDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$S:$U"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "Please turn in when Completed"
        .CenterHeader = "Printed &D"
        .RightHeader = "Page &P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Completed By ______________________"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
        ActiveWindow.View = xlPageBreakPreview
Application.ScreenUpdating = True

End Sub

Private Sub CommandButton3_Click()
Unload Me
 
 Application.ScreenUpdating = False

Worksheets(1).Select

Worksheets.Add(after:=Sheets(1)).Name = "Unique"

Worksheets(1).Select

Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=5, Criteria1:="BIN"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=8, Criteria1:="0"
    ActiveSheet.Range("$A$1:$P$1000000").AutoFilter Field:=4, Criteria1:="<>D*" _
        , Operator:=xlAnd, Criteria2:="<>*PDI*"
    Range("$A$1:$P$1000000").Select
    Selection.Copy
    Worksheets(2).Select
    Range("A1").Select
    ActiveSheet.Paste
    Worksheets(1).Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Range("A1").Select

        
    Worksheets(2).Select
    Range("A1").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Unique!R1C1:R500000C16", Version:=6).CreatePivotTable TableDestination:= _
        "Unique!R1C31", TableName:="PivotTable8", DefaultVersion:=6
    Worksheets(2).Select
    Cells(1, 31).Select
    With ActiveSheet.PivotTables("PivotTable8")
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
    End With
    With ActiveSheet.PivotTables("PivotTable8").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable8").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable8").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable8").AddDataField ActiveSheet.PivotTables( _
        "PivotTable8").PivotFields("Location"), "Count of Location", xlCount
    Columns("AE:AF").Select
    Selection.Copy
    Range("AC1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AE:AF").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Count of Locations"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = _
        "=INDEX(R2C29:R500000C30,MATCH(RC[-16],R2C29:R500000C29,0),2)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q" & Range("A" & Rows.Count).End(xlUp).Row)
    Columns("Q:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("F:F").Copy
    Range("R1").Select
    ActiveSheet.Paste
    Range("R1").FormulaR1C1 = "Qtys"
    Range("U1").Select
    
    Range("S1").FormulaR1C1 = "Zone"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-15],2)"
    Range("S2").Select
    Selection.AutoFill Destination:=Range("S2:S" & Range("A" & Rows.Count).End(xlUp).Row)
    Columns("S:S").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Unique!R1C1:R1048576C19", Version:=6).CreatePivotTable TableDestination:= _
        "Unique!R1C21", TableName:="PivotTable8", DefaultVersion:=6
    Sheets("Unique").Select
    Cells(1, 21).Select
    With ActiveSheet.PivotTables("PivotTable8")
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
    End With
    With ActiveSheet.PivotTables("PivotTable8").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable8").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable8").PivotFields("Count of Locations")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable8").PivotFields("Qty")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable8").PivotFields("Zone")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable8").PivotFields("Location")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable8").PivotFields("Sku")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable8").AddDataField ActiveSheet.PivotTables( _
        "PivotTable8").PivotFields("Qtys"), "Sum of Qtys", xlSum
    ActiveSheet.PivotTables("PivotTable8").PivotFields("Count of Locations"). _
        ClearAllFilters
    ActiveSheet.PivotTables("PivotTable8").PivotFields("Count of Locations"). _
        CurrentPage = "1"
    ActiveSheet.PivotTables("PivotTable8").PivotFields("Qty").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable8").PivotFields("Qty").CurrentPage = "1"
        
Range("AA5").Select
    ActiveCell.FormulaR1C1 = "Unique SKUs with 1 Location"
    Range("AA6").Select
    ActiveCell.FormulaR1C1 = "=OFFSET(R[-3]C[-5],COUNTA(C[-5])-2,0)"
    Range("AA6").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Range("AA8").Select
    ActiveCell.FormulaR1C1 = "Pages"
    Range("AA9").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-6])/47"
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
       End With
    Columns("AA:AA").EntireColumn.AutoFit
    
    ActiveSheet.PageSetup.PrintArea = "$U:$W"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$U:$W"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "Please turn in when Completed"
        .CenterHeader = "Printed &D"
        .RightHeader = "Page &P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "Completed By ______________________"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
        ActiveWindow.View = xlPageBreakPreview
 Application.ScreenUpdating = True

End Sub
