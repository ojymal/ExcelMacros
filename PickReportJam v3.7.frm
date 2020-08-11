VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Jam PPR Macro Version 3.7"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

    Unload UserForm1

Application.ScreenUpdating = False


' Correcting Units in PPR For A-N
 Range("A1").Select
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Units"
    Range("G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(RC[2]>RC[1],RC[2],RC[1])"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & Range("A" & Rows.Count).End(xlUp).Row)
    Columns("G:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("J:J").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("J1").FormulaR1C1 = "Move"
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft

'   Initial sort: Everything Flow Data for "Not Bin" related and "Bin Related"
Range("A1").Select
    Columns("A:L").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$L$200000").AutoFilter Field:=1, Criteria1:="<>B**", _
        Operator:=xlAnd, Criteria2:="<>M**"
    Columns("A:L").Select
    Selection.Copy
    Range("N1").Select
    ActiveSheet.Paste
    Range("AA1").Select
    ActiveSheet.Paste
    Range("BO1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    ActiveSheet.ShowAllData
    ActiveSheet.Range("$A$1:$L$200000").AutoFilter Field:=1, Criteria1:="=B**", _
        Operator:=xlOr, Criteria2:="=M**"
    Columns("A:L").Select
    Selection.Copy
    Range("AN1").Select
    ActiveSheet.Paste
    Range("BB1").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    
    
    '   Data for inventory in PDO destined for Bin B and M
        
    ActiveSheet.Range("$A$1:$L$200000").AutoFilter Field:=1, Criteria1:="=*PDO*" _
        , Operator:=xlAnd
    ActiveSheet.Range("$A$1:$L$200000").AutoFilter Field:=6, Criteria1:="=B*", _
        Operator:=xlOr, Criteria2:="=M*"
    Columns("A:L").Select
    Selection.Copy
    Range("CB1").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    
    
    ' All reserve data no bin
    
    Selection.AutoFilter
    Columns("N:Y").Select
    Selection.AutoFilter
    ActiveSheet.Range("$N$1:$Y$200000").AutoFilter Field:=4, Criteria1:= _
        "=bin*", Operator:=xlOr, Criteria2:="=*replenishment*"
    Range("N1:Y1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("N1").Select
    ActiveSheet.ShowAllData
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "N1:N200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("page").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    
    
    
    'Data for replenishment going to bin
    
    Columns("AA:AL").Select
    Selection.AutoFilter
    ActiveSheet.Range("$AA$1:$AL$200000").AutoFilter Field:=4, Criteria1:= _
        "<>*replenishment*", Operator:=xlAnd
    Range("AA1:AL1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("AA1").Select
    ActiveSheet.ShowAllData
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "AA1:AA200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("page").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    
    
    '   Data for replenishment in Bin
    
    Columns("AN:AY").Select
    Selection.AutoFilter
    ActiveSheet.Range("$AN$1:$AY$200000").AutoFilter Field:=4, Criteria1:= _
        "Bin picking"
    Range("AN1:AY1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    Range("AN1").Select
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "AN1:AN200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("page").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
     Selection.AutoFilter
     
     
     
    '   Bin Picking Data
    
     Columns("BB:BM").Select
    Selection.AutoFilter
    ActiveSheet.Range("$BB$1:$BM$200000").AutoFilter Field:=1, Criteria1:= _
        "=*pdi*", Operator:=xlOr, Criteria2:="=*cart*"
    Range("BB1:BM1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "BB1:BB200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("page").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    

' Everything Filter minus PDO

    Columns("BO:BZ").Select
    Selection.AutoFilter
      ActiveSheet.Range("$BO$1:$BZ$200000").AutoFilter Field:=1, Criteria1:="=*pdo*" _
        , Operator:=xlAnd
    Range("BO1:BZ1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "BO1:BO200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("page").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    
    '   Adding Zone to Replen
    Range("AZ1").FormulaR1C1 = "Zone"
    Range("AZ2").FormulaR1C1 = "=LEFT(RC[-12],5)"
    Range("AZ2").Select
    Selection.AutoFill Destination:=Range("AZ2:AZ" & Range("AN" & Rows.Count).End(xlUp).Row)
    Columns("AZ:AZ").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


    'Pivot Sheet creation and labeling

    Range("A1").Select
    
    Sheets.Add(After:=ActiveSheet).Name = "Pivot"
    ActiveCell.FormulaR1C1 = "Everything / Original"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Picking"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Replenishment"
    Range("A2").Select
    
    
    '   Generating Pivot Tables
    
        Sheets("page").Select
    Columns("BO:BZ").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C67:R200000C78", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R2C1", TableName:="PivotTable6", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(2, 1).Select
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
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
        "PivotTable6").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
        "PivotTable6").PivotFields("Units"), "Sum of Units", xlSum
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable6").CompactLayoutRowHeader = "Reserve Flow"
    
    
    
    Sheets("page").Select
    Columns("A:L").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C1:R200000C12", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R16C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(16, 1).Select
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
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Dest Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Units"), "Sum of Units", xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable1").CompactLayoutRowHeader = "Flow"
    
    
    
        Sheets("page").Select
    Columns("BB:BM").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C54:R200000C65", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R2C5", TableName:="PivotTable5", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(2, 5).Select
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
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Dest Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Units"), "Sum of Units", xlSum
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Free"), "Sum of Free", xlSum
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    On Error Resume Next
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Pick Type")
        .PivotItems("Bin replenishment, partial non-nested pallet (flow 2)"). _
        Visible = False
    End With
    On Error GoTo 0
    ActiveSheet.PivotTables("PivotTable5").CompactLayoutRowHeader = "Bin Pick Flow"
        
    

    Sheets("page").Select
    Columns("N:Y").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C14:R200000C25", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R10C5", TableName:="PivotTable2", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(10, 5).Select
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
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Dest Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Units"), "Sum of Units", xlSum
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable2").CompactLayoutRowHeader = "Reserve Picking Flow"
    
    
        
       Sheets("page").Select
    Columns("AN:AZ").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C40:R200000C52", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R2C10", TableName:="PivotTable4", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(2, 10).Select
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
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("Units"), "Sum of Units", xlSum
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable4").CompactLayoutRowHeader = "Bin Replenishment Flow"
    
        
    Sheets("page").Select
    Columns("CB:CM").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C80:R200000C91", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R8C10", TableName:="PivotTable9", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(8, 10).Select
    With ActiveSheet.PivotTables("PivotTable9")
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
    With ActiveSheet.PivotTables("PivotTable9").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable9").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable9").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    Range("J8").Select
    ActiveSheet.PivotTables("PivotTable9").AddDataField ActiveSheet.PivotTables( _
        "PivotTable9").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable9").AddDataField ActiveSheet.PivotTables( _
        "PivotTable9").PivotFields("Units"), "Sum of Units", xlSum
        
    On Error Resume Next
    With ActiveSheet.PivotTables("PivotTable9").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    On Error GoTo 0
    
    Range("J8").Select
    ActiveSheet.PivotTables("PivotTable9").CompactLayoutRowHeader = _
        "Bin Replenishment in PDO"
        
        
    Sheets("page").Select
    Columns("AA:AL").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C27:R200000C38", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R14C10", TableName:="PivotTable3", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(14, 10).Select
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
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Units"), "Sum of Units", xlSum
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable3").CompactLayoutRowHeader = "Bin Replenishment in Reserve"
    
    
    Range("A1").Select
    Columns("A:A").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Overflow", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("A1").Select
    
Application.ScreenUpdating = True
    
End Sub

Private Sub CommandButton2_Click()

    Unload UserForm1

Application.ScreenUpdating = False


' Correcting Units in PPR For A-N
 Range("A1").Select
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Units"
    Range("G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(RC[2]>RC[1],RC[2],RC[1])"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & Range("A" & Rows.Count).End(xlUp).Row)
    Columns("G:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("J:J").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("J1").FormulaR1C1 = "Move"
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft

'   Initial sort: Everything Flow Data for "Not Bin" related and "Bin Related"
Range("A1").Select
    Columns("A:L").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$L$200000").AutoFilter Field:=1, Criteria1:="<>B**", _
        Operator:=xlAnd, Criteria2:="<>M**"
    Columns("A:L").Select
    Selection.Copy
    Range("N1").Select
    ActiveSheet.Paste
    Range("AA1").Select
    ActiveSheet.Paste
    Range("BO1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    ActiveSheet.ShowAllData
    ActiveSheet.Range("$A$1:$L$200000").AutoFilter Field:=1, Criteria1:="=B**", _
        Operator:=xlOr, Criteria2:="=M**"
    Columns("A:L").Select
    Selection.Copy
    Range("AN1").Select
    ActiveSheet.Paste
    Range("BB1").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    
    
    '   Data for inventory in PDO destined for Bin B and M
        
    ActiveSheet.Range("$A$1:$L$200000").AutoFilter Field:=1, Criteria1:="=*PDO*" _
        , Operator:=xlAnd
    ActiveSheet.Range("$A$1:$L$200000").AutoFilter Field:=6, Criteria1:="=B*", _
        Operator:=xlOr, Criteria2:="=M*"
    Columns("A:L").Select
    Selection.Copy
    Range("CB1").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    
    
    ' All reserve data no bin
    
    Selection.AutoFilter
    Columns("N:Y").Select
    Selection.AutoFilter
    ActiveSheet.Range("$N$1:$Y$200000").AutoFilter Field:=4, Criteria1:= _
        "=bin*", Operator:=xlOr, Criteria2:="=*replenishment*"
    Range("N1:Y1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("N1").Select
    ActiveSheet.ShowAllData
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "N1:N200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("page").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    
    
    
    'Data for replenishment going to bin
    
    Columns("AA:AL").Select
    Selection.AutoFilter
    ActiveSheet.Range("$AA$1:$AL$200000").AutoFilter Field:=4, Criteria1:= _
        "<>*replenishment*", Operator:=xlAnd
    Range("AA1:AL1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("AA1").Select
    ActiveSheet.ShowAllData
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "AA1:AA200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("page").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    
    
    '   Data for replenishment in Bin
    
    Columns("AN:AY").Select
    Selection.AutoFilter
    ActiveSheet.Range("$AN$1:$AY$200000").AutoFilter Field:=4, Criteria1:= _
        "Bin picking"
    Range("AN1:AY1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    Range("AN1").Select
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "AN1:AN200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("page").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
     Selection.AutoFilter
     
     
     
    '   Bin Picking Data
    
     Columns("BB:BM").Select
    Selection.AutoFilter
    ActiveSheet.Range("$BB$1:$BM$200000").AutoFilter Field:=1, Criteria1:= _
        "=*pdi*", Operator:=xlOr, Criteria2:="=*cart*"
    Range("BB1:BM1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "BB1:BB200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("page").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    

' Everything Filter minus PDO

    Columns("BO:BZ").Select
    Selection.AutoFilter
      ActiveSheet.Range("$BO$1:$BZ$200000").AutoFilter Field:=1, Criteria1:="=*pdo*" _
        , Operator:=xlAnd
    Range("BO1:BZ1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("page").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "BO1:BO200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("page").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    
    
    '   Adding Zone to Replen
    Range("AZ1").FormulaR1C1 = "Zone"
    Range("AZ2").FormulaR1C1 = "=LEFT(RC[-12],5)"
    Range("AZ2").Select
    Selection.AutoFill Destination:=Range("AZ2:AZ" & Range("AN" & Rows.Count).End(xlUp).Row)
    Columns("AZ:AZ").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


    'Pivot Sheet creation and labeling

    Range("A1").Select
    
    Sheets.Add(After:=ActiveSheet).Name = "Pivot"
    ActiveCell.FormulaR1C1 = "Everything / Original"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Picking"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Replenishment"
    Range("A2").Select
    
    Sheets.Add(After:=ActiveSheet).Name = "ReplenInfo"
    
    
    
    'Generating Pivot Tables
    
    Sheets("page").Select
    
    Columns("BO:BZ").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C67:R200000C78", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R2C1", TableName:="PivotTable6", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(2, 1).Select
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
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
        "PivotTable6").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
        "PivotTable6").PivotFields("Units"), "Sum of Units", xlSum
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable6").CompactLayoutRowHeader = "Reserve Flow"
    
    
    
    Sheets("page").Select
    Columns("A:L").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C1:R200000C12", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R16C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(16, 1).Select
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
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Dest Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Units"), "Sum of Units", xlSum
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable1").CompactLayoutRowHeader = "Flow"
    
    
    
        Sheets("page").Select
    Columns("BB:BM").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C54:R200000C65", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R2C5", TableName:="PivotTable5", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(2, 5).Select
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
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Dest Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Units"), "Sum of Units", xlSum
    ActiveSheet.PivotTables("PivotTable5").AddDataField ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Free"), "Sum of Free", xlSum
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    On Error Resume Next
    With ActiveSheet.PivotTables("PivotTable5").PivotFields("Pick Type")
        .PivotItems("Bin replenishment, partial non-nested pallet (flow 2)"). _
        Visible = False
    End With
    On Error GoTo 0
    ActiveSheet.PivotTables("PivotTable5").CompactLayoutRowHeader = "Bin Pick Flow"
        
    

    Sheets("page").Select
    Columns("N:Y").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C14:R200000C25", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R10C5", TableName:="PivotTable2", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(10, 5).Select
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
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Dest Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Units"), "Sum of Units", xlSum
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable2").CompactLayoutRowHeader = "Reserve Picking Flow"
    
    
        
       Sheets("page").Select
    Columns("AN:AZ").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C40:R200000C52", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R2C10", TableName:="PivotTable4", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(2, 10).Select
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
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("Units"), "Sum of Units", xlSum
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable4").CompactLayoutRowHeader = "Bin Replenishment Flow"
    
        
    Sheets("page").Select
    Columns("CB:CM").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C80:R200000C91", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R8C10", TableName:="PivotTable9", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(8, 10).Select
    With ActiveSheet.PivotTables("PivotTable9")
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
    With ActiveSheet.PivotTables("PivotTable9").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable9").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable9").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    Range("J8").Select
    ActiveSheet.PivotTables("PivotTable9").AddDataField ActiveSheet.PivotTables( _
        "PivotTable9").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable9").AddDataField ActiveSheet.PivotTables( _
        "PivotTable9").PivotFields("Units"), "Sum of Units", xlSum
        
    On Error Resume Next
    With ActiveSheet.PivotTables("PivotTable9").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    On Error GoTo 0
    
    Range("J8").Select
    ActiveSheet.PivotTables("PivotTable9").CompactLayoutRowHeader = _
        "Bin Replenishment in PDO"
        
        
    Sheets("page").Select
    Columns("AA:AL").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C27:R200000C38", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R14C10", TableName:="PivotTable3", DefaultVersion:=6
    Sheets("Pivot").Select
    Cells(14, 10).Select
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
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Pick Type")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Pull Location"), "Count of Pull Location", xlCount
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Units"), "Sum of Units", xlSum
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Pick Type")
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable3").CompactLayoutRowHeader = "Bin Replenishment in Reserve"
    
    
    Range("A1").Select
    Columns("A:A").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Overflow", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("A1").Select

    
    
    Sheets("page").Select
    Columns("AN:AZ").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "page!R1C40:R1048576C52", Version:=6).CreatePivotTable TableDestination:= _
        "ReplenInfo!R1C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("ReplenInfo").Select
    Cells(1, 1).Select
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
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Zone")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Units"), "Cases", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Units"), "Total Units", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Units"), "Density", xlAverage
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Density")
        .NumberFormat = "0.0"
    End With
    ActiveSheet.PivotTables("PivotTable1").CompactLayoutRowHeader = _
        "Available by Zone"
    
    
        Range("E1").Select
    ActiveWorkbook.Worksheets("ReplenInfo").PivotTables("PivotTable1").PivotCache. _
        CreatePivotTable TableDestination:="ReplenInfo!R1C6", TableName:= _
        "PivotTable2", DefaultVersion:=6
    Sheets("ReplenInfo").Select
    Cells(1, 6).Select
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
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Zone")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Pull Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Units"), "Cases", xlCount
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Units"), "Total Units", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Units"), "Density", xlAverage
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Density")
        .NumberFormat = "0.0"
    End With
    Range("I3").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Pull Location").AutoSort _
        xlDescending, "Density", ActiveSheet.PivotTables("PivotTable2"). _
        PivotColumnAxis.PivotLines(3), 1
    
    Columns("A:I").EntireColumn.AutoFit
    
    Range("M1").Formula2R1C1 = "=page!C[27]"
    Range("N1").Formula2R1C1 = "=page!C[32]"
    Columns("M:N").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Range("$M:$N").AutoFilter Field:=1, Criteria1:="0"
    Range("M1:N1").Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.Range("$M:$N").AutoFilter Field:=1
    ActiveWorkbook.Worksheets("ReplenInfo").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ReplenInfo").AutoFilter.Sort.SortFields.Add2 Key:= _
        Columns("N:N"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ReplenInfo").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    Range("M1").Formula2R1C1 = "Case at Location"
    Range("N1").Formula2R1C1 = "Units"
    Columns("M:N").EntireColumn.AutoFit
    
    Range("A1").Select

    Sheets("Pivot").Select
    Range("A1").Select
    
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton3_Click()

    Unload UserForm1

MsgBox ("Macro Ended")

End Sub


