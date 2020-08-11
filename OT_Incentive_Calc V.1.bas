Attribute VB_Name = "Module1"
Sub OT_Incentive_Calc()
Attribute OT_Incentive_Calc.VB_ProcData.VB_Invoke_Func = " \n14"
'
' OT_Incentive_Calc Jam_Macro V.1
'

'

Application.ScreenUpdating = False

'   Clear data on results tab and insert headers

    Sheets("Results").Select

    Columns("A:I").ClearContents
    Range("A1").FormulaR1C1 = "Employee Name"
    Range("B1").FormulaR1C1 = "Employee ID"
    Range("C1").FormulaR1C1 = "PP End Date"
    Range("D1").FormulaR1C1 = "Straight"
    Range("E1").FormulaR1C1 = "OT"
    Range("F1").FormulaR1C1 = "Total Hours"
    Range("G1").FormulaR1C1 = "Standard"
    Range("H1").FormulaR1C1 = "Actual OT"
    Range("I1").FormulaR1C1 = "Incentive"
    Range("A2").Select
    
    
    
    
'   Format Report on OT Repot Tab
    
    Sheets("Staffing Report").Select
    Rows("1:2").Delete Shift:=xlUp
    
    Range("B:B").NumberFormat = "General"
    Range("B:B").Value = Range("B:B").Value
    
    
    
    
'   Format Report on OT Report Tab and paste results in Results Tab

    Sheets("OT Report").Select
    Rows("1:1").Delete Shift:=xlUp
    Range("A1").End(xlDown).EntireRow.Resize(ActiveCell.EntireRow.Rows.Count + 3, ActiveCell.EntireRow.Columns.Count).Offset(-3, 0).Delete Shift:=xlUp
    Columns("I:R").Delete Shift:=xlToLeft
    
    ActiveSheet.Range("$A:$H").RemoveDuplicates Columns:=3, Header:= _
        xlYes
        
    Range("B2:F2").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Sheets("Results").Select
    Range("A2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Sheets("OT Report").Select
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Copy
    Sheets("Results").Select
    Range("F2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    
    
    
'   Insert Formulas => autofill=> remove formulas and keep values
    
    Range("G2").FormulaR1C1 = _
        "=VLOOKUP(RC[-5],'Staffing Report'!C[-5]:C[13],19,FALSE)"
    Range("H2").FormulaR1C1 = "=IF(RC[-2]-RC[-1]>0,RC[-2]-RC[-1],0)"
    Range("I2").FormulaR1C1 = "=RC[-1]*R16C13"
    
    Range("G2:I2").Select
        Selection.AutoFill Destination:=Range("G2:I" & Range("A" & Rows.Count).End(xlUp).Row)
        
    Columns("A:H").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    
    
'   Final formatting and filter
    
    Columns("A:I").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Results").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Results").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("I:I"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Results").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("A:I").EntireColumn.AutoFit
    
    Range("A1").Select
    
Application.ScreenUpdating = True

End Sub

