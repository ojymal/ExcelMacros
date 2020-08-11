VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IdleTimeReport 
   Caption         =   "Idle Time Macro (FastTrak)"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "IdleTimeReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IdleTimeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

'Jam's Idle Time Report Macro V1.0

Dim user1 As Range
Dim ws As Worksheet

Set ws = Worksheets("Names")

For Each user1 In ws.Range("Users")
    With Me.ComboBox1
        .AddItem user1.Value
    End With
    With Me.ListBox1
        .AddItem user1.Value
    End With

Next user1
ListBox1.MultiSelect = 1

End Sub




Private Sub CommandButton1_Click()

'Jam's Idle Time Report Macro Single

Unload Me

Dim tmb As String
Dim tmb1 As String
Dim thr As Integer

thr = TextBox1

tmb = ComboBox1.Value
tmb1 = Application.WorksheetFunction.Index(Range("idxUsers"), Application.WorksheetFunction.Match(tmb, Range("Users"), 0), 2)

'___

    Worksheets("IdleTime").Select
    Range("$A$1").Select
    Range("$A:$G, $N:$W").Select
    Selection.ClearContents
    Range("$G$3").FormulaR1C1 = "Running"
    Range("$M$1").FormulaR1C1 = "All Activity:"
    Range("$A$1").Select
    Application.Wait (Now + TimeValue("0:00:001"))
Application.ScreenUpdating = False
    Range("$G$3").FormulaR1C1 = "Done"
    Range("$A$1").Select
    Range("$G$1") = tmb
    Worksheets(1).Select
    Range("A1").Select
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A:$AD").AutoFilter Field:=2, Criteria1:=tmb1
    ActiveWorkbook.Worksheets("UserTransactionReport").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("UserTransactionReport").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("UserTransactionReport").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("$A:$B,$D:$E").Select
    Selection.Copy
    Sheets("Calc").Select
    Range("A1").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.NumberFormat = "[$-en-US]h:mm AM/PM;@"
    Sheets("UserTransactionReport").Select
    Range("$A:$B,$D:$E,$I:$I,$U:$V,$Y:$Y").Select
    Selection.Copy
    Sheets("IdleTime").Select
    Range("$N$1").Select
    ActiveSheet.Paste
        Columns("N:N").Select
    Selection.NumberFormat = "[$-en-US]h:mm AM/PM;@"
    Sheets("UserTransactionReport").Select
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Sheets("Calc").Select
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Gap Minutes"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Test"
    Range("A1").Select
    Selection.AutoFilter
    Range("$A$1").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "=(RC[-4]-R[-1]C[-4])*24*60"
    Range("E3").Select
    On Error Resume Next
    Selection.AutoFill Destination:=Range("E3:E" & Range("A" & Rows.count).End(xlUp).Row)
    On Error GoTo 0
    Range(Selection, Selection.End(xlDown)).Select
    Columns("E:E").Select
    Selection.NumberFormat = "0"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IFS(RC[-1]>" & thr & ",1, AND(RC[-1]<" & thr & ",R[1]C[-1]>" _
        & thr & "),2,AND(RC[-1]<" & thr & ",R[1]C[-1]<" & thr & ",R[2]C[-1]>" & thr & "),3),0)"
    Range("F2").Select
    On Error Resume Next
    Selection.AutoFill Destination:=Range("F2:F" & Range("A" & Rows.count).End(xlUp).Row)
    On Error GoTo 0
    Range(Selection, Selection.End(xlDown)).Select
    Columns("A:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Range("$A:$F").AutoFilter Field:=6, Criteria1:=Array("1", _
        "2", "3"), Operator:=xlFilterValues
    Columns("A:F").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("IdleTime").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Calc").Select
    Range("A1").Select
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Columns("A:F").Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("IdleTime").Select
    Columns("A:F").Select
    Selection.AutoFilter
        ActiveSheet.Range("$A:$F").AutoFilter Field:=6, Criteria1:="2"
    Range("E1").Select
    ActiveCell.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    ActiveSheet.Range("$A:$F").AutoFilter Field:=6, Criteria1:="3"
    Range("$A$1:$F$1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Range("$F:$F").Select
    Selection.ClearContents
    Range("A1").Select
Application.ScreenUpdating = True

End Sub
  
Private Sub CommandButton2_Click()

'Jam's Idle Time Report Macro Multi

Unload Me

Dim tmb() As Variant
Dim tmb1() As Variant
Dim thr As Integer
Dim i As Integer, count As Integer

thr = TextBox2
count = 1

    Worksheets("IdleTime").Select
    Range("$A$1").Select
    Range("$A:$G, $N:$W").Select
    Selection.ClearContents
    Range("$G$3").FormulaR1C1 = "Running"
    Range("$M$1").FormulaR1C1 = "All Activity:"
    Range("$A$1").Select
    Application.Wait (Now + TimeValue("0:00:001"))
Application.ScreenUpdating = False
    Range("$G$3").FormulaR1C1 = "Done"
    Range("$A$1").Select

'___


For i = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(i) = True Then
ReDim Preserve tmb(count)
ReDim Preserve tmb1(count)
tmb(count) = ListBox1.List(i)
tmb1(count) = Application.WorksheetFunction.Index(Range("idxUsers"), Application.WorksheetFunction.Match(tmb(count), Range("Users"), 0), 2)


Sheets.Add(After:=Sheets(Sheets.count)).Name = tmb1(count)
Worksheets(count + 4).Select


    Range("$G$1") = tmb(count)
    Worksheets(1).Select
    Range("A1").Select
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A:$AD").AutoFilter Field:=2, Criteria1:=tmb1(count)
    ActiveWorkbook.Worksheets("UserTransactionReport").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("UserTransactionReport").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("UserTransactionReport").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("$A:$B,$D:$E").Select
    Selection.Copy
    Sheets("Calc").Select
    Range("A1").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.NumberFormat = "[$-en-US]h:mm AM/PM;@"
    Sheets("UserTransactionReport").Select
    Range("$A:$B,$D:$E,$I:$I,$L:$L,$U:$V,$Y:$Y,$AA:$AA").Select
    Selection.Copy
    Worksheets(count + 4).Select
    Range("$N$1").Select
    ActiveSheet.Paste
        Columns("N:N").Select
    Selection.NumberFormat = "[$-en-US]h:mm AM/PM;@"
    Sheets("UserTransactionReport").Select
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Sheets("Calc").Select
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Gap Minutes"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Test"
    Range("A1").Select
    Selection.AutoFilter
    Range("$A$1").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "=(RC[-4]-R[-1]C[-4])*24*60"
    Range("E3").Select
    On Error Resume Next
    Selection.AutoFill Destination:=Range("E3:E" & Range("A" & Rows.count).End(xlUp).Row)
    On Error GoTo 0
    Range(Selection, Selection.End(xlDown)).Select
    Columns("E:E").Select
    Selection.NumberFormat = "0"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IFS(RC[-1]>" & thr & ",1, AND(RC[-1]<" & thr & ",R[1]C[-1]>" _
        & thr & "),2,AND(RC[-1]<" & thr & ",R[1]C[-1]<" & thr & ",R[2]C[-1]>" & thr & "),3),0)"
    Range("F2").Select
    On Error Resume Next
    Selection.AutoFill Destination:=Range("F2:F" & Range("A" & Rows.count).End(xlUp).Row)
    On Error GoTo 0
    Range(Selection, Selection.End(xlDown)).Select
    Columns("A:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Range("$A:$F").AutoFilter Field:=6, Criteria1:=Array("1", _
        "2", "3"), Operator:=xlFilterValues
    Columns("A:F").Select
    Application.CutCopyMode = False
    Selection.Copy
    Worksheets(count + 4).Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Calc").Select
    Range("A1").Select
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Columns("A:F").Select
    Selection.ClearContents
    Range("A1").Select
    Worksheets(count + 4).Select
    Columns("A:F").Select
    Selection.AutoFilter
        ActiveSheet.Range("$A:$F").AutoFilter Field:=6, Criteria1:="2"
    Range("E1").Select
    ActiveCell.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    ActiveSheet.Range("$A:$F").AutoFilter Field:=6, Criteria1:="3"
    Range("$A$1:$F$1").Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Range("$F:$F").Select
    Selection.ClearContents
    Range("A1").Select

count = count + 1


End If
Next i

Worksheets("IdleTime").Select
Application.ScreenUpdating = True

End Sub
