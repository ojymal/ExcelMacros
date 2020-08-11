Attribute VB_Name = "Module1"





Public Sub Jam21DP()
'
' Macro2 Macro
'

'
On Error Resume Next
Worksheets("Dock Door Locations_1").Activate
    Range("A1").Select
    Rows("1:5").Select
    Selection.Delete Shift:=xlUp

    Selection.End(xlDown).Select
    ActiveCell.EntireRow.Resize(ActiveCell.EntireRow.Rows.Count + 2, ActiveCell.EntireRow.Columns.Count).Offset(-2, 0).Select
    Selection.Delete Shift:=xlUp
    
    Range("L1").Select
    ActiveCell.Formula2R1C1 = _
        "=COUNTA(UNIQUE(FILTER(IF(C[-10]=""IB-DD-EXP-STG"","""",C[-7]),(IF(C[-10]=""IB-DD-EXP-STG"","""",C[-7])<>""""))))-2"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-6])"
    
Worksheets("Pallet and Case PDI Location(").Activate
    Range("A1").Select
    Rows("1:5").Select
    Selection.Delete Shift:=xlUp

    Selection.End(xlDown).Select
    ActiveCell.EntireRow.Resize(ActiveCell.EntireRow.Rows.Count + 2, ActiveCell.EntireRow.Columns.Count).Offset(-2, 0).Select
    Selection.Delete Shift:=xlUp
    
    Range("L1").Select
    ActiveCell.Formula2R1C1 = "=COUNTA(UNIQUE(FILTER(C[-7],C[-7]<>"""")))-1"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-6])"
    
    
Worksheets("PE 001_4").Activate
    Range("A1").Select
    Rows("1:5").Select
    Selection.Delete Shift:=xlUp

    Selection.End(xlDown).Select
    ActiveCell.EntireRow.Resize(ActiveCell.EntireRow.Rows.Count + 2, ActiveCell.EntireRow.Columns.Count).Offset(-2, 0).Select
    Selection.Delete Shift:=xlUp
    
    Range("L1").Select
    ActiveCell.Formula2R1C1 = "=COUNTA(UNIQUE(FILTER(C[-7],C[-7]<>"""")))-1"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-6])"
    
    
    
Worksheets("QC Locations_6").Activate
    Range("A1").Select
    Rows("1:5").Select
    Selection.Delete Shift:=xlUp

    Selection.End(xlDown).Select
    ActiveCell.EntireRow.Resize(ActiveCell.EntireRow.Rows.Count + 2, ActiveCell.EntireRow.Columns.Count).Offset(-2, 0).Select
    Selection.Delete Shift:=xlUp
    
    Range("L1").Select
    ActiveCell.Formula2R1C1 = _
        "=IF((COUNTA(UNIQUE(FILTER(IF(C[-9]=""QC-N12"",0,C[-7]),C[-7]<>"""")))-2)<0,0,COUNTA(UNIQUE(FILTER(IF(C[-9]=""QC-N12"",0,C[-7]),C[-7]<>"""")))-2)"
    Range("M1").Select
    ActiveCell.Formula2R1C1 = "=SUM(IF(C[-10]=""QC-N12"",0,C[-6]))"

    On Error GoTo 0

Worksheets("Dock Door Locations_1").Activate
    Sheets.Add(Before:=ActiveSheet).Name = "Summary"
    Worksheets("Summary").Activate
    
    Range("A1").Formula2R1C1 = "Dock"
    Range("A2").Formula2R1C1 = "PDI"
    Range("A3").Formula2R1C1 = "PE"
    Range("A4").Formula2R1C1 = "QC"
    
    Range("B1").Formula2R1C1 = "='Dock Door Locations_1'!RC[10]:RC[11]"
    Range("B2").Select
    ActiveCell.Formula2R1C1 = _
        "='Pallet and Case PDI Location('!R[-1]C[10]:R[-1]C[11]"
    Range("B3").Select
    ActiveCell.Formula2R1C1 = "='PE 001_4'!R[-2]C[10]:R[-2]C[11]"
    Range("B4").Select
    ActiveCell.Formula2R1C1 = "='QC Locations_6'!R[-3]C[10]:R[-3]C[11]"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("F1").Select
        ActiveCell.FormulaR1C1 = _
        "=""Received not stowed: ""&IF(R[4]C[-3]=0,0,TEXT(R[4]C[-3],""#,#""))&"" units (""&IF(R[4]C[-4]=0,0,TEXT(R[4]C[-4],""#,#""))&"" pallet(s)).  Breakdown: Dock: ""&IF(RC[-3]=0,0,TEXT(RC[-3],""#,#""))&"" units (""&IF(RC[-4]=0,0,TEXT(RC[-4],""#,#""))&"" pallet(s)), PDI: ""&IF(R[1]C[-3]=0,0,TEXT(R[1]C[-3],""#,#""))&"" units (""&IF(R[1]C[-4]=0,0,TEXT(R[1]C[-4],""#,#""))&"" " & _
        "pallet(s)), PE: ""&IF(R[2]C[-3]=0,0,TEXT(R[2]C[-3],""#,#""))&"" units (""&IF(R[2]C[-4]=0,0,TEXT(R[2]C[-4],""#,#""))&"" pallet(s)). QC: ""&IF(R[3]C[-3]=0,0,TEXT(R[3]C[-3],""#,#""))&"" units (""&IF(R[3]C[-4]=0,0,TEXT(R[3]C[-4],""#,#""))&"" pallet(s))"""
    Range("A1").Select

End Sub




