VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Private Sub Worksheet_Change(ByVal Target As Range)


' Checks target column for changes then adds a time stamp on designated column on the same row.
' Does not update other cells.

If Target.Column = 1 Then

Application.EnableEvents = False

Cells(Target.Row, 5).Value = Time

Application.EnableEvents = True

End If



End Sub

