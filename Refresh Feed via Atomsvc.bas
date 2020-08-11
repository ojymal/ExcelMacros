Attribute VB_Name = "Module1"
Private Sub UpdateAndRefreshFiles()

Range("X1").FormulaR1C1 = "Running..."
Application.Wait (Now + TimeValue("0:00:005"))

Application.ScreenUpdating = False

'Dimension Sheet Variables
Dim User, Path, AllDataPath, StartingMonth, StartingDate, StartingYear, EndingMonth, EndingDate, _
EndingYear, AllDataCode As String
Dim sht As Worksheet
Dim sheet As Worksheet
Dim con As WorkbookConnection


'Define Sheet Range
Set sht = ActiveSheet


'Dimension File to create
Dim AllDataFile As Object


'Current File Location
Path = Application.ActiveWorkbook.Path


'Define File Path with current File Path
AllDataPath = Path & "\AllDataToday.atomsvc"



'Define Date Values
StartingMonth = Application.WorksheetFunction.Text(sht.Range("B1").Value, "MM")
StartingDate = Application.WorksheetFunction.Text(sht.Range("B1").Value, "DD")
StartingYear = Application.WorksheetFunction.Text(sht.Range("B1").Value, "YYYY")
EndingMonth = Application.WorksheetFunction.Text(sht.Range("B2").Value, "MM")
EndingDate = Application.WorksheetFunction.Text(sht.Range("B2").Value, "DD")
EndingYear = Application.WorksheetFunction.Text(sht.Range("B2").Value, "YYYY")


'Define Code
AllDataCode = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?><service xmlns:atom=""http://www.w3.org/2005/Atom"" xmlns:app=""http://www.w3.org/2007/app"" xmlns=""http://www.w3.org/2007/app""><workspace><atom:title>ReportUserTransactions</atom:title><collection href=""http://sqllist3015/ReportServer?%2FQVC%2FReportUserTransactions&amp;Event=%3CALL%3E&amp;User=&amp;StartDate=" & StartingMonth & "%2F" & StartingDate & "%2F" & StartingYear & "%2000%3A00%3A00&amp;EndDate=" & EndingMonth & "%2F" & EndingDate & "%2F" & EndingYear & "%2023%3A59%3A59&amp;rs%3AParameterLanguage=&amp;rs%3ACommand=Render&amp;rs%3AFormat=ATOM&amp;rc%3AItemPath=Tablix1""><atom:title>Tablix1</atom:title></collection></workspace></service>"

'Create ATOMSVC File

Open AllDataPath For Output As #1
Print #1, AllDataCode
Close #1


'Update connection string
ActiveWorkbook.Connections("Datafeed_All_Data").DataFeedConnection.Connection = "DATAFEED;Data Source=" & AllDataPath & ";Namespaces to Include=*;Max Received Message Size=4398046511104;Integrated Security=SSPI;Keep Alive=true;Persist Security Info=false;Service Document Url=" & AllDataPath


'Update all data feeds
On Error Resume Next
For Each con In ActiveWorkbook.Connections
    If Left(con.Name, 8) = "Datafeed" Then
    Cname = con.Name
        With ActiveWorkbook.Connections(Cname).DataFeedConnection
            .Refresh
        End With
    End If
Next
On Error GoTo 0


'Update Query/Tables
Worksheets("BinReplenQuery").Range("A1").ListObject.QueryTable.Refresh BackgroundQuery:=False
Worksheets("BinPickQuery").Range("A1").ListObject.QueryTable.Refresh BackgroundQuery:=False
Worksheets("AdHocDropQuery").Range("A1").ListObject.QueryTable.Refresh BackgroundQuery:=False
Worksheets("RTSQuery").Range("A1").ListObject.QueryTable.Refresh BackgroundQuery:=False
Worksheets("RTSSortQuery").Range("A1").ListObject.QueryTable.Refresh BackgroundQuery:=False
Worksheets("IdleTimeQuery").Range("A1").ListObject.QueryTable.Refresh BackgroundQuery:=False


'Update All PivotTables
For Each sheet In ThisWorkbook.Worksheets
    For Each Pivot In sheet.PivotTables
        Pivot.RefreshTable
        Pivot.Update
    Next
Next

Range("X1").FormulaR1C1 = "Ready!"

Application.ScreenUpdating = True

End Sub



Private Sub format_lines()
'
' format_lines Macro
'

'
    Rows("13:13").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B1").Select
End Sub

Private Sub clear_format_lines()
'
' clear_format_lines Macro
'

'
    Rows("13:13").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B1").Select
End Sub

