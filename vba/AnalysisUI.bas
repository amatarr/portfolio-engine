Option Explicit

' Each analysis sheet has its own spot selector, named Spot_<key>. Its
' dropdown offers "Follow Home" plus every contract in the book; changing
' it recomputes that sheet only. Each analysis sheet's own module calls
' OnSheetChange from Worksheet_Change.
Public Sub OnSheetChange(ByVal Target As Range)
    Dim key As String, sheetName As String

    If Target.Cells.CountLarge > 1 Then Exit Sub
    sheetName = Target.Worksheet.Name
    key = SpotKey(sheetName)
    If key = "" Then Exit Sub

    On Error GoTo Done
    If Intersect(Target, ThisWorkbook.Names("Spot_" & key).RefersToRange) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    RunPython "import xlwings_refresh; xlwings_refresh.refresh_sheet('" & sheetName & "')"

Done:
    Application.EnableEvents = True
End Sub

' Must match SHEETS in vol_calculator/xlwings_analysis.py.
Private Function SpotKey(sheetName As String) As String
    Select Case sheetName
        Case "Greeks": SpotKey = "Greeks"
        Case "Spot&Vol Ladder": SpotKey = "Ladder"
        Case "Heat Maps": SpotKey = "HeatMaps"
        Case "Stress": SpotKey = "Stress"
        Case "Payoff": SpotKey = "Payoff"
        Case "Breakevens": SpotKey = "Breakevens"
    End Select
End Function
