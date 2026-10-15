Option Explicit

' The Positions table is a single spilled formula, so its rows can't be
' deleted directly. Select the row(s) of the legs to remove -- any cell
' in each row is enough -- and run this. Rows are counted from the
' header, found by the PositionsView name, so the table can be moved.
Public Sub RemoveSelectedLegs()
    Dim header As Range, area As Range, r As Range, picked As String

    Set header = ThisWorkbook.Names("PositionsView").RefersToRange
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Worksheet.Name <> header.Worksheet.Name Then
        MsgBox "Select the leg(s) to remove in the Positions table first.", vbExclamation, "Remove legs"
        Exit Sub
    End If

    For Each area In Selection.Areas
        For Each r In area.Rows
            If r.Row > header.Row Then picked = picked & "," & (r.Row - header.Row)
        Next r
    Next area

    If picked = "" Then
        MsgBox "Select the leg(s) to remove in the Positions table first.", vbExclamation, "Remove legs"
        Exit Sub
    End If

    RunPython "import xlwings_refresh; xlwings_refresh.remove_legs('" & Mid(picked, 2) & "')"
End Sub
