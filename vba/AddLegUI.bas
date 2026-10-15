Option Explicit

' Add Leg behaves like the Streamlit form: picking a Type, Commodity or
' Contract fills every field with a suggestion that can be overwritten.
' Suggestions are built from the _Data sheet (ContractQuotes,
' ExpiryCandidates), which every refresh rewrites, so they fill instantly
' without calling Python. Only a Future Spot change runs Python, to
' recompute the whole model around the new spot.

Private Const PY_REFRESH As String = "import xlwings_refresh; xlwings_refresh.main()"

' The configurable rows, top to bottom. RowLabels and Suggestions return
' one entry per row, in this same order.
Private Function RowNames() As Variant
    RowNames = Array("AddLeg_Quantity", "AddLeg_Quantity2", "AddLeg_Strike1", "AddLeg_Strike2", _
        "AddLeg_Strike3", "AddLeg_ExpiryChoice", "AddLeg_Expiry", "AddLeg_Expiry2", _
        "AddLeg_Vol1", "AddLeg_Vol2", "AddLeg_Vol3", "AddLeg_OptionType", "AddLeg_EntryPrice")
End Function

' An empty label means this type doesn't use the row: it is greyed out and cleared.
Private Function RowLabels(legType As String) As Variant
    Select Case legType
        Case "Future"
            RowLabels = Array("Quantity", "", "", "", "", "", "", "", "", "", "", "", "")
        Case "Call", "Put"
            RowLabels = Array("Quantity", "", "Strike", "", "", "Suggested expiry", "Expiry (days)", "", _
                "Vol %", "", "", "", "Entry price (blank = Mark to Market)")
        Case "Call Spread", "Put Spread"
            RowLabels = Array("Quantity", "", "Strike long", "Strike short", "", "Suggested expiry", _
                "Expiry (days)", "", "Vol % long", "Vol % short", "", "", "")
        Case "Straddle"
            RowLabels = Array("Quantity", "", "Strike", "", "", "Suggested expiry", "Expiry (days)", "", _
                "Vol % call", "Vol % put", "", "", "")
        Case "Strangle"
            RowLabels = Array("Quantity", "", "Strike put", "Strike call", "", "Suggested expiry", _
                "Expiry (days)", "", "Vol % put", "Vol % call", "", "", "")
        Case "Butterfly"
            RowLabels = Array("Quantity", "", "Strike low", "Strike mid", "Strike high", "Suggested expiry", _
                "Expiry (days)", "", "Vol % low", "Vol % mid", "Vol % high", "Option type", "")
        Case "Collar"
            RowLabels = Array("Put quantity", "Call quantity", "Strike put", "Strike call", "", _
                "Suggested expiry", "Expiry (days)", "", "Vol % put", "Vol % call", "", "", "")
        Case "Calendar Spread"
            RowLabels = Array("Quantity", "", "Strike", "", "", "Suggested expiry (near)", _
                "Near expiry (days)", "Far expiry (days)", "Vol % near", "Vol % far", "", "Option type", "")
        Case Else
            RowLabels = Empty
    End Select
End Function

' Streamlit's own defaults per type, with strikes placed relative to the
' at-the-money strike instead of Streamlit's fixed 530.
Private Function Suggestions(legType As String, atm As Variant, expLabel As Variant, _
                             nearDays As Variant, farDays As Variant) As Variant
    Dim e As Variant
    e = Empty
    Select Case legType
        Case "Future"
            Suggestions = Array(-1, e, e, e, e, e, e, e, e, e, e, e, e)
        Case "Call", "Put"
            Suggestions = Array(1, e, Off(atm, 0), e, e, expLabel, nearDays, e, 20, e, e, e, e)
        Case "Call Spread"
            Suggestions = Array(1, e, Off(atm, -10), Off(atm, 30), e, expLabel, nearDays, e, 22, 20, e, e, e)
        Case "Put Spread"
            Suggestions = Array(1, e, Off(atm, 10), Off(atm, -30), e, expLabel, nearDays, e, 23, 25, e, e, e)
        Case "Straddle"
            Suggestions = Array(1, e, Off(atm, 0), e, e, expLabel, nearDays, e, 21, 21.5, e, e, e)
        Case "Strangle"
            Suggestions = Array(1, e, Off(atm, -30), Off(atm, 30), e, expLabel, nearDays, e, 24, 20, e, e, e)
        Case "Butterfly"
            Suggestions = Array(1, e, Off(atm, -30), Off(atm, 0), Off(atm, 30), expLabel, nearDays, e, _
                24, 21, 19, "call", e)
        Case "Collar"
            Suggestions = Array(1, -1, Off(atm, -30), Off(atm, 30), e, expLabel, nearDays, e, 24, 19, e, e, e)
        Case "Calendar Spread"
            Suggestions = Array(1, e, Off(atm, 0), e, e, expLabel, nearDays, farDays, 20, 22, e, "call", e)
    End Select
End Function

Private Function Off(atm As Variant, delta As Double) As Variant
    If IsEmpty(atm) Then Off = Empty Else Off = atm + delta
End Function

Private Function Cell(nm As String) As Range
    Set Cell = ThisWorkbook.Names(nm).RefersToRange
End Function

Private Function Hit(Target As Range, nm As String) As Boolean
    Hit = Not Intersect(Target, Cell(nm)) Is Nothing
End Function

' ContractQuotes rows are: contract, commodity, current price.
Private Function QuoteFor(contract As String) As Variant
    Dim data As Variant, i As Long
    QuoteFor = Empty
    data = Cell("ContractQuotes").Value
    If Not IsArray(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 1)) = contract Then QuoteFor = data(i, 3): Exit Function
    Next i
End Function

Private Function FirstContractFor(commodity As String) As Variant
    Dim data As Variant, i As Long
    FirstContractFor = Empty
    data = Cell("ContractQuotes").Value
    If Not IsArray(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 2)) = commodity Then FirstContractFor = data(i, 1): Exit Function
    Next i
End Function

Private Function AtTheMoney() As Variant
    Dim px As Variant
    AtTheMoney = Empty
    px = QuoteFor(CStr(Cell("AddLeg_Contract").Value))
    If IsNumeric(px) And Not IsEmpty(px) Then AtTheMoney = Application.WorksheetFunction.MRound(px, 10)
End Function

' ExpiryCandidates rows are: contract, label, days to expiry -- nearest first.
Private Sub ContractExpiries(ByRef firstLabel As Variant, ByRef firstDays As Variant, ByRef secondDays As Variant)
    Dim data As Variant, i As Long, found As Long, contract As String
    firstLabel = Empty: firstDays = Empty: secondDays = Empty
    contract = CStr(Cell("AddLeg_Contract").Value)
    data = Cell("ExpiryCandidates").Value
    If Not IsArray(data) Then Exit Sub
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 1)) = contract Then
            found = found + 1
            If found = 1 Then
                firstLabel = data(i, 2)
                firstDays = data(i, 3)
            Else
                secondDays = data(i, 3)
                Exit For
            End If
        End If
    Next i
End Sub

Private Function DaysForLabel(label As String) As Variant
    Dim data As Variant, i As Long, contract As String
    DaysForLabel = Empty
    contract = CStr(Cell("AddLeg_Contract").Value)
    data = Cell("ExpiryCandidates").Value
    If Not IsArray(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 1)) = contract And CStr(data(i, 2)) = label Then DaysForLabel = data(i, 3): Exit Function
    Next i
End Function

Public Sub FillSuggestions()
    Dim legType As String, labels As Variant, values As Variant, names As Variant
    Dim expLabel As Variant, nearDays As Variant, farDays As Variant
    Dim i As Long, prevEvents As Boolean

    legType = CStr(Cell("AddLeg_Type").Value)
    labels = RowLabels(legType)
    If IsEmpty(labels) Then Exit Sub

    ContractExpiries expLabel, nearDays, farDays
    If IsEmpty(farDays) And Not IsEmpty(nearDays) Then farDays = nearDays + 60
    values = Suggestions(legType, AtTheMoney(), expLabel, nearDays, farDays)
    names = RowNames()

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    For i = 0 To UBound(names)
        With Cell(CStr(names(i)))
            .Value = values(i)
            If labels(i) = "" Then
                .Offset(0, -1).Value = ""
                .Interior.Color = RGB(242, 242, 242)
                .Offset(0, -1).Interior.Color = RGB(242, 242, 242)
            Else
                .Offset(0, -1).Value = labels(i) & ":"
                .Interior.Color = RGB(255, 242, 204)
                .Offset(0, -1).Interior.Color = RGB(221, 235, 247)
            End If
        End With
    Next i
    Application.EnableEvents = prevEvents
End Sub

Public Sub OnHomeChange(ByVal Target As Range)
    If Target.Cells.CountLarge > 1 Then Exit Sub
    On Error GoTo Done
    Application.EnableEvents = False

    If Hit(Target, "FutureSpot") Then
        RunPython PY_REFRESH
    ElseIf Hit(Target, "AddLeg_Commodity") Then
        Cell("AddLeg_Contract").Value = FirstContractFor(CStr(Target.Value))
        FillSuggestions
    ElseIf Hit(Target, "AddLeg_Type") Or Hit(Target, "AddLeg_Contract") Then
        FillSuggestions
    ElseIf Hit(Target, "AddLeg_ExpiryChoice") Then
        Cell("AddLeg_Expiry").Value = DaysForLabel(CStr(Target.Value))
    End If

Done:
    Application.EnableEvents = True
End Sub
