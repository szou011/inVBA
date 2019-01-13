Function GetCombination(CoinsRange As Range, SumCellId As Double) As String
'updateby Extendoffice 20160506
    Dim xStr As String
    Dim xSum As Double
    Dim xCell As Range
    xSum = SumCellId
    For Each xCell In CoinsRange
        If Not (xSum / xCell < 1) Then
            xStr = xStr & Int(xSum / xCell) & " of " & xCell & "  "
            xSum = xSum - (Int(xSum / xCell)) * xCell
        End If
    Next
    GetCombination = xStr
End Function
