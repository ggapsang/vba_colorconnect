Function COLORCONNECT(rng As Range, clr As Range) As String
    Dim cell As Range
    Dim result As String
    For Each cell In rng
        If cell.Interior.ColorIndex = clr.Interior.ColorIndex Then
            result = result & cell.Value & " & "
        End If
    Next cell
    If Len(result) > 3 Then
        COLORCONNECT = Left(result, Len(result) - 3)
    Else
        COLORCONNECT = ""
    End If
End Function
