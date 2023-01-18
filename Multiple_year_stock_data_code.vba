Sub stock()

Dim ticker As String
Dim summaryrow As Integer
Dim volume As Double
Dim change As Double
Dim per_change As Double
Dim max As Double
Dim maxticker As String
Dim min As Double
Dim minticker As String
Dim maxvol As Double
Dim maxvolticker As String

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

volume = 0
change = 0
LastRow = Cells(Rows.Count, 2).End(xlUp).Row
summaryrow = 2
first = Cells(2, 3).Value

For i = 2 To LastRow

If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    ticker = Cells(i, 1).Value
    volume = volume + Cells(i, 7).Value
    change = Cells(i, 6).Value - first
    per_change = (Cells(i, 6).Value - first) / first

    Range("I" & summaryrow).Value = ticker
    Range("J" & summaryrow).Value = change
    Range("K" & summaryrow).Value = per_change
    Range("L" & summaryrow).Value = volume
    
    Range("I" & summaryrow).NumberFormat = "General"
    Range("J" & summaryrow).NumberFormat = "0.00"
    Range("K" & summaryrow).NumberFormat = "0.00%"
    Range("L" & summaryrow).NumberFormat = "#,##0"
    
    summaryrow = summaryrow + 1
    volume = 0
    change = 0
    first = Cells(i + 1, 3).Value
    

Else
    volume = volume + Cells(i, 7).Value
    
End If

Next i

Set R = Range("J2:J5001")
For Each Cell In R

If Cell.Value < 0 Then
    Cell.Interior.ColorIndex = 3
End If

If Cell.Value > 0 Then
    Cell.Interior.ColorIndex = 4
End If

Next

'BONUS

Set sheet = ActiveSheet

For i = 2 To 5000
    With sheet.Cells(i, 11)
     If .Value > max Then
     max = .Value
     maxticker = .Offset(0, -2).Value
     End If
    End With
Next i

For i = 2 To 5000
    With sheet.Cells(i, 11)
     If .Value < min Then
     min = .Value
     minticker = .Offset(0, -2).Value
     End If
    End With
Next i

For i = 2 To 5000
    With sheet.Cells(i, 12)
     If .Value > maxvol Then
     maxvol = .Value
     maxvolticker = .Offset(0, -3).Value
     End If
    End With
Next i

Cells(2, 16).Value = max
Cells(2, 15).Value = tag
Cells(3, 16).Value = min
Cells(3, 15).Value = minticker
Cells(4, 16).Value = maxvol
Cells(4, 15).Value = maxvolticker

Cells(2, 16).NumberFormat = "0.00%"
Cells(3, 16).NumberFormat = "0.00%"

End Sub