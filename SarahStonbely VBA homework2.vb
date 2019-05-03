Sub Total_Stock_Volume()

'Set variables for the stock name and the total volume for each stock
Dim Ticker As String
Dim Total_Stock_Volume As Double
Dim Summary_Table_Row1 As String
Dim Summary_Table_Row2 As Double

'Loop through down the rows
For i = 2 To 705715

'Tell it to stop when it finds a new ticker symbol
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'and set the ticker name
Ticker = Cells(i, 1).Value

End If

Next i

'Then tell it to add all of the volumes in the associated cells
For j = 2 To 705715

'tell it to loop through column G (7) and add the volume for each Ticker
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

Total_Stock_Volume = Cells(j + 1, 7)

End If

Next j

'Print the results in two columns to the right
Range("I" & Summary_Table_Row1).Value = Ticker
Range("J" & Summary_Table_Row2).Value = Total_Stock_Volume


End Sub••••ˇˇˇˇ