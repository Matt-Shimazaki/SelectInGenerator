Sub create_in()
Dim i As Integer
Dim LastRow As Integer
Dim in_text As String

in_text = in_text & "in ("
LastRow = Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row

For i = 1 To LastRow

    If i = 1 Then
        in_text = in_text & "'" & Sheets("Sheet1").Cells(i, 1).Value
    Else
        in_text = in_text & "','" & Sheets("Sheet1").Cells(i, 1).Value    
    End If

Next i

in_text = in_text & "')"
Sheets("Sheet1").Range("C1").Value = in_text

End Sub
