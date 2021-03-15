Attribute VB_Name = "Module1"
Sub AppendToExistingOnLeft()

Range("A1").Select
For y = 1 To 10
    For x = 1 To 548
        Cells(x, y).Select
        If ActiveCell.Value <> "" Then
            ActiveCell.Value = (("<th>") & ActiveCell.Value)
        End If
    Next x
Next y

End Sub

Sub AppendToExistingOnRight()

Range("A1").Select
For y = 1 To 10
    For x = 1 To 548
        Cells(x, y).Select
        If ActiveCell.Value <> "" Then
            ActiveCell.Value = (ActiveCell.Value & ("</th>"))
        End If
    Next x
Next y

End Sub

Sub Appendtr()

Range("A1").Select
For x = 1 To 548
    Cells(x, 1).Select
    If ActiveCell.Value <> "" Then
        ActiveCell.Value = (("<tr>") & ActiveCell.Value)
    End If
Next x

Range("J1").Select
For x = 1 To 548
    Cells(x, 10).Select
    If ActiveCell.Value <> "" Then
        ActiveCell.Value = (ActiveCell.Value & ("</tr>"))
    End If
Next x
End Sub
