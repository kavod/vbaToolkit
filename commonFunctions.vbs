Function line_col_max(myCell As Range) As Long()
'' Renvoie la dernière ligne et dernière colonne d'une feuille à partir d'une cellule de cette feuille
    Dim result(2) As Long
    If myCell(2, 1).value = "" Then
        result(1) = myCell.Row
    Else
        result(1) = myCell.End(xlDown).Row
    End If
    If myCell(1, 2) = "" Then
        result(2) = myCell.Column
    Else
        result(2) = myCell.End(xlToRight).Column
    End If
    line_col_max = result()
End Function
