Sub removeValidacaoDados(rng As Range)
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Especificações")
    
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    
    With rng.Validation
        .Delete
    End With
    
    Application.ScreenUpdating = True
    Range("L5:O5").Select
End Sub
