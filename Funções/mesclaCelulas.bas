Sub mesclaCelulas(rngArray As Variant, ws As Worksheet)
    Dim i As Integer
    Dim rng As Range
    Dim cell As Range
    
    For i = 0 To UBound(rngArray) Step 1 ' Começa do índice 0 e avança de 1 em 1
        If i Mod 2 = 0 Then ' Verifica se o índice é par
            If i + 1 <= UBound(rngArray) Then ' Verifica se há uma próxima posição no array
                Set rng = ws.Range(rngArray(i)) ' Define o range da primeira célula a ser mesclada
                Set rng = ws.Range(rng, rngArray(i + 1)) ' Expande o range para incluir a próxima célula a ser mesclada
                rng.Merge ' Mescla as células
            Else ' Caso não haja uma próxima posição no array, mescla somente a célula atual
                Set rng = ws.Range(rngArray(i))
                rng.Merge
            End If
        End If
    Next i

End Sub