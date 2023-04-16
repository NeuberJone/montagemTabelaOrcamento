Function formataCelulaComum(rng As Range)
    Dim cell As Range
    
    ' Loop através de cada célula no intervalo de células passado como parâmetro
    For Each cell In rng
        With cell
            .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
            .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
            .Font.Bold = False                      ' Aplica negrito (False para não aplicar negrito, True para aplicar)
            .Font.Name = "Calibri"                  ' Define a fonte como Calibri
            .Font.Size = 11                         ' Define o tamanho do corpo como 11
            .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
            .Borders.Weight = xlThin                ' Define a espessura da borda como fina
            .Interior.Color = xlNone                ' Define a cor de preenchimento como "nenhuma"
        End With
        aplicaBordaExterna rng
    Next cell
End Function