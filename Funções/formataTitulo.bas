Function formataTitulo(rng As Range)
    
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 20                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
    End With
    
    aplicaBordaExterna rng
    
End Function