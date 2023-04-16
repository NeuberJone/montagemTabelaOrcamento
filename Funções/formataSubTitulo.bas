Function formataSubTitulo(rng As Range)
    Dim cell As Range
    
    ' Loop através de cada célula no intervalo de células passado como parâmetro
    For Each cell In rng
        With cell
            .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
            .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
            .Font.Bold = True                       ' Aplica negrito
            .Font.Name = "Calibri"                  ' Define a fonte como Calibri
            .Font.Size = 11                         ' Define o tamanho do corpo como 20
            .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
            .Borders.Weight = xlThin                ' Define a espessura da borda como fina
            .Interior.Color = RGB(217, 217, 217)    ' Define o preenchimento com a cor #D9D9D9 (Cinza)
        End With
        aplicaBordaExterna rng
    Next cell
End Function