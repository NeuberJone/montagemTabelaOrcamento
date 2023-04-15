Sub larguraXprofundidadeXaltura()

    Application.ScreenUpdating = False

    ' Declaração de variáveis
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = ThisWorkbook.Sheets("Especificações") ' Define a planilha Planilha de Testes como objeto de planilha
    
    Set rng = ws.Range("L9")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        .Interior.Color = RGB(217, 217, 217)    ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("M9")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = False
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .NumberFormat = ";;;" ' Torna o conteúdo invisível
    End With

    Set rng = ws.Range("N9")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        .Interior.Color = RGB(217, 217, 217)    ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("O9")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        .Interior.Color = RGB(217, 217, 217)    ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("L10")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = False                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        '.Interior.Color = RGB(217, 217, 217)    ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("M10")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = False
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .NumberFormat = ";;;" ' Torna o conteúdo invisível
    End With

    Set rng = ws.Range("N10")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = False                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        '.Interior.Color = RGB(217, 217, 217)    ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("O10")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = False                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        '.Interior.Color = RGB(217, 217, 217)    ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("L9")             ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Largura"           ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("M9")             ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Profundidade"                   ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("N9")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Altura"                  ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("O9")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Tamanho"                     ' Insere o texto na célula atualmente selecionada

    Range("O10").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]="""","""",IF(RC[-2]="""",RC[-3]&"" cm"",IF(RC[-1]="""",RC[-3]&""x""&RC[-2]&"" cm"",RC[-3]&""x""&RC[-2]&""x""&RC[-1]&"" cm"")))"

    Range("S7").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Profundidade: "",R[3]C[-6],""cm"")"
    Range("S8").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Altura: "",R[2]C[-5],""cm"")"
    
    desbloqueiaM9eM10

    Range("L10").Select

    Application.ScreenUpdating = True
    
End Sub
