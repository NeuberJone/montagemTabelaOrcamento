Sub formataSacola()

    Application.ScreenUpdating = False

    ' Declaração de variáveis
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = ThisWorkbook.Sheets("Especificações") ' Define a planilha Planilha de Testes como objeto de planilha
    
    
    Set rng = ws.Range("K2:P2")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                   ' Mescla as células
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 20                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
    End With

    Set rng = ws.Range("L4:O4")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                   ' Mescla as células
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

    Set rng = ws.Range("L5:O5")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                   ' Mescla as células
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = False                      ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        '.Interior.Color = RGB(217, 217, 217)   ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("L7")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M7:O7")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                   ' Mescla as células
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = False                      ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        '.Interior.Color = RGB(217, 217, 217)   ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

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

    Set rng = ws.Range("L12")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M12")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("N12")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("O12")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("L14")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M14:O14")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                   ' Mescla as células
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = False                      ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        '.Interior.Color = RGB(217, 217, 217)   ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("L16")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M16")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("N16:O16")                  ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                     ' Mescla as células
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

    Set rng = ws.Range("L18")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M18:O18") ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                     ' Mescla as células
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

    Set rng = ws.Range("L20")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M20:O20")                  ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                     ' Mescla as células
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

    Set rng = ws.Range("L22")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M22:O22")                  ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                     ' Mescla as células
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

    Set rng = ws.Range("L24")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M24:O24")                  ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                     ' Mescla as células
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





        Set rng = ws.Range("L24")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M24:O24")                  ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                     ' Mescla as células
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

    Set rng = ws.Range("L26")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M26:O26")                  ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                     ' Mescla as células
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

    Set rng = ws.Range("L28")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("M28:O28")                  ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                     ' Mescla as células
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

    Set rng = ws.Range("K2:P29")
    ' Aplica a borda externa ao intervalo de células
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous ' Define o estilo de linha da borda superior como contínuo
        .Weight = xlThin ' Define a espessura da borda superior como fina
        .Color = RGB(0, 0, 0) ' Define a cor da borda superior como preto
    End With
    
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous ' Define o estilo de linha da borda inferior como contínuo
        .Weight = xlThin ' Define a espessura da borda inferior como fina
        .Color = RGB(0, 0, 0) ' Define a cor da borda inferior como preto
    End With

    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous ' Define o estilo de linha da borda esquerda como contínuo
        .Weight = xlThin ' Define a espessura da borda esquerda como fina
        .Color = RGB(0, 0, 0) ' Define a cor da borda esquerda como preto
    End With

    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous ' Define o estilo de linha da borda direita como contínuo
        .Weight = xlThin ' Define a espessura da borda direita como fina
        .Color = RGB(0, 0, 0) ' Define a cor da borda direita como preto
    End With

    Application.ScreenUpdating = True
    
End Sub
