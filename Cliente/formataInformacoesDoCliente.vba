Sub formataInformacoesDoCliente()

    Application.ScreenUpdating = False

    ' Declaração de variáveis
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = ThisWorkbook.Sheets("Especificações") ' Define a planilha Planilha de Testes como objeto de planilha
    
    Set rng = ws.Range("B2:I2")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("C4:D4")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                   ' Mescla as células
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlEdgeLeft         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        .Interior.Color = RGB(217, 217, 217)    ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("E4:H4")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("C6:H6")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("C7:H7")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("C9:H9")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("C10:H10")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("C12")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlEdgeLeft         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                      ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        .Interior.Color = RGB(217, 217, 217)   ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("D12:H12")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("C14")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlEdgeLeft         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        .Interior.Color = RGB(217, 217, 217)   ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("D14:H14")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("C16:D16")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                   ' Mescla as células
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlEdgeLeft         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                      ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        .Interior.Color = RGB(217, 217, 217)   ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("E16:H16")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                   ' Mescla as células
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = False                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        '.Interior.Color = RGB(217, 217, 217)   ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("C18:D18")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                   ' Mescla as células
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlEdgeLeft         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                       ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        .Interior.Color = RGB(217, 217, 217)   ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("E18:H18")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("C20:E20")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Merge                                   ' Mescla as células
    With rng                                    ' Formatação do texto na célula atualmente selecionada
        .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
        .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
        .Font.Bold = True                      ' Aplica negrito
        .Font.Name = "Calibri"                  ' Define a fonte como Calibri
        .Font.Size = 11                         ' Define o tamanho do corpo como 20
        .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
        .Borders.Weight = xlThin                ' Define a espessura da borda como fina
        .Interior.Color = RGB(217, 217, 217)   ' Define o preenchimento com a cor #D9D9D9 (Cinza)
    End With

    Set rng = ws.Range("F20:H20")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("B2:I21")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    With rng.Borders(xlEdgeTop)     ' Aplica a borda externa ao intervalo de células
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