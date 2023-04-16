Sub larguraXaltura()

    Application.ScreenUpdating = False

    ' Declaração de variáveis
    Dim ws As Worksheet
    Dim rng As Range

    If larguraXalturaStatus = False Then

        'desbloqueiaM9eM10

        Set ws = ThisWorkbook.Sheets("Especificações") ' Define a planilha Planilha de Testes como objeto de planilha
        
        tamanhoAltura = ws.Range("N10").Value
        ws.Range("N10").Value = Empty
        
        tamanhoProfundidade = ws.Range("M10").Value
        ws.Range("N10").Value = Empty

        Set rng = ws.Range("L9")
        rng.Value = "Largura"                       ' Define o intervalo de células a serem mescladas como objeto de intervalo e coloca o texto Largura
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

        Set rng = ws.Range("M9")
        rng.Value = "Largura"
        With rng
            .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
            .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
            .Font.Bold = True                       ' Aplica negrito
            .Font.Name = "Calibri"                  ' Define a fonte como Calibri
            .Font.Size = 11                         ' Define o tamanho do corpo como 20
            .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
            .Borders.Weight = xlThin                ' Define a espessura da borda como fina
            .Interior.Color = RGB(217, 217, 217)    ' Define o preenchimento com a cor #D9D9D9 (Cinza)
        End With

        Set rng = ws.Range("M10")
        rng = tamanhoAltura
        With rng
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = False
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Interior.ColorIndex = xlNone
        End With

        Set rng = ws.Range("O9")
        rng.Value = Empty
        With rng
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Borders.LineStyle = xlNone
            '.Borders.Weight = xlNone
            .Interior.Color = xlNone
        End With

        Set rng = ws.Range("O10")
        rng.Value = Empty
        With rng
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Borders.LineStyle = xlNone
            '.Borders.Weight = xlNone
            .Interior.Color = xlNone
        End With

        Set rng = ws.Range("N9")
        rng.Value = "Tamanho"               ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

        Set rng = ws.Range("N10")
        With rng                                    ' Formatação do texto na célula atualmente selecionada
            .HorizontalAlignment = xlCenter         ' Centraliza o texto horizontalmente
            .VerticalAlignment = xlCenter           ' Centraliza o texto verticalmente
            .Font.Bold = False                      ' Remove negrito
            .Font.Name = "Calibri"                  ' Define a fonte como Calibri
            .Font.Size = 11                         ' Define o tamanho do corpo como 20
            .Borders.LineStyle = xlContinuous       ' Aplica uma borda contínua
            .Borders.Weight = xlThin                ' Define a espessura da borda como fina
            .Interior.Color = xlNone                ' Remove preenchimento
        End With

        Range("N10").FormulaR1C1 = "=IFS(RC[-2]="""","""",RC[-1]="""",RC[-2]&"" cm"",RC[-1]<>"""",RC[-2]&""x""&RC[-1]&"" cm"")"

        Range("S7").FormulaR1C1 = "=CONCATENATE(""Altura: "",R[3]C[-6],""cm"")"

        Range("S8").FormulaR1C1 = Empty
        
        larguraXalturaStatus = True

        Range("L10").Select

        Application.ScreenUpdating = True

    Else

        MsgBox "A formatação de tamanho Largura x Altura já está aplicada."

        Range("L10").Select

        Application.ScreenUpdating = True

        Exit Sub
        
    End If
    
End Sub
