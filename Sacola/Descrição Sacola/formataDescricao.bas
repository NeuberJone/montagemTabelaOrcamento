Sub formataDescricao()

    Application.ScreenUpdating = False

    ' Declaração de variáveis
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = ThisWorkbook.Sheets("Especificações") ' Define a planilha Planilha de Testes como objeto de planilha
    
    Set rng = ws.Range("R2:T2")                 ' Define o intervalo de células a serem mescladas como objeto de intervalo
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

    Set rng = ws.Range("R34:T34")
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

    Set rng = ws.Range("R2:T49")
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
