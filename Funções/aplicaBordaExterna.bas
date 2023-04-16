Function aplicaBordaExterna(rng As Range)
    
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
    
End Function