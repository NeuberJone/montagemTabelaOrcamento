Sub restauraPlanilha()

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim coluna As Range
    
    ' Define a planilha Planilha de Testes como objeto de planilha
    Set ws = ThisWorkbook.Sheets("Especificações")
    
    ' Limpa a formatação das colunas A a Z (26 Colunas)
    ws.Range("A:Z").ClearFormats
    
    ' Apaga o conteúdo das células das colunas A a Z (26 Colunas)
    ws.Range("A:Z").ClearContents

    ' Loop que altera a largura das colunas A a Z (26 Colunas) para o tamanho padrão (8.43 pontos) (64 pixels)
    For i = 1 To 26
        ws.Columns(Chr(64 + i)).ColumnWidth = 8.43
    Next i
    
    formatacaoDeCondicionais
    
    ws.Rows(2).RowHeight = 15
    
    Range("A1").Select
    
    atualizaTudo

    Application.ScreenUpdating = True
    
End Sub