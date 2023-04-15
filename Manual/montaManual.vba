Sub montaManual()

    Application.ScreenUpdating = False
    
    ' Declaração de variáveis
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("Especificações") ' Define a planilha Planilha de Testes como objeto de planilha
    Set rng = ws.Range("K2:K29")
    
    desbloqueiaM9eM10
    
    dimensionaColunas
    limpaAreaProduto
    
    formataManual
    preencheManual
    adicionaFormulaManual
    adicionaValidacaoDadosManual
    
    removeValidacaoDados rng
    
    Application.ScreenUpdating = True
End Sub