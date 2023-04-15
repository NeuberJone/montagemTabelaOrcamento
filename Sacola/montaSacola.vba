Sub montaSacola()

    Application.ScreenUpdating = False
    
    ' Declaração de variáveis
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("Especificações") ' Define a planilha Planilha de Testes como objeto de planilha
    
    desbloqueiaM9eM10
    
    limpaAreaProduto
    dimensionaColunas
    
    formataSacola
    preencheSacola

    adicionaFormulasSacola
    adicionaValidacaoDadosSacola

    Application.ScreenUpdating = True
End Sub