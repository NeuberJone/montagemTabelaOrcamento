Sub montaDescricao()

    Application.ScreenUpdating = False
    
    ' Declaração de variáveis
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("Especificações") ' Define a planilha Planilha de Testes como objeto de planilha
    
    desbloqueiaM9eM10
    
    dimensionaColunas
    limpaAreaDescricao

    formataDescricao
    preencheDescricao
    adicionaFormulasDescricao

    Range("L5:O5").Select
    
    Application.ScreenUpdating = True
End Sub