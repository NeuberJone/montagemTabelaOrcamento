Sub montaInformacoesDoCliente()

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("Especificações") ' Define a planilha Planilha de Testes como objeto de planilha
    Set rng = ws.Range("B2:I21")
    
    desbloqueiaM9eM10
    
    dimensionaColunas
    limpaAreaCliente

    formataInformacoesDoCliente
    preencheInformacoesCliente
    
    removeValidacaoDados rng
    
    Range("C7:H7").Select
Application.ScreenUpdating = True

End Sub