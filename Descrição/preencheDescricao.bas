Sub preencheDescricao()

    Application.ScreenUpdating = False

    ' Declaração de variáveis
    Dim ws As Worksheet
    Dim rng As Range

    
    Set ws = ThisWorkbook.Sheets("Especificações")  ' Define a planilha Planilha de Testes como objeto de planilha
    
    Set rng = ws.Range("R2:T2")             ' Define o intervalo de células a ser editado
    rng.Value = "Descrição"                  ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("S19")
    rng.Value = "-> Valores com desconto para pagamento via PIX ou Transferência"                     ' Insere o texto na célula atualmente selecionada
    
    Set rng = ws.Range("S20")
    rng.Value = "(50% Entrada e 50% Entrega)"

    Set rng = ws.Range("S36")
    rng.Value = "-> Valores para pagamento via cartão de crédito (Sem Juros)"

    Application.ScreenUpdating = True
    
End Sub
