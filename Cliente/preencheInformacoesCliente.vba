Sub preencheInformacoesCliente()

    Application.ScreenUpdating = False

    ' Declaração de variáveis
    Dim ws As Worksheet
    Dim rng As Range

    
    Set ws = ThisWorkbook.Sheets("Especificações")  ' Define a planilha Planilha de Testes como objeto de planilha
    
    Set rng = ws.Range("B2:I2")             ' Define o intervalo de células a ser editado
    rng.Value = "Informações do cliente"    ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("C4:D4")             ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Número da Proposta"        ' Insere o texto na célula atualmente selecionada
    
    Set rng = ws.Range("C6:H6")             ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Nome do Cliente"           ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("C9:H9")             ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Empresa"                   ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("C12")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Telefone"                  ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("C14")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Email"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("C16:D16")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Forma de Pagamento"       ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("C18:D18")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Previsão de entrega"         ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("E18:H18")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "5" 

    Set rng = ws.Range("C20:E20")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Validade do Orçamento"    ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("F20:H20")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "7 dias"                    ' Insere o texto na célula atualmente selecionada

    Application.ScreenUpdating = True
    
End Sub
