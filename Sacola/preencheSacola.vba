Sub preencheSacola()

    Application.ScreenUpdating = False

    ' Declaração de variáveis
    Dim ws As Worksheet
    Dim rng As Range

    
    Set ws = ThisWorkbook.Sheets("Especificações")  ' Define a planilha Planilha de Testes como objeto de planilha
    
    Set rng = ws.Range("K2:P2")             ' Define o intervalo de células a ser editado
    rng.Value = "Sacola"                  ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L4:O4")             ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Nome do Material"        ' Insere o texto na célula atualmente selecionada
    
    Set rng = ws.Range("L7")             ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Papel"           ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L9")             ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Largura"           ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("M9")             ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Profundidade"                   ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("N9")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Altura"                  ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("O9")               ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Tamanho"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L12")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Cores"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("M12")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Selecione"         ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L14")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Lados"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("M14:O14")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Selecione"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L16")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Alça"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("M16")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Selecione"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L18")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento 1"                     ' Insere o texto na célula atualmente selecionada
    Set rng = ws.Range("M18:O18")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento / Complemento 1"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L20")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento 2"                     ' Insere o texto na célula atualmente selecionada
    Set rng = ws.Range("M20:O20")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento / Complemento 2"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L22")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento 3"                     ' Insere o texto na célula atualmente selecionada
    Set rng = ws.Range("M22:O22")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento / Complemento 3"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L24")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento 4"                     ' Insere o texto na célula atualmente selecionada
    Set rng = ws.Range("M24:O24")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento / Complemento 4"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L26")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento 5"                     ' Insere o texto na célula atualmente selecionada
    Set rng = ws.Range("M26:O26")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento / Complemento 5"                     ' Insere o texto na célula atualmente selecionada

    Set rng = ws.Range("L28")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento 6"                     ' Insere o texto na célula atualmente selecionada
    Set rng = ws.Range("M28:O28")           ' Define o intervalo de células a serem mescladas como objeto de intervalo
    rng.Value = "Acabamento / Complemento 6"                     ' Insere o texto na célula atualmente selecionada

    Application.ScreenUpdating = True
    
End Sub
