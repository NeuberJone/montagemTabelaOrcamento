Sub dimensionaColunas()
    
    Application.ScreenUpdating = False

    ' Declaração de variáveis

    Dim ws As Worksheet
    Dim intervalos As Variant
    Dim intervalo As Variant
    Dim limiteInicial As Integer
    Dim limiteFinal As Integer

    ' Define a planilha Planilha de Testes como objeto de planilha
    Set ws = ThisWorkbook.Sheets("Especificações")

    ' Array com os intervalos de colunas para cada largura de 2.14 no loop
    intervalos = Array("1 To 2", "9 To 11", "16 To 18", "20 To 22")
    
    ' Loop para colocar o tamanho de 2.14 nas colunas
    ' Loop para cada intervalo de colunas
    
    For Each intervalo In intervalos
        ' Extrai os limites do intervalo
        limiteInicial = Split(intervalo, " To ")(0)
        limiteFinal = Split(intervalo, " To ")(1)
        
        ' Loop para alterar a largura das colunas dentro do intervalo
        Dim i As Integer
        For i = limiteInicial To limiteFinal
            ws.Columns(Chr(64 + i)).ColumnWidth = 2.14
        Next i
    Next intervalo

    ' Array com os intervalos de colunas para cada largura de 8.43 no loop
    intervalos = Array("3 To 3", "6 To 8", "23 To 26")
    
    ' Loop para colocar o tamanho de 8.43 nas colunas
    ' Loop para cada intervalo de colunas
    For Each intervalo In intervalos
        ' Extrai os limites do intervalo
        limiteInicial = Split(intervalo, " To ")(0)
        limiteFinal = Split(intervalo, " To ")(1)
        
        ' Loop para alterar a largura das colunas dentro do intervalo
        For i = limiteInicial To limiteFinal
            ws.Columns(Chr(64 + i)).ColumnWidth = 8.43
        Next i
    Next intervalo

    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 1.43
    ws.Columns("S").ColumnWidth = 61.43

    For i = 12 To 15
        ws.Columns(Chr(64 + i)).ColumnWidth = 13.57
    Next i

    ' Altere a altura da linha 2 (Linha de títulos)
    ws.Rows(2).RowHeight = 26.25

    Application.ScreenUpdating = True
    
End Sub
