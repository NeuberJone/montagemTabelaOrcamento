Sub formataManual()

    Application.ScreenUpdating = False

    ' Declaração de variáveis
    Dim ws As Worksheet
    Dim rngArray As Variant
    
    Set ws = ThisWorkbook.Sheets("Especificações")  ' Define a planilha "Especificações" como objeto de planilha
    
    rngArray = Array("K2", "P2", "L4", "O4", "L5", "O5", "M7", "O7", "M14", "O14", "M16", "O16", "M18", "O18", _
                     "M20", "O20", "M22", "O22", "M24", "O24", "M26", "O26", "M28", "O28")
    
    ' Chamar a sub-rotina mesclaCelulas passando o array de células como argumento
    mesclaCelulas rngArray, ws

    formataTitulo ws.Range("K2:P2")
    
    formataSubTitulo ws.Range("L4:O4,M9,N9,O9,L7,L9,L12,L14,L16,L18,L20,L22,L24,L26,L28")
    
    formataCelulaComum ws.Range("L5:O5,M7:O7,L10,M10,N10,O10,M12,N12,O12,M14:O14,M16:O16,M18:O18,M20:O20,M22:O22,M24:O24,M26:O26,M28:O28")

    aplicaBordaExterna ws.Range("K2:P29")

    Application.ScreenUpdating = True
    
End Sub
