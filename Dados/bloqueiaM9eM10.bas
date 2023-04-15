Sub bloqueiaM9eM10()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Planilha ativa
    
    ' Define o intervalo de todas as células na planilha
    Set rng = ws.UsedRange
    
    ' Desbloqueia todas as células no intervalo definido
    rng.Locked = False
    
    ' Bloquear as células M9 e M10
    ws.Range("M9:M10").Locked = True
    
    ' Proteger a planilha
    ws.Protect
End Sub