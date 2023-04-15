Sub desbloqueiaM9eM10()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Planilha ativa
    
    ' Desproteger a planilha (caso esteja protegida)
    If ws.ProtectContents = True Then
        ws.Unprotect
    End If

    ' Desbloquear as células M9 e M10
    ws.Range("M9:M10").Locked = False
    
End Sub
