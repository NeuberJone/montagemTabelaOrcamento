Sub adicionaFormulaManual()

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Especificações")

    Range("M7:O7").Select
    ActiveCell.FormulaR1C1 = "=Dados!R35C54"
    
    Range("O10").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-3]="""","""",IF(RC[-2]="""",RC[-3]&"" cm"",IF(RC[-1]="""",RC[-3]&""x""&RC[-2]&"" cm"",RC[-3]&""x""&RC[-2]&""x""&RC[-1]&"" cm"")))"

    Range("N12").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>""1x0"",""Não se aplica"",""Selecione"")"
    
    Range("O12").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]=""Manual"",""Digite a cor"",""Não se aplica"")"

    Application.ScreenUpdating = True

    Range("L5:O5").Select
End Sub