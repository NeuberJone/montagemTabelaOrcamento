Sub adicionaFormulasDescricao()

    Application.ScreenUpdating = False

    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Especificações")

    Range("S4").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(IF(R[1]C[-7]="""",""Nome do Material"",R[1]C[-7]),"""")"
    Range("S5").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Tamanho: "",R[5]C[-4])"
    Range("S6").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Largura: "",R[4]C[-7],""cm"")"
    
    Range("S7").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Profundidade: "",R[3]C[-6],""cm"")"
    Range("S8").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Altura: "",R10C14,""cm"")"
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Altura: "",R[2]C[-5],""cm"")"

    If Range("L16").Value = "Alça" Then
        Range("S9").Select
        ActiveCell.FormulaR1C1 = _
            "=CONCATENATE(""Alça de "",IF(R[7]C[-6]=""Selecione"","""",IF(R[7]C[-6]=""Manual"",R[7]C[-5],R[7]C[-6])),"" (consulte as cores disponíveis)"")"
    Else
        Range("S9").Select
        ActiveCell.FormulaR1C1 = "=IF(R[7]C[-6]=0,""Info Adicional"",R[7]C[-6])"
    End If

    Range("S10").Select
    ActiveCell.FormulaR1C1 = "=""Papel: ""&R[-3]C[-6]"
    Range("S11").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[1]C[-6]="""",""Impressão: "",CONCATENATE(""Impressão: "",IF(R[1]C[-6]=""Selecione"","""",IF(R[1]C[-6]=""1x0"",R[1]C[-6]&"" - Cor (""&IF(R[1]C[-5]=""Selecione"","")"",IF(R[1]C[-5]=""Não se aplica"","")"",IF(R[1]C[-5]=""Manual"",R[1]C[-4]&"")"",R[1]C[-5]&"")""))),R[1]C[-6]&"" - Cores (Colorido)"")),IF(OR(R[3]C[-6]=""Selecione"",R[3]C[-6]=""Não se aplica""),""""," & _
        "IF(R[3]C[-6]="""","""","" - ""&R[3]C[-6]))))" & _
        ""
    Range("S12").Select
    ActiveCell.FormulaR1C1 = "=IF(R[6]C[-6]=0,"""",R[6]C[-6])"
    Range("S13").Select
    ActiveCell.FormulaR1C1 = "=IF(R[7]C[-6]=0,"""",R[7]C[-6])"
    Range("S14").Select
    ActiveCell.FormulaR1C1 = "=IF(R[8]C[-6]=0,"""",R[8]C[-6])"
    Range("S15").Select
    ActiveCell.FormulaR1C1 = "=IF(R[9]C[-6]=0,"""",R[9]C[-6])"
    Range("S16").Select
    ActiveCell.FormulaR1C1 = "=IF(R[10]C[-6]=0,"""",R[10]C[-6])"
    Range("S17").Select
    ActiveCell.FormulaR1C1 = "=IF(R[11]C[-6]=0,"""",R[11]C[-6])"
    Range("S22").Select
    ActiveCell.FormulaR1C1 = "=Dados!R6C14"
    Range("S23").Select
    ActiveCell.FormulaR1C1 = "=Dados!R7C14"
    Range("S25").Select
    ActiveCell.FormulaR1C1 = "=Dados!R9C14"
    Range("S26").Select
    ActiveCell.FormulaR1C1 = "=Dados!R10C14"
    Range("S28").Select
    ActiveCell.FormulaR1C1 = "=Dados!R12C14"
    Range("S29").Select
    ActiveCell.FormulaR1C1 = "=Dados!R13C14"
    Range("S31").Select
    ActiveCell.FormulaR1C1 = "=Dados!R15C14"
    Range("S32").Select
    ActiveCell.FormulaR1C1 = "=Dados!R16C14"
    Range("S38").Select
    ActiveCell.FormulaR1C1 = "=Dados!R21C14"
    Range("S39").Select
    ActiveCell.FormulaR1C1 = "=Dados!R22C14"
    Range("S41").Select
    ActiveCell.FormulaR1C1 = "=Dados!R24C14"
    Range("S42").Select
    ActiveCell.FormulaR1C1 = "=Dados!R25C14"
    Range("S44").Select
    ActiveCell.FormulaR1C1 = "=Dados!R27C14"
    Range("S45").Select
    ActiveCell.FormulaR1C1 = "=Dados!R28C14"
    Range("S47").Select
    ActiveCell.FormulaR1C1 = "=Dados!R30C14"
    Range("S48").Select
    ActiveCell.FormulaR1C1 = "=Dados!R31C14"

    Application.ScreenUpdating = True

End Sub