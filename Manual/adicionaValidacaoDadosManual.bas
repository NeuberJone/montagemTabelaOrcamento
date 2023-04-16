Sub adicionaValidacaoDadosManual()

    Dim ws As Worksheet
    Dim rng As Range

    Set ws = ThisWorkbook.Sheets("Especificações")

    Set rng = ws.Range("K2:P29")

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    With rng.Validation
        .Delete
    End With

    Application.EnableEvents = True

    Range("M12").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:= _
        "Selecione,1x0,4x0,4x1,4x4"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    Range("N12").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:= _
        "Selecione,Preto,Pantone,Manual"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    Application.ScreenUpdating = True
    Range("L5:O5").Select
    
End Sub
