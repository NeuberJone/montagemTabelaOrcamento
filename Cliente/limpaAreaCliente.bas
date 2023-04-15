Sub limpaAreaCliente()

    Range("X3").Select
    Selection.Copy
    Columns("B:I").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.ClearContents
    
End Sub