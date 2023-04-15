Sub limpaAreaProduto()

    Range("X3").Select
    Selection.Copy
    Columns("K:P").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.ClearContents
    
End Sub