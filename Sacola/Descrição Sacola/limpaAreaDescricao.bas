Sub limpaAreaDescricao()

    Range("X3").Select
    Selection.Copy
    Columns("R:T").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.ClearContents
    
End Sub
