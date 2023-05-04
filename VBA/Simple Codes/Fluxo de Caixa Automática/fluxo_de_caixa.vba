Sub Entradas()
'
' Entradas Macro
' Lança Macros em caixa
'

'
    Sheets("Receitas").Select
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E14").Select
    Sheets("Fluxo de Caixa Pessoal").Select
    Range("B4:F4").Select
    Selection.Copy
    Sheets("Receitas").Select
    Range("B4:F4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Range("B4:F4").Select
    Selection.Font.Bold = False
    Selection.Font.Size = 11
    Range("B4").NumberFormat = "dd-mm-yyyy;@"
    ' Formatar como kz Angola a baixo
    Range("F4").NumberFormat = "#,##0.00 [$Kz-pt-AO]"
    Range("D7").Select
    Sheets("Fluxo de Caixa Pessoal").Select
    Range("C6").Select
    Sheets("Fluxo de Caixa Pessoal").Select
    Range("B4:F4").Select
    Selection.ClearContents

    
    
End Sub


Sub Saidas()
'
' Saídas Macro
' Lança Saídas em caixa
'

'
    Sheets("Despesas").Select
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E14").Select
    Sheets("Fluxo de Caixa Pessoal").Select
    Range("B4:F4").Select
    Selection.Copy
    Sheets("Despesas").Select
    Range("B4:F4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:B").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Range("B4:F4").Select
    Selection.Font.Bold = False
    Selection.Font.Size = 11
    Range("B4").NumberFormat = "dd-mm-yyyy;@"
    ' Formatar como kz Angola a baixo
    Range("F4").NumberFormat = "#,##0.00 [$Kz-pt-AO]"
    ' Código Redundante Abaixo
    ' Range("D7").Select
    ' Sheets("Fluxo de Caixa Pessoal").Select
    ' Range("C6").Select
    Sheets("Fluxo de Caixa Pessoal").Select
    Range("B4:F4").Select
    Selection.ClearContents

    
End Sub














