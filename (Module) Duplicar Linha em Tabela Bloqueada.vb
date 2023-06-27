Sub CopiarLinhaTabela()
    Dim tbl As ListObject
    Dim linhaSelecionada As Range
    Dim novaLinha As Range
    Dim ws As Worksheet
    Dim col As ListColumn
    
    ' Verifica se uma célula está selecionada
    If TypeName(Selection) <> "Range" Then
        MsgBox "Selecione uma célula na tabela.", vbInformation
        Exit Sub
    End If
    
    ' Verifica se a célula selecionada está dentro de uma tabela
    On Error Resume Next
    Set tbl = Selection.ListObject
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Selecione uma célula dentro de uma tabela.", vbInformation
        Exit Sub
    End If
    
    ' Desprotege a planilha para permitir alterações
    Set ws = tbl.Parent
    ws.Unprotect Password:="Hus07468"
    
    ' Obtém a linha selecionada
    Set linhaSelecionada = Selection.EntireRow
    
    ' Insere uma nova linha abaixo da linha selecionada
    linhaSelecionada.Offset(1).Insert Shift:=xlDown
    
    ' Obtém a nova linha criada
    Set novaLinha = linhaSelecionada.Offset(1)
    
    ' Copia os valores da linha selecionada para a nova linha
    For Each col In tbl.ListColumns
        If col.Name = "Destino" Or col.Name = "Apto" Then
            novaLinha.Cells(col.Index).Value = "'" & linhaSelecionada.Cells(col.Index).Value
        Else
            novaLinha.Cells(col.Index).Value = linhaSelecionada.Cells(col.Index).Value
        End If
    Next col
    
    ' Bloqueia a planilha com a senha "Hus07468" e permite edição de conteúdos, mas não objetos de desenho
    ActiveSheet.Protect Password:="Hus07468", Contents:=True, DrawingObjects:=False, AllowFiltering:=True
    
End Sub