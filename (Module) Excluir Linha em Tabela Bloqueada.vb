Sub ExcluirLinhaUnica()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim linhaSelecionada As Range
    Dim unicoValue As Variant
    
    ' Desbloqueia a planilha com a senha "Hus07468"
    ActiveSheet.Unprotect Password:="Hus07468"
    
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
    
    ' Obtém a linha selecionada
    Set linhaSelecionada = Selection.EntireRow
    
    ' Obtém o valor da coluna "Único"
    unicoValue = linhaSelecionada.Cells(1, tbl.ListColumns("Único").Index).Value
    
    ' Verifica se o valor é superior a 1
    If unicoValue > 1 Then
        ' Permite a exclusão da linha
        linhaSelecionada.Delete
    Else
        ' Gera mensagem de erro para o usuário
        MsgBox "O valor é único e não pode ser excluído.", vbExclamation
    End If
    
    ' Bloqueia a planilha com a senha "Hus07468" e permite edição de conteúdos, mas não objetos de desenho
    ActiveSheet.Protect Password:="Hus07468", Contents:=True, DrawingObjects:=False, AllowFiltering:=True
End Sub