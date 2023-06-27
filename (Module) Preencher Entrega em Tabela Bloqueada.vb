Sub EntregaTotal()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filtro As Range
    Dim cel As Range
    Dim dataAtual As Date
    
    ' Define a planilha ativa
    Set ws = ActiveSheet
    
    ' Verifica se a tabela "kitControle" existe na planilha
    On Error Resume Next
    Set tbl = ws.ListObjects("kitControle")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "A tabela 'kitControle' não foi encontrada na planilha.", vbInformation
        Exit Sub
    End If
    
    ' Desprotege a planilha para permitir alterações
    tbl.Parent.Unprotect Password:="Hus07468"
    
    ' Verifica se há um intervalo filtrado na tabela
    On Error Resume Next
    Set filtro = tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If filtro Is Nothing Then
        MsgBox "Não há intervalo filtrado na tabela 'kitControle'.", vbInformation
        Exit Sub
    End If
    
    ' Desabilita a atualização automática e o cálculo
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual