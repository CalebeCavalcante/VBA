
Public Sub Exemplo_Get_Dados()
    Dim MySheet As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long

    Set MySheet = ThisWorkbook.Sheets("MySheetName")
    
    LastRow = getLastRow(MySheet)
    LastColumn = getLastColumn(MySheet)
    
    If LastRow < 1 Or LastColumn < 1 Then ErrorGetInfo
    
    '  Ok para utilizar as variáveis
    
    
    Exit Sub
    
ErrorGetInfo:
    MsgBox "Falha ao buscar dados da linha e coluna", vbInformation
    
End Sub
Public Function getUltimaLinha(ByRef wsPlan As Worksheet) As Long
    
    ' # Objetivo
    '   Verificar a maior qtde de linhas em uma planilha
    ' # Motivo
    '   Uma ou outra coluna pode ter dados em branco, sendo assim percorrer todas as colunas
    '   do titulo reduz o erro em verificar a última linha 
    ' # Observações
    '   Por que não usar UsedRange.Rows.Count ?
    '   É possível implementar sim um algoritimo com essa função, porém o UsedRange começa a partir da célula onde
    '   os dados começam (e.g. E50:E500), sendo assim, neste exemplo o UsedRange.Rows.Count seria 450 (i.e. Row(50) - (Row500) )
    '   porém a última linha ainda seria 500.
	
	Dim lColumns As Long, lRows As Long, lUltCol As Long, lUltRow As Long, lColunaTitulos As Long, lRowHead as Long
    
    On Error GoTo Fail
    
    '  lRowHead = Primeira Linha com dados das Informações
    '  Se lRowHead não for passado, será utilizado a primeira linha do Range que está sendo usado na sheet
    
    lRowHead = 1
    
    lUltCol = getUltimaColuna(wsPlan, lRowHead)
    lUltRow = 0
    
    For lColumns = 1 To lUltCol
     
     '  Obs.: cells: wsPlan.Rows.Count pois o arquivo pode ir até a linha 1048756(xlsx) ou 65536(xls).
     lRows = wsPlan.Cells(wsPlan.Rows.Count, lColumns).End(xlUp).Row
     
     '  Se a atual for maior, alterar para a atual
     If lUltRow < lRows Then lUltRow = lRows
     
    Next
    
    getUltimaLinha = lUltRow
    
    Exit Function
Fail:
    getUltimaLinha = 0
End Function
Public Function getUltimaColuna(ByRef wsPlan As Worksheet, Optional lRowHead As Long)
	
	' # Objetivo
    '   Verificar a ultima coluna com dados em uma planilha
    ' # Motivo
    '   Ter uma função que retornar a coluna de acordo com a linha de título agiliza o
    '   o processo de programação e reduz erros.
	' # Observações
    '   Por que não usar UsedRange.Columns.Count ?
    '   É possível implementar sim um algoritimo com essa função, porém o UsedRange começa a partir da célula onde
    '   os dados começam (e.g. E50:E500), sendo assim, neste exemplo o UsedRange.Columns.Count seria 1 (i.e. Column(E) to Column(E) )
    '   porém a última coluna ainda seria 5 (i.e. A to E).
	
    On Error GoTo Fail
   
    '  lRowHead = Primeira Linha com dados das Informações
    
    '  Se lRowHead não for passado, será utilizado a primeira linha do Range que está sendo usado na sheet (i.e. UsedRange.Rows(1).Row)
    If lRowHead < 1 Then lRowHead = wsPlan.UsedRange.Rows(1).Row
    
    '  Obs.: cells:  1 & wsPlan.Columns.Count pois o arquivo pode ir até a coluna XFD(xlsx) ou IV(xls).
    getUltimaColuna = wsPlan.Cells(lRowHead, wsPlan.Columns.Count).End(xlToLeft).Column
    
	Exit Function
Fail:
    getUltimaColuna = 0

End Function
