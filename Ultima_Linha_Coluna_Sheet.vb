
Public Sub Exemplo_Get_Dados()

    Dim MySheet As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long

    Set MySheet = ThisWorkbook.Sheets("MySheetName")
    
    LastRow = getLastRow(MySheet)
    LastColumn = getLastColumn(MySheet)
    
    If LastRow < 1 Or LastColumn < 1 Then ErrorGetInfo
    
    Rem Ok para utilizar as variáveis
    
        
    Exit Sub
    
ErrorGetInfo:
    MsgBox "Falha ao buscar dados da linha e coluna"
    
End Sub

Public Function getUltimaLinha(ByRef wsPlan As Worksheet) As Long

    Dim lColumns As Long, lRows As Long, lUltCol As Long, lUltRow As Long
    
    On Error GoTo Fail
    
    Rem Utilizando IV(.xls) em vez de XFD(.xlsx) por conta da versão .xls, e alguns arquivos estão como .xls
    lUltCol = getUltimaColuna(wsPlan)
    
    Rem Pode ter colunas em branco, sendo assim verificar a Coluna com a maior qtde de linhas, abaixo:
    For lColumns = 1 To lUltCol
     
     Rem Obs.: cells: .Rows.Count pq o arquivo pode ir até a linha 1048756(xlsx) ou 65536(xls).
     lRows = wsPlan.Cells(wsPlan.Rows.Count, lColumns).End(xlUp).Row
     
     If lUltRow < lRows Then lUltRow = lRows
     
    Next
    
    getUltimaLinha = lRows
    
    Exit Function
Fail:
    getUltimaLinha = 0
	
End Function

Public Function getUltimaColuna(ByRef wsPlan As Worksheet)

    On Error GoTo Fail
    
    Rem Obs.: cells:  1 & wsPlan.Columns.Count pq o arquivo pode ir até a coluna XFD(xlsx) ou IV(xls).
    getUltimaColuna = wsPlan.Cells(1, wsPlan.Columns.Count).End(xlToLeft).Column
    
Fail:
    getUltimaColuna = 0

End Function