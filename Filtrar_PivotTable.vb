    '@Author: Calebe Cavalcante <calebesantos.cavalcante@gmail.com>
    
    '@Parameters: 
    '#  wsSheet: Esperado variável com objeto WorkSheet setado (e.g. set ws = sheets("Plan1") )
    '#  sPivotTableName: Nome da Pivot na sheet informada
    '#  sPivotFieldName: Nome do campo. (Nome original da Base e não o nome que foi colocado na dinâmica)
    '#  sValueFilter: Valor do campo que será filtrado. Se o valor não for encontrado, será exibido todos os valores do campo
    
    '@Return: avoid 
    
    '@Exemplo: Filtrar_PivotTable(ActiveSheet, "PivotTable 1", "Serviço ativo", "sim")
    
    
    Sub Filtrar_PivotTable(wsSheet As Worksheet, sPivotTableName As String, sPivotFieldName As String, sValueFilter As String)
    Dim pvField As PivotField, index As Integer, indexFind As Integer

        Set pvField = wsSheet.PivotTables(sPivotTableName).PivotFields(sPivotFieldName)
    
        ' Limpar Qualquer seleção anterior
        pvField.ClearAllFilters
        
        index = 0
        
        'Percorrer todos os itens da tabela dinâmica para verificar se o valor existe no campo
        Do
            index = index + 1
            
            'Se o valor corresponde, então selecionar
            If pvField.PivotItems(index).Value = sValueFilter Then
                indexFind = index
            End If
            
        Loop While index < pvField.PivotItems.Count And indexFind = 0
        
        'Se o valor procurado for encontrado
        If indexFind > 0 Then
            'Colocar o index do campo encontrado visivel
            pvField.PivotItems(indexFind).Visible = True
        
            'Percorrer todos os itens da tabela dinâmica e ocultar
            For index = 1 To pvField.PivotItems.Count
                'Se o valor corresponde, então selecionar
                If indexFind <> index Then
                    pvField.PivotItems(index).Visible = False
                End If
            Next
        
        End If
        
    End Sub
