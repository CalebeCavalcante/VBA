    ' Autor: Calebe Cavalcante <calebesantos.cavalcante@gmail.com>
	
	' # Objetivo
    '   Percorrer todas as sheets do arquivo e aplicar Refresh em qualquer dinâmica
        	
	Sub Refresh_All_PivotTables()

		Dim pvtTable As PivotTable, pvtFields As PivotField
		Dim wsPlan  As Worksheet
		Dim sSheetName As String, sPivotName As String
		
		On Error GoTo Fail
		
		' Percorrer todas as sheets do arquivo
		For Each wsPlan In ThisWorkbook.Sheets
		
			' Para Cada Sheet percorrer todas as pivot tables da sheet
			For Each pvtTable In wsPlan.PivotTables
				
				' Carregar nomes caso ocorra algum erro
				sSheetName = wsPlan.Name
				sPivotName = pvtTable.Name
				
				'Atualizar Link
				pvtTable.PivotCache.Refresh
				
			Next
		Next
		
		MsgBox "Tabelas Dinâmicas atualziadas com sucesso!", vbInformation
		
		Exit Sub
		
	Fail:

		MsgBox "Falha ao atualizar a Dinâmica: " & sPivotName & " na Sheet " & sSheetName, vbCritical, "Refresh_PIvot_tables"
		
	End Sub