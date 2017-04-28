	
	' Autor: Calebe Cavalcante. <calebesantos.cavalcante@gmail.com>
	
	' # Objetivo
    '   Percorrer todos os arquivos em uma pasta no Windows e abrir os arquivos com extensão em Excel
    
	' # Motivo
    '   Este código se encaixa perfeitamente em um cenário onde é preciso consolidar vários arquivos com mesmo layout de colunas. 
	'   Entretanto pode também aplicar a qualquer cenário onde uma lista de arquivos precisa ser percorrida, cabendo ao desenvolvedor aplicar a lógica necessária para tratar os dados.
       
    ' # Observações
    ' 	Para Funcionar é necessário adicionar a DLL Scription Runtim.
	'	Para isto ir em Tools >> Refereces >> Microsoft Scripting Runtime ( marcar check box )
	
	Sub Percorrer_Arquivos_Pasta()

		Dim Fs As FileSystemObject
		Dim oFolder As Folder, oFile As File, sFolderPath As String
		Dim wkBook As Workbook
			
		' Infomar a pasta onde os arquivos estão
		sFolderPath = "C:\Users\CalebeCavalcante\Desktop\Arquivos Excel"

		' Iniciar novo objeto FileSystemObject
		Set Fs = New FileSystemObject
		
		' Carregar Objeto Folder na variável
		Set oFolder = Fs.GetFolder(sFolderPath)
		
		' Para Cara Arquivo dentro da pasta
		For Each oFile In oFolder.Files
			
			' Verificar se a extensão é de Excel (xlsx, xlsb, xlsm e etc.)
			If Left(Fs.GetExtensionName(oFile.Path), 3) = "xls" Then
				
				Set wkBook = Workbooks.Open(oFile.Path)
				
				' Executar demais ações com a variável wsBook
				
				' Fechar Arquivo
				wkBook.Close
				
			End If
		Next
		
		' Finalizar variável
		Set oFile = Nothing
		Set oFolder = Nothing
		Set Fs = Nothing
			
	End Sub