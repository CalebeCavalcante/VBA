  
  ' Autor: Calebe Cavalcante. <calebesantos.cavalcante@gmail.com>
	
	' # Objetivo
    '   Permitir que o usuário selecione arquivos em uma pasta no Windows para usar nas demais rotinas
    
	' # Motivo
    '   Este código permite ao usuário interagir com a rotina, podendo escolher o arquivo que precisa ser trabalhado no código.
	'   Entretanto pode também aplicar a qualquer cenário onde uma lista de arquivos precisa ser percorrida, cabendo ao desenvolvedor aplicar a lógica necessária para tratar os dados.
       
    ' # Observações
    ' 	Para Funcionar é necessário adicionar a DLL Scription Runtim.
	'	Para isto ir em Tools >> Refereces >> Microsoft Scripting Runtime ( marcar check box )

    Sub Exemplo_Uso_MinhaRotina()
    
        Dim UserFileSelect
        Dim UserFileList
        
        ' Uso para selecionar 1 arquivo
        UserFileSelect = Escolher_Arquivo()
        
        ' Uso para selecionar vários arquivos
        UserFileList = Escolher_Arquivo(True)
        
        ' Percorrer a lista de arquivos selecionados
        For Each Item In UserFileList
            
            UserFileSelect = Item
        
        Next
    
    End Sub

    Private Function Escolher_Arquivo(Optional bMultiSelect As Boolean)
        Dim Fs As FileSystemObject
        Dim oFile As File
        Dim arrFiles As Variant
        Dim arrRetorno() As String
        Dim lCount As Long
        
        On Error GoTo Erro:
        
        ' Iniciar Variável (-1) para começar o array no ponteiro zero (0)
        lCount = -1
        
        ' Abrir Opção para selecionar os Arquivos
        arrFiles = Application.GetOpenFilename("Todos os Arquivos (*.*), *.*", MultiSelect:=bMultiSelect)
        
        ' Iniciar novo objeto FileSystemObject
        Set Fs = New FileSystemObject
        
				
        ' Verificar se o retorno é um array.
        ' Obs.: Somente será array se a variável bMultiSelect = true
        If IsArray(arrFiles) = False Then
                        
			' ######## Exemplos - Desenvolvedor pode retornar ########
            
			' o nome do arquivo
			sNomeArquivo = oFile.Name

			' a pasta onde o arquivo está
			sPastaArquivo = oFile.ParentFolder

			' o nome completo ( Pasta + Nome Arquivo )
			sCaminhoCompleto = oFile.Path
			
			' ################
			
			' Se falso, retornar o caminho completo do arquivo
			Set oFile = Fs.GetFile(arrFiles)
			
            Escolher_Arquivo = oFile.Path
			
            Exit Function
            
        Else
            ' Se é um array, então percorrer os arquivos(Item) selecionados
            For Each Item In arrFiles
                
                ' Para Cada Item , carregar o arquivo
                Set oFile = Fs.GetFile(Item)
                
                ' Add o contador
                lCount = lCount + 1
                
                ' Redefinir o array de retorno (Redim) sem perder os dados já inseridos no array (Preserve)
                ReDim Preserve arrRetorno(lCount)
                
                ' Por Default está o caminho completo. Caso precise utilize aqui outra variável (Descritos no Top)
                arrRetorno(lCount) = oFile.Path
                
            Next
            
            ' Após percorrer os Itens, retornar com o array
            Escolher_Arquivo = arrRetorno
            
        End If
        
        Exit Function
    
Erro:
        ' Se bMultiSelect = true o retorno esperado é um array
        If bMultiSelect Then
            Escolher_Arquivo = Array()
        Else
            Escolher_Arquivo = ""
        End If
    
    End Function
