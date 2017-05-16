    '@Author: Calebe Cavalcante <calebesantos.cavalcante@gmail.com>
    
    '@Parameters: 
    '#  sFileName: Esperado nome de arquivo com a extensão no final
    
    '@Return: (String) Extensão do Arquivo passado
    
    '@Exemplo: GetExtensionByRegExp('Arquivo.extensao') Return: string(extensao)
    
    Function GetExtensionByRegExp(sFileName As String) As String
    Dim regex As Object, str As String
        
        On Error GoTo Fail
        
        ' Criando novo objeto RegExp
        Set regex = CreateObject("VBScript.RegExp")
        
        ' Padrão Para pegar os últimos caracteres do final do arquivo
        str = "\.\w{2,}$"
        
        'Carregando padrão no objeto regex
        With regex
          .Pattern = str
          .Global = True
        End With
        
        'Verificando se o padrão é encontrado
        If regex.Test(sFileName) Then
            ' Retornando a extensão fazendo:
            ' 1 replace regex.Replace, tirando a extensão, deixando apenas o nome do arquivo
            ' 2 replace interno, tirando o nome, deixando assim só a extensão
            ' 3 replace externo, tirando o "." da extensão
            GetExtensionByRegExp = Replace(Replace(sFileName, regex.Replace(sFileName, ""), ""), ".", "")
        Else
            GetExtensionByRegExp = Empty
        End If
        
        ' Finalizando variável
        Set regex = Nothing
        
        Exit Function
Fail:
        ' Qualquer erro na execução, retornar vazio
        GetExtensionByRegExp = Empty
        
    End Function