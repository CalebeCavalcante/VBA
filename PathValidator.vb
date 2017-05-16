    '@Author: Calebe Cavalcante <calebesantos.cavalcante@gmail.com>
    
    '@Parameters: 
    '#  sPath: Esperado nome de arquivo com a extensão ou diretório 
    '#  IsFolderOnly: se true indica que o valor passado é de uma pasta (Folder) e não de um arquivo com extensão
    
    '@Return: (Boolean) Retorna True se o Arquivo ou Pasta foi encontrado com sucesso
    
    '@Exemplo: PathValidator('c:\users\Desktop\Arquivo.extensao') Return: true
    
    Function PathValidator(sPath As String, Optional IsFolderOnly As Boolean) As Boolean
    Dim sCheck As String
    
        On Error GoTo Fail
        
        ' Se a variável IsFolderOnly indica que está sendo passado em sPath apenas o diretório sem o nome do arquivo
        If IsFolderOnly Then
            sCheck = Dir(sPath, vbDirectory)
        Else
            sCheck = Dir(sPath, vbNormal)
        End If
        
        ' Se o valor de retorno da função Dir for empty então retornar como False
        If sCheck = "" Or IsEmpty(sCheck) Then
            PathValidator = False
        Else
            PathValidator = True
        End If
        
        Exit Function
        
Fail:
        ' Qualquer erro na execução do código retorna como false
        PathValidator = False
    End Function