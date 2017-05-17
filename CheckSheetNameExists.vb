    '@Author: Calebe Cavalcante <calebesantos.cavalcante@gmail.com>
    
    '@Parameters: 
    '#  Book: Esperado variável do tipo Workbook onde a sheet(Plan) será procurada. Pode ser o mesmo arquivo da macro ou arquivo externo, setado em uma variável do tipo workbook
    '#  sSheetName: Esperado string com o nome da sheet (Plan) de procura
    '@Return: (Boolean) True se a sheet(Plan) for localizada no arquivo

    Public Function checkSheetExists(ByRef Book As Workbook, sSheetName As String) As Boolean
        Dim wsPlans As Worksheet
        
        On Error GoTo Fail
        
        'Tentando setar a sheet no arquivo passado. Caso ocorra erro irá para false no step Fail
        Set wsPlans = Book.Sheets(sSheetName)
        
        'Se o objeto for nothing indica que não foi encontrado
        If wsPlans Is Nothing Then
            checkSheetExists = False
        Else
            checkSheetExists = True
        End If
        
        Exit Function
        
    Fail:
        ' Se não entrar no if acima, então retornar como false
        checkSheetExists = False
        
    End Function