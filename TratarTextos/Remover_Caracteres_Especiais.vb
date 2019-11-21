Function remover_caracteres_especiais(sTexto As String)
    
    Dim sCaracter As String, sTemp As String
    Dim i As Long
  
    sCaracter = "!@#$%¨&*()_+=§"
          
    sTemp = sTexto
    
    For i = 1 To Len(sCaracter)
        sTemp = Replace(sTemp, Mid(sCaracter, i, 1), "")
    Next i
    
    remover_caracteres_especiais = sTemp
    
End Function
