Function remove_acentos(sString As String) As String
     
    Dim sAcentos As String, sSemAcentos As String, sTemp As String
    Dim i As Long
  
    sAcentos = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
    sSemAcentos = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
      
    sTemp = sString
    
    For i = 1 To Len(sAcentos)
        sTemp = Replace(sTemp, Mid(sAcentos, i, 1), Mid(sSemAcentos, i, 1))
    Next i
    
    remove_acentos = sTemp
      
End Function
