Sub remove_quebras_linha()
 
    Dim rngSearch As Range
    Dim vItem As Variant
    
    Set rngSearch = Worksheets("Planilha 1").Range("C2:J8562")
    
    For Each vItem In Array(9, 10, 13)
        find_and_replace rngSearch, Chr(vItem), ""
    Next

End Sub

Sub find_and_replace(rngSearch As Range, sFind As String, sReplace As String)

    With rngSearch
        Set c = .Find(sFind, LookIn:=xlValues)
        If Not c Is Nothing Then
            firstAddress = c.Address
            Do
                c.Value = Replace(c.Value, sFind, sReplace)
                Set c = .FindNext(c)
            Loop While Not c Is Nothing
        End If
    End With
    
End Sub
