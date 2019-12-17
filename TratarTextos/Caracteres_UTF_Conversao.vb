Sub Caracteres_UTF_Conversao()
    Dim vRanges As Variant
    Dim rngSearch As Range
    
    For Each vRanges In Array("E:E", "H:H")
    
        Set rngSearch = Worksheets("efetividade").Range(vRanges)
        Call Replace_Caracteres_UTF_Conversao(rngSearch)
    Next

End Sub

Sub Replace_Caracteres_UTF_Conversao(rngSearch As Range)
       
    arrItem = Array("&iacute;", "&ccedil;", "&atilde;", "&oacute;", "&otilde;")
    arrReplace = Array("í", "ç", "ã", "ó", "õ")
    
    For i = 0 To UBound(arrItem)
    
        rngSearch.Replace What:=arrItem(i), Replacement:=arrReplace(i), LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        
    Next
    
End Sub
