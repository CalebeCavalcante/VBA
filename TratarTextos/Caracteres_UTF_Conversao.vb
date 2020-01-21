Sub Caracteres_UTF_Conversao()
    Dim vRanges As Variant
    Dim rngSearch As Range
    
    For Each vRanges In Array("E:E", "H:H")
    
        Set rngSearch = Worksheets("efetividade").Range(vRanges)
        Call Replace_Caracteres_UTF_Conversao(rngSearch)
    Next

End Sub

Sub Replace_Caracteres_UTF_Conversao(rngSearch As Range)
       
    arrItem = Array("&iacute;", "&ccedil;", "&atilde;", "&oacute;", "&otilde;", "&Aacute;", "&aacute;", "&Acirc;", "&acirc;", "&Agrave;", "&agrave;", "&Aring;", "&aring;", "&Atilde;", "&atilde;", "&Auml;", "&auml;", "&AElig;", "&aelig;", "&Eacute;", "&eacute;", "&Ecirc;", "&ecirc;", "&Egrave;", "&egrave;", "&Euml;", "&euml;", "&ETH;", "&eth;", "&Iacute;", "&iacute;", "&Icirc;", "&icirc;", "&Igrave;", "&igrave;", "&Iuml;", "&iuml;", "&Oacute;", "&oacute;", "&Ocirc;", "&ocirc;", "&Ograve;", "&ograve;", "&Oslash;", "&oslash;", "&Otilde;", "&otilde;", "&Ouml;", "&ouml;", "&Uacute;", "&uacute;", "&Ucirc;", "&ucirc;", "&Ugrave;", "&ugrave;", "&Uuml;", "&uuml;", "&Ccedil;", "&ccedil;", "&Ntilde;", "&ntilde;", "&lt;", "&gt;", "&amp;", "&quot;", "&reg;", "&copy;", "&Yacute;", "&yacute;", "&THORN;", "&thorn;", "&szlig;")
    arrReplace = Array("í", "ç", "ã", "ó", "õ", "Á", "á", "Â", "â", "À", "à", "Å", "å", "Ã", "ã", "Ä", "ä", "Æ", "æ", "É", "é", "Ê", "ê", "È", "è", "Ë", "ë", "Ð", "ð", "Í", "í", "Î", "î", "Ì", "ì", "Ï", "ï", "Ó", "ó", "Ô", "ô", "Ò", "ò", "Ø", "ø", "Õ", "õ", "Ö", "ö", "Ú", "ú", "Û", "û", "Ù", "ù", "Ü", "ü", "Ç", "ç", "Ñ", "ñ", "<", ">", "&", """", "®", "©", "Ý", "ý", "Þ", "þ", "ß")
        
    For i = 0 To UBound(arrItem)
    
        rngSearch.Replace What:=arrItem(i), Replacement:=arrReplace(i), LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        
    Next
    
End Sub
