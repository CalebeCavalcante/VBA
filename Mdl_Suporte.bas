Attribute VB_Name = "Mdl_Suporte"
Sub showMessages(options As Boolean)

    Application.ScreenUpdating = options
    Application.DisplayAlerts = options
    
End Sub
Sub doDelay(Optional xTimes As Long)
Const t As Date = "00:00:05"

If xTimes < 1 Then xTimes = 1

Application.Wait Now + (t * xTimes)
DoEvents

End Sub
Sub clearPlan(Plan As Worksheet)
On Error Resume Next

    Plan.Rows("2:" & Rows.Count).ClearContents
    Plan.Rows("2:" & Rows.Count).Delete

End Sub
Sub activePlan(Plan As Worksheet)

On Error Resume Next
Plan.Activate
Plan.Select

End Sub
Function lastRow(Plan As Worksheet) As Long

Dim lColumn As Long, lMax As Long

For lColumn = 1 To 100

    lMax = Plan.Cells(Rows.Count, lColumn).End(xlUp).Row
    
    Select Case lMax

        Case Is > lastRow
            lastRow = lMax
    
    End Select

Next

End Function
Function Split97(sStr As Variant, sdelim As String) As Variant
''Tom Ogilvy
Split97 = Evaluate("{""" & _
                       Application.Substitute(sStr, sdelim, """,""") & """}")
End Function
Function implodeArrToString(arrDados As Variant)
    
    Dim dado As Variant
    
    For Each dado In arrDados
        If Left(dado, 7) = "to_date" Then
            implodeArrToString = implodeArrToString & CStr(dado) & ","
        Else
            implodeArrToString = implodeArrToString & Quote(CStr(dado), ",")
        End If
    Next

End Function
Function Quote(sText As String, Optional AddValue As String)
    Quote = "'" & sText & "'" & AddValue
End Function
Function RemoveLastComa(sText As String)
    RemoveLastComa = Left(Trim(sText), Len(Trim(sText)) - 1)
End Function
Function toNumber(sText As String)
On Error GoTo Fail

    toNumber = CLng(Trim(Replace(sText, " ", "")))
    Exit Function
Fail:
    toNumber = 0
    
End Function
Function isNumberInList(lngToFind As Long, sTextListOfNumbers As String) As Boolean

Rem EXEMPLO:
Rem RANGE("A1") = 1,2,3,45,6
Rem RANGE("A2") = 45
Rem Uso = isNumberInList(CLng(Range("A2").Value), Range("A1").Value)

Dim vList  As Variant, vItem As Variant

Rem Default
isNumberInList = False

On Error Resume Next

vList = Split97(sTextListOfNumbers, ",")

For Each vItem In vList
    If toNumber(CStr(vItem)) = lngToFind Then
        isNumberInList = True
        Exit Function
    End If
Next

End Function
Function addOne(ByRef value)
    value = value + 1
End Function
Function clearToCSV(sText As String)
    
    clearToCSV = Replace(Replace(Replace(sText, ",", ""), Chr(10), ""), Chr(13), "")

End Function
Function thisFilePath(Optional sSubFolder As String) As String
    thisFilePath = ThisWorkbook.Path & Application.PathSeparator & IIf(Len(sSubFolder), sSubFolder & Application.PathSeparator, "")
End Function
Sub gerarCSV(wsPlan As Worksheet)
    Dim sNameFile As String

    sNameFile = thisFilePath("BASES_EXPORT") & wsPlan.Name & ".csv"
    sNameFile = getFileNameAvailable(sNameFile)
        
    activePlan wsPlan
    
    wsPlan.Copy
    ActiveWorkbook.SaveAs sNameFile, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Close
    
End Sub
Function IsFileExists(sPath As String) As Boolean

Dim Fs As FileSystemObject
Rem Necessário habilitar em
Rem Tools >> References >> Microsoft Script Runtime
Rem para usar o FileSystemObject

    On Error GoTo Fail
    
    Set Fs = New FileSystemObject
    IsFileExists = Fs.FileExists(sPath)
      
    Set Fs = Nothing
    Exit Function
    
Fail:
    IsPathValid = False
    Set Fs = Nothing
    
End Function
Function getFileNameAvailable(sFileName As String) As String
    
    Dim Fs As FileSystemObject
    Rem Necessário habilitar em
    Rem Tools >> References >> Microsoft Script Runtime
    Rem para usar o FileSystemObject
    
    Dim sExtension As String, sFullName As String
    Dim lVersion As Long
    
    On Error GoTo Fail
    
    getFileNameAvailable = sFileName
    
    If IsFileExists(sFileName) Then
        
        Set Fs = New FileSystemObject
        sExtension = Fs.GetExtensionName(sFileName)
        sExtension = IIf(Len(sExtension), "." & sExtension, "")
        sFullName = Left(sFileName, Len(sFileName) - Len(sExtension))
        
        lVersion = 1
        
        Rem Enquanto o nome existir, tentar outras versoes
        getFileNameAvailable = sFullName & "_v" & lVersion & sExtension
        
        Debug.Print getFileNameAvailable
        
        Do While IsFileExists(getFileNameAvailable)
            addOne lVersion
            getFileNameAvailable = sFullName & "_v" & lVersion & sExtension
        Loop
        
    End If
    
    GoTo Final
    
    Exit Function
    
Fail:
    getFileNameAvailable = sFileName

Final:
    Set Fs = Nothing
    
End Function
Sub SaveContentHTML(sContent As String, sCodeFile As String)
    
    Rem Necessário habilitar em
    Rem Tools >> References >> Microsoft Script Runtime
    Rem para usar o FileSystemObject
    
    Dim Fs As FileSystemObject
    Dim tf As TextStream
    Dim sNameFile As String
    
    On Error GoTo Fim
    
    Set Fs = New FileSystemObject
    
    sNameFile = thisFilePath("DOCS_EXPORT") & "DOCNUMBER_" & sCodeFile & ".htm"
    sNameFile = getFileNameAvailable(sNameFile)
            
    Set tf = Fs.CreateTextFile(sNameFile, True)
    
    tf.Write sContent
    tf.Close
    
Fim:
  
    Set tf = Nothing
    Set Fs = Nothing
    
   
End Sub
Function upper(sText As String)
    upper = StrConv(sText, vbUpperCase)
End Function
Function lower(sText As String)
    lower = StrConv(sText, vbLowerCase)
End Function
Function proper(sText As String)
    proper = StrConv(sText, vbProperCase)
End Function
Function findColumnOnWs(wsPlan As Worksheet, sColumnName As String) As Long
    Dim rFind As Range
    
    Set rFind = wsPlan.Range("1:1").Find(sColumnName)
    
    If rFind Is Nothing Then
        findColumnOnWs = 0
    Else
        findColumnOnWs = rFind.Column
    End If
    
End Function

