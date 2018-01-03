
' Fonte: http://analystcave.com/excel-regex-tutorial/

' Like Operator

If "Animal" Like "[A-Z]*" then 
   Debug.Print "Match: String starting with Capital letter!"
End If

' TESTE MATCHING

Dim regex As Object, str As String
Set regex = CreateObject("VBScript.RegExp")
 
With regex
  .Pattern = "[0-9]+"
End With
     
str = "Hello 123 World!"
Debug.Print regex.Test(str) 'Result: True
 
str = "Hello World!"
Debug.Print regex.Test(str) 'Result: False  

' REPLACE 
Dim regex As Object, str As String
Set regex = CreateObject("VBScript.RegExp")
 
With regex
  .Pattern = "123-[0-9]+-123"
  .Global = True 'If False, would replace only first
End With
     
str = "321-123-000-123-643-123-888-123"
Debug.Print regex.Replace(str, "<Replace>") 
'Result: 321-<Replace>-643-<Replace>
