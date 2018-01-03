
' Fonte: http://analystcave.com/excel-regex-tutorial/

' ############## Like Operator ##############

' Fonte: http://analystcave.com/vba-like-operator/
' Regras para o Like Operator 
' * matches any number of characters
' ? matches any 1 character
' [ ] matches any 1 character specified between the brackets
' - matches any range of characters e.g. [a-z] matches any non-capital 1 letter of the alphabet
' # matches any digit character

If "Animal" Like "[A-Z]*" then 
   Debug.Print "Match: String starting with Capital letter!"
End If

If "3Q/2017" Like "#[Q|T|S]/#*" Then
   Debug.Print "Data Agrupada por Q, S ou T"
End If

' ############## Regex ##############

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
