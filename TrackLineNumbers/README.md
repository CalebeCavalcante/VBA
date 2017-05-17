# Descrição

Adicionar o número de linha dentro do código VBA. Mesma funcionalidade oferecida pelo [Total Visual CodeTools](http://www.fmsinc.com/MicrosoftAccess/VBACodingTools.html)

Numerando as linhas do seu código VBA, permite utilizar o <b>Err.Number</b> para apontar a linha exata onde o ocorreu o erro do código.

Maiores explicações sobre a funcionalidade Err.Number em [Track Line Numbers to Pinpoint the Location of a Crash](https://msdn.microsoft.com/en-us/library/ee358847(v=office.12).aspx)

# Uso

1. Abrir copiar todo o código VBA e colar no arquivo Word;
2. Alt + F8 para acessar a lista de macros do arquivo Word;
3. Existem duas macros: <b>Enumerar_Linhas_Arquivo</b> & <b>Retirar_Numeracao_Linhas_Arquivo</b>
  > <b>Enumerar_Linhas_Arquivo</b>: Irá colocar um número sequencial em todos os códigos que executam alguma lógica de programação. Excluiem-se desse conceito comentários, declarações e semelhantes
  > <b>Retirar_Numeracao_Linhas_Arquivo</b>:  Irá retirar todos os números das linhas. Com isso você pode reorganizar o arquivo e numerar novamente.
4. Agora apenas voltar o código para seu VBE.
