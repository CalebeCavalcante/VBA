REM Criar botões para ocultar e reexibir ranges
REM A ideia é atribuir a mesma função em todos os botões em vez de criar uma função para cada um
REM Sendo assim os Gráficos terão nome com prefixo "GRAF_" no caso abaixo:
REM os Objeto de Click com nome do Gráfico. Ex: "Barras". Que mostra o gráfico "GRAF_Barras" com mais alguma opção, se necessário.

Sub Click_Name_Shape()

    Dim vShape As String, vGrafico As String, OPCAO As String, vFigura As String
    
    Application.ScreenUpdating = False
    
    OPCAO = "_" & Range("OPCAO").Value
    
    vFigura = ActiveSheet.Shapes(Application.Caller).Name
    vGrafico = "GRAF_" & vFigura & OPCAO

    For Each Graf In Sheets("Visão_Gerencial").ChartObjects

        If vGrafico = Graf.Name Then

          Graf.Visible = True

        Else:

          Graf.Visible = False

        End If

    Next

    Application.ScreenUpdating = True
    
End Sub
