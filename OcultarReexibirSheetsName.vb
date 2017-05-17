    '@Author: Calebe Cavalcante <calebesantos.cavalcante@gmail.com>
    
    '@Parameters:
    '
    
    '@Exemplo: Funções para Ocultar/Reexibir todas as sheets de uma vez
    
    
    Sub Ocultar_SheetsNames()
        'Ocultar Nomes das Sheets
        ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.TabRatio = 0
    End Sub
    
    
    Sub Reexibir_Sheets()
        Dim wsPlan As Worksheet
        
        On Error Resume Next
        
        'Reexibir qualquer sheet oculta
        For Each wsPlan In ThisWorkbook.Sheets
            wsPlan.Visible = xlSheetVisible
        Next
        
        'Reexibir Nomes das Sheets
        ActiveWindow.DisplayWorkbookTabs = True
        ActiveWindow.TabRatio = 0.961
        
        'Reexibir Display de Formulas, Scroll e demais
        With Application
            .CommandBars("Visual Basic").Enabled = True
            .CommandBars("Visual Basic").Visible = True
            .CommandBars("Worksheet Menu Bar").Enabled = True
            .CommandBars("Standard").Enabled = True
            .DisplayAlerts = True
            .DisplayScrollBars = True
            .DisplayFormulaBar = True
            .DisplayFullScreen = False
            .ScreenUpdating = True
        End With
        
    End Sub