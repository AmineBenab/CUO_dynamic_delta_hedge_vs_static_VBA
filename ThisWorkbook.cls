Private Sub Workbook_Open()

    MsgBox "Calculations turned to manual"
    
    Application.Calculation = xlCalculationManual

End Sub
