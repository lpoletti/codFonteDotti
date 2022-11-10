Attribute VB_Name = "funcoes"
Function aplication(var As String)
    If var = "on" Then
        Application.Calculation = xlCalculationManual
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End If
End Function

Function calc_HorKm()
On Error Resume Next
    Dim x As Double
    x = 0
    If Len(sistema.txtMovMHorIn.Value) > 0 And Len(sistema.txtMovMHorFim.Value) Then x = x + sistema.txtMovMHorFim.Value - sistema.txtMovMHorIn.Value
    If Len(sistema.txtMovMKmIn.Value) > 0 And Len(sistema.txtMovMKmFim.Value) Then x = x + sistema.txtMovMKmFim.Value - sistema.txtMovMKmIn.Value
    sistema.txtMovMHorKm.Value = x
End Function
Function calc_KmTot()
On Error Resume Next
    Dim x As Double
    x = 0
    If Len(sistema.txtMovMKmIn.Value) > 0 And Len(sistema.txtMovMKmFim.Value) > 0 Then x = x + sistema.txtMovMKmFim.Value - sistema.txtMovMKmIn.Value
    sistema.txtMovMKmTotal.Value = x
End Function
Function calc_VlrTot()
On Error Resume Next
    Dim x As Double
    x = 0
    If Len(sistema.txtMovMValUnit.Value) > 0 And Len(sistema.txtMovMHorKm.Value) > 0 Then x = x + sistema.txtMovMValUnit.Value * sistema.txtMovMHorKm.Value
    sistema.txtMovMValTotal.Value = x
End Function



