Attribute VB_Name = "movMensal"
Option Explicit
Sub salvarMovM()
    Dim tabela, frota As ListObject
    Dim lin, id, result, i As Long
        
    Set tabela = Planilha8.ListObjects(1)
    Set frota = Planilha2.ListObjects(1)
    id = Range("idMovMes").Value
    lin = tabela.Range.Rows.Count
    i = 1
    sistema.cbMovMEquip.BoundColumn = 1
    
    If sistema.cbMovMEquip.ListIndex < 0 Then
        MsgBox "Selecione uma frota"
    Else
        result = CLng(sistema.txtMovMHorIn.Value) - CLng(frota.Range(sistema.cbMovMEquip.ListIndex + 2, 10).Value)
    End If
    If sistema.cbMovMCliente.Value = "" Or sistema.cbMovMEquip.Value = "" Or sistema.cbMovMFunc.Value = "" Then
        MsgBox "Os campos Cliente/Equipamento/Funcionários são obrigatórios", vbCritical, "Aviso"
    Else
        'MsgBox "Horimetro divergente " & result & " Horas" & Chr(13) & "Hora atual: " & frota.Range(sistema.cbMovMEquip.ListIndex + 2, 10).Value & Chr(13) & "Hora inserida: " & sistema.txtMovMHorIn.Value, vbCritical, "Aviso"
        MsgBox "Horimetro divergente", vbCritical
        For i = 1 To frota.Range.Rows.Count
            If CLng(sistema.cbMovMEquip.Value) = frota.Range(i, 8).Value Then
                frota.Range(i, 10).Value = sistema.txtMovMHorFim.Value
            End If
        Next
        
        tabela.Range(lin, 32).Value = sistema.cbMovMEquip.Value
        tabela.Range(lin, 31).Value = id
        tabela.Range(lin, 9).Value = sistema.txtMovMNOs1.Value
        tabela.Range(lin, 10).Value = CDate(sistema.txtMovMData.Value)
        tabela.Range(lin, 11).Value = sistema.cbMovMCliente.Value
        sistema.cbMovMEquip.BoundColumn = 2
        tabela.Range(lin, 12).Value = sistema.cbMovMEquip.Value
        tabela.Range(lin, 13).Value = sistema.txtMovMObra.Value
        tabela.Range(lin, 14).Value = sistema.txtMovMHorIn.Value
        tabela.Range(lin, 15).Value = sistema.txtMovMHorFim.Value
        tabela.Range(lin, 16).Value = sistema.txtMovMKmIn.Value
        tabela.Range(lin, 17).Value = sistema.txtMovMKmFim.Value
        tabela.Range(lin, 18).Value = sistema.txtMovMNumOs2.Value
        tabela.Range(lin, 19).Value = sistema.txtMovMHorKm.Value
        tabela.Range(lin, 20).Value = sistema.txtMovMKmTotal.Value
        tabela.Range(lin, 21).Value = sistema.txtMovMValUnit.Value
        tabela.Range(lin, 22).Value = sistema.txtMovMValTotal.Value
        tabela.Range(lin, 23).Value = sistema.cbMovMSit.Value
        If IsDate(sistema.txtMovMDataPgto.Value) Then
            tabela.Range(lin, 24).Value = CDate(sistema.txtMovMDataPgto.Value)
        Else
            tabela.Range(lin, 24).Value = sistema.txtMovMDataPgto.Value
        End If
        tabela.Range(lin, 25).Value = sistema.cbMovMFunc.Value
        If IsDate(sistema.txtMovMDataPgto.Value) Then
            tabela.Range(lin, 26).Value = CDate(sistema.txtMovMDataAdto.Value)
        Else
            tabela.Range(lin, 26).Value = sistema.txtMovMDataAdto.Value
        End If
        tabela.Range(lin, 27).Value = sistema.txtMovMValAdto.Value
        tabela.Range(lin, 28).Value = sistema.txtMovMDesc.Value
        tabela.Range(lin, 29).Value = sistema.txtMovMValNota.Value
        tabela.Range(lin, 30).Value = sistema.txtMovMObs.Value
    
        
        tabela.ListRows.Add
        Range("idMovMes").Value = id + 1
        MsgBox "Cadastrado com sucesso", vbInformation, "Cadastro Movimentação"
        Call limparMovM
    End If
End Sub
        
        
Sub excluirMovM()
    Dim tabela As ListObject
    Dim lin As Long
    Dim result As VbMsgBoxResult
        
    Set tabela = Planilha8.ListObjects(1)
    'For lin = 2 To tabela.Range.Rows.Count - 1
    'lin = sistema.lbFunc.ListIndex + 2
    result = MsgBox("Quer mesmo deletar?", vbYesNo, "Cuidado")
    
    
    If result = 6 Then
        For j = 2 To tabela.Range.Rows.Count
            'MsgBox tabela.Range(j, 31).Value
            If CLng(sistema.lbMovMId.Caption) = CLng(tabela.Range(j, 31).Value) Then
                sistema.lbMovMMovs.Clear
                tabela.Range(j, 1).EntireRow.Delete
                Call listaMovM
            End If
        Next
    Else
        Exit Sub
    End If

End Sub
Sub editarMovM()
    Dim tabela, frota As ListObject
    Dim lin, id, result, i, j As Long
        
    Set tabela = Planilha8.ListObjects(1)
    Set frota = Planilha2.ListObjects(1)
    id = Range("idMovMes").Value
    lin = tabela.Range.Rows.Count
    i = 1
    j = 1
    sistema.cbMovMEquip.BoundColumn = 1
    
    If sistema.cbMovMEquip.ListIndex < 0 Then
        MsgBox "Selecione uma frota"
    Else
        result = CLng(sistema.txtMovMHorIn.Value) - CLng(frota.Range(sistema.cbMovMEquip.ListIndex + 2, 10).Value)
    End If
    If sistema.cbMovMCliente.Value = "" Or sistema.cbMovMEquip.Value = "" Or sistema.cbMovMFunc.Value = "" Then
        MsgBox "Os campos Cliente/Equipamento/Funcionários são obrigatórios", vbCritical, "Aviso"
    Else
        MsgBox "Horimetro divergente " & result & " Horas" & Chr(13) & "Hora atual: " & frota.Range(sistema.cbMovMEquip.ListIndex + 2, 10).Value & Chr(13) & "Hora inserida: " & sistema.txtMovMHorIn.Value, vbCritical, "Aviso"
        
        For i = 1 To frota.Range.Rows.Count
            If CLng(sistema.cbMovMEquip.Value) = frota.Range(i, 8).Value Then
                frota.Range(i, 10).Value = sistema.txtMovMHorFim.Value
            End If
        Next
        For j = 2 To tabela.Range.Rows.Count
            'MsgBox tabela.Range(j, 31).Value
            If CLng(sistema.lbMovMId.Caption) = CLng(tabela.Range(j, 31).Value) Then
                'MsgBox tabela.Range(j, 31).Value
                tabela.Range(j, 32).Value = sistema.cbMovMEquip.Value
                tabela.Range(j, 9).Value = sistema.txtMovMNOs1.Value
                tabela.Range(j, 10).Value = CDate(sistema.txtMovMData.Value)
                tabela.Range(j, 11).Value = sistema.cbMovMCliente.Value
                sistema.cbMovMEquip.BoundColumn = 2
                tabela.Range(j, 12).Value = sistema.cbMovMEquip.Value
                tabela.Range(j, 13).Value = sistema.txtMovMObra.Value
                tabela.Range(j, 14).Value = sistema.txtMovMHorIn.Value
                tabela.Range(j, 15).Value = sistema.txtMovMHorFim.Value
                tabela.Range(j, 16).Value = sistema.txtMovMKmIn.Value
                tabela.Range(j, 17).Value = sistema.txtMovMKmFim.Value
                tabela.Range(j, 18).Value = sistema.txtMovMNumOs2.Value
                tabela.Range(j, 19).Value = sistema.txtMovMHorKm.Value
                tabela.Range(j, 20).Value = sistema.txtMovMKmTotal.Value
                tabela.Range(j, 21).Value = sistema.txtMovMValUnit.Value
                tabela.Range(j, 22).Value = sistema.txtMovMValTotal.Value
                tabela.Range(j, 23).Value = sistema.cbMovMSit.Value
                If IsDate(sistema.txtMovMDataPgto.Value) Then
                    tabela.Range(j, 24).Value = CDate(sistema.txtMovMDataPgto.Value)
                Else
                    tabela.Range(j, 24).Value = sistema.txtMovMDataPgto.Value
                End If
                tabela.Range(j, 25).Value = sistema.cbMovMFunc.Value
                If IsDate(sistema.txtMovMDataPgto.Value) Then
                    tabela.Range(j, 26).Value = CDate(sistema.txtMovMDataAdto.Value)
                Else
                    tabela.Range(j, 26).Value = sistema.txtMovMDataAdto.Value
                End If
                tabela.Range(j, 26).Value = CDate(sistema.txtMovMDataAdto.Value)
                tabela.Range(j, 27).Value = sistema.txtMovMValAdto.Value
                tabela.Range(j, 28).Value = sistema.txtMovMDesc.Value
                tabela.Range(j, 29).Value = sistema.txtMovMValNota.Value
                tabela.Range(j, 30).Value = sistema.txtMovMObs.Value
        
            
                MsgBox "Editado com sucesso", vbInformation, "Edição Movimentação"
                Call limparMovM
            End If
        Next
    End If
End Sub
Sub limparMovM()
        sistema.txtMovMNOs1.Value = ""
        sistema.txtMovMData.Value = ""
        sistema.cbMovMCliente.Value = ""
        sistema.cbMovMEquip.Value = ""
        sistema.txtMovMObra.Value = ""
        sistema.txtMovMHorIn.Value = ""
        sistema.txtMovMHorFim.Value = ""
        sistema.txtMovMKmIn.Value = ""
        sistema.txtMovMKmFim.Value = ""
        sistema.txtMovMNumOs2.Value = ""
        sistema.txtMovMHorKm.Value = ""
        sistema.txtMovMKmTotal.Value = ""
        sistema.txtMovMValUnit.Value = ""
        sistema.txtMovMValTotal.Value = ""
        sistema.cbMovMSit.Value = ""
        sistema.txtMovMDataPgto.Value = ""
        sistema.cbMovMFunc.Value = ""
        sistema.txtMovMDataAdto.Value = ""
        sistema.txtMovMValAdto.Value = ""
        sistema.txtMovMDesc.Value = ""
        sistema.txtMovMValNota.Value = ""
        sistema.txtMovMObs.Value = ""
End Sub
Sub listaMovM()
    Dim tabela As ListObject
    Dim i, y, z As Long
    
    Set tabela = Planilha8.ListObjects(1)
    y = 1
    z = 0
    'MsgBox WorksheetFunction.CountIf(tabela.DataBodyRange.Columns(3), CLng(sistema.cbMovMFiltro.Value))
    For i = 0 To tabela.DataBodyRange.Rows.Count - 1
    'If tabela.DataBodyRange.Cells(y, 3) <> vbString Then
            If CLng(tabela.DataBodyRange.Cells(y, 3)) = CLng(sistema.cbMovMFiltro.Value) Then
            
                With sistema.lbMovMMovs
                   .AddItem
                   .Column(0, z) = tabela.DataBodyRange.Cells(y, 9)
                   .Column(1, z) = tabela.DataBodyRange.Cells(y, 10)
                   .Column(2, z) = tabela.DataBodyRange.Cells(y, 11)
                   .Column(3, z) = tabela.DataBodyRange.Cells(y, 12)
                   .Column(4, z) = tabela.DataBodyRange.Cells(y, 13)
                   .Column(5, z) = FormatCurrency(tabela.DataBodyRange.Cells(y, 22))
                   .Column(6, z) = tabela.DataBodyRange.Cells(y, 23)
                   .Column(7, z) = tabela.DataBodyRange.Cells(y, 31)
                End With
              z = z + 1
            End If
         y = y + 1
    'End If
    Next
    
End Sub
Sub populaCbs()
    Dim tabela As ListObject
    Dim frotas As ListObject
    Dim clientesM As ListObject
    Dim funcs As ListObject
    Dim i, y As Long
        
    Set tabela = Planilha8.ListObjects(1)
    Set frotas = Planilha2.ListObjects(1)
    Set clientesM = Planilha3.ListObjects(1)
    Set funcs = Planilha4.ListObjects(1)
    y = 1
    For i = 0 To frotas.DataBodyRange.Rows.Count - 1
     With sistema.cbMovMEquip
        .AddItem frotas.DataBodyRange.Cells(y, 8)
        .Column(1, i) = frotas.DataBodyRange.Cells(y, 3)
     End With
     y = y + 1
    Next
    y = 1
    For i = 0 To clientesM.DataBodyRange.Rows.Count - 1
     With sistema.cbMovMCliente
        .AddItem clientesM.DataBodyRange.Cells(y, 13)
        .Column(1, i) = clientesM.DataBodyRange.Cells(y, 1)
     End With
     y = y + 1
    Next
    y = 1
    For i = 0 To funcs.DataBodyRange.Rows.Count - 1
     With sistema.cbMovMFunc
        .AddItem funcs.DataBodyRange.Cells(y, 12)
        .Column(1, i) = funcs.DataBodyRange.Cells(y, 1)
     End With
     y = y + 1
    Next
    
    sistema.cbMovMSit.RowSource = "'MOV MENSAL'!$X$6:$X$11"
    
    
End Sub
Sub populaMovM()
    Dim tabela As ListObject
    Dim lin, id As Long
    On Error Resume Next
    Set tabela = Planilha8.ListObjects(1)
    
    For lin = 2 To tabela.Range.Rows.Count - 1
        If sistema.tbMovMEditar = True And CLng(tabela.Range(lin, 31).Value) = CLng(sistema.lbMovMMovs.Value) Then
            sistema.txtMovMNOs1.Value = tabela.Range(lin, 9).Value
            sistema.txtMovMData.Value = tabela.Range(lin, 10).Value
            sistema.cbMovMCliente.Value = tabela.Range(lin, 11).Value
            sistema.cbMovMEquip.Value = tabela.Range(lin, 12).Value
            sistema.txtMovMObra.Value = tabela.Range(lin, 13).Value
            sistema.txtMovMHorIn.Value = tabela.Range(lin, 14).Value
            sistema.txtMovMHorFim.Value = tabela.Range(lin, 15).Value
            sistema.txtMovMKmIn.Value = tabela.Range(lin, 16).Value
            sistema.txtMovMKmFim.Value = tabela.Range(lin, 17).Value
            sistema.txtMovMNumOs2.Value = tabela.Range(lin, 18).Value
            sistema.txtMovMHorKm.Value = tabela.Range(lin, 19).Value
            sistema.txtMovMKmTotal.Value = tabela.Range(lin, 20).Value
            sistema.txtMovMValUnit.Value = tabela.Range(lin, 21).Value
            sistema.txtMovMValTotal.Value = tabela.Range(lin, 22).Value
            sistema.cbMovMSit.Value = tabela.Range(lin, 23).Value
            sistema.txtMovMDataPgto.Value = tabela.Range(lin, 24).Value
            sistema.cbMovMFunc.Value = tabela.Range(lin, 25).Value
            sistema.txtMovMDataAdto.Value = tabela.Range(lin, 26).Value
            sistema.txtMovMValAdto.Value = tabela.Range(lin, 27).Value
            sistema.txtMovMDesc.Value = tabela.Range(lin, 28).Value
            sistema.txtMovMValNota.Value = tabela.Range(lin, 29).Value
            sistema.txtMovMObs.Value = tabela.Range(lin, 30).Value
            sistema.lbMovMId.Caption = tabela.Range(lin, 31).Value
        Else
            'Call limparMovM
        End If
    Next
End Sub
Sub populaMovMFiltro()
    Dim i As Integer
    For i = 2017 To Year(Date)
        sistema.cbMovMFiltro.AddItem i
    Next
End Sub



