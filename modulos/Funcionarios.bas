Attribute VB_Name = "Funcionarios"
Option Explicit
Sub salvarFunc()
    Dim tabela As ListObject
    Dim lin, id As Long
        
    Set tabela = Planilha4.ListObjects(1)
    id = Range("idfuncionario").Value
    lin = tabela.Range.Rows.Count
    
    tabela.Range(lin, 12).Value = id
    tabela.Range(lin, 1).Value = sistema.txtFunNome.Value
    tabela.Range(lin, 2).Value = sistema.txtFunEnd.Value
    tabela.Range(lin, 3).Value = sistema.txtFunCtps.Value
    tabela.Range(lin, 4).Value = sistema.txtFunPis.Value
    tabela.Range(lin, 5).Value = sistema.txtFunSal.Value
    tabela.Range(lin, 6).Value = sistema.txtFunAluguel.Value
    tabela.Range(lin, 7).Value = sistema.txtFunAlim.Value
    tabela.Range(lin, 8).Value = sistema.txtFunValeT.Value
    tabela.Range(lin, 9).Value = sistema.txtFunBoni.Value

    
    tabela.ListRows.Add
    Range("idfuncionario").Value = id + 1
    MsgBox "Cadastrado com sucesso", vbInformation, "Cadastro Funcionário"
    Call limparFunc
    
End Sub
Sub excluirFunc()
    Dim tabela As ListObject
    Dim lin As Long
    Dim result As VbMsgBoxResult
        
    Set tabela = Planilha4.ListObjects(1)
    lin = sistema.lbFunc.ListIndex + 2
    result = MsgBox("Quer mesmo deletar?", vbYesNo, "Cuidado")
    
    
    If result = 6 Then
        sistema.lbFunc.RowSource = ""
        tabela.Range(lin, 1).EntireRow.Delete
        Call listaFunc
    Else
        Exit Sub
    End If

End Sub
Sub editarFunc()
    Dim tabela As ListObject
    Dim lin, id As Long
        
    Set tabela = Planilha4.ListObjects(1)
    lin = sistema.lbFunc.ListIndex + 2
    
        If sistema.tbFunEditar = True Then
            sistema.lbFunc.RowSource = ""
            tabela.Range(lin, 1).Value = sistema.txtFunNome.Value
            tabela.Range(lin, 2).Value = sistema.txtFunEnd.Value
            tabela.Range(lin, 3).Value = sistema.txtFunCtps.Value
            tabela.Range(lin, 4).Value = sistema.txtFunPis.Value
            tabela.Range(lin, 5).Value = sistema.txtFunSal.Value
            tabela.Range(lin, 6).Value = sistema.txtFunAluguel.Value
            tabela.Range(lin, 7).Value = sistema.txtFunAlim.Value
            tabela.Range(lin, 8).Value = sistema.txtFunValeT.Value
            tabela.Range(lin, 9).Value = sistema.txtFunBoni.Value
            Call listaFunc
            MsgBox "Alteração efetuada com sucesso!", vbInformation, "Sucesso"
        Else
            Call listaFunc
        End If
    
End Sub
Sub limparFunc()
    sistema.txtFunNome.Value = ""
    sistema.txtFunEnd.Value = ""
    sistema.txtFunCtps.Value = ""
    sistema.txtFunPis.Value = ""
    sistema.txtFunSal.Value = ""
    sistema.txtFunAluguel.Value = ""
    sistema.txtFunAlim.Value = ""
    sistema.txtFunValeT.Value = ""
    sistema.txtFunBoni.Value = ""
End Sub
Sub listaFunc()
    Dim tabela As ListObject
    
    Set tabela = Planilha4.ListObjects(1)
    sistema.lbFunc.RowSource = Planilha4.ListObjects(1)
    
End Sub
Sub populaFunc()
    Dim tabela As ListObject
    Dim lin, id As Long
        
    Set tabela = Planilha4.ListObjects(1)
    lin = sistema.lbFunc.ListIndex + 2

    If sistema.tbFunEditar = True Then
        sistema.txtFunNome.Value = tabela.Range(lin, 1).Value
        sistema.txtFunEnd.Value = tabela.Range(lin, 2).Value
        sistema.txtFunCtps.Value = tabela.Range(lin, 3).Value
        sistema.txtFunPis.Value = tabela.Range(lin, 4).Value
        sistema.txtFunSal.Value = tabela.Range(lin, 5).Value
        sistema.txtFunAluguel.Value = tabela.Range(lin, 6).Value
        sistema.txtFunAlim.Value = tabela.Range(lin, 7).Value
        sistema.txtFunValeT.Value = tabela.Range(lin, 8).Value
        sistema.txtFunBoni.Value = tabela.Range(lin, 9).Value
    Else
        Call limparFunc
    End If
End Sub


