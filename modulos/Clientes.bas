Attribute VB_Name = "Clientes"
Option Explicit
Sub salvarCliente()
    Dim tabela As ListObject
    Dim lin, id As Long
        
    Set tabela = Planilha3.ListObjects(1)
    id = Range("idcliente").Value
    lin = tabela.Range.Rows.Count
    
    tabela.Range(lin, 13).Value = id
    tabela.Range(lin, 1).Value = sistema.txtClNome.Value
    tabela.Range(lin, 2).Value = sistema.txtClCnpj.Value
    tabela.Range(lin, 3).Value = sistema.txtClInsEst.Value
    tabela.Range(lin, 4).Value = sistema.txtClInsMun.Value
    tabela.Range(lin, 5).Value = sistema.txtClCei.Value
    tabela.Range(lin, 6).Value = sistema.txtClEnd.Value
    tabela.Range(lin, 7).Value = sistema.txtClMun.Value
    tabela.Range(lin, 8).Value = sistema.cbClUf.Value
    tabela.Range(lin, 11).Value = sistema.txtClObs.Value
    tabela.Range(lin, 10).Value = sistema.txtClTel.Value
    tabela.Range(lin, 9).Value = sistema.txtClEmail.Value
    
    tabela.ListRows.Add
    Range("idcliente").Value = id + 1
    MsgBox "Cadastrado com sucesso", vbInformation, "Cadastro Cliente"
    Call limparCliente
    
End Sub
Sub excluirCliente()
    Dim tabela As ListObject
    Dim lin As Long
    Dim result As VbMsgBoxResult
        
    Set tabela = Planilha3.ListObjects(1)
    lin = sistema.lbClClientes.ListIndex + 2
    result = MsgBox("Quer mesmo deletar?", vbYesNo, "Cuidado")
    
    
    If result = 6 Then
        sistema.lbClClientes.RowSource = ""
        tabela.Range(lin, 1).EntireRow.Delete
        Call listaClientes
    Else
        Exit Sub
    End If

End Sub
Sub editarCliente()
    Dim tabela As ListObject
    Dim lin, id As Long
        
    Set tabela = Planilha3.ListObjects(1)
    lin = sistema.lbClClientes.ListIndex + 2
    
        If sistema.tbClEditar = True Then
            sistema.lbClClientes.RowSource = ""
            tabela.Range(lin, 1).Value = sistema.txtClNome.Value
            tabela.Range(lin, 2).Value = sistema.txtClCnpj.Value
            tabela.Range(lin, 3).Value = sistema.txtClInsEst.Value
            tabela.Range(lin, 4).Value = sistema.txtClInsMun.Value
            tabela.Range(lin, 5).Value = sistema.txtClCei.Value
            tabela.Range(lin, 6).Value = sistema.txtClEnd.Value
            tabela.Range(lin, 7).Value = sistema.txtClMun.Value
            tabela.Range(lin, 8).Value = sistema.cbClUf.Value
            tabela.Range(lin, 11).Value = sistema.txtClObs.Value
            tabela.Range(lin, 10).Value = sistema.txtClTel.Value
            tabela.Range(lin, 9).Value = sistema.txtClEmail.Value
            Call listaClientes
            MsgBox "Alteração efetuada com sucesso!", vbInformation, "Sucesso"

        Else
            Call listaClientes
        End If
    
End Sub
Sub limparCliente()
    sistema.txtClNome.Value = ""
    sistema.txtClCnpj.Value = ""
    sistema.txtClInsEst.Value = ""
    sistema.txtClInsMun.Value = ""
    sistema.txtClCei.Value = ""
    sistema.txtClEnd.Value = ""
    sistema.txtClMun.Value = ""
    sistema.cbClUf.Value = ""
    sistema.txtClObs.Value = ""
    sistema.txtClTel.Value = ""
    sistema.txtClEmail.Value = ""
End Sub
Sub listaClientes()
    Dim tabela As ListObject
    Dim frota As Long
    
    Set tabela = Planilha3.ListObjects(1)
    sistema.lbClClientes.RowSource = Planilha3.ListObjects(1)
    
End Sub
Sub populaCliente()
    Dim tabela As ListObject
    Dim lin, id As Long
        
    Set tabela = Planilha3.ListObjects(1)
    lin = sistema.lbClClientes.ListIndex + 2

    If sistema.tbClEditar = True Then
        sistema.txtClNome.Value = tabela.Range(lin, 1).Value
        sistema.txtClCnpj.Value = tabela.Range(lin, 2).Value
        sistema.txtClInsEst.Value = tabela.Range(lin, 3).Value
        sistema.txtClInsMun.Value = tabela.Range(lin, 4).Value
        sistema.txtClCei.Value = tabela.Range(lin, 5).Value
        sistema.txtClEnd.Value = tabela.Range(lin, 6).Value
        sistema.txtClMun.Value = tabela.Range(lin, 7).Value
        sistema.cbClUf.Value = tabela.Range(lin, 8).Value
        sistema.txtClObs.Value = tabela.Range(lin, 11).Value
        sistema.txtClTel.Value = tabela.Range(lin, 10).Value
        sistema.txtClEmail.Value = tabela.Range(lin, 9).Value
    Else
        Call limparFrota
    End If
End Sub

