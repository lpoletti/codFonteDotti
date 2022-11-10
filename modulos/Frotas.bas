Attribute VB_Name = "Frotas"
Option Explicit
Sub salvarFrota()
    Dim tabela As ListObject
    Dim lin, id As Long
        
    Set tabela = Planilha2.ListObjects(1)
    id = Range("id").Value
    lin = tabela.Range.Rows.Count
    
    tabela.Range(lin, 8).Value = id
    tabela.Range(lin, 1).Value = sistema.txtFtdesc.Value
    tabela.Range(lin, 2).Value = sistema.txtFtCor.Value
    tabela.Range(lin, 3).Value = sistema.txtFtPlaca.Value
    tabela.Range(lin, 4).Value = sistema.txtFtRenavam.Value
    tabela.Range(lin, 5).Value = sistema.txtFtModelo.Value
    tabela.Range(lin, 6).Value = sistema.txtFtAno.Value
    tabela.Range(lin, 7).Value = sistema.txtFtSigla.Value
    tabela.Range(lin, 9).Value = sistema.txtFtStatus.Value
    tabela.Range(lin, 10).Value = sistema.txtFtHorFim.Value
    
    tabela.ListRows.Add
    Range("id").Value = id + 1
    MsgBox "Cadastrado com sucesso", vbInformation, "Cadastro Frota"
    Call limparFrota
    
End Sub
Sub excluirFrota()
    Dim tabela As ListObject
    Dim lin As Long
    Dim result As VbMsgBoxResult
        
    Set tabela = Planilha2.ListObjects(1)
    lin = sistema.lbFrotas.ListIndex + 2
    result = MsgBox("Quer mesmo deletar?", vbYesNo, "Cuidado")
    
    If result = 6 Then
        sistema.lbFrotas.RowSource = ""
        tabela.Range(lin, 1).EntireRow.Delete
        Call listaFrotas
    Else
        Exit Sub
    End If

End Sub
Sub editarFrota()
    Dim tabela As ListObject
    Dim lin, id As Long
        
    Set tabela = Planilha2.ListObjects(1)
    lin = sistema.lbFrotas.ListIndex + 2
    
        If sistema.tgbHabEdicao = True Then
            sistema.lbFrotas.RowSource = ""
            tabela.Range(lin, 1).Value = sistema.txtFtdesc.Value
            tabela.Range(lin, 2).Value = sistema.txtFtCor.Value
            tabela.Range(lin, 3).Value = sistema.txtFtPlaca.Value
            tabela.Range(lin, 4).Value = sistema.txtFtRenavam.Value
            tabela.Range(lin, 5).Value = sistema.txtFtModelo.Value
            tabela.Range(lin, 6).Value = sistema.txtFtAno.Value
            tabela.Range(lin, 7).Value = sistema.txtFtSigla.Value
            tabela.Range(lin, 9).Value = sistema.txtFtStatus.Value
            tabela.Range(lin, 10).Value = sistema.txtFtHorFim.Value
            Call listaFrotas
            MsgBox "Alteração efetuada com sucesso!", vbInformation, "Sucesso"
            Call limparFrota
        Else
            Call limparFrota
        End If
    
End Sub
Sub limparFrota()
    sistema.txtFtdesc.Value = ""
    sistema.txtFtCor.Value = ""
    sistema.txtFtPlaca.Value = ""
    sistema.txtFtRenavam.Value = ""
    sistema.txtFtModelo.Value = ""
    sistema.txtFtAno.Value = ""
    sistema.txtFtSigla.Value = ""
    sistema.txtFtStatus.Value = ""
    sistema.txtFtHorFim.Value = ""
End Sub
Sub listaFrotas()
    Dim tabela As ListObject
    Dim frota As Long
    
    Set tabela = Planilha2.ListObjects(1)
    sistema.lbFrotas.RowSource = Planilha2.ListObjects(1)
    
End Sub

Sub abrir_arquivo()
    Workbooks.Open "C:\Users\Luan\Google Drive\LpinfoCaxias\CLIENTES\Cópia de CAIXA DOTTI ATUALIZADO.xlsm"
End Sub
