VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sistema 
   Caption         =   "Sistema Terraplenagem"
   ClientHeight    =   13875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15930
   OleObjectBlob   =   "sistema.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Lg As Single
Dim Ht As Single
Dim Fini As Boolean

Private Sub btFunExcluir_Click()
    If tbFunEditar = False Then
        MsgBox "Habilite a Edição e selecione um funcionário para poder excluir!", vbInformation, "Informação"
    ElseIf lbFunc.Value < 1 Or lbFunc.ListIndex = -1 Then
        MsgBox "Selecione um funcionário para excluir!", vbInformation, "Informação"
    Else
        Call excluirFunc
    End If
End Sub

Private Sub btFunLimpar_Click()
    Call limparFunc
End Sub

Private Sub btFunSalvar_Click()
    Call aplication("on")
    If tbFunEditar = False Then
        If txtFunNome <> "" Then
            Call salvarFunc
        Else
            MsgBox "Preencha os campos antes de salvar!!!", vbCritical, "Aviso!"
        End If
    ElseIf sistema.lbFunc.ListIndex = -1 And tbFunEditar = True Then
        MsgBox "Selecione uma opção para alterar", vbCritical, "Aviso"
    Else
        Call editarFunc
    End If
    Call aplication("off")
End Sub


Private Sub cbMovMExcluir_Click()
    Call excluirMovM
End Sub

Private Sub cbMovMFiltro_Change()
    If tbMovMEditar = True Then
        sistema.lbMovMMovs.Clear
        Call listaMovM
    End If
End Sub

Private Sub cbMovMLimpar_Click()
    Call limparMovM
End Sub

Private Sub cbMovMSalvar_Click()
Call aplication("on")
On Error Resume Next
    If tbMovMEditar = False Then
        If cbMovMCliente <> "" Or cbMovMEquip <> "" Or cbMovMFunc <> "" Then
            Call salvarMovM
        Else
            MsgBox "Preencha os campos antes de salvar!!!", vbCritical, "Aviso!"
        End If
    ElseIf sistema.lbMovMMovs.ListIndex = -1 And tbMovMEditar = True Then
        MsgBox "Selecione uma opção para alterar", vbCritical, "Aviso"
    Else
        Call editarMovM
    End If
Call aplication("off")
End Sub

Private Sub CommandButton2_Click()
    sistema.MultiPage2.Visible = True
End Sub



Private Sub lbFunc_Click()
    Call populaFunc
End Sub
Private Sub btnFtExcluir_Click()

    If tgbHabEdicao = False Then
        MsgBox "Habilite a Edição e selecione uma frota para poder excluir!", vbInformation, "Informação"
    ElseIf lbFrotas.Value < 1 Or lbFrotas.ListIndex = -1 Then
        MsgBox "Selecione uma frota para excluir!", vbInformation, "Informação"
    Else
        Call excluirFrota
    End If

End Sub

Private Sub btnSalvarFrota_Click()
Call aplication("on")
    If tgbHabEdicao = False Then
        If txtFtdesc <> "" Or txtFtPlaca <> "" Or txtFtRenavam <> "" Or txtFtAno <> "" Then
            Call salvarFrota
        Else
            MsgBox "Preencha os campos antes de salvar!!!", vbCritical, "Aviso!"
        End If
    ElseIf sistema.lbFrotas.ListIndex = -1 And tgbHabEdicao = True Then
        MsgBox "Selecione uma opção para alterar", vbCritical, "Aviso"
    Else
        Call editarFrota
    End If
    Call aplication("off")
End Sub
'exclui clientes após confirmação
Private Sub cbClExcluir_Click()
    If tbClEditar = False Then
        MsgBox "Habilite a Edição e selecione um cliente para poder excluir!", vbInformation, "Informação"
    ElseIf lbClClientes.Value < 1 Or lbClClientes.ListIndex = -1 Then
        MsgBox "Selecione um cliente para excluir!", vbInformation, "Informação"
    Else
        Call excluirCliente
    End If
    
End Sub
'limpa campos do cadastro de clientes
Private Sub cbClLimpar_Click()
    Call limparCliente
End Sub
'Salvar Clientes no banco
Private Sub cbClSalvar_Click()
    
    Call aplication("on")
    If tbClEditar = False Then
        If txtClNome <> "" Then
            Call salvarCliente
        Else
            MsgBox "Preencha os campos antes de salvar!!!", vbCritical, "Aviso!"
        End If
    ElseIf sistema.lbClClientes.ListIndex = -1 And tbClEditar = True Then
        MsgBox "Selecione uma opção para alterar", vbCritical, "Aviso"
    Else
        Call editarCliente
    End If
    Call aplication("off")
    
End Sub
'limpa campos das frotas
Private Sub CommandButton1_Click()
    Call limparFrota
End Sub



'Popula os campos com as informações das frotas
Private Sub lbClClientes_Click()
    Call populaCliente
End Sub


'Popula os campos com as informações das frotas
Private Sub lbFrotas_Click()
    Dim tabela As ListObject
    Dim lin, id As Long
        
    Set tabela = Planilha2.ListObjects(1)
    lin = sistema.lbFrotas.ListIndex + 2

    If sistema.tgbHabEdicao = True Then
        sistema.txtFtdesc.Value = tabela.Range(lin, 1).Value
        sistema.txtFtCor.Value = tabela.Range(lin, 2).Value
        sistema.txtFtPlaca.Value = tabela.Range(lin, 3).Value
        sistema.txtFtRenavam.Value = tabela.Range(lin, 4).Value
        sistema.txtFtModelo.Value = tabela.Range(lin, 5).Value
        sistema.txtFtAno.Value = tabela.Range(lin, 6).Value
        sistema.txtFtSigla.Value = tabela.Range(lin, 7).Value
        sistema.txtFtStatus.Value = tabela.Range(lin, 9).Value
        sistema.txtFtHorFim.Value = tabela.Range(lin, 10).Value
    Else
        Call limparFrota
    End If
End Sub



Private Sub lbMovMMovs_Click()
    Call populaMovM
End Sub

'Abre a planilha e fecha o sistema
Private Sub MultiPage1_Click(ByVal index As Long)
    If index = 7 Then
        Application.Visible = True
        sistema.Hide
    ElseIf index = 1 Then
        Application.Sheets(5).Select
        Application.Visible = True
        sistema.Hide
    ElseIf index = 2 Then
        Application.Sheets(6).Select
        Application.Visible = True
        sistema.Hide
    ElseIf index = 3 Then
        Application.Sheets(7).Select
        Application.Visible = True
        sistema.Hide
    ElseIf index = 5 Then
        Application.Sheets(9).Select
        Application.Visible = True
        sistema.Hide
    ElseIf index = 6 Then
        Application.Sheets(10).Select
        Application.Visible = True
        sistema.Hide
    End If
End Sub


Private Sub refreshMovM_Click()
    sistema.lbMovMMovs.Clear
    Call listaMovM
End Sub

Private Sub tbFunEditar_Click()
    If sistema.tbFunEditar = True Then
        sistema.lbFunc.Enabled = True
        sistema.lbFunc.BackColor = &H80000005
    Else
        sistema.lbFunc.Enabled = False
        sistema.lbFunc.BackColor = &H80000004
    End If
End Sub

Private Sub tbMovMEditar_Click()
Call aplication("on")
    sistema.lbMovMMovs.Clear
    If tbMovMEditar = True Then
        With sistema.lbMovMMovs
        .Enabled = True
        .BackColor = &H80000005
        End With
        Call listaMovM
    Else
        With sistema.lbMovMMovs
        .Enabled = False
        .BackColor = &H80000004
        End With
    End If
Call aplication("off")
End Sub

'Toglle button Editar do Cadastro de frotas
Private Sub tgbHabEdicao_Click()
    
    If sistema.tgbHabEdicao = True Then
        sistema.lbFrotas.Enabled = True
        sistema.lbFrotas.BackColor = &H80000005
    Else
        sistema.lbFrotas.Enabled = False
        sistema.lbFrotas.BackColor = &H80000004
    End If
    
End Sub
'Toglle button Editar do Cadastro de clientes
Private Sub tbClEditar_Click()
    If sistema.tbClEditar = True Then
        sistema.lbClClientes.Enabled = True
        sistema.lbClClientes.BackColor = &H80000005
    Else
        sistema.lbClClientes.Enabled = False
        sistema.lbClClientes.BackColor = &H80000004
    End If
End Sub

Private Sub txtMovMHorFim_Change()
    Call calc_HorKm
End Sub
Private Sub txtMovMHorIn_Change()
    Call calc_HorKm
End Sub

Private Sub txtMovMHorKm_Change()
    Call calc_VlrTot
End Sub

Private Sub txtMovMKmFim_Change()
    Call calc_HorKm
    Call calc_KmTot
End Sub

Private Sub txtMovMKmIn_Change()
    Call calc_HorKm
    Call calc_KmTot
End Sub

Private Sub txtMovMValUnit_Change()
    Call calc_VlrTot
End Sub

Private Sub UserForm_Activate()
    ActiveSheet.Select
End Sub

'Initialize do sistema
Private Sub UserForm_Initialize()
'RESIZE
    Dim i As Integer, L As Integer, TB
    
    
    
        InitMaxMin Me.Caption
        Ht = Me.Height
        Lg = Me.Width
        
        Application.WindowState = xlMaximized
    '    Me.Height = Application.Height
    '    Me.Width = Application.Width
        Me.Left = Application.Left
        Me.Top = Application.Top
'FECHA RESIZE

    'Deixa o Excel Maximizado
    'Application.WindowState = xlMaximized
    
    'Faz com que o UserForm1 recebe a largura da janela do Excel
    'sistema.Width = Application.UsableWidth
    
    'Faz com que o UserForm1 receba a altura da janela Excel mais 18 pontos
    'sistema.Height = Application.Height
    
    With sistema.lbFrotas
        .Enabled = False
        .BackColor = &H80000004
    End With
    With sistema.lbClClientes
        .Enabled = False
        .BackColor = &H80000004
    End With
    With sistema.lbFunc
        .Enabled = False
        .BackColor = &H80000004
    End With
    With sistema.lbMovMMovs
        .Enabled = False
        .BackColor = &H80000004
    End With
    sistema.txtMovMData.Value = Format(Date, "dd/mm/yyyy")
    sistema.cbMovMFiltro.Value = Year(Date)
    
    'Planilha1.Select
    
    Call listaFrotas
    Call listaClientes
    Call listaFunc
    
    Call populaCbs
    Call populaMovMFiltro
    
End Sub
'RESIZE
Private Sub UserForm_Resize()
    Dim RtL As Single, RtH As Single
        If Me.Width < 300 Or Me.Height < 200 Or Fini Then Exit Sub
        RtL = Me.Width / Lg
        RtH = Me.Height / Ht
        Me.Zoom = IIf(RtL < RtH, RtL, RtH) * 100
End Sub
Private Sub UserForm_Terminate()
    Fini = True
End Sub
'FECHA RESIZE
