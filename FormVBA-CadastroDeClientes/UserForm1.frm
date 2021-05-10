VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   8310.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15000
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cod As Integer

Private Sub CommandButton1_Click()
    MsgBox "Para excluir uma linha de cadastro, é necessário clicar na linha e depois ir no botão de excluir, para atualizar a linha de cadastro, é necessário dar dois clicks na tal linha.", vbOKOnly, "Instruções"
End Sub

Private Sub UserForm_Activate()
    CarregarLista
    lst_cadastro.ColumnWidths = "40;150;200;100;100;100;200;70;100;100;100"
    
End Sub

Private Function RetornaLinha(codbusca As Integer) As Double
    Dim ponteiro As Integer
    ponteiro = 1
    
    Do While Cells(ponteiro, 1).Value <> "" And Cells(ponteiro, 1).Value <> codbusca
        ponteiro = ponteiro + 1
    Loop
    
    If (Cells(ponteiro, 1)) = "" Then
        MsgBox "Código não encontrado"
        RetornaLinha = 0
    Else
        RetornaLinha = ponteiro
    End If
End Function

Private Sub PosicionaNovoCadastro()
    Dim ponteiro As Integer
    ponteiro = 1
    
    Do While Cells(ponteiro, 1).Value <> ""
        ponteiro = ponteiro + 1
    Loop

    Cells(ponteiro, 1).Select
End Sub

Private Sub CarregarLista()
    PosicionaNovoCadastro
    Dim ultimalinha As Double
    ultimalinha = ActiveCell.Row - 1
    lst_cadastro.RowSource = "Cadastro_de_Clientes!A2:K" & ultimalinha
End Sub

Private Sub LiberarTexto()
    txt_nome.Enabled = True
    txt_email.Enabled = True
    txt_cpf.Enabled = True
    txt_telefone.Enabled = True
    txt_cep.Enabled = True
    txt_logradouro.Enabled = True
    txt_numero.Enabled = True
    txt_bairro.Enabled = True
    txt_cidade.Enabled = True
    Txt_estado.Enabled = True
End Sub

Private Sub BloquearTexto()
    txt_nome.Enabled = False
    txt_email.Enabled = False
    txt_cpf.Enabled = False
    txt_telefone.Enabled = False
    txt_cep.Enabled = False
    txt_logradouro.Enabled = False
    txt_numero.Enabled = False
    txt_bairro.Enabled = False
    txt_cidade.Enabled = False
    Txt_estado.Enabled = False
    
    txt_nome.Text = ""
    txt_email.Text = ""
    txt_cpf.Text = ""
    txt_telefone.Text = ""
    txt_cep.Text = ""
    txt_logradouro.Text = ""
    txt_numero.Text = ""
    txt_bairro.Text = ""
    txt_cidade.Text = ""
    Txt_estado.Text = ""
    
End Sub

Private Sub btn_novo_Click()
    PosicionaNovoCadastro
    If (ActiveCell.Row = 2) Then
        cod = 1
        lbl_cod.Caption = "Cod: 1"
    Else
        cod = Cells(ActiveCell.Row - 1, 1).Value + 1
        lbl_cod.Caption = "Cod: " & CStr(cod)
    End If
    
    btn_cadastrar.Enabled = True
    btn_novo.Enabled = False
    
    LiberarTexto
    
End Sub

Private Sub btn_cadastrar_Click()
    Cells(ActiveCell.Row, 1).Value = cod
    Cells(ActiveCell.Row, 2).Value = txt_nome.Text
    Cells(ActiveCell.Row, 3).Value = txt_email.Text
    Cells(ActiveCell.Row, 4).Value = txt_cpf.Text
    Cells(ActiveCell.Row, 5).Value = txt_telefone.Text
    Cells(ActiveCell.Row, 6).Value = txt_cep.Text
    Cells(ActiveCell.Row, 7).Value = txt_logradouro.Text
    Cells(ActiveCell.Row, 8).Value = txt_numero.Text
    Cells(ActiveCell.Row, 9).Value = txt_bairro.Text
    Cells(ActiveCell.Row, 10).Value = txt_cidade.Text
    Cells(ActiveCell.Row, 11).Value = Txt_estado.Text
    
    lbl_cod.Caption = "Cod: "
    
    BloquearTexto
    
    btn_novo.Enabled = True
    btn_cadastrar.Enabled = False
    CarregarLista
End Sub

Private Sub btn_atualizar_Click()
    Cells(RetornaLinha(cod), 1).Value = cod
    Cells(RetornaLinha(cod), 2).Value = txt_nome.Text
    Cells(RetornaLinha(cod), 3).Value = txt_email.Text
    Cells(RetornaLinha(cod), 4).Value = txt_cpf.Text
    Cells(RetornaLinha(cod), 5).Value = txt_telefone.Text
    Cells(RetornaLinha(cod), 6).Value = txt_cep.Text
    Cells(RetornaLinha(cod), 7).Value = txt_logradouro.Text
    Cells(RetornaLinha(cod), 8).Value = txt_numero.Text
    Cells(RetornaLinha(cod), 9).Value = txt_bairro.Text
    Cells(RetornaLinha(cod), 10).Value = txt_cidade.Text
    Cells(RetornaLinha(cod), 11).Value = Txt_estado.Text
    
    lbl_cod.Caption = "Cod: "
    BloquearTexto
    
    btn_atualizar.Enabled = False
    CarregarLista
End Sub

Private Sub btn_excluir_Click()
    Dim linha As Double
    Dim confirmacao As VbMsgBoxResult
    
    confirmacao = MsgBox("Deseja realmente excluir?", vbYesNo, "Confirmação")
    
    If (confirmacao = vbYes) Then
        linha = RetornaLinha(cod)
        
        If linha <> 0 Then
            Rows(linha).Delete
            CarregarLista
            btn_excluir.Enabled = False
        End If
    End If
End Sub

Private Sub lst_cadastro_Click()
    cod = CInt(lst_cadastro.List(lst_cadastro.ListIndex, 0))
    btn_excluir.Enabled = True
End Sub

Private Sub lst_cadastro_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cod = CInt(lst_cadastro.List(lst_cadastro.ListIndex, 0))
    
    lbl_cod.Caption = "Cod: " & CStr(cod)
    
    txt_nome.Text = lst_cadastro.List(lst_cadastro.ListIndex, 1)
    txt_email.Text = lst_cadastro.List(lst_cadastro.ListIndex, 2)
    txt_cpf.Text = lst_cadastro.List(lst_cadastro.ListIndex, 3)
    txt_telefone.Text = lst_cadastro.List(lst_cadastro.ListIndex, 4)
    txt_cep.Text = lst_cadastro.List(lst_cadastro.ListIndex, 5)
    txt_logradouro.Text = lst_cadastro.List(lst_cadastro.ListIndex, 6)
    txt_numero.Text = lst_cadastro.List(lst_cadastro.ListIndex, 7)
    txt_bairro.Text = lst_cadastro.List(lst_cadastro.ListIndex, 8)
    txt_cidade.Text = lst_cadastro.List(lst_cadastro.ListIndex, 9)
    Txt_estado.Text = lst_cadastro.List(lst_cadastro.ListIndex, 10)
    
    btn_atualizar.Enabled = True
    LiberarTexto
    
End Sub

