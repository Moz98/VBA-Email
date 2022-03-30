
'Comandos iniciar da abertura do Forms
Private Sub UserForm_Activate()
'Conteúdo da Drop List (Situação)
With Me.ComboSituação
    .Clear
    .AddItem ""
    .AddItem "Ativo"
    .AddItem "Inativo"
End With
' Mostrar o Conteúdo da DropBox (Banco de Dados)
Call Refresh_data
End Sub
'======================== ATUALIZAÇÃO DA LISTBOX =================================
Sub Refresh_data()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Banco de Dados")
Dim last_row As Long
last_row = Application.WorksheetFunction.CountA(sh.Range("A:A"))

With Me.ListData
    .ColumnHeads = True
    .ColumnCount = 14
    
    If last_row = 1 Then
        .RowSource = "'Banco de Dados'!A2:N2"
    Else
        .RowSource = "'Banco de Dados'!A2:N" & last_row
    End If
End With

End Sub

'======================== BOTÃO ADICIONAR =================================
Private Sub cmdAddData_Click()
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Banco de Dados")
Dim last_row As Long
last_row = Application.WorksheetFunction.CountA(sh.Range("A:A"))

'------------------- Validação ---------------------------
                  'Situação
If Me.ComboSituação.Value = "" Then
MsgBox "Por Favor insira a situação da Empresa", vbExclamation
Exit Sub
End If
                  'Nome da Empresa
If Me.txtNome.Value = "" Then
MsgBox "Por Favor insira o nome da Empresa", vbExclamation
Exit Sub
End If
                  'Nome do Responsável
If Me.txtResponsável.Value = "" Then
MsgBox "Por Favor insira o nome do responsável", vbExclamation
Exit Sub
End If

'--------------------------------------------------------------
            'Valores dos Campos Automático
sh.Range("A" & last_row + 1).Value = Me.ComboSituação
sh.Range("B" & last_row + 1).Value = "=Row()-1" 'Preenchimento automático - ID
sh.Range("C" & last_row + 1).Value = Me.TxtCNPJ
sh.Range("D" & last_row + 1).Value = Me.txtSigla
sh.Range("E" & last_row + 1).Value = Me.txtNome
sh.Range("F" & last_row + 1).Value = Me.txtEndereço
sh.Range("G" & last_row + 1).Value = Me.txtComplemento
sh.Range("H" & last_row + 1).Value = Me.txtCEP
sh.Range("I" & last_row + 1).Value = Me.txtCidade
sh.Range("J" & last_row + 1).Value = Me.txtResponsável
sh.Range("K" & last_row + 1).Value = Me.txtCargo
sh.Range("L" & last_row + 1).Value = Me.txtEmail
sh.Range("M" & last_row + 1).Value = Me.txtTelefone
sh.Range("N" & last_row + 1).Value = Now 'Preenchimento automático - Data e Horário
'--------------------------------------------------------------

Me.ComboSituação.Value = ""
Me.TxtCNPJ.Value = ""
Me.txtSigla.Value = ""
Me.txtNome.Value = ""
Me.txtEndereço.Value = ""
Me.txtComplemento.Value = ""
Me.txtCEP.Value = ""
Me.txtCidade.Value = ""
Me.txtResponsável.Value = ""
Me.txtCargo.Value = ""
Me.txtEmail.Value = ""
Me.txtTelefone.Value = ""
'--------------------------------------------------------------

Call Refresh_data
End Sub
'======================== BOTÃO ATUALIZAR =================================
'Selecionar as informações na Lista do Banco de Dados
Private Sub ListData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.txtIndice.Value = Me.ListData.List(Me.ListData.ListIndex, 1) 'Demonstra o ID
Me.ComboSituação.Value = Me.ListData.List(Me.ListData.ListIndex, 0)
Me.TxtCNPJ.Value = Me.ListData.List(Me.ListData.ListIndex, 2)
Me.txtSigla.Value = Me.ListData.List(Me.ListData.ListIndex, 3)
Me.txtNome.Value = Me.ListData.List(Me.ListData.ListIndex, 4)
Me.txtEndereço.Value = Me.ListData.List(Me.ListData.ListIndex, 5)
Me.txtComplemento.Value = Me.ListData.List(Me.ListData.ListIndex, 6)
Me.txtCEP.Value = Me.ListData.List(Me.ListData.ListIndex, 7)
Me.txtCidade.Value = Me.ListData.List(Me.ListData.ListIndex, 8)
Me.txtResponsável.Value = Me.ListData.List(Me.ListData.ListIndex, 9)
Me.txtCargo.Value = Me.ListData.List(Me.ListData.ListIndex, 10)
Me.txtEmail.Value = Me.ListData.List(Me.ListData.ListIndex, 11)
Me.txtTelefone.Value = Me.ListData.List(Me.ListData.ListIndex, 12)

End Sub

Private Sub cmdAtualizar_Click()

'-----------------------------------------------------------
If Me.txtIndice.Value = "" Then
MsgBox "Selecione alguma empresa", vbInformation
Exit Sub
End If
'-----------------------------------------------------------
                  'Atualização
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Banco de Dados")
Dim selected_row As Long
selected_row = Application.WorksheetFunction.Match(CLng(Me.txtIndice.Value), sh.Range("B:B"), 0)

'------------------- Validação ---------------------------
                  'Situação
If Me.ComboSituação.Value = "" Then
MsgBox "Por Favor insira a situação da Empresa", vbExclamation
Exit Sub
End If
                  'Nome da Empresa
If Me.txtNome.Value = "" Then
MsgBox "Por Favor insira o nome da Empresa", vbExclamation
Exit Sub
End If
                  'Nome do Responsável
If Me.txtResponsável.Value = "" Then
MsgBox "Por Favor insira o nome do responsável", vbExclamation
Exit Sub
End If

'--------------------------------------------------------------
            'Valores dos Campos Automático
sh.Range("A" & selected_row).Value = Me.ComboSituação
sh.Range("C" & selected_row).Value = Me.TxtCNPJ
sh.Range("D" & selected_row).Value = Me.txtSigla
sh.Range("E" & selected_row).Value = Me.txtNome
sh.Range("F" & selected_row).Value = Me.txtEndereço
sh.Range("G" & selected_row).Value = Me.txtComplemento
sh.Range("H" & selected_row).Value = Me.txtCEP
sh.Range("I" & selected_row).Value = Me.txtCidade
sh.Range("J" & selected_row).Value = Me.txtResponsável
sh.Range("K" & selected_row).Value = Me.txtCargo
sh.Range("L" & selected_row).Value = Me.txtEmail
sh.Range("M" & selected_row).Value = Me.txtTelefone
sh.Range("N" & selected_row).Value = Now 'Preenchimento automático - Data e Horário
'--------------------------------------------------------------

Me.ComboSituação.Value = ""
Me.TxtCNPJ.Value = ""
Me.txtSigla.Value = ""
Me.txtNome.Value = ""
Me.txtEndereço.Value = ""
Me.txtComplemento.Value = ""
Me.txtCEP.Value = ""
Me.txtCidade.Value = ""
Me.txtResponsável.Value = ""
Me.txtCargo.Value = ""
Me.txtEmail.Value = ""
Me.txtTelefone.Value = ""
Me.txtIndice.Value = ""
'--------------------------------------------------------------

Call Refresh_data


End Sub
'======================== BOTÃO RESETAR =================================
Private Sub cmdReset_Click()
Me.ComboSituação.Value = ""
Me.TxtCNPJ.Value = ""
Me.txtSigla.Value = ""
Me.txtNome.Value = ""
Me.txtEndereço.Value = ""
Me.txtComplemento.Value = ""
Me.txtCEP.Value = ""
Me.txtCidade.Value = ""
Me.txtResponsável.Value = ""
Me.txtCargo.Value = ""
Me.txtEmail.Value = ""
Me.txtTelefone.Value = ""



End Sub
'======================== BOTÃO DELETAR =================================
Private Sub cmdDelete_Click()
If Me.txtIndice.Value = "" Then
MsgBox "Selecione alguma empresa", vbInformation
Exit Sub
End If

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Banco de Dados")
Dim selected_row As Long
selected_row = Application.WorksheetFunction.Match(CLng(Me.txtIndice.Value), sh.Range("B:B"), 0)
'---------------------------------------------------------------------------------------------------
sh.Range("A" & selected_row).EntireRow.Delete
'---------------------------------------------------------------------------------------------------
Me.ComboSituação.Value = ""
Me.TxtCNPJ.Value = ""
Me.txtSigla.Value = ""
Me.txtNome.Value = ""
Me.txtEndereço.Value = ""
Me.txtComplemento.Value = ""
Me.txtCEP.Value = ""
Me.txtCidade.Value = ""
Me.txtResponsável.Value = ""
Me.txtCargo.Value = ""
Me.txtEmail.Value = ""
Me.txtTelefone.Value = ""

Call Refresh_data


End Sub


'======================== BOTÃO SAIR =================================
Private Sub cmdExit_click()
Dim iExit As VbMsgBoxResult

iExit = MsgBox("Você deseja sair ?", vbQuestion + vbYesNo, "Sair")

If iExit = vbYes Then
Unload Me
End If

End Sub

