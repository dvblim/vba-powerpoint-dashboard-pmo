VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGerador 
   Caption         =   "Painel de Geração de Relatórios PMO"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   OleObjectBlob   =   "frmGerador.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGerador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Código que roda quando você clica no botão do formulário
Private Sub btnGerar_Click()
    Dim dias As Long
    
    ' 1. Verifica qual opção de dias foi marcada
    If opt30.Value = True Then dias = 30
    If opt60.Value = True Then dias = 60
    If opt90.Value = True Then dias = 90
    If optFull.Value = True Then dias = 0 ' 0 representa o total
    
    ' 2. Pequena validação: o usuário marcou algum slide?
    If chkSlide1.Value = False And chkSlide2.Value = False Then
        MsgBox "Por favor, selecione pelo menos um slide para gerar!", vbExclamation
        Exit Sub
    End If
    
    ' 3. Feedback visual
    lblStatus.Caption = "Processando PowerPoint... por favor aguarde."
    lblStatus.ForeColor = RGB(0, 0, 255)
    DoEvents ' Faz o Windows atualizar a frase na tela
    
    ' 4. CHAMA A MACRO PRINCIPAL enviando as escolhas do formulário
    Call GerarRelatorioMaster_Pro(dias, chkSlide1.Value, chkSlide2.Value)
    
    ' 5. Finalização
    lblStatus.Caption = "Relatório Gerado com Sucesso!"
    lblStatus.ForeColor = RGB(0, 120, 0)
    MsgBox "O PowerPoint foi gerado!", vbInformation
    Unload Me ' Fecha o formulário
End Sub

' Configuração que roda assim que o formulário abre
Private Sub UserForm_Initialize()
    opt90.Value = True ' Deixa os 90 dias marcados por padrão
    chkSlide1.Value = True
    chkSlide2.Value = True
End Sub
