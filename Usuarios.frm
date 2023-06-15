VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usuarios 
   Caption         =   "Registro de Usuário e Senha"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4800
   OleObjectBlob   =   "Usuarios.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBRegistrar_Click()

On Error GoTo Erro

If TUsuario = "" Or TSenha = "" Then
MsgBox "Precisa preencher todos os campos!", vbCritical, "REGISTRO"
TUsuario.SetFocus
Exit Sub
End If

Variaveis.Ref

Dim Ver As Double
Ver = WorksheetFunction.CountIf(Plan.Range("B:B"), TUsuario.Text)


If Ver > 0 Then
MsgBox "Usuário já cadastrado!", vbCritical, "REGISTRO"
Exit Sub
End If


With Plan

   Do
   Linha = Linha + 1
   Loop Until .Cells(Linha, ColunaUser).Value = ""
   
   .Cells(Linha, ColunaUser).Value = TUsuario.Text
   .Cells(Linha, ColunaSenha).Value = TSenha.Text
   .Cells(1, 1).Value = "OK"
    
    MsgBox "Usuário Registrado!", vbInformation, "REGISTRO"

    TUsuario = ""
    TSenha = ""

End With



Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "REGISTRAR"


End Sub

Private Sub CommandButton1_Click()

Permitir = "OK"

On Error Resume Next
ThisWorkbook.Save

Unload Me
Login.Show

End Sub


Private Sub TUsuario_Change()

TUsuario = VBA.UCase(TUsuario.Text)

End Sub

Private Sub UserForm_Initialize()

Variaveis.Ref

If Plan.Cells(1, 1).Value <> "" Then
Else
MsgBox "Seja Bem Vindo ao Registro de Usuários!", vbInformation, "REGISTRO"
End If

 Permitir = ""
  
 
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If Permitir <> "" Then
Else
Cancel = True
MsgBox "Utilize o botão Fechar!", vbCritical, "FECHAR"

End If


End Sub








