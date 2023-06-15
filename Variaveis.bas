Attribute VB_Name = "Variaveis"
Public Permitir As String
Public ColunaUser As Double
Public Plan As Object
Public Linha As Double
Public ColunaSenha As Double
Public Senha As String
Sub Ref()

Set Plan = Planilha1 'alterar
ColunaUser = 2 ' alterar
Linha = 1 ' alterar
ColunaSenha = ColunaUser + 1


End Sub


' RECORTE O CÓDIGO A SEGUIR E COLE NO EVENTO OPEN DA PLANILHA


'Application.Visible = False

'Variaveis.Ref

'If Plan.Range("A1") = "" Then
'Usuarios.Show
'Else
'Login.Show
'End If
