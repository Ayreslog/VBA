VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tela 
   Caption         =   "TELA"
   ClientHeight    =   8685.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16200
   OleObjectBlob   =   "Tela.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

On Error GoTo Erro

Variaveis.Ref

Dim Conf As String
Conf = InputBox("Informe sua senha!", "SENHA")

If Conf = "" Then
Exit Sub
End If


Do

            Linha = Linha + 1
            
            If Plan.Cells(Linha, ColunaUser).Text = Plan.Range("E1").Text Then
                        
                    Senha = Plan.Cells(Linha, ColunaSenha).Text
                        
                                If Conf <> Senha Then
                                
                                MsgBox "Senha incorreta!", vbCritical, "SENHA"
                                Exit Do
                                
                                Else
                                
                                Application.Visible = True

                                Permitir = "OK"
                                
                                Unload Me
                                
                                End If

                        Exit Do
            End If
            
            
            If Plan.Cells(Linha, ColunaUser).Value = "" Then
                        MsgBox "Usuário incorreto!", vbCritical, "USUÁRIO"
                        Exit Do
            End If

Loop Until Plan.Cells(Linha, ColunaUser).Value = ""


Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ABRIR"

End Sub

Private Sub CommandButton10_Click()

Permitir = "OK"

Unload Me
On Error Resume Next
Usuarios.Show

End Sub

Private Sub CommandButton3_Click()

Alterar_Senha.Show

End Sub

Private Sub CommandButton4_Click()

On Error Resume Next
ThisWorkbook.Save
Application.Quit


End Sub

Private Sub UserForm_Initialize()

Permitir = ""

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If Permitir <> "" Then
Exit Sub
End If


On Error Resume Next
ThisWorkbook.Save

Application.Quit

End Sub
