VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Login 
   Caption         =   "Login"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4785
   OleObjectBlob   =   "Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBEntrar_Click()

Unload Me
Tela.Show


End Sub

Private Sub CBSair_Click()

ThisWorkbook.Save
Application.Quit

End Sub

Private Sub TSenha_Change()

CBEntrar.Enabled = False

End Sub
Private Sub TSenha_Exit(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo Erro

        If TSenha = "" Then
                MsgBox "Senha Inv�lida!", vbCritical, "SENHA"
                Exit Sub
        End If

Variaveis.Ref

Do

Linha = Linha + 1

            If Plan.Cells(Linha, ColunaUser).Text = TUser.Text Then
            
            Senha = Plan.Cells(Linha, ColunaSenha).Text
                        
                    If TSenha <> Senha Then
                    
                            MsgBox "Senha incorreta!", vbCritical, "SENHA"
                            TSenha = ""
                            Exit Do
                            
                            Else
                            
                            Permitir = "OK"
                    
                           Plan.Cells(1, 5).Value = TUser.Text
                                                     
                           CBEntrar.Enabled = True
                    
                           MsgBox "Ol�:!" & TUser.Text, vbInformation, "LOGIN"
                    
                    End If
            
            
            Exit Do
            End If


If Plan.Cells(Linha, ColunaUser).Value = "" Then
MsgBox "Usu�rio incorreto!", vbCritical, "USU�RIO"
TUser = ""
Exit Do
End If

Loop Until Plan.Cells(Linha, ColunaUser).Value = ""


Exit Sub
Erro:
MsgBox "ERRO!", vbCritical, "SENHA"


End Sub
Private Sub TUser_Change()

TUser = VBA.UCase(TUser.Text)

CBEntrar.Enabled = False

End Sub
Private Sub TUser_Exit(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo Erro

    If TUser = "" Then
        MsgBox "Usu�rio Inv�lido!", vbCritical, "USU�RIO"
        TSenha.Enabled = False
        Exit Sub
    End If

Variaveis.Ref

Do

            Linha = Linha + 1
            
            If Plan.Cells(Linha, ColunaUser).Text = TUser.Text Then
            TSenha.Enabled = True
            Exit Do
            End If
            
            
            If Plan.Cells(Linha, ColunaUser).Value = "" Then
            MsgBox "Usu�rio incorreto!", vbCritical, "USU�RIO"
            TUser = ""
            Exit Do
            End If

Loop Until Plan.Cells(Linha, ColunaUser).Value = ""


Exit Sub

Erro:
MsgBox "Erro!", vbCritical, "USU�RIO"


End Sub
Private Sub UserForm_Initialize()

Permitir = ""

CBSair.SetFocus

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If Permitir <> "" Then
Exit Sub
End If


On Error Resume Next
ThisWorkbook.Save
Application.Quit


End Sub


