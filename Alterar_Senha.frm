VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Alterar_Senha 
   Caption         =   "TROCAR SENHA"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   OleObjectBlob   =   "Alterar_Senha.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Alterar_Senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBFechar_Click()

On Error Resume Next
ThisWorkbook.Save

Unload Me

End Sub

Private Sub TSenha_Exit(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo Erro

If TSenha = "" Then
MsgBox "Digite senha válida!", vbCritical, "SENHA"
Exit Sub
End If

Variaveis.Ref

Dim NSENHA As String

With Plan

    Do
    
    Linha = Linha + 1
    
        If .Cells(Linha, ColunaUser).Text = TUser.Text Then
        
            Senha = .Cells(Linha, ColunaSenha).Text
               
               If TSenha <> Senha Then
                  
                MsgBox "Senha incorreta!", vbCritical, "SENHA"
                TSenha = ""
                Exit Sub
                
                Else
                
                
                NSENHA = InputBox("Informe sua nova senha!", "SENHA")

                If NSENHA = "" Then
                
                        MsgBox "Senha Inválida!", vbCritical, "SENHA"
                        Exit Sub
                        
                        Else
                        
                        .Cells(Linha, ColunaSenha).Value = NSENHA
                        
                        MsgBox "Senha alterada para_!" & NSENHA, vbInformation, "NOVA SENHA"

                        TSenha = ""
                        
                        Exit Sub

                End If


               End If
               
        End If
        
    
    Loop Until .Cells(Linha, ColunaUser).Value = ""
    

End With

MsgBox "Usuário não encontrado!", vbCritical, "SENHA"


Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "TROCAR SENHA"


End Sub

Private Sub TUser_Change()

TUser = VBA.UCase(TUser.Text)

End Sub

Private Sub UserForm_Initialize()

Variaveis.Ref

TUser = Plan.Cells(1, 5).Value
CBFechar.SetFocus


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

On Error Resume Next
ThisWorkbook.Save

End Sub
