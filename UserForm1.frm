VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Selecionar a Planilha para Salvar em TXT"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    'Chama a rotina para salvar como txt
    'Será salvo um novo arquivo txt com base na planilha seleciona na lista de opções
    Call ExecutarSalvarTXT(Sheets(lstPlanilhas.Text), ThisWorkbook.Path)
    
    Unload Me   'Fecha o form
    
End Sub

Private Sub lstPlanilhas_Change()

End Sub

Private Sub UserForm_Initialize()
    
    'Chama a rotina para preencher a lista das planilha disponíveis no arquivo
    Call PreencheLista
    
End Sub

Private Sub PreencheLista()
Dim sht As Worksheet

    lstPlanilhas.Clear
    
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name <> "Principal" Then 'Não exibe a planilha Principal
            lstPlanilhas.AddItem sht.Name
        End If
    Next sht
    
End Sub
