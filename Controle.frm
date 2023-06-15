VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Controle 
   Caption         =   "CADASTRO DE PERNOITES"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11205
   OleObjectBlob   =   "Controle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Controle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBEditar_Click()


On Error GoTo Erro

If TCod.Value = "" Or TValor.Value = "" Or CConta.Value = "" Or CSubconta.Value = "" Or Ndoc.Value = "" Then
MsgBox "Favor preencher todos os campos!", vbCritical, "SALVAR"
Exit Sub
End If

Planilha1.Activate

Dim Plan As String
Plan = Planilha1.Name

With Worksheets(Plan).Range("B:B")

Set C = .Find(TCod.Value, LookIn:=xlValues, Lookat:=xlWhole)

If Not C Is Nothing Then
C.Select
Call Salvar_Saida
MsgBox "Editado com sucesso!", vbInformation, "EDITAR"
Call Limpar_Saidas
Else
MsgBox "Código não encontrado!", vbCritical, "EDITAR"
End If


End With


Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"


End Sub

Private Sub CBExcluir_Click()

On Error GoTo Erro

Planilha1.Activate

Dim Plan As String
Plan = Planilha1.Name

With Worksheets(Plan).Range("B:B")

Set C = .Find(TCod.Value, LookIn:=xlValues, Lookat:=xlWhole)

If Not C Is Nothing Then
C.Select
Selection.EntireRow.Delete
MsgBox "Excluído com sucesso!", vbInformation, "EDITAR"
Call Limpar_Saidas
Else
MsgBox "Código não encontrado!", vbCritical, "EDITAR"
End If


End With

Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"


End Sub

Private Sub CBNovo_Click()

Call Limpar_Saidas

End Sub

Private Sub CBSalvar_Click()

On Error GoTo Erro

If TCod.Value = "" Or TValor.Value = "" Or CConta.Value = "" Or CSubconta.Value = "" Then
MsgBox "Favor preencher todos os campos!", vbCritical, "SALVAR"
Exit Sub
End If

Dim Ver As Double
Ver = WorksheetFunction.CountIf(Planilha1.Range("B:B"), TCod)


If Ver > 0 Then
MsgBox "Código já cadastrado!", vbCritical, "SALVAR"
Exit Sub
End If

Planilha1.Activate
Planilha1.Range("B4").Select

Módulo1.Localizar_Celula_Vazia

Call Salvar_Saida

MsgBox "Salvo com sucesso!", vbInformation, "SALVAR"

Call Limpar_Saidas


Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"

End Sub

Private Sub CConta_Change()

On Error GoTo Erro

On Error Resume Next
CSubconta.Clear

Planilha3.Activate

Dim Plan As String
Plan = Planilha3.Name

With Worksheets(Plan).Rows(4)

Set C = .Find(CConta.Value, LookIn:=xlValues, Lookat:=xlWhole)

If Not C Is Nothing Then
C.Select
ActiveCell.Offset(1, 0).Select

If ActiveCell.Value <> "" Then

CSubconta.Enabled = True
Range(Selection, Selection.End(xlDown)).Select
CSubconta.List = Application.Transpose(Selection)
Else
CSubconta.Enabled = False
End If

End If

End With


Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"


End Sub

Private Sub CConta_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Pesquisar <> "" Then
Exit Sub
End If


Planilha1.Select

Dim Cod As Long
On Error Resume Next
Cod = Range("B100000").End(xlUp).Offset(0, 0).Value

TCod = Cod + 1

If Cod < 1 Then
TCod = "1"
End If


End Sub

Private Sub CommandButton5_Click()

ThisWorkbook.Save

End Sub


Sub Data_Atual()

LData.Caption = Date
LMes.Caption = Format(Now, "mmmm")
LMes.Caption = UCase(LMes.Caption)
LAno.Caption = Format(Now, "yyyy")

End Sub

Private Sub CommandButton6_Click()

Saidas.Show

End Sub

Private Sub CSubconta_Change()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub TMultas_Change()

End Sub

Private Sub TValor_Change()

End Sub

Private Sub TValor_Exit(ByVal Cancel As MSForms.ReturnBoolean)


If Pesquisar <> "" Then
Exit Sub
End If

Planilha1.Select

Dim Cod As Long
On Error Resume Next
Cod = Range("B100000").End(xlUp).Offset(0, 0).Value

TCod = Cod + 1

If Cod < 1 Then
TCod = "1"
End If


End Sub

Private Sub UserForm_Initialize()
Tmultas.AddItem ("ARUJÁ BULKY")
Tmultas.AddItem ("JUNDIAI BULKY")
Tmultas.AddItem ("PINHAIS BULKY")
Tmultas.AddItem ("RIO BULKY")
Tmultas.AddItem ("FULFILLMENT")
Tmultas.AddItem ("FCB SP")
Tmultas.AddItem ("FCB MG")
Tmultas.AddItem ("STL")
Call Data_Atual
Call Carregar_Conta

End Sub
Sub Limpar_Saidas()

TValor = ""
CConta = ""
CSubconta = ""
TDesconto = ""
Tmultas = ""
TCod = ""
Pesquisar = ""

CBEditar.Enabled = False
CBExcluir.Enabled = False
CBSalvar.Enabled = True

Call Data_Atual

End Sub
Sub Salvar_Saida()

On Error GoTo Erro

Dim Data As Date
Data = LData.Caption

Dim Ano As Double
Dim Valor, Multa, Desconto, Id As Double

Ano = LAno.Caption
Valor = TValor.Value
Id = TCod.Value

If Tmultas.Value <> "" Then
Multa = Tmultas.Value
End If

If TDesconto.Value <> "" Then
Desconto = TDesconto.Value
End If

ActiveCell.Value = Id
ActiveCell.Offset(0, 1).Value = Data
ActiveCell.Offset(0, 2).Value = LMes.Caption
ActiveCell.Offset(0, 3).Value = Ano
ActiveCell.Offset(0, 4).Value = Valor
ActiveCell.Offset(0, 5).Value = CConta.Text
ActiveCell.Offset(0, 6).Value = CSubconta.Text
ActiveCell.Offset(0, 7).Value = Multa
ActiveCell.Offset(0, 8).Value = Desconto
ActiveCell.Offset(0, 10).Value = Ndoc





Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"

End Sub
Sub Carregar_Conta()

On Error GoTo Erro

Dim Linha As Integer
Linha = 4

Do

Linha = Linha + 1
CConta.AddItem Planilha3.Cells(Linha, 3).Value

Loop Until Planilha3.Cells(Linha, 3).Value = ""


Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"

End Sub























