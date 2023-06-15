VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Saidas 
   Caption         =   "Saídas de Caixa"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13920
   OleObjectBlob   =   "Saidas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Saidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

On Error GoTo Erro


Dim F As Controle
Set F = Controle

Linha = ListBox1.ListIndex

F.TCod = ListBox1.List(Linha, 0)
F.LData = ListBox1.List(Linha, 1)
F.LMes = ListBox1.List(Linha, 2)
F.LAno = ListBox1.List(Linha, 3)
F.TValor = ListBox1.List(Linha, 4)
F.CConta = ListBox1.List(Linha, 5)
F.CSubconta = ListBox1.List(Linha, 6)
F.Tmultas = ListBox1.List(Linha, 7)
F.TDesconto = ListBox1.List(Linha, 8)

F.MultiPage1.Value = 0

F.CBEditar.Enabled = True
F.CBExcluir.Enabled = True
F.CBSalvar.Enabled = False

Pesquisar = "OK"

Unload Me




Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"

End Sub

Private Sub UserForm_Click()

End Sub

Sub Cabecalho()

With ListBox1
.AddItem
.List(0, 0) = "Código"
.List(0, 1) = "Data"
.List(0, 2) = "Mês"
.List(0, 3) = "Ano"
.List(0, 4) = "Valor"
.List(0, 5) = "Conta"
.List(0, 6) = "Subconta"
.List(0, 7) = "Multas"
.List(0, 8) = "Descontos"
End With

ListBox1.ColumnWidths = "40;50;60;50;50;125;125;50;50"

End Sub
Sub Filtro()

On Error GoTo Erro

Dim Linha, LinhaListbox As Long
Dim ULTIMALINHA As Variant


LinhaListbox = 1
Linha = 5

ListBox1.Clear

Call Cabecalho

Planilha1.Activate

With Planilha1

        While .Cells(Linha, 2).Value <> ""
        
                With ListBox1
                .AddItem
                .List(LinhaListbox, 0) = Planilha1.Cells(Linha, 2)
                .List(LinhaListbox, 1) = Planilha1.Cells(Linha, 3).Text
                .List(LinhaListbox, 2) = Planilha1.Cells(Linha, 4)
                .List(LinhaListbox, 3) = Planilha1.Cells(Linha, 5)
                .List(LinhaListbox, 4) = Planilha1.Cells(Linha, 6)
                .List(LinhaListbox, 5) = Planilha1.Cells(Linha, 7)
                .List(LinhaListbox, 6) = Planilha1.Cells(Linha, 8)
                .List(LinhaListbox, 7) = Planilha1.Cells(Linha, 9)
                .List(LinhaListbox, 8) = Planilha1.Cells(Linha, 10)
                End With
                
                LinhaListbox = LinhaListbox + 1
                
                Linha = Linha + 1
        
        Wend
    


End With

Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"


End Sub

Private Sub UserForm_Initialize()

Call Filtro

Pesquisar = ""

End Sub












