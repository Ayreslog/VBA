Attribute VB_Name = "Módulo1"
Sub SalvarComoTXT()
    UserForm1.Show
End Sub

Sub ExecutarSalvarTXT(mPlan As Worksheet, mPathSave As String)
Dim NovoArquivoXLS As Workbook

    'Cria um novo arquivo excel
    Set NovoArquivoXLS = Application.Workbooks.Add

    'Copia a planilha para o novo arquivo criado
    mPlan.Copy Before:=NovoArquivoXLS.Sheets(1)

    'Salva o arquivo
    Application.DisplayAlerts = False
    NovoArquivoXLS.SaveAs mPathSave & "\" & mPlan.Name & ".txt", _
        FileFormat:=xlText, CreateBackup:=False
    
    NovoArquivoXLS.Close
    Set NovoArquivoXLS = Nothing
    Application.DisplayAlerts = True
    
    MsgBox "Novo arquivo salvo em: " & mPathSave & "\" & mPlan.Name & ".txt", vbInformation

End Sub



