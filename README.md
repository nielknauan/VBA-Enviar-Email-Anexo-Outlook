HHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHH
'Desenvolvido por: Nielk Nauan Carvalho dos Santos
'Nome da Macro: Sub enviaremail()
'Data de Desenvolvimento: 17/01/2023
'Função da Macro: enviar email automatico
'HHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHH
    
Sub EnviarAbaAnexada()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim Destinatario As String
    Dim Assunto As String
    Dim Mensagem As String
    Dim Anexo As String
    Dim NomeAnexo As String
    Dim NomeAba As String
    
    Destinatario = "email@exemplo.com"
    Assunto = "Aba anexada"
    Mensagem = "Segue aba anexada"
    NomeAba = "Planilha01"
    Anexo = "C:\Caminho\valido\coleta.xlsx"
    NomeAnexo = "coleta"
    
    ThisWorkbook.Sheets(3).Copy
    ActiveWorkbook.SaveAs Filename:=Anexo
    ActiveWorkbook.Close
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    OutMail.Display
    
    With OutMail
        .To = Destinatario
        .CC = ""
        .BCC = ""
        .Subject = Assunto
        .Body = Mensagem
        .Attachments.Add Anexo, 1, 1, NomeAnexo
        .Send
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub

