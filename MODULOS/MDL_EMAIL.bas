Attribute VB_Name = "MDL_EMAIL"
Sub enviarEmailRelatorio()
Dim arquivo As String
Dim titulo() As String
Dim filename As String
Dim formatedTitle As String
Dim id As Variant
id = Robot.Cells(ActiveCell.Row, 6).value
arquivo = Application.WorksheetFunction.VLookup(id, TBL_CADASTRO_RELATORIOS.Range("A:D"), 4, False)
filename = Robot.Cells(ActiveCell.Row, 7).value
titulo = Split(filename, ".")
Debug.Print titulo(0)
Debug.Print emailsTo(1)
End Sub


Function emailSignature() As String
emailSignature = "" & _
"</html>" & _
"</html>"
End Function

Function emailsTo(id As Integer) As String
    emailsTo = "moises@gmail.com"
End Function

Function emailsCopy(id As Integer) As String
    emailsCopy = "moises@gmail.com"
End Function

Function emailsBCC(id As Integer) As String
    emailsBCC = "jsilva.moises@gmail.com"
End Function


Public Sub EnviarEmailComAssinatura()
On Error GoTo TRATAR_ERRO

Dim outlook_ As outlook.Application
Dim email  As outlook.MailItem
Dim AskResult As Integer

Set outlook_ = New outlook.Application
Set email = outlook_.CreateItem(olMailItem)


With email
    .Display
    .Session.Accounts.Item (1)
    .To = "moises.juvenal@gmail.com"
    .CC = "moises.juvenal@gmail.com"
    .BCC = "jsilva.moises@gmail.com"
    .Subject = "Assunto"
    .Body = "Olá"
    
    .HTMLBody = "<html>" & _
    "<h1>Bem Vindo Moisés Juvenal da Silva</h2>" & _
    "</html>" & .HTMLBody
    .Display
    .Attachments.Add ("C:\Users\Usuario\Dropbox\@@@ExcelVBA\README.md")
     AskResult = MsgBox("Confirma o envio do Email?", vbYesNo)
     Select Case AskResult
     Case vbYes
          .Send
     Case vbNo
          .Delete
          Set outlook_ = Nothing
          Set email = Nothing
     End Select
End With

Set outlook_ = Nothing
Set email = Nothing
On Error GoTo 0
Exit Sub
TRATAR_ERRO:
Dim erro As String
erro = Err.Number & vbNewLine & Err.Description
MsgBox erro, vbCritical + vbOKOnly, "Enviar Email"



End Sub
