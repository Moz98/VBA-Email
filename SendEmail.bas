' Módulo para o envio de E-mail



Sub Send_Email_()

Dim emailApplication As Object
Dim emailItem As Object


Set emailApplication = CreateObject("Outlook.Application")
Set emailItem = emailApplication.CreateItem(0)

'Definindo de onde serão coleatada as informações do e-mail
emailItem.To = Range("'Consulta e Envio'!D11").Value
emailItem.Cc = Range("'Consulta e Envio'!F11").Value
emailItem.Subject = Range("'Consulta e Envio'!H11").Value
emailItem.Body = Range("'Consulta e Envio'!F17").Value

emailItem.display

'Limpando os campos anteriormente definidos
Set emailItem = Nothing
Set emailApplication = Nothing

End Sub
