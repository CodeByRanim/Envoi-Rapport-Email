Sub SendEmailWithAttachment()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim FilePath As String
    
    ' Spécifiez le chemin du fichier à envoyer
    FilePath = "C:\chemin\vers\le\rapport.xlsx"
    
    ' Créer une instance d'Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    ' Définir les propriétés de l'email
    With OutlookMail
        .To = "destinataire@example.com"
        .Subject = "Rapport mensuel"
        .Body = "Bonjour, veuillez trouver ci-joint le rapport mensuel."
        .Attachments.Add FilePath
        .Send
    End With
    
    MsgBox "Email envoyé avec succès !"
End Sub
