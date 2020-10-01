Attribute VB_Name = "macro"
Public Sub SaveAttachments(Item As Outlook.MailItem)

If Item.Attachments.Count > 0 Then
 
Dim objAttachments As Outlook.Attachments
Dim lngCount As Integer
Dim strFile As String
Dim sFileType As String
  Dim i As Integer
Dim olFrom As String
Dim olExt As String


'Definindo o email do remetente, extensão do arquivo e local para salvar os anexos
olFrom = "someone@email.com"
olExt = "xlsx"
strFolderPath = "C:\Temp\"


Set objAttachments = Item.Attachments
    lngCount = objAttachments.Count
 
    
    For i = lngCount To 1 Step -1
    
    If Item.SenderEmailAddress = olFrom And Right$(objAttachments.Item(i).FileName, 4) = olExt Then

       ' Obtendo o nome do arquivo anexo.
       sFile = objAttachments.Item(i).FileName

       
       ' Concatenando a pasta com o nome do anexo.
       strFile = strFolderPath & sFile

       ' Salvando anexo no local definido.
       objAttachments.Item(i).SaveAsFile strFile
    
    End If

    Next i

End If

End Sub
