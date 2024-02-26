Public Sub SaveAttachmentsAndForward(Item As Outlook.MailItem)
    Dim attachment As Outlook.Attachment
    Dim targetFolder As String
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim folderExists As Boolean
    folderExists = False
    
    ' Define your base path
    Dim basePath As String
    basePath = "S:\Touchstone\Catrader\Boston\Deals\CATBonds\"
    
    For Each attachment In Item.Attachments
        If InStr(attachment.FileName, "Investor Presentation") > 0 Then
            ' Generate the string based on the file name
            Dim folderName As String
            folderName = GenerateFolderName(attachment.FileName)
            
            ' Check if folder exists
            If fs.FolderExists(basePath & folderName) Then
                folderExists = True
                Exit For
            Else
                ' Create folder and save attachment
                fs.CreateFolder basePath & folderName
                attachment.SaveAsFile basePath & folderName & "\" & attachment.FileName
            End If
        End If
    Next attachment
    
    If folderExists Then
        MsgBox "Folder already exists."
    Else
        ' Forward the email
        Dim fwdMail As Outlook.MailItem
        Set fwdMail = Item.Forward
        fwdMail.Recipients.Add "philip.buonomo-ext@amundi.com"
        fwdMail.Body = "Hi all," & vbNewLine & "The files below have been saved to " & basePath & folderName & vbNewLine & "Thanks," & vbNewLine & "Philip" & vbNewLine & vbNewLine & fwdMail.Body
        fwdMail.Send
    End If
    
    Set fs = Nothing
End Sub

Function GenerateFolderName(fileName As String) As String
    Dim baseName As String
    Dim endIndex As Integer
    Dim startIndex As Integer
    Dim cleanName As String
    
    ' Find the position of "Investor Presentation" to identify the relevant part of the fileName
    endIndex = InStr(fileName, "Investor Presentation")
    
    ' If "Investor Presentation" is found, work backward to find the start of the name segment
    If endIndex > 0 Then
        ' Extract up to just before "Investor Presentation"
        baseName = Left(fileName, endIndex - 1)
        
        ' Remove any trailing punctuation or spaces that might be left after cutting the string
        baseName = Trim(baseName) ' Trim spaces
        Do While Right(baseName, 1) = "." Or Right(baseName, 1) = "-"
            baseName = Left(baseName, Len(baseName) - 1)
            baseName = Trim(baseName) ' Trim spaces again if necessary
        Loop
        
        ' Replace spaces with nothing to concatenate words
        cleanName = Replace(baseName, " ", "")
        
        ' Optional: Replace other characters as needed (e.g., periods, commas)
        cleanName = Replace(cleanName, ".", "")
        
    Else
        ' If "Investor Presentation" not found, return a default or modified fileName
        cleanName = "DefaultFolderName" ' Adjust according to your needs
    End If
    
    GenerateFolderName = cleanName
End Function

