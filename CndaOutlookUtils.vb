Imports System.IO.Path
Module CndaOutlookUtils
    ''' <summary>
    ''' Create a copy of a reference email based on the Cnda Info, attaches a file it it exists and moves to current draft folder
    ''' </summary>
    ''' <param name="AttachmentName">Name of file to attach.  If Nothing, then no attachment will be made</param>
    ''' <param name="Info"></param>
    ''' <param name="RefMail"></param>
    Public Sub CreateEmailWithAttachment(AttachmentName As String, Info As CndaBaseClasses.CndaCustInfo, RefMail As Outlook.MailItem)
        If (RefMail IsNot Nothing) Then
            Dim curMail As Outlook.MailItem = RefMail.Copy
            If File.Exists(AttachmentName) Then
                Dim unused = curMail.Attachments.Add(Source:=AttachmentName)
            End If
            For Each addr As CndaMailListItem In Info.AddrList
                Dim recipient As Outlook.Recipient = curMail.Recipients.Add(addr.Address)
                Select Case addr.AddressType
                    Case CndaMailListItem.AddressTypeEnum.MailTo
                        recipient.Type = Outlook.OlMailRecipientType.olTo
                    Case CndaMailListItem.AddressTypeEnum.MailCC
                        recipient.Type = Outlook.OlMailRecipientType.olCC
                    Case CndaMailListItem.AddressTypeEnum.MailBCC
                        recipient.Type = Outlook.OlMailRecipientType.olBCC
                End Select
            Next addr

            'Dim folder As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(My.Settings.MailFolder)
            Dim folder As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetFolderFromID(My.Settings.MailFolderId)
            If folder Is Nothing Then
                MsgBox($"Error cannot find {My.Settings.MailFolderId} folder in Outlook", MsgBoxStyle.Critical)
            Else
                curMail.Move(folder)
            End If
        End If
    End Sub
End Module
