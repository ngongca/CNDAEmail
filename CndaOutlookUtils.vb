Imports System.IO
Module CndaOutlookUtils
    ''' <summary>
    ''' Create a copy of a reference email based on the Cnda Info, attaches a file it it exists and moves to outlook <paramref name="Folder"/>
    ''' </summary>
    ''' <param name="AttachmentName">Name of file to attach.  If Nothing, then no attachment will be made</param>
    ''' <param name="Info"></param>
    ''' <param name="RefMail"></param>
    ''' <param name="Folder"></param>
    Public Sub CreateEmailWithAttachment(AttachmentName As String,
                                         Info As CndaCustInfo,
                                         RefMail As Outlook.MailItem,
                                         Folder As Outlook.Folder)
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

            If Folder Is Nothing Then
                MsgBox($"Error cannot find {Folder.MailFolderId} folder in Outlook", MsgBoxStyle.Critical)
            Else
                curMail.Move(Folder)
            End If
        End If
    End Sub
End Module
