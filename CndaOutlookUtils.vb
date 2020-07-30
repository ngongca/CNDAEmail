Imports System.IO
Module CndaOutlookUtils
    ''' <summary>
    ''' Create a copy of a reference email based on the Cnda Info, attaches a file it it exists./>
    ''' </summary>
    ''' <param name="AttachmentName">Name of file to attach.  If Nothing, then no attachment will be made</param>
    ''' <param name="Info"></param>
    ''' <param name="RefMail"></param>
    Public Sub CreateEmail(AttachmentName As String,
                           Info As CndaCustInfo,
                           RefMail As Outlook.MailItem,
                           AttachPDf As Boolean)
        If (RefMail IsNot Nothing) Then
            Dim curMail As Outlook.MailItem = RefMail.Copy
            If AttachPDf Then
                If File.Exists(AttachmentName) Then
                    Dim unused = curMail.Attachments.Add(Source:=AttachmentName)
                Else
                    MsgBox($"Error cannot find {AttachmentName} to attach to email. Email will not contain attachment", MsgBoxStyle.Information)
                End If
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
            curMail.Save()
        End If
    End Sub
End Module
