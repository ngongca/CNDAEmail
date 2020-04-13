Imports Microsoft.Office.Interop
Imports Microsoft.Office.Tools.Ribbon
Imports System.IO


Public Class CNDAExportAndEmail

    Private WithEvents GenPdf As CndaOutlookGenPDFandEmailFileDialog
    Private thisEmail As Outlook.MailItem

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        'set default folder
        If My.Settings.MailFolderId Is "" Then
            Dim df As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
            My.Settings.MailFolderId = df.EntryID
            My.Settings.Save()
        End If
    End Sub

    Private Sub CndaEmailExportAndEmail_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles CNDAExportAndEmail_Button.Click
        Dim m As Outlook.Inspector = e.Control.Context
        thisEmail = TryCast(m.CurrentItem, Outlook.MailItem)
        GenPdf = New CndaOutlookGenPDFandEmailFileDialog With {
            .GeneratePdf = True
        }
        GenPdf.PptFileInstructionLabel.Text = "CNDA Outlook Generate PDF"
        GenPdf.ShowDialog()
        If MsgBox("Email generation complete. See your Drafts folder." & vbCrLf & "Do you with to remove the current email?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            thisEmail.Close(Outlook.OlInspectorClose.olDiscard)
        End If
    End Sub
    ''' <summary>
    ''' Gets PPT and XLS file from user and then generates CNDA emails using existing PDF files that were generated using the NDA tools.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CndaEmailButton_Click(sender As Object, e As RibbonControlEventArgs) Handles CNDAEmailButton.Click
        Dim m As Outlook.Inspector = e.Control.Context
        thisEmail = TryCast(m.CurrentItem, Outlook.MailItem)
        GenPdf = New CndaOutlookGenPDFandEmailFileDialog With {
            .GeneratePdf = False
        }
        GenPdf.PptFileInstructionLabel.Text = "PPT file used to Generate PDF"
        GenPdf.ShowDialog()
        If MsgBox("Email generation complete. See your Drafts folder." & vbCrLf & "Do you with to remove the current email?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            thisEmail.Close(Outlook.OlInspectorClose.olDiscard)
        End If
    End Sub
    Private Sub GenPdfEventHandler(pptFilename As String, xmlFilename As String) Handles GenPdf.GeneratePdfEvent
        Dim xlCndaInfo As CndaAllInfo = CndaXmlToAllInfo(xmlFilename)
        CNDAPowerPoint.PptToPDFs(pptFilename, xlCndaInfo)
    End Sub

    Private Sub GenEmailEventHandler(pptFilename As String, xmlFilename As String) Handles GenPdf.GenerateEmailEvent
        If thisEmail IsNot Nothing Then
            Dim xlCndaInfo As CndaAllInfo = CndaXmlToAllInfo(XmlFileName:=xmlFilename)
            For Each c As CndaCustInfo In xlCndaInfo.CndaInfos
                Dim pdfFileName As String = CNDAPowerPoint.CndaPdfString(PptFilename:=pptFilename, c.Cnda, c.CustName)
                If File.Exists(pdfFileName) Then
                    CreateEmailWithAttachment(pdfFileName, c, thisEmail)
                Else
                    If MsgBox($"{$"Could not find pdf file {pdfFileName}, no email generated"}{vbCrLf}Continue?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' A click on this button generates emails only based on Cnda Info without attachments
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CndaEmailOnlyButton_Click(sender As Object, e As RibbonControlEventArgs) Handles CNDAEmailOnlyButton.Click
        Dim dlg As New CndaOutlookEmailOnlyForm()
        If dlg.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim m As Outlook.Inspector = e.Control.Context
            Dim mailItem As Outlook.MailItem = TryCast(m.CurrentItem, Outlook.MailItem)
            If mailItem IsNot Nothing Then
                Dim xlCndaInfo As CndaAllInfo = CndaXmlToAllInfo(dlg.XmlFilename)
                For Each c As CndaCustInfo In xlCndaInfo.CndaInfos
                    CreateEmailWithAttachment("", c, mailItem)
                Next
                If MsgBox("Email generation complete. See your Drafts folder." & vbCrLf & "Do you with to remove the current email?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    mailItem.Close(Outlook.OlInspectorClose.olDiscard)
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' Create a copy of a reference email based on the Cnda Info, attaches a file it it exists and moves to current draft folder
    ''' </summary>
    ''' <param name="AttachmentName">Name of file to attach.  If Nothing, then no attachment will be made</param>
    ''' <param name="Info"></param>
    ''' <param name="RefMail"></param>
    Private Sub CreateEmailWithAttachment(AttachmentName As String, Info As CndaBaseClasses.CndaCustInfo, RefMail As Outlook.MailItem)
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
End Class
