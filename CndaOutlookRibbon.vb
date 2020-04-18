Imports Microsoft.Office.Interop
Imports Microsoft.Office.Tools.Ribbon
Imports System.IO
Public Class CndaOutlookRibbon
    Private CndaOlkMdl As CndaOutlookModel
    Private thisEmail As Outlook.MailItem
    Private WithEvents GenPdf As CndaOutlookGenPDFandEmailFileDialog

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        CndaOlkMdl = New CndaOutlookModel()
        CndaOlkMdl.InitModel()
    End Sub

    Private Sub CndaEmailExportAndEmail_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles CNDAExportAndEmail_Button.Click
        Dim m As Outlook.Inspector = e.Control.Context
        thisEmail = TryCast(m.CurrentItem, Outlook.MailItem)
        GenPdf = New CndaOutlookGenPDFandEmailFileDialog With {
            .GeneratePdf = True
        }
        GenPdf.PptFileInstructionLabel.Text = "CNDA Outlook Generate PDF"
        GenPdf.ShowDialog()
        If MsgBox("Email generation complete" & vbCrLf & "Do you with to remove the current email?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
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
        If MsgBox($"Email generation complete" & vbCrLf & "Do you with to remove the current email?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            thisEmail.Close(Outlook.OlInspectorClose.olDiscard)
        End If
    End Sub
    Private Sub GenPdfEventHandler(pptFilename As String, ByRef obj As List(Of CndaCustInfo)) Handles GenPdf.GeneratePdfEvent
        CNDAPowerPoint.PptToPDFs(pptFilename, obj)
    End Sub

    Private Sub GenEmailEventHandler(pptFilename As String, ByRef mailCnt As Integer) Handles GenPdf.GenerateEmailEvent
        mailCnt = 0
        If thisEmail IsNot Nothing Then
            For Each c As CndaCustInfo In CndaOlkMdl.CustInfoList
                Dim pdfFileName As String = CNDAPowerPoint.CndaPdfString(PptFilename:=pptFilename, c.Cnda, c.CustName)
                If File.Exists(pdfFileName) Then
                    CndaOlkMdl.CreateEmailWithAttachment(pdfFileName, c, thisEmail)
                    mailCnt += 1
                Else
                    If MsgBox($"{$"Could not find pdf file {pdfFileName}, no email generated"}{vbCrLf}Continue?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub
    Private Sub XmlFileChangeEventHander(ByVal xmlFilename As String,
                                         ByRef custCollection As List(Of CndaCustInfo)) Handles GenPdf.XmlFileChangeEvent
        My.Settings.XmlFileName = xmlFilename
        My.Settings.Save()
        CndaOlkMdl.UpdateModel(xmlFilename:=xmlFilename)
        custCollection = CndaOlkMdl.CustInfoList
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
                CndaOlkMdl.UpdateModel(dlg.XmlFilename)
                For Each c As CndaCustInfo In CndaOlkMdl.CustInfoList
                    CndaOlkMdl.CreateEmailWithAttachment("", c, mailItem)
                Next
                If MsgBox("Email generation complete. See your Drafts folder." & vbCrLf & "Do you with to remove the current email?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    mailItem.Close(Outlook.OlInspectorClose.olDiscard)
                End If
            End If
        End If
    End Sub

End Class
