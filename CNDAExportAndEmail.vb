Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports Microsoft.Office.Tools.Ribbon

Public Class CNDAExportAndEmail
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub CNDAUpdateEmail_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles CNDAUpdateEmail_Button.Click
        Dim df As New GetFileDialog
        If df.ShowDialog() = Global.System.Windows.Forms.DialogResult.OK Then
            Dim m As Outlook.Inspector = e.Control.Context
            Dim mailItem As Outlook.MailItem = TryCast(m.CurrentItem, Outlook.MailItem)
            If mailItem IsNot Nothing Then
                Dim xlCndaInfo As New CndaAllInfo()
                'TODO create constructor of CndaAllInfo with Excel filename)
                xlCndaInfo = CndaExcel.ExtractCndaInfo(df.GetXlsFilename)
                If CNDAPowerPoint.PptToPDFs(df.GetPptFilename, xlCndaInfo) > 0 Then
                    For Each c As CndaInfo In xlCndaInfo.CndaInfos
                        CreateEmailWithAttachment(CNDAPowerPoint.CndaPdfString(df.GetPptFilename, c.Cnda, c.CustName), c,
                            mailItem)
                    Next
                    If MsgBox("Email generation complete. See your Drafts folder." & vbCrLf & "Do you with to remove the current email?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        mailItem.Close(Outlook.OlInspectorClose.olDiscard)
                    End If
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Create a copy of a reference email based on the Cnda Info, attaches a file and moves to current draft folder
    ''' </summary>
    ''' <param name="AttachmentName"></param>
    ''' <param name="Info"></param>
    ''' <param name="RefMail"></param>
    Private Sub CreateEmailWithAttachment(AttachmentName As String, Info As CndaBaseClasses.CndaInfo, RefMail As Outlook.MailItem)
        If (RefMail IsNot Nothing) Then
            Dim curMail As Outlook.MailItem = RefMail.Copy
            Dim unused = curMail.Attachments.Add(Source:=AttachmentName)
            For Each c As String In Info.ToList
                Dim recipient1 As Outlook.Recipient = curMail.Recipients.Add(c)
                recipient1.Type = Outlook.OlMailRecipientType.olTo
            Next
            For Each c As String In Info.CcList
                Dim recipient1 As Outlook.Recipient = curMail.Recipients.Add(c)
                recipient1.Type = Outlook.OlMailRecipientType.olCC
            Next
            For Each c As String In Info.BccList
                Dim recipient1 As Outlook.Recipient = curMail.Recipients.Add(c)
                recipient1.Type = Outlook.OlMailRecipientType.olBCC
            Next
            Dim folder As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
            If folder Is Nothing Then
                MsgBox("Error cannot find Drafts folder in Outlook", MsgBoxStyle.Critical)
            Else
                curMail.Move(folder)
            End If
        End If
    End Sub
End Class
