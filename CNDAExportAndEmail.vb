Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports Microsoft.Office.Tools.Ribbon

Public Class CNDAExportAndEmail
    Private Const TO_COL As String = "C2:C50"
    Private Const CC_COL As String = "D2:D50"
    Private Const BCC_COL As String = "E2:E50"
    Private Const NAME_CELL As String = "A2"
    Private Const CNDA_CELL As String = "B2"
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub CNDAUpdateEmail_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles CNDAUpdateEmail_Button.Click
        Dim df As New GetFileDialog
        If df.ShowDialog() = Global.System.Windows.Forms.DialogResult.OK Then
            Dim m As Outlook.Inspector = e.Control.Context
            Dim mailItem As Outlook.MailItem = TryCast(m.CurrentItem, Outlook.MailItem)
            If mailItem IsNot Nothing Then
                Dim pptFilename As String = df.GetPptFilename()
                Dim xlsFilename As String = df.GetXlsFilename()
                ExportAndEmailAll(pptFilename, xlsFilename, mailItem)
                If MsgBox("Email generation complete. See your Drafts folder." & vbCrLf & "Do you with to remove the current email?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    mailItem.Close(Outlook.OlInspectorClose.olDiscard)
                End If
            End If
        End If
    End Sub
    Private Sub ExportAndEmailAll(pptFilename As String, xlsFilename As String, ByVal mailItem As Outlook.MailItem)
        Dim xlCndaInfo As New CndaBaseClasses.CndaAllInfo()
        xlCndaInfo = CndaExcel.ExtractCndaInfo(xlsFilename)
        Dim i As Integer = CNDAPowerPoint.PptToPDFs(pptFilename, xlCndaInfo)
        If i > 0 Then
            For Each c As CndaInfo In xlCndaInfo.CndaInfos
                Dim pdfname As String = CNDAPowerPoint.CndaPdfString(pptFilename, c.Cnda, c.CustName)
                CreateCndaEmail(pdfname, c, mailItem)
            Next
        End If
    End Sub

    Private Sub CreateCndaEmail(pdfFilename As String, info As CndaBaseClasses.CndaInfo, RefMail As Outlook.MailItem)
        If (RefMail IsNot Nothing) Then
            Dim curMail As Outlook.MailItem = RefMail.Copy
            Dim unused = curMail.Attachments.Add(Source:=pdfFilename)
            For Each c As String In info.ToList
                Dim recipient1 As Outlook.Recipient = curMail.Recipients.Add(c)
                recipient1.Type = Outlook.OlMailRecipientType.olTo
            Next
            For Each c As String In info.CcList
                Dim recipient1 As Outlook.Recipient = curMail.Recipients.Add(c)
                recipient1.Type = Outlook.OlMailRecipientType.olCC
            Next
            For Each c As String In info.BccList
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
    Private Sub CreateEmailFromWorksheet(pdfFilename As String, xlWs As Excel.Worksheet, ByVal refMail As Outlook.MailItem)
        If (refMail IsNot Nothing) Then
            Dim curMail As Outlook.MailItem = refMail.Copy
            Dim unused = curMail.Attachments.Add(Source:=pdfFilename)
            Dim range As Excel.Range = xlWs.Range(TO_COL)
            For Each c In range
                If c.Text <> "" And c.Row <> 1 Then
                    Dim recipient1 As Outlook.Recipient = curMail.Recipients.Add(c.Text)
                    recipient1.Type = Outlook.OlMailRecipientType.olTo
                ElseIf c.Text = "" Then
                    Exit For
                End If
            Next
            range = xlWs.Range(CC_COL)
            For Each c In range
                If c.Text <> "" And c.Row <> 1 Then
                    Dim recipient As Outlook.Recipient = curMail.Recipients.Add(Name:=c.Value2)
                    recipient.Type = Outlook.OlMailRecipientType.olCC
                ElseIf c.Value = "" Then
                    Exit For
                End If
            Next
            range = xlWs.Range(BCC_COL)
            For Each c In range
                If c.Value <> "" And c.Row <> 1 Then
                    Dim recipient2 As Outlook.Recipient = curMail.Recipients.Add(c.Value)
                    recipient2.Type = Outlook.OlMailRecipientType.olBCC
                ElseIf c.Value = "" Then
                    Exit For
                End If
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
