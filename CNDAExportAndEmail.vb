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

    Public Function ExtractCndaInfo(xlsFilename As String) As CndaUtils.CndaAllInfo
        Dim xlApp As New Excel.Application
        Dim xlWb As Excel.Workbook = xlApp.Workbooks.Open(xlsFilename,, True)
        Dim xlWs As Excel.Worksheet
        Dim xlAllInfo As New CndaUtils.CndaAllInfo()
        For Each xlWs In xlWb.Sheets
            Dim xlInfo As New CndaUtils.CndaInfo()
            With xlInfo
                .CustName = xlWs.Range(NAME_CELL).Text
                .Cnda = xlWs.Range(CNDA_CELL).Text
            End With
            For Each c As Excel.Range In xlWs.Range(TO_COL)
                If c.Text <> "" And c.Row <> 1 Then
                    xlInfo.ToList.Add(c.Text)
                ElseIf c.Text = "" Then
                    Exit For
                End If
            Next
            For Each c As Excel.Range In xlWs.Range(CC_COL)
                If c.Text <> "" And c.Row <> 1 Then
                    xlInfo.CcList.Add(c.Text)
                ElseIf c.Text = "" Then
                    Exit For
                End If
            Next
            For Each c As Excel.Range In xlWs.Range(BCC_COL)
                If c.Text <> "" And c.Row <> 1 Then
                    xlInfo.BccList.Add(c.Text)
                ElseIf c.Text = "" Then
                    Exit For
                End If
            Next
        Next
        ExtractCndaInfo = xlAllInfo
    End Function

    Private Sub ExportAndEmailAll(pptFilename As String, xlsFilename As String, ByVal mailItem As Outlook.MailItem)
        Dim xlApp As New Excel.Application
        Dim xlWb As Excel.Workbook = xlApp.Workbooks.Open(xlsFilename,, True)
        Dim xlWs As Excel.Worksheet
        For Each xlWs In xlWb.Sheets
            Dim name As String = xlWs.Range(NAME_CELL).Text
            Dim cnda As String = xlWs.Range(CNDA_CELL).Text
            If MsgBox($"Generate email for {name} with {pptFilename}?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Dim pdfFilename As String = CNDAPowerPoint.PptToPDF(pptFilename, name, cnda)
                CreateEmailFromWorksheet(pdfFilename, xlWs, mailItem)
            End If
        Next
        xlApp.Quit()
        xlWb = Nothing
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
