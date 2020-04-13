Imports System.Windows.Forms

Public Class CndaOutlookEmailOnlyForm
    Public Property XmlFilename As String

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        DialogResult = System.Windows.Forms.DialogResult.OK
        Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        Close()
    End Sub

    Private Sub CndaOutlookEmaiOnlyForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Settings.XmlFileName = "" Then
            XmlFilename = "<select XML file>"
        Else
            XmlFilename = My.Settings.XmlFileName
        End If
        XlsFilenameLabel.Text = XmlFilename
        Dim mf As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetFolderFromID(My.Settings.MailFolderId)
        If mf IsNot Nothing Then
            EmailFolderLabel.Text = mf.Name
        Else
            EmailFolderLabel.Text = "<select folder>"
        End If

    End Sub

    Private Sub GetXlsButton_Click(sender As Object, e As EventArgs) Handles GetXlsButton.Click
        If XmlOpenFileDialog2.ShowDialog = DialogResult.OK Then
            XmlFilename = XmlOpenFileDialog2.FileName
            My.Settings.XmlFileName = XmlFilename
            XlsFilenameLabel.Text = XmlFilename
            My.Settings.Save()
        End If
    End Sub

    Private Sub GetEmailFolderButton_Click(sender As Object, e As EventArgs) Handles GetEmailFolderButton.Click
        Dim dg As Outlook.Folder = Globals.ThisAddIn.Application.Session.PickFolder()
        If dg IsNot Nothing Then
            My.Settings.MailFolderId = dg.FolderPath
            EmailFolderLabel.Text = dg.Name
            My.Settings.Save()
        End If
    End Sub
End Class
