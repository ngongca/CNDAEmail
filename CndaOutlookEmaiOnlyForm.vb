Imports System.Windows.Forms

Public Class CndaOutlookEmailOnlyForm
    Public Property XlsFilename As String

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub CndaOutlookEmaiOnlyForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        XlsFilename = My.Settings.XlsFileName
        XlsFilenameLabel.Text = XlsFilename
        Dim mf As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetFolderFromID(My.Settings.MailFolderId)
        EmailFolderLabel.Text = mf.Name
    End Sub

    Private Sub GetXlsButton_Click(sender As Object, e As EventArgs) Handles GetXlsButton.Click
        If XmlOpenFileDialog2.ShowDialog = DialogResult.OK Then
            XlsFilename = XmlOpenFileDialog2.FileName
            My.Settings.XlsFileName = XlsFilename
            XlsFilenameLabel.Text = XlsFilename
        End If
    End Sub

    Private Sub GetEmailFolderButton_Click(sender As Object, e As EventArgs) Handles GetEmailFolderButton.Click
        Dim dg As Outlook.Folder = Globals.ThisAddIn.Application.Session.PickFolder()
        If dg IsNot Nothing Then
            My.Settings.MailFolderId = dg.FolderPath
            EmailFolderLabel.Text = dg.Name
        End If
    End Sub
End Class
