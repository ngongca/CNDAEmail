Imports System

Public Class CndaOutlookGenPDFandEmailFileDialog
    Private PptFilename As String = ""
    Private XmlFilename As String = ""
    Property GeneratePdf As Boolean = True

    Public Event GeneratePdfEvent(ByVal pptFilename As String, ByVal xmlFilename As String)
    Public Event GenerateEmailEvent(ByVal pptFilename As String, ByVal xmlFilename As String)

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If PptFilename = "" Then
            Dim msgbxstatus As MsgBoxResult = MsgBox("Error PPT file not entered", MsgBoxStyle.RetryCancel)
            If msgbxstatus = MsgBoxResult.Cancel Then
                DialogResult = System.Windows.Forms.DialogResult.Cancel
                Close()
            End If
        ElseIf XmlFilename = "" Then
            Dim msgbxstatus1 As MsgBoxResult = MsgBox("Error XLS file not entered", MsgBoxStyle.RetryCancel)
            If msgbxstatus1 = MsgBoxResult.Cancel Then
                DialogResult = System.Windows.Forms.DialogResult.Cancel
                Close()
            End If
        Else
            OK_Button.Enabled = False
            If GeneratePdf Then
                InfoLabel.Visible = True
                InfoLabel.Text = "Generating PDFs"
                Update()
                RaiseEvent GeneratePdfEvent(PptFilename, XmlFilename)
            End If
            InfoLabel.Text = "Saving emails"
            OK_Button.Enabled = False
            Update()
            RaiseEvent GenerateEmailEvent(PptFilename, XmlFilename)
            DialogResult = System.Windows.Forms.DialogResult.OK
            Close()
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim f As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetFolderFromID(My.Settings.MailFolderId)
        EmailFolderLabel.Text = f.Name
        XmlFilename = My.Settings.XmlFileName
        XLS_Label.Text = XmlFilename
    End Sub

    Private Sub SelectPPT_Button_Click(sender As Object, e As EventArgs) Handles SelectPPT_Button.Click
        OpenPPTFileDialog.ShowDialog()
        PptFilename = OpenPPTFileDialog.FileName
        PPT_Label.Text = PptFilename
    End Sub

    Private Sub SelectXLS_Button_Click(sender As Object, e As EventArgs) Handles SelectXLS_Button.Click
        OpenXMLFileDialog.ShowDialog()
        XmlFilename = OpenXMLFileDialog.FileName
        XLS_Label.Text = XmlFilename
        My.Settings.XmlFileName = XmlFilename
        My.Settings.Save()
    End Sub


    Private Sub PickEmailFolderButton_Click(sender As Object, e As EventArgs) Handles PickEmailFolderButton.Click
        Dim dg As Outlook.Folder = Globals.ThisAddIn.Application.Session.PickFolder()
        If dg IsNot Nothing Then
            My.Settings.MailFolderId = dg.EntryID
            EmailFolderLabel.Text = dg.Name
            My.Settings.Save()
        End If
    End Sub

End Class
