Imports System

Public Class CndaOutlookGetFileDialog
    Public Property PptFilename As String = ""
    Public Property XlsFilename As String = ""

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If PptFilename = "" Then
            Dim msgbxstatus As MsgBoxResult = MsgBox("Error PPT file not entered", MsgBoxStyle.RetryCancel)
            If msgbxstatus = MsgBoxResult.Cancel Then
                DialogResult = System.Windows.Forms.DialogResult.Cancel
                Close()
            End If
        ElseIf XlsFilename = "" Then
            Dim msgbxstatus1 As MsgBoxResult = MsgBox("Error XLS file not entered", MsgBoxStyle.RetryCancel)
            If msgbxstatus1 = MsgBoxResult.Cancel Then
                DialogResult = System.Windows.Forms.DialogResult.Cancel
                Close()
            End If
        Else
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
    End Sub

    Private Sub SelectPPT_Button_Click(sender As Object, e As EventArgs) Handles SelectPPT_Button.Click
        OpenPPTFileDialog.ShowDialog()
        PptFilename = OpenPPTFileDialog.FileName
        PPT_Label.Text = PptFilename
    End Sub

    Private Sub SelectXLS_Button_Click(sender As Object, e As EventArgs) Handles SelectXLS_Button.Click
        OpenXLSFileDialog.ShowDialog()
        XlsFilename = OpenXLSFileDialog.FileName
        XLS_Label.Text = XlsFilename
    End Sub

    Private Sub PickEmailFolderButton_Click(sender As Object, e As EventArgs) Handles PickEmailFolderButton.Click
        Dim dg As Outlook.Folder = Globals.ThisAddIn.Application.Session.PickFolder()
        If dg IsNot Nothing Then
            My.Settings.MailFolderId = dg.EntryID
            EmailFolderLabel.Text = dg.Name
        End If
    End Sub
End Class
