Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook

Public Class CndaOutlookEmailView
    Implements CNDAEmail.ICndaOutlookEvents
    Public Property XmlFilename As String = "<enter xml file>"
    Public Property MailFolderName As String = "<enter mail folder>"
    Private EmailsGenerated As Boolean = False

    Public Event XmlFileChangeEvent(ByVal xmlFilename As String,
                                    ByRef objList As CheckedListBox.ObjectCollection) Implements ICndaOutlookEvents.XmlFileChangeEvent
    Public Event EmailFolderChangeEvent(ByRef emailFolder As Outlook.Folder) Implements ICndaOutlookEvents.EmailFolderChangeEvent
    Public Event SendEmailsEvent(ByRef objList As CheckedListBox.CheckedItemCollection,
                                 ByRef count As Integer) Implements ICndaOutlookEvents.SendEmailsEvent
    Public Event PptFileChangeEvent(pptFilename As String) Implements ICndaOutlookEvents.PptFileChangeEvent

    Private Sub CndaOutlookEmailView_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        XlsFilenameLabel.Text = XmlFilename
        RaiseEvent XmlFileChangeEvent(XmlFilename, EmailViewCheckedListBox.Items)
        For i = 0 To EmailViewCheckedListBox.Items.Count - 1
            EmailViewCheckedListBox.SetItemChecked(i, CheckState.Checked)
        Next
        EmailFolderLabel1.Text = MailFolderName
        Update()
    End Sub


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button1.Click
        If XmlFilename = "" Then
            Dim msgbxstatus1 As MsgBoxResult = MsgBox("Error XLS file not entered", MsgBoxStyle.RetryCancel)
            If msgbxstatus1 = MsgBoxResult.Cancel Then
                DialogResult = System.Windows.Forms.DialogResult.Cancel
                Close()
            End If
        Else
            If Not EmailsGenerated Then
                WorkingLabel.Visible = True
                WorkingLabel.Text = "Generating Emails..."
                Dim count As Integer
                RaiseEvent SendEmailsEvent(EmailViewCheckedListBox.CheckedItems, count)
                WorkingLabel.Text = $"CNDA generated {count} emails in your {MailFolderName} folder" & vbCrLf _
                & "Do you wish to delete the current email?"
                OK_Button1.Text = "YES"
                Cancel_Button1.Text = "NO"
                EmailsGenerated = True
            Else
                DialogResult = DialogResult.Yes
                Close()
            End If
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button1.Click
        If EmailsGenerated Then
            DialogResult = DialogResult.No
        Else
            DialogResult = System.Windows.Forms.DialogResult.Cancel
            Close()
        End If
    End Sub
    Private Sub GetXlsButton_Click(sender As Object, e As EventArgs) Handles GetXlsButton.Click
        If XmlOpenFileDialog2.ShowDialog = DialogResult.OK Then
            XmlFilename = XmlOpenFileDialog2.FileName
            RaiseEvent XmlFileChangeEvent(XmlFilename, EmailViewCheckedListBox.Items)
            For i = 0 To EmailViewCheckedListBox.Items.Count - 1
                EmailViewCheckedListBox.SetItemChecked(i, CheckState.Checked)
            Next
            Update()
        End If
    End Sub

    Private Sub GetEmailFolderButton_Click(sender As Object, e As EventArgs) Handles GetEmailFolderButton.Click
        Dim folder As Folder = Globals.ThisAddIn.Application.Session.PickFolder()
        If folder IsNot Nothing Then
            RaiseEvent EmailFolderChangeEvent(folder)
            MailFolderName = folder.Name
            EmailFolderLabel1.Text = folder.Name
        End If
    End Sub

End Class
