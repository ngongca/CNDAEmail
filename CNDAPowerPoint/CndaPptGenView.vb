Imports System.Windows.Forms
Public Class CndaPptGenView
    Public Property XmlFilename As String
    Private IsComplete As Boolean = False
    Public Event GenPdfEvent(ByRef objList As CheckedListBox.CheckedItemCollection,
                             ByRef count As Integer)
    Public Event PptXmlFileChangeEvent(ByVal xmlFilename As String,
                                    ByRef objList As CheckedListBox.ObjectCollection)

    Private Sub CndaPptGenView_Load() Handles MyBase.Load
        If XmlFilename <> "" Then
            XmlFilenameLabel.Text = XmlFilename
            RaiseEvent PptXmlFileChangeEvent(XmlFilename, PptCheckedListBox.Items)
            For i = 0 To PptCheckedListBox.Items.Count - 1
                PptCheckedListBox.SetItemCheckState(i, CheckState.Checked)
            Next
            Update()
        End If
    End Sub

    Private Sub GetXmlButton_Click() Handles GetXmlButton.Click
        If XmlOpenFileDialog2.ShowDialog() = DialogResult.OK Then
            XmlFilename = XmlOpenFileDialog2.FileName
            XmlFilenameLabel.Text = XmlFilename
            RaiseEvent PptXmlFileChangeEvent(XmlFilename, PptCheckedListBox.Items)
            For i = 0 To PptCheckedListBox.Items.Count - 1
                PptCheckedListBox.SetItemCheckState(i, CheckState.Checked)
            Next
            Update()
        End If
    End Sub

    Private Sub OK_Button_Click() Handles OK_Button.Click
        If IsComplete Then
            DialogResult = DialogResult.OK
            Close()
        Else
            Dim cnt As Integer = 0
            If XmlFilename = "" Then
                Dim msgbxstatus1 As MsgBoxResult = MsgBox("Error XLS file not entered", MsgBoxStyle.RetryCancel)
                If msgbxstatus1 = MsgBoxResult.Cancel Then
                    DialogResult = DialogResult.Cancel
                    Close()
                End If
            Else
                OK_Button.Enabled = False
                WorkingLabel.Visible = True
                WorkingLabel.Text = $"Generating {PptCheckedListBox.CheckedItems.Count} PDFs..."
                Update()
                RaiseEvent GenPdfEvent(PptCheckedListBox.CheckedItems, cnt)
                OK_Button.Enabled = True
                OK_Button.Text = My.Resources.OKString
                Cancel_Button.Enabled = False
                Cancel_Button.Visible = False
                WorkingLabel.Text = $"Generated {cnt} PDFs" + vbCrLf + vbCrLf + "Click OK to continue"
                Update()
                IsComplete = True
            End If
        End If
    End Sub
    Private Sub Cancel_Button_Click() Handles Cancel_Button.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub
End Class