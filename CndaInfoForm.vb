Public Class CndaInfoForm
    Private Sub CndaInfoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Public Event ReadyToBuild(ByRef FileCount As Integer)

    Private Sub CndaInfoForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        InfoLabel.Text = "Working"
        UseWaitCursor = True
        Update()
        Dim pdfCount As Integer
        RaiseEvent ReadyToBuild(pdfCount)
        InfoLabel.Text = $"Done.  Generated {pdfCount} PDFs"
        PptDoneButton.Visible = True
        PptDoneButton.Enabled = True
        UseWaitCursor = False
        Update()
    End Sub

    Private Sub PptDoneButton_Click(sender As Object, e As EventArgs) Handles PptDoneButton.Click
        DialogResult = System.Windows.Forms.DialogResult.OK
        Close()
    End Sub
End Class