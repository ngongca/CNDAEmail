Imports System.Windows.Forms

Public Class GetFileDialog
    Dim pptFilename As String = ""
    Dim xlsFilename As String = ""

    Public Function GetPptFilename() As String
        GetPptFilename = pptFilename
    End Function

    Public Function GetXlsFilename() As String
        GetXlsFilename = xlsFilename
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If pptFilename = "" Then
            Dim msgbxstatus As MsgBoxResult = MsgBox("Error PPT file not entered", MsgBoxStyle.RetryCancel)
            If msgbxstatus = MsgBoxResult.Cancel Then
                Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        ElseIf xlsFilename = "" Then
            Dim msgbxstatus1 As MsgBoxResult = MsgBox("Error XLS file not entered", MsgBoxStyle.RetryCancel)
            If msgbxstatus1 = MsgBoxResult.Cancel Then
                Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub SelectPPT_Button_Click(sender As Object, e As EventArgs) Handles SelectPPT_Button.Click
        OpenPPTFileDialog.ShowDialog()
        pptFilename = OpenPPTFileDialog.FileName
        PPT_Label.Text = pptFilename
    End Sub

    Private Sub SelectXLS_Button_Click(sender As Object, e As EventArgs) Handles SelectXLS_Button.Click
        OpenXLSFileDialog.ShowDialog()
        xlsFilename = OpenXLSFileDialog.FileName
        XLS_Label.Text = xlsFilename
    End Sub

End Class
