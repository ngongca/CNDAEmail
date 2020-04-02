Imports Microsoft.Office.Tools.Ribbon

Public Class CNDAPowerPointRibbon
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub GeneratePDFButton_Click(sender As Object, e As RibbonControlEventArgs) Handles GeneratePDFButton.Click
        If PptOpenXlsFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim xlsFileName As String = PptOpenXlsFileDialog.FileName
            Dim xlsCndaInfo As New CndaAllInfo()
            xlsCndaInfo = CndaExcel.ExtractCndaInfo(xlsFilename:=xlsFileName)
            Dim pptApp As PowerPoint.Application = Globals.ThisAddIn.Application
            Dim pptPres As PowerPoint.Presentation = pptApp.ActivePresentation
            Dim pdfCnt As Integer = CndaPPTUtils.PptToPDFs(PptPres:=pptPres, CndaData:=xlsCndaInfo)
            MsgBox($"PDF Generation completed writing {pdfCnt} files")
        End If
    End Sub

    Private Sub PptSettingsButton_Click(sender As Object, e As RibbonControlEventArgs) Handles PptSettingsButton.Click
        Dim settingsDialog As New SettingsDialog()
        If settingsDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            My.Settings.Save()
            MsgBox($"here is the custname setting {My.Settings.CNDACustMatch}")
        End If
    End Sub
End Class
