Imports Microsoft.Office.Tools.Ribbon

Public Class CNDAPowerPointRibbon
    Private WithEvents InfoDialog As CndaInfoForm
    Private xmlFilename As String
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub GeneratePDFButton_Click(sender As Object, e As RibbonControlEventArgs) Handles GeneratePDFButton.Click
        PptOpenXMLFileDialog.Title = "Select CNDA data file XML"
        If PptOpenXMLFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            xmlFilename = PptOpenXMLFileDialog.FileName
            InfoDialog = New CndaInfoForm
            With InfoDialog
                .Text = "CNDA PDFs"
                .StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
                .ShowDialog()
            End With
        End If
    End Sub

    Private Sub ReadyToBuildPdfs(ByRef count As Integer) Handles infoDialog.ReadyToBuild
        Dim xmlCndaInfo As CndaAllInfo = CndaXmlToAllInfo(XmlFileName:=xmlFilename)
        Dim pptApp As PowerPoint.Application = Globals.ThisAddIn.Application
        Dim pptPres As PowerPoint.Presentation = pptApp.ActivePresentation
        count = PptToPDFs(PptPres:=pptPres, CndaData:=xmlCndaInfo)
    End Sub

    Private Sub PptSettingsButton_Click(sender As Object, e As RibbonControlEventArgs) Handles PptSettingsButton.Click
        Dim settingsDialog As New SettingsDialog()
        If settingsDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            My.Settings.Save()
            MsgBox($"here is the custname setting {My.Settings.CNDACustMatch}")
        End If
    End Sub
End Class
