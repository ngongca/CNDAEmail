Imports Microsoft.Office.Interop
Imports Microsoft.Office.Tools.Ribbon
Imports System.IO
Public Class CndaOutlookRibbon
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
    End Sub

    Private Sub CndaEmailExportAndEmail_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles CNDAExportAndEmail_Button.Click
        Dim cntrl = New CndaOutlookEmailController()
        cntrl.RunExportAndEmail()
    End Sub

    ''' <summary>
    ''' Gets PPT and XLS file from user and then generates CNDA emails using existing PDF files that were generated using the NDA tools.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CndaEmailButton_Click(sender As Object, e As RibbonControlEventArgs) Handles CNDAEmailButton.Click
        Dim ctl = New CndaOutlookEmailController()
        ctl.RunAttacheEmail()
    End Sub


    ''' <summary>
    ''' A click on this button generates emails only based on Cnda Info without attachments
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CndaEmailOnlyButton_Click(sender As Object, e As RibbonControlEventArgs) Handles CNDAEmailOnlyButton.Click
        Dim ctl = New CndaOutlookEmailController()
        ctl.RunEmailOnly()
    End Sub

End Class
