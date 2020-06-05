Imports Microsoft.Office.Interop
Imports Microsoft.Office.Tools.Ribbon
Imports System.IO
Public Class CndaOutlookRibbon

    ''' <summary>
    ''' Gets PPT and XLS file from user and then generates CNDA emails using existing PDF files that were generated using the NDA tools.
    ''' </summary>
    Private Sub CndaEmailButton_Click() Handles CNDAEmailButton.Click
        Dim ctl = New CndaOutlookEmailController()
        ctl.RunAttachEmail()
    End Sub


    ''' <summary>
    ''' A click on this button generates emails only based on Cnda Info without attachments
    ''' </summary>
    Private Sub CndaEmailOnlyButton_Click() Handles CNDAEmailOnlyButton.Click
        Dim ctl = New CndaOutlookEmailController()
        ctl.RunEmailOnly()
    End Sub

End Class
