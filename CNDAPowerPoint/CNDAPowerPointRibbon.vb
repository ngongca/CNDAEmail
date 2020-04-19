Imports Microsoft.Office.Tools.Ribbon

Public Class CNDAPowerPointRibbon
    Private Sub GeneratePDFButton_Click() Handles GeneratePDFButton.Click
        Dim PptController As New CndaPptGenController()
        PptController.Run()
    End Sub

End Class
