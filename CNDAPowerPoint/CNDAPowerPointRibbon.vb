Imports Microsoft.Office.Tools.Ribbon

Public Class CNDAPowerPointRibbon

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub GeneratePDFButton_Click(sender As Object, e As RibbonControlEventArgs) Handles GeneratePDFButton.Click
        Dim PptController As New CndaPptGenController()
        PptController.Run()
    End Sub


End Class
