Imports System.Windows.Forms
Public Class CndaPptGenController

    Private WithEvents PptView As CndaPptGenView
    Private pptModel As CndaModel

    Public Sub Run()
        pptModel = New CndaModel(My.Settings.PptXmlFilename)
        PptView = New CndaPptGenView()
        PptView.ShowDialog()
    End Sub

    Private Sub GenPdfEventHandler(ByRef objList As CheckedListBox.CheckedItemCollection,
                                  ByRef Count As Integer) Handles PptView.GenPdfEvent
        Dim custList As New List(Of CndaCustInfo)
        For Each o As CndaCustInfo In objList
            custList.Add(o)
        Next
        Count = CNDAPowerPoint.PptToPDFs(PptPres:=Globals.ThisAddIn.Application.ActivePresentation, CustList:=custList)
    End Sub

    Private Sub PptViewEvents_XmlFileChangeEvent(xmlFilename As String,
                                                 ByRef objList As CheckedListBox.ObjectCollection) Handles PptView.PptXmlFileChangeEvent
        If xmlFilename <> "" Then
            pptModel.UpdateModel(XmlFilename:=xmlFilename)
            My.Settings.PptXmlFilename = xmlFilename
            My.Settings.Save()
            objList.Clear()
            For Each o As Object In pptModel.CustInfoList
                objList.Add(o)
            Next
        End If
    End Sub
End Class
