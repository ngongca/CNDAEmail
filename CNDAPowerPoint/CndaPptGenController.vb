Imports System.Windows.Forms
Public Class CndaPptGenController

    Private WithEvents PptView As CndaPptGenView
    Private pptModel As CndaPptModel

    Public Sub Run()
        pptModel = New CndaPptModel()
        PptView = New CndaPptGenView With {
            .XmlFilename = pptModel.XmlFilename
        }
        PptView.ShowDialog()
    End Sub

    Private Sub GenPdfEventHandler(ByRef objList As CheckedListBox.CheckedItemCollection,
                                  ByRef Count As Integer) Handles PptView.GenPdfEvent
        Dim custList As New List(Of CndaCustInfo)
        For Each o As CndaCustInfo In objList
            custList.Add(o)
        Next
        Count = PptToPDFs(PptPres:=Globals.ThisAddIn.Application.ActivePresentation, CustList:=custList)
    End Sub

    Private Sub PptViewEvents_XmlFileChangeEvent(xmlFilename As String,
                                                 ByRef objList As CheckedListBox.ObjectCollection) Handles PptView.PptXmlFileChangeEvent
        If xmlFilename <> "" Then
            pptModel.UpdateModel(XmlFilename:=xmlFilename)
            objList.Clear()
            For Each o As Object In pptModel.CustInfoList
                objList.Add(o)
            Next
        End If
    End Sub
End Class
