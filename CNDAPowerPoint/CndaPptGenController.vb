Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Core

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
        Count = 0
        Dim tempfile As String = Path.GetTempFileName()
        Dim pptApp As PowerPoint.Application = Globals.ThisAddIn.Application
        Dim PptPres As PowerPoint.Presentation = pptApp.ActivePresentation
        PptPres.SaveCopyAs(tempfile)
        For Each c As CndaCustInfo In objList
            Dim tpres As PowerPoint.Presentation = pptApp.Presentations.Open(tempfile, [ReadOnly]:=MsoTriState.msoTrue,
                                                                             WithWindow:=MsoTriState.msoFalse)
            FindReplaceAll(tpres, c)
            Dim fullName As String = CndaPdfString(PptPres.FullName, c.Cnda, c.CustName)
            tpres.ExportAsFixedFormat(Path:=fullName,
                                    FixedFormatType:=PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                    Intent:=PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen)
            tpres.Close()
            Count += 1
        Next c
        If File.Exists(tempfile) Then
            File.Delete(tempfile)
        End If
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
