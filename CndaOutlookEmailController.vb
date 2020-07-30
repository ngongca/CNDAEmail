Imports Microsoft.Office.Interop.Outlook
Imports System.Windows.Forms
Public Class CndaOutlookEmailController
    Private ReadOnly mdl As CndaOutlookModel
    Private WithEvents OtlEmailEvents As ICndaOutlookEvents

    Public Sub New()
        mdl = New CndaOutlookModel()
    End Sub

    Public Sub RunEmailOnly()
        Dim OtlEmailView As New CndaOutlookEmailView With {
            .XmlFilename = mdl.XmlFileName
        }
        With mdl
            .AttachPdf = False
        End With
        OtlEmailEvents = OtlEmailView
        If OtlEmailView.ShowDialog = System.Windows.Forms.DialogResult.Yes Then
            mdl.CurEmail.Close(OlInspectorClose.olDiscard)
        End If
        OtlEmailView.Dispose()
    End Sub

    Public Sub RunAttachEmail()
        Dim OtlPptEmailView As New CndaOtlPptEmailView With {
            .XmlFilename = mdl.XmlFileName
        }
        With mdl
            .AttachPdf = True
        End With
        OtlEmailEvents = OtlPptEmailView
        If OtlPptEmailView.ShowDialog = DialogResult.Yes Then
            mdl.CurEmail.Close(OlInspectorClose.olDiscard)
        End If
        OtlPptEmailView.Dispose()
    End Sub


    Private Sub SendEmailsEventHandler(ByRef objList As CheckedListBox.CheckedItemCollection,
                                       ByRef count As Integer) Handles OtlEmailEvents.SendEmailsEvent
        count = 0
        If objList IsNot Nothing Then
            For Each obj As CndaCustInfo In objList
                Dim pdfFilename As String = ""
                If mdl.PptFileName <> "" Then
                    pdfFilename = CNDAPowerPoint.CndaPdfString(mdl.PptFileName, obj.Cnda, obj.CustName)
                End If
                'Send the mail
                CreateEmail(pdfFilename, obj, mdl.CurEmail, mdl.AttachPdf)
                count += 1
            Next obj
        End If
    End Sub
    Private Sub XmlFileChangeEventHander(ByVal xmlFilename As String,
                                         ByRef objList As CheckedListBox.ObjectCollection) Handles OtlEmailEvents.XmlFileChangeEvent
        If xmlFilename <> "" Then
            mdl.UpdateModel(xmlFilename:=xmlFilename)
            objList.Clear()
            For Each o As Object In mdl.CustInfoList
                objList.Add(o)
            Next
        End If
    End Sub
    Private Sub PptFileChangeEventHandler(ByVal pptFilename As String) Handles OtlEmailEvents.PptFileChangeEvent
        If pptFilename <> "" Then
            mdl.PptFileName = pptFilename
        End If
    End Sub
End Class
