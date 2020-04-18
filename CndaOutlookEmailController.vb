Imports Microsoft.Office.Interop.Outlook
Imports System.Windows.Forms
Public Class CndaOutlookEmailController
    Private OtlEmailView As CndaOutlookEmailView
    Private ReadOnly mdl As CndaOutlookModel
    Private WithEvents OtlEmailEvents As ICndaOutlookEvents

    Public Sub New()
        mdl = New CndaOutlookModel()
    End Sub

    Public Sub RunEmailOnly()
        OtlEmailView = New CndaOutlookEmailView With {
            .XmlFilename = mdl.XmlFileName,
            .MailFolderName = mdl.EmailFolder.Name
        }
        OtlEmailEvents = OtlEmailView
        If OtlEmailView.ShowDialog = System.Windows.Forms.DialogResult.Yes Then
            mdl.CurEmail.Close(OlInspectorClose.olDiscard)
        End If
    End Sub

    Public Sub RunAttacheEmail()

    End Sub

    Public Sub RunExportAndEmail()
        '  GenPdf = New CndaOutlookPptView With {
        '    .GeneratePdf = True
        '}
        '  GenPdf.PptFileInstructionLabel.Text = "CNDA Outlook Generate PDF"
        '  GenPdf.ShowDialog()
        '  If MsgBox("Email generation complete" & vbCrLf & "Do you with to remove the current email?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
        '      thisEmail.Close(Outlook.OlInspectorClose.olDiscard)
        '  End If
    End Sub
    'Private Sub GenPdfEventHandler(pptFilename As String, ByRef obj As List(Of CndaCustInfo)) Handles GenPdf.GeneratePdfEvent
    '    CNDAPowerPoint.PptToPDFs(pptFilename, obj)
    'End Sub
    'Private Sub GenEmailEventHandler(pptFilename As String, ByRef mailCnt As Integer) Handles GenPdf.GenerateEmailEvent
    '    'mailCnt = 0
    '    'If thisEmail IsNot Nothing Then
    '    '    For Each c As CndaCustInfo In mdl.CustInfoList
    '    '        Dim pdfFileName As String = CNDAPowerPoint.CndaPdfString(PptFilename:=pptFilename, c.Cnda, c.CustName)
    '    '        If File.Exists(pdfFileName) Then
    '    '            mdl.CreateEmailWithAttachment(pdfFileName, c, thisEmail)
    '    '            mailCnt += 1
    '    '        Else
    '    '            If MsgBox($"{$"Could not find pdf file {pdfFileName}, no email generated"}{vbCrLf}Continue?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
    '    '                Exit For
    '    '            End If
    '    '        End If
    '    '    Next
    '    'End If
    'End Sub

    Private Sub SendEmailsEventHandler(ByRef objList As CheckedListBox.CheckedItemCollection,
                                       ByRef count As Integer,
                                       ByVal pptFilename As String,
                                       ByVal genPdf As Boolean) Handles OtlEmailEvents.SendEmailsEvent
        count = 0
        If objList IsNot Nothing Then
            For Each obj As CndaCustInfo In objList
                Dim pdfFilename As String = ""
                If pptFilename <> "" Then
                    pdfFilename = CNDAPowerPoint.CndaPdfString(pptFilename, obj.Cnda, obj.CustName)
                    If genPdf Then
                        'TODO generate the pdf file
                    End If
                End If
                'Send the mail
                CreateEmailWithAttachment(pdfFilename, obj, mdl.CurEmail, mdl.EmailFolder)
                count += 1
            Next obj
        End If
    End Sub

    Private Sub EmailFolderChangeEventHandler(ByRef folder As Outlook.Folder) Handles OtlEmailEvents.EmailFolderChangeEvent
        If folder IsNot Nothing Then
            mdl.EmailFolder = folder
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
End Class
