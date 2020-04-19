Imports System.IO
Public Class CndaOutlookModel
    Public ReadOnly Property CustInfoList As New List(Of CndaCustInfo)


    Private _EmailFolder As Outlook.Folder
    Public Property EmailFolder() As Outlook.Folder
        Get
            Return _EmailFolder
        End Get
        Set(AutoPropertyValue As Outlook.Folder)
            If AutoPropertyValue IsNot Nothing Then
                _EmailFolder = AutoPropertyValue
                My.Settings.MailFolderId = _EmailFolder.EntryID
                My.Settings.Save()
            End If
        End Set
    End Property
    Public Property PptFileName As String = ""
    Public Property XmlFileName As String = "<enter xml file>"
    Public Property CurEmail As Outlook.MailItem
    Public Property GenPdf As Boolean = False
    Public Property AttachPdf As Boolean = False


    Public Sub New()

        'My.Settings.Reset()
        'My.Settings.Save()
        If My.Settings.XmlFileName <> "" Then
            CustInfoList = CndaXmlToCustInfo(My.Settings.XmlFileName)
            XmlFileName = My.Settings.XmlFileName
        End If
        'set up this email

        Dim selObject As Object = Globals.ThisAddIn.Application.ActiveInspector.CurrentItem
        If (TypeOf selObject Is Outlook.MailItem) Then
            CurEmail = TryCast(selObject, Outlook.MailItem)
        End If


        'set default folder
        If My.Settings.MailFolderId Is "" Then
            EmailFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
            My.Settings.MailFolderId = EmailFolder.EntryID
            My.Settings.Save()
        Else
            EmailFolder = Globals.ThisAddIn.Application.Session.GetFolderFromID(My.Settings.MailFolderId)
        End If
    End Sub


    Public Sub UpdateModel(ByVal xmlFilename As String)
        My.Settings.XmlFileName = xmlFilename
        My.Settings.Save()
        CustInfoList.Clear()
        For Each cust As CndaCustInfo In CndaXmlToCustInfo(XmlFileName:=xmlFilename)
            CustInfoList.Add(cust)
        Next
    End Sub



End Class
