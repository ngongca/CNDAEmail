Imports System.IO
Public Class CndaOutlookModel
    Public Property CustInfoList As New List(Of CndaCustInfo)

    Public Sub InitModel()
        My.Settings.Reset()
        My.Settings.Save()
        If My.Settings.XmlFileName <> "" Then
            '    CustInfoList = CndaXmlToCustInfo(My.Settings.XmlFileName)
            '_AllInfo = CndaXmlToAllInfo(My.Settings.XmlFileName)
        End If

        'set default folder
        If My.Settings.MailFolderId Is "" Then
            Dim df As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
            My.Settings.MailFolderId = df.EntryID
            My.Settings.Save()
        End If
    End Sub

    Public Sub UpdateModel(ByVal xmlFilename As String)
        CustInfoList = CndaXmlToCustInfo(XmlFileName:=xmlFilename)
        '_AllInfo = CndaXmlToAllInfo(XmlFileName:=xmlFilename)
    End Sub


    Public Function CreateCustList(listObject As System.Windows.Forms.CheckedListBox.CheckedItemCollection) As List(Of CndaCustInfo)
        CreateCustList = New List(Of CndaCustInfo)
        For Each obj As CndaCustInfo In listObject
            CreateCustList.Add(obj)
        Next
    End Function



End Class
