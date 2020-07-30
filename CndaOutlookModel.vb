Imports System.IO
Public Class CndaOutlookModel
    Public ReadOnly Property CustInfoList As New List(Of CndaCustInfo)

    Public Property PptFileName As String = ""
    Public Property XmlFileName As String = ""
    Public Property CurEmail As Outlook.MailItem
    Public Property AttachPdf As Boolean = False

    Public Sub New()
        Dim selObject As Object = Globals.ThisAddIn.Application.ActiveInspector.CurrentItem
        If (TypeOf selObject Is Outlook.MailItem) Then
            CurEmail = TryCast(selObject, Outlook.MailItem)
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
