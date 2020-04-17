
Public Class CndaPptModel
    Public Property CustInfoList As New List(Of CndaCustInfo)
    Public Property XmlFilename As String

    Public Sub New()
        Me.XmlFilename = My.Settings.XmlFilename
        If XmlFilename <> "" Then
            CustInfoList = CndaXmlToCustInfo(XmlFilename)
        End If
    End Sub

    Public Sub UpdateModel(ByVal XmlFilename As String)
        Me.XmlFilename = XmlFilename
        CustInfoList = CndaXmlToCustInfo(XmlFileName:=Me.XmlFilename)
    End Sub


    Public Function CreateCustList(listObject As System.Windows.Forms.CheckedListBox.CheckedItemCollection) As List(Of CndaCustInfo)
        CreateCustList = New List(Of CndaCustInfo)
        For Each obj As CndaCustInfo In listObject
            CreateCustList.Add(obj)
        Next
    End Function

End Class
