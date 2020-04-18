
Public Class CndaModel
    Public Property CustInfoList As New List(Of CndaCustInfo)
    Public Property XmlFilename As String

    Public Sub New(XmlFilename As String)
        Me.XmlFilename = XmlFilename
        If XmlFilename <> "" Then
            CustInfoList = CndaXmlToCustInfo(XmlFilename)
        End If
    End Sub

    Public Sub UpdateModel(ByVal XmlFilename As String)
        Me.XmlFilename = XmlFilename
        CustInfoList = CndaXmlToCustInfo(XmlFileName:=Me.XmlFilename)
    End Sub
End Class
