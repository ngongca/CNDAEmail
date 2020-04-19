
Public Class CndaPptModel
    Public ReadOnly Property CustInfoList As New List(Of CndaCustInfo)
    Public Property XmlFilename As String

    Public Sub New()
        If My.Settings.PptXmlFilename <> "" Then
            XmlFilename = My.Settings.PptXmlFilename
            If XmlFilename <> "" Then
                CustInfoList.Clear()
                For Each c As CndaCustInfo In CndaXmlToCustInfo(XmlFilename)
                    CustInfoList.Add(c)
                Next
            End If
        End If
    End Sub

    Public Sub UpdateModel(ByVal XmlFilename As String)
        Me.XmlFilename = XmlFilename
        My.Settings.PptXmlFilename = XmlFilename
        CustInfoList.Clear()
        For Each c As CndaCustInfo In CndaXmlToCustInfo(XmlFileName:=Me.XmlFilename)
            CustInfoList.Add(c)
        Next
    End Sub
End Class
