Imports System.Xml.Serialization
''' <summary>
''' Contains a collection of <see cref="CndaCustInfo"/> customer data objects
''' </summary>
Public Class CndaAllInfo
    ''' <summary>
    ''' Collection of <see cref="CndaCustInfo"/> objects
    ''' </summary>
    ''' <returns></returns>
    <XmlElement(ElementName:="customer")>
    Public Property CndaInfos As New List(Of CndaCustInfo)
End Class
