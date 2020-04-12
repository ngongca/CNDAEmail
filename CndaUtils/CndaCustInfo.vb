Imports System.Xml.Serialization
''' <summary>
''' Stores Cnda and email information for a single customer
''' </summary>
Public Class CndaCustInfo
    <XmlAttribute(AttributeName:="name")>
    Public Property CustName As String
    <XmlAttribute(AttributeName:="cnda")>
    Public Property Cnda As String
    <XmlElement(ElementName:="edit")>
    Public Property EditList As New List(Of CndaEditPair)
    <XmlElement(ElementName:="address")>
    Public Property AddrList As New List(Of CndaMailListItem)
End Class
