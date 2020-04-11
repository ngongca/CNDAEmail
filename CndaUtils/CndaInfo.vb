Imports System.Xml.Serialization
''' <summary>
''' Stores Cnda and email information for a single customer
''' </summary>
<XmlType(TypeName:="CustInfo")>
Public Class CndaInfo
    Public Property EditList As New List(Of CndaEditPair)
    Public Property Tolist As New List(Of String)
    Public Property CcList As New List(Of String)
    Public Property BccList As New List(Of String)
    Public Property Cnda As String
    Public Property CustName As String
End Class
