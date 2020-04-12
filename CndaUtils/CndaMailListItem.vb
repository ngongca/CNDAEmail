Imports System.Xml.Serialization
<XmlType(TypeName:="email")>
Public Class CndaMailListItem
    Public Enum AddressTypeEnum
        MailTo
        MailCC
        MailBCC
    End Enum
    <XmlAttribute(AttributeName:="type")>
    Public Property AddressType As AddressTypeEnum
    <XmlText>
    Public Property Address As String
End Class
