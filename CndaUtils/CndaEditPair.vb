Imports System.Xml.Serialization
''' <summary>
''' Contains an editing pair to apply to the powerpoint presentation.
''' </summary>
''' 
Public Class CndaEditPair
    <XmlAttribute(AttributeName:="key")>
    Property FindRegExPattern As String
    <XmlText>
    Property ReplaceValue As String
End Class
