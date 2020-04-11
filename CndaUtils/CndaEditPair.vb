Imports System.Xml.Serialization
''' <summary>
''' Contains an editing pair to apply to the powerpoint presentation.
''' </summary>
''' 
<XmlType(TypeName:="Data")>
Public Class CndaEditPair
    <XmlAttribute()>
    Property FindRegExPattern As String
    <XmlText>
    Property ReplaceValue As String
End Class
