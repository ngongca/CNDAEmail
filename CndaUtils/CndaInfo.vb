''' <summary>
''' Stores Cnda and email information for a single customer
''' </summary>
Public Class CndaInfo
    Public Property CustName As String
    Public Property ToList As New List(Of String)
    Public Property Cnda As String
    Public Property CcList As New List(Of String)
    Public Property BccList As New List(Of String)
End Class
