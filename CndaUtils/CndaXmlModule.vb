Public Module CndaXmlModule
    Public Function CndaXmlToAllInfo(XmlFileName As String) As CndaAllInfo
        Dim reader As New System.Xml.Serialization.XmlSerializer(GetType(CndaAllInfo))
        Dim file As New System.IO.StreamReader(XmlFileName)
        CndaXmlToAllInfo = CType(reader.Deserialize(file), CndaAllInfo)
    End Function
End Module
