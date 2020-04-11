Public Module CndaXmlModule
    Public Function CndaXmlToAllInfo(XmlFileName As String) As CndaAllInfo
        Dim reader As New Xml.Serialization.XmlSerializer(GetType(CndaAllInfo))
        Dim file As New IO.StreamReader(XmlFileName)
        CndaXmlToAllInfo = CType(reader.Deserialize(file), CndaAllInfo)
    End Function
    ''' <summary>
    ''' code to serialize the objects in case they change.
    ''' </summary>
    ''' <param name="XmlFilename"></param>
    ''' <param name="Info"></param>
    Public Sub CndaAllInfoToXml(XmlFilename As String, Info As CndaAllInfo)
        Dim mySerializer As New Xml.Serialization.XmlSerializer(GetType(CndaAllInfo))
        ' To write to a file, create a StreamWriter object. 
        MsgBox($"Writing file {XmlFilename}")
        Dim myWriter As New IO.StreamWriter(XmlFilename)
        mySerializer.Serialize(myWriter, Info)
        myWriter.Close()
    End Sub
End Module
