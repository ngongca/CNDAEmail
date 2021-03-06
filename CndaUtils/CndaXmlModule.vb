﻿Public Module CndaXmlModule
    ''' <summary>
    ''' Module to generate a test XML file using the latest base classes.
    ''' </summary>
    Public Sub GenTestXML()
        Dim e1 As New CndaEditPair
        With e1
            .FindRegExPattern = "CNDA#+"
            .ReplaceValue = "11111"
        End With
        Dim e2 As New CndaEditPair
        With e2
            .FindRegExPattern = "CustName"
            .ReplaceValue = "Customer 1"
        End With
        Dim a1 As New CndaMailListItem
        With a1
            .Address = "to1@c1.com"
            .AddressType = CndaMailListItem.AddressTypeEnum.MailTo
        End With
        Dim a2 As New CndaMailListItem
        With a2
            .Address = "to2@c1.com"
            .AddressType = CndaMailListItem.AddressTypeEnum.MailTo
        End With
        Dim a3 As New CndaMailListItem
        With a3
            .Address = "cc1@c1.com"
            .AddressType = CndaMailListItem.AddressTypeEnum.MailCC
        End With
        Dim t As New CndaCustInfo
        With t
            .EditList.Add(e1)
            .EditList.Add(e2)
            .AddrList.Add(a1)
            .AddrList.Add(a2)
            .AddrList.Add(a3)
            .Cnda = "11111"
            .CustName = "Customer 1"
        End With
        Dim e3 As New CndaEditPair
        With e3
            .FindRegExPattern = "CNDA#+"
            .ReplaceValue = "22222"
        End With
        Dim e4 As New CndaEditPair
        With e4
            .FindRegExPattern = "CustName"
            .ReplaceValue = "Customer 2"
        End With
        Dim a4 As New CndaMailListItem
        With a4
            .Address = "to1@c2.com"
            .AddressType = CndaMailListItem.AddressTypeEnum.MailTo
        End With
        Dim a5 As New CndaMailListItem
        With a5
            .Address = "cc1@c2.com"
            .AddressType = CndaMailListItem.AddressTypeEnum.MailCC
        End With
        Dim a6 As New CndaMailListItem
        With a6
            .Address = "bcc1@c2.com"
            .AddressType = CndaMailListItem.AddressTypeEnum.MailBCC
        End With
        Dim t2 As New CndaCustInfo
        With t2
            .EditList.Add(e3)
            .EditList.Add(e4)
            .AddrList.Add(a4)
            .AddrList.Add(a5)
            .AddrList.Add(a6)
            .Cnda = "22222"
            .CustName = "Customer 2"
        End With
        Dim ci As New List(Of CndaCustInfo) From {
            t,
            t2
        }
        Dim tfile As String = System.IO.Path.GetTempFileName()
        MsgBox($"Writing xml {tfile}")
        CndaCustInfoToXml(tfile, ci)
    End Sub
    Public Sub CndaCustInfoToXml(XmlFilename As String, Info As List(Of CndaCustInfo))
        Dim mySerializer As New Xml.Serialization.XmlSerializer(GetType(List(Of CndaCustInfo)))
        ' To write to a file, create a StreamWriter object. 
        MsgBox($"Writing file {XmlFilename}")
        Dim myWriter As New IO.StreamWriter(XmlFilename)
        mySerializer.Serialize(myWriter, Info)
        myWriter.Close()
    End Sub
    Public Function CndaXmlToCustInfo(XmlFileName As String) As List(Of CndaCustInfo)
        Dim reader As New Xml.Serialization.XmlSerializer(GetType(List(Of CndaCustInfo)))
        Dim file As New IO.StreamReader(XmlFileName)
        CndaXmlToCustInfo = CType(reader.Deserialize(file), List(Of CndaCustInfo))
    End Function
End Module
