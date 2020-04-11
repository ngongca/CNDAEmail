''' <summary>
''' Cnda utilities that require access to Excel spreadsheets
''' </summary>
Public Module CndaExcelUtils
    'TODO remove hard code of columns
    Private Const TO_COL As String = "C2:C50"
    Private Const CC_COL As String = "D2:D50"
    Private Const BCC_COL As String = "E2:E50"
    Private Const NAME_CELL As String = "A2"
    Private Const CNDA_CELL As String = "B2"
    ''' <summary>
    ''' Parses Excel datasheet for cnda and email information
    ''' </summary>
    ''' <param name="xlsFilename">Fully qualified Excel filename to parse</param>
    ''' <returns>fully populated <c cref="CndaBaseClasses.CndaAllInfo">object</c></returns>
    Public Function ExtractCndaInfo(xlsFilename As String) As CndaAllInfo
        Dim xlApp As New Excel.Application
        Dim xlWb As Excel.Workbook = xlApp.Workbooks.Open(Filename:=xlsFilename, [ReadOnly]:=True)
        Dim xlAllInfo As CndaAllInfo = Nothing
        If xlWb IsNot Nothing Then
            Dim xlWs As Excel.Worksheet
            xlAllInfo = New CndaAllInfo()
            For Each xlWs In xlWb.Sheets
                Dim xlInfo As New CndaInfo()
                With xlInfo
                    .CustName = xlWs.Range(NAME_CELL).Text
                    .Cnda = xlWs.Range(CNDA_CELL).Text
                End With
                For Each c As Excel.Range In xlWs.Range(TO_COL)
                    If c.Text <> "" And c.Row <> 1 Then
                        xlInfo.Tolist.Add(c.Text)
                    ElseIf c.Text = "" Then
                        Exit For
                    End If
                Next
                For Each c As Excel.Range In xlWs.Range(CC_COL)
                    If c.Text <> "" And c.Row <> 1 Then
                        xlInfo.CcList.Add(c.Text)
                    ElseIf c.Text = "" Then
                        Exit For
                    End If
                Next
                For Each c As Excel.Range In xlWs.Range(BCC_COL)
                    If c.Text <> "" And c.Row <> 1 Then
                        xlInfo.BccList.Add(c.Text)
                    ElseIf c.Text = "" Then
                        Exit For
                    End If
                Next
                xlAllInfo.CndaInfos.Add(xlInfo)
            Next
            xlWb.Close()
            xlWb = Nothing
        Else
            MsgBox($"Error - could not open {xlsFilename}", MsgBoxStyle.Exclamation)
        End If
        xlApp.Quit()
        xlApp = Nothing
        Return xlAllInfo
    End Function

End Module
