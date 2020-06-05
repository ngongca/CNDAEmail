Imports Microsoft.Office.Core
Imports System.Text.RegularExpressions
Imports System.IO

''' <summary>
''' Cnda utilities that work on PowerPoint files
''' </summary>
Public Module CndaPPTUtils
    ''' <summary>
    ''' Generate single PDF for customer
    ''' </summary>
    ''' <param name="PptFilename">File location of Presentation to export to PDF</param>
    ''' <param name="Cust"><see cref="CndaCustInfo"/> containing edit streams and customer NDA info</param>
    Public Sub PptToPDF(PptFilename As String, Cust As CndaCustInfo)
        Dim pptApp As New PowerPoint.Application
        Dim pptPres As PowerPoint.Presentation = pptApp.Presentations.Open(PptFilename, WithWindow:=MsoTriState.msoFalse,
                                                                               ReadOnly:=MsoTriState.msoTrue)
        If pptPres IsNot Nothing And Cust IsNot Nothing Then
            FindReplaceAll(pptPres, Cust)
            Dim fullName As String = CndaPdfString(pptPres.FullName, Cust.Cnda, Cust.CustName)
            pptPres.ExportAsFixedFormat(Path:=fullName,
                                            FixedFormatType:=PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                            Intent:=PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen)
            pptPres.Close()
        End If
    End Sub
    ''' <summary>
    ''' Generate multiple PDFs for customers in <paramref name="CustList"/>
    ''' </summary>
    ''' <param name="PptFilename">File path to Presentation to export to PDF</param>
    ''' <param name="CustList">List of <see cref="CndaCustInfo"/> containing edit streams and customer NDA info</param>
    ''' <returns></returns>
    Public Function PptToPDFs(PptFilename As String, CustList As List(Of CndaCustInfo)) As Integer
        Dim retVal As Integer = 0
        Dim pptApp As New PowerPoint.Application
        Dim pptPres As PowerPoint.Presentation = pptApp.Presentations.Open(PptFilename, WithWindow:=MsoTriState.msoFalse,
                                                                           ReadOnly:=MsoTriState.msoTrue)
        If pptPres IsNot Nothing And CustList IsNot Nothing Then
            Dim tempfile As String = System.IO.Path.GetTempFileName()
            pptPres.SaveCopyAs(tempfile)
            For Each c As CndaCustInfo In CustList
                Dim tpres As PowerPoint.Presentation = pptApp.Presentations.Open(tempfile, [ReadOnly]:=MsoTriState.msoTrue,
                                                                                 WithWindow:=MsoTriState.msoFalse)
                FindReplaceAll(tpres, c)
                Dim fullName As String = CndaPdfString(pptPres.FullName, c.Cnda, c.CustName)
                tpres.ExportAsFixedFormat(Path:=fullName,
                                        FixedFormatType:=PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                        Intent:=PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen)
                tpres.Close()
                retVal += 1
            Next
            If File.Exists(tempfile) Then
                File.Delete(tempfile)
            End If
        End If
        Return retVal
    End Function

    ''' <summary>
    ''' Generate a standard PDF filename from a PowerPoint filename
    ''' </summary>
    ''' <param name="PptFilename">PowerPoint filename</param>
    ''' <param name="Cnda"></param>
    ''' <param name="CustName"></param>
    ''' <returns>Fully qualified PDF file name with path</returns>
    Public Function CndaPdfString(PptFilename As String, Cnda As String, CustName As String) As String
        'Write out pdf
        Dim wBase As String = Path.ChangeExtension(PptFilename, vbNullString)
        Return wBase & "_" & CustName & "_CNDA" & Cnda & ".pdf"
    End Function

    ''' <summary>
    ''' Replace all text in the presentation per the <see cref="CndaCustInfo.EditList"/>.
    ''' </summary>
    ''' <param name="pres">Presentation to make changes</param>
    ''' <param name="Info">Customer list containing edit information in EditList</param>
    Public Sub FindReplaceAll(ByRef pres As PowerPoint.Presentation, Info As CndaCustInfo)
        Dim sld As PowerPoint.Slide
        Dim shp As PowerPoint.Shape

        If pres IsNot Nothing And Info IsNot Nothing Then
            For Each sld In pres.Slides
                For Each shp In sld.Shapes
                    If shp.HasTextFrame Then
                        If shp.TextFrame.HasText Then
                            For Each pair In Info.EditList
                                Dim m As Match = Regex.Match(shp.TextFrame.TextRange.Text, pair.FindRegExPattern, RegexOptions.IgnoreCase)
                                If m.Success Then
                                    shp.TextFrame.TextRange.Replace(m.Value, pair.ReplaceValue)
                                End If
                            Next pair
                        End If
                    End If
                Next shp
                ' Check footer as well
                If sld.HeadersFooters.Footer.Visible Then
                    For Each pair In Info.EditList
                        sld.HeadersFooters.Footer.Text = Regex.Replace(sld.HeadersFooters.Footer.Text,
                                                                       pair.FindRegExPattern,
                                                                       pair.ReplaceValue, RegexOptions.IgnoreCase)
                    Next pair
                End If
            Next sld
        End If
    End Sub

End Module
