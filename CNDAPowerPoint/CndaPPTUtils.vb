Imports Microsoft.Office.Core
Imports System.Text.RegularExpressions

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
                Dim tpres As PowerPoint.Presentation = pptApp.Presentations.Open(tempfile, [ReadOnly]:=MsoTriState.msoCTrue,
                                                                                 WithWindow:=MsoTriState.msoFalse)
                FindReplaceAll(tpres, c)
                Dim fullName As String = CndaPdfString(pptPres.FullName, c.Cnda, c.CustName)
                tpres.ExportAsFixedFormat(Path:=fullName,
                                        FixedFormatType:=PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                        Intent:=PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen)
                tpres.Close()
                retVal += 1
            Next
            If System.IO.File.Exists(tempfile) Then
                System.IO.File.Delete(tempfile)
            End If
        End If
        Return retVal
    End Function
    ''' <summary>
    ''' Generate multiple PDFs for customers in <paramref name="CustList"/>
    ''' </summary>
    ''' <param name="PptPres">Presentation to export to PDF</param>
    ''' <param name="CustList">List of <see cref="CndaCustInfo"/> containing edit streams and customer NDA info</param>
    ''' <returns></returns>
    Public Function PptToPDFs(ByRef PptPres As PowerPoint.Presentation, CustList As List(Of CndaCustInfo)) As Integer
        Dim retVal As Integer = 0
        If PptPres IsNot Nothing And CustList IsNot Nothing Then
            Dim tempfile As String = System.IO.Path.GetTempFileName()
            PptPres.SaveCopyAs(tempfile)
            Dim pptApp As PowerPoint.Application = Globals.ThisAddIn.Application
            For Each c As CndaCustInfo In CustList
                Dim tpres As PowerPoint.Presentation = pptApp.Presentations.Open(tempfile, [ReadOnly]:=MsoTriState.msoCTrue,
                                                                                 WithWindow:=MsoTriState.msoFalse)
                FindReplaceAll(tpres, c)
                Dim fullName As String = CndaPdfString(PptPres.FullName, c.Cnda, c.CustName)
                tpres.ExportAsFixedFormat(Path:=fullName,
                                        FixedFormatType:=PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                        Intent:=PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen)
                tpres.Close()
                retVal += 1
            Next
            If System.IO.File.Exists(tempfile) Then
                System.IO.File.Delete(tempfile)
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
        Dim wPath As String = CreateObject("Scripting.FileSystemObject").GetParentFolderName(PptFilename)
        Dim wName As String = CreateObject("Scripting.FileSystemObject").GetBaseName(PptFilename)
        Return wPath & "\" & wName & "_" & CustName & "_CNDA" & Cnda & ".pdf"
    End Function
    Private Sub FindReplaceAll(ByRef pres As PowerPoint.Presentation, FindWord As String, ReplaceWord As String)
        Dim sld As PowerPoint.Slide
        Dim shp As PowerPoint.Shape

        For Each sld In pres.Slides
            For Each shp In sld.Shapes
                If shp.HasTextFrame Then
                    If shp.TextFrame.HasText Then
                        shp.TextFrame.TextRange.Text = Regex.Replace(shp.TextFrame.TextRange.Text, FindWord, ReplaceWord, RegexOptions.IgnoreCase)
                    End If
                End If
            Next shp
            ' Check footer as well
            If sld.HeadersFooters.Footer.Visible Then
                sld.HeadersFooters.Footer.Text = Regex.Replace(sld.HeadersFooters.Footer.Text, FindWord, ReplaceWord, RegexOptions.IgnoreCase)
            End If
        Next sld
    End Sub
    Private Sub FindReplaceAll(ByRef pres As PowerPoint.Presentation, Info As CndaCustInfo)
        Dim sld As PowerPoint.Slide
        Dim shp As PowerPoint.Shape

        For Each sld In pres.Slides
            For Each shp In sld.Shapes
                If shp.HasTextFrame Then
                    If shp.TextFrame.HasText Then
                        For Each pair In Info.EditList
                            shp.TextFrame.TextRange.Text = Regex.Replace(shp.TextFrame.TextRange.Text, pair.FindRegExPattern,
                                                                         pair.ReplaceValue, RegexOptions.IgnoreCase)
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
    End Sub
    'Find the FIRST occurance of myPattern in the powerpoint and return the value
    Private Function FindRegExp(ByVal pres As PowerPoint.Presentation, myPattern As String)
        FindRegExp = myPattern
        Dim sld As PowerPoint.Slide
        'Loop through each slide in Presentation
        For Each sld In pres.Slides
            Dim shp As PowerPoint.Shape
            For Each shp In sld.Shapes
                If shp.HasTextFrame Then
                    If shp.TextFrame.HasText Then
                        'Test whether the String can be compared.
                        Dim colMatches As Match = Regex.Match(shp.TextFrame.TextRange.Text, myPattern, RegexOptions.IgnoreCase)
                        If colMatches.Success Then
                            FindRegExp = colMatches.Value
                            Exit For
                        End If
                    End If
                End If
            Next shp
        Next sld
    End Function
End Module
