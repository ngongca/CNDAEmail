﻿Imports Microsoft.Office.Core
Imports System.Text.RegularExpressions

''' <summary>
''' Cnda utilities that work on PowerPoint files
''' </summary>
Public Module CndaPPTUtils
    ''' <summary>
    ''' Generates PDF files using NDA information from the <paramref name="CndaData"/> parameter to edit a PowerPoint deck.  
    ''' For each worksheet in the workbook, a PDF file is generated.
    ''' </summary>
    ''' <param name="PptFilename">PowerPoint deck to edit</param>
    ''' <param name="CndaData">contains CNDA information</param>
    ''' <returns>number of PDF files generated</returns>
    Public Function PptToPDFs(PptFilename As String, CndaData As CndaAllInfo) As Integer
        Dim retVal As Integer = 0
        Dim pptApp As New PowerPoint.Application
        Dim pptPres As PowerPoint.Presentation = pptApp.Presentations.Open(PptFilename, WithWindow:=MsoTriState.msoFalse,
                                                                           ReadOnly:=MsoTriState.msoTrue)
        If pptPres IsNot Nothing Then
            For Each c As CndaCustInfo In CndaData.CndaInfos
                Dim cnda As String = c.Cnda
                Dim name As String = c.CustName

                Dim CndaXXX As String = FindRegExp(pptPres, My.Settings.CNDARegEx)
                FindReplaceAll(pptPres, CndaXXX, cnda)
                FindReplaceAll(pptPres, My.Settings.CNDACustMatch, name)

                Dim fullName As String = CndaPdfString(PptFilename, cnda, name)
                pptPres.ExportAsFixedFormat(Path:=fullName,
                                        FixedFormatType:=PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                        Intent:=PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen)
                FindReplaceAll(pptPres, cnda, CndaXXX)
                FindReplaceAll(pptPres, name, My.Settings.CNDACustMatch)
                retVal += 1
            Next
            pptPres.Close()
            pptPres = Nothing
            pptApp.Quit()
            pptApp = Nothing
        End If
        Return retVal
    End Function
    ''' <summary>4b
    ''' 
    ''' Generates PDF files using data from the <see cref="CndaAllInfo"/> information to edit and export from the
    ''' <see cref="PowerPoint.Presentation"/> that is provided.
    ''' </summary>
    ''' <param name="PptPres">Presentation that will be edited</param>
    ''' <param name="CndaData">contains CNDA information</param>
    ''' <returns>Number of files generated</returns>
    Public Function PptToPDFs(ByRef PptPres As PowerPoint.Presentation, CndaData As CndaAllInfo) As Integer
        Dim retVal As Integer = 0
        If PptPres IsNot Nothing Then
            Dim tempfile As String = System.IO.Path.GetTempFileName()
            PptPres.SaveCopyAs(tempfile)
            Dim pptApp As PowerPoint.Application = Globals.ThisAddIn.Application
            For Each c As CndaCustInfo In CndaData.CndaInfos
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
