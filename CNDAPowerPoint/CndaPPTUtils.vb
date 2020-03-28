Imports Microsoft.Office.Core

Public Module CndaPPTUtils
    ''' <summary>
    ''' Generates a PDF file from an editied
    ''' </summary>
    ''' <param name="PptFilename"></param>
    ''' <param name="CndaData"></param>
    ''' <returns></returns>
    Public Function PptToPDFs(PptFilename As String, CndaData As CndaAllInfo) As Integer
        Dim retVal As Integer = 0
        Dim pptApp As New PowerPoint.Application
        Dim pptPres As PowerPoint.Presentation = pptApp.Presentations.Open(PptFilename, WithWindow:=MsoTriState.msoFalse,
                                                                           ReadOnly:=MsoTriState.msoTrue)
        If pptPres IsNot Nothing Then
            For Each c As CndaInfo In CndaData.CndaInfos
                Dim cnda As String = c.Cnda
                Dim name As String = c.CustName
                Dim CndaXXX As String = FindRegExp(pptPres, "CNDA#+")
                FindReplaceAll(pptPres, CndaXXX, cnda)
                FindReplaceAll(pptPres, "CustName", name)

                Dim fullName As String = CndaPdfString(PptFilename, cnda, name)
                pptPres.ExportAsFixedFormat(Path:=fullName,
                                        FixedFormatType:=PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                        Intent:=PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen)
                FindReplaceAll(pptPres, cnda, CndaXXX)
                FindReplaceAll(pptPres, name, "CustName")
                retVal += 1
            Next
            pptPres.Close()
            pptPres = Nothing
            pptApp.Quit()
            pptApp = Nothing
        End If
        Return retVal
    End Function

    Public Function CndaPdfString(PptFilename As String, Cnda As String, CustName As String) As String
        'Write out pdf
        Dim wPath As String = CreateObject("Scripting.FileSystemObject").GetParentFolderName(PptFilename)
        Dim wName As String = CreateObject("Scripting.FileSystemObject").GetBaseName(PptFilename)
        Return wPath & "\" & wName & "_" & CustName & "_CNDA" & Cnda & ".pdf"
    End Function

    Public Function PptToPDF(PPTFilename As String, Name As String, Cnda As String) As String
        Dim pptApp As New PowerPoint.Application
        Dim pptPres As PowerPoint.Presentation = pptApp.Presentations.Open(PPTFilename, WithWindow:=MsoTriState.msoFalse)

        Dim CndaXXX As String = FindRegExp(pptPres, "CNDA#+")
        FindReplaceAll(pptPres, CndaXXX, Cnda)
        FindReplaceAll(pptPres, "CustName", Name)

        'Write out pdf
        Dim wPath As String = CreateObject("Scripting.FileSystemObject").GetParentFolderName(PPTFilename)
        Dim wName As String = CreateObject("Scripting.FileSystemObject").GetBaseName(PPTFilename)
        Dim fullName As String = wPath & "\" & wName & "_" & Name & "_CNDA" & Cnda & ".pdf"
        pptPres.ExportAsFixedFormat(Path:=fullName,
                                    FixedFormatType:=PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                    Intent:=PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen)
        PptToPDF = fullName
        pptPres.Close()
        pptPres = Nothing
        pptApp.Quit()
        pptApp = Nothing
    End Function
    Private Sub FindReplaceAll(ByVal pres As PowerPoint.Presentation, FindWord As String, ReplaceWord As String)
        Dim sld As PowerPoint.Slide
        Dim shp As PowerPoint.Shape
        Dim ShpTxt As PowerPoint.TextRange
        Dim TmpTxt As PowerPoint.TextRange

        'Loop through each slide in Presentation
        For Each sld In pres.Slides
            For Each shp In sld.Shapes
                If shp.HasTextFrame Then
                    'Store text into a variable
                    ShpTxt = shp.TextFrame.TextRange
                    'Find First Instance of "Find" word (if exists)
                    TmpTxt = ShpTxt.Replace(FindWhat:=FindWord, ReplaceWhat:=ReplaceWord)
                End If
            Next shp
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
                    'Store text into a variable
                    Dim ShpTxt As String = shp.TextFrame.TextRange.Text
                    ' Create a regular expression object.
                    Dim objRegExp As New RegularExpressions.Regex(myPattern, RegularExpressions.RegexOptions.IgnoreCase)
                    'Create objects.
                    'Test whether the String can be compared.
                    Dim colMatches As RegularExpressions.Match = objRegExp.Match(ShpTxt)
                    If (colMatches.Success) Then
                        FindRegExp = colMatches.Value
                        Exit For
                    End If
                End If
            Next shp
        Next sld
    End Function
End Module
