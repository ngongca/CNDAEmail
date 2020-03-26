Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports Microsoft.Office.Tools.Ribbon

Public Class CNDAExportAndEmail
    Private Const TO_COL As String = "C2:C50"
    Private Const CC_COL As String = "D2:D50"
    Private Const BCC_COL As String = "E2:E50"
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub CNDAUpdateEmail_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles CNDAUpdateEmail_Button.Click
        Dim df As New GetFileDialog
        If df.ShowDialog() = Global.System.Windows.Forms.DialogResult.OK Then
            Dim m As Outlook.Inspector = e.Control.Context
            Dim mailItem As Outlook.MailItem = TryCast(m.CurrentItem, Outlook.MailItem)
            If mailItem IsNot Nothing Then
                Dim pptFilename As String = df.GetPptFilename()
                Dim xlsFilename As String = df.GetXlsFilename()
                ExportAndEmailAll(pptFilename, xlsFilename, mailItem)
                If MsgBox("Email generation complete. See your Drafts folder." & vbCrLf & "Do you with to remove the current email?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    mailItem.Close(Outlook.OlInspectorClose.olDiscard)
                End If
            End If
        End If
    End Sub

    Private Sub ExportAndEmailAll(pptFilename As String, xlsFilename As String, ByVal mailItem As Outlook.MailItem)
        Dim xlApp As New Excel.Application
        Dim xlWb As Excel.Workbook = xlApp.Workbooks.Open(xlsFilename,, True)
        Dim xlWs As Excel.Worksheet
        For Each xlWs In xlWb.Sheets
            Dim name As String = xlWs.Range("A2").Text
            Dim cnda As String = xlWs.Range("B2").Text
            If MsgBox($"Generate email for {name} with {pptFilename}?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Dim pdfFilename As String = PowerPointToPDF(pptFilename, name, cnda)
                CreateEmailFromWorksheet(pdfFilename, xlWs, mailItem)
            End If
        Next
        xlApp.Quit()
        xlWb = Nothing
    End Sub

    Private Sub CreateEmailFromWorksheet(pdfFilename As String, xlWs As Excel.Worksheet, ByVal refMail As Outlook.MailItem)
        If (refMail IsNot Nothing) Then
            Dim curMail As Outlook.MailItem = refMail.Copy
            Dim unused = curMail.Attachments.Add(Source:=pdfFilename)
            Dim range As Excel.Range = xlWs.Range(TO_COL)
            For Each c In range
                If c.Text <> "" And c.Row <> 1 Then
                    Dim recipient1 As Outlook.Recipient = curMail.Recipients.Add(c.Text)
                    recipient1.Type = Outlook.OlMailRecipientType.olTo
                ElseIf c.Text = "" Then
                    Exit For
                End If
            Next
            range = xlWs.Range(CC_COL)
            For Each c In range
                If c.Text <> "" And c.Row <> 1 Then
                    Dim recipient As Outlook.Recipient = curMail.Recipients.Add(Name:=c.Value2)
                    recipient.Type = Outlook.OlMailRecipientType.olCC
                ElseIf c.Value = "" Then
                    Exit For
                End If
            Next
            range = xlWs.Range(BCC_COL)
            For Each c In range
                If c.Value <> "" And c.Row <> 1 Then
                    Dim recipient2 As Outlook.Recipient = curMail.Recipients.Add(c.Value)
                    recipient2.Type = Outlook.OlMailRecipientType.olBCC
                ElseIf c.Value = "" Then
                    Exit For
                End If
            Next
            Dim folder As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts)
            If folder Is Nothing Then
                MsgBox("Error cannot find Drafts folder in Outlook", MsgBoxStyle.Critical)
            Else
                curMail.Move(folder)
            End If
        End If
    End Sub

    Function PowerPointToPDF(PPTFilename As String, Name As String, Cnda As String) As String
        Dim pptApp As New PowerPoint.Application
        Dim pptPres As PowerPoint.Presentation = pptApp.Presentations.Open(PPTFilename, WithWindow:=MsoTriState.msoFalse)

        Dim CndaXXX As String = FindRegExp(pptPres, "CNDA#+")
        FindReplaceAll(pptPres, CndaXXX, Cnda)
        FindReplaceAll(pptPres, "CustName", Name)

        'Write out pdf
        Dim wPath As String = CreateObject("Scripting.FileSystemObject").GetParentFolderName(PPTFilename)
        Dim wName As String = CreateObject("Scripting.FileSystemObject").GetBaseName(PPTFilename)
        Dim fullName As String = wPath & "\" & wName & "_" & Name & "_" & Cnda & ".pdf"
        pptPres.ExportAsFixedFormat(Path:=fullName,
                                    FixedFormatType:=PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                    Intent:=PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentScreen)
        PowerPointToPDF = fullName
        pptPres.Close()
        pptPres = Nothing
        pptApp.Quit()
        pptApp = Nothing
    End Function
    Sub FindReplaceAll(ByVal pres As PowerPoint.Presentation, FindWord As String, ReplaceWord As String)
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
    Function FindRegExp(ByVal pres As PowerPoint.Presentation, myPattern As String)
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
End Class
