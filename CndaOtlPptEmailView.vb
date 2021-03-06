﻿Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook

Public Class CndaOtlPptEmailView
    Implements ICndaOutlookEvents

    Private PptFilename As String = ""
    Public Property XmlFilename As String = ""
    Private EmailsGenerated As Boolean = False

    Public Event XmlFileChangeEvent(xmlFilename As String,
                                    ByRef objList As CheckedListBox.ObjectCollection) Implements ICndaOutlookEvents.XmlFileChangeEvent
    Public Event SendEmailsEvent(ByRef objList As CheckedListBox.CheckedItemCollection,
                                 ByRef count As Integer) Implements ICndaOutlookEvents.SendEmailsEvent
    Public Event PptFileChangeEvent(pptFilename As String) Implements ICndaOutlookEvents.PptFileChangeEvent
    Private Sub CndaOtlPptEmailView_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OtlPptXmlLabel.Text = XmlFilename
        RaiseEvent XmlFileChangeEvent(XmlFilename, OtlPptCheckedListBox.Items)
        For i = 0 To OtlPptCheckedListBox.Items.Count - 1
            OtlPptCheckedListBox.SetItemChecked(i, CheckState.Checked)
        Next
        Update()
    End Sub
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OtlPptOK_Button.Click
        If PptFilename = "" Then
            Dim msgbxstatus1 As MsgBoxResult = MsgBox("Error PPT file not entered", MsgBoxStyle.RetryCancel)
            If msgbxstatus1 = MsgBoxResult.Cancel Then
                DialogResult = DialogResult.Cancel
                Close()
            End If
        ElseIf XmlFilename = "" Then
            Dim msgbxstatus1 As MsgBoxResult = MsgBox("Error XLS file not entered", MsgBoxStyle.RetryCancel)
            If msgbxstatus1 = MsgBoxResult.Cancel Then
                DialogResult = DialogResult.Cancel
                Close()
            End If
        Else
            If Not EmailsGenerated Then
                OtlPptWorkingLabel.Visible = True
                OtlPptWorkingLabel.Text = $"Generating {OtlPptCheckedListBox.CheckedItems.Count} Emails..." & vbCrLf _
                    & "...Please stand by"
                Dim count As Integer
                RaiseEvent SendEmailsEvent(OtlPptCheckedListBox.CheckedItems, count)
                OtlPptWorkingLabel.Text = $"CNDA generated {count} emails" & vbCrLf _
                & "Do you wish to delete the current email?"
                OtlPptOK_Button.Text = My.Resources.YESString
                OtlPptCancel_Button.Text = My.Resources.NOString

                EmailsGenerated = True
            Else
                DialogResult = DialogResult.Yes
                Close()
            End If
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OtlPptCancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub OtlPptXmlButton_Click(sender As Object, e As EventArgs) Handles OtlPptXmlButton.Click
        If OtlPptXmlOpenFileDialog.ShowDialog = DialogResult.OK Then
            XmlFilename = OtlPptXmlOpenFileDialog.FileName
            RaiseEvent XmlFileChangeEvent(XmlFilename, OtlPptCheckedListBox.Items)
            For i = 0 To OtlPptCheckedListBox.Items.Count - 1
                OtlPptCheckedListBox.SetItemChecked(i, CheckState.Checked)
            Next
            Update()
        End If
    End Sub

    Private Sub OtlPptPptButton_Click(sender As Object, e As EventArgs) Handles OtlPptPptButton.Click
        If OtlPptPptOpenFileDialog.ShowDialog = DialogResult.OK Then
            PptFilename = OtlPptPptOpenFileDialog.FileName
            RaiseEvent PptFileChangeEvent(PptFilename)
            OtlPptPptLabel.Text = PptFilename
            Update()
        End If
    End Sub


End Class
