Imports System

Public Class CndaOutlookGenPDFandEmailFileDialog
    Private PptFilename As String = ""
    Private XmlFilename As String = ""
    Property GeneratePdf As Boolean = True
    Private _generated As Boolean = False

    Public Event GeneratePdfEvent(ByVal pptFilename As String,
                                  ByRef objList As List(Of CndaCustInfo))
    Public Event GenerateEmailEvent(ByVal pptFilename As String, ByRef mailCnt As Integer)
    Public Event XmlFileChangeEvent(ByVal xmlFilename As String,
                                    ByRef objList As List(Of CndaCustInfo))

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If _generated Then
            DialogResult = System.Windows.Forms.DialogResult.OK
            Close()
        Else
            If PptFilename = "" Then
                Dim msgbxstatus As MsgBoxResult = MsgBox("Error PPT file not entered", MsgBoxStyle.RetryCancel)
                If msgbxstatus = MsgBoxResult.Cancel Then
                    DialogResult = System.Windows.Forms.DialogResult.Cancel
                    Close()
                End If
            ElseIf XmlFilename = "" Then
                Dim msgbxstatus1 As MsgBoxResult = MsgBox("Error XLS file not entered", MsgBoxStyle.RetryCancel)
                If msgbxstatus1 = MsgBoxResult.Cancel Then
                    DialogResult = System.Windows.Forms.DialogResult.Cancel
                    Close()
                End If
            Else
                OK_Button.Enabled = False
                InfoLabel.Visible = True
                If GeneratePdf Then
                    InfoLabel.Text = "Generating PDFs..."
                    Update()
                    Dim objList As New List(Of CndaCustInfo)
                    For Each obj As CndaCustInfo In CheckedListBox1.CheckedItems
                        objList.Add(obj)
                    Next
                    RaiseEvent GeneratePdfEvent(PptFilename, objList)
                End If
                InfoLabel.Text = "Saving emails..."
                OK_Button.Enabled = False
                Update()
                Dim cnt As Integer
                RaiseEvent GenerateEmailEvent(PptFilename, cnt)
                _generated = True
                OK_Button.Enabled = True
                OK_Button.Text = "OK"
                Cancel_Button.Enabled = False
                Cancel_Button.Visible = False
                InfoLabel.Text = $"Generated {cnt} Emails" + vbCrLf + vbCrLf + "Click OK to continue"
                Update()
            End If
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Dialog1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim f As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetFolderFromID(My.Settings.MailFolderId)
        EmailFolderLabel.Text = f.Name
        XmlFilename = My.Settings.XmlFileName
        XLS_Label.Text = XmlFilename
        Dim custList As List(Of CndaCustInfo) = Nothing
        RaiseEvent XmlFileChangeEvent(XmlFilename, custList)
        CheckedListBox1.DisplayMember = "CustName"
        For Each cust As Object In custList
            CheckedListBox1.Items.Add(cust)
        Next cust
        Update()
        _generated = False
    End Sub

    Private Sub SelectPPT_Button_Click(sender As Object, e As EventArgs) Handles SelectPPT_Button.Click
        OpenPPTFileDialog.ShowDialog()
        PptFilename = OpenPPTFileDialog.FileName
        PPT_Label.Text = PptFilename
    End Sub

    Private Sub SelectXml_Button_Click(sender As Object, e As EventArgs) Handles SelectXml_Button.Click
        OpenXMLFileDialog.ShowDialog()
        XmlFilename = OpenXMLFileDialog.FileName
        XLS_Label.Text = XmlFilename
        Dim custList As List(Of CndaCustInfo) = Nothing
        RaiseEvent XmlFileChangeEvent(XmlFilename, custList)
        CheckedListBox1.DisplayMember = "CustName"
        For Each cust As Object In custList
            CheckedListBox1.Items.Add(cust)
        Next cust
        Update()
    End Sub


    Private Sub PickEmailFolderButton_Click(sender As Object, e As EventArgs) Handles PickEmailFolderButton.Click
        Dim dg As Outlook.Folder = Globals.ThisAddIn.Application.Session.PickFolder()
        If dg IsNot Nothing Then
            My.Settings.MailFolderId = dg.EntryID
            EmailFolderLabel.Text = dg.Name
            My.Settings.Save()
        End If
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged

    End Sub
End Class
