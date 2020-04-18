<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CndaOutlookEmailView
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button1 = New System.Windows.Forms.Button()
        Me.Cancel_Button1 = New System.Windows.Forms.Button()
        Me.XmlLabel1 = New System.Windows.Forms.Label()
        Me.GetXlsButton = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GetEmailFolderButton = New System.Windows.Forms.Button()
        Me.XlsFilenameLabel = New System.Windows.Forms.Label()
        Me.EmailFolderLabel1 = New System.Windows.Forms.Label()
        Me.XmlOpenFileDialog2 = New System.Windows.Forms.OpenFileDialog()
        Me.EmailViewCheckedListBox = New System.Windows.Forms.CheckedListBox()
        Me.WorkingLabel = New System.Windows.Forms.Label()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.OK_Button1, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Cancel_Button1, 1, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(395, 181)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel2.TabIndex = 0
        '
        'OK_Button1
        '
        Me.OK_Button1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button1.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button1.Name = "OK_Button1"
        Me.OK_Button1.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button1.TabIndex = 0
        Me.OK_Button1.Text = "GO"
        '
        'Cancel_Button1
        '
        Me.Cancel_Button1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button1.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button1.Name = "Cancel_Button1"
        Me.Cancel_Button1.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button1.TabIndex = 1
        Me.Cancel_Button1.Text = "Cancel"
        '
        'XmlLabel1
        '
        Me.XmlLabel1.AutoSize = True
        Me.XmlLabel1.Location = New System.Drawing.Point(13, 62)
        Me.XmlLabel1.Name = "XmlLabel1"
        Me.XmlLabel1.Size = New System.Drawing.Size(263, 13)
        Me.XmlLabel1.TabIndex = 1
        Me.XmlLabel1.Text = "XML file containing Customer info and email addresses"
        '
        'GetXlsButton
        '
        Me.GetXlsButton.Location = New System.Drawing.Point(13, 79)
        Me.GetXlsButton.Name = "GetXlsButton"
        Me.GetXlsButton.Size = New System.Drawing.Size(92, 23)
        Me.GetXlsButton.TabIndex = 2
        Me.GetXlsButton.Text = "Select XML File"
        Me.GetXlsButton.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(161, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Folder to place generated Emails"
        '
        'GetEmailFolderButton
        '
        Me.GetEmailFolderButton.Location = New System.Drawing.Point(13, 30)
        Me.GetEmailFolderButton.Name = "GetEmailFolderButton"
        Me.GetEmailFolderButton.Size = New System.Drawing.Size(92, 23)
        Me.GetEmailFolderButton.TabIndex = 4
        Me.GetEmailFolderButton.Text = "Email Folder"
        Me.GetEmailFolderButton.UseVisualStyleBackColor = True
        '
        'XlsFilenameLabel
        '
        Me.XlsFilenameLabel.AutoSize = True
        Me.XlsFilenameLabel.Location = New System.Drawing.Point(111, 84)
        Me.XlsFilenameLabel.Name = "XlsFilenameLabel"
        Me.XlsFilenameLabel.Size = New System.Drawing.Size(76, 13)
        Me.XlsFilenameLabel.TabIndex = 5
        Me.XlsFilenameLabel.Text = "<xml filename>"
        '
        'EmailFolderLabel1
        '
        Me.EmailFolderLabel1.AutoSize = True
        Me.EmailFolderLabel1.Location = New System.Drawing.Point(111, 35)
        Me.EmailFolderLabel1.Name = "EmailFolderLabel1"
        Me.EmailFolderLabel1.Size = New System.Drawing.Size(72, 13)
        Me.EmailFolderLabel1.TabIndex = 6
        Me.EmailFolderLabel1.Text = "<email folder>"
        '
        'XmlOpenFileDialog2
        '
        Me.XmlOpenFileDialog2.FileName = "OpenXmlFileDialog2"
        Me.XmlOpenFileDialog2.Filter = "XML File|*.xml|All Files|*.*"
        '
        'EmailViewCheckedListBox
        '
        Me.EmailViewCheckedListBox.CheckOnClick = True
        Me.EmailViewCheckedListBox.FormattingEnabled = True
        Me.EmailViewCheckedListBox.Location = New System.Drawing.Point(13, 109)
        Me.EmailViewCheckedListBox.Name = "EmailViewCheckedListBox"
        Me.EmailViewCheckedListBox.Size = New System.Drawing.Size(198, 94)
        Me.EmailViewCheckedListBox.Sorted = True
        Me.EmailViewCheckedListBox.TabIndex = 7
        '
        'WorkingLabel
        '
        Me.WorkingLabel.Enabled = False
        Me.WorkingLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WorkingLabel.Location = New System.Drawing.Point(217, 109)
        Me.WorkingLabel.Name = "WorkingLabel"
        Me.WorkingLabel.Size = New System.Drawing.Size(324, 61)
        Me.WorkingLabel.TabIndex = 8
        Me.WorkingLabel.Text = "WorkingLabel"
        Me.WorkingLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.WorkingLabel.Visible = False
        '
        'CndaOutlookEmailView
        '
        Me.AcceptButton = Me.OK_Button1
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button1
        Me.ClientSize = New System.Drawing.Size(553, 222)
        Me.Controls.Add(Me.WorkingLabel)
        Me.Controls.Add(Me.EmailViewCheckedListBox)
        Me.Controls.Add(Me.EmailFolderLabel1)
        Me.Controls.Add(Me.XlsFilenameLabel)
        Me.Controls.Add(Me.GetEmailFolderButton)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GetXlsButton)
        Me.Controls.Add(Me.XmlLabel1)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CndaOutlookEmailView"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Generate Emails From XML Customer Data"
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button1 As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button1 As System.Windows.Forms.Button
    Friend WithEvents XmlLabel1 As System.Windows.Forms.Label
    Friend WithEvents GetXlsButton As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GetEmailFolderButton As System.Windows.Forms.Button
    Friend WithEvents XlsFilenameLabel As System.Windows.Forms.Label
    Friend WithEvents EmailFolderLabel1 As System.Windows.Forms.Label
    Friend WithEvents XmlOpenFileDialog2 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents EmailViewCheckedListBox As System.Windows.Forms.CheckedListBox
    Friend WithEvents WorkingLabel As System.Windows.Forms.Label
End Class
