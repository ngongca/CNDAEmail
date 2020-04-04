<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CndaOutlookEmailOnlyForm
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GetXlsButton = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GetEmailFolderButton = New System.Windows.Forms.Button()
        Me.XlsFilenameLabel = New System.Windows.Forms.Label()
        Me.EmailFolderLabel = New System.Windows.Forms.Label()
        Me.XlsOpenFileDialog2 = New System.Windows.Forms.OpenFileDialog()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(379, 113)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(240, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "XLS file containing NDA info and email addresses"
        '
        'GetXlsButton
        '
        Me.GetXlsButton.Location = New System.Drawing.Point(13, 30)
        Me.GetXlsButton.Name = "GetXlsButton"
        Me.GetXlsButton.Size = New System.Drawing.Size(92, 23)
        Me.GetXlsButton.TabIndex = 2
        Me.GetXlsButton.Text = "Select XLS File"
        Me.GetXlsButton.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(161, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Folder to place generated Emails"
        '
        'GetEmailFolderButton
        '
        Me.GetEmailFolderButton.Location = New System.Drawing.Point(13, 77)
        Me.GetEmailFolderButton.Name = "GetEmailFolderButton"
        Me.GetEmailFolderButton.Size = New System.Drawing.Size(92, 23)
        Me.GetEmailFolderButton.TabIndex = 4
        Me.GetEmailFolderButton.Text = "Email Folder"
        Me.GetEmailFolderButton.UseVisualStyleBackColor = True
        '
        'XlsFilenameLabel
        '
        Me.XlsFilenameLabel.AutoSize = True
        Me.XlsFilenameLabel.Location = New System.Drawing.Point(111, 35)
        Me.XlsFilenameLabel.Name = "XlsFilenameLabel"
        Me.XlsFilenameLabel.Size = New System.Drawing.Size(73, 13)
        Me.XlsFilenameLabel.TabIndex = 5
        Me.XlsFilenameLabel.Text = "<xls filename>"
        '
        'EmailFolderLabel
        '
        Me.EmailFolderLabel.AutoSize = True
        Me.EmailFolderLabel.Location = New System.Drawing.Point(111, 82)
        Me.EmailFolderLabel.Name = "EmailFolderLabel"
        Me.EmailFolderLabel.Size = New System.Drawing.Size(72, 13)
        Me.EmailFolderLabel.TabIndex = 6
        Me.EmailFolderLabel.Text = "<email folder>"
        '
        'XlsOpenFileDialog2
        '
        Me.XlsOpenFileDialog2.FileName = "OpenXlsFileDialog2"
        Me.XlsOpenFileDialog2.Filter = "Excel File|*.xls?|All Files|*.*"
        '
        'CndaOutlookEmailOnlyForm
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(537, 154)
        Me.Controls.Add(Me.EmailFolderLabel)
        Me.Controls.Add(Me.XlsFilenameLabel)
        Me.Controls.Add(Me.GetEmailFolderButton)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GetXlsButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CndaOutlookEmailOnlyForm"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Generate Emails From Excel Data"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GetXlsButton As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GetEmailFolderButton As System.Windows.Forms.Button
    Friend WithEvents XlsFilenameLabel As System.Windows.Forms.Label
    Friend WithEvents EmailFolderLabel As System.Windows.Forms.Label
    Friend WithEvents XlsOpenFileDialog2 As System.Windows.Forms.OpenFileDialog
End Class
