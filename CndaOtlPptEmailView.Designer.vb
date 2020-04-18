<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CndaOtlPptEmailView
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
        Me.OtlPptOK_Button = New System.Windows.Forms.Button()
        Me.OtlPptCancel_Button = New System.Windows.Forms.Button()
        Me.OtlPptCheckedListBox = New System.Windows.Forms.CheckedListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OtlPptPptButton = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.OtlPptFolderButton = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.OtlPptXmlButton = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.OtlPptWorkingLabel = New System.Windows.Forms.Label()
        Me.OtlPptXmlLabel = New System.Windows.Forms.Label()
        Me.OtlPptFolderLabel = New System.Windows.Forms.Label()
        Me.OtlPptPptLabel = New System.Windows.Forms.Label()
        Me.OtlPptXmlOpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.OtlPptPptOpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OtlPptOK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.OtlPptCancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(383, 244)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OtlPptOK_Button
        '
        Me.OtlPptOK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OtlPptOK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OtlPptOK_Button.Name = "OtlPptOK_Button"
        Me.OtlPptOK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OtlPptOK_Button.TabIndex = 0
        Me.OtlPptOK_Button.Text = "GO"
        '
        'OtlPptCancel_Button
        '
        Me.OtlPptCancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OtlPptCancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.OtlPptCancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.OtlPptCancel_Button.Name = "OtlPptCancel_Button"
        Me.OtlPptCancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.OtlPptCancel_Button.TabIndex = 1
        Me.OtlPptCancel_Button.Text = "Cancel"
        '
        'OtlPptCheckedListBox
        '
        Me.OtlPptCheckedListBox.FormattingEnabled = True
        Me.OtlPptCheckedListBox.Location = New System.Drawing.Point(16, 162)
        Me.OtlPptCheckedListBox.Name = "OtlPptCheckedListBox"
        Me.OtlPptCheckedListBox.Size = New System.Drawing.Size(175, 109)
        Me.OtlPptCheckedListBox.Sorted = True
        Me.OtlPptCheckedListBox.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "PowerPoint File"
        '
        'OtlPptPptButton
        '
        Me.OtlPptPptButton.Location = New System.Drawing.Point(16, 25)
        Me.OtlPptPptButton.Name = "OtlPptPptButton"
        Me.OtlPptPptButton.Size = New System.Drawing.Size(97, 23)
        Me.OtlPptPptButton.TabIndex = 3
        Me.OtlPptPptButton.Text = "Select PPT File"
        Me.OtlPptPptButton.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(161, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Folder to place generated Emails"
        '
        'OtlPptFolderButton
        '
        Me.OtlPptFolderButton.Location = New System.Drawing.Point(16, 71)
        Me.OtlPptFolderButton.Name = "OtlPptFolderButton"
        Me.OtlPptFolderButton.Size = New System.Drawing.Size(97, 23)
        Me.OtlPptFolderButton.TabIndex = 5
        Me.OtlPptFolderButton.Text = "Email Folder"
        Me.OtlPptFolderButton.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 103)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(115, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Customer info XML File"
        '
        'OtlPptXmlButton
        '
        Me.OtlPptXmlButton.Location = New System.Drawing.Point(16, 119)
        Me.OtlPptXmlButton.Name = "OtlPptXmlButton"
        Me.OtlPptXmlButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OtlPptXmlButton.Size = New System.Drawing.Size(97, 23)
        Me.OtlPptXmlButton.TabIndex = 7
        Me.OtlPptXmlButton.Text = "Select XML File"
        Me.OtlPptXmlButton.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 146)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Customer List"
        '
        'OtlPptWorkingLabel
        '
        Me.OtlPptWorkingLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OtlPptWorkingLabel.Location = New System.Drawing.Point(197, 162)
        Me.OtlPptWorkingLabel.Name = "OtlPptWorkingLabel"
        Me.OtlPptWorkingLabel.Size = New System.Drawing.Size(325, 79)
        Me.OtlPptWorkingLabel.TabIndex = 9
        Me.OtlPptWorkingLabel.Text = "Working"
        Me.OtlPptWorkingLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.OtlPptWorkingLabel.Visible = False
        '
        'OtlPptXmlLabel
        '
        Me.OtlPptXmlLabel.AutoSize = True
        Me.OtlPptXmlLabel.Location = New System.Drawing.Point(119, 124)
        Me.OtlPptXmlLabel.Name = "OtlPptXmlLabel"
        Me.OtlPptXmlLabel.Size = New System.Drawing.Size(76, 13)
        Me.OtlPptXmlLabel.TabIndex = 10
        Me.OtlPptXmlLabel.Text = "<xml filename>"
        '
        'OtlPptFolderLabel
        '
        Me.OtlPptFolderLabel.AutoSize = True
        Me.OtlPptFolderLabel.Location = New System.Drawing.Point(119, 76)
        Me.OtlPptFolderLabel.Name = "OtlPptFolderLabel"
        Me.OtlPptFolderLabel.Size = New System.Drawing.Size(72, 13)
        Me.OtlPptFolderLabel.TabIndex = 11
        Me.OtlPptFolderLabel.Text = "<email folder>"
        '
        'OtlPptPptLabel
        '
        Me.OtlPptPptLabel.AutoSize = True
        Me.OtlPptPptLabel.Location = New System.Drawing.Point(119, 30)
        Me.OtlPptPptLabel.Name = "OtlPptPptLabel"
        Me.OtlPptPptLabel.Size = New System.Drawing.Size(76, 13)
        Me.OtlPptPptLabel.TabIndex = 12
        Me.OtlPptPptLabel.Text = "<ppt filename>"
        '
        'OtlPptXmlOpenFileDialog
        '
        Me.OtlPptXmlOpenFileDialog.FileName = "OtlPptXmlOpenFileDialog"
        Me.OtlPptXmlOpenFileDialog.Filter = "XML Files|*.xml|all files|*.*"
        '
        'OtlPptPptOpenFileDialog
        '
        Me.OtlPptPptOpenFileDialog.FileName = "OtlPptPptOpenFileDialog"
        Me.OtlPptPptOpenFileDialog.Filter = "PPT files|*.ppt?|All Files|*.*"
        '
        'CndaOtlPptEmailView
        '
        Me.AcceptButton = Me.OtlPptOK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.OtlPptCancel_Button
        Me.ClientSize = New System.Drawing.Size(541, 285)
        Me.Controls.Add(Me.OtlPptPptLabel)
        Me.Controls.Add(Me.OtlPptFolderLabel)
        Me.Controls.Add(Me.OtlPptXmlLabel)
        Me.Controls.Add(Me.OtlPptWorkingLabel)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.OtlPptXmlButton)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.OtlPptFolderButton)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.OtlPptPptButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.OtlPptCheckedListBox)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CndaOtlPptEmailView"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Generate PDF and Emails"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OtlPptOK_Button As System.Windows.Forms.Button
    Friend WithEvents OtlPptCancel_Button As System.Windows.Forms.Button
    Friend WithEvents OtlPptCheckedListBox As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents OtlPptPptButton As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OtlPptFolderButton As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents OtlPptXmlButton As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents OtlPptWorkingLabel As System.Windows.Forms.Label
    Friend WithEvents OtlPptXmlLabel As System.Windows.Forms.Label
    Friend WithEvents OtlPptFolderLabel As System.Windows.Forms.Label
    Friend WithEvents OtlPptPptLabel As System.Windows.Forms.Label
    Friend WithEvents OtlPptXmlOpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents OtlPptPptOpenFileDialog As System.Windows.Forms.OpenFileDialog
End Class
