﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CndaOutlookGenPDFandEmailFileDialog
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.SelectPPT_Button = New System.Windows.Forms.Button()
        Me.SelectXml_Button = New System.Windows.Forms.Button()
        Me.PPT_Label = New System.Windows.Forms.Label()
        Me.XLS_Label = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PptFileInstructionLabel = New System.Windows.Forms.Label()
        Me.OpenPPTFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.OpenXMLFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.PickEmailFolderButton = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.EmailFolderLabel = New System.Windows.Forms.Label()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.InfoLabel = New System.Windows.Forms.Label()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(368, 229)
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
        Me.OK_Button.Text = "Go"
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
        'SelectPPT_Button
        '
        Me.SelectPPT_Button.Location = New System.Drawing.Point(26, 25)
        Me.SelectPPT_Button.Name = "SelectPPT_Button"
        Me.SelectPPT_Button.Size = New System.Drawing.Size(91, 23)
        Me.SelectPPT_Button.TabIndex = 1
        Me.SelectPPT_Button.Text = "Select PPT File"
        Me.SelectPPT_Button.UseVisualStyleBackColor = True
        '
        'SelectXml_Button
        '
        Me.SelectXml_Button.Location = New System.Drawing.Point(26, 113)
        Me.SelectXml_Button.Name = "SelectXml_Button"
        Me.SelectXml_Button.Size = New System.Drawing.Size(91, 23)
        Me.SelectXml_Button.TabIndex = 2
        Me.SelectXml_Button.Text = "Select XML File"
        Me.SelectXml_Button.UseVisualStyleBackColor = True
        '
        'PPT_Label
        '
        Me.PPT_Label.AutoEllipsis = True
        Me.PPT_Label.Location = New System.Drawing.Point(123, 30)
        Me.PPT_Label.Name = "PPT_Label"
        Me.PPT_Label.Size = New System.Drawing.Size(297, 18)
        Me.PPT_Label.TabIndex = 3
        Me.PPT_Label.Text = "<no ppt file selected>"
        '
        'XLS_Label
        '
        Me.XLS_Label.AutoEllipsis = True
        Me.XLS_Label.Location = New System.Drawing.Point(123, 118)
        Me.XLS_Label.Name = "XLS_Label"
        Me.XLS_Label.Size = New System.Drawing.Size(297, 18)
        Me.XLS_Label.TabIndex = 4
        Me.XLS_Label.Text = "<no xml file selected>"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(26, 97)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(242, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "XML file containing customer and email addresses"
        '
        'PptFileInstructionLabel
        '
        Me.PptFileInstructionLabel.AutoSize = True
        Me.PptFileInstructionLabel.Location = New System.Drawing.Point(26, 9)
        Me.PptFileInstructionLabel.Name = "PptFileInstructionLabel"
        Me.PptFileInstructionLabel.Size = New System.Drawing.Size(181, 13)
        Me.PptFileInstructionLabel.TabIndex = 6
        Me.PptFileInstructionLabel.Text = "PPT file to update and export to PDF"
        '
        'OpenPPTFileDialog
        '
        Me.OpenPPTFileDialog.Filter = "Powerpoint files|*.ppt?"
        Me.OpenPPTFileDialog.Title = "Select PPT file to Export"
        '
        'OpenXMLFileDialog
        '
        Me.OpenXMLFileDialog.Filter = "XML Files|*.xml|All Files|*.*"
        Me.OpenXMLFileDialog.Title = "Select CNDA XML Customer File"
        '
        'PickEmailFolderButton
        '
        Me.PickEmailFolderButton.Location = New System.Drawing.Point(26, 67)
        Me.PickEmailFolderButton.Name = "PickEmailFolderButton"
        Me.PickEmailFolderButton.Size = New System.Drawing.Size(91, 23)
        Me.PickEmailFolderButton.TabIndex = 7
        Me.PickEmailFolderButton.Text = "Email folder"
        Me.PickEmailFolderButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(26, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(161, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Folder to place generated Emails"
        '
        'EmailFolderLabel
        '
        Me.EmailFolderLabel.AutoSize = True
        Me.EmailFolderLabel.Location = New System.Drawing.Point(123, 72)
        Me.EmailFolderLabel.Name = "EmailFolderLabel"
        Me.EmailFolderLabel.Size = New System.Drawing.Size(103, 13)
        Me.EmailFolderLabel.TabIndex = 9
        Me.EmailFolderLabel.Text = "<no folder selected>"
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.CheckOnClick = True
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(29, 142)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(197, 94)
        Me.CheckedListBox1.Sorted = True
        Me.CheckedListBox1.TabIndex = 10
        '
        'InfoLabel
        '
        Me.InfoLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InfoLabel.Location = New System.Drawing.Point(232, 142)
        Me.InfoLabel.Name = "InfoLabel"
        Me.InfoLabel.Size = New System.Drawing.Size(282, 71)
        Me.InfoLabel.TabIndex = 11
        Me.InfoLabel.Text = "Working..."
        Me.InfoLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.InfoLabel.Visible = False
        '
        'CndaOutlookGenPDFandEmailFileDialog
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(526, 270)
        Me.Controls.Add(Me.InfoLabel)
        Me.Controls.Add(Me.CheckedListBox1)
        Me.Controls.Add(Me.EmailFolderLabel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PickEmailFolderButton)
        Me.Controls.Add(Me.PptFileInstructionLabel)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.XLS_Label)
        Me.Controls.Add(Me.PPT_Label)
        Me.Controls.Add(Me.SelectXml_Button)
        Me.Controls.Add(Me.SelectPPT_Button)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CndaOutlookGenPDFandEmailFileDialog"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "CNDA PDF and Email"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents SelectPPT_Button As System.Windows.Forms.Button
    Friend WithEvents SelectXml_Button As System.Windows.Forms.Button
    Friend WithEvents PPT_Label As System.Windows.Forms.Label
    Friend WithEvents XLS_Label As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PptFileInstructionLabel As System.Windows.Forms.Label
    Friend WithEvents OpenPPTFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents OpenXMLFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents PickEmailFolderButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents EmailFolderLabel As System.Windows.Forms.Label
    Friend WithEvents CheckedListBox1 As System.Windows.Forms.CheckedListBox
    Friend WithEvents InfoLabel As System.Windows.Forms.Label
End Class
