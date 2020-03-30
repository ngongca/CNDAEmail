<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GetFileDialog
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
        Me.SelectPPT_Button = New System.Windows.Forms.Button()
        Me.SelectXLS_Button = New System.Windows.Forms.Button()
        Me.PPT_Label = New System.Windows.Forms.Label()
        Me.XLS_Label = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PptFileInstructionLabel = New System.Windows.Forms.Label()
        Me.OpenPPTFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.OpenXLSFileDialog = New System.Windows.Forms.OpenFileDialog()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(277, 106)
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
        'SelectPPT_Button
        '
        Me.SelectPPT_Button.Location = New System.Drawing.Point(26, 25)
        Me.SelectPPT_Button.Name = "SelectPPT_Button"
        Me.SelectPPT_Button.Size = New System.Drawing.Size(91, 23)
        Me.SelectPPT_Button.TabIndex = 1
        Me.SelectPPT_Button.Text = "Select PPT File"
        Me.SelectPPT_Button.UseVisualStyleBackColor = True
        '
        'SelectXLS_Button
        '
        Me.SelectXLS_Button.Location = New System.Drawing.Point(26, 72)
        Me.SelectXLS_Button.Name = "SelectXLS_Button"
        Me.SelectXLS_Button.Size = New System.Drawing.Size(91, 23)
        Me.SelectXLS_Button.TabIndex = 2
        Me.SelectXLS_Button.Text = "Select XLS File"
        Me.SelectXLS_Button.UseVisualStyleBackColor = True
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
        Me.XLS_Label.Location = New System.Drawing.Point(123, 77)
        Me.XLS_Label.Name = "XLS_Label"
        Me.XLS_Label.Size = New System.Drawing.Size(297, 18)
        Me.XLS_Label.TabIndex = 4
        Me.XLS_Label.Text = "<no xls file selected>"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(32, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(240, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "XLS file containing NDA info and email addresses"
        '
        'PptFileInstructionLabel
        '
        Me.PptFileInstructionLabel.AutoSize = True
        Me.PptFileInstructionLabel.Location = New System.Drawing.Point(32, 9)
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
        'OpenXLSFileDialog
        '
        Me.OpenXLSFileDialog.Filter = "Excel Files|*.xls?"
        Me.OpenXLSFileDialog.Title = "Select CNDA Excel File"
        '
        'GetFileDialog
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(435, 147)
        Me.Controls.Add(Me.PptFileInstructionLabel)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.XLS_Label)
        Me.Controls.Add(Me.PPT_Label)
        Me.Controls.Add(Me.SelectXLS_Button)
        Me.Controls.Add(Me.SelectPPT_Button)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "GetFileDialog"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Enter File Data"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents SelectPPT_Button As Windows.Forms.Button
    Friend WithEvents SelectXLS_Button As Windows.Forms.Button
    Friend WithEvents PPT_Label As Windows.Forms.Label
    Friend WithEvents XLS_Label As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents PptFileInstructionLabel As Windows.Forms.Label
    Friend WithEvents OpenPPTFileDialog As Windows.Forms.OpenFileDialog
    Friend WithEvents OpenXLSFileDialog As Windows.Forms.OpenFileDialog
End Class
