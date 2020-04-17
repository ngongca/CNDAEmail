<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CndaPptGenView
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
        Me.XmlFilenameLabel = New System.Windows.Forms.Label()
        Me.GetXmlButton = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.XmlOpenFileDialog2 = New System.Windows.Forms.OpenFileDialog()
        Me.WorkingLabel = New System.Windows.Forms.Label()
        Me.PptOpenFileDialog2 = New System.Windows.Forms.OpenFileDialog()
        Me.PptCheckedListBox = New System.Windows.Forms.CheckedListBox()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(297, 126)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 25)
        Me.TableLayoutPanel1.TabIndex = 7
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 19)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "GO"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 19)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'XmlFilenameLabel
        '
        Me.XmlFilenameLabel.AutoSize = True
        Me.XmlFilenameLabel.Location = New System.Drawing.Point(110, 31)
        Me.XmlFilenameLabel.Name = "XmlFilenameLabel"
        Me.XmlFilenameLabel.Size = New System.Drawing.Size(76, 13)
        Me.XmlFilenameLabel.TabIndex = 12
        Me.XmlFilenameLabel.Text = "<xml filename>"
        '
        'GetXmlButton
        '
        Me.GetXmlButton.Location = New System.Drawing.Point(12, 26)
        Me.GetXmlButton.Name = "GetXmlButton"
        Me.GetXmlButton.Size = New System.Drawing.Size(92, 23)
        Me.GetXmlButton.TabIndex = 9
        Me.GetXmlButton.Text = "Select XML File"
        Me.GetXmlButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(263, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "XML file containing Customer info and email addresses"
        '
        'XmlOpenFileDialog2
        '
        Me.XmlOpenFileDialog2.FileName = "OpenXmlFileDialog2"
        Me.XmlOpenFileDialog2.Filter = "XML File|*.xml|All Files|*.*"
        '
        'WorkingLabel
        '
        Me.WorkingLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WorkingLabel.Location = New System.Drawing.Point(215, 55)
        Me.WorkingLabel.Name = "WorkingLabel"
        Me.WorkingLabel.Size = New System.Drawing.Size(237, 65)
        Me.WorkingLabel.TabIndex = 14
        Me.WorkingLabel.Text = "Working"
        Me.WorkingLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.WorkingLabel.Visible = False
        '
        'PptOpenFileDialog2
        '
        Me.PptOpenFileDialog2.FileName = "PptOpenFileDialog2"
        Me.PptOpenFileDialog2.Filter = "PPT files|*.ppt?|All files|*.*"
        '
        'PptCheckedListBox
        '
        Me.PptCheckedListBox.CheckOnClick = True
        Me.PptCheckedListBox.FormattingEnabled = True
        Me.PptCheckedListBox.Location = New System.Drawing.Point(12, 55)
        Me.PptCheckedListBox.Name = "PptCheckedListBox"
        Me.PptCheckedListBox.Size = New System.Drawing.Size(197, 94)
        Me.PptCheckedListBox.Sorted = True
        Me.PptCheckedListBox.TabIndex = 15
        '
        'CndaPptGenView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(455, 163)
        Me.Controls.Add(Me.PptCheckedListBox)
        Me.Controls.Add(Me.WorkingLabel)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.XmlFilenameLabel)
        Me.Controls.Add(Me.GetXmlButton)
        Me.Controls.Add(Me.Label1)
        Me.Name = "CndaPptGenView"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Generate PDF Files"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents XmlFilenameLabel As System.Windows.Forms.Label
    Friend WithEvents GetXmlButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents XmlOpenFileDialog2 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents WorkingLabel As System.Windows.Forms.Label
    Friend WithEvents PptOpenFileDialog2 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents PptCheckedListBox As System.Windows.Forms.CheckedListBox
End Class
