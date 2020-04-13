<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CndaInfoForm
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
        Me.InfoLabel = New System.Windows.Forms.Label()
        Me.PptDoneButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'InfoLabel
        '
        Me.InfoLabel.AutoEllipsis = True
        Me.InfoLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InfoLabel.Location = New System.Drawing.Point(12, 9)
        Me.InfoLabel.Name = "InfoLabel"
        Me.InfoLabel.Size = New System.Drawing.Size(244, 22)
        Me.InfoLabel.TabIndex = 0
        Me.InfoLabel.Text = "Working"
        Me.InfoLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.InfoLabel.UseWaitCursor = True
        '
        'PptDoneButton
        '
        Me.PptDoneButton.Enabled = False
        Me.PptDoneButton.Location = New System.Drawing.Point(91, 51)
        Me.PptDoneButton.Name = "PptDoneButton"
        Me.PptDoneButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.PptDoneButton.Size = New System.Drawing.Size(75, 23)
        Me.PptDoneButton.TabIndex = 1
        Me.PptDoneButton.Text = "Dismiss"
        Me.PptDoneButton.UseCompatibleTextRendering = True
        Me.PptDoneButton.UseVisualStyleBackColor = True
        Me.PptDoneButton.Visible = False
        '
        'CndaInfoForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(268, 86)
        Me.Controls.Add(Me.PptDoneButton)
        Me.Controls.Add(Me.InfoLabel)
        Me.Name = "CndaInfoForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "CndaInfoForm"
        Me.UseWaitCursor = True
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents InfoLabel As System.Windows.Forms.Label
    Friend WithEvents PptDoneButton As System.Windows.Forms.Button
End Class
