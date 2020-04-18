<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CndaOutlookPptView
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
        Me.SuspendLayout()
        '
        'CndaOutlookEmailView
        '
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Name = "CndaOutlookEmailView"
        Me.ResumeLayout(False)

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
