Partial Class CNDAPowerPointRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.PPTTab1 = Me.Factory.CreateRibbonTab
        Me.CNDAGroup = Me.Factory.CreateRibbonGroup
        Me.GeneratePDFButton = Me.Factory.CreateRibbonButton
        Me.PptOpenXlsFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.PPTTab1.SuspendLayout()
        Me.CNDAGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'PPTTab1
        '
        Me.PPTTab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.PPTTab1.Groups.Add(Me.CNDAGroup)
        Me.PPTTab1.Label = "TabAddIns"
        Me.PPTTab1.Name = "PPTTab1"
        '
        'CNDAGroup
        '
        Me.CNDAGroup.Items.Add(Me.GeneratePDFButton)
        Me.CNDAGroup.Label = "CNDA Tools"
        Me.CNDAGroup.Name = "CNDAGroup"
        '
        'GeneratePDFButton
        '
        Me.GeneratePDFButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.GeneratePDFButton.Label = "Generate PDF"
        Me.GeneratePDFButton.Name = "GeneratePDFButton"
        Me.GeneratePDFButton.OfficeImageId = "ExportFile"
        Me.GeneratePDFButton.ShowImage = True
        '
        'PptOpenXlsFileDialog
        '
        Me.PptOpenXlsFileDialog.Filter = "Excel Files|*.xls?|All|*.*"
        '
        'CNDAPowerPointRibbon
        '
        Me.Name = "CNDAPowerPointRibbon"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.PPTTab1)
        Me.PPTTab1.ResumeLayout(False)
        Me.PPTTab1.PerformLayout()
        Me.CNDAGroup.ResumeLayout(False)
        Me.CNDAGroup.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PPTTab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents CNDAGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents GeneratePDFButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PptOpenXlsFileDialog As System.Windows.Forms.OpenFileDialog
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As CNDAPowerPointRibbon
        Get
            Return Me.GetRibbon(Of CNDAPowerPointRibbon)()
        End Get
    End Property
End Class
