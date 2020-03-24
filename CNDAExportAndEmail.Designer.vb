﻿Partial Class CNDAExportAndEmail
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
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.CNDA_Group = Me.Factory.CreateRibbonGroup
        Me.CNDAUpdateEmail_Button = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.CNDA_Group.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.CNDA_Group)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'CNDA_Group
        '
        Me.CNDA_Group.Items.Add(Me.CNDAUpdateEmail_Button)
        Me.CNDA_Group.Label = "CNDA Tools"
        Me.CNDA_Group.Name = "CNDA_Group"
        '
        'CNDAUpdateEmail_Button
        '
        Me.CNDAUpdateEmail_Button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.CNDAUpdateEmail_Button.Label = "CNDA Update and Email"
        Me.CNDAUpdateEmail_Button.Name = "CNDAUpdateEmail_Button"
        Me.CNDAUpdateEmail_Button.OfficeImageId = "SendAsPdfAttachment"
        Me.CNDAUpdateEmail_Button.ShowImage = True
        '
        'CNDAExportAndEmail
        '
        Me.Name = "CNDAExportAndEmail"
        Me.RibbonType = "Microsoft.Outlook.Mail.Compose"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.CNDA_Group.ResumeLayout(False)
        Me.CNDA_Group.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents CNDA_Group As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents CNDAUpdateEmail_Button As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As CNDAExportAndEmail
        Get
            Return Me.GetRibbon(Of CNDAExportAndEmail)()
        End Get
    End Property
End Class
