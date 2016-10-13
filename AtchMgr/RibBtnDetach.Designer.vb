Partial Class RibBtnDetach
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabAttachments = Me.Factory.CreateRibbonTab
        Me.GroupDetach = Me.Factory.CreateRibbonGroup
        Me.DoDetach = Me.Factory.CreateRibbonButton
        Me.TabAttachments.SuspendLayout()
        Me.GroupDetach.SuspendLayout()
        '
        'TabAttachments
        '
        Me.TabAttachments.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.TabAttachments.ControlId.OfficeId = "TabAttachments"
        Me.TabAttachments.Groups.Add(Me.GroupDetach)
        Me.TabAttachments.Label = "TabAttachments"
        Me.TabAttachments.Name = "TabAttachments"
        '
        'GroupDetach
        '
        Me.GroupDetach.Items.Add(Me.DoDetach)
        Me.GroupDetach.Label = "Detach Attachment(s)"
        Me.GroupDetach.Name = "GroupDetach"
        '
        'DoDetach
        '
        Me.DoDetach.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DoDetach.Label = "Detach"
        Me.DoDetach.Name = "DoDetach"
        Me.DoDetach.OfficeImageId = "SaveSentItemsMenu"
        Me.DoDetach.ShowImage = True
        '
        'RibBtnDetach
        '
        Me.Name = "RibBtnDetach"
        Me.RibbonType = "Microsoft.Outlook.Mail.Read"
        Me.Tabs.Add(Me.TabAttachments)
        Me.TabAttachments.ResumeLayout(False)
        Me.TabAttachments.PerformLayout()
        Me.GroupDetach.ResumeLayout(False)
        Me.GroupDetach.PerformLayout()

    End Sub

    Friend WithEvents TabAttachments As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents GroupDetach As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents DoDetach As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property RibBtnDetach() As RibBtnDetach
        Get
            Return Me.GetRibbon(Of RibBtnDetach)()
        End Get
    End Property
End Class
