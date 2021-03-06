Partial Class AMLAS
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AMLAS))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "AMLAS"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button3)
        Me.Group1.Label = "Create new"
        Me.Group1.Name = "Group1"
        '
        'Button3
        '
        Me.Button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Label = "Create new doc"
        Me.Button3.Name = "Button3"
        Me.Button3.ScreenTip = "Create copy from template"
        Me.Button3.ShowImage = True
        Me.Button3.SuperTip = "Create new"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button1)
        Me.Group2.Items.Add(Me.Button4)
        Me.Group2.Items.Add(Me.Button2)
        Me.Group2.Label = "Navigation"
        Me.Group2.Name = "Group2"
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Label = "Back to overview"
        Me.Button1.Name = "Button1"
        Me.Button1.ScreenTip = "Overview page"
        Me.Button1.ShowImage = True
        Me.Button1.SuperTip = "Go back to overview page"
        '
        'Button4
        '
        Me.Button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.ImageName = "Toggle"
        Me.Button4.Label = "Toggle "
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'Button2
        '
        Me.Button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Label = "Back to previous page"
        Me.Button2.Name = "Button2"
        Me.Button2.ScreenTip = "Toggle between overview and previous page"
        Me.Button2.ShowImage = True
        Me.Button2.SuperTip = "Toggle page"
        '
        'AMLAS
        '
        Me.Name = "AMLAS"
        Me.RibbonType = "Microsoft.Visio.Drawing"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As AMLAS
        Get
            Return Me.GetRibbon(Of AMLAS)()
        End Get
    End Property
End Class
