Partial Class SapHFMRibbon
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SapHFMRibbon))
        Me.SapHfm = Me.Factory.CreateRibbonTab
        Me.Generate = Me.Factory.CreateRibbonGroup
        Me.ButtonGenData = Me.Factory.CreateRibbonButton
        Me.SapBiHfm = Me.Factory.CreateRibbonGroup
        Me.ButtonUpdateHFMMulti = Me.Factory.CreateRibbonButton
        Me.SapBiHfmLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonSAPLogon = Me.Factory.CreateRibbonButton
        Me.ButtonSAPLogoff = Me.Factory.CreateRibbonButton
        Me.SapHfm.SuspendLayout()
        Me.Generate.SuspendLayout()
        Me.SapBiHfm.SuspendLayout()
        Me.SapBiHfmLogon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapHfm
        '
        Me.SapHfm.Groups.Add(Me.Generate)
        Me.SapHfm.Groups.Add(Me.SapBiHfm)
        Me.SapHfm.Groups.Add(Me.SapBiHfmLogon)
        Me.SapHfm.Label = "SAP BI HFM"
        Me.SapHfm.Name = "SapHfm"
        '
        'Generate
        '
        Me.Generate.Items.Add(Me.ButtonGenData)
        Me.Generate.Label = "Generate"
        Me.Generate.Name = "Generate"
        '
        'ButtonGenData
        '
        Me.ButtonGenData.Image = CType(resources.GetObject("ButtonGenData.Image"), System.Drawing.Image)
        Me.ButtonGenData.Label = "Generate Data"
        Me.ButtonGenData.Name = "ButtonGenData"
        Me.ButtonGenData.ScreenTip = "Generate the Output Data"
        Me.ButtonGenData.ShowImage = True
        '
        'SapBiHfm
        '
        Me.SapBiHfm.Items.Add(Me.ButtonUpdateHFMMulti)
        Me.SapBiHfm.Label = "Sap BI HFM"
        Me.SapBiHfm.Name = "SapBiHfm"
        '
        'ButtonUpdateHFMMulti
        '
        Me.ButtonUpdateHFMMulti.Image = CType(resources.GetObject("ButtonUpdateHFMMulti.Image"), System.Drawing.Image)
        Me.ButtonUpdateHFMMulti.Label = "Update Account Mapping"
        Me.ButtonUpdateHFMMulti.Name = "ButtonUpdateHFMMulti"
        Me.ButtonUpdateHFMMulti.ShowImage = True
        '
        'SapBiHfmLogon
        '
        Me.SapBiHfmLogon.Items.Add(Me.ButtonSAPLogon)
        Me.SapBiHfmLogon.Items.Add(Me.ButtonSAPLogoff)
        Me.SapBiHfmLogon.Label = "Logon"
        Me.SapBiHfmLogon.Name = "SapBiHfmLogon"
        '
        'ButtonSAPLogon
        '
        Me.ButtonSAPLogon.Image = CType(resources.GetObject("ButtonSAPLogon.Image"), System.Drawing.Image)
        Me.ButtonSAPLogon.Label = "SAP Logon"
        Me.ButtonSAPLogon.Name = "ButtonSAPLogon"
        Me.ButtonSAPLogon.ShowImage = True
        '
        'ButtonSAPLogoff
        '
        Me.ButtonSAPLogoff.Image = CType(resources.GetObject("ButtonSAPLogoff.Image"), System.Drawing.Image)
        Me.ButtonSAPLogoff.Label = "SAP Logoff"
        Me.ButtonSAPLogoff.Name = "ButtonSAPLogoff"
        Me.ButtonSAPLogoff.ShowImage = True
        '
        'SapHFMRibbon
        '
        Me.Name = "SapHFMRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapHfm)
        Me.SapHfm.ResumeLayout(False)
        Me.SapHfm.PerformLayout()
        Me.Generate.ResumeLayout(False)
        Me.Generate.PerformLayout()
        Me.SapBiHfm.ResumeLayout(False)
        Me.SapBiHfm.PerformLayout()
        Me.SapBiHfmLogon.ResumeLayout(False)
        Me.SapBiHfmLogon.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapHfm As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SapBiHfm As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSAPLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SapBiHfmLogon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonSAPLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonUpdateHFMMulti As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Generate As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonGenData As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As SapHFMRibbon
        Get
            Return Me.GetRibbon(Of SapHFMRibbon)()
        End Get
    End Property
End Class
