﻿Partial Class Ribbon1
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
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.SapBiHfm = Me.Factory.CreateRibbonGroup
        Me.ButtonTransferHFM = Me.Factory.CreateRibbonButton
        Me.ButtonUpdateHFM = Me.Factory.CreateRibbonButton
        Me.ButtonSAPLogoff = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.SapBiHfm.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.SapBiHfm)
        Me.Tab1.Label = "Sap BI HFM"
        Me.Tab1.Name = "Tab1"
        '
        'SapBiHfm
        '
        Me.SapBiHfm.Items.Add(Me.ButtonTransferHFM)
        Me.SapBiHfm.Items.Add(Me.ButtonUpdateHFM)
        Me.SapBiHfm.Items.Add(Me.ButtonSAPLogoff)
        Me.SapBiHfm.Label = "Sap BI HFM"
        Me.SapBiHfm.Name = "SapBiHfm"
        '
        'ButtonTransferHFM
        '
        Me.ButtonTransferHFM.Label = "Transfer HFM Mapping"
        Me.ButtonTransferHFM.Name = "ButtonTransferHFM"
        '
        'ButtonUpdateHFM
        '
        Me.ButtonUpdateHFM.Label = "Update HFM Mapping"
        Me.ButtonUpdateHFM.Name = "ButtonUpdateHFM"
        '
        'ButtonSAPLogoff
        '
        Me.ButtonSAPLogoff.Label = "SAP Logoff"
        Me.ButtonSAPLogoff.Name = "ButtonSAPLogoff"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.SapBiHfm.ResumeLayout(False)
        Me.SapBiHfm.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SapBiHfm As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonUpdateHFM As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonTransferHFM As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSAPLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
