' Copyright 2016 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SapExcelDestinationConfiguration

    Private Shared inMemoryDestinationConfiguration As New SapInMemoryDestinationConfiguration()

    Public Shared Sub SetUp()
        '' register the in-memory destination configuration -- called before executing any of the examples
        RfcDestinationManager.RegisterDestinationConfiguration(inMemoryDestinationConfiguration)
    End Sub

    Public Shared Sub TearDown()
        '' unregister the in-memory destination configuration -- called after we are done working with the examples 
        RfcDestinationManager.UnregisterDestinationConfiguration(inMemoryDestinationConfiguration)
    End Sub

    Public Shared Sub ExcelAddOrChangeDestination(pWSname As String)
        Dim parameters As New RfcConfigParameters()

        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        aWB = Globals.SapHFMAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets(pWSname)
        Catch Exc As System.Exception
            MsgBox("No " & pWSname & " Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            Exit Sub
        End Try

        parameters(RfcConfigParameters.Name) = aPws.Cells(5, 2).value
        parameters(RfcConfigParameters.Language) = aPws.Cells(6, 2).value
        parameters(RfcConfigParameters.PeakConnectionsLimit) = "5"
        parameters(RfcConfigParameters.ConnectionIdleTimeout) = "600" '' 600 seconds, i.e. 10 minutes
        If aPws.Cells(7, 2).value <> "" Then
            parameters(RfcConfigParameters.AppServerHost) = aPws.Cells(7, 2).value
            parameters(RfcConfigParameters.SystemNumber) = CInt(aPws.Cells(8, 2).value)
        ElseIf aPws.Cells(10, 2).value <> "" Then
            parameters(RfcConfigParameters.MessageServerHost) = aPws.Cells(10, 2).value
            parameters(RfcConfigParameters.LogonGroup) = aPws.Cells(11, 2).value
        End If
        parameters(RfcConfigParameters.SystemID) = aPws.Cells(9, 2).value
        If aPws.Cells(12, 2).value <> "" Then
            parameters(RfcConfigParameters.Trace) = aPws.Cells(12, 2).value
        End If
        If aPws.Cells(13, 2).value <> "" Then
            parameters(RfcConfigParameters.Client) = aPws.Cells(13, 2).value
        End If
        If aPws.Cells(14, 2).value <> "" Then
            parameters(RfcConfigParameters.Language) = aPws.Cells(14, 2).value
        End If
        inMemoryDestinationConfiguration.AddOrEditDestination(parameters)
    End Sub
End Class
