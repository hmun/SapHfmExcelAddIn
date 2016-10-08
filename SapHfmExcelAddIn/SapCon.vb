' Copyright 2016 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector
Public Class SapCon
    Const aParamWs As String = "Parameter"
    Private aSapExcelDestinationConfiguration As SapExcelDestinationConfiguration
    Private aDest As String
    Private destination As RfcCustomDestination

    Public Sub New()
        Dim parameters As New RfcConfigParameters()

        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        aWB = Globals.SapHFMAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets(aParamWs)
        Catch Exc As System.Exception
            MsgBox("No " & aParamWs & " Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM - SapCon")
            Exit Sub
        End Try
        aDest = aPws.Cells(5, 2).Value
        aSapExcelDestinationConfiguration = New SapExcelDestinationConfiguration
        aSapExcelDestinationConfiguration.ExcelAddOrChangeDestination(aParamWs)
        aSapExcelDestinationConfiguration.SetUp()
    End Sub

    Public Function checkCon() As Integer
        Dim dest As RfcDestination
        If destination Is Nothing Then
            dest = RfcDestinationManager.GetDestination(aDest)
            destination = dest.CreateCustomDestination()
        End If
        If destination.User = "" Then
            Dim oForm As New FormLogon
            Dim aClient As String
            Dim aUserName As String
            Dim aPassword As String
            Dim aLanguage As String
            Dim aRet As VariantType
            If Not destination.Client Is Nothing Then
                oForm.Client.Text = destination.Client
            End If
            If Not destination.Language Is Nothing Then
                oForm.Language.Text = destination.Language
            End If
            aRet = oForm.ShowDialog()
            If aRet = System.Windows.Forms.DialogResult.OK Then
                aClient = oForm.Client.Text
                aUserName = oForm.UserName.Text
                aPassword = oForm.Password.Text
                aLanguage = oForm.Language.Text
                setCredentials(aClient, aUserName, aPassword, aLanguage)
            End If
        End If
        Try
            destination.Ping()
            checkCon = True
        Catch ex As RfcInvalidParameterException
            clearCredentials()
            MsgBox("Connecting to SAP failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            checkCon = 4
        Catch ex As RfcBaseException
            clearCredentials()
            MsgBox("Connecting to SAP failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            checkCon = 8
        End Try
    End Function

    Public Sub setCredentials(aClient As String, aUsername As String, aPassword As String, aLanguage As String)
        destination.Client = aClient
        destination.User = aUsername
        destination.Password = aPassword
        destination.Language = aLanguage
    End Sub

    Public Sub SAPlogoff()
        destination = Nothing
    End Sub

    Public Sub clearCredentials()
        destination.User = ""
        destination.Password = Nothing
    End Sub

    Public Function getDestination() As RfcCustomDestination
        getDestination = destination
    End Function

End Class
