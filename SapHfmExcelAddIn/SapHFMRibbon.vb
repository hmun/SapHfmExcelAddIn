' Copyright 2016 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon

Public Class SapHFMRibbon
    Const HFMRow = 5
    Const HFMCol = 9

    Private aSapCon As SapCon

    Private Sub ButtonTransferHFM_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonTransferHFM.Click
        Dim aPws As Excel.Worksheet
        Dim aTws As Excel.Worksheet
        Dim aDws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aRange As Excel.Range
        Dim aChartofAccounts As String
        Dim i As Integer
        Dim aSAPYPNUMItem As SAPYPNUMItem
        Dim aDict As New Dictionary(Of String, SAPYPNUMItem)
        Dim aHFMSign As Integer

        aWB = Globals.SapHFMAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            Exit Sub
        End Try
        aChartofAccounts = aPws.Cells(2, 2).Value
        Try
            aTws = aWB.Worksheets("Table")
        Catch Exc As System.Exception
            MsgBox("No Table Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            Exit Sub
        End Try
        Try
            aDws = aWB.Worksheets("Data")
        Catch Exc As System.Exception
            MsgBox("No Data Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            Exit Sub
        End Try

        i = HFMRow
        Do
            aSAPYPNUMItem = New SAPYPNUMItem
            If Not aDict.ContainsKey(aTws.Cells(i, HFMCol).Value) Then
                If aTws.Cells(i, HFMCol + 6).Value = "#" Then
                    aHFMSign = 0
                Else
                    aHFMSign = CInt(aTws.Cells(i, HFMCol + 6).Value)
                End If
                If aTws.Cells(i, HFMCol + 1).Value <> "#" And aTws.Cells(i, HFMCol + 2).Value <> "#" And aTws.Cells(i, HFMCol + 3).Value <> "#" _
                   And aTws.Cells(i, HFMCol + 4).Value <> "#" And aTws.Cells(i, HFMCol + 5).Value <> "#" Then
                    aSAPYPNUMItem = aSAPYPNUMItem.create(aChartofAccounts, aTws.Cells(i, HFMCol).Value, aTws.Cells(i, HFMCol + 1).Value,
                                                         aTws.Cells(i, HFMCol + 2).Value, aTws.Cells(i, HFMCol + 3).Value,
                                                         aTws.Cells(i, HFMCol + 4).Value, aTws.Cells(i, HFMCol + 5).Value, aHFMSign)
                    aDict.Add(aTws.Cells(i, HFMCol).Value, aSAPYPNUMItem)
                End If
            End If
            i = i + 1
        Loop While aTws.Cells(i, HFMCol).Value <> ""

        aDws.Activate()
        If aDws.Cells(2, 1).Value <> "" Then
            aRange = aDws.Range("A2")
            i = 2
            Do
                i = i + 1
            Loop While aDws.Cells(i, 1).Value <> ""
            aRange = aDws.Range(aRange, aDws.Cells(i, 1))
            aRange.EntireRow.Delete()
        End If

        i = 2
        For Each Item In aDict.Values
            aDws.Cells(i, 1) = Item.YMPNUM
            aDws.Cells(i, 2) = Item.YHFMACC
            aDws.Cells(i, 3) = Item.YHFMCU1
            aDws.Cells(i, 4) = Item.YHFMCU2
            aDws.Cells(i, 5) = Item.YHFMCU3
            aDws.Cells(i, 6) = Item.YHFMICP
            aDws.Cells(i, 7) = Item.YHFMSIGN
            i = i + 1
        Next

    End Sub

    Private Sub ButtonUpdateHFM_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonUpdateHFM.Click
        Dim aPws As Excel.Worksheet
        Dim aDws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aConRet As Integer
        Dim aChartofAccounts As String
        Dim aRet As Integer
        Dim i As Integer

        aWB = Globals.SapHFMAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            Exit Sub
        End Try
        aChartofAccounts = aPws.Cells(2, 2).Value
        Try
            aDws = aWB.Worksheets("Data")
        Catch Exc As System.Exception
            MsgBox("No Data Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            Exit Sub
        End Try

        If aSapCon Is Nothing Then
            aSapCon = New SapCon
        End If

        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
        End If

        Dim aSAP_ZFAGL_UPD_YMPNUM As SAP_ZFAGL_UPD_YMPNUM

        Try
            aSAP_ZFAGL_UPD_YMPNUM = New SAP_ZFAGL_UPD_YMPNUM(aSapCon)
        Catch ex As System.Exception
            MsgBox("Exception in SAP_ZFAGL_UPD_YMPNUM.update! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_ZFAGL_UPD_YMPNUM")
            Exit Sub
        End Try

        aDws.Activate()
        i = 2
        Do
            aRet = aSAP_ZFAGL_UPD_YMPNUM.update(aChartofAccounts, aDws.Cells(i, 1).Value, aDws.Cells(i, 2).Value, aDws.Cells(i, 3).Value,
                                                aDws.Cells(i, 4).Value, aDws.Cells(i, 5).Value, aDws.Cells(i, 6).Value, CInt(aDws.Cells(i, 7).Value))
            If aRet = 0 Then
                aDws.Cells(i, 8) = "OK"
            Else
                aDws.Cells(i, 8) = "Error: " & aRet
                Exit Sub
            End If
            i = i + 1
        Loop While aDws.Cells(i, 1).Value <> ""
        MsgBox("Update completed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "SAP BI HFM")

    End Sub

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
    End Sub

    Private Sub ButtonSAPLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSAPLogoff.Click
        If Not aSapCon Is Nothing Then
            aSapCon = New SapCon
        End If

    End Sub
End Class
