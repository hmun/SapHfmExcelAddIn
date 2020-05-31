' Copyright 2016 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon

Public Class SapHFMRibbon
    Const HFMRow = 5
    Const HFMCol = 9

    Private aSapCon As SapCon
    Private aSapGeneral
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Sub SapHFMRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        aSapGeneral = New SapGeneral
    End Sub

    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        checkCon = False
        log.Debug("checkCon - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        log.Debug("checkCon - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            Try
                aSapCon = New SapCon()
            Catch ex As SystemException
                log.Warn("checkCon-New SapCon - )" & ex.ToString)
            End Try
        End If
        Try
            aSapConRet = aSapCon.checkCon()
        Catch ex As SystemException
            log.Warn("checkCon-aSapCon.checkCon - )" & ex.ToString)
        End Try
        If aSapConRet = 0 Then
            log.Debug("checkCon - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("checkCon - )" & ex.ToString)
            End Try
            log.Debug("checkCon - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("checkCon - " & "checkCon = True")
                checkCon = True
            Else
                log.Debug("checkCon - " & "connection check failed")
            End If
        End If
    End Function

    Private Sub ButtonTransferHFM_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonTransferHFM.Click
        Dim aPws As Excel.Worksheet
        Dim aTws As Excel.Worksheet
        Dim aDws As Excel.Worksheet
        Dim aEws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aRange As Excel.Range
        Dim aChartofAccounts As String
        Dim aMappingFamily As String
        Dim aCreateCDN_MONTHEND As String
        Dim i As Integer
        Dim aSAPYPNUMItem As SAPYPNUMItem
        Dim aDict As New Dictionary(Of String, SAPYPNUMItem)
        Dim aHFMSign As Integer
        Dim aHFMConstHelper As HFMConstHelper

        aWB = Globals.SapHFMAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            Exit Sub
        End Try
        aChartofAccounts = CStr(aPws.Cells(2, 2).Value)
        aMappingFamily = CStr(aPws.Cells(3, 2).Value)
        aCreateCDN_MONTHEND = CStr(aPws.Cells(4, 2).Value)

        aHFMConstHelper = New HFMConstHelper

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

        Globals.SapHFMAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapHFMAddIn.Application.EnableEvents = False
        Globals.SapHFMAddIn.Application.ScreenUpdating = False
        i = HFMRow
        Do
            aSAPYPNUMItem = New SAPYPNUMItem
            If Not aDict.ContainsKey(aTws.Cells(i, HFMCol).Value) Then
                If CStr(aTws.Cells(i, HFMCol + 6).Value) = "#" Then
                    aHFMSign = 0
                Else
                    aHFMSign = CInt(aTws.Cells(i, HFMCol + 6).Value)
                End If
                If CStr(aTws.Cells(i, HFMCol + 1).Value) <> "#" And CStr(aTws.Cells(i, HFMCol + 2).Value) <> "#" And CStr(aTws.Cells(i, HFMCol + 3).Value) <> "#" _
                   And CStr(aTws.Cells(i, HFMCol + 4).Value) <> "#" And CStr(aTws.Cells(i, HFMCol + 5).Value) <> "#" Then
                    aSAPYPNUMItem = aSAPYPNUMItem.create(aChartofAccounts, CStr(aTws.Cells(i, HFMCol).Value), CStr(aTws.Cells(i, HFMCol + 1).Value),
                                                         CStr(aTws.Cells(i, HFMCol + 2).Value), CStr(aTws.Cells(i, HFMCol + 3).Value),
                                                         CStr(aTws.Cells(i, HFMCol + 4).Value), CStr(aTws.Cells(i, HFMCol + 5).Value), aHFMSign)
                    aDict.Add(aTws.Cells(i, HFMCol).Value, aSAPYPNUMItem)
                End If
            End If
            i = i + 1
        Loop While CStr(aTws.Cells(i, HFMCol).Value) <> ""

        ' Fill the HFM-Export if it exists
        Try
            aEws = aWB.Worksheets("HFM-Export")
            aEws.Activate()
            If aEws.Cells(2, 1).Value <> "" Then
                aRange = aEws.Range("A2")
                i = 2
                Do
                    i = i + 1
                Loop While aEws.Cells(i, 1).Value <> ""
                aRange = aEws.Range(aRange, aEws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            i = 1
            ' ACCOUNT
            aEws.Cells(i, 1) = "ACCOUNT"
            aEws.Cells(i, 2) = "*"
            aEws.Cells(i, 3) = "*"
            aEws.Cells(i, 4) = "ZW"
            aEws.Cells(i, 5) = "Explicit"
            aEws.Cells(i, 6) = ""
            i = i + 1
            For Each Item In aDict.Values
                aEws.Cells(i, 1) = "ACCOUNT"
                aEws.Cells(i, 2) = CStr(Item.YMPNUM)
                If CInt(Item.YHFMSIGN) = 0 Then
                    aEws.Cells(i, 3) = CStr(Item.YHFMACC)
                Else
                    aEws.Cells(i, 3) = "-" & CStr(Item.YHFMACC)
                End If
                aEws.Cells(i, 5) = "Explicit"
                aEws.Cells(i, 6) = ""
                i = i + 1
            Next
            ' ICP
            aEws.Cells(i, 1) = "ICP"
            aEws.Cells(i, 2) = "*"
            aEws.Cells(i, 3) = "*"
            aEws.Cells(i, 4) = "ZW"
            aEws.Cells(i, 5) = "Explicit"
            aEws.Cells(i, 6) = ""
            i = i + 1
            For Each Item In aDict.Values
                aEws.Cells(i, 1) = "ICP"
                aEws.Cells(i, 2) = CStr(Item.YMPNUM)
                aEws.Cells(i, 3) = aHFMConstHelper.getConstantICP(CStr(Item.YHFMACC), CStr(Item.YHFMICP))
                aEws.Cells(i, 5) = "Explicit"
                aEws.Cells(i, 6) = ""
                i = i + 1
            Next
            ' UD1
            aEws.Cells(i, 1) = "UD1"
            aEws.Cells(i, 2) = "*"
            aEws.Cells(i, 3) = "*"
            aEws.Cells(i, 4) = "ZW"
            aEws.Cells(i, 5) = "Explicit"
            aEws.Cells(i, 6) = ""
            i = i + 1
            For Each Item In aDict.Values
                aEws.Cells(i, 1) = "UD1"
                aEws.Cells(i, 2) = CStr(Item.YMPNUM)
                aEws.Cells(i, 3) = aHFMConstHelper.getConstantUD1(CStr(Item.YHFMACC), CStr(Item.YHFMCU1))
                aEws.Cells(i, 5) = "Explicit"
                aEws.Cells(i, 6) = ""
                i = i + 1
            Next
            ' UD2
            aEws.Cells(i, 1) = "UD2"
            aEws.Cells(i, 2) = "*"
            aEws.Cells(i, 3) = "*"
            aEws.Cells(i, 4) = "ZW"
            aEws.Cells(i, 5) = "Explicit"
            aEws.Cells(i, 6) = ""
            i = i + 1
            For Each Item In aDict.Values
                aEws.Cells(i, 1) = "UD2"
                aEws.Cells(i, 2) = CStr(Item.YMPNUM)
                aEws.Cells(i, 3) = aHFMConstHelper.getConstantUD2(CStr(Item.YHFMACC), CStr(Item.YHFMCU2))
                aEws.Cells(i, 5) = "Explicit"
                aEws.Cells(i, 6) = ""
                i = i + 1
            Next
            ' UD3
            aEws.Cells(i, 1) = "UD3"
            aEws.Cells(i, 2) = "*"
            aEws.Cells(i, 3) = "*"
            aEws.Cells(i, 4) = "ZW"
            aEws.Cells(i, 5) = "Explicit"
            aEws.Cells(i, 6) = ""
            i = i + 1
            For Each Item In aDict.Values
                aEws.Cells(i, 1) = "UD3"
                aEws.Cells(i, 2) = CStr(Item.YMPNUM)
                aEws.Cells(i, 3) = aHFMConstHelper.getConstantUD3(CStr(Item.YHFMACC), CStr(Item.YHFMCU3))
                aEws.Cells(i, 5) = "Explicit"
                aEws.Cells(i, 6) = ""
                i = i + 1
            Next
            ' UD4
            aEws.Cells(i, 1) = "UD4"
            aEws.Cells(i, 2) = "*"
            aEws.Cells(i, 3) = "Input"
            aEws.Cells(i, 4) = "ZW"
            aEws.Cells(i, 5) = "Explicit"
            aEws.Cells(i, 6) = ""
            i = i + 1
            If aCreateCDN_MONTHEND = "X" Then
                aEws.Cells(i, 1) = "UD4"
                aEws.Cells(i, 2) = "CDN_MONTHEND"
                aEws.Cells(i, 3) = "CDN_MONTHEND"
                aEws.Cells(i, 4) = ""
                aEws.Cells(i, 5) = "Explicit"
                aEws.Cells(i, 6) = ""
            End If
        Catch Exc As System.Exception
        End Try

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

        Dim aCells As Excel.Range
        i = 2
        For Each Item In aDict.Values
            aCells = aDws.Range(aDws.Cells(i, 1), aDws.Cells(i, 8))
            aCells.NumberFormat = "@"
            aDws.Cells(i, 1).Value2 = Item.YMPNUM
            aDws.Cells(i, 2).Value2 = Item.YHFMACC
            aDws.Cells(i, 3).Value2 = Item.YHFMCU1
            aDws.Cells(i, 4).Value2 = Item.YHFMCU2
            aDws.Cells(i, 5).Value2 = Item.YHFMCU3
            aDws.Cells(i, 6).Value2 = Item.YHFMICP
            aDws.Cells(i, 7).Value2 = Item.YHFMSIGN
            i = i + 1
        Next
        Globals.SapHFMAddIn.Application.EnableEvents = True
        Globals.SapHFMAddIn.Application.ScreenUpdating = True
        Globals.SapHFMAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
    End Sub

    Private Sub ButtonUpdateHFM_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonUpdateHFM.Click
        Dim aPws As Excel.Worksheet
        Dim aDws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aChartofAccounts As String
        Dim aMappingFamily As String
        Dim aSourceSystem As String
        Dim aRet As Integer
        Dim i As Integer

        aWB = Globals.SapHFMAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            Exit Sub
        End Try
        aChartofAccounts = CStr(aPws.Cells(2, 2).Value)
        aMappingFamily = CStr(aPws.Cells(3, 2).Value)
        aSourceSystem = CStr(aPws.Cells(5, 2).Value)
        Try
            aDws = aWB.Worksheets("Data")
        Catch Exc As System.Exception
            MsgBox("No Data Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            Exit Sub
        End Try

        If checkCon() = False Then
            Exit Sub
        End If

        Dim aSAP_ZFAGL_UPD_YMPNUM As SAP_ZFAGL_UPD_YMPNUM

        Try
            log.Debug("ButtonUpdateHFM_Click - " & "creating aSAP_ZFAGL_UPD_YMPNUM")
            aSAP_ZFAGL_UPD_YMPNUM = New SAP_ZFAGL_UPD_YMPNUM(aSapCon)
        Catch ex As System.Exception
            MsgBox("Exception in SAP_ZFAGL_UPD_YMPNUM.update! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_ZFAGL_UPD_YMPNUM")
            Exit Sub
        End Try

        log.Debug("ButtonUpdateHFM_Click - " & "processing data - disabling events, screen update, cursor")
        Try
            aDws.Activate()
            i = 2
            Globals.SapHFMAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapHFMAddIn.Application.EnableEvents = False
            Globals.SapHFMAddIn.Application.ScreenUpdating = False
            Do
                log.Debug("ButtonUpdateHFM_Click - " & "calling aSAP_ZFAGL_UPD_YMPNUM.update_mf")
                aRet = aSAP_ZFAGL_UPD_YMPNUM.update_mf(aSourceSystem, aChartofAccounts, aMappingFamily, CStr(aDws.Cells(i, 1).Value), CStr(aDws.Cells(i, 2).Value), CStr(aDws.Cells(i, 3).Value),
                                                       CStr(aDws.Cells(i, 4).Value), CStr(aDws.Cells(i, 5).Value), CStr(aDws.Cells(i, 6).Value), CInt(aDws.Cells(i, 7).Value))
                If aRet = 0 Then
                    aDws.Cells(i, 8) = "OK"
                Else
                    aDws.Cells(i, 8) = "Error: " & aRet
                    Exit Sub
                End If
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""
            log.Debug("ButtonUpdateHFM_Click - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapHFMAddIn.Application.EnableEvents = True
            Globals.SapHFMAddIn.Application.ScreenUpdating = True
            Globals.SapHFMAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Update completed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "SAP BI HFM")
        Catch ex As System.Exception
            Globals.SapHFMAddIn.Application.EnableEvents = True
            Globals.SapHFMAddIn.Application.ScreenUpdating = True
            Globals.SapHFMAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Update failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            log.Error("ButtonUpdateHFM_Click - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Private Sub ButtonSAPLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSAPLogoff.Click
        log.Debug("ButtonLogoff_Click - " & "starting logoff")
        If Not aSapCon Is Nothing Then
            log.Debug("ButtonLogoff_Click - " & "calling aSapCon.SAPlogoff()")
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
        log.Debug("ButtonLogoff_Click - " & "exit")
    End Sub

    Private Sub ButtonSAPLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSAPLogon.Click
        Dim aConRet As Integer

        log.Debug("ButtonLogon_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            log.Debug("ButtonLogon_Click - " & "Version check failed")
            Exit Sub
        End If
        log.Debug("ButtonLogon_Click - " & "creating SapCon")
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        log.Debug("ButtonLogon_Click - " & "calling SapCon.checkCon()")
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            log.Debug("ButtonLogon_Click - " & "connection successfull")
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap Accounting")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub
End Class
