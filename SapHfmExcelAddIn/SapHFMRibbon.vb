' Copyright 2016 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Interop.Excel
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

    Private Sub ButtonTransferHFM_Click(sender As Object, e As RibbonControlEventArgs) 
        Dim aPws As Excel.Worksheet
        Dim aTws As Excel.Worksheet
        Dim aDws As Excel.Worksheet
        Dim aEws As Excel.Worksheet
        Dim aAws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aRange As Excel.Range
        Dim aChartofAccounts As String
        Dim aMappingFamily As String
        Dim aCreateCDN_MONTHEND As String
        Dim aAllowEmpty As String
        Dim i As Integer
        Dim aSAPYPNUMItem As SAPYPNUMItem
        Dim aDict As New Dictionary(Of String, SAPYPNUMItem)
        Dim aHFMSign As Integer
        Dim aZOSUD4 As String
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
        aAllowEmpty = CStr(aPws.Cells(5, 2).Value)

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
                If CStr(aTws.Cells(HFMRow - 1, HFMCol + 7).Value) = "UD4" Then
                    aZOSUD4 = CStr(aTws.Cells(i, HFMCol + 7).Value)
                Else
                    aZOSUD4 = ""
                End If
                If CStr(aTws.Cells(i, HFMCol + 1).Value) <> "#" And
                    (String.IsNullOrEmpty(aAllowEmpty) Or
                    (CStr(aTws.Cells(i, HFMCol + 2).Value) <> "#" And CStr(aTws.Cells(i, HFMCol + 3).Value) <> "#" And CStr(aTws.Cells(i, HFMCol + 4).Value) <> "#" And CStr(aTws.Cells(i, HFMCol + 5).Value) <> "#")) Then
                    aSAPYPNUMItem = aSAPYPNUMItem.create(aChartofAccounts, CStr(aTws.Cells(i, HFMCol).Value), CStr(aTws.Cells(i, HFMCol + 1).Value),
                                                         CStr(aTws.Cells(i, HFMCol + 2).Value), CStr(aTws.Cells(i, HFMCol + 3).Value),
                                                         CStr(aTws.Cells(i, HFMCol + 4).Value), CStr(aTws.Cells(i, HFMCol + 5).Value), aHFMSign, aZOSUD4)
                    aDict.Add(aTws.Cells(i, HFMCol).Value, aSAPYPNUMItem)
                End If
            End If
            i = i + 1
        Loop While CStr(aTws.Cells(i, HFMCol).Value) <> ""

        ' Fill the HFM-Export if it exists
        Try
            aEws = aWB.Worksheets("HFM-Export")
            Dim aLastRow As Integer = aEws.Cells(aEws.Cells.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            If aLastRow >= 1 Then
                Dim aDelRange As Excel.Range = aEws.Range(aEws.Cells(1, 1), aEws.Cells(aLastRow, 1))
                Dim unused = aDelRange.EntireRow.Delete()
            End If
            i = 0
            Dim aOutArray((aDict.Values.Count + 1) * 5 - 1, 5) As Object
            ' ACCOUNT
            aOutArray(0, 0) = "ACCOUNT"
            aOutArray(0, 1) = "*"
            aOutArray(0, 2) = "*"
            aOutArray(0, 3) = "ZW"
            aOutArray(0, 4) = "Explicit"
            aOutArray(0, 5) = ""
            i = i + 1
            For Each Item In aDict.Values
                aOutArray(i - 1, 0) = "ACCOUNT"
                aOutArray(i - 1, 1) = CStr(Item.YMPNUM)
                If CInt(Item.YHFMSIGN) = 0 Then
                    aOutArray(i - 1, 2) = CStr(Item.YHFMACC)
                Else
                    aOutArray(i - 1, 2) = "-" & CStr(Item.YHFMACC)
                End If
                aOutArray(i - 1, 4) = "Explicit"
                aOutArray(i - 1, 5) = ""
                i = i + 1
            Next
            ' ICP
            aOutArray(i - 1, 0) = "ICP"
            aOutArray(i - 1, 1) = "*"
            aOutArray(i - 1, 2) = "*"
            aOutArray(i - 1, 3) = "ZW"
            aOutArray(i - 1, 4) = "Explicit"
            aOutArray(i - 1, 5) = ""
            i = i + 1
            For Each Item In aDict.Values
                aOutArray(i - 1, 0) = "ICP"
                aOutArray(i - 1, 1) = CStr(Item.YMPNUM)
                aOutArray(i - 1, 2) = aHFMConstHelper.getConstantICP(CStr(Item.YHFMACC), CStr(Item.YHFMICP))
                aOutArray(i - 1, 4) = "Explicit"
                aOutArray(i - 1, 5) = ""
                i = i + 1
            Next
            ' UD1
            aOutArray(i - 1, 0) = "UD1"
            aOutArray(i - 1, 1) = "*"
            aOutArray(i - 1, 2) = "*"
            aOutArray(i - 1, 3) = "ZW"
            aOutArray(i - 1, 4) = "Explicit"
            aOutArray(i - 1, 5) = ""
            i = i + 1
            For Each Item In aDict.Values
                aOutArray(i - 1, 0) = "UD1"
                aOutArray(i - 1, 1) = CStr(Item.YMPNUM)
                aOutArray(i - 1, 2) = aHFMConstHelper.getConstantUD1(CStr(Item.YHFMACC), CStr(Item.YHFMCU1))
                aOutArray(i - 1, 4) = "Explicit"
                aOutArray(i - 1, 5) = ""
                i = i + 1
            Next
            ' UD2
            aOutArray(i - 1, 0) = "UD2"
            aOutArray(i - 1, 1) = "*"
            aOutArray(i - 1, 2) = "*"
            aOutArray(i - 1, 3) = "ZW"
            aOutArray(i - 1, 4) = "Explicit"
            aOutArray(i - 1, 5) = ""
            i = i + 1
            For Each Item In aDict.Values
                aOutArray(i - 1, 0) = "UD2"
                aOutArray(i - 1, 1) = CStr(Item.YMPNUM)
                aOutArray(i - 1, 2) = aHFMConstHelper.getConstantUD2(CStr(Item.YHFMACC), CStr(Item.YHFMCU2))
                aOutArray(i - 1, 4) = "Explicit"
                aOutArray(i - 1, 5) = ""
                i = i + 1
            Next
            ' UD3
            aOutArray(i - 1, 0) = "UD3"
            aOutArray(i - 1, 1) = "*"
            aOutArray(i - 1, 2) = "*"
            aOutArray(i - 1, 3) = "ZW"
            aOutArray(i - 1, 4) = "Explicit"
            aOutArray(i - 1, 5) = ""
            i = i + 1
            For Each Item In aDict.Values
                aOutArray(i - 1, 0) = "UD3"
                aOutArray(i - 1, 1) = CStr(Item.YMPNUM)
                aOutArray(i - 1, 2) = aHFMConstHelper.getConstantUD3(CStr(Item.YHFMACC), CStr(Item.YHFMCU3))
                aOutArray(i - 1, 4) = "Explicit"
                aOutArray(i - 1, 5) = ""
                i = i + 1
            Next
            ' Output the Array
            Dim aOutRange = aEws.Range(aEws.Cells(1, 1), aEws.Cells(aOutArray.GetUpperBound(0) + 1, aOutArray.GetUpperBound(1) + 1))
            aOutRange.Value = aOutArray
            ' Add The lines from HFM-Additions
            i = aOutArray.GetUpperBound(0) + 1
            Try
                aAws = aWB.Worksheets("HFM-Additions")
                Dim aHFMAdditions As New Collection
                Dim l As Integer = 1
                Do While Not String.IsNullOrEmpty(aAws.Cells(l, 1).Value)
                    Dim aArray(4) As String
                    aArray(0) = CStr(aAws.Cells(l, 1).Value)
                    aArray(1) = CStr(aAws.Cells(l, 2).Value)
                    aArray(2) = CStr(aAws.Cells(l, 3).Value)
                    aArray(3) = CStr(aAws.Cells(l, 4).Value)
                    aArray(4) = CStr(aAws.Cells(l, 5).Value)
                    l += 1
                    aHFMAdditions.Add(aArray)
                Loop
                Dim aArr As Object
                For Each aArr In aHFMAdditions
                    Dim aR As Excel.Range
                    aR = aEws.Range(aEws.Cells(i, 1), aEws.Cells(i, 5))
                    aR.Value = aArr
                    i += 1
                Next
            Catch Exc As System.Exception
            End Try
        Catch Exc As System.Exception
        End Try


        'aDws.Activate()
        'If aDws.Cells(2, 1).Value <> "" Then
        '    aRange = aDws.Range("A2")
        '    i = 2
        '    Do
        '        i = i + 1
        '    Loop While aDws.Cells(i, 1).Value <> ""
        '    aRange = aDws.Range(aRange, aDws.Cells(i, 1))
        '    aRange.EntireRow.Delete()
        'End If

        'Dim aCells As Excel.Range
        'i = 2
        'For Each Item In aDict.Values
        '    aCells = aDws.Range(aDws.Cells(i, 1), aDws.Cells(i, 8))
        '    aCells.NumberFormat = "@"
        '    aDws.Cells(i, 1).Value2 = Item.YMPNUM
        '    aDws.Cells(i, 2).Value2 = Item.YHFMACC
        '    aDws.Cells(i, 3).Value2 = Item.YHFMCU1
        '    aDws.Cells(i, 4).Value2 = Item.YHFMCU2
        '    aDws.Cells(i, 5).Value2 = Item.YHFMCU3
        '    aDws.Cells(i, 6).Value2 = Item.YHFMICP
        '    aDws.Cells(i, 7).Value2 = Item.YHFMSIGN
        '    If CStr(aDws.Cells(1, 8).Value) = "ZOSUD4" Then
        '        aDws.Cells(i, 8).Value2 = Item.ZOSUD4
        '    End If
        '    i = i + 1
        'Next
        Globals.SapHFMAddIn.Application.EnableEvents = True
        Globals.SapHFMAddIn.Application.ScreenUpdating = True
        Globals.SapHFMAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
    End Sub

    Private Sub ButtonUpdateHFM_Click(sender As Object, e As RibbonControlEventArgs) 
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
            Dim aHasUD4 As Boolean = If(CStr(aDws.Cells(1, 8).Value) = "ZOSUD4", True, False)
            Dim aMsgCol As Integer = If(aHasUD4, 9, 8)
            i = 2
            Globals.SapHFMAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapHFMAddIn.Application.EnableEvents = False
            Globals.SapHFMAddIn.Application.ScreenUpdating = False
            Do
                log.Debug("ButtonUpdateHFM_Click - " & "calling aSAP_ZFAGL_UPD_YMPNUM.update_mf")
                Dim aUD4 As String = If(aHasUD4, CStr(aDws.Cells(i, 8).Value), "")
                aRet = aSAP_ZFAGL_UPD_YMPNUM.update_mf(aSourceSystem, aChartofAccounts, aMappingFamily, CStr(aDws.Cells(i, 1).Value), CStr(aDws.Cells(i, 2).Value), CStr(aDws.Cells(i, 3).Value),
                                                       CStr(aDws.Cells(i, 4).Value), CStr(aDws.Cells(i, 5).Value), CStr(aDws.Cells(i, 6).Value), CInt(aDws.Cells(i, 7).Value), aUD4)
                If aRet = 0 Then
                    aDws.Cells(i, aMsgCol) = "OK"
                Else
                    aDws.Cells(i, aMsgCol) = "Error: " & aRet
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

    Private Sub ButtonUpdateHFMMulti_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonUpdateHFMMulti.Click
        Dim aPws As Excel.Worksheet
        Dim aDws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aChartofAccounts As String
        Dim aMappingFamily As String
        Dim aSourceSystem As String
        Dim aRet As Integer
        Dim i As Integer
        Dim aIntPar As New SAPCommon.TStr
        Dim aItems As New TData(aIntPar)
        Dim aKey As String
        Dim aIndex As Integer

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
            log.Debug("ButtonUpdateHFMMulti_Click - " & "creating aSAP_ZFAGL_UPD_YMPNUM")
            aSAP_ZFAGL_UPD_YMPNUM = New SAP_ZFAGL_UPD_YMPNUM(aSapCon)
        Catch ex As System.Exception
            MsgBox("Exception in SAP_ZFAGL_UPD_YMPNUM.update! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_ZFAGL_UPD_YMPNUM")
            Exit Sub
        End Try

        log.Debug("ButtonUpdateHFMMulti_Click - " & "processing data - disabling events, screen update, cursor")
        Try
            aDws.Activate()
            i = 2
            Dim aHasUD4 As Boolean = If(CStr(aDws.Cells(1, 8).Value) = "ZOSUD4", True, False)
            Dim aMsgCol As Integer = If(aHasUD4, 9, 8)
            Globals.SapHFMAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapHFMAddIn.Application.EnableEvents = False
            Globals.SapHFMAddIn.Application.ScreenUpdating = False
            Do
                aKey = CStr(i)
                aItems.addValue(aKey, "T_XMPNUMMF-/BIC/XCHRT_ACC", aChartofAccounts, "", "", pEmty:=False, pEmptyChar:="")
                aItems.addValue(aKey, "T_XMPNUMMF-/BIC/XHFMMF", aMappingFamily, "", "", pEmty:=False, pEmptyChar:="")
                aItems.addValue(aKey, "T_XMPNUMMF-SOURSYSTEM", aSourceSystem, "", "", pEmty:=False, pEmptyChar:="")
                aItems.addValue(aKey, "T_XMPNUMMF-/BIC/XMPNUMMF", CStr(aDws.Cells(i, 1).value), "", "", pEmty:=False, pEmptyChar:="")
                aItems.addValue(aKey, "T_XMPNUMMF-/BIC/XHFMACC", CStr(aDws.Cells(i, 2).value), "", "", pEmty:=False, pEmptyChar:="")
                aItems.addValue(aKey, "T_XMPNUMMF-/BIC/XHFMCU1", CStr(aDws.Cells(i, 3).value), "", "", pEmty:=False, pEmptyChar:="")
                aItems.addValue(aKey, "T_XMPNUMMF-/BIC/XHFMCU2", CStr(aDws.Cells(i, 4).value), "", "", pEmty:=False, pEmptyChar:="")
                aItems.addValue(aKey, "T_XMPNUMMF-/BIC/XHFMCU3", CStr(aDws.Cells(i, 5).value), "", "", pEmty:=False, pEmptyChar:="")
                aItems.addValue(aKey, "T_XMPNUMMF-/BIC/XHFMICP", CStr(aDws.Cells(i, 6).value), "", "", pEmty:=False, pEmptyChar:="")
                aItems.addValue(aKey, "T_XMPNUMMF-/BIC/XHFMSIGN", CStr(CInt(aDws.Cells(i, 7).value)), "", "", pEmty:=False, pEmptyChar:="")
                If aHasUD4 Then
                    aItems.addValue(aKey, "T_XMPNUMMF-/BIC/XOSUD4", CStr(aDws.Cells(i, 8).value), "", "", pEmty:=False, pEmptyChar:="")
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""

            log.Debug("ButtonUpdateHFMMulti_Click - " & "calling aSAP_ZFAGL_UPD_YMPNUM.update_mf")
            aRet = aSAP_ZFAGL_UPD_YMPNUM.update_mf_multi(aItems, aIndex)
            If aRet = 0 Then
                aDws.Cells(i, aMsgCol) = "OK, LastLine = " & aIndex
            Else
                aDws.Cells(i, aMsgCol) = "Error: " & aRet & ", LastLine = " & aIndex
                Exit Sub
            End If

            log.Debug("ButtonUpdateHFMMulti_Click - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapHFMAddIn.Application.EnableEvents = True
            Globals.SapHFMAddIn.Application.ScreenUpdating = True
            Globals.SapHFMAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Update completed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "SAP BI HFM")
        Catch ex As System.Exception
            Globals.SapHFMAddIn.Application.EnableEvents = True
            Globals.SapHFMAddIn.Application.ScreenUpdating = True
            Globals.SapHFMAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Update failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            log.Error("ButtonUpdateHFMMulti_Click - " & "Exception=" & ex.ToString)
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

    Private Sub ButtonGenData_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonGenData.Click
        Dim aSapHFMRibbon_Gen As New SapHFMRibbon_Gen
        aSapHFMRibbon_Gen.GenerateData()
    End Sub

End Class
