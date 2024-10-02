' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Tools.Ribbon
Imports SAPCommon
Imports SAPLogon
Imports SC = SAPCommon

Public Class SapHFMRibbon_Gen

    Private app = Globals.SapHFMAddIn.Application

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Function getIntParameters(ByRef pIntPar As SC.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = app.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SAP MM Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap MM Md")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SC.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub GenerateData()
        Dim aIntPar As New SC.TStr
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If
        ' get the ruleset limits
        Dim aGenNrFrom As Integer = If(aIntPar.value("GEN", "RULESET_FROM") <> "", CInt(aIntPar.value("GEN", "RULESET_FROM")), 0)
        Dim aGenNrTo As Integer = If(aIntPar.value("GEN", "RULESET_TO") <> "", CInt(aIntPar.value("GEN", "RULESET_TO")), 0)
        Dim aGenNr As String = ""
        For i As Integer = aGenNrFrom To aGenNrTo
            Dim aNr As String = If(i = 0, "", CStr(i))
            GenerateData_exec(pIntPar:=aIntPar, pNr:=aNr)
        Next
        MsgBox("Generate completed", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Generation")
    End Sub


    Private Sub GenerateData_exec(ByRef pIntPar As SC.TStr, Optional pNr As String = "")
        Dim aMigHelperVsto As MigHelperVsto
        Dim aBWs As Excel.Worksheet
        Dim aOWs As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aUselocal As Boolean = False
        Dim i As UInt32
        ' get internal parameters
        Dim aOwsName As String = If(pIntPar.value("GEN" & pNr, "WS_DATA") <> "", pIntPar.value("GEN" & pNr, "WS_DATA"), "Data")
        Dim aBwsName As String = If(pIntPar.value("GEN" & pNr, "WS_BASE") <> "", pIntPar.value("GEN" & pNr, "WS_BASE"), "Base")
        Dim aDeleteData As String = If(pIntPar.value("GEN" & pNr, "DELETE_DATA") <> "", pIntPar.value("GEN" & pNr, "DELETE_DATA"), "X")
        Dim aGenDeleteData As Boolean = If(aDeleteData = "X", True, False)
        Dim aLOff As Integer = If(pIntPar.value("GEN" & pNr, "LOFF_DATA") <> "", CInt(pIntPar.value("GEN" & pNr, "LOFF_DATA")), 4)
        Dim aLOffBData As Integer = If(pIntPar.value("GEN" & pNr, "LOFF_BDATA") <> "", CInt(pIntPar.value("GEN" & pNr, "LOFF_BDATA")), 1)
        Dim aLOffBNames As Integer = If(pIntPar.value("GEN" & pNr, "LOFF_BNAMES") <> "", CInt(pIntPar.value("GEN" & pNr, "LOFF_BNAMES")), 0)
        Dim aLOffTNames As Integer = If(pIntPar.value("GEN" & pNr, "LOFF_TNAMES") <> "", CInt(pIntPar.value("GEN" & pNr, "LOFF_TNAMES")), aLOff - 1)
        Dim aLineOut As Integer = If(pIntPar.value("GEN" & pNr, "LINE_OUT") <> "", CInt(pIntPar.value("GEN" & pNr, "LINE_OUT")), 0)
        Dim aBaseColFrom As Integer = If(pIntPar.value("GEN" & pNr, "BASE_COLFROM") <> "", CInt(pIntPar.value("GEN" & pNr, "BASE_COLFROM")), 1)
        Dim aBaseColTo As Integer = If(pIntPar.value("GEN" & pNr, "BASE_COLTO") <> "", CInt(pIntPar.value("GEN" & pNr, "BASE_COLTO")), 100)
        log.Debug("GenerateData_exec - " & "Basis Sheet")
        aWB = app.ActiveWorkbook
        Dim aGenLocalRules As String = If(pIntPar.value("GEN", "LOCAL_RULES") <> "", CStr(pIntPar.value("GEN", "LOCAL_RULES")), "")
        If aGenLocalRules = "X" Then
            aUselocal = True
            log.Debug("GenerateData_exec - " & "aUselocal = True")
        End If
        Try
            aBWs = aWB.Worksheets("InvoiceData")
        Catch Ex As System.Exception
            Try
                aBWs = aWB.Worksheets(aBwsName)
            Catch Exc As System.Exception
                log.Warn("GenerateData_exec - " & "No InvoiceData or " & aBwsName & " in current workbook.")
                MsgBox("No InvoiceData Sheet or " & aBwsName & " in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
                Exit Sub
            End Try
        End Try
        Try
            aBWs = aWB.Worksheets(aBwsName)
        Catch Exc As System.Exception
            log.Warn("GenerateData_exec - " & "No InvoiceData or " & aBwsName & " in current workbook.")
            MsgBox("No InvoiceData Sheet or " & aBwsName & " in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            Exit Sub
        End Try
        If String.IsNullOrEmpty(CStr(aBWs.Cells(aLOffBData + 1, aBaseColFrom).Value)) Then
            MsgBox("Base data cell row=" & aLOffBData + 1 & ", column=" & aBaseColFrom & " is empty. Check if the current workbook contains data and your internal parameters are correct!",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            Exit Sub
        End If
        If String.IsNullOrEmpty(CStr(aBWs.Cells(aLOffBNames + 1, aBaseColFrom).Value)) Then
            MsgBox("Base data name cell row=" & aLOffBNames + 1 & ", column=" & aBaseColFrom & " is empty. Check if the current workbook contains data and your internal parameters are correct!",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            Exit Sub
        End If
        '        aBWs.Activate()
        aMigHelperVsto = New MigHelperVsto(pIntPar:=pIntPar, pNr:=pNr, pUselocal:=aUselocal)
        ' process the data
        Try
            log.Debug("ButtonGenGLData_Click - " & "processing data - disabling events, screen update, cursor")
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            app.EnableEvents = False
            app.ScreenUpdating = False
            ' read the base lines
            app.StatusBar = "Reading the base data"
            i = aLOffBData + 1
            Dim aLastRow As UInt64
            If Not String.IsNullOrEmpty(CStr(aBWs.Cells(aLOffBData + 2, aBaseColFrom).Value)) Then
                aLastRow = aBWs.Cells(aLOffBData, aBaseColFrom).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
            Else
                aLastRow = aLOffBData + 1
            End If
            Dim aNamRange As Excel.Range = aBWs.Range(aBWs.Cells(aLOffBNames + 1, aBaseColFrom), aBWs.Cells(aLOffBNames + 1, aBaseColTo))
            Dim aNamArray As Object(,) = CType(aNamRange.Value, Object(,))
            Dim aValRange As Excel.Range = aBWs.Range(aBWs.Cells(aLOffBData + 1, aBaseColFrom), aBWs.Cells(aLastRow, aBaseColTo))
            Dim aValArray As Object(,) = CType(aValRange.Value, Object(,))
            Dim aMigEngine As SC.MigEngine = New SC.MigEngine(aMigHelperVsto.mh, pIntPar, pNr)
            ' migrating data
            app.StatusBar = "Migrating Data"
            aMigEngine.migrate(aNamArray, aValArray)
            ' prepare the output
            app.StatusBar = "Preparing Output"
            Dim aColDelData As Integer = If(pIntPar.value("GEN" & pNr, "DATA_COLDEL") <> "", CInt(pIntPar.value("GEN" & pNr, "DATA_COLDEL")), 1)
            Try
                aOWs = aWB.Worksheets(aOwsName)
            Catch Exc As System.Exception
                app.EnableEvents = True
                app.ScreenUpdating = True
                app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                MsgBox("No " & aOwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Migration Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
                Exit Sub
            End Try
            i = aLOff + 1
            aLastRow = aOWs.Cells(aOWs.Cells.Rows.Count, aColDelData).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            If aGenDeleteData And aLastRow >= aLOff + 1 Then
                app.StatusBar = "Deleting existing " & aLastRow - (aLOff + 1) & " lines in " & aOwsName
                Dim aRange As Excel.Range = aOWs.Range(aOWs.Cells(aLOff + 1, 1), aOWs.Cells(aLastRow, 1))
                Dim unused = aRange.EntireRow.Delete()
            End If
            Dim jMax As Integer = 0
            Do
                jMax += 1
            Loop While CStr(aOWs.Cells(aLOff, jMax + 1).value) <> ""
            Dim aOutLine = If(aLineOut <> 0, aLineOut, If(aGenDeleteData, aLOff + 1, aLastRow + 1))
            Dim aKeyRange As Excel.Range = aOWs.Range(aOWs.Cells(aLOff, 1), aOWs.Cells(aLOff, jMax))
            Dim aKeyArray As Object(,) = CType(aKeyRange.Value, Object(,))
            Dim aValueColumns As New SC.TOutData
            Dim aFormulaColumns As New SC.TOutData
            ' convert to output columns
            app.StatusBar = "Converting Output Columns"
            aMigEngine.ToTOutData(aKeyArray, aValueColumns, aFormulaColumns)
            ' write output to target sheet
            app.StatusBar = "Writing Value Columns"
            aMigHelperVsto.writeOutData(aOWs, aValueColumns, aOutLine, pType:="V")
            app.StatusBar = "Writing Formula Columns"
            aMigHelperVsto.writeOutData(aOWs, aFormulaColumns, aOutLine, pType:="F")
            aValueColumns = Nothing
            aFormulaColumns = Nothing
            aMigEngine = Nothing
            aMigHelperVsto = Nothing
            app.StatusBar = "Migration completed"
            app.EnableEvents = True
            app.ScreenUpdating = True
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            app.EnableEvents = True
            app.ScreenUpdating = True
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("ButtonGenGLData_Click failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            log.Error("ButtonGenGLData_Click - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub
End Class
