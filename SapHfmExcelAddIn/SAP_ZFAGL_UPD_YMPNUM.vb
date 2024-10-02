' Copyright 2016 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector
Imports SAPCommon

Public Class SAP_ZFAGL_UPD_YMPNUM
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private oRfcFunctionMf As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private aStrucDic As Dictionary(Of String, RfcStructureMetadata)
    Private aParamDic As Dictionary(Of String, RfcParameterMetadata)

    Sub New(aSapCon As SapCon)
        Try
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            Try
                log.Debug("New - " & "creating Function ZFAGL_UPD_YMPNUM")
                oRfcFunction = destination.Repository.CreateFunction("ZFAGL_UPD_YMPNUM")
                log.Debug("New - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
            Catch ex As System.Exception
                log.Debug("New - " & "creating Function ZFAGL_UPD_YMPNUM - Ignoring Exception=" & ex.ToString)
            End Try
            log.Debug("New - " & "creating Function ZFAGL_UPD_YMPNUMMF")
            oRfcFunctionMf = destination.Repository.CreateFunction("ZFAGL_UPD_YMPNUMMF")
            log.Debug("New - " & "oRfcFunctionMf.Metadata.Name=" & oRfcFunctionMf.Metadata.Name)
            aStrucDic = New Dictionary(Of String, RfcStructureMetadata)
            aParamDic = New Dictionary(Of String, RfcParameterMetadata)
            getMeta_update_mf(aParamDic, aStrucDic)
        Catch ex As System.Exception
            log.Error("New - Exception=" & ex.ToString)
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngActivityAlloc")
        End Try
    End Sub

    Private Sub addToStrucDic(pArrayName As String, pRfcStructureMetadata As RfcStructureMetadata, ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        If pStrucDic.ContainsKey(pArrayName) Then
            pStrucDic.Remove(pArrayName)
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Private Sub addToFieldDic(pArrayName As String, pRfcStructureMetadata As RfcParameterMetadata, ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata))
        If pFieldDic.ContainsKey(pArrayName) Then
            pFieldDic.Remove(pArrayName)
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Public Sub getMeta_update_mf(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {}
        Dim aImports As String() = {}
        Dim aTables As String() = {"T_XMPNUMMF"}
        Try
            log.Debug("getMeta_update - " & "creating Function ZFAGL_UPD_YMPNUM")
            oRfcFunction = destination.Repository.CreateFunction("ZFAGL_UPD_YMPNUM")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To oRfcFunction.Metadata.ParameterCount - 1
                If oRfcFunction.Metadata.Item(s).Direction = SAP.Middleware.Connector.RfcDirection.IMPORT Then
                    addToFieldDic("I|" & oRfcFunction.Metadata.Item(s).Name, oRfcFunction.Metadata.Item(s), pFieldDic)
                End If
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_update - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPBiHfm")
        Finally
            log.Debug("getMeta_update - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function update(pSourceSystem As String, pCHRT_ACCTS As String, pYMPNUM As String, pYHFMACC As String, pYHFMCU1 As String, pYHFMCU2 As String,
                           pYHFMCU3 As String, pYHFMICP As String, pYHFMSIGN As Integer, Optional pZOSUD4 As String = "") As Integer
        sapcon.checkCon()
        log.Debug("update - " & "BeginContext")
        RfcSessionManager.BeginContext(destination)
        Try
            log.Debug("update - " & "setting values")
            If pSourceSystem <> "" Then
                oRfcFunction.SetValue("I_SOURSYSTEM", pSourceSystem)
            End If
            oRfcFunction.SetValue("I_CHRT_ACCTS", pCHRT_ACCTS)
            oRfcFunction.SetValue("I_YMPNUM", pYMPNUM)
            oRfcFunction.SetValue("I_YHFMACC", pYHFMACC)
            oRfcFunction.SetValue("I_YHFMCU1", pYHFMCU1)
            oRfcFunction.SetValue("I_YHFMCU2", pYHFMCU2)
            oRfcFunction.SetValue("I_YHFMCU3", pYHFMCU3)
            oRfcFunction.SetValue("I_YHFMICP", pYHFMICP)
            oRfcFunction.SetValue("I_YHFMSIGN", pYHFMSIGN)
            If aParamDic.ContainsKey("I|" & "I_ZOSUD4") Then
                oRfcFunction.SetValue("I_ZOSUD4", pZOSUD4)
            End If
            log.Debug("update - " & "invoking " & oRfcFunction.Metadata.Name)
            oRfcFunction.Invoke(destination)
            update = oRfcFunction.GetValue("E_RETURN")
            log.Debug("update - " & "update=" & CStr(update))
        Catch ex As Exception
            log.Error("update - in SAP_ZFAGL_UPD_YMPNUM.update=" & ex.ToString)
            MsgBox("Exception in SAP_ZFAGL_UPD_YMPNUM.update! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            update = 4
        Finally
            log.Debug("update - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function update_mf(pSourceSystem As String, pCHRT_ACCTS As String, pYHFMMF As String, pYMPNUMMF As String, pYHFMACC As String, pYHFMCU1 As String, pYHFMCU2 As String,
                              pYHFMCU3 As String, pYHFMICP As String, pYHFMSIGN As Integer, Optional pZOSUD4 As String = "") As Integer
        sapcon.checkCon()
        log.Debug("update_mf - " & "BeginContext")
        RfcSessionManager.BeginContext(destination)
        Try
            log.Debug("update_mf - " & "setting values")
            If pSourceSystem <> "" Then
                oRfcFunctionMf.SetValue("I_SOURSYSTEM", pSourceSystem)
            End If
            oRfcFunctionMf.SetValue("I_CHRT_ACCTS", pCHRT_ACCTS)
            oRfcFunctionMf.SetValue("I_YHFMMF", pYHFMMF)
            oRfcFunctionMf.SetValue("I_YMPNUMMF", pYMPNUMMF)
            oRfcFunctionMf.SetValue("I_YHFMACC", pYHFMACC)
            oRfcFunctionMf.SetValue("I_YHFMCU1", pYHFMCU1)
            oRfcFunctionMf.SetValue("I_YHFMCU2", pYHFMCU2)
            oRfcFunctionMf.SetValue("I_YHFMCU3", pYHFMCU3)
            oRfcFunctionMf.SetValue("I_YHFMICP", pYHFMICP)
            oRfcFunctionMf.SetValue("I_YHFMSIGN", pYHFMSIGN)
            If aParamDic.ContainsKey("I|" & "I_ZOSUD4") Then
                oRfcFunction.SetValue("I_ZOSUD4", pZOSUD4)
            End If
            log.Debug("update_mf - " & "invoking " & oRfcFunctionMf.Metadata.Name)
            oRfcFunctionMf.Invoke(destination)
            update_mf = oRfcFunctionMf.GetValue("E_RETURN")
            log.Debug("update_mf - " & "update_mf=" & CStr(update_mf))
        Catch ex As Exception
            log.Error("update_mf - in SAP_ZFAGL_UPD_YMPNUM.update_mf=" & ex.ToString)
            MsgBox("Exception in SAP_ZFAGL_UPD_YMPNUM.update_mf! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            update_mf = 4
        Finally
            log.Debug("update_mf - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function update_mf_multi(pData As TData, ByRef pIndex As Integer) As Integer
        sapcon.checkCon()
        log.Debug("update_mf_multi - " & "BeginContext")
        RfcSessionManager.BeginContext(destination)
        Dim oT_XMPNUMMF As IRfcTable = oRfcFunctionMf.GetTable("T_XMPNUMMF")
        oT_XMPNUMMF.Clear()
        pIndex = 0
        Try
            log.Debug("update_mf_multi - " & "setting values")
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            Dim aTStrRec As SAPCommon.TStrRec
            For Each aKvP In pData.aTDataDic
                Dim oT_XMPNUMMFAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "T_XMPNUMMF"
                            If Not oT_XMPNUMMFAppended Then
                                oT_XMPNUMMF.Append()
                                oT_XMPNUMMFAppended = True
                            End If
                            oT_XMPNUMMF.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            log.Debug("update_mf_multi - " & "invoking " & oRfcFunctionMf.Metadata.Name)
            oRfcFunctionMf.Invoke(destination)
            update_mf_multi = oRfcFunctionMf.GetValue("E_RETURN")
            pIndex = oRfcFunctionMf.GetValue("E_LAST_INDEX")
            log.Debug("update_mf_multi - " & "update_mf=" & CStr(update_mf_multi))
        Catch ex As Exception
            log.Error("update_mf_multi - in SAP_ZFAGL_UPD_YMPNUM.update_mf=" & ex.ToString)
            MsgBox("Exception in SAP_ZFAGL_UPD_YMPNUM.update_mf! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            update_mf_multi = 4
        Finally
            log.Debug("update_mf_multi - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Sub RemoveFunction()
        destination.Repository.RemoveFunctionMetadata("ZFAGL_UPD_YMPNUM")
        destination.Repository.RemoveFunctionMetadata("ZFAGL_UPD_YMPNUMMF")
    End Sub

End Class
