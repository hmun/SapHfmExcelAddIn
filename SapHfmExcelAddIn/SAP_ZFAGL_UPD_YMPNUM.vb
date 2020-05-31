' Copyright 2016 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAP_ZFAGL_UPD_YMPNUM
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private oRfcFunctionMf As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

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
            Catch ex As System.Exception
                log.Error("New - Exception=" & ex.ToString)
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngActivityAlloc")
        End Try
    End Sub

    Public Function update(pSourceSystem As String, pCHRT_ACCTS As String, pYMPNUM As String, pYHFMACC As String, pYHFMCU1 As String, pYHFMCU2 As String,
                           pYHFMCU3 As String, pYHFMICP As String, pYHFMSIGN As Integer) As Integer
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
                              pYHFMCU3 As String, pYHFMICP As String, pYHFMSIGN As Integer) As Integer
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

    Public Sub RemoveFunction()
        destination.Repository.RemoveFunctionMetadata("ZFAGL_UPD_YMPNUM")
        destination.Repository.RemoveFunctionMetadata("ZFAGL_UPD_YMPNUMMF")
    End Sub

End Class
