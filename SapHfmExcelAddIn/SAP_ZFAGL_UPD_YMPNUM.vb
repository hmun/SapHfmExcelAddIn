' Copyright 2016 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAP_ZFAGL_UPD_YMPNUM
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        destination = aSapCon.getDestination()
        oRfcFunction = destination.Repository.CreateFunction("ZFAGL_UPD_YMPNUM")
    End Sub

    Public Function update(pCHRT_ACCTS As String, pYMPNUM As String, pYHFMACC As String, pYHFMCU1 As String, pYHFMCU2 As String,
                           pYHFMCU3 As String, pYHFMICP As String, pYHFMSIGN As Integer) As Integer
        sapcon.checkCon()
        RfcSessionManager.BeginContext(destination)
        Try
            oRfcFunction.SetValue("I_CHRT_ACCTS", pCHRT_ACCTS)
            oRfcFunction.SetValue("I_YMPNUM", pYMPNUM)
            oRfcFunction.SetValue("I_YHFMACC", pYHFMACC)
            oRfcFunction.SetValue("I_YHFMCU1", pYHFMCU1)
            oRfcFunction.SetValue("I_YHFMCU2", pYHFMCU2)
            oRfcFunction.SetValue("I_YHFMCU3", pYHFMCU3)
            oRfcFunction.SetValue("I_YHFMICP", pYHFMICP)
            oRfcFunction.SetValue("I_YHFMSIGN", pYHFMSIGN)
            oRfcFunction.Invoke(destination)
            update = oRfcFunction.GetValue("E_RETURN")
        Catch ex As Exception
            MsgBox("Exception in SAP_ZFAGL_UPD_YMPNUM.update! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP BI HFM")
            update = 4
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Sub RemoveFunction()
        destination.Repository.RemoveFunctionMetadata("ZFAGL_UPD_YMPNUM")
    End Sub

End Class
