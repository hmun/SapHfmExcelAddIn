' Copyright 2024 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ListView
Imports SAP.Middleware.Connector

Public Class SAP_ZFAGL_UPD_ZOSMSUD

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            log.Debug("New - " & "checking connection")
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            log.Error("New - Exception=" & ex.ToString)
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_ZFAGL_UPD_ZOSMSUD")
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

    Public Sub getMeta_Update(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {}
        Dim aImports As String() = {"I_UDNR"}
        Dim aTables As String() = {"T_ZOSMSUD"}
        Try
            log.Debug("getMeta_Update - " & "creating Function ZFAGL_UPD_ZOSMSUD")
            oRfcFunction = destination.Repository.CreateFunction("ZFAGL_UPD_ZOSMSUD")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
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
            log.Error("getMeta_Update - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_ZFAGL_UPD_ZOSMSUD")
        Finally
            log.Debug("getMeta_Update - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function Update(pData As TSAP_OsMdData, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        Update = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("ZFAGL_UPD_ZOSMSUD")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("E_T_MESSAGES")
            Dim oT_ZOSMSUD As IRfcTable = oRfcFunction.GetTable("T_ZOSMSUD")
            oT_ZOSMSUD.Clear()
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="T_ZOSMSUD", pIRfcTable:=oT_ZOSMSUD)
            ' call the BAPI
            oRfcFunction.Invoke(destination)

            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("MSGTY") <> "I" And oRETURN(i).GetValue("MSGTY") <> "W" Then
                    Update = Update & ";" & oRETURN(i).GetValue("MSGTXTP")
                    If oRETURN(i).GetValue("MSGTY") <> "S" And oRETURN(i).GetValue("MSGTY") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            Update = If(Update = "", pOKMsg, If(aErr = False, pOKMsg & Update, "Error" & Update))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_ZFAGL_UPD_ZOSMSUD")
            Update = "Error: Exception in Update"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
