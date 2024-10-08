﻿' Copyright 2016 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/
Public Class SAPYPNUMItem
    Public CHRT_ACCTS As String
    Public YMPNUM As String
    Public YHFMACC As String
    Public YHFMCU1 As String
    Public YHFMCU2 As String
    Public YHFMCU3 As String
    Public YHFMICP As String
    Public YHFMSIGN As Integer
    Public ZOSUD4 As String


    Public Function create(pCHRT_ACCTS As String, pYMPNUM As String, pYHFMACC As String, pYHFMCU1 As String, pYHFMCU2 As String,
                       pYHFMCU3 As String, pYHFMICP As String, pYHFMSIGN As Integer, pZOSUD4 As String) As SAPYPNUMItem
        Dim aRef As New SAPYPNUMItem
        aRef.CHRT_ACCTS = pCHRT_ACCTS
        aRef.YMPNUM = pYMPNUM
        aRef.YHFMACC = pYHFMACC
        aRef.YHFMCU1 = If(pYHFMCU1 <> "#", pYHFMCU1, "")
        aRef.YHFMCU2 = If(pYHFMCU2 <> "#", pYHFMCU2, "")
        aRef.YHFMCU3 = If(pYHFMCU3 <> "#", pYHFMCU3, "")
        aRef.YHFMICP = If(pYHFMICP <> "#", pYHFMICP, "")
        aRef.YHFMSIGN = pYHFMSIGN
        aRef.ZOSUD4 = If(pZOSUD4 <> "#", pZOSUD4, "")
        create = aRef
    End Function

End Class
