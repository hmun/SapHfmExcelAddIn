' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class HFMConst
    Public Account As TField
    Public UD1 As TField
    Public UD2 As TField
    Public UD3 As TField
    Public ICP As TField

    Private sTField As TField

    Public Sub New()
        sTField = New TField
        Account = New TField
        UD1 = New TField
        UD2 = New TField
        UD3 = New TField
        ICP = New TField
    End Sub

    Public Function setValues(pAccount As String, pUD1 As String, pUD2 As String, pUD3 As String, pICP As String)
        Account = sTField.create("Account", CStr(pAccount))
        UD1 = sTField.create("UD1", CStr(pUD1))
        UD2 = sTField.create("UD2", CStr(pUD2))
        UD3 = sTField.create("UD3", CStr(pUD3))
        ICP = sTField.create("ICP", CStr(pICP))
    End Function

    Public Function setValue(pField As String, pValue As String)
        If pField = "Account" Then
            Account = sTField.create(pField, CStr(pValue))
        ElseIf pField = "UD1" Then
            UD1 = sTField.create(pField, CStr(pValue))
        ElseIf pField = "UD2" Then
            UD2 = sTField.create(pField, CStr(pValue))
        ElseIf pField = "UD3" Then
            UD3 = sTField.create(pField, CStr(pValue))
        ElseIf pField = "ICP" Then
            ICP = sTField.create(pField, CStr(pValue))
        End If
    End Function

    Public Function getKey() As String
        Dim aKey As String
        aKey = Account.Value
        getKey = aKey
    End Function

    Public Function getRKey() As String
        Dim aKey As String
        aKey = Account.Value
        getRKey = aKey
    End Function

End Class

Public Class HFMConstHelper
    Public aHFMConstCol As Collection
    Private sTField As TField

    Public Sub New()
        sTField = New TField
        aHFMConstCol = New Collection

        Dim aCws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer
        aWB = Globals.SapHFMAddIn.Application.ActiveWorkbook
        Try
            aCws = aWB.Worksheets("HFM-Const")
            i = 2
            While CStr(aCws.Cells(i, 1).Value) <> ""
                addHFMConst(CStr(aCws.Cells(i, 1).Value), CStr(aCws.Cells(i, 2).Value), CStr(aCws.Cells(i, 3).Value), CStr(aCws.Cells(i, 4).Value), CStr(aCws.Cells(i, 5).Value))
            End While
        Catch Exc As System.Exception
            Exit Sub
        End Try

    End Sub

    Public Function addHFMConst(pAccount As String, pUD1 As String, pUD2 As String, pUD3 As String, pICP As String)
        Dim aHFMConst As HFMConst
        Dim aKey As String
        aKey = pAccount
        If contains(aHFMConstCol, aKey, "obj") Then
            aHFMConst = aHFMConstCol(aKey)
            aHFMConst.setValues(pAccount, pUD1, pUD2, pUD3, pICP)
        Else
            aHFMConst = New HFMConst
            aHFMConst.setValues(pAccount, pUD1, pUD2, pUD3, pICP)
            aHFMConstCol.Add(aHFMConst, aKey)
        End If
    End Function

    Public Function addConValue(pAccount As String, pField As String, pValue As String)
        Dim aHFMConst As HFMConst
        Dim aKey As String
        aKey = pAccount
        If contains(aHFMConstCol, aKey, "obj") Then
            aHFMConst = aHFMConstCol(aKey)
            aHFMConst.setValue(pField, pValue)
        Else
            aHFMConst = New HFMConst
            aHFMConst.setValue("Account", pAccount)
            aHFMConst.setValue(pField, pValue)
            aHFMConstCol.Add(aHFMConst, aKey)
        End If
    End Function

    Public Function getConstant(pAccount As String) As HFMConst
        Dim aKey As String
        Dim aHFMConst As New HFMConst
        aKey = pAccount
        If contains(aHFMConstCol, aKey, "obj") Then
            aHFMConst = aHFMConstCol(aKey)
        Else
            aHFMConst = Nothing
        End If
        getConstant = aHFMConst
    End Function

    Public Function getConstantUD1(pAccount As String, pVal As String) As String
        Dim aKey As String
        Dim aHFMConst As HFMConst
        aKey = pAccount
        If contains(aHFMConstCol, aKey, "obj") Then
            aHFMConst = aHFMConstCol(aKey)
            If aHFMConst.UD1.Value <> "" Then
                getConstantUD1 = aHFMConst.UD1.Value
            Else
                getConstantUD1 = pVal
            End If
        Else
            getConstantUD1 = pVal
        End If
    End Function

    Public Function getConstantUD2(pAccount As String, pVal As String) As String
        Dim aKey As String
        Dim aHFMConst As HFMConst
        aKey = pAccount
        If contains(aHFMConstCol, aKey, "obj") Then
            aHFMConst = aHFMConstCol(aKey)
            If aHFMConst.UD2.Value <> "" Then
                getConstantUD2 = aHFMConst.UD2.Value
            Else
                getConstantUD2 = pVal
            End If
        Else
            getConstantUD2 = pVal
        End If
    End Function

    Public Function getConstantUD3(pAccount As String, pVal As String) As String
        Dim aKey As String
        Dim aHFMConst As HFMConst
        aKey = pAccount
        If contains(aHFMConstCol, aKey, "obj") Then
            aHFMConst = aHFMConstCol(aKey)
            If aHFMConst.UD3.Value <> "" Then
                getConstantUD3 = aHFMConst.UD3.Value
            Else
                getConstantUD3 = pVal
            End If
        Else
            getConstantUD3 = pVal
        End If
    End Function

    Public Function getConstantICP(pAccount As String, pVal As String) As String
        Dim aKey As String
        Dim aHFMConst As HFMConst
        aKey = pAccount
        If contains(aHFMConstCol, aKey, "obj") Then
            aHFMConst = aHFMConstCol(aKey)
            If aHFMConst.ICP.Value <> "" Then
                getConstantICP = aHFMConst.ICP.Value
            Else
                getConstantICP = pVal
            End If
        Else
            getConstantICP = pVal
        End If
    End Function

    Private Function contains(col As Collection, Key As String, Optional aType As String = "var") As Boolean
        Dim obj As Object
        Dim var As Object
        On Error GoTo err
        contains = True
        If aType = "obj" Then
            obj = col(Key)
        Else
            var = col(Key)
        End If
        Exit Function
err:
        contains = False
    End Function

End Class